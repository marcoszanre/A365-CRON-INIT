# =============================================================================
# Test-ProactiveTeamsMessage.ps1 - Test Blueprint Auth + Teams MCP
# =============================================================================
# This script tests the Agent User Impersonation flow to send a Teams message
# using the blueprint credentials from .env
#
# Flow: Blueprint → T1 → T2 → Resource Token → Teams Graph API
# =============================================================================

param(
    [Parameter(Mandatory=$false)]
    [string]$TargetUserEmail,
    
    [Parameter(Mandatory=$false)]
    [string]$Message = "Hello! This is a proactive test message from Contoso Agent."
)

# ===========================
# LOAD .ENV FILE
# ===========================

$EnvFile = Join-Path $PSScriptRoot "..\.env"
if (Test-Path $EnvFile) {
    Get-Content $EnvFile | ForEach-Object {
        if ($_ -match '^\s*([^#][^=]+)=(.*)$') {
            $key = $matches[1].Trim()
            $value = $matches[2].Trim()
            [Environment]::SetEnvironmentVariable($key, $value, "Process")
        }
    }
    Write-Host "✅ Loaded .env file" -ForegroundColor Green
} else {
    Write-Host "❌ .env file not found at $EnvFile" -ForegroundColor Red
    exit 1
}

# ===========================
# CONFIGURATION FROM .ENV
# ===========================

$BlueprintClientId = $env:CONNECTIONS__SERVICE_CONNECTION__SETTINGS__CLIENTID
$BlueprintClientSecret = $env:CONNECTIONS__SERVICE_CONNECTION__SETTINGS__CLIENTSECRET
$TenantId = $env:CONNECTIONS__SERVICE_CONNECTION__SETTINGS__TENANTID
$AgentIdentityClientId = $env:AGENT_IDENTITY_CLIENT_ID
$AgentUserUPN = $env:AGENT_USER_UPN

# Use param if provided, otherwise use .env value
if (-not $TargetUserEmail) {
    $TargetUserEmail = $env:TARGET_USER_EMAIL
}

# Validate required values
$missing = @()
if (-not $BlueprintClientId) { $missing += "CONNECTIONS__SERVICE_CONNECTION__SETTINGS__CLIENTID" }
if (-not $BlueprintClientSecret) { $missing += "CONNECTIONS__SERVICE_CONNECTION__SETTINGS__CLIENTSECRET" }
if (-not $TenantId) { $missing += "CONNECTIONS__SERVICE_CONNECTION__SETTINGS__TENANTID" }
if (-not $AgentIdentityClientId) { $missing += "AGENT_IDENTITY_CLIENT_ID" }
if (-not $AgentUserUPN) { $missing += "AGENT_USER_UPN" }
if (-not $TargetUserEmail) { $missing += "TARGET_USER_EMAIL" }

if ($missing.Count -gt 0) {
    Write-Host "❌ Missing required .env values:" -ForegroundColor Red
    $missing | ForEach-Object { Write-Host "   - $_" -ForegroundColor Yellow }
    exit 1
}

Write-Host ""
Write-Host "=== Configuration ===" -ForegroundColor Cyan
Write-Host "Blueprint Client ID : $BlueprintClientId"
Write-Host "Tenant ID           : $TenantId"
Write-Host "Agent Identity ID   : $AgentIdentityClientId"
Write-Host "Agent User UPN      : $AgentUserUPN"
Write-Host "Target User         : $TargetUserEmail"
Write-Host ""

Read-Host "Press ENTER to start the token flow..."

# ===========================
# STEP 1: Get Blueprint Exchange Token (T1)
# ===========================

Write-Host ""
Write-Host "🔐 Step 1: Acquiring Blueprint Exchange Token (T1)..." -ForegroundColor Yellow
Write-Host "   POST $tokenUrl" -ForegroundColor DarkGray
Write-Host "   client_id     = $BlueprintClientId (Blueprint)" -ForegroundColor DarkGray
Write-Host "   scope         = api://AzureADTokenExchange/.default" -ForegroundColor DarkGray
Write-Host "   fmi_path      = $AgentIdentityClientId (Agent Identity)" -ForegroundColor DarkGray
Write-Host "   grant_type    = client_credentials" -ForegroundColor DarkGray

$tokenUrl = "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token"

$body1 = @{
    client_id     = $BlueprintClientId
    scope         = "api://AzureADTokenExchange/.default"
    grant_type    = "client_credentials"
    client_secret = $BlueprintClientSecret
    fmi_path      = $AgentIdentityClientId
}

try {
    $response1 = Invoke-RestMethod -Uri $tokenUrl -Method Post -Body $body1 -ContentType "application/x-www-form-urlencoded"
    $T1 = $response1.access_token
    Write-Host "✅ Got T1 (Blueprint Exchange Token)" -ForegroundColor Green
    Write-Host "   Token length: $($T1.Length) chars" -ForegroundColor DarkGray
    Write-Host "   Expires in: $($response1.expires_in) seconds" -ForegroundColor DarkGray
} catch {
    Write-Host "❌ Failed to get T1: $($_.Exception.Message)" -ForegroundColor Red
    if ($_.ErrorDetails.Message) {
        Write-Host "   Details: $($_.ErrorDetails.Message)" -ForegroundColor Yellow
    }
    exit 1
}

Read-Host "Press ENTER to continue to Step 2..."

# ===========================
# STEP 2: Get Agent Identity Exchange Token (T2)
# ===========================

Write-Host ""
Write-Host "🔐 Step 2: Acquiring MCP Token (Autonomous App Flow)..." -ForegroundColor Yellow
Write-Host "   POST $tokenUrl" -ForegroundColor DarkGray
Write-Host "   client_id            = $AgentIdentityClientId (Agent Identity)" -ForegroundColor DarkGray
Write-Host "   scope                = api://ea9ffc3e-8a23-4a7d-836d-234d7c7565c1/McpServers.Teams.All" -ForegroundColor DarkGray
Write-Host "   client_assertion     = T1 (from step 1)" -ForegroundColor DarkGray
Write-Host "   grant_type           = client_credentials" -ForegroundColor DarkGray

$body2 = @{
    client_id              = $AgentIdentityClientId
    scope                  = "api://ea9ffc3e-8a23-4a7d-836d-234d7c7565c1/McpServers.Teams.All"
    grant_type             = "client_credentials"
    client_assertion_type  = "urn:ietf:params:oauth:client-assertion-type:jwt-bearer"
    client_assertion       = $T1
}

try {
    $response2 = Invoke-RestMethod -Uri $tokenUrl -Method Post -Body $body2 -ContentType "application/x-www-form-urlencoded"
    $McpToken = $response2.access_token
    Write-Host "✅ Got MCP Token (Autonomous)" -ForegroundColor Green
    Write-Host "   Token length: $($McpToken.Length) chars" -ForegroundColor DarkGray
    Write-Host "   Expires in: $($response2.expires_in) seconds" -ForegroundColor DarkGray
} catch {
    Write-Host "❌ Failed to get MCP Token: $($_.Exception.Message)" -ForegroundColor Red
    if ($_.ErrorDetails.Message) {
        Write-Host "   Details: $($_.ErrorDetails.Message)" -ForegroundColor Yellow
    }
    exit 1
}

Read-Host "Press ENTER to continue to Step 3 (Create Chat via MCP)..."

# ===========================
# STEP 3: Create 1:1 Chat via Teams MCP Server
# ===========================

$McpServerUrl = "https://agent365.svc.cloud.microsoft/agents/servers/mcp_TeamsServer"

Write-Host ""
Write-Host "💬 Step 3: Creating 1:1 Chat via Teams MCP..." -ForegroundColor Yellow
Write-Host "   POST $McpServerUrl/mcp" -ForegroundColor DarkGray
Write-Host "   Tool: mcp_graph_chat_createChat" -ForegroundColor DarkGray
Write-Host "   chatType = oneOnOne" -ForegroundColor DarkGray
Write-Host "   members  = $TargetUserEmail" -ForegroundColor DarkGray

$headers = @{
    Authorization  = "Bearer $McpToken"
    "Content-Type" = "application/json"
}

# MCP JSON-RPC request format
$createChatRequest = @{
    jsonrpc = "2.0"
    id = 1
    method = "tools/call"
    params = @{
        name = "mcp_graph_chat_createChat"
        arguments = @{
            chatType = "oneOnOne"
            members = @($TargetUserEmail)
        }
    }
} | ConvertTo-Json -Depth 10

try {
    $chatResponse = Invoke-RestMethod -Uri "$McpServerUrl/mcp" -Method Post -Headers $headers -Body $createChatRequest
    
    # Parse MCP response
    if ($chatResponse.error) {
        throw "MCP Error: $($chatResponse.error.message)"
    }
    
    $chatResult = $chatResponse.result
    $chatId = $chatResult.content[0].text | ConvertFrom-Json | Select-Object -ExpandProperty id
    Write-Host "✅ Chat created/found: $chatId" -ForegroundColor Green
} catch {
    Write-Host "❌ Failed to create chat: $($_.Exception.Message)" -ForegroundColor Red
    if ($_.ErrorDetails.Message) {
        Write-Host "   Details: $($_.ErrorDetails.Message)" -ForegroundColor Yellow
    }
    exit 1
}

Read-Host "Press ENTER to continue to Step 4 (Send Message)..."

# ===========================
# STEP 4: Post Message to Chat via Teams MCP
# ===========================

Write-Host ""
Write-Host "📤 Step 4: Posting message via Teams MCP..." -ForegroundColor Yellow
Write-Host "   POST $McpServerUrl/mcp" -ForegroundColor DarkGray
Write-Host "   Tool: mcp_graph_chat_postMessage" -ForegroundColor DarkGray
Write-Host "   chatId: $chatId" -ForegroundColor DarkGray
Write-Host "   Message: $Message" -ForegroundColor DarkGray

$postMessageRequest = @{
    jsonrpc = "2.0"
    id = 2
    method = "tools/call"
    params = @{
        name = "mcp_graph_chat_postMessage"
        arguments = @{
            chatId = $chatId
            content = $Message
        }
    }
} | ConvertTo-Json -Depth 10

try {
    $messageResponse = Invoke-RestMethod -Uri "$McpServerUrl/mcp" -Method Post -Headers $headers -Body $postMessageRequest
    
    if ($messageResponse.error) {
        throw "MCP Error: $($messageResponse.error.message)"
    }
    
    Write-Host "✅ Message sent successfully!" -ForegroundColor Green
    Write-Host ""
    Write-Host "=== Result ===" -ForegroundColor Cyan
    Write-Host "Chat ID    : $chatId"
    Write-Host "Response   : $($messageResponse.result | ConvertTo-Json -Compress)"
    Write-Host ""
    Write-Host "🎉 Proactive Teams message sent via MCP from $AgentUserUPN to $TargetUserEmail" -ForegroundColor Green
} catch {
    Write-Host "❌ Failed to send message: $($_.Exception.Message)" -ForegroundColor Red
    if ($_.ErrorDetails.Message) {
        Write-Host "   Details: $($_.ErrorDetails.Message)" -ForegroundColor Yellow
    }
    exit 1
}
