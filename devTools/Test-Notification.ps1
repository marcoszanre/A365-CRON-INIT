# =============================================================================
# Test-Notification.ps1 - Test Agent 365 Notifications via PowerShell
# =============================================================================
# This script sends custom activities to test notification handling
# It bypasses the Agents Playground channelId bug
# =============================================================================

# ===========================
# CONFIGURATION - UPDATE THESE
# ===========================

# Get a fresh token with: a365 develop get-token
$BearerToken = "YOUR_BEARER_TOKEN_HERE"

# Your agent endpoint (local dev)
$AgentEndpoint = "http://localhost:3978/api/messages"

# Conversation ID - get from Agents Playground or use any UUID
$ConversationId = "00000000-0000-0000-0000-000000000001"

# Tenant ID (from a365.config.json)
$TenantId = "YOUR_TENANT_ID_HERE"

# Agent App ID (clientAppId from a365.config.json)
$AgentAppId = "YOUR_AGENT_APP_ID_HERE"

# Agent UPN (agentUserPrincipalName from a365.config.json)
$AgentUPN = "YOUR_AGENT_UPN_HERE"

# User email (managerEmail from a365.config.json)
$UserEmail = "YOUR_USER_EMAIL_HERE"

# ===========================
# NOTIFICATION TYPE SELECTION
# ===========================
# Sub-channel IDs from MS Docs:
#   - email      : Email notification (agent mentioned/addressed in email)
#   - word       : Word document comment notification
#   - excel      : Excel document comment notification
#   - powerpoint : PowerPoint document comment notification

$NotificationType = "email"  # Change to: email, word, excel, powerpoint

# ===========================
# HELPER FUNCTION
# ===========================

function Send-AgentNotification {
    param(
        [string]$SubChannel,
        [string]$NotificationName,
        [string]$ActivityName,  # emailNotification, wpxComment, agentLifecycle
        [hashtable]$NotificationValue
    )
    
    $ActivityId = "test-notif-" + (Get-Random -Maximum 99999)
    
    # Build the activity following MS Docs format:
    # https://learn.microsoft.com/en-us/microsoft-agent-365/developer/testing#email-notification
    # Key points:
    #   - type = "message" (not "event")
    #   - name = notification type (emailNotification, wpxComment, etc.)
    #   - channelId = "agents"
    #   - Notification data goes in "entities" array with type matching the notification
    $Activity = @{
        type = "message"  # Must be "message" per MS docs
        name = $ActivityName  # emailNotification, wpxComment, agentLifecycle
        channelId = "agents"  # REQUIRED
        id = $ActivityId
        timestamp = (Get-Date).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ss+00:00")
        serviceUrl = "http://localhost:3978"
        locale = "en-US"
        
        from = @{
            id = $UserEmail
            name = "Agent Manager"
            role = "user"
        }
        
        recipient = @{
            id = $AgentUPN
            name = "Agent"
            agenticUserId = $AgentUPN
            agenticAppId = $AgentAppId
            tenantId = $TenantId
        }
        
        conversation = @{
            conversationType = "personal"
            tenantId = $TenantId
            id = $ConversationId
        }
        
        # Notification data goes in entities array per MS docs
        # Note: Email notification already has 'type' in $NotificationValue
        entities = @(
            @{
                id = $SubChannel
                type = "productInfo"
            },
            @{
                type = "clientInfo"
                locale = "en-US"
            },
            # The actual notification entity (already includes type for email)
            $NotificationValue
        )
        
        channelData = @{
            tenant = @{
                id = $TenantId
            }
        }
        
        membersAdded = @()
        membersRemoved = @()
        reactionsAdded = @()
        reactionsRemoved = @()
        attachments = @()
        listenFor = @()
        textHighlights = @()
    }
    
    $JsonBody = $Activity | ConvertTo-Json -Depth 10
    
    Write-Host ""
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host "Sending $SubChannel notification" -ForegroundColor Cyan
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host "Activity ID: $ActivityId" -ForegroundColor Gray
    Write-Host ""
    
    # Show JSON for debugging (uncomment to see payload)
    # Write-Host "JSON Body:" -ForegroundColor DarkGray
    # Write-Host $JsonBody -ForegroundColor DarkGray
    
    try {
        $Response = Invoke-RestMethod -Uri $AgentEndpoint `
            -Method POST `
            -Headers @{
                "Authorization" = "Bearer $BearerToken"
                "Content-Type" = "application/json"
            } `
            -Body $JsonBody `
            -ErrorAction Stop
        
        Write-Host "✅ Notification sent successfully!" -ForegroundColor Green
        Write-Host "Response: $($Response | ConvertTo-Json -Depth 5)" -ForegroundColor Gray
    }
    catch {
        $StatusCode = $_.Exception.Response.StatusCode.value__
        Write-Host "❌ Error: HTTP $StatusCode" -ForegroundColor Red
        Write-Host $_.Exception.Message -ForegroundColor Red
        
        # Note: 404 errors on reply are EXPECTED in local dev (no connector)
        if ($StatusCode -eq 500) {
            Write-Host ""
            Write-Host "TIP: Check agent logs - the notification may have been processed" -ForegroundColor Yellow
            Write-Host "     even though the reply failed (no connector endpoint locally)" -ForegroundColor Yellow
        }
    }
}

# ===========================
# NOTIFICATION PAYLOADS
# Per SDK Pydantic models - use snake_case field names
# See: microsoft_agents_a365/notifications/models/email_reference.py
# ===========================

# Email Notification Entity
# SDK EmailReference expects: id, conversation_id, html_body (snake_case!)
$EmailNotification = @{
    type = "emailNotification"
    id = (New-Guid).ToString()
    conversation_id = $ConversationId
    html_body = @"
<body dir="ltr">
<div class="elementToProof" style="font-family: Aptos, Aptos_EmbeddedFont, Aptos_MSFontService, Calibri, Helvetica, sans-serif; font-size: 12pt; color: rgb(0, 0, 0);">
Hello Agent! Please create a Word document summarizing the Q1 sales report.
</div>
</body>
"@
}

# Word Comment Notification Entity (type will be added as "wpxCommentNotification")
$WordCommentNotification = @{
    documentId = (New-Guid).ToString()
    commentId = (New-Guid).ToString()
    driveId = (New-Guid).ToString()
    commentText = "@agent Please review this section and provide feedback"
    documentUrl = "https://example.sharepoint.com/sites/test/Document.docx"
    authorEmail = $UserEmail
}

# Excel Comment Notification Entity
$ExcelCommentNotification = @{
    documentId = (New-Guid).ToString()
    commentId = (New-Guid).ToString()
    driveId = (New-Guid).ToString()
    commentText = "@agent Please analyze this data"
    documentUrl = "https://example.sharepoint.com/sites/test/Spreadsheet.xlsx"
    authorEmail = $UserEmail
}

# PowerPoint Comment Notification Entity
$PowerPointCommentNotification = @{
    documentId = (New-Guid).ToString()
    commentId = (New-Guid).ToString()
    driveId = (New-Guid).ToString()
    commentText = "@agent Please improve this slide"
    documentUrl = "https://example.sharepoint.com/sites/test/Presentation.pptx"
    authorEmail = $UserEmail
}

# ===========================
# SEND THE NOTIFICATION
# ===========================

switch ($NotificationType) {
    "email" {
        Send-AgentNotification -SubChannel "email" `
            -NotificationName "emailNotification" `
            -ActivityName "emailNotification" `
            -NotificationValue $EmailNotification
    }
    "word" {
        Send-AgentNotification -SubChannel "word" `
            -NotificationName "wpxCommentNotification" `
            -ActivityName "wpxComment" `
            -NotificationValue $WordCommentNotification
    }
    "excel" {
        Send-AgentNotification -SubChannel "excel" `
            -NotificationName "wpxCommentNotification" `
            -ActivityName "wpxComment" `
            -NotificationValue $ExcelCommentNotification
    }
    "powerpoint" {
        Send-AgentNotification -SubChannel "powerpoint" `
            -NotificationName "wpxCommentNotification" `
            -ActivityName "wpxComment" `
            -NotificationValue $PowerPointCommentNotification
    }
    default {
        Write-Host "Unknown notification type: $NotificationType" -ForegroundColor Red
        Write-Host "Valid types: email, word, excel, powerpoint" -ForegroundColor Yellow
    }
}

Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Check agent logs for:" -ForegroundColor Cyan
Write-Host "  📬 NotificationTypes.EMAIL_NOTIFICATION" -ForegroundColor White
Write-Host "  or" -ForegroundColor Gray
Write-Host "  📬 NotificationTypes.WPX_COMMENT" -ForegroundColor White
Write-Host "========================================" -ForegroundColor Cyan
