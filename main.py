#!/usr/bin/env python3
# Copyright (c) Microsoft. All rights reserved.

"""
Main Entry Point

Start the A365 Agent Framework server with the Contoso Agent.

Usage:
    python main.py
    
    # Or with uv:
    uv run main.py
"""

import sys


def main() -> int:
    """Main entry point."""
    try:
        print("🚀 Starting A365 Agent Framework...")
        print()
        
        # Import here to ensure proper module loading
        from a365_agent import create_and_run_host
        from agents import ContosoAgent
        
        # Start the server
        create_and_run_host(ContosoAgent)
        
        return 0
        
    except ImportError as e:
        print(f"❌ Import error: {e}")
        print("Please ensure all dependencies are installed:")
        print("  uv pip install -e .")
        return 1
        
    except Exception as e:
        print(f"❌ Failed to start server: {e}")
        import traceback
        traceback.print_exc()
        return 1


if __name__ == "__main__":
    sys.exit(main())
