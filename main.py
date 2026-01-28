#!/usr/bin/env python3
"""
╔═══════════════════════════════════════════════════════════════════════════════╗
║                                                                               ║
║   ██╗  ██╗███████╗██╗     ██████╗                                             ║
║   ██║ ██╔╝██╔════╝██║     ██╔══██╗                                            ║
║   █████╔╝ █████╗  ██║     ██████╔╝                                            ║
║   ██╔═██╗ ██╔══╝  ██║     ██╔═══╝                                             ║
║   ██║  ██╗███████╗███████╗██║                                                 ║
║   ╚═╝  ╚═╝╚══════╝╚══════╝╚═╝                                                 ║
║                                                                               ║
║   AI-Powered Investment Teaser Generation Pipeline                           ║
║   ─────────────────────────────────────────────────                           ║
║   Automated Deal Flow • Privacy-Preserving • GPU-Accelerated                 ║
║                                                                               ║
╚═══════════════════════════════════════════════════════════════════════════════╝

KELP transforms raw company markdown data into professional investment teasers
with AI-powered content generation, real-time web research, and beautiful
PowerPoint presentations.

Usage:
    python main.py                        # Process all companies
    python main.py --company kalyani      # Process specific company
    python main.py --help                 # Show help

Author: Shubrojyoti Dey
License: MIT
"""

import sys
from pathlib import Path

# Add project root to path
PROJECT_ROOT = Path(__file__).parent
sys.path.insert(0, str(PROJECT_ROOT))

# Import and run the main pipeline
from pipeline_v5_enhanced import main

if __name__ == "__main__":
    print("""
╔═══════════════════════════════════════════════════════════════════════════════╗
║  KELP - AI Investment Teaser Pipeline                                        ║
║  ─────────────────────────────────────                                        ║
║  🚀 GPU-Accelerated LLM Content Generation                                   ║
║  🌐 Gemini-Style Deep Web Research                                           ║
║  📊 Data-Dense Professional Presentations                                    ║
╚═══════════════════════════════════════════════════════════════════════════════╝
    """)
    main()
