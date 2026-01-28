"""
Kelp Deal Flow Pipeline - Configuration Settings
"""
from pathlib import Path
from dataclasses import dataclass, field
from typing import Dict, List, Optional

# Base paths
BASE_DIR = Path(__file__).parent.parent
COMPANY_DATA_DIR = BASE_DIR / "Company Data"
OUTPUT_DIR = BASE_DIR / "output"
PPTX_OUTPUT_DIR = OUTPUT_DIR / "pptx"
CITATIONS_OUTPUT_DIR = OUTPUT_DIR / "citations"

# Ensure output directories exist
PPTX_OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
CITATIONS_OUTPUT_DIR.mkdir(parents=True, exist_ok=True)


@dataclass
class KelpBranding:
    """Kelp Branding Guidelines"""
    # Colors (RGB tuples)
    primary_dark_indigo: tuple = (45, 35, 75)  # Dark Indigo/Violet
    primary_violet: tuple = (75, 55, 120)
    accent_pink: tuple = (255, 105, 135)  # Pink gradient start
    accent_orange: tuple = (255, 165, 90)  # Orange gradient end
    accent_cyan: tuple = (0, 188, 212)  # Cyan Blue for icons
    background_white: tuple = (255, 255, 255)
    text_white: tuple = (255, 255, 255)
    text_dark_grey: tuple = (60, 60, 60)
    text_light_grey: tuple = (128, 128, 128)
    
    # Typography
    heading_font: str = "Arial"
    heading_font_bold: str = "Arial"
    body_font: str = "Arial"
    heading_size_large: int = 24  # Points
    heading_size_medium: int = 20
    body_size: int = 11
    footer_size: int = 9
    
    # Footer text
    footer_text: str = "Strictly Private & Confidential – Prepared by Kelp M&A Team"
    logo_text: str = "KELP"
    
    # Layout
    slide_width_inches: float = 13.333  # Widescreen 16:9
    slide_height_inches: float = 7.5


@dataclass
class SectorConfig:
    """Sector-specific configuration"""
    name: str
    keywords: List[str]
    slide1_title: str
    slide1_sections: List[str]
    slide2_title: str
    slide2_metrics: List[str]
    slide3_title: str
    slide3_hooks: List[str]
    image_queries: List[str]


# Sector configurations
SECTOR_CONFIGS: Dict[str, SectorConfig] = {
    "manufacturing": SectorConfig(
        name="Manufacturing & Industrials",
        keywords=["forge", "manufacturing", "machining", "industrial", "engineering", "production", "factory", "plant", "fabrication", "automotive components", "forging"],
        slide1_title="Business Profile & Infrastructure",
        slide1_sections=["Product Segments", "End-User Industries", "Manufacturing Footprint", "Certifications"],
        slide2_title="Financial & Operational Scale",
        slide2_metrics=["Revenue Growth (CAGR)", "EBITDA Margins", "Export Contribution", "Customer Count", "Capacity Utilization"],
        slide3_title="Investment Highlights",
        slide3_hooks=["Proprietary product portfolio with high entry barriers", "Strategic geographic advantage", "Robust financial performance with industry-leading margins", "Platform for consolidation"],
        image_queries=["industrial manufacturing facility", "precision machining factory", "metal forging plant", "engineering workshop"]
    ),
    "electronics": SectorConfig(
        name="Electronics & Defense",
        keywords=["electronics", "defence", "defense", "aerospace", "space", "satellite", "pcb", "embedded", "avionics", "radar", "centum", "drdo", "isro", "payload", "strategic electronics", "electronic equipment", "fpga"],
        slide1_title="Business Profile & Capabilities",
        slide1_sections=["Core Competencies", "Strategic Sectors", "Manufacturing Infrastructure", "Key Partnerships"],
        slide2_title="Financial Performance & Order Book",
        slide2_metrics=["Revenue Growth", "Order Book Value", "EBITDA Margins", "Export Revenue", "R&D Investment"],
        slide3_title="Investment Highlights",
        slide3_hooks=["Mission-critical electronics supplier to defense & space sectors", "Strong order book providing revenue visibility", "Strategic partnerships with global OEMs", "Indigenous technology capabilities"],
        image_queries=["electronics manufacturing facility", "aerospace components", "defense electronics", "satellite technology"]
    ),
    "pharma": SectorConfig(
        name="Pharmaceuticals",
        keywords=["pharma", "pharmaceutical", "medicine", "drug", "api", "formulation", "therapeutic", "healthcare", "clinical"],
        slide1_title="Business Profile & Product Portfolio",
        slide1_sections=["Therapeutic Areas", "Product Divisions", "Manufacturing Capabilities", "Regulatory Certifications"],
        slide2_title="Market Position & Financial Performance",
        slide2_metrics=["Revenue Growth", "Gross Margins", "Export Markets", "Product Portfolio Size", "R&D Pipeline"],
        slide3_title="Investment Highlights",
        slide3_hooks=["Diversified product portfolio across therapeutic areas", "Strong regulatory compliance and certifications", "Established distribution network", "Growing export presence"],
        image_queries=["pharmaceutical manufacturing", "medicine production facility", "pharma research lab", "drug manufacturing plant"]
    ),
    "technology": SectorConfig(
        name="Technology & IT Services",
        keywords=["software", "technology", "it services", "saas", "cloud", "digital", "ai", "machine learning", "salesforce", "development"],
        slide1_title="Business Profile & Service Portfolio",
        slide1_sections=["Core Services", "Technology Stack", "Client Industries", "Global Presence"],
        slide2_title="Growth Metrics & Unit Economics",
        slide2_metrics=["Revenue Growth (CAGR)", "Client Retention Rate", "Employee Count", "Profit Margins", "Order Pipeline"],
        slide3_title="Investment Highlights",
        slide3_hooks=["Strong expertise in emerging technologies", "High client retention demonstrating service quality", "Scalable business model with global delivery", "Strategic partnerships with technology leaders"],
        image_queries=["modern tech office", "software development team", "cloud computing data center", "IT services company"]
    ),
    "logistics": SectorConfig(
        name="Logistics & Supply Chain",
        keywords=["logistics", "supply chain", "express", "delivery", "transportation", "warehousing", "distribution", "freight", "courier"],
        slide1_title="Network & Infrastructure",
        slide1_sections=["Geographic Coverage", "Network Assets", "Service Offerings", "Technology Platform"],
        slide2_title="Operational Scale & Efficiency",
        slide2_metrics=["Revenue Growth", "Network Reach", "Delivery Volume", "EBITDA Margins", "Fleet Size"],
        slide3_title="Investment Highlights",
        slide3_hooks=["Pan-India network with extensive reach", "Capital-efficient operations with improving unit economics", "Technology-enabled logistics platform", "Consolidation opportunity in fragmented market"],
        image_queries=["logistics warehouse", "delivery fleet trucks", "distribution center", "supply chain operations"]
    ),
    "entertainment": SectorConfig(
        name="Entertainment & Media",
        keywords=["cinema", "entertainment", "multiplex", "media", "theatre", "movie", "film", "exhibition", "screens"],
        slide1_title="Business Profile & Market Presence",
        slide1_sections=["Screen Network", "Geographic Presence", "Service Offerings", "Brand Positioning"],
        slide2_title="Operating Metrics & Financial Performance",
        slide2_metrics=["Revenue Growth", "Screen Count", "Occupancy Rate", "Average Ticket Price", "F&B Revenue Per Head"],
        slide3_title="Investment Highlights",
        slide3_hooks=["Strategic presence in underserved Tier 2/3 markets", "Asset-light franchise model enabling rapid expansion", "Strong revenue per screen metrics", "Growing advertising revenue stream"],
        image_queries=["modern cinema multiplex", "movie theatre interior", "cinema auditorium", "entertainment venue"]
    )
}


@dataclass 
class LLMConfig:
    """LLM Configuration for Ollama"""
    model_name: str = "qwen2.5:7b"  # Use qwen2.5 which is available
    base_url: str = "http://localhost:11434"
    temperature_factual: float = 0.3  # For data extraction
    temperature_creative: float = 0.7  # For anonymization/rewriting
    max_tokens: int = 2048
    timeout: int = 120


@dataclass
class ImageConfig:
    """Image sourcing configuration"""
    unsplash_access_key: Optional[str] = None  # Set via environment variable
    pexels_api_key: Optional[str] = None  # Set via environment variable
    images_per_slide: int = 2
    min_width: int = 800
    min_height: int = 600


# Company mapping from folder names
COMPANY_FOLDERS = {
    "kalyani_forge": "automotive-kalyani-forge",
    "centum": "electronics-centum", 
    "ind_swift": "pharma-ind-swift",
    "ksolves": "technology-ksolves",
    "gati": "logistics-gati",
    "connplex": "entertainment-connplex"
}

# Initialize default configs
BRANDING = KelpBranding()
LLM_CONFIG = LLMConfig()
IMAGE_CONFIG = ImageConfig()
