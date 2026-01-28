"""
Data Ingestion Module - Parses Markdown files and extracts structured data
"""
import re
import json
from pathlib import Path
from dataclasses import dataclass, field
from typing import Dict, List, Any, Optional
import sys
sys.path.insert(0, str(Path(__file__).parent.parent.parent))

from config.settings import COMPANY_DATA_DIR


@dataclass
class CompanyData:
    """Structured company data extracted from one-pager"""
    name: str = ""
    sector_template: str = ""
    description: str = ""
    website: str = ""
    products_services: List[str] = field(default_factory=list)
    industries_served: List[str] = field(default_factory=list)
    key_operational_indicators: List[str] = field(default_factory=list)
    shareholders: List[Dict[str, Any]] = field(default_factory=list)
    milestones: List[Dict[str, str]] = field(default_factory=list)
    clients: List[str] = field(default_factory=list)
    awards_certifications: List[str] = field(default_factory=list)
    market_size_data: List[Dict[str, Any]] = field(default_factory=list)
    swot: Dict[str, List[str]] = field(default_factory=dict)
    global_presence: List[str] = field(default_factory=list)
    future_plans: List[str] = field(default_factory=list)
    financials: Dict[str, Any] = field(default_factory=dict)
    channel_mix: Dict[str, Any] = field(default_factory=dict)
    product_portfolio: List[Dict[str, str]] = field(default_factory=list)
    founded: str = ""
    headquarters: str = ""
    domain: str = ""
    segment: str = ""
    partners: List[str] = field(default_factory=list)
    
    # Metadata for citations
    source_file: str = ""
    raw_sections: Dict[str, str] = field(default_factory=dict)


class MarkdownParser:
    """Parses company one-pager markdown files"""
    
    def __init__(self, file_path: Path):
        self.file_path = file_path
        self.content = ""
        self.sections: Dict[str, str] = {}
        
    def load(self) -> str:
        """Load markdown content from file"""
        with open(self.file_path, 'r', encoding='utf-8') as f:
            self.content = f.read()
        return self.content
    
    def extract_sections(self) -> Dict[str, str]:
        """Extract all sections from markdown"""
        # Split by ## headers
        pattern = r'^## (.+?)$'
        lines = self.content.split('\n')
        
        current_section = "header"
        section_content = []
        
        for line in lines:
            header_match = re.match(pattern, line)
            if header_match:
                # Save previous section
                if section_content:
                    self.sections[current_section] = '\n'.join(section_content).strip()
                current_section = header_match.group(1).strip()
                section_content = []
            else:
                section_content.append(line)
        
        # Save last section
        if section_content:
            self.sections[current_section] = '\n'.join(section_content).strip()
            
        return self.sections
    
    def parse_table(self, table_text: str) -> List[Dict[str, str]]:
        """Parse markdown table into list of dicts"""
        lines = [l.strip() for l in table_text.strip().split('\n') if l.strip()]
        if len(lines) < 2:
            return []
        
        # Find header row (first row with |)
        header_idx = 0
        for i, line in enumerate(lines):
            if '|' in line and '---' not in line:
                header_idx = i
                break
        
        if header_idx >= len(lines):
            return []
            
        # Extract headers
        header_line = lines[header_idx]
        headers = [h.strip() for h in header_line.split('|') if h.strip()]
        
        # Extract data rows (skip separator)
        results = []
        for line in lines[header_idx + 1:]:
            if '---' in line or not line.strip():
                continue
            if '|' not in line:
                continue
                
            values = [v.strip() for v in line.split('|') if v.strip() or line.count('|') > len(headers)]
            # Handle empty cells
            all_values = []
            parts = line.split('|')[1:-1] if line.startswith('|') else line.split('|')
            all_values = [p.strip() for p in parts]
            
            if len(all_values) >= len(headers):
                row = {headers[i]: all_values[i] for i in range(len(headers))}
                results.append(row)
                
        return results
    
    def parse_list(self, text: str) -> List[str]:
        """Parse markdown list items"""
        items = []
        for line in text.split('\n'):
            line = line.strip()
            # Match - item or * item or numbered list
            match = re.match(r'^[-*]\s+(.+)$', line)
            if match:
                items.append(match.group(1).strip())
            elif re.match(r'^\d+\.\s+(.+)$', line):
                items.append(re.match(r'^\d+\.\s+(.+)$', line).group(1).strip())
        return items
    
    def parse_nested_list(self, text: str) -> List[str]:
        """Parse nested markdown list with sub-items"""
        items = []
        current_main = ""
        
        for line in text.split('\n'):
            line = line.strip()
            if not line:
                continue
                
            # Main item with bold
            main_match = re.match(r'^[-*]\s+\*\*(.+?)\*\*(?:\s*\((.+?)\))?', line)
            if main_match:
                current_main = main_match.group(1).strip()
                sub_items = main_match.group(2) if main_match.group(2) else ""
                if sub_items:
                    items.append(f"{current_main}: {sub_items}")
                else:
                    items.append(current_main)
            elif re.match(r'^[-*]\s+(.+)$', line):
                items.append(re.match(r'^[-*]\s+(.+)$', line).group(1).strip())
                
        return items
    
    def extract_company_name(self) -> str:
        """Extract company name from filename or content"""
        # Try from filename
        name = self.file_path.stem.replace('-OnePager', '').replace('-', ' ')
        
        # Try from content header
        match = re.search(r'^# .+?Template:.+?\n\n## Business Description\n\n(.+?) (?:is|are|Ltd|Limited|Inc)', 
                         self.content, re.MULTILINE | re.DOTALL)
        if match:
            name = match.group(1).strip()
        else:
            # Try simpler pattern
            desc_section = self.sections.get('Business Description', '')
            if desc_section:
                words = desc_section.split()[:5]
                for i, word in enumerate(words):
                    if word.lower() in ['is', 'are', 'limited', 'ltd', 'inc']:
                        name = ' '.join(words[:i])
                        break
                        
        return name.strip()
    
    def extract_template_type(self) -> str:
        """Extract the template type from header"""
        match = re.search(r'Template:\s*(.+?)(?:\n|$)', self.content)
        if match:
            return match.group(1).strip()
        return "Default"


class CompanyDataExtractor:
    """Extracts structured company data from parsed markdown"""
    
    def __init__(self, parser: MarkdownParser):
        self.parser = parser
        self.data = CompanyData()
        
    def extract_all(self) -> CompanyData:
        """Extract all company data"""
        self.parser.load()
        self.parser.extract_sections()
        
        self.data.source_file = str(self.parser.file_path)
        self.data.raw_sections = self.parser.sections.copy()
        
        # Extract each field
        self.data.name = self.parser.extract_company_name()
        self.data.sector_template = self.parser.extract_template_type()
        self.data.description = self.parser.sections.get('Business Description', '')
        self.data.website = self.parser.sections.get('Website', '').strip()
        
        # Products & Services
        ps_section = self.parser.sections.get('Product & Services', '')
        self.data.products_services = self.parser.parse_nested_list(ps_section)
        
        # Industries served
        ind_section = self.parser.sections.get('Application areas / Industries served', '')
        if ind_section:
            self.data.industries_served = [i.strip() for i in ind_section.split(',')]
        
        # Key Operational Indicators
        koi_section = self.parser.sections.get('Key Operational Indicators', '')
        if koi_section:
            self.data.key_operational_indicators = self._parse_operational_indicators(koi_section)
        
        # Shareholders
        sh_section = self.parser.sections.get('Shareholders', '')
        if sh_section:
            self.data.shareholders = self.parser.parse_table(sh_section)
        
        # Milestones
        ms_section = self.parser.sections.get('Key Milestones', '')
        if ms_section:
            self.data.milestones = self.parser.parse_table(ms_section)
        
        # Clients
        cl_section = self.parser.sections.get('Clients', '')
        if cl_section:
            self.data.clients = [c.strip() for c in cl_section.replace('\n', ' ').split() if c.strip()]
        
        # Awards and Certifications
        ac_section = self.parser.sections.get('Awards and Certifications', '')
        if ac_section:
            self.data.awards_certifications = self.parser.parse_list(ac_section)
        
        # Market Size
        mkt_section = self.parser.sections.get('Market Size', '')
        if mkt_section:
            self.data.market_size_data = self.parser.parse_table(mkt_section)
        
        # SWOT
        swot_section = self.parser.sections.get('SWOT', '')
        if swot_section:
            self.data.swot = self._parse_swot(swot_section)
        
        # Global Presence
        gp_section = self.parser.sections.get('Global Presence', '')
        if gp_section:
            self.data.global_presence = [g.strip() for g in gp_section.split(',')]
        
        # Future Plans
        fp_section = self.parser.sections.get('Future Plan', '')
        if fp_section:
            self.data.future_plans = self.parser.parse_list(fp_section)
        
        # Channel Mix (for entertainment/retail)
        cm_section = self.parser.sections.get('Channel Mix', '')
        if cm_section:
            self.data.channel_mix = self._parse_channel_mix(cm_section)
        
        # Product Portfolio (for pharma)
        pp_section = self.parser.sections.get('Product Portfolio', '')
        if pp_section:
            self.data.product_portfolio = self.parser.parse_table(pp_section)
        
        # Partners
        pt_section = self.parser.sections.get('Partners', '')
        if pt_section and pt_section.lower() != 'not available':
            self.data.partners = [p.strip() for p in pt_section.split(',')]
        
        # Details section parsing
        details_section = self.parser.sections.get('Details', '')
        if details_section:
            self._parse_details(details_section)
        
        return self.data
    
    def _parse_operational_indicators(self, text: str) -> List[str]:
        """Parse key operational indicators"""
        indicators = []
        for line in text.split('\n'):
            line = line.strip()
            if line.startswith('*'):
                # Extract the content after **Key:** Value pattern
                match = re.match(r'\*\s*\*\*(.+?):\*\*\s*(.+)', line)
                if match:
                    key = match.group(1).strip()
                    value = match.group(2).strip()
                    indicators.append(f"{key}: {value}")
                else:
                    indicators.append(line.lstrip('* '))
        return indicators
    
    def _parse_swot(self, text: str) -> Dict[str, List[str]]:
        """Parse SWOT analysis"""
        swot = {"Strengths": [], "Weaknesses": [], "Opportunities": [], "Threats": []}
        current_section = None
        
        for line in text.split('\n'):
            line = line.strip()
            
            # Check for section headers
            if '### Strengths' in line or line == 'Strengths':
                current_section = 'Strengths'
            elif '### Weaknesses' in line or line == 'Weaknesses':
                current_section = 'Weaknesses'
            elif '### Opportunities' in line or line == 'Opportunities':
                current_section = 'Opportunities'
            elif '### Threats' in line or line == 'Threats':
                current_section = 'Threats'
            elif current_section and line.startswith('-'):
                item = line.lstrip('- ').strip()
                if item:
                    swot[current_section].append(item)
                    
        return swot
    
    def _parse_channel_mix(self, text: str) -> Dict[str, Any]:
        """Parse channel mix data"""
        channel_data = {}
        for line in text.split('\n'):
            line = line.strip()
            if ':' in line and line.startswith('-'):
                key, value = line.lstrip('- ').split(':', 1)
                channel_data[key.strip()] = value.strip()
        return channel_data
    
    def _parse_details(self, text: str) -> None:
        """Parse details section for founded, headquarters, etc."""
        patterns = {
            'founded': r'Founded:\s*\*\*(.+?)\*\*',
            'headquarters': r'Headquarters:\s*\*\*(.+?)\*\*',
            'domain': r'Domain:\s*\*\*(.+?)\*\*',
            'segment': r'Segment:\s*\*\*(.+?)\*\*'
        }
        
        for field, pattern in patterns.items():
            match = re.search(pattern, text)
            if match:
                setattr(self.data, field, match.group(1).strip())


def load_company_data(company_folder: str) -> CompanyData:
    """Load and parse company data from folder"""
    folder_path = COMPANY_DATA_DIR / company_folder
    
    # Find the markdown file
    md_files = list(folder_path.glob("*.md"))
    if not md_files:
        raise FileNotFoundError(f"No markdown files found in {folder_path}")
    
    md_file = md_files[0]
    
    parser = MarkdownParser(md_file)
    extractor = CompanyDataExtractor(parser)
    
    return extractor.extract_all()


def load_all_companies() -> Dict[str, CompanyData]:
    """Load all company data from Company Data directory"""
    companies = {}
    
    for folder in COMPANY_DATA_DIR.iterdir():
        if folder.is_dir():
            try:
                data = load_company_data(folder.name)
                companies[folder.name] = data
                print(f"✓ Loaded: {data.name} ({folder.name})")
            except Exception as e:
                print(f"✗ Error loading {folder.name}: {e}")
    
    return companies


if __name__ == "__main__":
    # Test loading
    print("Loading all company data...\n")
    companies = load_all_companies()
    
    print(f"\n{'='*60}")
    print(f"Loaded {len(companies)} companies:")
    for folder, data in companies.items():
        print(f"\n{data.name}:")
        print(f"  - Sector: {data.sector_template}")
        print(f"  - Products: {len(data.products_services)}")
        print(f"  - Industries: {len(data.industries_served)}")
        print(f"  - Clients: {len(data.clients)}")
        print(f"  - Awards: {len(data.awards_certifications)}")
