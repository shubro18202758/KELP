"""
Content Generation Module - LLM-powered content generation and anonymization
Uses Ollama for local LLM inference
"""
import re
import json
from typing import Dict, List, Optional, Any
from dataclasses import dataclass
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent.parent))

from config.settings import LLM_CONFIG


@dataclass
class GeneratedContent:
    """Container for LLM-generated content"""
    original: str
    anonymized: str
    source: str
    citations: List[str]


class OllamaInterface:
    """Interface for Ollama LLM"""
    
    def __init__(self, model_name: str = None, base_url: str = None):
        self.model_name = model_name or LLM_CONFIG.model_name
        self.base_url = base_url or LLM_CONFIG.base_url
        self._available = None
        
    def is_available(self) -> bool:
        """Check if Ollama is running and model is available"""
        if self._available is not None:
            return self._available
            
        try:
            import requests
            response = requests.get(f"{self.base_url}/api/tags", timeout=5)
            if response.status_code == 200:
                models = response.json().get('models', [])
                model_names = [m.get('name', '').split(':')[0] for m in models]
                self._available = any(self.model_name.split(':')[0] in name for name in model_names)
            else:
                self._available = False
        except Exception:
            self._available = False
            
        return self._available
    
    def generate(self, prompt: str, temperature: float = 0.5, max_tokens: int = 1024) -> str:
        """Generate text using Ollama"""
        try:
            import requests
            
            response = requests.post(
                f"{self.base_url}/api/generate",
                json={
                    "model": self.model_name,
                    "prompt": prompt,
                    "stream": False,
                    "options": {
                        "temperature": temperature,
                        "num_predict": max_tokens
                    }
                },
                timeout=LLM_CONFIG.timeout
            )
            
            if response.status_code == 200:
                return response.json().get('response', '')
            else:
                print(f"Ollama error: {response.status_code}")
                return ""
                
        except Exception as e:
            print(f"Ollama generation failed: {e}")
            return ""


class ContentAnonymizer:
    """Anonymizes company content while preserving facts"""
    
    def __init__(self, llm: Optional[OllamaInterface] = None):
        self.llm = llm or OllamaInterface()
        self.use_llm = self.llm.is_available()
        
        if not self.use_llm:
            print("⚠ Ollama not available. Using rule-based anonymization.")
    
    def anonymize_description(self, description: str, company_name: str, sector: str) -> str:
        """Anonymize company description"""
        if self.use_llm:
            return self._llm_anonymize(description, company_name, sector)
        return self._rule_based_anonymize(description, company_name, sector)
    
    def _llm_anonymize(self, description: str, company_name: str, sector: str) -> str:
        """Use LLM for smart anonymization"""
        prompt = f"""You are writing a "blind" investment teaser for M&A purposes. 
Rewrite the following company description to ANONYMIZE the company identity while preserving all factual information.

Rules:
1. Remove all company names, founder names, and specific location names
2. Replace with generic descriptions like "The Company", "A leading [sector] company"
3. Keep all numbers, percentages, and metrics EXACTLY as stated
4. Keep the professional, investment-ready tone
5. Do not add any information not in the original
6. Keep it concise (2-3 sentences max)

Company Name to Remove: {company_name}
Sector: {sector}

Original Description:
{description}

Anonymized Description (2-3 sentences):"""

        result = self.llm.generate(prompt, temperature=0.3, max_tokens=300)
        
        # Clean up result
        result = result.strip()
        if result.startswith('"'):
            result = result[1:]
        if result.endswith('"'):
            result = result[:-1]
            
        # Fallback if LLM fails
        if not result or len(result) < 50:
            return self._rule_based_anonymize(description, company_name, sector)
            
        return result
    
    def _rule_based_anonymize(self, description: str, company_name: str, sector: str) -> str:
        """Rule-based anonymization fallback"""
        text = description
        
        # Remove company name variations
        name_parts = company_name.split()
        for part in name_parts:
            if len(part) > 3:  # Avoid removing common words
                text = re.sub(rf'\b{re.escape(part)}\b', '', text, flags=re.IGNORECASE)
        
        # Replace common patterns
        text = re.sub(r'\b(Ltd|Limited|Inc|Corporation|Corp|Pvt|Private)\b\.?', '', text, flags=re.IGNORECASE)
        
        # Clean up
        text = re.sub(r'\s+', ' ', text).strip()
        text = re.sub(r'\s+([,.])', r'\1', text)
        
        # Add generic intro if needed
        if not text.lower().startswith(('a ', 'the ', 'an ')):
            sector_desc = {
                'manufacturing': 'A leading manufacturing company',
                'electronics': 'A technology solutions provider',
                'pharma': 'An established pharmaceutical company',
                'technology': 'A software development company',
                'logistics': 'A logistics and supply chain company',
                'entertainment': 'An entertainment company'
            }
            prefix = sector_desc.get(sector.lower(), 'A company')
            text = f"{prefix} {text}"
        
        return text
    
    def anonymize_metric(self, metric: str, company_name: str) -> str:
        """Anonymize a metric while keeping numbers"""
        text = metric
        
        # Remove company name
        name_parts = company_name.split()
        for part in name_parts:
            if len(part) > 3:
                text = re.sub(rf'\b{re.escape(part)}\b', 'The Company', text, flags=re.IGNORECASE)
        
        return text
    
    def anonymize_highlight(self, highlight: str, company_name: str) -> str:
        """Anonymize investment highlight"""
        return self.anonymize_metric(highlight, company_name)


class ContentGenerator:
    """Generates polished content for slides"""
    
    def __init__(self, llm: Optional[OllamaInterface] = None):
        self.llm = llm or OllamaInterface()
        self.anonymizer = ContentAnonymizer(self.llm)
        self.use_llm = self.llm.is_available()
    
    def generate_slide1_content(self, raw_content: Dict, company_name: str, sector: str) -> Dict:
        """Generate polished content for Slide 1"""
        content = raw_content.copy()
        
        # Anonymize description
        if content.get('company_description'):
            content['company_description'] = self.anonymizer.anonymize_description(
                content['company_description'], company_name, sector
            )
        
        return content
    
    def generate_slide2_content(self, raw_content: Dict, company_name: str) -> Dict:
        """Generate polished content for Slide 2"""
        content = raw_content.copy()
        
        # Anonymize metrics
        if content.get('metrics'):
            content['metrics'] = [
                self.anonymizer.anonymize_metric(m, company_name) 
                for m in content['metrics'] if m
            ]
        
        return content
    
    def generate_slide3_content(self, raw_content: Dict, company_name: str, sector: str) -> Dict:
        """Generate polished content for Slide 3"""
        content = raw_content.copy()
        
        # Anonymize highlights
        if content.get('highlights'):
            content['highlights'] = [
                self.anonymizer.anonymize_highlight(h, company_name)
                for h in content['highlights']
            ]
        
        # Generate compelling hooks if LLM available
        if self.use_llm and content.get('highlights'):
            enhanced = self._enhance_highlights(content['highlights'], sector)
            if enhanced:
                content['highlights'] = enhanced
        
        return content
    
    def _enhance_highlights(self, highlights: List[str], sector: str) -> Optional[List[str]]:
        """Use LLM to enhance investment highlights"""
        prompt = f"""You are writing investment highlights for an M&A teaser in the {sector} sector.

Given these draft highlights, rewrite them to be:
1. Compelling and professional
2. Suitable for sophisticated investors
3. Action-oriented and forward-looking
4. Concise (max 15 words each)
5. Anonymized (no company names)

Draft Highlights:
{chr(10).join(f'- {h}' for h in highlights[:5])}

Rewrite as 4 polished bullet points (one per line, no bullet markers):"""

        result = self.llm.generate(prompt, temperature=0.5, max_tokens=200)
        
        if not result:
            return None
        
        # Parse result
        lines = [l.strip() for l in result.strip().split('\n') if l.strip()]
        lines = [l.lstrip('•-*1234567890. ') for l in lines]
        lines = [l for l in lines if len(l) > 10 and len(l) < 150]
        
        if len(lines) >= 3:
            return lines[:4]
        return None
    
    def generate_executive_summary(self, company_data: Any, sector: str) -> str:
        """Generate a brief executive summary for the cover or intro"""
        if not self.use_llm:
            return f"A leading player in the {sector} sector with strong operational capabilities and growth potential."
        
        prompt = f"""Write a 2-sentence executive summary for an investment teaser.

Sector: {sector}
Key Facts:
- Industries: {', '.join(company_data.industries_served[:5]) if company_data.industries_served else 'Various'}
- Products: {', '.join(company_data.products_services[:3]) if company_data.products_services else 'Multiple'}
- Presence: {', '.join(company_data.global_presence) if company_data.global_presence else 'National'}

Write an anonymous, compelling 2-sentence summary for investors:"""

        result = self.llm.generate(prompt, temperature=0.5, max_tokens=150)
        
        if result and len(result) > 50:
            return result.strip()
        
        return f"A leading player in the {sector} sector with strong operational capabilities and growth potential."


def generate_all_content(slide_content: Dict, company_data: Any, sector: str) -> Dict:
    """
    Main function to generate all slide content.
    Returns enhanced and anonymized content for all slides.
    """
    generator = ContentGenerator()
    company_name = company_data.name
    
    generated = {
        "slide1": generator.generate_slide1_content(
            slide_content['slide1'], company_name, sector
        ),
        "slide2": generator.generate_slide2_content(
            slide_content['slide2'], company_name
        ),
        "slide3": generator.generate_slide3_content(
            slide_content['slide3'], company_name, sector
        ),
        "executive_summary": generator.generate_executive_summary(company_data, sector)
    }
    
    return generated


if __name__ == "__main__":
    # Test content generation
    print("Testing Content Generation...\n")
    
    # Check Ollama availability
    llm = OllamaInterface()
    print(f"Ollama available: {llm.is_available()}")
    
    if llm.is_available():
        # Test generation
        result = llm.generate("Say 'Hello, I am working!' in exactly those words.", temperature=0.1)
        print(f"Test response: {result}")
    
    # Test anonymization
    anonymizer = ContentAnonymizer()
    test_desc = "Kalyani Forge Ltd. is a leading Indian engineering company specializing in high-quality forged products."
    anon = anonymizer.anonymize_description(test_desc, "Kalyani Forge", "manufacturing")
    print(f"\nOriginal: {test_desc}")
    print(f"Anonymized: {anon}")
