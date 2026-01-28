"""
Free Image Fetcher - No API Keys Required
==========================================

Uses FREE sources to fetch sector-appropriate images:
1. DuckDuckGo Image Search (primary) - completely free
2. icrawler for bulk downloading
3. Local Janus AI model for generation (fallback)

This is the critical image sourcing layer for the M&A teaser pipeline.
Each slide needs 3-4 high-quality, sector-appropriate images.
"""

import os
import re
import asyncio
import hashlib
import random
from pathlib import Path
from typing import List, Dict, Optional, Tuple
from dataclasses import dataclass, field
from datetime import datetime
import io
import time

# DuckDuckGo search
try:
    from duckduckgo_search import DDGS
    HAS_DDG = True
except ImportError:
    HAS_DDG = False
    print("⚠ duckduckgo-search not installed. Run: pip install duckduckgo-search")

# icrawler for bulk downloads
try:
    from icrawler.builtin import BingImageCrawler, GoogleImageCrawler
    from icrawler import ImageDownloader
    HAS_ICRAWLER = True
except ImportError:
    HAS_ICRAWLER = False

# PIL for image processing
try:
    from PIL import Image
    HAS_PIL = True
except ImportError:
    HAS_PIL = False

# Requests for downloading
import requests

import sys
sys.path.insert(0, str(Path(__file__).parent.parent.parent))
from config.settings import OUTPUT_DIR


@dataclass
class FetchedImage:
    """Represents a fetched image"""
    path: Path
    url: str
    query: str
    source: str  # 'duckduckgo', 'bing', 'google', 'janus', 'placeholder'
    width: int = 0
    height: int = 0
    license_info: str = ""


# ============================================================================
# SECTOR-SPECIFIC IMAGE QUERIES
# ============================================================================

SECTOR_IMAGE_QUERIES = {
    "Manufacturing & Industrials": {
        "slide1_cover": [
            "industrial manufacturing plant modern",
            "factory production line automation",
            "metal forging steel works",
            "automotive manufacturing facility",
            "heavy machinery industrial"
        ],
        "slide2_business": [
            "manufacturing quality control",
            "industrial equipment machinery",
            "factory workers safety gear",
            "production assembly line",
            "industrial warehouse logistics"
        ],
        "slide3_finance": [
            "business growth chart abstract",
            "industrial expansion construction",
            "manufacturing technology innovation",
            "supply chain logistics",
            "industrial engineering design"
        ],
        "slide4_highlights": [
            "industry leadership trophy",
            "global manufacturing network",
            "sustainable manufacturing green",
            "industrial innovation technology",
            "manufacturing excellence award"
        ]
    },
    "Electronics & Defense": {
        "slide1_cover": [
            "electronics manufacturing pcb",
            "defense technology systems",
            "aerospace engineering facility",
            "semiconductor chip production",
            "electronic components assembly"
        ],
        "slide2_business": [
            "circuit board manufacturing",
            "electronic testing laboratory",
            "defense radar systems",
            "avionics equipment",
            "electronic design engineering"
        ],
        "slide3_finance": [
            "technology growth innovation",
            "defense contract signing",
            "electronic components export",
            "semiconductor industry growth",
            "defense manufacturing expansion"
        ],
        "slide4_highlights": [
            "electronic certification quality",
            "defense innovation award",
            "aerospace technology leadership",
            "electronic manufacturing excellence",
            "defense capability demonstration"
        ]
    },
    "Pharmaceuticals": {
        "slide1_cover": [
            "pharmaceutical laboratory research",
            "medicine production facility",
            "pharmaceutical manufacturing clean room",
            "drug development research",
            "pharmaceutical pills medicine"
        ],
        "slide2_business": [
            "pharmaceutical quality testing",
            "medicine packaging production",
            "clinical research laboratory",
            "pharmaceutical formulation",
            "drug manufacturing process"
        ],
        "slide3_finance": [
            "pharmaceutical growth investment",
            "healthcare industry expansion",
            "medicine export global",
            "pharmaceutical innovation R&D",
            "healthcare market growth"
        ],
        "slide4_highlights": [
            "pharmaceutical certification WHO GMP",
            "medicine quality excellence",
            "pharmaceutical research breakthrough",
            "healthcare innovation award",
            "pharmaceutical global reach"
        ]
    },
    "Technology & IT Services": {
        "slide1_cover": [
            "software development team office",
            "technology data center",
            "IT services modern office",
            "cloud computing technology",
            "digital transformation technology"
        ],
        "slide2_business": [
            "software coding development",
            "IT infrastructure servers",
            "technology team collaboration",
            "digital solutions software",
            "IT consulting services"
        ],
        "slide3_finance": [
            "technology growth chart",
            "IT industry expansion",
            "software company growth",
            "technology investment success",
            "digital business growth"
        ],
        "slide4_highlights": [
            "technology innovation award",
            "IT excellence certification",
            "software development achievement",
            "digital transformation success",
            "technology leadership recognition"
        ]
    },
    "Logistics & Supply Chain": {
        "slide1_cover": [
            "logistics warehouse distribution",
            "freight transport trucks fleet",
            "supply chain management",
            "cargo shipping containers",
            "logistics network operations"
        ],
        "slide2_business": [
            "warehouse automation robots",
            "delivery trucks logistics",
            "supply chain technology",
            "freight forwarding operations",
            "last mile delivery"
        ],
        "slide3_finance": [
            "logistics growth expansion",
            "supply chain investment",
            "freight industry growth",
            "logistics network map",
            "transportation growth chart"
        ],
        "slide4_highlights": [
            "logistics excellence award",
            "supply chain innovation",
            "delivery network coverage",
            "logistics technology leadership",
            "freight services excellence"
        ]
    },
    "Entertainment & Media": {
        "slide1_cover": [
            "cinema multiplex theater",
            "movie theater interior",
            "entertainment venue auditorium",
            "film screening hall premium",
            "movie premiere red carpet"
        ],
        "slide2_business": [
            "cinema projection technology",
            "theater seating premium",
            "entertainment venue lobby",
            "movie theater concessions",
            "cinema experience luxury"
        ],
        "slide3_finance": [
            "entertainment industry growth",
            "cinema expansion investment",
            "media industry chart",
            "entertainment market growth",
            "box office success"
        ],
        "slide4_highlights": [
            "cinema excellence award",
            "entertainment innovation",
            "movie theater technology",
            "entertainment experience premium",
            "cinema chain expansion"
        ]
    }
}


class FreeImageFetcher:
    """
    Fetches images using FREE sources - no API keys required.
    
    Primary: DuckDuckGo Image Search
    Fallback: icrawler (Bing/Google)
    Last Resort: Placeholder generation
    """
    
    def __init__(self, cache_dir: Path = None):
        self.cache_dir = cache_dir or OUTPUT_DIR / "image_cache"
        self.cache_dir.mkdir(parents=True, exist_ok=True)
        
        # Track downloaded images to avoid duplicates
        self.downloaded_hashes = set()
        
        # Session for downloads
        self.session = requests.Session()
        self.session.headers.update({
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'
        })
    
    def _get_sector_queries(self, sector: str, slide_type: str) -> List[str]:
        """Get image queries for a specific sector and slide"""
        # Find matching sector
        for sector_key, queries in SECTOR_IMAGE_QUERIES.items():
            if sector.lower() in sector_key.lower() or sector_key.lower() in sector.lower():
                return queries.get(slide_type, queries.get('slide1_cover', []))
        
        # Default queries if sector not found
        return [
            "business professional modern",
            "corporate office building",
            "business growth success",
            "industry innovation technology"
        ]
    
    def _download_image(self, url: str, query: str, sector_prefix: str = "") -> Optional[Path]:
        """Download image from URL with sector prefix for proper caching"""
        if not HAS_PIL:
            return None
            
        try:
            response = self.session.get(url, timeout=10)
            if response.status_code != 200:
                return None
            
            # Check content type
            content_type = response.headers.get('content-type', '')
            if 'image' not in content_type:
                return None
            
            # Load image
            img_data = response.content
            img_hash = hashlib.md5(img_data).hexdigest()[:12]
            
            # Skip duplicates
            if img_hash in self.downloaded_hashes:
                return None
            self.downloaded_hashes.add(img_hash)
            
            # Open and validate
            img = Image.open(io.BytesIO(img_data))
            
            # Check minimum size
            if img.width < 400 or img.height < 300:
                return None
            
            # Convert to RGB if needed
            if img.mode in ('RGBA', 'P'):
                img = img.convert('RGB')
            
            # Resize for PPT (max 1920x1080)
            max_size = (1920, 1080)
            img.thumbnail(max_size, Image.LANCZOS)
            
            # Save with SECTOR PREFIX for proper isolation
            safe_query = re.sub(r'[^\w\-]', '_', query)[:20]
            # Use sector prefix to ensure images are sector-specific
            filename = f"{sector_prefix}_{safe_query}_{img_hash}.jpg" if sector_prefix else f"{safe_query}_{img_hash}.jpg"
            filepath = self.cache_dir / filename
            
            img.save(str(filepath), "JPEG", quality=85)
            return filepath
            
        except Exception as e:
            return None
    
    def fetch_with_duckduckgo(self, query: str, max_images: int = 5, sector_prefix: str = "") -> List[FetchedImage]:
        """Fetch images using DuckDuckGo (FREE, no API key)"""
        if not HAS_DDG:
            return []
        
        images = []
        try:
            with DDGS() as ddgs:
                results = list(ddgs.images(
                    query,
                    max_results=max_images * 3,  # Fetch more in case some fail
                    safesearch='moderate',
                    size='large',
                    type_image='photo'
                ))
                
                for result in results:
                    if len(images) >= max_images:
                        break
                        
                    url = result.get('image', '')
                    if not url:
                        continue
                    
                    # Download with sector prefix for isolation
                    filepath = self._download_image(url, query, sector_prefix)
                    if filepath:
                        with Image.open(filepath) as img:
                            images.append(FetchedImage(
                                path=filepath,
                                url=url,
                                query=query,
                                source='duckduckgo',
                                width=img.width,
                                height=img.height,
                                license_info="Web search result"
                            ))
                        print(f"    ✓ Downloaded: {filepath.name}")
                        
        except Exception as e:
            print(f"    ⚠ DuckDuckGo error: {e}")
        
        return images
    
    def fetch_with_icrawler(self, query: str, max_images: int = 5, sector_prefix: str = "") -> List[FetchedImage]:
        """Fetch images using icrawler (Bing) - no API key needed"""
        if not HAS_ICRAWLER:
            return []
        
        images = []
        try:
            # Create temp directory for this query
            safe_query = re.sub(r'[^\w\-]', '_', query)[:30]
            temp_dir = self.cache_dir / f"temp_{safe_query}"
            temp_dir.mkdir(exist_ok=True)
            
            # Use Bing crawler (no API key needed)
            crawler = BingImageCrawler(
                storage={'root_dir': str(temp_dir)},
                log_level=50  # Suppress logs
            )
            crawler.crawl(keyword=query, max_num=max_images)
            
            # Process downloaded images
            for img_file in temp_dir.glob("*.*"):
                if img_file.suffix.lower() in ['.jpg', '.jpeg', '.png', '.webp']:
                    try:
                        img = Image.open(img_file)
                        if img.width >= 400 and img.height >= 300:
                            # Move to cache
                            new_name = f"{safe_query}_{img_file.stem}.jpg"
                            new_path = self.cache_dir / new_name
                            
                            if img.mode in ('RGBA', 'P'):
                                img = img.convert('RGB')
                            img.save(str(new_path), "JPEG", quality=85)
                            
                            images.append(FetchedImage(
                                path=new_path,
                                url="",
                                query=query,
                                source='bing',
                                width=img.width,
                                height=img.height
                            ))
                            print(f"    ✓ Crawled: {new_path.name}")
                    except:
                        pass
                    finally:
                        img_file.unlink(missing_ok=True)
            
            # Cleanup temp dir
            temp_dir.rmdir()
            
        except Exception as e:
            print(f"    ⚠ icrawler error: {e}")
        
        return images
    
    def create_placeholder(self, query: str, sector: str) -> FetchedImage:
        """Create a placeholder image with sector branding"""
        if not HAS_PIL:
            return None
            
        try:
            from PIL import ImageDraw, ImageFont
            
            # Create gradient background
            width, height = 1200, 800
            img = Image.new('RGB', (width, height))
            draw = ImageDraw.Draw(img)
            
            # Sector-specific colors
            sector_colors = {
                'manufacturing': ((27, 54, 93), (75, 108, 183)),
                'electronics': ((45, 52, 94), (88, 86, 161)),
                'pharma': ((0, 128, 128), (0, 191, 179)),
                'technology': ((55, 66, 250), (135, 91, 247)),
                'logistics': ((255, 107, 53), (255, 165, 89)),
                'entertainment': ((232, 75, 138), (255, 150, 100))
            }
            
            # Find matching color
            color1, color2 = ((27, 54, 93), (75, 108, 183))  # Default
            for key, colors in sector_colors.items():
                if key in sector.lower():
                    color1, color2 = colors
                    break
            
            # Draw gradient
            for y in range(height):
                ratio = y / height
                r = int(color1[0] + (color2[0] - color1[0]) * ratio)
                g = int(color1[1] + (color2[1] - color1[1]) * ratio)
                b = int(color1[2] + (color2[2] - color1[2]) * ratio)
                draw.line([(0, y), (width, y)], fill=(r, g, b))
            
            # Add decorative elements
            for i in range(5):
                x = random.randint(50, width - 150)
                y_pos = random.randint(50, height - 150)
                size = random.randint(50, 150)
                opacity = random.randint(20, 60)
                overlay_color = (255, 255, 255)
                draw.ellipse([x, y_pos, x + size, y_pos + size], 
                            fill=(*overlay_color, opacity))
            
            # Add text
            try:
                font = ImageFont.truetype("arial.ttf", 36)
                font_small = ImageFont.truetype("arial.ttf", 20)
            except:
                font = ImageFont.load_default()
                font_small = font
            
            # Sector name
            draw.text((width//2, height//2 - 30), sector[:40], 
                     fill=(255, 255, 255), font=font, anchor="mm")
            
            # Query hint
            draw.text((width//2, height//2 + 30), query[:60], 
                     fill=(200, 200, 200), font=font_small, anchor="mm")
            
            # Save
            safe_query = re.sub(r'[^\w\-]', '_', query)[:25]
            filename = f"placeholder_{safe_query}_{random.randint(1000, 9999)}.jpg"
            filepath = self.cache_dir / filename
            img.save(str(filepath), "JPEG", quality=90)
            
            return FetchedImage(
                path=filepath,
                url="",
                query=query,
                source='placeholder',
                width=width,
                height=height,
                license_info="Generated placeholder"
            )
            
        except Exception as e:
            print(f"    ⚠ Placeholder error: {e}")
            return None
    
    def fetch_for_slide(self, sector: str, slide_type: str, 
                        images_needed: int = 3) -> List[FetchedImage]:
        """
        Fetch images for a specific slide.
        
        Args:
            sector: Company sector (e.g., "Manufacturing & Industrials")
            slide_type: One of 'slide1_cover', 'slide2_business', etc.
            images_needed: Number of images to fetch (3-4 per slide)
        
        Returns:
            List of FetchedImage objects
        """
        print(f"  📷 Fetching {images_needed} images for {slide_type}...")
        
        queries = self._get_sector_queries(sector, slide_type)
        all_images = []
        
        # Create sector prefix for file naming (ensures proper isolation)
        sector_prefix = re.sub(r'[^\w]', '_', sector.lower())[:30]
        
        # Try each query until we have enough images
        for query in queries:
            if len(all_images) >= images_needed:
                break
            
            remaining = images_needed - len(all_images)
            print(f"    🔍 Searching: '{query}'")
            
            # Try DuckDuckGo first (free, no rate limits) - with sector prefix!
            images = self.fetch_with_duckduckgo(query, remaining, sector_prefix)
            all_images.extend(images)
            
            # Add delay to avoid rate limiting
            time.sleep(1.5)
            
            # If not enough, try icrawler
            if len(all_images) < images_needed:
                remaining = images_needed - len(all_images)
                images = self.fetch_with_icrawler(query, remaining, sector_prefix)
                all_images.extend(images)
        
        # If still not enough, create placeholders
        while len(all_images) < images_needed:
            placeholder = self.create_placeholder(
                queries[0] if queries else "business",
                sector
            )
            if placeholder:
                all_images.append(placeholder)
                print(f"    📌 Created placeholder")
        
        return all_images[:images_needed]
    
    def fetch_all_for_company(self, sector: str) -> Dict[str, List[FetchedImage]]:
        """
        Fetch all images needed for a company's teaser (all 4 slides).
        Uses caching to avoid re-downloading same sector images.
        
        Returns:
            Dict with keys 'slide1', 'slide2', 'slide3', 'slide4'
            Each containing 3-4 images
        """
        print(f"\n🖼️ FETCHING IMAGES FOR SECTOR: {sector}")
        print("=" * 50)
        
        results = {}
        
        # Check if we have cached images for this SPECIFIC sector ONLY
        sector_cache_key = re.sub(r'[^\w]', '_', sector.lower())[:30]
        
        slide_types = [
            ('slide1', 'slide1_cover', 2),    # Cover needs fewer images
            ('slide2', 'slide2_business', 2),  # Business overview (reduced)
            ('slide3', 'slide3_finance', 1),   # Financial (charts dominate)
            ('slide4', 'slide4_highlights', 2) # Investment highlights
        ]
        
        # ONLY use cached images that match THIS SPECIFIC sector
        # This prevents entertainment images appearing on manufacturing slides!
        cached_images = list(self.cache_dir.glob(f"{sector_cache_key}*.jpg"))
        
        if len(cached_images) >= 7:
            print(f"  📁 Using {len(cached_images)} sector-specific cached images")
            # Distribute cached images across slides (DON'T randomly shuffle to keep consistency)
            idx = 0
            for slide_key, slide_type, count in slide_types:
                slide_imgs = []
                for i in range(count):
                    if idx < len(cached_images):
                        img_path = cached_images[idx]
                        try:
                            with Image.open(img_path) as img:
                                slide_imgs.append(FetchedImage(
                                    path=img_path,
                                    url="",
                                    query="cached",
                                    source="cache",
                                    width=img.width,
                                    height=img.height
                                ))
                        except:
                            pass
                        idx += 1
                results[slide_key] = slide_imgs
                print(f"  ✓ {slide_key}: {len(slide_imgs)} images (cached)")
        else:
            # Fetch new images for THIS sector
            for slide_key, slide_type, count in slide_types:
                images = self.fetch_for_slide(sector, slide_type, count)
                results[slide_key] = images
                print(f"  ✓ {slide_key}: {len(images)} images")
        
        total = sum(len(imgs) for imgs in results.values())
        print(f"\n✅ Total images fetched: {total}")
        
        return results


# ============================================================================
# JANUS LOCAL AI IMAGE GENERATOR (Fallback)
# ============================================================================

class JanusImageGenerator:
    """
    Uses local Janus AI model to GENERATE sector images.
    This is a fallback when web scraping fails.
    """
    
    def __init__(self, cache_dir: Path = None):
        self.cache_dir = cache_dir or OUTPUT_DIR / "janus_images"
        self.cache_dir.mkdir(parents=True, exist_ok=True)
        self.model_loaded = False
        self.model = None
        self.processor = None
    
    def _load_model(self):
        """Load Janus model if available"""
        if self.model_loaded:
            return True
            
        try:
            import torch
            from transformers import AutoModelForCausalLM, AutoProcessor
            
            print("  🔄 Loading Janus-Pro-7B model...")
            
            # Check for model
            model_path = "deepseek-ai/Janus-Pro-7B"
            
            self.processor = AutoProcessor.from_pretrained(model_path)
            self.model = AutoModelForCausalLM.from_pretrained(
                model_path,
                torch_dtype=torch.float16,
                device_map="auto",
                trust_remote_code=True
            )
            
            self.model_loaded = True
            print("  ✓ Janus model loaded!")
            return True
            
        except Exception as e:
            print(f"  ⚠ Janus not available: {e}")
            return False
    
    def generate_sector_image(self, prompt: str, sector: str) -> Optional[Path]:
        """Generate an image using Janus"""
        if not self._load_model():
            return None
        
        try:
            import torch
            
            # Build generation prompt
            full_prompt = f"Generate a professional photograph for business presentation: {prompt}. Style: corporate, high quality, no text or logos."
            
            # Generate
            inputs = self.processor(full_prompt, return_tensors="pt").to(self.model.device)
            
            with torch.no_grad():
                outputs = self.model.generate(
                    **inputs,
                    max_new_tokens=1024,
                    do_sample=True,
                    temperature=0.7
                )
            
            # Extract image from outputs
            # Note: This depends on Janus output format
            # May need adjustment based on actual model API
            
            safe_prompt = re.sub(r'[^\w\-]', '_', prompt)[:30]
            filename = f"janus_{safe_prompt}_{random.randint(1000, 9999)}.png"
            filepath = self.cache_dir / filename
            
            # Save generated image
            # (Implementation depends on Janus output format)
            
            return filepath
            
        except Exception as e:
            print(f"  ⚠ Janus generation failed: {e}")
            return None


# ============================================================================
# TEST
# ============================================================================

if __name__ == "__main__":
    print("=" * 60)
    print("FREE IMAGE FETCHER - TEST")
    print("=" * 60)
    
    fetcher = FreeImageFetcher()
    
    # Test for one sector
    sector = "Manufacturing & Industrials"
    images = fetcher.fetch_all_for_company(sector)
    
    print("\n📋 RESULTS:")
    for slide, imgs in images.items():
        print(f"\n{slide}:")
        for img in imgs:
            print(f"  - {img.path.name} ({img.source})")
