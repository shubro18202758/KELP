"""
Image Intelligence Module - Sources and manages images for investment teasers
Uses free APIs (Unsplash, Pexels) to source sector-appropriate generic images
"""
import os
import re
import json
import hashlib
import urllib.request
from pathlib import Path
from typing import Dict, List, Optional, Tuple
from dataclasses import dataclass, field
import sys

sys.path.insert(0, str(Path(__file__).parent.parent.parent))

from config.settings import IMAGE_CONFIG, SECTOR_CONFIGS


@dataclass
class ImageResult:
    """Result from image search"""
    url: str
    thumbnail_url: str
    width: int
    height: int
    description: str
    photographer: str
    source: str  # 'unsplash', 'pexels', 'placeholder'
    download_url: Optional[str] = None


@dataclass
class SectorImagePlan:
    """Image plan for a sector"""
    sector: str
    queries: List[str]
    slide1_images: List[str] = field(default_factory=list)  # Queries for slide 1
    slide2_images: List[str] = field(default_factory=list)  # Queries for slide 2
    slide3_images: List[str] = field(default_factory=list)  # Queries for slide 3


class ImageQueryGenerator:
    """Generates sector-appropriate image search queries"""
    
    SECTOR_QUERIES = {
        "manufacturing": {
            "slide1": ["industrial manufacturing factory", "precision machinery production line", 
                      "metal forging plant equipment", "quality control laboratory"],
            "slide2": ["business growth chart abstract", "industrial dashboard analytics",
                      "factory automation technology"],
            "slide3": ["modern industrial facility exterior", "engineering innovation technology",
                      "global manufacturing network"]
        },
        "electronics": {
            "slide1": ["electronics manufacturing pcb", "aerospace technology components",
                      "defense electronics systems", "satellite technology clean room"],
            "slide2": ["technology growth metrics", "electronics testing equipment",
                      "circuit board production"],
            "slide3": ["high tech research facility", "innovation electronics lab",
                      "aerospace engineering center"]
        },
        "pharma": {
            "slide1": ["pharmaceutical manufacturing plant", "medicine production laboratory",
                      "drug research facility", "pharma quality testing"],
            "slide2": ["healthcare data analytics", "pharmaceutical research equipment",
                      "medicine packaging production"],
            "slide3": ["modern pharma research center", "healthcare innovation laboratory",
                      "pharmaceutical clean room facility"]
        },
        "technology": {
            "slide1": ["modern tech office workspace", "software development team",
                      "cloud computing data center", "digital transformation technology"],
            "slide2": ["technology dashboard analytics", "software metrics visualization",
                      "cloud infrastructure servers"],
            "slide3": ["innovative tech campus", "modern software company office",
                      "technology startup workspace"]
        },
        "logistics": {
            "slide1": ["logistics warehouse distribution", "delivery fleet trucks",
                      "supply chain operations center", "freight transportation hub"],
            "slide2": ["logistics tracking technology", "supply chain analytics dashboard",
                      "warehouse automation systems"],
            "slide3": ["modern logistics center aerial", "transportation network infrastructure",
                      "express delivery operations"]
        },
        "entertainment": {
            "slide1": ["modern cinema multiplex interior", "movie theater auditorium",
                      "entertainment venue design", "film exhibition technology"],
            "slide2": ["entertainment industry metrics", "cinema projection technology",
                      "movie theater experience"],
            "slide3": ["premium cinema experience", "entertainment complex exterior",
                      "modern movie theater design"]
        }
    }
    
    def get_queries_for_sector(self, sector: str) -> SectorImagePlan:
        """Get image queries for a sector"""
        sector_lower = sector.lower()
        
        # Find matching sector
        queries = self.SECTOR_QUERIES.get(sector_lower, self.SECTOR_QUERIES.get("manufacturing"))
        
        return SectorImagePlan(
            sector=sector,
            queries=list(set(sum(queries.values(), []))),
            slide1_images=queries.get("slide1", []),
            slide2_images=queries.get("slide2", []),
            slide3_images=queries.get("slide3", [])
        )


class UnsplashClient:
    """Client for Unsplash API"""
    
    BASE_URL = "https://api.unsplash.com"
    
    def __init__(self, access_key: str = None):
        self.access_key = access_key or os.environ.get('UNSPLASH_ACCESS_KEY')
        self.available = bool(self.access_key)
    
    def search(self, query: str, per_page: int = 3) -> List[ImageResult]:
        """Search for images"""
        if not self.available:
            return []
        
        try:
            url = f"{self.BASE_URL}/search/photos"
            params = f"query={urllib.parse.quote(query)}&per_page={per_page}&orientation=landscape"
            full_url = f"{url}?{params}"
            
            req = urllib.request.Request(full_url)
            req.add_header('Authorization', f'Client-ID {self.access_key}')
            
            with urllib.request.urlopen(req, timeout=10) as response:
                data = json.loads(response.read().decode())
                
            results = []
            for photo in data.get('results', []):
                results.append(ImageResult(
                    url=photo['urls']['regular'],
                    thumbnail_url=photo['urls']['thumb'],
                    width=photo['width'],
                    height=photo['height'],
                    description=photo.get('description') or photo.get('alt_description') or query,
                    photographer=photo['user']['name'],
                    source='unsplash',
                    download_url=photo['links']['download']
                ))
            
            return results
            
        except Exception as e:
            print(f"Unsplash search failed: {e}")
            return []


class PexelsClient:
    """Client for Pexels API"""
    
    BASE_URL = "https://api.pexels.com/v1"
    
    def __init__(self, api_key: str = None):
        self.api_key = api_key or os.environ.get('PEXELS_API_KEY')
        self.available = bool(self.api_key)
    
    def search(self, query: str, per_page: int = 3) -> List[ImageResult]:
        """Search for images"""
        if not self.available:
            return []
        
        try:
            url = f"{self.BASE_URL}/search"
            params = f"query={urllib.parse.quote(query)}&per_page={per_page}&orientation=landscape"
            full_url = f"{url}?{params}"
            
            req = urllib.request.Request(full_url)
            req.add_header('Authorization', self.api_key)
            
            with urllib.request.urlopen(req, timeout=10) as response:
                data = json.loads(response.read().decode())
                
            results = []
            for photo in data.get('photos', []):
                results.append(ImageResult(
                    url=photo['src']['large'],
                    thumbnail_url=photo['src']['medium'],
                    width=photo['width'],
                    height=photo['height'],
                    description=photo.get('alt') or query,
                    photographer=photo['photographer'],
                    source='pexels',
                    download_url=photo['src']['original']
                ))
            
            return results
            
        except Exception as e:
            print(f"Pexels search failed: {e}")
            return []


class PlaceholderImageGenerator:
    """Generates placeholder images when APIs are unavailable"""
    
    # Placeholder image URLs from placeholder services
    PLACEHOLDER_URLS = {
        "manufacturing": [
            "https://images.unsplash.com/photo-1565043666747-69f6646db940?w=800",  # Factory
            "https://images.unsplash.com/photo-1581091226825-a6a2a5aee158?w=800",  # Industrial
        ],
        "electronics": [
            "https://images.unsplash.com/photo-1518770660439-4636190af475?w=800",  # Circuit
            "https://images.unsplash.com/photo-1531297484001-80022131f5a1?w=800",  # Tech
        ],
        "pharma": [
            "https://images.unsplash.com/photo-1532187863486-abf9dbad1b69?w=800",  # Lab
            "https://images.unsplash.com/photo-1579154204601-01588f351e67?w=800",  # Medicine
        ],
        "technology": [
            "https://images.unsplash.com/photo-1451187580459-43490279c0fa?w=800",  # Tech abstract
            "https://images.unsplash.com/photo-1460925895917-afdab827c52f?w=800",  # Dashboard
        ],
        "logistics": [
            "https://images.unsplash.com/photo-1586528116311-ad8dd3c8310d?w=800",  # Warehouse
            "https://images.unsplash.com/photo-1553413077-190dd305871c?w=800",  # Delivery
        ],
        "entertainment": [
            "https://images.unsplash.com/photo-1489599849927-2ee91cede3ba?w=800",  # Cinema
            "https://images.unsplash.com/photo-1517604931442-7e0c8ed2963c?w=800",  # Theater
        ]
    }
    
    def get_placeholder(self, sector: str, index: int = 0) -> ImageResult:
        """Get a placeholder image for a sector"""
        sector_lower = sector.lower()
        urls = self.PLACEHOLDER_URLS.get(sector_lower, self.PLACEHOLDER_URLS["manufacturing"])
        url = urls[index % len(urls)]
        
        return ImageResult(
            url=url,
            thumbnail_url=url.replace("w=800", "w=400"),
            width=800,
            height=533,
            description=f"Stock image for {sector} sector",
            photographer="Stock Photo",
            source="placeholder"
        )


class ImageSourcer:
    """Main class for sourcing images"""
    
    def __init__(self):
        self.unsplash = UnsplashClient()
        self.pexels = PexelsClient()
        self.placeholder = PlaceholderImageGenerator()
        self.query_generator = ImageQueryGenerator()
        self._cache: Dict[str, List[ImageResult]] = {}
    
    def _get_cache_key(self, query: str) -> str:
        """Generate cache key for query"""
        return hashlib.md5(query.encode()).hexdigest()
    
    def search(self, query: str, count: int = 2) -> List[ImageResult]:
        """Search for images using available APIs"""
        cache_key = self._get_cache_key(query)
        
        if cache_key in self._cache:
            return self._cache[cache_key][:count]
        
        results = []
        
        # Try Unsplash first
        if self.unsplash.available:
            results = self.unsplash.search(query, count)
        
        # Try Pexels if Unsplash fails
        if not results and self.pexels.available:
            results = self.pexels.search(query, count)
        
        # Cache results
        if results:
            self._cache[cache_key] = results
        
        return results[:count]
    
    def get_sector_images(self, sector: str) -> Dict[str, List[ImageResult]]:
        """Get all images needed for a sector's slides"""
        plan = self.query_generator.get_queries_for_sector(sector)
        
        images = {
            "slide1": [],
            "slide2": [],
            "slide3": []
        }
        
        # Get images for each slide
        for slide_key, queries in [("slide1", plan.slide1_images), 
                                   ("slide2", plan.slide2_images),
                                   ("slide3", plan.slide3_images)]:
            for query in queries[:2]:  # Max 2 queries per slide
                results = self.search(query, count=1)
                if results:
                    images[slide_key].extend(results)
            
            # Use placeholders if no API results
            if not images[slide_key]:
                idx = {"slide1": 0, "slide2": 1, "slide3": 0}[slide_key]
                images[slide_key].append(self.placeholder.get_placeholder(sector, idx))
        
        return images
    
    def get_images_for_citations(self, images: Dict[str, List[ImageResult]]) -> List[Dict]:
        """Convert images to citation format"""
        citations = []
        
        for slide, img_list in images.items():
            for img in img_list:
                citations.append({
                    "description": img.description,
                    "url": img.url,
                    "source": f"{img.source.title()} - {img.photographer}",
                    "slide": slide.replace("slide", "Slide ")
                })
        
        return citations


def source_images_for_sector(sector: str) -> Tuple[Dict[str, List[ImageResult]], List[Dict]]:
    """
    Main function to source images for a sector.
    
    Returns:
        Tuple of (images dict, citations list)
    """
    sourcer = ImageSourcer()
    images = sourcer.get_sector_images(sector)
    citations = sourcer.get_images_for_citations(images)
    
    return images, citations


if __name__ == "__main__":
    # Test image sourcing
    print("Testing Image Intelligence Module...\n")
    
    # Test query generation
    generator = ImageQueryGenerator()
    plan = generator.get_queries_for_sector("manufacturing")
    print(f"Sector: {plan.sector}")
    print(f"Slide 1 queries: {plan.slide1_images}")
    print(f"Slide 2 queries: {plan.slide2_images}")
    print(f"Slide 3 queries: {plan.slide3_images}")
    
    # Test image sourcing (will use placeholders if no API keys)
    print("\nSourcing images...")
    images, citations = source_images_for_sector("manufacturing")
    
    for slide, img_list in images.items():
        print(f"\n{slide}: {len(img_list)} images")
        for img in img_list:
            print(f"  - {img.source}: {img.description[:50]}...")
    
    print(f"\n✓ Generated {len(citations)} image citations")
