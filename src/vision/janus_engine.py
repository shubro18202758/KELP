"""
DeepSeek Janus-Pro-7B Vision-Language Engine
=============================================
Multimodal model for image understanding AND generation.
Used for:
- Generating sector-relevant images for PPT slides
- Analyzing web screenshots
- Creating charts and infographics
"""
import os
import torch
import base64
import io
import random
from pathlib import Path
from typing import Dict, List, Any, Optional, Tuple, Union
from dataclasses import dataclass, field
from datetime import datetime
from PIL import Image
import sys

sys.path.insert(0, str(Path(__file__).parent.parent.parent))
from config.settings import OUTPUT_DIR


@dataclass
class JanusConfig:
    """Configuration for Janus-Pro-7B"""
    model_name: str = "deepseek-ai/Janus-Pro-7B"
    device: str = "cuda" if torch.cuda.is_available() else "cpu"
    torch_dtype: torch.dtype = torch.bfloat16
    max_new_tokens: int = 512
    temperature: float = 0.7
    
    # Image generation settings
    image_size: int = 384  # Janus generates 384x384 images
    cfg_weight: float = 5.0
    num_inference_steps: int = 25
    
    # GPU optimization
    use_flash_attention: bool = True
    offload_to_cpu: bool = False


class JanusProEngine:
    """
    DeepSeek Janus-Pro-7B Engine for multimodal tasks.
    
    Capabilities:
    - Image understanding (analyze images, screenshots)
    - Image generation (create sector-relevant visuals)
    - Text generation with visual context
    """
    
    def __init__(self, config: JanusConfig = None):
        self.config = config or JanusConfig()
        self.model = None
        self.processor = None
        self.tokenizer = None
        self.image_gen_processor = None
        self._initialized = False
        
        # Output directories
        self.image_output_dir = OUTPUT_DIR / "generated_images"
        self.image_output_dir.mkdir(parents=True, exist_ok=True)
        
    def _initialize(self):
        """Lazy initialization of model"""
        if self._initialized:
            return
        
        print("   🔄 Loading DeepSeek Janus-Pro-7B...")
        
        try:
            from transformers import AutoModelForCausalLM, AutoConfig
            from janus.models import MultiModalityCausalLM, VLChatProcessor
            
            # Load model configuration
            config = AutoConfig.from_pretrained(self.config.model_name)
            
            # Check for flash attention
            if self.config.use_flash_attention:
                try:
                    config._attn_implementation = "flash_attention_2"
                except:
                    pass
            
            # Load model with optimizations
            self.model = AutoModelForCausalLM.from_pretrained(
                self.config.model_name,
                config=config,
                torch_dtype=self.config.torch_dtype,
                trust_remote_code=True,
                device_map="auto" if self.config.device == "cuda" else None,
            )
            
            # Load processors
            self.processor = VLChatProcessor.from_pretrained(self.config.model_name)
            self.tokenizer = self.processor.tokenizer
            
            # For image generation
            from janus.models import VLMImageProcessor
            self.image_gen_processor = VLMImageProcessor.from_pretrained(self.config.model_name)
            
            if self.config.device == "cuda" and not self.config.offload_to_cpu:
                self.model = self.model.to(self.config.device)
            
            self.model.eval()
            self._initialized = True
            print("   ✓ Janus-Pro-7B loaded successfully")
            
        except ImportError as e:
            print(f"   ⚠ Janus dependencies not installed: {e}")
            print("   📦 Install with: pip install janus-pro")
            self._initialized = False
        except Exception as e:
            print(f"   ⚠ Failed to load Janus: {e}")
            self._initialized = False
    
    def is_available(self) -> bool:
        """Check if model is available"""
        if not self._initialized:
            try:
                self._initialize()
            except:
                pass
        return self._initialized
    
    def generate_sector_image(self, sector: str, image_type: str = "generic",
                             seed: int = None) -> Optional[Path]:
        """
        Generate a sector-relevant image for PPT slides.
        
        Args:
            sector: Industry sector (manufacturing, pharma, tech, etc.)
            image_type: Type of image (product, facility, abstract, chart)
            seed: Random seed for reproducibility
            
        Returns:
            Path to generated image or None
        """
        if not self.is_available():
            return self._generate_placeholder_image(sector, image_type)
        
        # Build prompt based on sector and type
        prompts = self._get_image_prompts(sector, image_type)
        prompt = random.choice(prompts) if isinstance(prompts, list) else prompts
        
        try:
            # Set seed for reproducibility if provided
            if seed:
                torch.manual_seed(seed)
            
            # Generate image using Janus
            with torch.no_grad():
                # Prepare conversation for image generation
                conversation = [
                    {"role": "user", "content": f"Generate an image: {prompt}"}
                ]
                
                # Process input
                inputs = self.processor(
                    conversations=conversation,
                    images=None,
                    force_batchify=True,
                    return_tensors="pt"
                ).to(self.config.device)
                
                # Generate
                outputs = self.model.generate(
                    **inputs,
                    max_new_tokens=self.config.image_size * self.config.image_size,
                    temperature=self.config.temperature,
                    do_sample=True,
                )
                
                # Decode image from tokens
                image = self.image_gen_processor.decode(
                    outputs[0],
                    image_size=self.config.image_size
                )
                
            # Save image
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = f"{sector}_{image_type}_{timestamp}.png"
            output_path = self.image_output_dir / filename
            image.save(output_path)
            
            print(f"   🖼 Generated: {filename}")
            return output_path
            
        except Exception as e:
            print(f"   ⚠ Image generation failed: {e}")
            return self._generate_placeholder_image(sector, image_type)
    
    def _get_image_prompts(self, sector: str, image_type: str) -> List[str]:
        """Get appropriate prompts for sector and image type"""
        
        sector_prompts = {
            "manufacturing": {
                "product": [
                    "Professional photograph of precision engineered metal components on white background, industrial quality, no logos",
                    "Modern manufacturing equipment in clean factory setting, professional lighting, no visible branding",
                    "High-quality machined parts arrangement, industrial photography style, neutral background",
                ],
                "facility": [
                    "Modern manufacturing plant interior with automated machinery, bright lighting, professional photograph",
                    "Clean industrial workspace with robotic arms, modern factory aesthetic, no logos visible",
                    "Aerial view of modern industrial complex, professional photography, generic design",
                ],
                "abstract": [
                    "Abstract representation of manufacturing precision, geometric shapes in blue and silver, professional design",
                    "Modern industrial infographic background, clean lines, professional color scheme",
                ]
            },
            "pharmaceuticals": {
                "product": [
                    "Pharmaceutical laboratory with modern equipment, clean room aesthetic, professional photograph",
                    "Medicine capsules and tablets arrangement on white background, pharmaceutical quality",
                    "Scientific research equipment in modern lab setting, no branding visible",
                ],
                "facility": [
                    "Modern pharmaceutical research facility interior, clean and bright, professional",
                    "GMP certified production line, clean room environment, generic design",
                ],
                "abstract": [
                    "Abstract molecular structure visualization, pharmaceutical blue tones, professional design",
                    "DNA helix with pharmaceutical elements, modern scientific aesthetic",
                ]
            },
            "technology": {
                "product": [
                    "Modern server room with blue LED lighting, technology aesthetic, professional",
                    "Software development workspace, multiple monitors with code, modern office",
                    "Cloud computing visualization, abstract data center concept",
                ],
                "facility": [
                    "Modern tech office space with collaborative areas, bright and innovative design",
                    "Data center interior with server racks, blue lighting, professional photograph",
                ],
                "abstract": [
                    "Abstract network connectivity visualization, blue nodes and connections",
                    "Digital transformation concept, modern technology infographic style",
                ]
            },
            "logistics": {
                "product": [
                    "Modern logistics warehouse interior, organized shelving, professional lighting",
                    "Fleet of delivery trucks, aerial view, professional photograph, no logos",
                    "Automated sorting facility, conveyor systems, modern logistics",
                ],
                "facility": [
                    "Large distribution center aerial view, modern logistics hub design",
                    "Supply chain visualization, warehouse to delivery network",
                ],
                "abstract": [
                    "Abstract supply chain network, connected nodes and routes, professional design",
                    "Global logistics map visualization, modern infographic style",
                ]
            },
            "electronics": {
                "product": [
                    "Printed circuit boards arrangement, electronic components, professional macro photography",
                    "Modern electronic assembly, semiconductor chips, clean room environment",
                    "Aerospace electronics modules, high-precision components, professional",
                ],
                "facility": [
                    "Electronics manufacturing clean room, automated assembly lines",
                    "R&D laboratory with testing equipment, modern electronics facility",
                ],
                "abstract": [
                    "Abstract circuit pattern, electronic pathways, blue and gold tones",
                    "Technology innovation concept, electronic components in modern design",
                ]
            },
            "entertainment": {
                "product": [
                    "Modern cinema interior, comfortable seating, ambient lighting",
                    "Movie theater projection room, professional equipment",
                    "Entertainment venue lobby, modern design aesthetic",
                ],
                "facility": [
                    "Multiplex cinema exterior, modern architecture, evening lighting",
                    "Entertainment complex aerial view, contemporary design",
                ],
                "abstract": [
                    "Abstract entertainment concept, film reel and digital elements",
                    "Movie and media visualization, creative industry design",
                ]
            }
        }
        
        # Get sector-specific prompts or default
        sector_key = sector.lower().split()[0]  # Take first word
        sector_data = sector_prompts.get(sector_key, sector_prompts.get("manufacturing"))
        
        return sector_data.get(image_type, sector_data.get("abstract", ["Professional business image"]))
    
    def _generate_placeholder_image(self, sector: str, image_type: str) -> Path:
        """Generate a placeholder image when model not available"""
        from PIL import Image, ImageDraw, ImageFont
        
        # Create gradient background
        width, height = 800, 600
        image = Image.new('RGB', (width, height))
        draw = ImageDraw.Draw(image)
        
        # Sector colors
        sector_colors = {
            "manufacturing": [(45, 35, 75), (0, 191, 179)],
            "pharmaceuticals": [(45, 35, 75), (126, 87, 194)],
            "technology": [(45, 35, 75), (0, 150, 255)],
            "logistics": [(45, 35, 75), (254, 185, 95)],
            "electronics": [(45, 35, 75), (232, 77, 138)],
            "entertainment": [(45, 35, 75), (255, 105, 135)],
        }
        
        colors = sector_colors.get(sector.lower().split()[0], [(45, 35, 75), (0, 191, 179)])
        
        # Create gradient
        for y in range(height):
            r = int(colors[0][0] + (colors[1][0] - colors[0][0]) * y / height)
            g = int(colors[0][1] + (colors[1][1] - colors[0][1]) * y / height)
            b = int(colors[0][2] + (colors[1][2] - colors[0][2]) * y / height)
            draw.line([(0, y), (width, y)], fill=(r, g, b))
        
        # Add sector text
        try:
            font = ImageFont.truetype("arial.ttf", 40)
            small_font = ImageFont.truetype("arial.ttf", 20)
        except:
            font = ImageFont.load_default()
            small_font = font
        
        sector_text = sector.upper()
        type_text = f"[ {image_type.upper()} ]"
        
        # Center text
        draw.text((width//2, height//2 - 30), sector_text, fill=(255, 255, 255), 
                  anchor="mm", font=font)
        draw.text((width//2, height//2 + 30), type_text, fill=(200, 200, 200),
                  anchor="mm", font=small_font)
        
        # Save
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"{sector.split()[0].lower()}_{image_type}_{timestamp}.png"
        output_path = self.image_output_dir / filename
        image.save(output_path)
        
        return output_path
    
    def analyze_image(self, image_path: Union[str, Path]) -> str:
        """
        Analyze an image and generate description.
        
        Args:
            image_path: Path to image file
            
        Returns:
            Text description of image
        """
        if not self.is_available():
            return "Image analysis not available - model not loaded"
        
        try:
            image = Image.open(image_path).convert("RGB")
            
            conversation = [
                {
                    "role": "user",
                    "content": [
                        {"type": "image", "image": image},
                        {"type": "text", "text": "Describe this image in detail, focusing on business-relevant elements."}
                    ]
                }
            ]
            
            inputs = self.processor(
                conversations=conversation,
                images=[image],
                force_batchify=True,
                return_tensors="pt"
            ).to(self.config.device)
            
            with torch.no_grad():
                outputs = self.model.generate(
                    **inputs,
                    max_new_tokens=self.config.max_new_tokens,
                    temperature=0.3,
                    do_sample=False,
                )
            
            response = self.tokenizer.decode(outputs[0], skip_special_tokens=True)
            return response
            
        except Exception as e:
            return f"Analysis failed: {e}"
    
    def generate_chart_image(self, chart_type: str, data: Dict, 
                            title: str = "") -> Optional[Path]:
        """
        Generate a chart/graph image.
        
        Args:
            chart_type: bar, pie, line, etc.
            data: Chart data
            title: Chart title
            
        Returns:
            Path to generated chart image
        """
        # Use matplotlib for reliable chart generation
        try:
            import matplotlib.pyplot as plt
            import matplotlib
            matplotlib.use('Agg')
            
            # Kelp color palette
            colors = ['#00BFB3', '#E84D8A', '#FEB95F', '#7B68EE', '#4ECDC4']
            
            fig, ax = plt.subplots(figsize=(10, 6), facecolor='white')
            
            if chart_type == "bar":
                categories = data.get('categories', ['A', 'B', 'C'])
                values = data.get('values', [10, 20, 30])
                bars = ax.bar(categories, values, color=colors[:len(categories)])
                ax.set_ylabel(data.get('ylabel', 'Value'))
                
            elif chart_type == "pie":
                labels = data.get('labels', ['A', 'B', 'C'])
                sizes = data.get('values', [30, 40, 30])
                ax.pie(sizes, labels=labels, colors=colors[:len(labels)], 
                      autopct='%1.0f%%', startangle=90)
                ax.axis('equal')
                
            elif chart_type == "line":
                x = data.get('x', range(5))
                y = data.get('y', [10, 15, 13, 18, 22])
                ax.plot(x, y, color=colors[0], linewidth=3, marker='o')
                ax.fill_between(x, y, alpha=0.3, color=colors[0])
                ax.set_xlabel(data.get('xlabel', ''))
                ax.set_ylabel(data.get('ylabel', ''))
            
            if title:
                ax.set_title(title, fontsize=14, fontweight='bold', color='#2D234B')
            
            ax.spines['top'].set_visible(False)
            ax.spines['right'].set_visible(False)
            
            # Save
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = f"chart_{chart_type}_{timestamp}.png"
            output_path = self.image_output_dir / filename
            plt.savefig(output_path, dpi=150, bbox_inches='tight', facecolor='white')
            plt.close()
            
            return output_path
            
        except Exception as e:
            print(f"   ⚠ Chart generation failed: {e}")
            return None


# Singleton instance
_janus_engine = None

def get_janus_engine() -> JanusProEngine:
    """Get or create Janus engine singleton"""
    global _janus_engine
    if _janus_engine is None:
        _janus_engine = JanusProEngine()
    return _janus_engine


def generate_sector_images(sector: str, count: int = 3) -> List[Path]:
    """
    Generate multiple sector-relevant images.
    
    Args:
        sector: Industry sector
        count: Number of images to generate
        
    Returns:
        List of paths to generated images
    """
    engine = get_janus_engine()
    images = []
    
    image_types = ["product", "facility", "abstract"]
    
    for i in range(min(count, len(image_types))):
        image_path = engine.generate_sector_image(sector, image_types[i])
        if image_path:
            images.append(image_path)
    
    return images


if __name__ == "__main__":
    print("Testing DeepSeek Janus-Pro Engine...")
    print("=" * 50)
    
    engine = get_janus_engine()
    
    # Test placeholder generation (always works)
    print("\nGenerating placeholder images...")
    for sector in ["Manufacturing", "Pharmaceuticals", "Technology"]:
        path = engine._generate_placeholder_image(sector, "product")
        print(f"  {sector}: {path}")
    
    # Test chart generation
    print("\nGenerating charts...")
    chart_path = engine.generate_chart_image(
        "bar",
        {"categories": ["FY22", "FY23", "FY24", "FY25"], "values": [120, 145, 180, 220]},
        "Revenue Growth (₹ Cr)"
    )
    print(f"  Bar chart: {chart_path}")
    
    pie_path = engine.generate_chart_image(
        "pie",
        {"labels": ["Products", "Services", "Exports"], "values": [45, 30, 25]},
        "Revenue Mix"
    )
    print(f"  Pie chart: {pie_path}")
