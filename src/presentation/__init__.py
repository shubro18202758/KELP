"""
Presentation Module - PowerPoint generation with native charts
"""
from .ppt_generator import (
    SlideLayout,
    KelpPPTGenerator,
    generate_teaser_ppt
)

__all__ = [
    'SlideLayout',
    'KelpPPTGenerator',
    'generate_teaser_ppt'
]
