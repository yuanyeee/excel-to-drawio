"""Conversion configuration."""
import dataclasses

@dataclasses.dataclass
class ConvertConfig:
    """Conversion settings. All fields have sensible defaults."""
    scale: float = 1.0
    char_width: int = 7
    point_to_px: float = 96 / 72
    emu_per_px: int = 9525
    embed_images: bool = True
    skip_hidden: bool = False
    merge_fills: bool = True
    render_borders: bool = True
    render_fills: bool = True
    render_labels: bool = True
    render_shapes: bool = True
    render_images: bool = True
    page_mode: bool = True

