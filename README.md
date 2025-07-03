# PowerPoint Watermark Script

A Python script to automatically add watermarks to all slides in PowerPoint presentations. Supports both standard positioning and ribbon-style watermarks that span across slides.

## ✨ Features

- **Multiple positioning options**: Standard corners, center, or full ribbon styles
- **Smart ribbon sizing**: Automatically adjusts ribbon height based on watermark aspect ratio
- **Opacity control**: Set transparency from 0% to 100%
- **Aspect ratio preservation**: No image stretching or distortion
- **Layering control**: Place watermarks on top or underneath slide content
- **Batch processing**: Applies watermark to all slides at once

## 🚀 Quick Start

### Installation

1. **Clone or download** this repository
2. **Install dependencies**:
   ```bash
   pip install -r requirements.txt
   ```

### Basic Usage

```python
from main import add_watermark_to_pptx

# Simple watermark in bottom-right corner
add_watermark_to_pptx(
    pptx_path="your_presentation.pptx",
    watermark_path="your_logo.png"
)
```

## 📋 Requirements

- Python 3.7+
- python-pptx>=1.0.2
- Pillow>=10.0.0

## 🎯 Position Options

### Standard Positions
- `center` - Center of slide
- `bottom-right` - Bottom right corner with margin
- `bottom-left` - Bottom left corner with margin  
- `top-right` - Top right corner with margin
- `top-left` - Top left corner with margin

### Ribbon Positions
- `diagonal-ribbon` - Diagonal ribbon across entire slide
- `horizontal-ribbon` - Horizontal ribbon across slide width
- `vertical-ribbon` - Vertical ribbon across slide height

## 🛠️ Parameters

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `pptx_path` | str | Required | Path to input PowerPoint file |
| `watermark_path` | str | Required | Path to watermark image |
| `output_path` | str | `None` | Output file path (auto-generated if None) |
| `opacity` | float | `0.5` | Watermark opacity (0.0 to 1.0) |
| `position` | str | `"center"` | Watermark position (see options above) |
| `size_percentage` | float | `0.3` | Size relative to slide width (ignored for ribbons) |
| `on_top` | bool | `True` | Whether watermark appears above content |

## 📚 Examples

### Example 1: Standard Logo in Corner
```python
add_watermark_to_pptx(
    pptx_path="presentation.pptx",
    watermark_path="company_logo.png",
    opacity=0.7,
    position="bottom-right",
    size_percentage=0.15
)
```

### Example 2: Diagonal "CONFIDENTIAL" Ribbon
```python
add_watermark_to_pptx(
    pptx_path="sensitive_doc.pptx",
    watermark_path="confidential.png",
    output_path="confidential_presentation.pptx",
    opacity=0.3,
    position="diagonal-ribbon"
)
```

### Example 3: Subtle Background Watermark
```python
add_watermark_to_pptx(
    pptx_path="presentation.pptx",
    watermark_path="background_logo.png",
    opacity=0.1,
    position="center",
    size_percentage=0.5,
    on_top=False  # Behind content
)
```

### Example 4: Horizontal Brand Banner
```python
add_watermark_to_pptx(
    pptx_path="brand_presentation.pptx",
    watermark_path="brand_banner.png",
    opacity=0.4,
    position="horizontal-ribbon"
)
```

## 🎨 Smart Ribbon Sizing

The script automatically adjusts ribbon height based on your watermark's aspect ratio:

- **Wide images** (width/height > 2.5): 20% of slide height
- **Normal/tall images**: 35% of slide height

This ensures wide text watermarks don't create overly tall ribbons, while square logos get appropriate prominence.

## 🔄 Aspect Ratio Preservation

Unlike simple image insertion, this script:
- ✅ Maintains original image proportions
- ✅ Scales appropriately for different positions
- ✅ Prevents stretching or distortion
- ✅ Automatically calculates optimal dimensions

## 💡 Tips

1. **For text watermarks**: Use wide PNG images with transparent backgrounds
2. **For logo watermarks**: Square or vertical orientations work best for ribbons
3. **Opacity suggestions**: 
   - Logos: 0.7-1.0
   - Background watermarks: 0.1-0.3
   - Text overlays: 0.3-0.6
4. **File formats**: PNG with transparency recommended for best results

## 🐛 Troubleshooting

**Import errors**: Make sure you're in the virtual environment:
```bash
source venv/bin/activate  # On macOS/Linux
# or
venv\Scripts\activate     # On Windows
```

**Watermark too big/small**: Adjust `size_percentage` for standard positions or check your image's aspect ratio for ribbons.

**Watermark not visible**: Check `opacity` setting and ensure `on_top=True` if you want it above content.

## 📄 License

This project is open source. Feel free to modify and distribute as needed.

---

**Need help?** Check the examples above or examine the script's docstrings for detailed parameter information. 