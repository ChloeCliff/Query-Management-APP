#!/usr/bin/env python3
"""Generate PNG icons and a Windows .ico from qbox-logo.svg."""
import os
import subprocess
import sys

def generate_icons():
    """Generate PNG icons in multiple sizes"""
    sizes = [256, 128, 64, 32, 16]
    svg_file = "qbox-logo.svg"
    use_windows = sys.platform.startswith("win")
    
    for size in sizes:
        png_file = f"qbox-icon-{size}.png"
        try:
            # First choice: cairosvg (pure Python path, no shell tool ambiguity)
            import cairosvg
            cairosvg.svg2png(url=svg_file, write_to=png_file, output_width=size, output_height=size)
            print(f"✓ Generated {png_file}")
        except ImportError:
            try:
                # Fallback: ImageMagick via `magick` (on Windows, avoid plain `convert`)
                cmd = ["magick", "convert", "-density", "300", "-resize", f"{size}x{size}", svg_file, png_file] if use_windows else ["magick", "-density", "300", "-resize", f"{size}x{size}", svg_file, png_file]
                subprocess.run(cmd, check=True, capture_output=True, text=True, timeout=30)
                print(f"✓ Generated {png_file}")
            except (FileNotFoundError, subprocess.CalledProcessError, subprocess.TimeoutExpired):
                print(f"✗ Could not generate {png_file} - install cairosvg or ImageMagick (magick)")
                return False
    
    try:
        from PIL import Image

        png_256 = "qbox-icon-256.png"
        if not os.path.exists(png_256):
            print("✗ Could not create qbox.ico - qbox-icon-256.png missing")
            return False

        img = Image.open(png_256)
        img.save("qbox.ico", format="ICO", sizes=[(256, 256), (128, 128), (64, 64), (32, 32), (16, 16)])
        print("✓ Generated qbox.ico")
    except ImportError:
        print("✗ Could not generate qbox.ico - install Pillow")
        return False

    print("\n✓ All icons generated successfully!")
    print("\nFiles created:")
    for size in sizes:
        print(f"  - qbox-icon-{size}.png")
    print("  - qbox.ico")
    return True

if __name__ == "__main__":
    success = generate_icons()
    sys.exit(0 if success else 1)
