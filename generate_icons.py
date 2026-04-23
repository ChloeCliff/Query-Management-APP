#!/usr/bin/env python3
"""Generate PNG icons from qbox-logo.svg"""
import subprocess
import sys

def generate_icons():
    """Generate PNG icons in multiple sizes"""
    sizes = [256, 128, 64, 32, 16]
    svg_file = "qbox-logo.svg"
    
    for size in sizes:
        png_file = f"qbox-icon-{size}.png"
        try:
            # Try ImageMagick
            subprocess.run([
                "convert", "-density", "300", "-resize", f"{size}x{size}",
                svg_file, png_file
            ], check=True, capture_output=True)
            print(f"✓ Generated {png_file}")
        except (FileNotFoundError, subprocess.CalledProcessError):
            try:
                # Fallback: try with cairosvg if available
                import cairosvg
                cairosvg.svg2png(url=svg_file, write_to=png_file, output_width=size, output_height=size)
                print(f"✓ Generated {png_file}")
            except ImportError:
                print(f"✗ Could not generate {png_file} - install ImageMagick or cairosvg")
                return False
    
    print("\n✓ All icons generated successfully!")
    print("\nFiles created:")
    for size in sizes:
        print(f"  - qbox-icon-{size}.png")
    return True

if __name__ == "__main__":
    success = generate_icons()
    sys.exit(0 if success else 1)
