#!/usr/bin/env python3
"""
Icon Generation Script for Rewordly
Generates PNG icons from SVG source for the Outlook Add-in
"""

import os
import sys
from pathlib import Path

def generate_icons():
    """Generate PNG icons from SVG source"""
    
    # Check if cairosvg is available
    try:
        import cairosvg
    except ImportError:
        print("❌ cairosvg not found. Installing...")
        os.system("pip install cairosvg")
        try:
            import cairosvg
        except ImportError:
            print("❌ Failed to install cairosvg. Please install manually:")
            print("   pip install cairosvg")
            return False
    
    # Icon sizes required by the manifest
    sizes = [16, 32, 80]
    
    # Source SVG file
    svg_file = "icon.svg"
    
    if not os.path.exists(svg_file):
        print(f"❌ Source SVG file '{svg_file}' not found!")
        return False
    
    print("🎨 Generating PNG icons...")
    
    # Generate icons for each size
    for size in sizes:
        output_file = f"icon-{size}.png"
        
        try:
            # Convert SVG to PNG
            cairosvg.svg2png(
                url=svg_file,
                write_to=output_file,
                output_width=size,
                output_height=size
            )
            print(f"✅ Generated {output_file} ({size}x{size})")
        except Exception as e:
            print(f"❌ Failed to generate {output_file}: {e}")
            return False
    
    # Generate high-res icon (2x the largest size)
    try:
        cairosvg.svg2png(
            url=svg_file,
            write_to="hi-res-icon.png",
            output_width=160,
            output_height=160
        )
        print("✅ Generated hi-res-icon.png (160x160)")
    except Exception as e:
        print(f"❌ Failed to generate hi-res-icon.png: {e}")
        return False
    
    print("\n🎉 All icons generated successfully!")
    print("\nNext steps:")
    print("1. Review the generated PNG files")
    print("2. Update manifest.xml URLs to point to these files")
    print("3. For production, host these files on a CDN")
    
    return True

def update_manifest_urls():
    """Update manifest.xml to use local icon files"""
    
    manifest_path = "../manifest.xml"
    
    if not os.path.exists(manifest_path):
        print(f"❌ Manifest file not found at {manifest_path}")
        return False
    
    print("📝 Updating manifest.xml with local icon URLs...")
    
    # Read manifest content
    with open(manifest_path, 'r') as f:
        content = f.read()
    
    # Update icon URLs to use local files
    updates = {
        'https://www.contoso.com/assets/icon-32.png': './assets/icon-32.png',
        'https://www.contoso.com/assets/hi-res-icon.png': './assets/hi-res-icon.png',
        'https://www.contoso.com/assets/icon-16.png': './assets/icon-16.png',
        'https://www.contoso.com/assets/icon-80.png': './assets/icon-80.png'
    }
    
    for old_url, new_url in updates.items():
        content = content.replace(old_url, new_url)
    
    # Write updated content
    with open(manifest_path, 'w') as f:
        f.write(content)
    
    print("✅ Manifest.xml updated with local icon URLs")
    return True

if __name__ == "__main__":
    print("🤖 Rewordly Icon Generator")
    print("=" * 30)
    
    # Change to assets directory
    os.chdir(os.path.dirname(os.path.abspath(__file__)))
    
    # Generate icons
    if generate_icons():
        # Update manifest
        update_manifest_urls()
        print("\n🎉 Icon generation complete!")
    else:
        print("\n❌ Icon generation failed!")
        sys.exit(1) 