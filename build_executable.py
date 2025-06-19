#!/usr/bin/env python3
"""
Build script for creating Auto Consolidator executable
This script will package the Auto Consolidator into a standalone executable
"""

import os
import sys
import subprocess
import shutil
from pathlib import Path

def check_pyinstaller():
    """Check if PyInstaller is installed, install if needed"""
    try:
        import PyInstaller
        print("✓ PyInstaller is already installed")
        return True
    except ImportError:
        print("PyInstaller not found. Installing...")
        try:
            subprocess.check_call([sys.executable, "-m", "pip", "install", "pyinstaller"])
            print("✓ PyInstaller installed successfully")
            return True
        except subprocess.CalledProcessError:
            print("✗ Failed to install PyInstaller")
            return False

def build_executable():
    """Build the executable using PyInstaller"""
    
    # Get current directory
    current_dir = Path(__file__).parent
    script_file = current_dir / "auto_consolidator.py"
    icon_file = Path(r"C:\Users\AaronMelton\Downloads\Auto Consolidator Icon.png")
    
    # Check if main script exists
    if not script_file.exists():
        print(f"✗ Main script not found: {script_file}")
        return False
    
    # Check if icon exists
    icon_exists = icon_file.exists()
    if not icon_exists:
        print(f"⚠ Icon file not found: {icon_file}")
        print("  Continuing without custom icon...")
    else:
        print(f"✓ Icon file found: {icon_file}")
    
    # Build PyInstaller command
    cmd = [
        "pyinstaller",
        "--onefile",                    # Create single executable file
        "--windowed",                   # Hide console window (GUI app)
        "--name", "Auto_Consolidator",  # Executable name
        "--clean",                      # Clean PyInstaller cache
        "--noconfirm",                  # Overwrite output without asking
        # Add hidden imports for packages that might not be detected
        "--hidden-import", "openpyxl",
        "--hidden-import", "pandas",
        "--hidden-import", "tkinter",
        "--hidden-import", "ttkthemes",
        # Exclude unnecessary packages to reduce size
        "--exclude-module", "matplotlib",
        "--exclude-module", "IPython",
        "--exclude-module", "jupyter",
        "--exclude-module", "notebook",
    ]
    
    # Add icon if available (convert PNG to ICO first if needed)
    if icon_exists:
        ico_file = current_dir / "auto_consolidator.ico"
        try:
            # Try to convert PNG to ICO using PIL
            from PIL import Image
            img = Image.open(icon_file)
            # Resize to standard icon sizes and save as ICO
            img.save(ico_file, format='ICO', sizes=[(16,16), (32,32), (48,48), (64,64), (128,128), (256,256)])
            cmd.extend(["--icon", str(ico_file)])
            print(f"✓ Converted PNG icon to ICO: {ico_file}")
        except ImportError:
            print("⚠ PIL/Pillow not available for icon conversion")
            print("  Install with: pip install Pillow")
        except Exception as e:
            print(f"⚠ Could not convert icon: {e}")
    
    # Add the main script
    cmd.append(str(script_file))
    
    print(f"Building executable with command:")
    print(f"  {' '.join(cmd)}")
    print()
    
    try:
        # Run PyInstaller
        result = subprocess.run(cmd, cwd=current_dir, capture_output=True, text=True)
        
        if result.returncode == 0:
            print("✓ Executable built successfully!")
            
            # Find the executable
            dist_dir = current_dir / "dist"
            exe_file = dist_dir / "Auto_Consolidator.exe"
            
            if exe_file.exists():
                file_size = exe_file.stat().st_size / (1024 * 1024)  # Size in MB
                print(f"✓ Executable created: {exe_file}")
                print(f"  Size: {file_size:.1f} MB")
                
                # Copy Cell Map file to dist folder
                cell_map_source = current_dir / "Cell Map.xlsx"
                if cell_map_source.exists():
                    cell_map_dest = dist_dir / "Cell Map.xlsx"
                    shutil.copy2(cell_map_source, cell_map_dest)
                    print(f"✓ Copied Cell Map.xlsx to distribution folder")
                
                return True
            else:
                print("✗ Executable not found in dist folder")
                return False
        else:
            print("✗ Build failed!")
            print("STDOUT:", result.stdout)
            print("STDERR:", result.stderr)
            return False
            
    except FileNotFoundError:
        print("✗ PyInstaller not found in PATH")
        return False
    except Exception as e:
        print(f"✗ Build failed with error: {e}")
        return False

def create_distribution_package():
    """Create a complete distribution package"""
    current_dir = Path(__file__).parent
    dist_dir = current_dir / "dist"
    package_dir = current_dir / "Auto_Consolidator_Package"
    
    if not dist_dir.exists():
        print("✗ No dist folder found. Run build first.")
        return False
    
    # Create package directory
    if package_dir.exists():
        shutil.rmtree(package_dir)
    package_dir.mkdir()
    
    # Copy executable
    exe_file = dist_dir / "Auto_Consolidator.exe"
    if exe_file.exists():
        shutil.copy2(exe_file, package_dir / "Auto_Consolidator.exe")
        print(f"✓ Copied executable to package")
    
    # Copy Cell Map if it exists
    cell_map_file = dist_dir / "Cell Map.xlsx"
    if cell_map_file.exists():
        shutil.copy2(cell_map_file, package_dir / "Cell Map.xlsx")
        print(f"✓ Copied Cell Map.xlsx to package")
    
    # Create README file
    readme_content = """Auto Consolidator - Installation and Usage Guide

WHAT IS THIS?
============
Auto Consolidator is a tool for consolidating data from multiple Excel estimate files 
into a single consolidation spreadsheet with live formula links.

INSTALLATION
===========
1. Copy this entire folder to your desired location
2. No additional installation required - this is a standalone executable

USAGE
=====
1. Double-click "Auto_Consolidator.exe" to start the program
2. Configure your Cell Map file (Cell Map.xlsx is included as a template)
3. Select your consolidation target file
4. Add estimate files to process
5. Click "Run Consolidation"

FILES INCLUDED
=============
- Auto_Consolidator.exe: The main program
- Cell Map.xlsx: Template mapping file (configure this for your needs)
- README.txt: This file

REQUIREMENTS
===========
- Windows 7 or later
- Excel 2016 or later (for .xlsx file support)

TROUBLESHOOTING
==============
- If the program doesn't start, check that you have permission to run executables
- Ensure all Excel files are closed before running consolidation
- Check that your Cell Map.xlsx file has the correct column headers:
  * Source Sheet
  * Source Cell  
  * Destination Column (Consolidation)

SUPPORT
=======
Contact your IT administrator or the program developer for assistance.

Version: 2.0
Built: """ + str(Path(__file__).stat().st_mtime) + """
"""
    
    readme_file = package_dir / "README.txt"
    readme_file.write_text(readme_content)
    print(f"✓ Created README.txt")
    
    print(f"\n✓ Distribution package created: {package_dir}")
    print(f"  Ready to copy to company server or share with estimators")
    
    return True

def main():
    """Main build process"""
    print("Auto Consolidator - Executable Builder")
    print("=" * 40)
    
    # Check PyInstaller
    if not check_pyinstaller():
        return False
    
    # Build executable
    print("\n1. Building executable...")
    if not build_executable():
        return False
    
    # Create distribution package
    print("\n2. Creating distribution package...")
    if not create_distribution_package():
        return False
    
    print("\n" + "=" * 40)
    print("✓ BUILD COMPLETE!")
    print("\nNext steps:")
    print("1. Test the executable in the 'Auto_Consolidator_Package' folder")
    print("2. Copy the entire package folder to your company server")
    print("3. Share the package folder location with estimators")
    print("4. Estimators can run 'Auto_Consolidator.exe' directly")
    
    return True

if __name__ == "__main__":
    success = main()
    input("\nPress Enter to exit...")
    sys.exit(0 if success else 1)
