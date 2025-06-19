# Auto Consolidator - Deployment Guide

## Overview
This guide explains how to package the Auto Consolidator into a standalone executable and deploy it to your company server for estimators to use.

## Prerequisites
- Python 3.7 or later installed on your development machine
- All Auto Consolidator files in one folder
- Your custom icon file (Auto Consolidator Icon.png)

## Step 1: Prepare for Building

1. **Install Dependencies**
   ```bash
   pip install -r requirements.txt
   ```

2. **Copy Your Icon**
   - Copy `Auto Consolidator Icon.png` to the same folder as `auto_consolidator.py`
   - Or update the path in `build_executable.py` line 31

## Step 2: Build the Executable

### Option A: Quick Build (Recommended)
1. Double-click `build.bat`
2. Wait for the build process to complete
3. Check the `Auto_Consolidator_Package` folder

### Option B: Manual Build
1. Run: `python build_executable.py`
2. Follow the on-screen instructions

### Option C: Advanced Build
1. Run: `pyinstaller auto_consolidator.spec`
2. Manually copy files to distribution folder

## Step 3: Test the Executable

1. Navigate to `Auto_Consolidator_Package`
2. Double-click `Auto_Consolidator.exe`
3. Test with sample files to ensure everything works
4. Verify the icon appears correctly

## Step 4: Deploy to Company Server

### For Network Drive Deployment:
1. Copy the entire `Auto_Consolidator_Package` folder to your company server
2. Set appropriate permissions (Read & Execute for estimators)
3. Share the network path with estimators

### For Email Distribution:
1. Zip the `Auto_Consolidator_Package` folder
2. Email to estimators with installation instructions
3. Include the README.txt for user guidance

## Step 5: User Instructions

Send this to your estimators:

```
Auto Consolidator Installation
=============================

1. Copy the Auto_Consolidator_Package folder to your computer
2. Double-click Auto_Consolidator.exe to run
3. Configure your Cell Map.xlsx file for your specific needs
4. Use as normal - no additional installation required

Requirements:
- Windows 7 or later
- Excel 2016 or later
- No admin rights needed
```

## Troubleshooting

### Build Issues:
- **"PyInstaller not found"**: Run `pip install pyinstaller`
- **"Module not found"**: Run `pip install -r requirements.txt`
- **Icon issues**: Install Pillow with `pip install Pillow`

### Runtime Issues:
- **"Cannot run executable"**: Check antivirus settings
- **"Missing files"**: Ensure entire package folder is copied
- **Excel errors**: Verify Excel version compatibility

### File Size:
- Expected size: 25-40 MB
- If larger: Check excludes in build script
- If smaller: Some libraries may be missing

## File Structure After Build:
```
Auto_Consolidator_Package/
├── Auto_Consolidator.exe    # Main executable
├── Cell Map.xlsx           # Template mapping file
└── README.txt              # User instructions
```

## Security Considerations:
- Executable is not code-signed (may trigger antivirus warnings)
- Consider adding to antivirus whitelist
- Test on target machines before wide deployment
- Keep source code for future updates

## Updating the Program:
1. Modify source code as needed
2. Re-run build process
3. Replace old package folder with new one
4. Notify users of updates

## Version Control:
- Tag releases in your version control system
- Keep built executables for rollback purposes
- Document changes in README.txt
