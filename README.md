# Auto Consolidator

A powerful Excel consolidation tool that automatically links data from multiple estimate files into a single consolidated spreadsheet with live formulas.

![Auto Consolidator](https://img.shields.io/badge/Python-3.7%2B-blue.svg)
![License](https://img.shields.io/badge/License-MIT-green.svg)
![Platform](https://img.shields.io/badge/Platform-Windows-lightgrey.svg)

## Features

- **Automated Data Consolidation**: Consolidate data from multiple Excel estimate files
- **Live Formula Linking**: Creates Excel formulas that link back to source files
- **Automatic Item Numbering**: Sequential numbering for consolidated entries
- **User-Friendly GUI**: Modern, intuitive interface built with tkinter
- **Configurable Mappings**: Define custom cell mappings via Excel template
- **Progress Tracking**: Real-time progress updates during processing
- **Error Handling**: Comprehensive validation and error reporting
- **Executable Distribution**: Standalone .exe file for easy deployment

## Screenshot

![Auto Consolidator Interface](![image](https://github.com/user-attachments/assets/142ad51f-d554-472a-9985-03bc359e7c23))
*Modern, user-friendly interface for Excel consolidation*

## Quick Start

### For End Users (Executable)

1. Download the latest release from the [Releases](../../releases) page
2. Extract the package folder
3. Double-click `Auto_Consolidator.exe` to run
4. Configure your `Cell Map.xlsx` file
5. Select files and run consolidation

### For Developers (Source Code)

```bash
# Clone the repository
git clone https://github.com/StrongARMDevelopment/auto-consolidator.git
cd auto-consolidator

# Install dependencies
pip install -r requirements.txt

# Run the application
python auto_consolidator.py
```

## Installation

### System Requirements

- Windows 7 or later
- Excel 2016 or later (for .xlsx support)
- Python 3.7+ (for development only)

### Dependencies

The application requires the following Python packages:

- `pandas` - Data manipulation and analysis
- `openpyxl` - Excel file handling
- `tkinter` - GUI framework (included with Python)
- `ttkthemes` - Modern GUI themes
- `Pillow` - Image processing for icons

Install all dependencies with:
```bash
pip install -r requirements.txt
```

## Usage

### 1. Configure Cell Mapping

Edit the `Cell Map.xlsx` file to define:
- **Source Sheet**: Sheet name in estimate files
- **Source Cell**: Cell reference (e.g., "C5")
- **Destination Column**: Target column in consolidation file

### 2. Prepare Consolidation File

Create an Excel file with:
- Header row (default: row 4)
- Column names matching your cell map destinations
- Data start row (default: row 5)

### 3. Run Consolidation

1. Launch Auto Consolidator
2. Select your Cell Map file
3. Select your Consolidation file
4. Add estimate files to process
5. Configure settings (header row, data start row)
6. Click "Run Consolidation"

### 4. Review Results

- Output file is saved with timestamp
- Formulas link back to source files
- Item numbers are automatically assigned
- Open output file or folder directly from the app

## Building Executable

To create a standalone executable:

```bash
# Quick build
build.bat

# Or manual build with PyInstaller
pyinstaller --onefile --windowed --name "Auto_Consolidator" auto_consolidator.py
```

The executable will be created in the `dist` folder.

## Project Structure

```
auto-consolidator/
├── auto_consolidator.py       # Main application
├── requirements.txt           # Python dependencies
├── build.bat                  # Quick build script
├── Cell Map.xlsx             # Template mapping file
├── README.md                 # This file
├── LICENSE                   # Project license
├── .gitignore               # Git ignore rules
└── consolidator_errors.log  # Application error log (created at runtime)
```

## Configuration

### Cell Map Format

The Cell Map Excel file must contain these columns:

| Source Sheet | Source Cell | Destination Column (Consolidation) |
|--------------|-------------|-----------------------------------|
| Summary      | C5          | Total Cost                        |
| Materials    | D10         | Material Cost                     |
| Labor        | E15         | Labor Hours                       |

### Default Settings

- Header Row: 4
- Data Start Row: 5
- Consolidation Sheet: "General Consolidation"
- Auto-clear existing data: Enabled

## Error Handling

The application includes comprehensive error handling:

- **File Validation**: Checks file existence and format
- **Sheet Validation**: Verifies required sheets exist
- **Cell Validation**: Confirms cell references are valid
- **Path Security**: Prevents path traversal attacks
- **Excel Formula Protection**: Guards against formula injection

Errors are logged to `consolidator_errors.log` and displayed in the GUI.

## Contributing

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/amazing-feature`)
3. Commit your changes (`git commit -m 'Add amazing feature'`)
4. Push to the branch (`git push origin feature/amazing-feature`)
5. Open a Pull Request

### Development Setup

```bash
# Clone your fork
git clone https://github.com/yourusername/auto-consolidator.git

# Create virtual environment
python -m venv venv
source venv/bin/activate  # On Windows: venv\Scripts\activate

# Install development dependencies
pip install -r requirements.txt
pip install -r requirements-dev.txt  # If available

# Run tests
python -m pytest tests/  # If tests are available
```

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## Support

- **Issues**: Report bugs and request features via [GitHub Issues](../../issues)
- **Documentation**: See [DEPLOYMENT_GUIDE.md](DEPLOYMENT_GUIDE.md) for detailed setup instructions
- **Discussions**: Ask questions in [GitHub Discussions](../../discussions)

## Changelog

### Version 2.0.0
- Complete refactor with improved architecture
- Added automatic item numbering
- Enhanced error handling and validation
- Modern GUI with progress tracking
- Executable packaging support
- Security improvements

### Version 1.0.0
- Initial release
- Basic consolidation functionality
- Simple GUI interface

## Acknowledgments

- Built with Python and tkinter
- Uses openpyxl for Excel file handling
- Pandas for data manipulation
- PyInstaller for executable creation

---

**Made for estimators, by estimators** 📊

For questions or support, please open an issue or contact the development team.