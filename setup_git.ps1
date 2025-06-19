#!/bin/bash
# Auto Consolidator - GitHub Repository Setup Script
# Run this in PowerShell from your project directory

echo "Auto Consolidator - GitHub Repository Setup"
echo "==========================================="
echo ""

# Check if git is installed
if (Get-Command git -ErrorAction SilentlyContinue) {
    echo "✓ Git is installed"
} else {
    echo "✗ Git is not installed. Please install Git first:"
    echo "  https://git-scm.com/download/win"
    exit 1
}

# Initialize git repository
echo "Initializing Git repository..."
git init

# Add all files
echo "Adding files to git..."
git add .

# Create initial commit
echo "Creating initial commit..."
git commit -m "Initial commit: Auto Consolidator v2.0

- Complete refactor with improved architecture
- Added automatic item numbering
- Enhanced error handling and validation
- Modern GUI with progress tracking
- Executable packaging support
- Security improvements"

# Check if main branch exists, if not create it
$currentBranch = git branch --show-current
if ($currentBranch -ne "main") {
    echo "Renaming branch to main..."
    git branch -M main
}

echo ""
echo "Local repository initialized!"
echo ""
echo "Next steps:"
echo "1. Create a new repository on GitHub.com"
echo "2. Copy the repository URL"
echo "3. Run: git remote add origin https://github.com/yourusername/auto-consolidator.git"
echo "4. Run: git push -u origin main"
echo ""
echo "Repository is ready for GitHub!"
