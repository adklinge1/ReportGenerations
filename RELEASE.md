# ReportGenerator Release Process

## Automated Release Setup

This repository is configured with GitHub Actions to automatically build and release the ReportGenerator Windows Forms application.

## How to Create a Release

### Method 1: Git Tag (Recommended)
```bash
# Make your changes and commit them
git add .
git commit -m "Your changes description"

# Create a version tag
git tag v1.0.0

# Push the tag to trigger the release
git push origin v1.0.0
```

### Method 2: Manual Trigger
1. Go to your GitHub repository
2. Click "Actions" tab
3. Select "Build and Release ReportGenerator"
4. Click "Run workflow"
5. Enter version number (e.g., v1.0.1)
6. Click "Run workflow"

## What Happens Automatically

1. **Build Process**: Compiles the WindowsFormsApp1 project in Release mode
2. **Package Creation**: Creates a zip file with:
   - ReportGenerator.exe
   - All required DLL files
   - CSV data files (PalmSpeciesToScientificNameAndRate.csv, TreeSpeciesToScientificNameAndRate.csv)
   - HTML template files
   - Excel template file
3. **Release Creation**: Creates a GitHub release with:
   - Professional release page
   - Downloadable zip file
   - Auto-generated release notes
   - Installation instructions

## Customer Distribution

### Your Release Page
`https://github.com/adklinge1/ReportGenerations/releases`

### Customer Instructions
1. Visit the releases page
2. Download the latest `ReportGenerator-vX.X.X-win-x64.zip`
3. Extract the zip file
4. Run `ReportGenerator.exe`

## Version Numbering

Use semantic versioning: `v{major}.{minor}.{patch}`
- **Major**: Breaking changes
- **Minor**: New features  
- **Patch**: Bug fixes

Examples:
- `v1.0.0` - Initial release
- `v1.0.1` - Bug fix
- `v1.1.0` - New feature
- `v2.0.0` - Breaking changes

## Testing the Workflow

Before using this workflow in production, it's highly recommended to test it first. See [TESTING_RELEASE_WORKFLOW.md](TESTING_RELEASE_WORKFLOW.md) for comprehensive testing instructions.

### Quick Test (Recommended)
1. Go to your repository's "Actions" tab
2. Select "Build and Release ReportGenerator" 
3. Click "Run workflow"
4. Enter a test version like `v0.0.1-test`
5. Click "Run workflow" and monitor the build

This will test the build process without creating an actual release.

### Before Your First Production Release
- [ ] Complete the testing guide in [TESTING_RELEASE_WORKFLOW.md](TESTING_RELEASE_WORKFLOW.md)
- [ ] Update placeholder URLs in `release.bat` and this file
- [ ] Verify the generated package works on a clean machine
- [ ] Test the application functionality thoroughly

## Troubleshooting

### Build Fails
1. Check the Actions tab for error details
2. Ensure all NuGet packages restore correctly
3. Verify project builds locally in Visual Studio

### Missing Files in Release
The workflow automatically copies:
- Main executable and DLLs from `build/Release/`
- CSV files from `WindowsFormsApp1/ExcelReader/`
- HTML files from `WindowsFormsApp1/TreeCalculator/`
- Excel template from `WindowsFormsApp1/`

## Benefits

✅ **No more manual zipping**  
✅ **Professional distribution**  
✅ **Version control integration**  
✅ **Always up-to-date customer access**  
✅ **Automatic dependency inclusion**  
✅ **Self-documenting releases**