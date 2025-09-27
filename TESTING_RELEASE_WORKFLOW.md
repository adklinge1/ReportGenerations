# Testing the GitHub Actions Release Workflow

This guide provides step-by-step instructions for testing the GitHub Actions release workflow before using it in production.

## Prerequisites

Before testing, ensure you have:
- [ ] A GitHub repository with the workflow files committed
- [ ] GitHub Actions enabled on your repository
- [ ] Write access to the repository
- [ ] Git configured locally with your GitHub credentials

## Setup Instructions

### 1. Update Repository URLs
Before testing, update the placeholder URLs in your files:

1. **In `release.bat`** (lines 66 and 69):
   - Replace `adklinge1` with your GitHub username
   - Replace `ReportGenerations` with your repository name

2. **In `RELEASE.md`** (line 48):
   - Replace `adklinge1` with your GitHub username  
   - Replace `ReportGenerations` with your repository name

### 2. Verify Workflow File
Ensure `.github/workflows/release.yml` is present and properly configured.

## Testing Methods

### Method 1: Manual Workflow Dispatch (Recommended for Testing)

This method allows you to test without creating permanent git tags.

#### Steps:
1. **Navigate to Actions Tab**
   - Go to your GitHub repository
   - Click the "Actions" tab
   - Look for "Build and Release ReportGenerator" workflow

2. **Trigger Manual Run**
   - Click on the workflow name
   - Click "Run workflow" button (top right)
   - Enter a test version like `v0.0.1-test`
   - Click "Run workflow"

3. **Monitor the Build**
   - Watch the workflow execution in real-time
   - Check each step for errors or warnings
   - Typical build time: 3-5 minutes

#### Expected Results:
- ✅ All steps should complete successfully
- ✅ Build artifacts should be created
- ❌ No GitHub release will be created (manual dispatch doesn't create releases)

### Method 2: Git Tag Testing (Creates Actual Release)

⚠️ **Warning**: This creates a real release. Use only when ready for production.

#### Steps:
1. **Prepare Your Code**
   ```bash
   git add .
   git commit -m "Prepare for release testing"
   ```

2. **Create and Push Tag**
   ```bash
   git tag v0.0.1-test
   git push origin v0.0.1-test
   ```

3. **Monitor Workflow**
   - Go to Actions tab immediately after pushing
   - Watch the automated workflow execution

#### Expected Results:
- ✅ Workflow should trigger automatically
- ✅ GitHub release should be created
- ✅ Release should contain the zip file
- ✅ Release should have proper description and notes

### Method 3: Using release.bat Script

Test the convenience script for creating releases.

#### Steps:
1. **Run the Script**
   ```cmd
   release.bat v0.0.2-test
   ```

2. **Follow Prompts**
   - Enter commit message when prompted
   - Script will handle git operations automatically

3. **Verify Results**
   - Check that tag was created and pushed
   - Monitor workflow in GitHub Actions

## Verification Checklist

After running any test method, verify the following:

### Workflow Execution
- [ ] All workflow steps completed without errors
- [ ] Build step completed successfully
- [ ] Dependencies were copied correctly
- [ ] Package was created with correct naming

### Build Artifacts
- [ ] ReportGenerator executable is present
- [ ] All required DLL files are included
- [ ] CSV data files are included
- [ ] HTML template files are included
- [ ] Excel template file is included

### Release Creation (Tag method only)
- [ ] GitHub release was created automatically
- [ ] Release has correct version number
- [ ] Zip file is attached to release
- [ ] Release notes are auto-generated
- [ ] Download link works correctly

### Package Contents
Download and extract the release package to verify:
- [ ] `ReportGenerator.exe` runs without errors
- [ ] All functionality works as expected
- [ ] No missing dependencies
- [ ] Application launches successfully

## Troubleshooting

### Common Issues and Solutions

#### Build Fails - NuGet Restore Error
**Problem**: NuGet packages fail to restore
**Solution**: 
- Check that `packages.config` exists in project
- Verify project file paths in workflow
- Ensure NuGet package sources are accessible

#### Build Fails - MSBuild Error
**Problem**: MSBuild compilation fails
**Solution**:
- Test local build first: Open in Visual Studio and build
- Check for syntax errors or missing references
- Verify .NET Framework version compatibility

#### Missing Files in Release Package
**Problem**: Required files not included in zip
**Solution**:
- Check file paths in workflow copy steps
- Verify files exist in repository
- Review PowerShell copy commands for errors

#### Workflow Doesn't Trigger
**Problem**: Pushing tag doesn't start workflow
**Solution**:
- Verify tag format matches pattern `v*.*.*`
- Check that workflow file is in `main` branch
- Ensure GitHub Actions is enabled

#### Release Not Created
**Problem**: Workflow runs but no release appears
**Solution**:
- Verify you used tag method (not manual dispatch)
- Check repository permissions
- Review workflow logs for release step errors

### Debug Steps

1. **Check Workflow Logs**
   - Go to Actions tab → Select failed run
   - Expand each step to see detailed logs
   - Look for red error messages

2. **Verify File Paths**
   - Ensure all referenced files exist
   - Check case sensitivity in file paths
   - Verify directory structure matches workflow

3. **Test Local Build**
   ```cmd
   # Test the exact build command locally
   msbuild WindowsFormsApp1/GenerateReport.csproj /p:Configuration=Release /p:Platform="Any CPU"
   ```

## Pre-Production Checklist

Before using for actual releases:

- [ ] Successfully completed Method 1 (Manual Dispatch)
- [ ] Successfully completed Method 2 (Git Tag) with test version
- [ ] Downloaded and tested the generated package
- [ ] Verified all application functionality
- [ ] Updated all placeholder URLs in scripts
- [ ] Documented any project-specific requirements
- [ ] Tested on clean Windows machine (if possible)

## Production Workflow

Once testing is complete:

1. **Clean up test releases** (optional)
   - Delete test tags: `git tag -d v0.0.1-test && git push origin :refs/tags/v0.0.1-test`
   - Delete test releases from GitHub UI

2. **Create your first production release**
   ```bash
   release.bat v1.0.0
   ```

3. **Monitor and validate**
   - Ensure release completes successfully
   - Test download and installation
   - Share release URL with end users

## Support

If you encounter issues:
1. Check the troubleshooting section above
2. Review GitHub Actions documentation
3. Examine workflow logs for specific error messages
4. Verify local build works before debugging workflow

---

**Next Steps**: After successful testing, see [RELEASE.md](RELEASE.md) for production release instructions.