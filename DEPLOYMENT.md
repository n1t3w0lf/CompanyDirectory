# People Directory WebPart - Deployment Guide

This guide provides step-by-step instructions for deploying the People Directory WebPart to your SharePoint Online environment.

## Prerequisites

Before deployment, ensure you have:

✅ SharePoint Online Administrator access
✅ Global Administrator or Application Administrator access (for API permissions)
✅ Access to SharePoint App Catalog
✅ Node.js v16.13.0+ or v18.17.1+ installed locally
✅ Microsoft 365 tenant with Azure AD

## Pre-Deployment Checklist

- [ ] Verify you have administrative access to SharePoint Admin Center
- [ ] Confirm App Catalog exists (create if necessary)
- [ ] Review API permissions requirements with security team
- [ ] Plan communication to end users
- [ ] Identify pilot users/sites for initial rollout

## Step 1: Build the Solution Package

### Option A: Build from Source

```bash
# Clone the repository
git clone <repository-url>
cd CompanyDirectory

# Install dependencies
npm install

# Build the production package
npm run package
```

The package will be created at: `solution/people-directory-webpart.sppkg`

### Option B: Use Pre-built Package

If a pre-built package is provided, skip to Step 2.

## Step 2: Deploy to App Catalog

### 2.1 Access App Catalog

1. Navigate to SharePoint Admin Center: `https://{tenant}-admin.sharepoint.com`
2. In the left menu, click **More features**
3. Under **Apps**, click **Open**
4. Click **App Catalog**

### 2.2 Create App Catalog (if needed)

If you don't have an App Catalog:

1. Click **Create a new app catalog site**
2. Fill in the required information:
   - **Title**: App Catalog
   - **Web Site Address**: apps
   - **Administrator**: Your email
3. Click **OK** and wait 10-15 minutes for provisioning

### 2.3 Upload Package

1. Go to **Apps for SharePoint**
2. Click **Upload** or drag and drop `people-directory-webpart.sppkg`
3. A dialog appears asking **Do you trust people-directory-webpart?**
4. Check **Make this solution available to all sites in the organization**
5. Click **Deploy**

✅ **Checkpoint**: The app should now appear in the App Catalog with status "Available"

## Step 3: Approve API Permissions

**CRITICAL STEP**: The web part will not function without these permissions.

### 3.1 Navigate to API Management

1. Go to SharePoint Admin Center: `https://{tenant}-admin.sharepoint.com`
2. In the left menu, expand **Advanced**
3. Click **API access**

### 3.2 Approve Permissions

You should see pending requests for:

| Resource | Permission | Type | Status |
|----------|-----------|------|--------|
| Microsoft Graph | User.ReadBasic.All | Delegated | Pending |
| Microsoft Graph | User.Read.All | Delegated | Pending |

For each pending request:
1. Select the request
2. Click **Approve**
3. Confirm by clicking **Approve** again

⏱️ **Time**: Permissions typically activate within 1-2 minutes

✅ **Checkpoint**: Both permissions should show status "Approved"

### 3.3 Verify Permissions

```bash
# Optional: Verify via PowerShell
Connect-PnPOnline -Url https://{tenant}-admin.sharepoint.com -Interactive

Get-PnPTenantServicePrincipalPermissionGrants |
  Where-Object { $_.Resource -eq "Microsoft Graph" } |
  Format-Table Resource, Scope
```

## Step 4: Test Deployment

### 4.1 Create Test Page

1. Navigate to any SharePoint site
2. Create a new modern page or edit an existing one
3. Click **+** to add a web part
4. Search for "People Directory"
5. Add the web part to the page

### 4.2 Verify Functionality

Test the following:
- [ ] Web part loads without errors
- [ ] Search box is functional
- [ ] Searching returns results (try your own name)
- [ ] User cards display correctly
- [ ] Clicking a user card opens details panel
- [ ] Profile photos load (if available)
- [ ] Contact buttons work (email, phone)
- [ ] Filters panel opens and functions

### 4.3 Check Browser Console

Open browser developer tools (F12) and check for:
- ✅ No red errors in Console tab
- ✅ No 403 (Forbidden) errors in Network tab
- ✅ Successful Graph API calls

## Step 5: SharePoint List Creation

The web part automatically creates a cache list on first use.

### 5.1 Verify List Creation

1. Go to **Site Contents** on the site where you added the web part
2. Look for a list named **PeopleDirectoryCache**
3. The list should have custom columns for user data

### 5.2 List Permissions

**Default**: The list inherits permissions from the parent site.

**Recommended**: Restrict direct user access:
1. Go to list settings
2. Click **Permissions for this list**
3. Click **Stop Inheriting Permissions**
4. Remove user edit permissions (keep read-only)

This prevents users from manually editing cached data.

## Step 6: Configure Web Part Properties

### 6.1 Web Part Settings

1. Edit the page where the web part is added
2. Click the **Edit** (pencil) icon on the web part
3. The property pane appears on the right

Available settings:
- **Title**: Display title (default: "People Directory")
- **Description**: Subtitle text

### 6.2 Recommended Settings

For a typical deployment:
```
Title: Find People
Description: Search for colleagues across the organization
```

For executive/leadership pages:
```
Title: Leadership Directory
Description: Connect with organizational leaders
```

## Step 7: Rollout Strategy

### Option A: Phased Rollout (Recommended)

**Phase 1: Pilot (Week 1)**
- Deploy to 2-3 pilot sites
- Gather feedback from 50-100 users
- Monitor performance and errors

**Phase 2: Department Rollout (Week 2-3)**
- Deploy to departmental sites
- Train department admins
- Document common questions

**Phase 3: Organization-wide (Week 4+)**
- Make available to all sites
- Announce via company communication channels
- Provide training resources

### Option B: Immediate Rollout

1. Add to home page/intranet hub
2. Send announcement email with:
   - Feature overview
   - How to use
   - Screenshot
   - Support contact

## Step 8: Monitoring and Maintenance

### 8.1 Monitor Performance

**Week 1**: Check daily for:
- Error reports from users
- Browser console errors
- Performance issues

**Month 1**: Monitor weekly

**Ongoing**: Monitor monthly or as needed

### 8.2 Cache Management

The cache is self-managing with LRU eviction. Manual intervention is rarely needed.

**To clear cache** (if needed):
1. Navigate to **PeopleDirectoryCache** list
2. Delete all items
3. Cache will repopulate on next use

### 8.3 Update Deployments

When a new version is released:

```bash
# Build new package
npm run package

# Upload to App Catalog (same location)
# Select "Replace" when prompted
# Version number will auto-increment
```

**Note**: Updates may take 24 hours to propagate to all sites.

## Step 9: User Training

### Training Materials Needed

1. **Quick Start Guide** (1 page)
   - How to search
   - How to filter
   - How to view details

2. **Video Tutorial** (2-3 minutes)
   - Screen recording of basic usage
   - Include voiceover or subtitles

3. **FAQ Document**
   - Common questions
   - Troubleshooting tips

### Sample Training Email

```
Subject: New Feature: People Directory

Hi Team,

We're excited to announce the launch of our new People Directory!

What is it?
A searchable directory of everyone in our organization with up-to-date
information from Azure AD.

Where to find it?
Available on the Intranet home page and departmental sites.

How to use it?
1. Type a name, email, or department in the search box
2. Use filters to narrow results
3. Click any person to view their full profile
4. Click email/phone icons for quick contact

Questions?
Contact IT Support at [email/phone]

[Link to detailed guide]
```

## Troubleshooting Common Issues

### Issue: "Insufficient permissions" error

**Cause**: API permissions not approved

**Solution**:
1. Verify permissions in SharePoint Admin Center → API access
2. Ensure both Graph permissions are approved
3. Wait 2-3 minutes for propagation
4. Refresh the page

### Issue: Web part doesn't appear in add menu

**Cause**: Package not deployed or not made available tenant-wide

**Solution**:
1. Go to App Catalog
2. Find the package
3. Verify "Available to all sites" is checked
4. Wait 10-15 minutes for propagation
5. Refresh SharePoint page

### Issue: SharePoint list not created

**Cause**: User lacks permissions to create lists

**Solution**:
1. Grant site owner permissions temporarily
2. Load the web part (list will be created)
3. Remove elevated permissions

### Issue: Search returns no results

**Possible causes**:
- API permissions not approved
- Graph API temporary outage
- User has no access to directory

**Solutions**:
1. Check API permissions (Step 3)
2. Test with your own user account
3. Verify user is in Azure AD
4. Check browser console for specific errors

### Issue: Profile photos not loading

**Cause**: User.Read.All permission not approved

**Solution**:
1. Verify User.Read.All is approved
2. Some users may not have photos (expected)
3. Fallback initials should display

## Security Best Practices

✅ **Do's**
- ✅ Review API permissions with security team before deployment
- ✅ Monitor API usage via Azure AD logs
- ✅ Keep the solution package updated
- ✅ Document customizations

❌ **Don'ts**
- ❌ Don't modify the SharePoint cache list directly
- ❌ Don't grant User.Read.All unless needed (User.ReadBasic.All may suffice)
- ❌ Don't disable error logging
- ❌ Don't bypass the caching mechanism

## Performance Considerations

### Expected Performance

| Metric | Target | Notes |
|--------|--------|-------|
| Initial load | < 2 seconds | First load, no cache |
| Search response | < 1 second | With cache |
| Photo loading | < 100ms per image | Lazy loaded |
| Scroll performance | 60 FPS | Virtual scrolling |

### Optimization Tips

1. **First Search Warm-up**: Pre-cache popular searches
2. **Network**: Ensure good connectivity to Microsoft Graph
3. **Browser**: Recommend Chromium-based browsers for best performance

## Rollback Procedure

If you need to rollback the deployment:

1. Go to App Catalog
2. Find **people-directory-webpart**
3. Click **...** → **Remove**
4. Confirm removal

**Note**: This will disable the web part on all sites immediately.

To restore:
1. Re-upload the previous version package
2. Click **Deploy**
3. Re-approve API permissions if needed

## Support Contacts

After deployment, establish support contacts:

- **Level 1 Support**: IT Helpdesk
- **Level 2 Support**: SharePoint Admins
- **Level 3 Support**: Development Team

Create a support process:
1. User reports issue → IT Helpdesk
2. Helpdesk documents issue, checks common fixes
3. Escalate to SharePoint Admin if unresolved
4. SharePoint Admin investigates, escalates to Dev if needed

## Success Metrics

Track these metrics post-deployment:

- **Adoption**: Number of sites using the web part
- **Usage**: Number of searches per day
- **Performance**: Average response time
- **Support**: Number of support tickets
- **Satisfaction**: User feedback scores

## Post-Deployment Checklist

One week after deployment:

- [ ] Review error logs
- [ ] Check cache list size (should be < 5000 items)
- [ ] Verify API permission usage
- [ ] Collect user feedback
- [ ] Update documentation based on feedback
- [ ] Plan for any needed improvements

## Additional Resources

- [SharePoint Admin Center](https://docs.microsoft.com/en-us/sharepoint/get-started-new-admin-center)
- [API Access Management](https://docs.microsoft.com/en-us/sharepoint/api-access)
- [App Catalog](https://docs.microsoft.com/en-us/sharepoint/use-app-catalog)
- [SPFx Deployment](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/toolchain/deploy-your-spfx-solution-to-a-sharepoint-site)

## Contact

For deployment assistance, contact:
- **Technical Lead**: [Name/Email]
- **Project Manager**: [Name/Email]
- **Support**: [Support Email/Phone]
