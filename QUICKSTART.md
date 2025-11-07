# Quick Start Guide - People Directory WebPart

Get started with the People Directory in 5 minutes!

## For End Users

### How to Search for People

1. **Open the People Directory**
   - Navigate to your SharePoint site
   - Find the People Directory web part (usually on the home page)

2. **Start Searching**
   - Type at least 2 characters in the search box
   - Results appear automatically after 300ms

3. **What You Can Search**
   - Name (first or last)
   - Email address
   - Department name
   - Job title

### How to Filter Results

1. **Open Filters**
   - Click the filter icon (funnel) in the top-right corner

2. **Apply Filters**
   - Select a department from the dropdown
   - Select an office location from the dropdown
   - Filters can be combined

3. **Clear Filters**
   - Click the X on any active filter badge
   - Or select "All Departments" / "All Locations" in dropdowns

### How to View Profile Details

1. **Click Any User Card**
   - Click on any person in the search results

2. **View Information**
   - Contact details (email, phone)
   - Organization info (department, location)
   - Manager information
   - Last updated timestamp

3. **Quick Actions**
   - Click email icon to send email
   - Click phone icon to call
   - Information is always current from Azure AD

### Tips for Best Results

✅ **DO:**
- Use at least 2 characters when searching
- Try different search terms (name, email, department)
- Use filters to narrow large result sets
- Click profile cards to verify current information

❌ **DON'T:**
- Don't use special characters in search
- Don't expect instant results (300ms debounce)
- Don't be concerned if first search is slow (cache warming)

---

## For Administrators

### Quick Deployment (10 Minutes)

**Prerequisites:**
- SharePoint Admin access
- Package file ready

**Steps:**

1. **Upload Package (2 min)**
   ```
   SharePoint Admin Center → Apps → App Catalog
   → Upload → people-directory-webpart.sppkg
   → Deploy (check "all sites")
   ```

2. **Approve Permissions (2 min)**
   ```
   SharePoint Admin Center → Advanced → API access
   → Approve "User.ReadBasic.All"
   → Approve "User.Read.All"
   ```

3. **Add to Page (2 min)**
   ```
   Edit SharePoint page → + Add web part
   → Search "People Directory" → Add
   ```

4. **Test (4 min)**
   - Search for your name
   - Verify results appear
   - Click your profile
   - Check photos load
   - Test filters

✅ **Done!** The web part is ready to use.

### Quick Troubleshooting

**Problem: "Insufficient permissions"**
→ **Fix**: Approve API permissions in Step 2 above

**Problem: No results**
→ **Fix**: Wait 2-3 minutes after approving permissions

**Problem: Photos not loading**
→ **Fix**: Verify User.Read.All is approved

**Problem: Web part not in menu**
→ **Fix**: Wait 10-15 minutes after deployment, refresh page

---

## For Developers

### Quick Setup (5 Minutes)

```bash
# Clone and install
git clone <repo-url>
cd CompanyDirectory
npm install

# Start development server
npm run serve

# Build for production
npm run package
```

### Quick File Reference

```
Key Files:
├── src/services/
│   ├── GraphService.ts        # Microsoft Graph API calls
│   ├── ListService.ts          # SharePoint list operations
│   └── PeopleService.ts        # Main orchestration logic
├── src/webparts/peopleDirectory/components/
│   ├── PeopleDirectory.tsx     # Main component
│   ├── UserCard.tsx            # User card component
│   └── UserDetailsPanel.tsx    # Details panel
└── src/models/
    ├── IUserProfile.ts         # TypeScript interfaces
    └── Constants.ts            # Configuration constants
```

### Quick Customization

**Change Search Debounce Time:**
```typescript
// src/models/Constants.ts
public static readonly SEARCH_DEBOUNCE_MS = 300; // Change this
```

**Change Cache TTL:**
```typescript
// src/models/Constants.ts
public static readonly CLIENT_CACHE_TTL = 30 * 60 * 1000; // 30 min
public static readonly LIST_CACHE_TTL = 24 * 60 * 60 * 1000; // 24 hr
```

**Change Page Size:**
```typescript
// src/models/Constants.ts
public static readonly DEFAULT_PAGE_SIZE = 50; // Change this
```

### Quick Testing

```bash
# Local workbench (limited functionality)
gulp serve

# SharePoint workbench (full functionality)
gulp serve --nobrowser
# Then go to: https://{tenant}.sharepoint.com/_layouts/workbench.aspx
```

---

## Common Use Cases

### Use Case 1: Find a Colleague
**Goal**: Find John Smith in Marketing

**Steps**:
1. Open People Directory
2. Type "john smith"
3. Click filter icon
4. Select "Marketing" department
5. Click John's card to view details
6. Click email icon to contact

**Time**: 30 seconds

---

### Use Case 2: Find All People in a Department
**Goal**: See everyone in IT Department

**Steps**:
1. Open People Directory
2. Type "it" (or any IT person's name)
3. Click filter icon
4. Select "IT" department
5. View all IT team members

**Time**: 20 seconds

---

### Use Case 3: Find Someone by Office Location
**Goal**: Find colleagues in New York office

**Steps**:
1. Open People Directory
2. Type any search term (optional)
3. Click filter icon
4. Select "New York" location
5. View all New York employees

**Time**: 20 seconds

---

### Use Case 4: Quick Contact
**Goal**: Quickly call or email someone

**Steps**:
1. Search for person
2. On the user card:
   - Click email icon → opens email client
   - Click phone icon → initiates call
3. Or click card for full contact details

**Time**: 15 seconds

---

## Keyboard Shortcuts

| Key | Action |
|-----|--------|
| `/` | Focus search box |
| `Tab` | Navigate between elements |
| `Enter` | Open selected user card |
| `Esc` | Close details panel |
| `Ctrl/Cmd + F` | Focus search (browser default) |

---

## Mobile Usage

### Mobile-Specific Tips

📱 **On Phone**:
- Search bar is full width
- User cards stack vertically
- Swipe to scroll
- Tap to view details

📱 **On Tablet**:
- Two-column card layout
- Side panel for filters
- Optimized touch targets

---

## Getting Help

### Self-Service Help

1. **Check the FAQ** (in full documentation)
2. **Try a different search term**
3. **Clear filters and try again**
4. **Refresh the page**

### Contact Support

If issues persist:

📧 **Email**: support@company.com
📞 **Phone**: (555) 123-4567
💬 **Teams**: IT Support Team

**Include in your support request**:
- What you were trying to do
- What happened instead
- Screenshot (if applicable)
- Browser and device type

---

## Pro Tips

💡 **Tip 1**: First search may be slower (cache warming)
💡 **Tip 2**: Use filters to narrow large result sets
💡 **Tip 3**: Profile details are always current from Azure AD
💡 **Tip 4**: Photos may take a moment to load (lazy loading)
💡 **Tip 5**: Try searching by department for team discovery

---

## Next Steps

✅ **For Users**: Start searching for colleagues!
✅ **For Admins**: See DEPLOYMENT.md for advanced configuration
✅ **For Developers**: See README.md for development guide

---

## Quick Reference Card (Print This!)

```
PEOPLE DIRECTORY - QUICK REFERENCE

SEARCH:
□ Type 2+ characters
□ Search name, email, dept

FILTER:
□ Click filter icon
□ Select dept or location
□ Combine filters

VIEW DETAILS:
□ Click any card
□ See contact info
□ Click email/phone icons

TIPS:
□ First search = slower
□ Use filters for large results
□ Info always current from AD

SUPPORT:
support@company.com | (555) 123-4567
```

---

**Questions?** See the full README.md or contact your IT department.
