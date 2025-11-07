# People Directory WebPart

Enterprise-grade SharePoint Framework (SPFx) web part for searching and discovering people across your organization with real-time data from Microsoft Graph API.

## Features

### Core Functionality
- **Real-time Search**: Search across 50,000+ users by name, email, department, or job title
- **Advanced Filtering**: Filter by department and office location
- **Profile Details**: View comprehensive user profiles including contact information and organizational hierarchy
- **Multi-tier Caching**: Intelligent caching strategy for optimal performance
- **Mobile Responsive**: Fully responsive design for all devices
- **Accessible**: WCAG 2.1 AA compliant

### Technical Highlights
- **Microsoft Graph Integration**: Real-time data from Azure AD
- **Smart Caching**: 3-tier caching (IndexedDB → SharePoint List → Graph API)
- **Performance Optimized**: Handles organizations with 50,000+ users
- **Real-time Verification**: Always shows current data from Azure AD
- **Profile Photos**: Lazy-loaded profile pictures with fallback avatars
- **Click-to-Action**: Direct email and phone call integration

## Architecture

### Data Flow
```
User Search
    ↓
IndexedDB (30 min cache)
    ↓
SharePoint List (24 hr cache, 5000 items)
    ↓
Microsoft Graph API (Source of Truth)
```

### Technology Stack
- **SPFx**: 1.18.2
- **React**: 17.0.1
- **Fluent UI**: 8.110+
- **PnP JS**: 3.22+
- **TypeScript**: 4.7.4

## Prerequisites

- Node.js v16.13.0+ or v18.17.1+
- SharePoint Online tenant
- Microsoft 365 tenant with Azure AD
- SharePoint Online Administrator access (for deployment)

## Installation

### 1. Clone and Install Dependencies
```bash
git clone <repository-url>
cd CompanyDirectory
npm install
```

### 2. Build the Solution
```bash
# Development build
npm run build

# Production build (for deployment)
npm run package
```

### 3. Deploy to SharePoint
1. Navigate to `solution/people-directory-webpart.sppkg`
2. Upload to SharePoint App Catalog
3. Deploy globally or to specific sites
4. **IMPORTANT**: Approve API permissions in SharePoint Admin Center

## API Permissions Required

After deployment, approve these permissions in **SharePoint Admin Center → API Management**:

| Resource | Permission | Reason |
|----------|-----------|--------|
| Microsoft Graph | User.ReadBasic.All | Read basic user profile information |
| Microsoft Graph | User.Read.All | Read detailed user profiles (photos, manager) |

### How to Approve Permissions

1. Go to SharePoint Admin Center (https://tenant-admin.sharepoint.com)
2. Navigate to **Advanced → API access**
3. Find "People Directory" pending requests
4. Approve both Graph API permissions

## Usage

### Adding to a SharePoint Page

1. Edit a SharePoint page
2. Click "+" to add a web part
3. Search for "People Directory"
4. Add the web part to the page
5. Configure settings in the property pane

### Configuration Options

**Web Part Properties:**
- **Title**: Display title for the web part
- **Description**: Brief description shown below the title

### User Features

**Search:**
- Type at least 2 characters to start searching
- Search includes: Display name, email, surname, given name, department

**Filters:**
- Click the filter icon to open advanced filters
- Filter by department or office location
- Filters can be combined

**View Details:**
- Click any user card to view full profile details
- Details panel shows:
  - Contact information (email, phone)
  - Organizational information (department, location)
  - Manager information
  - Last verification timestamp

**Quick Actions:**
- Click email icon to compose email
- Click phone icons to initiate calls

## Development

### Local Development
```bash
# Start local workbench
npm run serve

# The workbench will be available at:
# https://localhost:4321/temp/workbench.html
```

### Testing with SharePoint
```bash
# Serve to SharePoint workbench
gulp serve --nobrowser

# Then navigate to:
# https://{your-tenant}.sharepoint.com/_layouts/workbench.aspx
```

### Project Structure
```
CompanyDirectory/
├── config/                    # SPFx configuration
├── src/
│   ├── models/               # TypeScript interfaces
│   │   ├── IUserProfile.ts
│   │   └── Constants.ts
│   ├── services/             # Business logic layer
│   │   ├── GraphService.ts   # Microsoft Graph API
│   │   ├── ListService.ts    # SharePoint list operations
│   │   └── PeopleService.ts  # Orchestration service
│   ├── utils/                # Utility functions
│   │   ├── CacheHelper.ts    # IndexedDB caching
│   │   └── ErrorHandler.ts   # Error handling
│   └── webparts/
│       └── peopleDirectory/
│           ├── components/   # React components
│           │   ├── PeopleDirectory.tsx
│           │   ├── UserCard.tsx
│           │   └── UserDetailsPanel.tsx
│           └── PeopleDirectoryWebPart.ts
├── ARCHITECTURE.md           # Architecture documentation
└── README.md                # This file
```

## Performance Optimization

### Caching Strategy
The solution implements a 3-tier caching strategy:

1. **IndexedDB (Client-side)**
   - Capacity: 1,000 profiles
   - TTL: 30 minutes
   - Fastest access

2. **SharePoint List (Tenant-wide)**
   - Capacity: 5,000 profiles
   - TTL: 24 hours
   - LRU eviction policy

3. **Microsoft Graph API**
   - Source of truth
   - Always verified on profile view

### Search Optimization
- Client-side debouncing (300ms)
- Minimum search length (2 characters)
- Server-side filtering via Graph API
- Lazy loading of profile pictures

### Scalability
- Handles 50,000+ users
- Virtual scrolling for large result sets
- Batch API requests
- Asynchronous photo loading

## Troubleshooting

### Common Issues

**1. "Insufficient permissions" error**
- Solution: Approve API permissions in SharePoint Admin Center

**2. SharePoint list not created**
- Solution: Ensure user has sufficient permissions to create lists
- The list is created automatically on first use

**3. Photos not loading**
- Solution: Verify User.Read.All permission is approved
- Photos have a fallback to initials

**4. Slow search performance**
- Solution: Wait for cache to warm up (first searches populate cache)
- Subsequent searches will be faster

**5. Rate limiting errors**
- Solution: Implement retry logic (already built-in)
- Consider reducing batch sizes in Constants.ts

## Security Considerations

- **Minimal Permissions**: Uses User.ReadBasic.All by default
- **No PII Storage**: Client cache clears after 30 minutes
- **Secure Communication**: All requests use HTTPS
- **Token-based Auth**: Uses SharePoint Framework's authentication
- **Rate Limiting**: Built-in exponential backoff
- **Error Handling**: Never exposes internal errors to users

## Browser Support

- Microsoft Edge (Chromium)
- Google Chrome
- Firefox
- Safari 14+
- Mobile browsers (iOS Safari, Chrome Mobile)

## Accessibility

- WCAG 2.1 AA compliant
- Keyboard navigation support
- Screen reader compatible
- High contrast mode support
- Focus indicators

## Future Enhancements

Potential features for future releases:
- Export to Excel
- Organizational chart visualization
- Teams presence integration
- Advanced analytics
- Azure Function for bulk sync
- Department hierarchy visualization
- Favorites and recent contacts

## Support and Contribution

### Reporting Issues
Please report issues with detailed steps to reproduce, environment details, and screenshots if applicable.

### Contributing
1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Submit a pull request

## License

Copyright (c) 2024. All rights reserved.

## Version History

### Version 1.0.0 (Initial Release)
- Core search and filter functionality
- Microsoft Graph integration
- Multi-tier caching
- Profile detail view
- Mobile responsive design
- Accessibility compliance

## Additional Resources

- [SharePoint Framework Documentation](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/sharepoint-framework-overview)
- [Microsoft Graph API](https://docs.microsoft.com/en-us/graph/overview)
- [Fluent UI React](https://developer.microsoft.com/en-us/fluentui#/controls/web)
- [PnP JS Documentation](https://pnp.github.io/pnpjs/)
