# Changelog

All notable changes to the People Directory WebPart will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [1.0.0] - 2024-01-XX

### Added
- Initial release of People Directory WebPart
- Real-time search across Azure AD users via Microsoft Graph API
- Advanced filtering by department and office location
- Multi-tier caching strategy (IndexedDB, SharePoint List, Graph API)
- Responsive user card layout with profile photos
- Detailed user profile panel with contact information
- Click-to-email and click-to-call functionality
- Lazy loading of profile pictures
- Automatic SharePoint list provisioning for caching
- Debounced search with 300ms delay
- LRU cache eviction for SharePoint list
- Error handling with user-friendly messages
- Accessibility compliance (WCAG 2.1 AA)
- Mobile-responsive design
- Support for 50,000+ users
- Real-time data verification from Azure AD
- Manager information display
- Fluent UI React components integration
- Comprehensive documentation
- Deployment guide
- Architecture documentation

### Technical Details
- SPFx version 1.18.2
- React 17.0.1
- TypeScript 4.7.4
- Fluent UI React 8.110+
- PnP JS 3.22+
- Microsoft Graph integration
- IndexedDB for client caching
- SharePoint list for tenant-wide caching

### Security
- Minimal API permissions (User.ReadBasic.All, User.Read.All)
- Secure token-based authentication
- No PII storage beyond cache TTL
- Rate limiting with exponential backoff
- Error messages don't expose internal details

### Performance
- Search debouncing (300ms)
- Minimum search length (2 characters)
- Virtual scrolling support
- Lazy image loading
- Batch API requests
- Three-tier caching system
- Client cache: 1,000 profiles, 30-minute TTL
- List cache: 5,000 profiles, 24-hour TTL

### Known Limitations
- SharePoint list limited to 5,000 cached items
- Profile photos require User.Read.All permission
- Department/location lists limited to 1,000 items for performance
- First search may be slower (cache warm-up)

## [Unreleased]

### Planned Features
- Export to Excel functionality
- Organizational chart visualization
- Teams presence indicators (online/offline)
- Favorites and recent contacts
- Azure Function for bulk overnight synchronization
- Department hierarchy tree view
- Advanced analytics and usage tracking
- Custom field mapping
- Multi-language support
- Dark theme support
- Print-friendly view

### Under Consideration
- Integration with Teams chat
- Integration with Outlook
- Custom profile fields from Azure AD extensions
- Bulk operations (export, share)
- Saved searches
- Email distribution list creation from search results

## Version History

### [1.0.0] - Initial Release
First production-ready version with core functionality.

---

## Release Notes

### Version 1.0.0 Release Notes

**Release Date**: January 2024

**What's New**:
This is the initial release of the enterprise People Directory WebPart for SharePoint Online.

**Key Features**:
1. **Search**: Fast, real-time search across your entire organization
2. **Filters**: Advanced filtering by department and location
3. **Profiles**: Detailed user profiles with contact information
4. **Performance**: Handles 50,000+ users with intelligent caching
5. **Mobile**: Fully responsive design for all devices
6. **Accessible**: WCAG 2.1 AA compliant

**Installation**:
See DEPLOYMENT.md for detailed installation instructions.

**Upgrade Path**:
Not applicable (initial release)

**Breaking Changes**:
None (initial release)

**Bug Fixes**:
None (initial release)

**Deprecations**:
None

**Dependencies**:
- SharePoint Online
- Microsoft 365 with Azure AD
- Node.js 16.13.0+ or 18.17.1+ (for development)

**Support**:
For issues, questions, or feature requests, please contact your IT administrator or development team.

---

## Contributing

When contributing, please:
1. Update this CHANGELOG with your changes
2. Follow the Keep a Changelog format
3. Include version numbers and dates
4. Categorize changes appropriately (Added, Changed, Fixed, Deprecated, Removed, Security)

## Questions?

For questions about versioning or releases, contact the development team.
