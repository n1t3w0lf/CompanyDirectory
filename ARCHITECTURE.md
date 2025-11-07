# People Directory WebPart - Architecture Document

## Executive Summary
Enterprise-grade People Directory SPFx WebPart for organizations with 50,000+ users across 1,000+ departments.

## Architecture Overview

### Data Flow Architecture
```
User Search Input
    ↓
Client-Side Debouncing (300ms)
    ↓
Check Local Cache (IndexedDB) ←─────┐
    ↓ (if expired/missing)           │
Microsoft Graph API Query            │
    ↓                                │
Update SharePoint List Cache         │
    ↓                                │
Update IndexedDB ────────────────────┘
    ↓
Display Results (Virtual Scroll)
    ↓
On User Card Click
    ↓
Real-time Graph API Verification
    ↓
Update Cache if Changed
    ↓
Display User Details
```

## Component Architecture

### 1. Data Layer
- **Microsoft Graph Client**: Primary data source for user profiles
- **SharePoint List**: Intelligent cache (5,000 most accessed profiles)
- **IndexedDB**: Client-side session cache (30-minute TTL)
- **PnP JS**: SharePoint list operations

### 2. Service Layer
- **GraphService**: User profile retrieval, search, photo fetching
- **CacheService**: Multi-tier caching strategy
- **ListService**: SharePoint list provisioning and CRUD
- **SyncService**: Background synchronization logic

### 3. Presentation Layer
- **PeopleDirectoryWebPart**: Main web part component
- **SearchBar**: Debounced search input with suggestions
- **FilterPanel**: Advanced filtering (department, location, etc.)
- **PeopleGrid**: Virtual scrolling list of user cards
- **UserCard**: Individual profile card with lazy-loaded photo
- **UserDetailsPanel**: Detailed view with real-time verification

## Technical Specifications

### SharePoint List Schema
**List Name**: PeopleDirectoryCache

| Field Name | Type | Indexed | Notes |
|------------|------|---------|-------|
| UserPrincipalName | Single line text | Yes | Primary key |
| DisplayName | Single line text | Yes | For search |
| Email | Single line text | Yes | Contact info |
| Department | Single line text | Yes | For filtering |
| JobTitle | Single line text | Yes | For filtering |
| OfficeLocation | Single line text | Yes | For filtering |
| BusinessPhones | Multi-line text | No | JSON array |
| MobilePhone | Single line text | No | Contact info |
| ProfilePhotoUrl | Hyperlink | No | Stored in SharePoint or Graph URL |
| LastVerified | Date/Time | Yes | For cache invalidation |
| AccessCount | Number | Yes | For LRU eviction |
| Manager | Single line text | No | Display name |

### Microsoft Graph Permissions Required
- User.Read.All
- User.ReadBasic.All (minimum)

### Performance Targets
- Initial load: < 2 seconds
- Search response: < 1 second
- Filter application: < 500ms
- Photo loading: Lazy, < 100ms per image
- Scroll performance: 60 FPS

## Scalability Strategy

### 1. Caching Layers
**Tier 1 - IndexedDB (Client)**
- Capacity: 1,000 profiles
- TTL: 30 minutes
- Eviction: LRU

**Tier 2 - SharePoint List (Tenant)**
- Capacity: 5,000 profiles
- TTL: 24 hours
- Eviction: LRU based on AccessCount

**Tier 3 - Microsoft Graph (Source of Truth)**
- Always queried for real-time verification
- Batch requests where possible

### 2. Search Strategy
- Client-side filtering when < 100 results cached
- Server-side (Graph) for fresh searches
- Debouncing: 300ms
- Minimum search length: 2 characters

### 3. Photo Optimization
- Thumbnail size: 96x96 (Graph API parameter)
- Lazy loading with Intersection Observer
- Fallback to initials avatar
- Optional: Store in SharePoint document library for CDN caching

## Security Considerations

1. **API Permissions**: Minimal required permissions (User.ReadBasic.All)
2. **Data Privacy**: No PII stored client-side beyond session
3. **Secure Communication**: All HTTPS, token-based auth
4. **List Permissions**: Restrict to web part, not user-editable
5. **Rate Limiting**: Implement exponential backoff
6. **Error Handling**: Never expose internal errors to users

## Deployment Strategy

### Phase 1: MVP (Week 1-2)
- Basic search and display
- Microsoft Graph integration
- Simple SharePoint list cache
- Basic UI with Fluent components

### Phase 2: Optimization (Week 3)
- Multi-tier caching
- Virtual scrolling
- Advanced filters
- Photo optimization

### Phase 3: Enhancement (Week 4)
- Real-time verification
- Analytics/telemetry
- Mobile optimization
- Accessibility audit

## Future Enhancements
- Export to Excel
- Organizational chart view
- Teams integration
- Presence indicators (online/offline)
- Azure Function for overnight sync
- Department hierarchy visualization
