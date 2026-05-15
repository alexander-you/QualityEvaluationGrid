# Architecture

This document describes the solution concept, component design, and data flow of the Enhanced Quality Evaluation Grid.

## Solution Concept

The control is a **standard dataset-bound PCF control** that replaces Dynamics 365's native subgrid renderer. It intercepts the platform's dataset API to read records, columns, sorting, paging, and filtering — then renders a custom React UI that provides richer visualization and interaction patterns than the default grid.

```
┌──────────────────────────────────────────────────────────────────┐
│                    Dynamics 365 UCI Shell                        │
│                                                                  │
│  ┌────────────────────────────────────────────────────────────┐  │
│  │  PCF Control (QualityEvaluationGrid)                       │  │
│  │                                                            │  │
│  │  ┌─────────────┐    ┌──────────────────────────────────┐  │  │
│  │  │  index.ts    │    │  QualityGrid.tsx (React)         │  │  │
│  │  │  PCF         │───▶│                                  │  │  │
│  │  │  Lifecycle   │    │  ┌────────┐ ┌────────────────┐  │  │  │
│  │  │              │    │  │ Cards  │ │  Filter Bar     │  │  │  │
│  │  │  • init      │    │  └────────┘ └────────────────┘  │  │  │
│  │  │  • updateView│    │  ┌──────────────────────────────┐│  │  │
│  │  │  • destroy   │    │  │  DetailsList (Fluent UI)     ││  │  │
│  │  │              │    │  │  • Columns with renderers    ││  │  │
│  │  │  Metadata    │    │  │  • Sticky header             ││  │  │
│  │  │  fetches     │    │  │  • ScrollablePane            ││  │  │
│  │  │  (parallel)  │    │  └──────────────────────────────┘│  │  │
│  │  │              │    │  ┌──────────────────────────────┐│  │  │
│  │  │  Aggregate   │    │  │  Footer (paging + counts)    ││  │  │
│  │  │  queries     │    │  └──────────────────────────────┘│  │  │
│  │  └─────────────┘    └──────────────────────────────────┘  │  │
│  └────────────────────────────────────────────────────────────┘  │
│                                                                  │
│  ┌──────────────┐  ┌──────────────┐  ┌──────────────────────┐   │
│  │  D365 Web API │  │ Dataset API  │  │  Xrm.App.sidePanes  │   │
│  │  (metadata,   │  │ (records,    │  │  (evaluation         │   │
│  │   aggregates) │  │  filtering)  │  │   drill-through)     │   │
│  └──────────────┘  └──────────────┘  └──────────────────────┘   │
└──────────────────────────────────────────────────────────────────┘
```

## Component Design

### index.ts — PCF Lifecycle Controller

Responsibilities:
- **init()**: Captures container reference, enables resize tracking, resolves `clientUrl`.
- **updateView()**: Orchestrates the full render cycle on every dataset change:
  1. Reads dataset columns, records, sorting, and paging state.
  2. Triggers one-time metadata fetches (option colors, entity set name).
  3. Builds `colorConfig` from defaults + optional user JSON override.
  4. Computes `scoreRanges` from the color config.
  5. Defines callback handlers (`onSort`, `onFilter`, `onRowClick`, `onLookupClick`).
  6. Renders the React component via `ReactDOM.render()`.
- **destroy()**: Unmounts React, closes the side pane if open.

Private methods:
- **_fetchOptionColors()**: Parallel metadata queries for Picklist, Status, and State attributes. Extracts option values, labels, and colors.
- **_fetchEntitySetName()**: Resolves the OData entity set name for FetchXML aggregate queries.
- **_fetchCardData()**: Executes 5 parallel FetchXML aggregate queries for KPI cards. Respects current filter state. Deduplicates via filter hash.

### QualityGrid.tsx — React UI Component

Responsibilities:
- **Filter bar**: Renders Fluent UI Dropdowns for Picklist/State columns, chip selectors for record type and score ranges, and a "Clear all" button.
- **KPI cards**: Three interactive cards (average score, critical questions, evaluator due dates) positioned to the right of the filter bar. Each card segment is clickable and applies server-side filters.
- **DetailsList**: The main data grid with custom column renderers for:
  - Primary links (bold, clickable)
  - Lookup fields (user avatars via Persona, or clickable links)
  - Score fields (color-coded pills with `/100` suffix)
  - Status/Picklist fields (color-coded pills from D365 metadata)
  - Record type fields (smart entity icons resolved from the regarding object)
- **Footer**: Paging controls with total record count, bilingual support.

### SmartEntityIcon — Async Icon Resolver

A React component that determines the correct Fluent UI icon for a record type by:
1. Checking the entity type (Case → TextDocument, Email → Mail, PhoneCall → Phone).
2. For conversations (`msdyn_ocliveworkitem`): querying `msdyn_channel` via `Xrm.WebApi` to distinguish Voice (Phone) from Chat (OfficeChat).
3. Caching results in a global `Map` to avoid repeated API calls.

## Data Flow

### Initial Load

```
updateView() called
  ├─ Dataset not loaded? → return (wait for next call)
  ├─ First call?
  │   ├─ _fetchOptionColors() ──▶ 3 parallel metadata API calls
  │   ├─ _fetchEntitySetName() ─▶ 1 metadata API call
  │   └─ _fetchCardData([]) ────▶ 5 parallel FetchXML aggregate queries
  ├─ Build columnsConfig from dataset.columns
  ├─ Build items from dataset.sortedRecordIds
  ├─ Compute scoreRanges from colorConfig
  └─ ReactDOM.render(QualityGrid, props)
```

### Filter Interaction

```
User clicks filter (dropdown / chip / card segment)
  ├─ React state updated (activeFilters, activeRecordType, activeScoreRange, etc.)
  ├─ applyFilters() called → assembles combined filter conditions
  ├─ props.onFilter() called in index.ts
  │   ├─ Builds FilterExpression conditions (eq, ge, le, lt)
  │   ├─ dataset.filtering.setFilter() or clearFilter()
  │   ├─ _fetchCardData(conditions) → 5 parallel aggregate queries with same filters
  │   └─ dataset.refresh() → triggers new updateView() cycle
  └─ Grid re-renders with filtered data + updated card metrics
```

### Row Click Navigation

```
User clicks evaluation row
  ├─ Extract regarding object (entity type + ID)
  ├─ context.navigation.openForm() → opens source record (conversation/case/email)
  └─ window.top.setTimeout(1500ms)
      └─ Xrm.App.sidePanes.createPane() or navigate()
          └─ Opens evaluation details in side pane (340px width)
```

The 1500ms delay is intentional — it prevents a UCI navigation conflict (error 2415919106) that occurs when `openForm()` and `sidePanes.navigate()` fire simultaneously.

## API Dependencies

| API | Purpose | Frequency |
|-----|---------|-----------|
| `EntityDefinitions/Attributes/PicklistAttributeMetadata` | Option colors and labels | Once per session |
| `EntityDefinitions/Attributes/StatusAttributeMetadata` | Status option colors | Once per session |
| `EntityDefinitions/Attributes/StateAttributeMetadata` | State option colors | Once per session |
| `EntityDefinitions?$select=EntitySetName` | OData entity set name | Once per session |
| FetchXML aggregate (avg score) | KPI card data | On each filter change |
| FetchXML aggregate (critical status counts) | KPI card data | On each filter change |
| FetchXML aggregate (due date counts × 3) | KPI card data | On each filter change |
| `Xrm.WebApi.retrieveRecord` (msdyn_channel) | Smart icon resolution | Once per conversation record |
| `dataset.filtering.setFilter()` | Server-side grid filtering | On each filter change |
| `Xrm.App.sidePanes` | Side pane navigation | On row click |

## Security Considerations

- All API calls use `credentials: 'include'` — they inherit the user's D365 session and security roles.
- No data is stored outside D365. The control is stateless across sessions.
- The `ColorConfig` JSON input is parsed with `try/catch` fallback to prevent injection of malformed config.
- Smart icon cache is in-memory only and cleared on page refresh.
