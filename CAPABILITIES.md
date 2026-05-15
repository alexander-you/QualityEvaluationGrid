# Capabilities

A detailed reference of every feature in the Enhanced Quality Evaluation Grid.

---

## 1. Interactive KPI Cards

Three smart cards are displayed in the filter bar area, providing real-time aggregate metrics across all records matching the current filters.

### Average Score Card
- Displays the mean evaluation score computed via a FetchXML `avg` aggregate on `msdyn_score`.
- Renders as a bold number (e.g., **78**).
- Updates automatically when any filter changes.

### Critical Questions Card
- Shows the count of evaluations grouped by critical question pass/fail status.
- Computed via a FetchXML grouped `count` aggregate on `msdyn_criticalquestionstatus`.
- Pass/fail classification is determined by matching option-set labels (case-insensitive contains "pass" or "fail").
- **Interactive**: Click "Pass" to filter the grid to only passing evaluations. Click "Fail" to filter to failing. Click again to deselect.
- Visual indicators: colored dot (green for pass, red for fail) with subtle background highlight on active state.

### Evaluator Expiration Date Card
- Displays three counts: evaluations due **this week**, **next week**, and **this month**.
- Week boundaries: Monday to Sunday (ISO standard).
- Month boundaries: first day of current month to first day of next month.
- Computed via three separate FetchXML count queries with date range filters (`ge`/`lt`).
- **Interactive**: Click any segment to filter the grid to that date range. Active segment highlights in blue. Click again to deselect.
- Segments are separated by thin vertical dividers for visual clarity.

### Card Behavior
- All cards execute **5 parallel FetchXML aggregate queries** on initial load and re-execute whenever filters change.
- Filter conditions applied in the filter bar are also applied to the card queries, so card metrics always reflect the filtered dataset.
- Cards are positioned to the right of the filter bar with `marginLeft: auto` for a clean split layout.
- Card headers are vertically aligned across all three cards.

---

## 2. Filter Bar

A comprehensive filtering system with multiple filter types, all performing server-side filtering via the D365 dataset filtering API.

### Dropdown Filters
- Automatically generated for every **Picklist**, **Status**, and **State** column in the view.
- Options are sourced from D365 entity metadata (all possible values, not just those on the current page).
- Styled with Fluent UI Dropdown: rounded corners, 160–220px width, 32px height.
- Default selection: "All" (no filter applied).

### Record Type Chip Filters
- Pill/chip-style toggle buttons for the record type field.
- When entity metadata is available: shows **all possible values** (e.g., Email, Conversation, Case) regardless of what's on the current page.
- When metadata is unavailable: falls back to extracting unique values from the current page.
- Server-side filtering when metadata is available; client-side filtering as fallback.
- Active state: blue border and background. Includes a cancel (×) icon for clarity.

### Score Range Chip Filters
- Three chips derived from the color configuration:
  - **Failed** (0–60) — Red styling
  - **Medium** (61–85) — Yellow styling
  - **Excellent** (86–100) — Green styling
- Applies server-side range filtering using `GreaterThanOrEqual` and `LessThanOrEqual` conditions.
- Active chip uses the corresponding score range color for border and background.

### Clear All
- A "Clear all" button appears when any filter is active.
- Resets all dropdown selections, chip toggles, score range, card filters, and date range in a single click.
- Clears all server-side filter conditions.

### Filter Composition
- All filter types combine using AND logic.
- Example: selecting "Conversation" (record type) + "Failed" (score range) + "This week" (due date) filters to evaluations of conversations that scored ≤60 and are due this week.

---

## 3. Data Grid

Built on Fluent UI's `DetailsList` with custom column renderers.

### Column Types

| Column Type | Rendering | Detection |
|-------------|-----------|-----------|
| **Primary** | Bold 14px link, triggers row click navigation | `isPrimary` flag from dataset |
| **Lookup (User)** | Fluent UI Persona with avatar image from D365 entity image endpoint | `dataType` starts with "Lookup" + entity type is `systemuser` |
| **Lookup (Other)** | Clickable link, opens the related record form | `dataType` starts with "Lookup", "Customer", or "Owner" |
| **Score** | Color-coded pill badge with `/100` suffix (e.g., `87/100`) | Field name contains "score" |
| **Status / Picklist** | Color-coded pill badge | Field name contains "status", or `dataType` is "State" or "Picklist" |
| **Record Type** | Smart entity icon + text label | Field name contains "recordtype" |
| **Other** | Plain text, 13px font | Default fallback |

### Color Resolution Priority
1. **D365 metadata colors** — pulled from option-set definition at runtime (highest priority).
2. **ColorConfig JSON** — user-provided override via the control property.
3. **Default color map** — hardcoded sensible defaults (lowest priority).

Contrast calculation ensures text remains readable: if the background luminance is high, text is dark (#333333); if low, text is white (#FFFFFF).

### Smart Entity Icons
For the record type column, icons are resolved based on the regarding object:

| Entity | Icon |
|--------|------|
| Case (`incident`) | TextDocument |
| Email | Mail |
| Phone Call | Phone |
| Conversation (Chat) | OfficeChat |
| Conversation (Voice) | Phone |
| Other | Page |

Conversation channel type is resolved via an async `Xrm.WebApi` call to `msdyn_ocliveworkitem.msdyn_channel`. Results are cached in a global `Map` to avoid redundant API calls.

### Grid Features
- **Column sorting**: Click any column header to sort. Supports ascending/descending toggle. Uses server-side sorting via `dataset.sorting`.
- **Column resizing**: All columns are resizable. Regarding column has a wider max (250px); others default to 150px.
- **Sticky headers**: Column headers remain visible during vertical scrolling via Fluent UI's `Sticky` + `ScrollablePane`.
- **Scrollable content**: Virtual scrolling via `ScrollablePane` with auto-visibility scrollbars.

---

## 4. Row Click Navigation

Clicking an evaluation row triggers a dual-navigation pattern:

1. **Main form**: Opens the source interaction record (conversation, case, or email) using `context.navigation.openForm()`. The target is determined by the `msdyn_regardingobjectid` lookup field.
2. **Side pane**: After a 1500ms delay, opens the evaluation record in a 340px-wide side pane using `Xrm.App.sidePanes`. The delay prevents a UCI navigation conflict (error 2415919106).

Side pane behavior:
- If a pane with ID `evaluationPane` already exists, it navigates within the existing pane.
- If not, a new pane is created with `canClose: true`, `alwaysRender: true`.
- The `Xrm` reference is captured before navigation to survive the PCF iframe teardown during form transition.
- `window.top.setTimeout` is used instead of `window.setTimeout` so the timer survives the PCF iframe being destroyed.

---

## 5. User Avatars

Lookup columns pointing to `systemuser` records render with Fluent UI's `Persona` component:
- Avatar image URL: `{clientUrl}/Image/download.aspx?Entity=systemuser&Attribute=entityimage&Id={userId}`
- Fallback: two-letter initials from the user's name.
- Size: 24px (compact).
- Cache-busted with a timestamp query parameter.

---

## 6. Paging Footer

A fixed footer at the bottom of the grid:
- **Left side**: "Total records: **2300** | Page: 1"
- **Right side**: Previous/Next page buttons with chevron icons.
- Total result count uses the actual server count when available (`paging.totalResultCount`), or falls back to `items.length+` for indeterminate counts.

### Localization
- **English** (default): "Total records", "Page", "Previous", "Next"
- **Hebrew** (language ID 1037): Localized strings with RTL-aware chevron direction (Previous = ChevronRight, Next = ChevronLeft).

---

## 7. Metadata Integration

The control fetches D365 entity metadata on first load to power several features:

| Metadata Type | Used For |
|---------------|----------|
| `PicklistAttributeMetadata` | Dropdown filter options, chip labels, option-set colors |
| `StatusAttributeMetadata` | Status field colors and labels |
| `StateAttributeMetadata` | State field colors and labels |
| `EntityDefinitions.EntitySetName` | FetchXML aggregate query endpoint |

All metadata is fetched once per session and cached in memory. Three metadata API calls run in parallel.

---

## 8. Configuration

### Control Properties

| Property | Type | Required | Description |
|----------|------|----------|-------------|
| `dataSet` | Dataset | Yes | The bound dataset (evaluation records) |
| `ColorConfig` | SingleLine.Text | No | JSON string for custom color overrides |

### Required Platform Features

| Feature | Purpose |
|---------|---------|
| `Utility` | Access to `context.mode`, `context.userSettings` |
| `WebAPI` | Access to `Xrm.WebApi` for smart icon resolution |

### Manifest

- **Namespace**: `YurpolskyV2`
- **Constructor**: `QualityEvaluationGrid`
- **Control type**: Standard (dataset-bound)
- **Current version**: 1.0.15
