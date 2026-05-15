# Changelog

All notable changes to the Enhanced Quality Evaluation Grid are documented here.

## [1.0.15] — 2026-05-15

### Fixed
- Date range filter operator: corrected `LessThan` condition operator from `6` (Like) to `3` (LessThan), fixing "Next week" and "This month" card filters.

## [1.0.14] — 2026-05-15

### Added
- **Clickable KPI cards**: All three smart cards now act as interactive filters.
  - Critical questions card: click "Pass" or "Fail" to filter the grid.
  - Evaluator expiration date card: click "This week", "Next week", or "This month" to filter by date range.
- Active card segments show visual highlight (colored background for critical, blue for dates).
- All card filters combine with existing dropdown/chip/score filters.
- "Clear all" button resets card filters.

### Fixed
- Aligned card headers to the same vertical baseline across all three cards.

## [1.0.13] — 2026-05-15

### Changed
- **Critical questions card**: Refined from uppercase "PASS/FAIL" to sentence case with colored dot indicators and subtle `·` separator.
- **Evaluator expiration date card**: Replaced `>` separator with vertical segments — bold number on top, muted label below, thin dividers between.

## [1.0.12] — 2026-05-15

### Changed
- Moved smart cards from a separate row above the filter bar into the filter bar itself, positioned to the right with `marginLeft: auto`.
- Reduced card padding and font sizes for a compact, integrated look.

## [1.0.11] — 2026-05-15

### Added
- **Smart KPI cards**: Three cards displayed above the grid:
  - Average score (FetchXML `avg` aggregate on `msdyn_score`).
  - Critical question pass/fail breakdown (grouped count on `msdyn_criticalquestionstatus`).
  - Evaluator expiration date bucketed by this week / next week / this month.
- **Score range chip filters**: Failed (0–60), Medium (61–85), Excellent (86–100) with color-coded styling.
- Cards re-fetch on every filter change to reflect the filtered dataset.
- All 5 aggregate queries execute in parallel.

### Changed
- Extended `onFilter` to support score range conditions (`ge`/`le`).

## [1.0.10] — 2026-05-15

### Changed
- Record type chips now sourced from entity metadata (all possible values) instead of extracting from the current page only.
- Server-side filtering for record type when metadata is available.
- `applyFilters()` function combines dropdown filters with record type for unified server-side filtering.

## [1.0.9] — 2026-05-15

### Added
- Record type chip/pill filters always visible in the filter bar.

## [1.0.8] — 2026-05-15

### Added
- Styled Fluent UI dropdowns with rounded corners and compact sizing.
- Dynamic filter bar auto-discovers all Picklist/State columns from metadata.

## [1.0.7] — 2026-05-15

### Added
- Filter bar with dropdown filters for status columns.

## [1.0.6] — 2026-05-15

### Added
- D365 option-set colors fetched from entity metadata at runtime.
- Color priority: metadata > ColorConfig JSON > defaults.
- Contrast-aware text color calculation.

## [1.0.5] — 2026-05-15

### Fixed
- Publisher name corrected to "Alexander Yurpolsky".

## [1.0.4] — 2026-05-15

### Changed
- Control renamed to "Enhanced Quality Evaluation Grid".

## [1.0.3] — 2026-05-15

### Fixed
- Side pane navigation conflict: added 1500ms delay after `openForm()` to prevent UCI error 2415919106.
- Used `window.top.setTimeout` to survive PCF iframe teardown.
- Captured `Xrm` reference before navigation to avoid stale context.

## [1.0.2] — 2026-05-15

### Added
- English localization for footer (Total records, Page, Previous, Next).
- RTL-aware chevron direction for Hebrew.

## [1.0.1] — 2026-05-15

### Added
- Initial release with core grid functionality.
- Color-coded score pills and status badges.
- Smart entity icons (Case, Email, Conversation with voice/chat detection).
- User avatar rendering via D365 entity image endpoint.
- Side pane drill-through for evaluation details.
- Hebrew footer localization.
- Configurable colors via `ColorConfig` JSON property.
