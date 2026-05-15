# Enhanced Quality Evaluation Grid

A custom **PowerApps Component Framework (PCF)** control for **Dynamics 365 Customer Service** that replaces the default evaluation subgrid with a feature-rich, interactive data grid purpose-built for quality management workflows.

![Demo](QualityEvaluationGrid.mp4)

## Key Benefits

- **At-a-glance performance visibility** — Color-coded score pills, status badges, and smart entity icons make it easy to scan hundreds of evaluations and spot issues instantly.
- **Real-time KPI cards** — Aggregate metrics (average score, critical question pass/fail, evaluator deadlines) are computed server-side via FetchXML and update live as filters change.
- **Powerful multi-dimensional filtering** — Dropdown filters, chip selectors, score range chips, and clickable KPI cards work together to let managers build precise queries without leaving the grid.
- **Contextual drill-through** — Clicking an evaluation opens the source interaction (conversation, case, or email) as the main form while simultaneously displaying the evaluation in a side pane — no context switching.
- **Zero-configuration colors** — Option-set colors are automatically pulled from D365 metadata. No JSON configuration required for most deployments.
- **Bilingual support** — English and Hebrew (RTL-aware) footer localization out of the box.

## Prerequisites

- Dynamics 365 Customer Service with the Quality Evaluation Agent (QEA) feature
- Node.js 18+ and npm
- .NET SDK 6.0+
- [Power Platform CLI (`pac`)](https://learn.microsoft.com/en-us/power-platform/developer/cli/introduction)

## Quick Start

### 1. Clone and install

```bash
git clone https://github.com/YourUsername/QualityEvaluationGrid.git
cd QualityEvaluationGrid
npm install
```

### 2. Build the control

```bash
npm run build
```

### 3. Build the solution

```bash
cd Solution
dotnet build -c Release
```

### 4. Deploy to Dynamics 365

```bash
# Authenticate (first time only)
pac auth create --url https://your-org.crm.dynamics.com

# Import
pac solution import --path Solution/bin/Release/Solution.zip --force-overwrite --publish-changes

# Clean up the transport solution
pac solution delete --solution-name Solution
```

### 5. Configure the control

1. Open a model-driven app in the maker portal.
2. Navigate to the view or form where evaluation records are displayed.
3. Replace the default subgrid with the **Enhanced Quality Evaluation Grid** custom control.
4. (Optional) Provide a `ColorConfig` JSON string to override default score/status colors.

## Color Configuration

The control works out of the box with sensible defaults and D365 metadata colors. For custom overrides, pass a JSON string via the `ColorConfig` property:

```json
{
  "status": {
    "0": { "bg": "#E0F1FF", "color": "#0078D4" },
    "1": { "bg": "#E6FFEC", "color": "#107C10" },
    "2": { "bg": "#FDE7E9", "color": "#D13438" }
  },
  "score": [
    { "max": 60, "style": { "bg": "#FDE7E9", "color": "#A80000" } },
    { "max": 85, "style": { "bg": "#FFF4CE", "color": "#795900" } },
    { "max": 100, "style": { "bg": "#DFF6DD", "color": "#107C10" } }
  ]
}
```

| Key | Purpose |
|-----|---------|
| `status.<value>` | Maps a raw option-set value to a background/text color pair |
| `score[].max` | Upper bound of a score range |
| `score[].style` | Color pair applied to scores within this range |

D365 metadata colors (configured in the option-set definition) take priority over `ColorConfig` for Picklist fields.

## Development

```bash
# Start the PCF test harness
npm start

# Start with hot reload
npm run start:watch

# Lint
npm run lint

# Clean build artifacts
npm run clean
```

## Tech Stack

| Layer | Technology |
|-------|-----------|
| Framework | PowerApps Component Framework (PCF) |
| UI | React 17 + Fluent UI React 8 |
| Language | TypeScript 5.8 |
| Build | pcf-scripts + Webpack |
| Data | D365 Web API, FetchXML aggregates, dataset filtering API |
| Deployment | .NET SDK + Power Platform CLI |

## Project Structure

```
QualityEvaluationGrid/
├── QualityGrid/
│   ├── ControlManifest.Input.xml   # PCF manifest (version, properties, features)
│   ├── index.ts                     # PCF lifecycle (init, updateView, destroy)
│   ├── QualityGrid.tsx              # React component (grid, filters, cards, footer)
│   ├── css/QualityGrid.css          # Base styles
│   ├── generated/ManifestTypes.d.ts # Auto-generated type definitions
│   └── strings/QualityGrid.1033.resx # Display name resource
├── Solution/
│   ├── Solution.cdsproj             # Solution project
│   └── src/Other/Solution.xml       # Solution metadata
├── package.json
├── tsconfig.json
├── eslint.config.mjs
└── pcfconfig.json
```

## Documentation

| Document | Description |
|----------|-------------|
| [ARCHITECTURE.md](ARCHITECTURE.md) | Solution concept, data flow, and component design |
| [CAPABILITIES.md](CAPABILITIES.md) | Detailed feature reference with technical specifics |
| [CHANGELOG.md](CHANGELOG.md) | Version history and release notes |

## Author

**Alexander Yurpolsky**

## License

This project is licensed under the MIT License — see the [LICENSE](LICENSE) file for details.
