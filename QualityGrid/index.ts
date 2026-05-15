import { IInputs, IOutputs } from "./generated/ManifestTypes";
import * as React from "react";
import * as ReactDOM from "react-dom";
import { QualityGrid, IGridItem, IColumnConfig, ICardData, IScoreRange, IDueDateRanges } from "./QualityGrid";

export class QualityEvaluationGrid implements ComponentFramework.StandardControl<IInputs, IOutputs> {
    private _container: HTMLDivElement;
    private _clientUrl: string;
    private _context: ComponentFramework.Context<IInputs>;
    private _optionColors: Record<string, Record<string, string>> = {};
    private _optionMeta: Record<string, Array<{ value: number; label: string; color: string }>> = {};
    private _optionColorsFetchStarted = false;
    private _entitySetName: string = '';
    private _cardData: ICardData | null = null;
    private _cardDataFetchInFlight = false;
    private _lastFilterHash: string = '';
    private _dueDateRanges: IDueDateRanges | null = null;

    constructor() {}

    public init(context: ComponentFramework.Context<IInputs>, notifyOutputChanged: () => void, state: ComponentFramework.Dictionary, container: HTMLDivElement): void {
        this._container = container;
        this._context = context;
        context.mode.trackContainerResize(true);
        this._container.style.height = "100%";
        this._container.style.position = "relative"; 
        try { this._clientUrl = (context as any).page?.getClientUrl() || ""; } catch (e) { this._clientUrl = ""; }
    }

    public updateView(context: ComponentFramework.Context<IInputs>): void {
        this._context = context;
        try {
            const dataset = context.parameters.dataSet;
            if (!dataset || dataset.loading) return;

            const gridEntityName = dataset.getTargetEntityType();

            // Fetch option-set colors and entity set name from D365 metadata (once per entity)
            if (!this._optionColorsFetchStarted && this._clientUrl) {
                this._optionColorsFetchStarted = true;
                Promise.all([
                    this._fetchOptionColors(gridEntityName),
                    this._fetchEntitySetName(gridEntityName)
                ]).then(() => {
                    // Initial card data fetch with no filters
                    this._fetchCardData([]);
                    this.updateView(this._context);
                }).catch(() => {});
            }

            // --- 1. הגדרת ברירת מחדל חזקה (כדי שהצבעים יעבדו גם בלי הגדרות) ---
            const defaultColorConfig = {
                "status": {
                    // ערכים סטנדרטיים
                    "0": { "bg": "#E0F1FF", "color": "#0078D4" }, // Active
                    "1": { "bg": "#E6FFEC", "color": "#107C10" }, // Completed
                    "2": { "bg": "#FDE7E9", "color": "#D13438" }, // Failed
                    
                    // ערכים ספציפיים שלך (Agent Status)
                    "700610000": { "bg": "#FFF4CE", "color": "#795900" }, 
                    "700610001": { "bg": "#E0F1FF", "color": "#0078D4" }, 
                    "700610002": { "bg": "#E6FFEC", "color": "#107C10" }, 
                    "700610003": { "bg": "#FDE7E9", "color": "#D13438" }, 
                    "700610004": { "bg": "#F3F2F1", "color": "#605E5C" }, 
                    "700610005": { "bg": "#F3F2F1", "color": "#605E5C" }
                },
                "score": [
                    { "max": 60, "style": { "bg": "#FDE7E9", "color": "#A80000" } },
                    { "max": 85, "style": { "bg": "#FFF4CE", "color": "#795900" } },
                    { "max": 100, "style": { "bg": "#DFF6DD", "color": "#107C10" } }
                ]
            };

            let colorConfig = defaultColorConfig;

            // --- 2. ניסיון טעינה מההגדרות בזהירות ---
            try {
                // בדיקה אם הפרמטר קיים בכלל (הגנה מפני גרסה ישנה)
                const params = context.parameters as any;
                if (params.ColorConfig && params.ColorConfig.raw) {
                    const configString = params.ColorConfig.raw;
                    if (configString.trim().length > 0) {
                        const userConfig = JSON.parse(configString);
                        // מיזוג: לוקחים את ברירת המחדל ודורסים עם מה שהמשתמש הזין
                        colorConfig = { 
                            ...defaultColorConfig, 
                            ...userConfig,
                            status: { ...defaultColorConfig.status, ...(userConfig.status || {}) }
                        };
                    }
                }
            } catch (e) {
                console.log("PCF: Error loading JSON config, using default colors.");
            }

            // --- המשך הקוד הרגיל ---
            const activeSort = (dataset.sorting && dataset.sorting.length > 0) ? dataset.sorting[0] : null;

            const columnsConfig: IColumnConfig[] = (dataset.columns || [])
                .filter(c => !c.isHidden && c.order >= 0)
                .sort((a, b) => a.order - b.order)
                .map(c => ({
                    fieldName: c.name,
                    displayName: c.displayName,
                    dataType: c.dataType,
                    minWidth: c.visualSizeFactor > 100 ? c.visualSizeFactor : 100,
                    isPrimary: c.isPrimary,
                    isSorted: activeSort ? activeSort.name === c.name : false,
                    isSortedDescending: activeSort ? activeSort.name === c.name && activeSort.sortDirection === 1 : false
                }));

            const items: IGridItem[] = (dataset.sortedRecordIds || []).map(id => {
                const record = dataset.records[id];
                const item: IGridItem = { key: id };
                columnsConfig.forEach(col => {
                    try {
                        item[col.fieldName] = record.getFormattedValue(col.fieldName);
                        item[col.fieldName + '_raw'] = record.getValue(col.fieldName);
                    } catch (ex) { item[col.fieldName] = ""; }
                });
                return item;
            });

            const onSort = (columnName: string, isDesc: boolean) => {
                dataset.sorting = [{ name: columnName, sortDirection: (isDesc ? 1 : 0) as any }];
                dataset.refresh();
            };

            const onRowClick = (evaluationId: string) => {
                const record = dataset.records[evaluationId];
                let targetMainId = "", targetMainEntity = "";
                const regardingVal = record.getValue("msdyn_regardingobjectid");
                
                if (regardingVal != null && typeof regardingVal === "object") {
                    // @ts-ignore
                    targetMainId = regardingVal.id.guid || regardingVal.id;
                    // @ts-ignore
                    targetMainEntity = regardingVal.etn;
                } else {
                    targetMainId = evaluationId; targetMainEntity = gridEntityName;
                }

                // Capture Xrm reference before any navigation to avoid stale context
                let xrmToUse: any = null;
                try { xrmToUse = (window.top as any).Xrm; } catch(e) {}
                if (!xrmToUse) xrmToUse = (window as any).Xrm;

                const navigatePayload = {
                    pageType: "entityrecord" as const,
                    entityName: gridEntityName,
                    entityId: evaluationId,
                    formId: "a063fd0a-bb08-f011-bae3-0022482ad58f"
                };

                // Open main form first — this is the user's primary action
                context.navigation.openForm({ entityName: targetMainEntity, entityId: targetMainId, openInNewWindow: false });

                // Delay side pane to avoid competing navigation conflict (UCI cancels simultaneous navigations).
                // Use window.top.setTimeout so the timer survives the PCF iframe being torn down during form transition.
                const topWindow: any = window.top || window;
                topWindow.setTimeout(() => {
                    try {
                        if (xrmToUse && xrmToUse.App && xrmToUse.App.sidePanes) {
                            const paneId = "evaluationPane";
                            const existingPane = xrmToUse.App.sidePanes.getPane(paneId);
                            if (existingPane) {
                                existingPane.navigate(navigatePayload);
                                existingPane.select();
                            } else {
                                xrmToUse.App.sidePanes.createPane({
                                    title: "Evaluation Pane",
                                    paneId: paneId,
                                    imageSrc: "WebResources/msdyn_formSparkleRegular.svg",
                                    hideHeader: true,
                                    canClose: true,
                                    width: 340,
                                    alwaysRender: true
                                }).then((pane: any) => {
                                    pane.navigate(navigatePayload);
                                    pane.select();
                                }).catch(() => {});
                            }
                        }
                    } catch (e) {}
                }, 1500);
            };

            const onLookupClick = (typeName: string, id: string) => {
                context.navigation.openForm({ entityName: typeName, entityId: id, openInNewWindow: false });
            };

            // Compute score ranges from colorConfig for chip filters
            const scoreRanges: IScoreRange[] = (colorConfig.score || []).map((range: any, idx: number, arr: any[]) => {
                const prevMax = idx === 0 ? -1 : arr[idx - 1].max;
                const labels = ['Failed', 'Medium', 'Excellent'];
                return {
                    label: labels[idx] || `Range ${idx + 1}`,
                    min: prevMax + 1,
                    max: range.max,
                    style: range.style
                };
            });

            const onFilter = (filters: Record<string, number | null>, scoreRange?: { min: number; max: number } | null, dateRange?: { field: string; start: string; end: string } | null) => {
                const conditions: Array<{ attributeName: string; conditionOperator: number; value: string }> = [];
                // Picklist/state equals conditions
                for (const [field, value] of Object.entries(filters)) {
                    if (value !== null) {
                        conditions.push({ attributeName: field, conditionOperator: 0, value: String(value) });
                    }
                }
                // Score range conditions (ge/le)
                if (scoreRange) {
                    conditions.push({ attributeName: 'msdyn_score', conditionOperator: 4, value: String(scoreRange.min) });
                    conditions.push({ attributeName: 'msdyn_score', conditionOperator: 5, value: String(scoreRange.max) });
                }
                // Date range conditions (ge/lt)
                if (dateRange) {
                    conditions.push({ attributeName: dateRange.field, conditionOperator: 4, value: dateRange.start });
                    conditions.push({ attributeName: dateRange.field, conditionOperator: 3, value: dateRange.end });
                }
                if (conditions.length === 0) {
                    dataset.filtering.clearFilter();
                } else {
                    dataset.filtering.setFilter({
                        conditions: conditions,
                        filterOperator: 0
                    } as any);
                }
                // Re-fetch card data with current filter conditions
                this._fetchCardData(conditions);
                dataset.refresh();
            };

            const paging = dataset.paging;
            
            ReactDOM.render(
                React.createElement(QualityGrid, {
                    items: items,
                    columnsConfig: columnsConfig,
                    onRowClick: onRowClick,
                    onLookupClick: onLookupClick,
                    onSort: onSort, 
                    clientUrl: this._clientUrl,
                    hasNextPage: paging.hasNextPage,
                    hasPreviousPage: paging.hasPreviousPage,
                    onLoadNextPage: () => paging.hasNextPage && paging.loadNextPage(),
                    onLoadPreviousPage: () => paging.hasPreviousPage && paging.loadPreviousPage(),
                    totalResultCount: paging.totalResultCount,
                    currentPage: (paging as any).pageNumber || 1,
                    colorConfig: colorConfig,
                    languageId: context.userSettings.languageId,
                    optionColors: this._optionColors,
                    optionMeta: this._optionMeta,
                    onFilter: onFilter,
                    cardData: this._cardData,
                    scoreRanges: scoreRanges,
                    dueDateRanges: this._dueDateRanges
                }),
                this._container
            );

        } catch (error) {
            this._container.innerHTML = `<div style="color:red">Error: ${(error as any).message}</div>`;
        }
    }

    private async _fetchOptionColors(entityName: string): Promise<void> {
        const metadataTypes = [
            'Microsoft.Dynamics.CRM.PicklistAttributeMetadata',
            'Microsoft.Dynamics.CRM.StatusAttributeMetadata',
            'Microsoft.Dynamics.CRM.StateAttributeMetadata'
        ];
        try {
            const fetches = metadataTypes.map(type => {
                const url = `${this._clientUrl}/api/data/v9.2/EntityDefinitions(LogicalName='${entityName}')/Attributes/${type}?$select=LogicalName&$expand=OptionSet`;
                return fetch(url, {
                    credentials: 'include',
                    headers: { 'Accept': 'application/json', 'OData-MaxVersion': '4.0', 'OData-Version': '4.0' }
                }).then(r => r.ok ? r.json() : null).catch(() => null);
            });
            const results = await Promise.all(fetches);
            for (const data of results) {
                if (!data || !data.value) continue;
                for (const attr of data.value) {
                    this._processAttributeOptions(attr);
                }
            }
        } catch (e) {
            console.log("PCF: Error fetching option colors from metadata.");
        }
    }

    private async _fetchEntitySetName(entityName: string): Promise<void> {
        try {
            const url = `${this._clientUrl}/api/data/v9.2/EntityDefinitions(LogicalName='${entityName}')?$select=EntitySetName`;
            const resp = await fetch(url, {
                credentials: 'include',
                headers: { 'Accept': 'application/json', 'OData-MaxVersion': '4.0', 'OData-Version': '4.0' }
            });
            if (resp.ok) {
                const data = await resp.json();
                this._entitySetName = data.EntitySetName || '';
            }
        } catch (e) {
            console.log('PCF: Error fetching entity set name.');
        }
    }

    private async _fetchCardData(filterConditions: Array<{ attributeName: string; conditionOperator: number; value: string }>): Promise<void> {
        if (!this._entitySetName || !this._clientUrl) return;
        // Deduplicate: skip if same filters already in-flight
        const hash = JSON.stringify(filterConditions);
        if (hash === this._lastFilterHash && this._cardData) return;
        this._lastFilterHash = hash;

        // Build FetchXML filter block from current conditions
        let filterXml = '';
        if (filterConditions.length > 0) {
            const opMap: Record<number, string> = { 0: 'eq', 3: 'lt', 4: 'ge', 5: 'le' };
            const condXml = filterConditions.map(c =>
                `<condition attribute="${c.attributeName}" operator="${opMap[c.conditionOperator] || 'eq'}" value="${c.value}" />`
            ).join('');
            filterXml = `<filter type="and">${condXml}</filter>`;
        }

        const entityName = this._context.parameters.dataSet.getTargetEntityType();

        // 1. Average score
        const avgFetch = `<fetch aggregate="true"><entity name="${entityName}"><attribute name="msdyn_score" aggregate="avg" alias="avg_score" />${filterXml}</entity></fetch>`;

        // 2. Critical question status counts
        const critFetch = `<fetch aggregate="true"><entity name="${entityName}"><attribute name="msdyn_criticalquestionstatus" groupby="true" alias="cq_status" /><attribute name="msdyn_criticalquestionstatus" aggregate="count" alias="cnt" />${filterXml}</entity></fetch>`;

        // 3. Evaluator due date: this week, next week, this month
        const now = new Date();
        const dayOfWeek = now.getDay(); // 0=Sun
        const mondayOffset = dayOfWeek === 0 ? -6 : 1 - dayOfWeek;
        const weekStart = new Date(now); weekStart.setDate(now.getDate() + mondayOffset); weekStart.setHours(0, 0, 0, 0);
        const weekEnd = new Date(weekStart); weekEnd.setDate(weekStart.getDate() + 7);
        const nextWeekEnd = new Date(weekStart); nextWeekEnd.setDate(weekStart.getDate() + 14);
        const monthStart = new Date(now.getFullYear(), now.getMonth(), 1);
        const monthEnd = new Date(now.getFullYear(), now.getMonth() + 1, 1);

        const fmt = (d: Date) => d.toISOString().split('T')[0];

        // Store date boundaries for card click filtering
        this._dueDateRanges = {
            thisWeek: { start: fmt(weekStart), end: fmt(weekEnd) },
            nextWeek: { start: fmt(weekEnd), end: fmt(nextWeekEnd) },
            thisMonth: { start: fmt(monthStart), end: fmt(monthEnd) }
        };

        const thisWeekFetch = `<fetch aggregate="true"><entity name="${entityName}"><attribute name="msdyn_evaluatorduedate" aggregate="count" alias="cnt" />${filterXml}<filter type="and"><condition attribute="msdyn_evaluatorduedate" operator="ge" value="${fmt(weekStart)}" /><condition attribute="msdyn_evaluatorduedate" operator="lt" value="${fmt(weekEnd)}" /></filter></entity></fetch>`;
        const nextWeekFetch = `<fetch aggregate="true"><entity name="${entityName}"><attribute name="msdyn_evaluatorduedate" aggregate="count" alias="cnt" />${filterXml}<filter type="and"><condition attribute="msdyn_evaluatorduedate" operator="ge" value="${fmt(weekEnd)}" /><condition attribute="msdyn_evaluatorduedate" operator="lt" value="${fmt(nextWeekEnd)}" /></filter></entity></fetch>`;
        const thisMonthFetch = `<fetch aggregate="true"><entity name="${entityName}"><attribute name="msdyn_evaluatorduedate" aggregate="count" alias="cnt" />${filterXml}<filter type="and"><condition attribute="msdyn_evaluatorduedate" operator="ge" value="${fmt(monthStart)}" /><condition attribute="msdyn_evaluatorduedate" operator="lt" value="${fmt(monthEnd)}" /></filter></entity></fetch>`;

        try {
            const doFetch = (fetchXml: string) =>
                fetch(`${this._clientUrl}/api/data/v9.2/${this._entitySetName}?fetchXml=${encodeURIComponent(fetchXml)}`, {
                    credentials: 'include',
                    headers: { 'Accept': 'application/json', 'OData-MaxVersion': '4.0', 'OData-Version': '4.0' }
                }).then(r => r.ok ? r.json() : null).catch(() => null);

            const [avgResult, critResult, twResult, nwResult, tmResult] = await Promise.all([
                doFetch(avgFetch), doFetch(critFetch), doFetch(thisWeekFetch), doFetch(nextWeekFetch), doFetch(thisMonthFetch)
            ]);

            // Parse average score
            let avgScore: number | null = null;
            if (avgResult?.value?.[0]) {
                avgScore = Math.round(avgResult.value[0].avg_score ?? 0);
            }

            // Parse critical question counts — need to map option values to pass/fail
            let criticalPass = 0, criticalFail = 0;
            if (critResult?.value) {
                const critMeta = this._optionMeta['msdyn_criticalquestionstatus'] || [];
                for (const row of critResult.value) {
                    const statusVal = row.cq_status;
                    const count = row.cnt || 0;
                    // Determine pass/fail by label match
                    const metaItem = critMeta.find(m => m.value === statusVal);
                    const label = (metaItem?.label || '').toLowerCase();
                    if (label.includes('pass')) criticalPass += count;
                    else if (label.includes('fail')) criticalFail += count;
                }
            }

            // Parse due date counts
            const extractCount = (r: any) => (r?.value?.[0]?.cnt) || 0;

            this._cardData = {
                avgScore,
                criticalPass,
                criticalFail,
                dueDateThisWeek: extractCount(twResult),
                dueDateNextWeek: extractCount(nwResult),
                dueDateThisMonth: extractCount(tmResult)
            };
            // Re-render with card data
            this.updateView(this._context);
        } catch (e) {
            console.log('PCF: Error fetching card data.', e);
        }
    }

    private _processAttributeOptions(attr: any): void {
        if (!attr.OptionSet || !attr.OptionSet.Options) return;
        const fieldColors: Record<string, string> = {};
        const fieldOptions: Array<{ value: number; label: string; color: string }> = [];
        for (const opt of attr.OptionSet.Options) {
            if (opt.Color) {
                fieldColors[String(opt.Value)] = opt.Color;
            }
            const label = (opt.Label && opt.Label.UserLocalizedLabel && opt.Label.UserLocalizedLabel.Label) || String(opt.Value);
            fieldOptions.push({ value: opt.Value, label: label, color: opt.Color || '' });
        }
        if (Object.keys(fieldColors).length > 0) {
            this._optionColors[attr.LogicalName] = fieldColors;
        }
        if (fieldOptions.length > 0) {
            this._optionMeta[attr.LogicalName] = fieldOptions;
        }
    }

    public getOutputs(): IOutputs { return {}; }
    public destroy(): void {
        ReactDOM.unmountComponentAtNode(this._container);
        try {
            const paneId = "evaluationPane";
            const topXrm = (window.top as any).Xrm;
            if (topXrm?.App?.sidePanes?.getPane(paneId)) topXrm.App.sidePanes.getPane(paneId).close();
            const localXrm = (window as any).Xrm;
            if (localXrm?.App?.sidePanes?.getPane(paneId)) localXrm.App.sidePanes.getPane(paneId).close();
        } catch (e) {}
    }
}