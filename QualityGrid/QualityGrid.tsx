import * as React from 'react';
import { useState, useEffect } from 'react';
import { DetailsList, DetailsListLayoutMode, SelectionMode, IColumn, ConstrainMode } from '@fluentui/react/lib/DetailsList';
import { ScrollablePane, ScrollbarVisibility } from '@fluentui/react/lib/ScrollablePane';
import { Sticky, StickyPositionType } from '@fluentui/react/lib/Sticky';
import { Persona, PersonaSize, PersonaPresence } from '@fluentui/react/lib/Persona';
import { Icon } from '@fluentui/react/lib/Icon';
import { Link } from '@fluentui/react/lib/Link';
import { IconButton, DefaultButton } from '@fluentui/react/lib/Button';
import { Dropdown, IDropdownOption, IDropdownStyles } from '@fluentui/react/lib/Dropdown';
import { mergeStyles } from '@fluentui/react/lib/Styling';

// זיכרון גלובלי לאייקונים
const iconTypeCache = new Map<string, string>();

export interface IGridItem {
    key: string;
    [key: string]: any;
}

export interface IColumnConfig {
    fieldName: string;
    displayName: string;
    dataType: string;
    minWidth: number;
    isPrimary?: boolean;
    isSorted?: boolean;
    isSortedDescending?: boolean;
}

export interface IScoreRange {
    label: string;
    min: number;
    max: number;
    style: { bg: string; color: string };
}

export interface ICardData {
    avgScore: number | null;
    criticalPass: number;
    criticalFail: number;
    dueDateThisWeek: number;
    dueDateNextWeek: number;
    dueDateThisMonth: number;
}

export interface IDueDateRanges {
    thisWeek: { start: string; end: string };
    nextWeek: { start: string; end: string };
    thisMonth: { start: string; end: string };
}

export interface IQualityGridProps {
    items: IGridItem[];
    columnsConfig: IColumnConfig[];
    onRowClick: (id: string) => void;
    onLookupClick: (entityType: string, id: string) => void;
    onSort: (columnName: string, isDesc: boolean) => void;
    clientUrl: string;
    hasNextPage: boolean;
    hasPreviousPage: boolean;
    onLoadNextPage: () => void;
    onLoadPreviousPage: () => void;
    totalResultCount: number; 
    currentPage: number;
    colorConfig: any;
    languageId: number;
    optionColors: Record<string, Record<string, string>>;
    optionMeta: Record<string, Array<{ value: number; label: string; color: string }>>;
    onFilter: (filters: Record<string, number | null>, scoreRange?: { min: number; max: number } | null, dateRange?: { field: string; start: string; end: string } | null) => void;
    cardData: ICardData | null;
    scoreRanges: IScoreRange[];
    dueDateRanges: IDueDateRanges | null;
}

const localizedStrings: Record<string, { totalRecords: string; page: string; previous: string; next: string; isRTL: boolean }> = {
    "1037": { totalRecords: '\u05e1\u05d4"\u05db \u05e8\u05e9\u05d5\u05de\u05d5\u05ea', page: '\u05e2\u05de\u05d5\u05d3', previous: '\u05d4\u05e7\u05d5\u05d3\u05dd', next: '\u05d4\u05d1\u05d0', isRTL: true },
    "default": { totalRecords: 'Total records', page: 'Page', previous: 'Previous', next: 'Next', isRTL: false }
};

// --- תיקון גודל: הקטנה ל-14px (30% פחות) ---
const iconClass = mergeStyles({ 
    fontSize: 14, // <--- גודל עדין ומדויק
    marginRight: 8, 
    verticalAlign: 'middle', 
    color: '#333333', // אפור כהה/שחור קלאסי
    fontWeight: 'normal'
});

const pillBaseStyle = mergeStyles({
    padding: '4px 12px',
    borderRadius: '16px',
    fontSize: '12px',
    fontWeight: 700,
    display: 'inline-flex',
    alignItems: 'center',
    justifyContent: 'center',
    minWidth: '60px',
    textAlign: 'center'
});

const rootContainerStyle = mergeStyles({
    height: '100%',
    position: 'relative',
    display: 'flex',
    flexDirection: 'column'
});

const gridContainerStyle = mergeStyles({
    flexGrow: 1,
    position: 'relative',
    overflow: 'hidden'
});

const footerStyle = mergeStyles({
    height: '40px',
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'space-between',
    padding: '0 20px',
    borderTop: '1px solid #edebe9',
    backgroundColor: '#f8f8f8',
    flexShrink: 0
});

const filterBarStyle = mergeStyles({
    display: 'flex',
    alignItems: 'flex-end',
    padding: '8px 16px 10px',
    borderBottom: '1px solid #e1dfdd',
    backgroundColor: '#faf9f8',
    flexShrink: 0,
    gap: '16px',
    flexWrap: 'wrap'
});

const filterLabelStyle = mergeStyles({
    display: 'flex',
    alignItems: 'center',
    gap: '6px',
    fontSize: '13px',
    fontWeight: 600,
    color: '#323130',
    paddingBottom: '6px',
    flexShrink: 0
});

const chipContainerStyle = mergeStyles({
    display: 'flex',
    alignItems: 'center',
    gap: '6px',
    flexWrap: 'wrap',
    paddingBottom: '4px'
});

const chipStyle = (isActive: boolean) => mergeStyles({
    padding: '4px 12px',
    borderRadius: '16px',
    fontSize: '12px',
    fontWeight: 600,
    cursor: 'pointer',
    border: isActive ? '1.5px solid #0078D4' : '1px solid #c8c6c4',
    backgroundColor: isActive ? '#E0F1FF' : '#ffffff',
    color: isActive ? '#0078D4' : '#605E5C',
    display: 'inline-flex',
    alignItems: 'center',
    gap: '4px',
    userSelect: 'none',
    transition: 'all 0.15s ease',
    ':hover': {
        backgroundColor: isActive ? '#CCE4F7' : '#f3f2f1',
        borderColor: isActive ? '#0078D4' : '#a19f9d'
    }
});

const styledDropdownStyles: Partial<IDropdownStyles> = {
    root: { minWidth: 160, maxWidth: 220 },
    dropdown: {
        fontSize: 13,
        borderRadius: 6,
        border: '1px solid #c8c6c4',
        selectors: {
            ':hover': { borderColor: '#a19f9d' },
            ':focus::after': { borderRadius: 6 }
        }
    },
    title: { borderRadius: 6, fontSize: 13, height: 32, lineHeight: '30px' },
    caretDownWrapper: { top: 4 },
    label: { fontSize: 12, fontWeight: 600, color: '#605E5C', paddingBottom: 2 }
};

const UserWithImage: React.FunctionComponent<{ name: string, userId: string, clientUrl: string }> = (props) => {
    try {
        let cleanId = props.userId;
        if (cleanId && cleanId.startsWith('{')) cleanId = cleanId.replace('{', '').replace('}', '');
        
        // --- התיקון הגדול ---
        // במקום לפנות ל-API שמחזיר Base64 (טקסט), אנו פונים ל-download.aspx שמחזיר תמונה אמיתית
        const imageUrl = cleanId && props.clientUrl 
            ? `${props.clientUrl}/Image/download.aspx?Entity=systemuser&Attribute=entityimage&Id=${cleanId}&Timestamp=${new Date().getTime()}` 
            : undefined;

        return (
            <div style={{ display: 'flex', alignItems: 'center' }}>
                <Persona 
                    text={props.name} 
                    size={PersonaSize.size24} 
                    imageUrl={imageUrl} // עכשיו זה יקבל קובץ תמונה אמיתי
                    imageInitials={props.name ? props.name.substring(0,2).toUpperCase() : "?"} 
                    presence={PersonaPresence.none} 
                    hidePersonaDetails={false} 
                />
            </div>
        );
    } catch (e) {
        return <span>{props.name}</span>;
    }
};

// --------------------------------------------------------------------------
// רכיב האייקונים החכם - מעודכן עם האייקונים שביקשת
// --------------------------------------------------------------------------
const SmartEntityIcon: React.FunctionComponent<{ entityType: string, recordId: string, text: string }> = (props) => {
    // ברירת מחדל
    const [iconName, setIconName] = useState<string>('Page');

    useEffect(() => {
        const checkIcon = async () => {
            const type = (props.entityType || "").toLowerCase();

            // 1. טיפול ב-Case (אירוע) -> TextDocument
            if (type === 'incident' || type === 'case') {
                setIconName('TextDocument'); // <--- לבקשתך
                return;
            }

            // 2. טיפול במייל וטלפון רגיל
            if (type === 'email') { setIconName('Mail'); return; }
            if (type === 'phonecall') { setIconName('Phone'); return; } // <--- Phone

            // 3. הטיפול המיוחד בשיחה (Conversation - Voice vs Chat)
            if (type === 'msdyn_ocliveworkitem') {
                
                if (iconTypeCache.has(props.recordId)) {
                    setIconName(iconTypeCache.get(props.recordId) || 'OfficeChat');
                    return;
                }

                try {
                    // @ts-ignore
                    const result = await Xrm.WebApi.retrieveRecord("msdyn_ocliveworkitem", props.recordId, "?$select=msdyn_channel");
                    
                    let detectedIcon = 'OfficeChat'; // <--- OfficeChat לבקשתך
                    
                    // 192440000 = Voice Call
                    if (result && result.msdyn_channel === 192440000) {
                        detectedIcon = 'Phone'; // <--- Phone לבקשתך
                    }

                    iconTypeCache.set(props.recordId, detectedIcon);
                    setIconName(detectedIcon);

                } catch (e) {
                    console.error("Failed to fetch channel type", e);
                    setIconName('OfficeChat'); 
                }
            } else {
                setIconName('Page');
            }
        };

        checkIcon();
    }, [props.entityType, props.recordId]);

    return <span><Icon iconName={iconName} className={iconClass} />{props.text}</span>;
};


const getContrastColor = (hexColor: string): string => {
    try {
        const hex = hexColor.replace('#', '');
        const r = parseInt(hex.substring(0, 2), 16);
        const g = parseInt(hex.substring(2, 4), 16);
        const b = parseInt(hex.substring(4, 6), 16);
        const luminance = (0.299 * r + 0.587 * g + 0.114 * b) / 255;
        return luminance > 0.5 ? '#333333' : '#FFFFFF';
    } catch {
        return '#333333';
    }
};

export const QualityGrid: React.FunctionComponent<IQualityGridProps> = (props) => {

    const [activeFilters, setActiveFilters] = useState<Record<string, number | null>>({});
    const [activeRecordType, setActiveRecordType] = useState<number | string | null>(null);
    const [activeScoreRange, setActiveScoreRange] = useState<IScoreRange | null>(null);
    const [activeDueDateKey, setActiveDueDateKey] = useState<string | null>(null);
    const [activeCriticalFilter, setActiveCriticalFilter] = useState<'pass' | 'fail' | null>(null);

    // Identify all filterable columns (Picklist/State that have option metadata)
    // Exclude the record type column — it gets its own chip filter
    const filterableColumns = React.useMemo(() => {
        return props.columnsConfig.filter(c =>
            !c.fieldName.includes('recordtype')
            && (c.dataType === 'Picklist' || c.dataType === 'State' || c.fieldName.includes('status'))
            && props.optionMeta
            && props.optionMeta[c.fieldName]
            && props.optionMeta[c.fieldName].length > 0
        );
    }, [props.columnsConfig, props.optionMeta]);

    // Identify record type column
    const recordTypeCol = React.useMemo(() => {
        return props.columnsConfig.find(c => c.fieldName.includes('recordtype'));
    }, [props.columnsConfig]);

    // Get record type values: prefer metadata (all values) over current items (partial)
    const recordTypeHasMeta = !!(recordTypeCol && props.optionMeta && props.optionMeta[recordTypeCol.fieldName] && props.optionMeta[recordTypeCol.fieldName].length > 0);

    const recordTypeChips = React.useMemo(() => {
        if (!recordTypeCol) return [];
        // If metadata is available, use it — this gives all possible values
        if (recordTypeHasMeta) {
            return props.optionMeta[recordTypeCol.fieldName].map(o => ({ label: o.label, value: o.value }));
        }
        // Fallback: extract unique labels from current items (client-side only)
        const seen = new Set<string>();
        for (const item of props.items) {
            const val = item[recordTypeCol.fieldName];
            if (val && typeof val === 'string' && val.trim()) seen.add(val.trim());
        }
        return Array.from(seen).sort().map(label => ({ label, value: label }));
    }, [props.items, recordTypeCol, recordTypeHasMeta, props.optionMeta]);

    // Filter items: server-side when metadata exists, client-side fallback
    const displayItems = React.useMemo(() => {
        if (!activeRecordType || !recordTypeCol) return props.items;
        // When using server-side filtering (metadata path), dataset is already filtered
        if (recordTypeHasMeta) return props.items;
        // Client-side fallback: filter by label
        return props.items.filter(item => item[recordTypeCol.fieldName] === String(activeRecordType));
    }, [props.items, activeRecordType, recordTypeCol, recordTypeHasMeta]);

    const hasAnyFilter = Object.values(activeFilters).some(v => v !== null) || activeRecordType !== null || activeScoreRange !== null || activeDueDateKey !== null || activeCriticalFilter !== null;

    const applyFilters = (dropdownFilters: Record<string, number | null>, recordType: number | string | null, scoreRange?: IScoreRange | null, dueDateKey?: string | null, critFilter?: 'pass' | 'fail' | null) => {
        const combined = { ...dropdownFilters };
        // Include record type in server-side filter when metadata is available
        if (recordTypeCol && recordTypeHasMeta) {
            combined[recordTypeCol.fieldName] = typeof recordType === 'number' ? recordType : null;
        }
        // Critical question status filter
        const cf = critFilter !== undefined ? critFilter : activeCriticalFilter;
        if (cf && props.optionMeta && props.optionMeta['msdyn_criticalquestionstatus']) {
            const match = props.optionMeta['msdyn_criticalquestionstatus'].find(m => m.label.toLowerCase().includes(cf));
            if (match) combined['msdyn_criticalquestionstatus'] = match.value;
        }
        const sr = scoreRange !== undefined ? scoreRange : activeScoreRange;
        // Date range
        const ddk = dueDateKey !== undefined ? dueDateKey : activeDueDateKey;
        let dateRange: { field: string; start: string; end: string } | null = null;
        if (ddk && props.dueDateRanges) {
            const r = (props.dueDateRanges as any)[ddk];
            if (r) dateRange = { field: 'msdyn_evaluatorduedate', start: r.start, end: r.end };
        }
        props.onFilter(combined, sr ? { min: sr.min, max: sr.max } : null, dateRange);
    };

    const clearAllFilters = () => {
        const emptyFilters: Record<string, number | null> = {};
        setActiveFilters(emptyFilters);
        setActiveRecordType(null);
        setActiveScoreRange(null);
        setActiveDueDateKey(null);
        setActiveCriticalFilter(null);
        props.onFilter(emptyFilters, null, null);
    };

    const getStatusStyle = (rawValue: any, fieldName?: string) => {
        const key = String(rawValue);
        // For Picklist fields, prefer D365-configured option colors
        if (fieldName && props.optionColors && props.optionColors[fieldName] && props.optionColors[fieldName][key]) {
            const bgColor = props.optionColors[fieldName][key];
            return { bg: bgColor, color: getContrastColor(bgColor) };
        }
        if (props.colorConfig && props.colorConfig.status && props.colorConfig.status[key]) {
            return props.colorConfig.status[key];
        }
        return { bg: '#F3F2F1', color: '#605E5C' };
    };

    const getScoreStyle = (score: number) => {
        if (props.colorConfig && props.colorConfig.score && Array.isArray(props.colorConfig.score)) {
            for (const range of props.colorConfig.score) {
                if (score <= range.max) {
                    return range.style;
                }
            }
        }
        return { bg: '#F3F2F1', color: '#605E5C' };
    };

    const onRenderDetailsHeader = (headerProps: any, defaultRender: any) => {
        if (!headerProps) return null;
        return (
            <Sticky stickyPosition={StickyPositionType.Header} isScrollSynced>
                {defaultRender(headerProps)}
            </Sticky>
        );
    };

    const columns: IColumn[] = props.columnsConfig.map(col => {
        return {
            key: col.fieldName,
            name: col.displayName,
            fieldName: col.fieldName,
            minWidth: col.minWidth,
            maxWidth: col.fieldName === 'msdyn_regardingobjectid' ? 250 : 150,
            isResizable: true,
            isSorted: col.isSorted,
            isSortedDescending: col.isSortedDescending,
            onColumnClick: (ev: React.MouseEvent<HTMLElement>, column: IColumn) => {
                const newDesc = column.isSorted ? !column.isSortedDescending : false;
                props.onSort(col.fieldName, newDesc);
            },
            
            onRender: (item: IGridItem) => {
                try {
                    const rawValue = item[col.fieldName + '_raw'];
                    const textValue = item[col.fieldName];

                    // נתונים לצורך זיהוי איקון חכם (גם אם אנחנו בעמודה אחרת)
                    let relatedId = "";
                    let relatedType = "";
                    
                    // מנסים לחלץ את ה-Regarding משדות אפשריים בשורה
                    const regardingData = item['msdyn_regardingobjectid_raw'] || item['regardingobjectid_raw'];
                    if (regardingData) {
                         if (Array.isArray(regardingData) && regardingData[0]) { 
                             relatedId = regardingData[0].id.guid; 
                             relatedType = regardingData[0].etn; 
                         } else if (regardingData.id) { 
                             relatedId = regardingData.id.guid; 
                             relatedType = regardingData.etn; 
                         }
                    }

                    if (col.isPrimary) {
                        return <Link onClick={() => props.onRowClick(item.key)} style={{ fontWeight: '700', fontSize: '14px' }}>{textValue}</Link>;
                    }

                    const isLookup = col.dataType.startsWith('Lookup') || col.dataType === 'Customer' || col.dataType === 'Owner';
                    if (isLookup) {
                         let targetId = "", targetType = "";
                         if (Array.isArray(rawValue) && rawValue[0]) { targetId = rawValue[0].id.guid; targetType = rawValue[0].etn; }
                         else if (rawValue && rawValue.id) { targetId = rawValue.id.guid; targetType = rawValue.etn; }

                         if (targetType === 'systemuser') return <UserWithImage name={textValue} userId={targetId} clientUrl={props.clientUrl} />;
                         
                         // כאן בעמודת regarding מחזירים רק לינק רגיל (בלי איקון) כי העברנו אותו למקום אחר
                         return <Link onClick={() => props.onLookupClick(targetType, targetId)}>{textValue}</Link>;
                    }

                    if (col.fieldName.includes('score')) {
                        const score = Number(rawValue) || 0;
                        const style = getScoreStyle(score);
                        return <span className={pillBaseStyle} style={{ backgroundColor: style.bg, color: style.color }}>{textValue}/100</span>;
                    }

                    if (col.fieldName.includes('status') || col.dataType === 'State' || col.dataType === 'Picklist') {
                        const style = getStatusStyle(rawValue, col.fieldName);
                        return <span className={pillBaseStyle} style={{ backgroundColor: style.bg, color: style.color }}>{textValue}</span>;
                    }

                    // --- כאן השינוי הגדול: מיקום האייקון ---
                    // בדיקה: האם העמודה הנוכחית היא "סוג רשומה"? (Record Type)
                    if (col.fieldName.includes('recordtype')) {
                        // אם הצלחנו לזהות את ה-Regarding, נשתמש באייקון החכם
                        if (relatedId && relatedType) {
                            return <SmartEntityIcon entityType={relatedType} recordId={relatedId} text={textValue} />;
                        }
                        // גיבוי למקרה שאין Regarding: אייקון רגיל
                        return <span><Icon iconName="Page" className={iconClass} />{textValue}</span>;
                    }

                    return <span style={{fontSize: '13px'}}>{textValue}</span>;

                } catch (error) { return <span>{item[col.fieldName]}</span>; }
            }
        };
    });

    return (
        <div className={rootContainerStyle}>
            {(filterableColumns.length > 0 || recordTypeChips.length > 0 || props.scoreRanges.length > 0 || props.cardData) && (
                <div className={filterBarStyle}>
                    <span className={filterLabelStyle}>
                        <Icon iconName="Filter" style={{ fontSize: 14 }} />
                        Filters
                    </span>
                    {filterableColumns.map(col => {
                        const options: IDropdownOption[] = [
                            { key: '__all__', text: 'All' },
                            ...props.optionMeta[col.fieldName].map(o => ({ key: o.value, text: o.label }))
                        ];
                        return (
                            <Dropdown
                                key={col.fieldName}
                                label={col.displayName}
                                placeholder="All"
                                selectedKey={activeFilters[col.fieldName] === undefined || activeFilters[col.fieldName] === null ? '__all__' : activeFilters[col.fieldName]}
                                options={options}
                                onChange={(_ev, option) => {
                                    const newFilters = { ...activeFilters };
                                    if (!option || option.key === '__all__') {
                                        newFilters[col.fieldName] = null;
                                    } else {
                                        newFilters[col.fieldName] = option.key as number;
                                    }
                                    setActiveFilters(newFilters);
                                    applyFilters(newFilters, activeRecordType);
                                }}
                                styles={styledDropdownStyles}
                            />
                        );
                    })}
                    {recordTypeChips.length > 0 && recordTypeCol && (
                        <div style={{ display: 'flex', flexDirection: 'column' }}>
                            <span style={{ fontSize: 12, fontWeight: 600, color: '#605E5C', paddingBottom: 2 }}>{recordTypeCol.displayName}</span>
                            <div className={chipContainerStyle}>
                                {recordTypeChips.map(chip => (
                                    <span
                                        key={String(chip.value)}
                                        className={chipStyle(activeRecordType === chip.value)}
                                        onClick={() => {
                                            const newVal = activeRecordType === chip.value ? null : chip.value;
                                            setActiveRecordType(newVal);
                                            applyFilters(activeFilters, newVal);
                                        }}
                                    >
                                        {chip.label}
                                        {activeRecordType === chip.value && <Icon iconName="Cancel" style={{ fontSize: 10, cursor: 'pointer' }} />}
                                    </span>
                                ))}
                            </div>
                        </div>
                    )}
                    {props.scoreRanges.length > 0 && (
                        <div style={{ display: 'flex', flexDirection: 'column' }}>
                            <span style={{ fontSize: 12, fontWeight: 600, color: '#605E5C', paddingBottom: 2 }}>Score</span>
                            <div className={chipContainerStyle}>
                                {props.scoreRanges.map(range => {
                                    const isActive = activeScoreRange?.label === range.label;
                                    return (
                                        <span
                                            key={range.label}
                                            style={{
                                                padding: '4px 12px', borderRadius: 16, fontSize: 12, fontWeight: 600,
                                                cursor: 'pointer', display: 'inline-flex', alignItems: 'center', gap: 4,
                                                userSelect: 'none', transition: 'all 0.15s ease',
                                                border: isActive ? `1.5px solid ${range.style.color}` : '1px solid #c8c6c4',
                                                backgroundColor: isActive ? range.style.bg : '#ffffff',
                                                color: isActive ? range.style.color : '#605E5C'
                                            }}
                                            onClick={() => {
                                                const newRange = isActive ? null : range;
                                                setActiveScoreRange(newRange);
                                                applyFilters(activeFilters, activeRecordType, newRange);
                                            }}
                                        >
                                            {range.label} ({range.min}–{range.max})
                                            {isActive && <Icon iconName="Cancel" style={{ fontSize: 10, cursor: 'pointer' }} />}
                                        </span>
                                    );
                                })}
                            </div>
                        </div>
                    )}
                    {hasAnyFilter && (
                        <DefaultButton
                            text="Clear all"
                            iconProps={{ iconName: 'ClearFilter' }}
                            onClick={clearAllFilters}
                            styles={{ root: { height: 32, borderRadius: 6, fontSize: 12 }, icon: { fontSize: 12 } }}
                        />
                    )}
                    {props.cardData && (
                        <div style={{ display: 'flex', gap: 10, marginLeft: 'auto', alignItems: 'stretch', flexShrink: 0 }}>
                            {/* Card 1: Average Score */}
                            <div style={{ border: '1.5px solid #D13438', borderRadius: 8, padding: '6px 16px', backgroundColor: '#fff', textAlign: 'center', display: 'flex', flexDirection: 'column', minWidth: 110 }}>
                                <div style={{ fontSize: 11, fontWeight: 600, color: '#D13438', lineHeight: 1.2 }}>Average score</div>
                                <div style={{ fontSize: 22, fontWeight: 700, color: '#323130', lineHeight: 1.2, flex: 1, display: 'flex', alignItems: 'center', justifyContent: 'center' }}>{props.cardData.avgScore ?? '—'}</div>
                            </div>
                            {/* Card 2: Critical Questions (clickable) */}
                            <div style={{ border: '1.5px solid #D13438', borderRadius: 8, padding: '6px 16px', backgroundColor: '#fff', textAlign: 'center', display: 'flex', flexDirection: 'column', minWidth: 150 }}>
                                <div style={{ fontSize: 11, fontWeight: 600, color: '#D13438', lineHeight: 1.2 }}>Critical questions</div>
                                <div style={{ display: 'flex', justifyContent: 'center', alignItems: 'center', gap: 12, fontSize: 13, fontWeight: 600, flex: 1 }}>
                                    <span
                                        style={{ display: 'inline-flex', alignItems: 'center', gap: 4, color: '#107C10', cursor: 'pointer', padding: '2px 6px', borderRadius: 4, backgroundColor: activeCriticalFilter === 'pass' ? '#DFF6DD' : 'transparent', transition: 'background-color 0.15s' }}
                                        onClick={() => { const nv = activeCriticalFilter === 'pass' ? null : 'pass' as const; setActiveCriticalFilter(nv); applyFilters(activeFilters, activeRecordType, undefined, undefined, nv); }}
                                    >
                                        <span style={{ width: 7, height: 7, borderRadius: '50%', backgroundColor: '#107C10', display: 'inline-block', flexShrink: 0 }} />
                                        {props.cardData.criticalPass} Pass
                                    </span>
                                    <span style={{ color: '#c8c6c4' }}>·</span>
                                    <span
                                        style={{ display: 'inline-flex', alignItems: 'center', gap: 4, color: '#A80000', cursor: 'pointer', padding: '2px 6px', borderRadius: 4, backgroundColor: activeCriticalFilter === 'fail' ? '#FDE7E9' : 'transparent', transition: 'background-color 0.15s' }}
                                        onClick={() => { const nv = activeCriticalFilter === 'fail' ? null : 'fail' as const; setActiveCriticalFilter(nv); applyFilters(activeFilters, activeRecordType, undefined, undefined, nv); }}
                                    >
                                        <span style={{ width: 7, height: 7, borderRadius: '50%', backgroundColor: '#A80000', display: 'inline-block', flexShrink: 0 }} />
                                        {props.cardData.criticalFail} Fail
                                    </span>
                                </div>
                            </div>
                            {/* Card 3: Evaluator Expiration Date (clickable segments) */}
                            <div style={{ border: '1.5px solid #D13438', borderRadius: 8, padding: '6px 16px', backgroundColor: '#fff', display: 'flex', flexDirection: 'column', minWidth: 200 }}>
                                <div style={{ fontSize: 11, fontWeight: 600, color: '#D13438', lineHeight: 1.2, textAlign: 'center' }}>Evaluator expiration date</div>
                                <div style={{ display: 'flex', justifyContent: 'center', alignItems: 'center', gap: 0, flex: 1 }}>
                                    {[
                                        { key: 'thisWeek', label: 'This week', count: props.cardData.dueDateThisWeek },
                                        { key: 'nextWeek', label: 'Next week', count: props.cardData.dueDateNextWeek },
                                        { key: 'thisMonth', label: 'This month', count: props.cardData.dueDateThisMonth }
                                    ].map((seg, idx) => (
                                        <React.Fragment key={seg.key}>
                                            {idx > 0 && <div style={{ width: 1, height: 24, backgroundColor: '#e1dfdd', flexShrink: 0 }} />}
                                            <div
                                                style={{ textAlign: 'center', padding: '2px 10px', cursor: 'pointer', borderRadius: 4, backgroundColor: activeDueDateKey === seg.key ? '#E0F1FF' : 'transparent', transition: 'background-color 0.15s' }}
                                                onClick={() => { const nv = activeDueDateKey === seg.key ? null : seg.key; setActiveDueDateKey(nv); applyFilters(activeFilters, activeRecordType, undefined, nv); }}
                                            >
                                                <div style={{ fontSize: 16, fontWeight: 700, color: activeDueDateKey === seg.key ? '#0078D4' : '#323130', lineHeight: 1.1 }}>{seg.count}</div>
                                                <div style={{ fontSize: 10, color: activeDueDateKey === seg.key ? '#0078D4' : '#605E5C', fontWeight: 500 }}>{seg.label}</div>
                                            </div>
                                        </React.Fragment>
                                    ))}
                                </div>
                            </div>
                        </div>
                    )}
                </div>
            )}
            <div className={gridContainerStyle}>
                <ScrollablePane scrollbarVisibility={ScrollbarVisibility.auto}>
                    <DetailsList
                        items={displayItems}
                        columns={columns}
                        setKey="set"
                        layoutMode={DetailsListLayoutMode.fixedColumns}
                        selectionMode={SelectionMode.none}
                        onItemInvoked={(item) => props.onRowClick(item.key)}
                        onRenderDetailsHeader={onRenderDetailsHeader}
                        constrainMode={ConstrainMode.unconstrained}
                    />
                </ScrollablePane>
            </div>
            
            <div className={footerStyle}>
                {(() => {
                    const strings = localizedStrings[String(props.languageId)] || localizedStrings["default"];
                    const prevIcon = strings.isRTL ? 'ChevronRight' : 'ChevronLeft';
                    const nextIcon = strings.isRTL ? 'ChevronLeft' : 'ChevronRight';
                    return (
                        <>
                            <div style={{ fontSize: '12px', color: '#333' }}>
                                {strings.totalRecords}: <b>{props.totalResultCount !== -1 ? props.totalResultCount : props.items.length + '+'}</b>
                                &nbsp; | {strings.page}: {props.currentPage}
                            </div>
                            <div>
                                <IconButton iconProps={{ iconName: prevIcon }} title={strings.previous} disabled={!props.hasPreviousPage} onClick={props.onLoadPreviousPage} />
                                <IconButton iconProps={{ iconName: nextIcon }} title={strings.next} disabled={!props.hasNextPage} onClick={props.onLoadNextPage} />
                            </div>
                        </>
                    );
                })()}
            </div>
        </div>
    );
};