"use strict";

import powerbi from "powerbi-visuals-api";
import VisualConstructorOptions = powerbi.extensibility.visual.VisualConstructorOptions;
import VisualUpdateOptions = powerbi.extensibility.visual.VisualUpdateOptions;
import IVisual = powerbi.extensibility.visual.IVisual;
import IVisualHost = powerbi.extensibility.visual.IVisualHost;
import ITooltipService = powerbi.extensibility.ITooltipService;
import ISelectionManager = powerbi.extensibility.ISelectionManager;
import VisualTooltipDataItem = powerbi.extensibility.VisualTooltipDataItem;
import EnumerateVisualObjectInstancesOptions = powerbi.EnumerateVisualObjectInstancesOptions;
import VisualObjectInstance = powerbi.VisualObjectInstance;
import VisualObjectInstanceEnumerationObject = powerbi.VisualObjectInstanceEnumerationObject;
import * as d3 from "d3";

interface ShiftItem {
    id: string;
    location: string;
    provider: string;
    category: string;
    startDate: Date;
    endDate: Date;
    caseCount: number;
    extraMetric: number;
    subRowIndex: number;
    selectionId: powerbi.visuals.ISelectionId;
}

interface ProviderRow {
    providerName: string;
    shifts: ShiftItem[];
    maxSubRows: number;
    totalCases: number;
    earliestShift: Date;
}

interface LocationGroup {
    locationName: string;
    providers: ProviderRow[];
    totalCases: number;
}

interface VisualSettings {
    colorPaletteSettings: { palette: string };
    legendSettings: { show: boolean; position: string; fontSize: number; fontColor: string };
    swimlaneSettings: {
        rowHeight: number;
        barHeightPercent: number;
        providerColumnWidth: number;
        providerFontSize: number;
        providerFontColor: string;
        showZebraStriping: boolean;
        evenRowColor: string;
        oddRowColor: string;
        showRowBorders: boolean;
        rowBorderColor: string;
    };
    locationHeaderSettings: {
        showHeader: boolean;
        headerHeight: number;
        headerFontSize: number;
        headerFontColor: string;
        headerBgColor: string;
        showCaseBadge: boolean;
        badgeBgColor: string;
        badgeTextColor: string;
    };
    barSettings: {
        barOpacity: number;
        barCornerRadius: number;
        showBorder: boolean;
        borderColor: string;
        borderWidth: number;
        showCaseNotches: boolean;
        overlapMode: string;
    };
    caseLabelSettings: {
        show: boolean;
        labelPosition: string;
        labelContent: string;
        fontSize: number;
        fontColor: string;
        badgeBgColor: string;
        prefix: string;
        suffix: string;
    };
    timeAxisSettings: {
        timeFormat: string;
        clampToSurgicalDay: boolean;
        dayStartHour: number;
        dayEndHour: number;
        fontSize: number;
        fontColor: string;
        showGridlines: boolean;
        gridlineColor: string;
    };
    referenceLineSettings: {
        showNowLine: boolean;
        nowLineColor: string;
        showShiftChangeLines: boolean;
        shiftChangeColor: string;
    };
    overtimeSettings: {
        highlightLongShifts: boolean;
        thresholdHours: number;
        warningBorderColor: string;
    };
    sortSettings: {
        locationSort: string;
        providerSort: string;
    };
}

function cleanAffix(sym: string): string {
    if (!sym || sym === "none") return "";
    return sym;
}

export class Visual implements IVisual {
    private target: HTMLElement;
    private host: IVisualHost;
    private tooltipService: ITooltipService;
    private selectionManager: ISelectionManager;
    private svg: d3.Selection<SVGSVGElement, unknown, null, undefined>;
    private legendContainer: d3.Selection<SVGGElement, unknown, null, undefined>;
    private chartContainer: d3.Selection<SVGGElement, unknown, null, undefined>;
    private settings: VisualSettings;

    constructor(options: VisualConstructorOptions) {
        this.target = options.element;
        this.host = options.host;
        this.tooltipService = options.host.tooltipService;
        this.selectionManager = options.host.createSelectionManager();

        this.target.style.overflow = "auto";
        this.svg = d3.select(this.target).append("svg");
        this.legendContainer = this.svg.append("g").attr("class", "legend-container");
        this.chartContainer = this.svg.append("g").attr("class", "chart-container");

        this.settings = this.getDefaultSettings();
    }

    private getDefaultSettings(): VisualSettings {
        return {
            colorPaletteSettings: { palette: "powerBi" },
            legendSettings: { show: true, position: "Top", fontSize: 10, fontColor: "#444444" },
            swimlaneSettings: {
                rowHeight: 38,
                barHeightPercent: 72,
                providerColumnWidth: 155,
                providerFontSize: 10,
                providerFontColor: "#333333",
                showZebraStriping: true,
                evenRowColor: "#FFFFFF",
                oddRowColor: "#F8FAFC",
                showRowBorders: true,
                rowBorderColor: "#E2E8F0"
            },
            locationHeaderSettings: {
                showHeader: true,
                headerHeight: 28,
                headerFontSize: 11,
                headerFontColor: "#0F172A",
                headerBgColor: "#E2E8F0",
                showCaseBadge: true,
                badgeBgColor: "#3B82F6",
                badgeTextColor: "#FFFFFF"
            },
            barSettings: {
                barOpacity: 92,
                barCornerRadius: 4,
                showBorder: true,
                borderColor: "#FFFFFF",
                borderWidth: 1.5,
                showCaseNotches: true,
                overlapMode: "stagger"
            },
            caseLabelSettings: {
                show: true,
                labelPosition: "pillBadge",
                labelContent: "casesOnly",
                fontSize: 9,
                fontColor: "#FFFFFF",
                badgeBgColor: "rgba(15, 23, 42, 0.75)",
                prefix: "none",
                suffix: " cases"
            },
            timeAxisSettings: {
                timeFormat: "auto",
                clampToSurgicalDay: false,
                dayStartHour: 7,
                dayEndHour: 19,
                fontSize: 10,
                fontColor: "#64748B",
                showGridlines: true,
                gridlineColor: "#E2E8F0"
            },
            referenceLineSettings: {
                showNowLine: true,
                nowLineColor: "#EF4444",
                showShiftChangeLines: true,
                shiftChangeColor: "#94A3B8"
            },
            overtimeSettings: {
                highlightLongShifts: true,
                thresholdHours: 10,
                warningBorderColor: "#DC2626"
            },
            sortSettings: {
                locationSort: "natural",
                providerSort: "natural"
            }
        };
    }

    private parseSettings(dataView: powerbi.DataView): VisualSettings {
        const defaults = this.getDefaultSettings();
        const objects = dataView?.metadata?.objects;
        if (!objects) return defaults;

        const cp = objects["colorPaletteSettings"];
        const leg = objects["legendSettings"];
        const sl = objects["swimlaneSettings"];
        const lh = objects["locationHeaderSettings"];
        const bar = objects["barSettings"];
        const cl = objects["caseLabelSettings"];
        const ax = objects["timeAxisSettings"];
        const ref = objects["referenceLineSettings"];
        const ot = objects["overtimeSettings"];
        const srt = objects["sortSettings"];

        return {
            colorPaletteSettings: { palette: String(cp?.["palette"] ?? defaults.colorPaletteSettings.palette) },
            legendSettings: {
                show: leg?.["show"] !== undefined ? Boolean(leg["show"]) : defaults.legendSettings.show,
                position: String(leg?.["position"] ?? defaults.legendSettings.position),
                fontSize: leg?.["fontSize"] !== undefined ? Number(leg["fontSize"]) : defaults.legendSettings.fontSize,
                fontColor: (leg?.["fontColor"] as { solid?: { color?: string } })?.solid?.color ?? defaults.legendSettings.fontColor
            },
            swimlaneSettings: {
                rowHeight: sl?.["rowHeight"] !== undefined ? Number(sl["rowHeight"]) : defaults.swimlaneSettings.rowHeight,
                barHeightPercent: sl?.["barHeightPercent"] !== undefined ? Number(sl["barHeightPercent"]) : defaults.swimlaneSettings.barHeightPercent,
                providerColumnWidth: sl?.["providerColumnWidth"] !== undefined ? Number(sl["providerColumnWidth"]) : defaults.swimlaneSettings.providerColumnWidth,
                providerFontSize: sl?.["providerFontSize"] !== undefined ? Number(sl["providerFontSize"]) : defaults.swimlaneSettings.providerFontSize,
                providerFontColor: (sl?.["providerFontColor"] as { solid?: { color?: string } })?.solid?.color ?? defaults.swimlaneSettings.providerFontColor,
                showZebraStriping: sl?.["showZebraStriping"] !== undefined ? Boolean(sl["showZebraStriping"]) : defaults.swimlaneSettings.showZebraStriping,
                evenRowColor: (sl?.["evenRowColor"] as { solid?: { color?: string } })?.solid?.color ?? defaults.swimlaneSettings.evenRowColor,
                oddRowColor: (sl?.["oddRowColor"] as { solid?: { color?: string } })?.solid?.color ?? defaults.swimlaneSettings.oddRowColor,
                showRowBorders: sl?.["showRowBorders"] !== undefined ? Boolean(sl["showRowBorders"]) : defaults.swimlaneSettings.showRowBorders,
                rowBorderColor: (sl?.["rowBorderColor"] as { solid?: { color?: string } })?.solid?.color ?? defaults.swimlaneSettings.rowBorderColor
            },
            locationHeaderSettings: {
                showHeader: lh?.["showHeader"] !== undefined ? Boolean(lh["showHeader"]) : defaults.locationHeaderSettings.showHeader,
                headerHeight: lh?.["headerHeight"] !== undefined ? Number(lh["headerHeight"]) : defaults.locationHeaderSettings.headerHeight,
                headerFontSize: lh?.["headerFontSize"] !== undefined ? Number(lh["headerFontSize"]) : defaults.locationHeaderSettings.headerFontSize,
                headerFontColor: (lh?.["headerFontColor"] as { solid?: { color?: string } })?.solid?.color ?? defaults.locationHeaderSettings.headerFontColor,
                headerBgColor: (lh?.["headerBgColor"] as { solid?: { color?: string } })?.solid?.color ?? defaults.locationHeaderSettings.headerBgColor,
                showCaseBadge: lh?.["showCaseBadge"] !== undefined ? Boolean(lh["showCaseBadge"]) : defaults.locationHeaderSettings.showCaseBadge,
                badgeBgColor: (lh?.["badgeBgColor"] as { solid?: { color?: string } })?.solid?.color ?? defaults.locationHeaderSettings.badgeBgColor,
                badgeTextColor: (lh?.["badgeTextColor"] as { solid?: { color?: string } })?.solid?.color ?? defaults.locationHeaderSettings.badgeTextColor
            },
            barSettings: {
                barOpacity: bar?.["barOpacity"] !== undefined ? Number(bar["barOpacity"]) : defaults.barSettings.barOpacity,
                barCornerRadius: bar?.["barCornerRadius"] !== undefined ? Number(bar["barCornerRadius"]) : defaults.barSettings.barCornerRadius,
                showBorder: bar?.["showBorder"] !== undefined ? Boolean(bar["showBorder"]) : defaults.barSettings.showBorder,
                borderColor: (bar?.["borderColor"] as { solid?: { color?: string } })?.solid?.color ?? defaults.barSettings.borderColor,
                borderWidth: bar?.["borderWidth"] !== undefined ? Number(bar["borderWidth"]) : defaults.barSettings.borderWidth,
                showCaseNotches: bar?.["showCaseNotches"] !== undefined ? Boolean(bar["showCaseNotches"]) : defaults.barSettings.showCaseNotches,
                overlapMode: String(bar?.["overlapMode"] ?? defaults.barSettings.overlapMode)
            },
            caseLabelSettings: {
                show: cl?.["show"] !== undefined ? Boolean(cl["show"]) : defaults.caseLabelSettings.show,
                labelPosition: String(cl?.["labelPosition"] ?? defaults.caseLabelSettings.labelPosition),
                labelContent: String(cl?.["labelContent"] ?? defaults.caseLabelSettings.labelContent),
                fontSize: cl?.["fontSize"] !== undefined ? Number(cl["fontSize"]) : defaults.caseLabelSettings.fontSize,
                fontColor: (cl?.["fontColor"] as { solid?: { color?: string } })?.solid?.color ?? defaults.caseLabelSettings.fontColor,
                badgeBgColor: (cl?.["badgeBgColor"] as { solid?: { color?: string } })?.solid?.color ?? defaults.caseLabelSettings.badgeBgColor,
                prefix: String(cl?.["prefix"] ?? defaults.caseLabelSettings.prefix),
                suffix: String(cl?.["suffix"] ?? defaults.caseLabelSettings.suffix)
            },
            timeAxisSettings: {
                timeFormat: String(ax?.["timeFormat"] ?? defaults.timeAxisSettings.timeFormat),
                clampToSurgicalDay: ax?.["clampToSurgicalDay"] !== undefined ? Boolean(ax["clampToSurgicalDay"]) : defaults.timeAxisSettings.clampToSurgicalDay,
                dayStartHour: ax?.["dayStartHour"] !== undefined ? Number(ax["dayStartHour"]) : defaults.timeAxisSettings.dayStartHour,
                dayEndHour: ax?.["dayEndHour"] !== undefined ? Number(ax["dayEndHour"]) : defaults.timeAxisSettings.dayEndHour,
                fontSize: ax?.["fontSize"] !== undefined ? Number(ax["fontSize"]) : defaults.timeAxisSettings.fontSize,
                fontColor: (ax?.["fontColor"] as { solid?: { color?: string } })?.solid?.color ?? defaults.timeAxisSettings.fontColor,
                showGridlines: ax?.["showGridlines"] !== undefined ? Boolean(ax["showGridlines"]) : defaults.timeAxisSettings.showGridlines,
                gridlineColor: (ax?.["gridlineColor"] as { solid?: { color?: string } })?.solid?.color ?? defaults.timeAxisSettings.gridlineColor
            },
            referenceLineSettings: {
                showNowLine: ref?.["showNowLine"] !== undefined ? Boolean(ref["showNowLine"]) : defaults.referenceLineSettings.showNowLine,
                nowLineColor: (ref?.["nowLineColor"] as { solid?: { color?: string } })?.solid?.color ?? defaults.referenceLineSettings.nowLineColor,
                showShiftChangeLines: ref?.["showShiftChangeLines"] !== undefined ? Boolean(ref["showShiftChangeLines"]) : defaults.referenceLineSettings.showShiftChangeLines,
                shiftChangeColor: (ref?.["shiftChangeColor"] as { solid?: { color?: string } })?.solid?.color ?? defaults.referenceLineSettings.shiftChangeColor
            },
            overtimeSettings: {
                highlightLongShifts: ot?.["highlightLongShifts"] !== undefined ? Boolean(ot["highlightLongShifts"]) : defaults.overtimeSettings.highlightLongShifts,
                thresholdHours: ot?.["thresholdHours"] !== undefined ? Number(ot["thresholdHours"]) : defaults.overtimeSettings.thresholdHours,
                warningBorderColor: (ot?.["warningBorderColor"] as { solid?: { color?: string } })?.solid?.color ?? defaults.overtimeSettings.warningBorderColor
            },
            sortSettings: {
                locationSort: String(srt?.["locationSort"] ?? defaults.sortSettings.locationSort),
                providerSort: String(srt?.["providerSort"] ?? defaults.sortSettings.providerSort)
            }
        };
    }

    private getColorScale(categories: string[]): (cat: string) => string {
        const pal = this.settings.colorPaletteSettings.palette;
        if (pal === "powerBi") {
            const map = new Map<string, string>();
            categories.forEach(cat => map.set(cat, this.host.colorPalette.getColor(cat).value));
            return (cat: string) => map.get(cat) || "#3B82F6";
        }

        let scheme: readonly string[] = d3.schemeTableau10;
        if (pal === "category10") scheme = d3.schemeCategory10;
        else if (pal === "paired") scheme = d3.schemePaired;
        else if (pal === "accent") scheme = d3.schemeAccent;
        else if (pal === "dark2") scheme = d3.schemeDark2;
        else if (pal === "set2") scheme = d3.schemeSet2;

        const scale = d3.scaleOrdinal(scheme).domain(categories);
        return (cat: string) => scale(cat);
    }

    public update(options: VisualUpdateOptions) {
        this.legendContainer.selectAll("*").remove();
        this.chartContainer.selectAll("*").remove();

        const dataView = options.dataViews && options.dataViews[0];
        if (!dataView || !dataView.categorical) return;

        this.settings = this.parseSettings(dataView);

        const viewportWidth = options.viewport.width;
        const categories = dataView.categorical.categories || [];
        const values = dataView.categorical.values || [];
        if (categories.length < 5) return;

        const locCol = categories.find(c => c.source.roles && c.source.roles["location"]) || categories[0];
        const provCol = categories.find(c => c.source.roles && c.source.roles["provider"]) || categories[1];
        const catCol = categories.find(c => c.source.roles && c.source.roles["category"]) || categories[2];
        const startCol = categories.find(c => c.source.roles && c.source.roles["startDate"]) || categories[3];
        const endCol = categories.find(c => c.source.roles && c.source.roles["endDate"]) || categories[4];

        const caseCol = values.find(v => v.source.roles && v.source.roles["caseCount"]) || values[0];
        const extraCol = values.find(v => v.source.roles && v.source.roles["extraMetric"]) || values[1];

        const rowCount = locCol.values.length;
        const rawShifts: ShiftItem[] = [];
        let minDate: Date = new Date(8640000000000000);
        let maxDate: Date = new Date(-8640000000000000);

        for (let i = 0; i < rowCount; i++) {
            const rawStart = startCol.values[i];
            const rawEnd = endCol.values[i];
            if (!rawStart) continue;

            const startDate = new Date(String(rawStart));
            let endDate = rawEnd ? new Date(String(rawEnd)) : new Date(startDate.getTime() + 3600000);
            if (isNaN(startDate.getTime())) continue;
            if (isNaN(endDate.getTime()) || endDate <= startDate) {
                endDate = new Date(startDate.getTime() + 3600000);
            }

            if (startDate < minDate) minDate = startDate;
            if (endDate > maxDate) maxDate = endDate;

            const selectionId = this.host.createSelectionIdBuilder()
                .withCategory(provCol, i)
                .withCategory(catCol, i)
                .createSelectionId();

            rawShifts.push({
                id: `${i}`,
                location: String(locCol.values[i] ?? "Main OR"),
                provider: String(provCol.values[i] ?? "Staff"),
                category: String(catCol.values[i] ?? "Clinical"),
                startDate,
                endDate,
                caseCount: Number(caseCol?.values[i]) || 0,
                extraMetric: Number(extraCol?.values[i]) || 0,
                subRowIndex: 0,
                selectionId
            });
        }

        if (rawShifts.length === 0) return;

        // Optional Day Clamping (e.g., 07:00 - 19:00 surgical day window)
        if (this.settings.timeAxisSettings.clampToSurgicalDay) {
            const startH = Math.max(0, Math.min(23, this.settings.timeAxisSettings.dayStartHour));
            const endH = Math.max(1, Math.min(24, this.settings.timeAxisSettings.dayEndHour));
            const clampedMin = new Date(minDate);
            clampedMin.setHours(startH, 0, 0, 0);
            const clampedMax = new Date(maxDate);
            clampedMax.setHours(endH, 0, 0, 0);
            if (clampedMax > clampedMin) {
                minDate = clampedMin;
                maxDate = clampedMax;
            }
        }

        if (minDate.getTime() === maxDate.getTime()) {
            maxDate = new Date(minDate.getTime() + 86400000);
        }

        const uniqueCategories = Array.from(new Set(rawShifts.map(s => s.category)));
        const colorScale = this.getColorScale(uniqueCategories);

        // Group Hierarchy: Location -> Provider with Concurrency Layout
        const locMap = new Map<string, Map<string, ShiftItem[]>>();
        rawShifts.forEach(s => {
            if (!locMap.has(s.location)) locMap.set(s.location, new Map<string, ShiftItem[]>());
            const pMap = locMap.get(s.location)!;
            if (!pMap.has(s.provider)) pMap.set(s.provider, []);
            pMap.get(s.provider)!.push(s);
        });

        const locations: LocationGroup[] = [];
        locMap.forEach((pMap, locName) => {
            const providers: ProviderRow[] = [];
            let locCases = 0;

            pMap.forEach((shiftList, pName) => {
                shiftList.sort((a, b) => a.startDate.getTime() - b.startDate.getTime());

                // Stagger algorithm for overlapping concurrent shifts
                let maxSubRows = 1;
                if (this.settings.barSettings.overlapMode === "stagger") {
                    const lanes: Date[] = [];
                    shiftList.forEach(s => {
                        let assignedLane = -1;
                        for (let l = 0; l < lanes.length; l++) {
                            if (s.startDate >= lanes[l]) {
                                assignedLane = l;
                                lanes[l] = s.endDate;
                                break;
                            }
                        }
                        if (assignedLane === -1) {
                            assignedLane = lanes.length;
                            lanes.push(s.endDate);
                        }
                        s.subRowIndex = assignedLane;
                    });
                    maxSubRows = Math.max(1, lanes.length);
                }

                let pCases = 0;
                let earliest = shiftList[0].startDate;
                shiftList.forEach(s => {
                    pCases += s.caseCount;
                    if (s.startDate < earliest) earliest = s.startDate;
                });
                locCases += pCases;

                providers.push({
                    providerName: pName,
                    shifts: shiftList,
                    maxSubRows,
                    totalCases: pCases,
                    earliestShift: earliest
                });
            });

            // Provider sorting inside location
            const pSort = this.settings.sortSettings.providerSort;
            if (pSort === "asc") providers.sort((a, b) => a.providerName.localeCompare(b.providerName));
            else if (pSort === "desc") providers.sort((a, b) => b.providerName.localeCompare(a.providerName));
            else if (pSort === "caseDesc") providers.sort((a, b) => b.totalCases - a.totalCases);
            else if (pSort === "caseAsc") providers.sort((a, b) => a.totalCases - b.totalCases);
            else if (pSort === "startAsc") providers.sort((a, b) => a.earliestShift.getTime() - b.earliestShift.getTime());

            locations.push({ locationName: locName, providers, totalCases: locCases });
        });

        // Location Sorting
        const lSort = this.settings.sortSettings.locationSort;
        if (lSort === "asc") locations.sort((a, b) => a.locationName.localeCompare(b.locationName));
        else if (lSort === "desc") locations.sort((a, b) => b.locationName.localeCompare(a.locationName));
        else if (lSort === "caseDesc") locations.sort((a, b) => b.totalCases - a.totalCases);
        else if (lSort === "caseAsc") locations.sort((a, b) => a.totalCases - b.totalCases);

        // Render Legend
        let legendMarginTop = 10;
        if (this.settings.legendSettings.show && uniqueCategories.length > 0) {
            legendMarginTop = this.renderLegend(uniqueCategories, colorScale, viewportWidth);
        }

        const provColWidth = this.settings.swimlaneSettings.providerColumnWidth;
        const rightMargin = 25;
        const timelineX = provColWidth + 12;
        const timelineWidth = Math.max(100, viewportWidth - timelineX - rightMargin);

        const timeScale = d3.scaleTime()
            .domain([minDate, maxDate])
            .range([0, timelineWidth]);

        const baseRowHeight = this.settings.swimlaneSettings.rowHeight;
        const headerHeight = this.settings.locationHeaderSettings.headerHeight;
        const timeAxisHeight = 42;

        let totalHeight = timeAxisHeight;
        locations.forEach(loc => {
            if (this.settings.locationHeaderSettings.showHeader) totalHeight += headerHeight;
            loc.providers.forEach(p => {
                totalHeight += baseRowHeight * p.maxSubRows;
            });
            totalHeight += 10;
        });

        this.svg.attr("width", viewportWidth).attr("height", Math.max(options.viewport.height, totalHeight + legendMarginTop + 20));
        this.chartContainer.attr("transform", `translate(0, ${legendMarginTop})`);

        this.svg.on("click", () => {
            this.selectionManager.clear();
            this.svg.selectAll(".gantt-bar").attr("opacity", this.settings.barSettings.barOpacity / 100);
        });

        // Sticky Timeline Axis
        const axisGroup = this.chartContainer.append("g")
            .attr("class", "time-axis-group")
            .attr("transform", `translate(${timelineX}, ${timeAxisHeight - 12})`);

        const timeAxis = d3.axisTop(timeScale)
            .ticks(Math.max(4, Math.floor(timelineWidth / 115)))
            .tickFormat((d) => {
                const date = d as Date;
                if (this.settings.timeAxisSettings.timeFormat === "12hour") return d3.timeFormat("%I:%M %p")(date);
                if (this.settings.timeAxisSettings.timeFormat === "24hour") return d3.timeFormat("%H:%M")(date);
                if (this.settings.timeAxisSettings.timeFormat === "dateTimeShort") return d3.timeFormat("%b %d, %H:%M")(date);
                return d3.timeFormat("%b %d, %H:%M")(date);
            });

        axisGroup.call(timeAxis);
        axisGroup.selectAll("text")
            .attr("font-size", `${this.settings.timeAxisSettings.fontSize}px`)
            .attr("fill", this.settings.timeAxisSettings.fontColor);

        // Render Locations, Provider Rows & Shifts
        let currentY = timeAxisHeight;
        const prefix = cleanAffix(this.settings.caseLabelSettings.prefix);
        const suffix = cleanAffix(this.settings.caseLabelSettings.suffix);

        locations.forEach(loc => {
            // Location Header Banner
            if (this.settings.locationHeaderSettings.showHeader) {
                const locHeader = this.chartContainer.append("g")
                    .attr("class", "location-header")
                    .attr("transform", `translate(0, ${currentY})`);

                locHeader.append("rect")
                    .attr("x", 0)
                    .attr("y", 0)
                    .attr("width", viewportWidth - rightMargin + 8)
                    .attr("height", headerHeight)
                    .attr("fill", this.settings.locationHeaderSettings.headerBgColor)
                    .attr("rx", 3);

                locHeader.append("text")
                    .attr("x", 12)
                    .attr("y", headerHeight / 2 + 4)
                    .attr("font-size", `${this.settings.locationHeaderSettings.headerFontSize}px`)
                    .attr("font-weight", "bold")
                    .attr("fill", this.settings.locationHeaderSettings.headerFontColor)
                    .text(loc.locationName);

                if (this.settings.locationHeaderSettings.showCaseBadge) {
                    const badgeText = `${loc.totalCases} Total Cases`;
                    const badgeWidth = Math.max(80, badgeText.length * 7 + 16);
                    const badgeX = timelineX;

                    locHeader.append("rect")
                        .attr("x", badgeX)
                        .attr("y", 4)
                        .attr("width", badgeWidth)
                        .attr("height", headerHeight - 8)
                        .attr("rx", (headerHeight - 8) / 2)
                        .attr("fill", this.settings.locationHeaderSettings.badgeBgColor);

                    locHeader.append("text")
                        .attr("x", badgeX + badgeWidth / 2)
                        .attr("y", headerHeight / 2 + 3.5)
                        .attr("text-anchor", "middle")
                        .attr("font-size", "9px")
                        .attr("font-weight", "bold")
                        .attr("fill", this.settings.locationHeaderSettings.badgeTextColor)
                        .text(badgeText);
                }

                currentY += headerHeight;
            }

            // Provider Swimlanes
            loc.providers.forEach((prov, pIdx) => {
                const computedRowHeight = baseRowHeight * prov.maxSubRows;
                const rowGroup = this.chartContainer.append("g")
                    .attr("class", "provider-row")
                    .attr("transform", `translate(0, ${currentY})`);

                // Alternating row background
                const rowBgColor = pIdx % 2 === 0 ? this.settings.swimlaneSettings.evenRowColor : this.settings.swimlaneSettings.oddRowColor;
                if (this.settings.swimlaneSettings.showZebraStriping) {
                    rowGroup.append("rect")
                        .attr("x", 0)
                        .attr("y", 0)
                        .attr("width", viewportWidth - rightMargin + 8)
                        .attr("height", computedRowHeight)
                        .attr("fill", rowBgColor);
                }

                // Row Divider
                if (this.settings.swimlaneSettings.showRowBorders) {
                    rowGroup.append("line")
                        .attr("x1", 0)
                        .attr("x2", viewportWidth - rightMargin + 8)
                        .attr("y1", computedRowHeight)
                        .attr("y2", computedRowHeight)
                        .attr("stroke", this.settings.swimlaneSettings.rowBorderColor)
                        .attr("stroke-width", 1);
                }

                // Provider Name Label
                rowGroup.append("text")
                    .attr("x", 12)
                    .attr("y", computedRowHeight / 2 + 4)
                    .attr("font-size", `${this.settings.swimlaneSettings.providerFontSize}px`)
                    .attr("font-weight", "600")
                    .attr("fill", this.settings.swimlaneSettings.providerFontColor)
                    .text(prov.providerName);

                // Vertical Gridlines across row
                if (this.settings.timeAxisSettings.showGridlines) {
                    const ticks = timeScale.ticks(Math.max(4, Math.floor(timelineWidth / 115)));
                    ticks.forEach(t => {
                        const tx = timelineX + timeScale(t);
                        rowGroup.append("line")
                            .attr("x1", tx)
                            .attr("x2", tx)
                            .attr("y1", 0)
                            .attr("y2", computedRowHeight)
                            .attr("stroke", this.settings.timeAxisSettings.gridlineColor)
                            .attr("stroke-dasharray", "2,2");
                    });
                }

                // Render Shift Bars
                const subSlotHeight = computedRowHeight / prov.maxSubRows;
                const barH = subSlotHeight * (this.settings.swimlaneSettings.barHeightPercent / 100);
                const barPadY = (subSlotHeight - barH) / 2;

                prov.shifts.forEach(shift => {
                    const x0 = timelineX + timeScale(shift.startDate);
                    const x1 = timelineX + timeScale(shift.endDate);
                    const barW = Math.max(5, x1 - x0);
                    const barY = shift.subRowIndex * subSlotHeight + barPadY;

                    const durationHrs = (shift.endDate.getTime() - shift.startDate.getTime()) / 3600000;
                    const isOvertime = this.settings.overtimeSettings.highlightLongShifts && (durationHrs >= this.settings.overtimeSettings.thresholdHours);
                    const pacing = shift.caseCount > 0 ? (shift.caseCount / durationHrs).toFixed(1) : "0";

                    const tooltipItems: VisualTooltipDataItem[] = [
                        { displayName: "Location", value: shift.location },
                        { displayName: "Provider", value: shift.provider },
                        { displayName: "Assignment", value: shift.category },
                        { displayName: "Start", value: d3.timeFormat("%m/%d %H:%M")(shift.startDate) },
                        { displayName: "End", value: d3.timeFormat("%m/%d %H:%M")(shift.endDate) },
                        { displayName: "Shift Duration", value: `${durationHrs.toFixed(2)} hrs` },
                        { displayName: "Cases Completed", value: `${shift.caseCount}` },
                        { displayName: "Case Pacing", value: `${pacing} cases/hr` },
                        ...(shift.extraMetric > 0 ? [{ displayName: "Units / Metric", value: d3.format(",.2f")(shift.extraMetric) }] : [])
                    ];

                    const strokeColor = isOvertime ? this.settings.overtimeSettings.warningBorderColor : this.settings.barSettings.borderColor;
                    const strokeWidth = isOvertime ? 2.5 : (this.settings.barSettings.showBorder ? this.settings.barSettings.borderWidth : 0);

                    const barRect = rowGroup.append("rect")
                        .attr("class", "gantt-bar")
                        .attr("x", x0)
                        .attr("y", barY)
                        .attr("width", barW)
                        .attr("height", barH)
                        .attr("rx", this.settings.barSettings.barCornerRadius)
                        .attr("fill", colorScale(shift.category))
                        .attr("fill-opacity", this.settings.barSettings.barOpacity / 100)
                        .attr("stroke", strokeColor)
                        .attr("stroke-width", strokeWidth)
                        .attr("cursor", "pointer");

                    barRect
                        .on("mouseover", (event: MouseEvent) => {
                            barRect.attr("fill-opacity", 1).attr("stroke", "#0F172A");
                            this.tooltipService.show({
                                coordinates: [event.clientX, event.clientY],
                                isTouchEvent: false,
                                dataItems: tooltipItems,
                                identities: []
                            });
                        })
                        .on("mousemove", (event: MouseEvent) => {
                            this.tooltipService.move({
                                coordinates: [event.clientX, event.clientY],
                                isTouchEvent: false,
                                dataItems: tooltipItems,
                                identities: []
                            });
                        })
                        .on("mouseout", () => {
                            barRect.attr("fill-opacity", this.settings.barSettings.barOpacity / 100).attr("stroke", strokeColor);
                            this.tooltipService.hide({ immediately: true, isTouchEvent: false });
                        })
                        .on("click", (event: MouseEvent) => {
                            event.stopPropagation();
                            this.selectionManager.select(shift.selectionId, event.ctrlKey || event.metaKey);
                        });

                    // Micro Case Notches along bottom edge of the bar
                    if (this.settings.barSettings.showCaseNotches && shift.caseCount > 0 && barW > 25) {
                        const notchCount = Math.min(shift.caseCount, 15);
                        const step = barW / (notchCount + 1);
                        for (let k = 1; k <= notchCount; k++) {
                            rowGroup.append("line")
                                .attr("x1", x0 + k * step)
                                .attr("x2", x0 + k * step)
                                .attr("y1", barY + barH - 4)
                                .attr("y2", barY + barH)
                                .attr("stroke", "rgba(255, 255, 255, 0.8)")
                                .attr("stroke-width", 1.5)
                                .attr("pointer-events", "none");
                        }
                    }

                    // Case Data Labels & Floating Badges
                    if (this.settings.caseLabelSettings.show && shift.caseCount > 0) {
                        let content = `${prefix}${shift.caseCount}${suffix}`;
                        if (this.settings.caseLabelSettings.labelContent === "casesAndHours") {
                            content = `${shift.caseCount} cs (${durationHrs.toFixed(1)}h)`;
                        } else if (this.settings.caseLabelSettings.labelContent === "hoursOnly") {
                            content = `${durationHrs.toFixed(1)}h`;
                        }

                        const pos = this.settings.caseLabelSettings.labelPosition;
                        const fontSize = this.settings.caseLabelSettings.fontSize;

                        if (pos === "pillBadge") {
                            const badgeW = Math.max(30, content.length * 6 + 12);
                            const badgeH = Math.min(barH - 4, 18);
                            const badgeX = x0 + Math.max(4, (barW - badgeW) / 2);
                            const badgeY = barY + (barH - badgeH) / 2;

                            if (barW > 35) {
                                rowGroup.append("rect")
                                    .attr("x", badgeX)
                                    .attr("y", badgeY)
                                    .attr("width", badgeW)
                                    .attr("height", badgeH)
                                    .attr("rx", badgeH / 2)
                                    .attr("fill", this.settings.caseLabelSettings.badgeBgColor)
                                    .attr("pointer-events", "none");

                                rowGroup.append("text")
                                    .attr("x", badgeX + badgeW / 2)
                                    .attr("y", badgeY + badgeH / 2 + 3)
                                    .attr("text-anchor", "middle")
                                    .attr("font-size", `${fontSize}px`)
                                    .attr("font-weight", "bold")
                                    .attr("fill", this.settings.caseLabelSettings.fontColor)
                                    .attr("pointer-events", "none")
                                    .text(content);
                            }
                        } else if (pos === "right") {
                            rowGroup.append("text")
                                .attr("x", x0 + barW + 6)
                                .attr("y", barY + barH / 2 + 3.5)
                                .attr("text-anchor", "start")
                                .attr("font-size", `${fontSize}px`)
                                .attr("font-weight", "bold")
                                .attr("fill", "#1E293B")
                                .attr("pointer-events", "none")
                                .text(content);
                        } else if (pos === "above") {
                            rowGroup.append("text")
                                .attr("x", x0 + barW / 2)
                                .attr("y", barY - 3)
                                .attr("text-anchor", "middle")
                                .attr("font-size", `${fontSize}px`)
                                .attr("font-weight", "bold")
                                .attr("fill", "#1E293B")
                                .attr("pointer-events", "none")
                                .text(content);
                        } else {
                            if (barW > 32) {
                                rowGroup.append("text")
                                    .attr("x", x0 + barW / 2)
                                    .attr("y", barY + barH / 2 + 3.5)
                                    .attr("text-anchor", "middle")
                                    .attr("font-size", `${fontSize}px`)
                                    .attr("font-weight", "bold")
                                    .attr("fill", this.settings.caseLabelSettings.fontColor)
                                    .attr("pointer-events", "none")
                                    .text(content);
                            }
                        }
                    }
                });

                currentY += computedRowHeight;
            });

            currentY += 8;
        });

        // Shift Turnover & Now Reference Lines
        const plotBottomY = currentY;
        if (this.settings.referenceLineSettings.showShiftChangeLines) {
            const shiftHours = [7, 15, 19];
            const startDay = new Date(minDate);
            startDay.setHours(0, 0, 0, 0);

            for (let dayOffset = 0; dayOffset <= Math.ceil((maxDate.getTime() - minDate.getTime()) / 86400000); dayOffset++) {
                shiftHours.forEach(hr => {
                    const lineTime = new Date(startDay.getTime() + dayOffset * 86400000 + hr * 3600000);
                    if (lineTime >= minDate && lineTime <= maxDate) {
                        const rx = timelineX + timeScale(lineTime);
                        this.chartContainer.append("line")
                            .attr("x1", rx)
                            .attr("x2", rx)
                            .attr("y1", timeAxisHeight)
                            .attr("y2", plotBottomY)
                            .attr("stroke", this.settings.referenceLineSettings.shiftChangeColor)
                            .attr("stroke-width", 1.5)
                            .attr("stroke-dasharray", "3,3")
                            .attr("pointer-events", "none");
                    }
                });
            }
        }

        if (this.settings.referenceLineSettings.showNowLine) {
            const now = new Date();
            if (now >= minDate && now <= maxDate) {
                const nx = timelineX + timeScale(now);
                this.chartContainer.append("line")
                    .attr("x1", nx)
                    .attr("x2", nx)
                    .attr("y1", timeAxisHeight - 15)
                    .attr("y2", plotBottomY)
                    .attr("stroke", this.settings.referenceLineSettings.nowLineColor)
                    .attr("stroke-width", 2)
                    .attr("pointer-events", "none");

                this.chartContainer.append("text")
                    .attr("x", nx)
                    .attr("y", timeAxisHeight - 18)
                    .attr("text-anchor", "middle")
                    .attr("font-size", "9px")
                    .attr("font-weight", "bold")
                    .attr("fill", this.settings.referenceLineSettings.nowLineColor)
                    .text("NOW");
            }
        }
    }

    private renderLegend(categories: string[], colorScale: (cat: string) => string, totalWidth: number): number {
        const fontSize = this.settings.legendSettings.fontSize;
        const fontColor = this.settings.legendSettings.fontColor;

        let runningX = 20;
        const startY = 18;

        const legItems = this.legendContainer.selectAll(".legend-item")
            .data(categories)
            .enter()
            .append("g")
            .attr("class", "legend-item")
            .attr("cursor", "pointer")
            .on("click", (event: MouseEvent, d: string) => {
                event.stopPropagation();
                this.svg.selectAll(".gantt-bar").each(function () {
                    const el = d3.select(this);
                    const fill = el.attr("fill");
                    el.attr("opacity", fill === colorScale(d) ? 1 : 0.15);
                });
            });

        legItems.each(function (cat) {
            const item = d3.select(this);
            item.attr("transform", `translate(${runningX}, ${startY})`);
            item.append("circle").attr("r", 5).attr("fill", colorScale(cat));
            const txt = item.append("text")
                .attr("x", 10)
                .attr("y", 4)
                .attr("font-size", `${fontSize}px`)
                .attr("fill", fontColor)
                .text(cat);
            const node = txt.node() as SVGGraphicsElement;
            const textWidth = node ? node.getBBox().width : 40;
            runningX += textWidth + 24;
        });

        return 38;
    }

    public getFormattingModel(): powerbi.visuals.FormattingModel {
        const cards: powerbi.visuals.FormattingCard[] = [
            {
                displayName: "Colors & Themes",
                uid: "colorPaletteCard_uid",
                groups: [
                    {
                        displayName: "Palette",
                        uid: "colorPaletteGroup_uid",
                        slices: [
                            {
                                uid: "palette_slice",
                                displayName: "Theme Palette",
                                control: {
                                    type: powerbi.visuals.FormattingComponent.Dropdown,
                                    properties: {
                                        descriptor: { objectName: "colorPaletteSettings", propertyName: "palette" },
                                        value: this.settings.colorPaletteSettings.palette
                                    }
                                }
                            }
                        ]
                    }
                ]
            },
            {
                displayName: "Legend",
                uid: "legendCard_uid",
                groups: [
                    {
                        displayName: "Display",
                        uid: "legendDisplayGroup_uid",
                        slices: [
                            {
                                uid: "legendShow_slice",
                                displayName: "Show Legend",
                                control: {
                                    type: powerbi.visuals.FormattingComponent.ToggleSwitch,
                                    properties: {
                                        descriptor: { objectName: "legendSettings", propertyName: "show" },
                                        value: this.settings.legendSettings.show
                                    }
                                }
                            },
                            {
                                uid: "legendFontSize_slice",
                                displayName: "Text Size",
                                control: {
                                    type: powerbi.visuals.FormattingComponent.NumUpDown,
                                    properties: {
                                        descriptor: { objectName: "legendSettings", propertyName: "fontSize" },
                                        value: this.settings.legendSettings.fontSize
                                    }
                                }
                            },
                            {
                                uid: "legendFontColor_slice",
                                displayName: "Text Color",
                                control: {
                                    type: powerbi.visuals.FormattingComponent.ColorPicker,
                                    properties: {
                                        descriptor: { objectName: "legendSettings", propertyName: "fontColor" },
                                        value: { value: this.settings.legendSettings.fontColor }
                                    }
                                }
                            }
                        ]
                    }
                ]
            },
            {
                displayName: "Swimlanes & Rows",
                uid: "swimlaneCard_uid",
                groups: [
                    {
                        displayName: "Row Sizing",
                        uid: "swimlaneSizeGroup_uid",
                        slices: [
                            {
                                uid: "rowHeight_slice",
                                displayName: "Row Height (px)",
                                control: {
                                    type: powerbi.visuals.FormattingComponent.NumUpDown,
                                    properties: {
                                        descriptor: { objectName: "swimlaneSettings", propertyName: "rowHeight" },
                                        value: this.settings.swimlaneSettings.rowHeight
                                    }
                                }
                            },
                            {
                                uid: "barHeightPercent_slice",
                                displayName: "Bar Thickness (%)",
                                control: {
                                    type: powerbi.visuals.FormattingComponent.Slider,
                                    properties: {
                                        descriptor: { objectName: "swimlaneSettings", propertyName: "barHeightPercent" },
                                        value: this.settings.swimlaneSettings.barHeightPercent
                                    }
                                }
                            },
                            {
                                uid: "providerColWidth_slice",
                                displayName: "Provider Column Width (px)",
                                control: {
                                    type: powerbi.visuals.FormattingComponent.NumUpDown,
                                    properties: {
                                        descriptor: { objectName: "swimlaneSettings", propertyName: "providerColumnWidth" },
                                        value: this.settings.swimlaneSettings.providerColumnWidth
                                    }
                                }
                            },
                            {
                                uid: "providerFontSize_slice",
                                displayName: "Provider Font Size",
                                control: {
                                    type: powerbi.visuals.FormattingComponent.NumUpDown,
                                    properties: {
                                        descriptor: { objectName: "swimlaneSettings", propertyName: "providerFontSize" },
                                        value: this.settings.swimlaneSettings.providerFontSize
                                    }
                                }
                            },
                            {
                                uid: "providerFontColor_slice",
                                displayName: "Provider Font Color",
                                control: {
                                    type: powerbi.visuals.FormattingComponent.ColorPicker,
                                    properties: {
                                        descriptor: { objectName: "swimlaneSettings", propertyName: "providerFontColor" },
                                        value: { value: this.settings.swimlaneSettings.providerFontColor }
                                    }
                                }
                            }
                        ]
                    },
                    {
                        displayName: "Row Styling",
                        uid: "swimlaneStripeGroup_uid",
                        slices: [
                            {
                                uid: "showZebra_slice",
                                displayName: "Alternating Row Fills",
                                control: {
                                    type: powerbi.visuals.FormattingComponent.ToggleSwitch,
                                    properties: {
                                        descriptor: { objectName: "swimlaneSettings", propertyName: "showZebraStriping" },
                                        value: this.settings.swimlaneSettings.showZebraStriping
                                    }
                                }
                            },
                            {
                                uid: "evenRowColor_slice",
                                displayName: "Even Row Background",
                                control: {
                                    type: powerbi.visuals.FormattingComponent.ColorPicker,
                                    properties: {
                                        descriptor: { objectName: "swimlaneSettings", propertyName: "evenRowColor" },
                                        value: { value: this.settings.swimlaneSettings.evenRowColor }
                                    }
                                }
                            },
                            {
                                uid: "oddRowColor_slice",
                                displayName: "Odd Row Background",
                                control: {
                                    type: powerbi.visuals.FormattingComponent.ColorPicker,
                                    properties: {
                                        descriptor: { objectName: "swimlaneSettings", propertyName: "oddRowColor" },
                                        value: { value: this.settings.swimlaneSettings.oddRowColor }
                                    }
                                }
                            },
                            {
                                uid: "showRowBorders_slice",
                                displayName: "Show Row Dividers",
                                control: {
                                    type: powerbi.visuals.FormattingComponent.ToggleSwitch,
                                    properties: {
                                        descriptor: { objectName: "swimlaneSettings", propertyName: "showRowBorders" },
                                        value: this.settings.swimlaneSettings.showRowBorders
                                    }
                                }
                            },
                            {
                                uid: "rowBorderColor_slice",
                                displayName: "Divider Color",
                                control: {
                                    type: powerbi.visuals.FormattingComponent.ColorPicker,
                                    properties: {
                                        descriptor: { objectName: "swimlaneSettings", propertyName: "rowBorderColor" },
                                        value: { value: this.settings.swimlaneSettings.rowBorderColor }
                                    }
                                }
                            }
                        ]
                    }
                ]
            },
            {
                displayName: "Location Headers",
                uid: "locationHeaderCard_uid",
                groups: [
                    {
                        displayName: "Banner Appearance",
                        uid: "locationHeaderGroup_uid",
                        slices: [
                            {
                                uid: "showLocHeader_slice",
                                displayName: "Show Location Header",
                                control: {
                                    type: powerbi.visuals.FormattingComponent.ToggleSwitch,
                                    properties: {
                                        descriptor: { objectName: "locationHeaderSettings", propertyName: "showHeader" },
                                        value: this.settings.locationHeaderSettings.showHeader
                                    }
                                }
                            },
                            {
                                uid: "headerHeight_slice",
                                displayName: "Banner Height (px)",
                                control: {
                                    type: powerbi.visuals.FormattingComponent.NumUpDown,
                                    properties: {
                                        descriptor: { objectName: "locationHeaderSettings", propertyName: "headerHeight" },
                                        value: this.settings.locationHeaderSettings.headerHeight
                                    }
                                }
                            },
                            {
                                uid: "headerFontSize_slice",
                                displayName: "Header Font Size",
                                control: {
                                    type: powerbi.visuals.FormattingComponent.NumUpDown,
                                    properties: {
                                        descriptor: { objectName: "locationHeaderSettings", propertyName: "headerFontSize" },
                                        value: this.settings.locationHeaderSettings.headerFontSize
                                    }
                                }
                            },
                            {
                                uid: "headerFontColor_slice",
                                displayName: "Header Font Color",
                                control: {
                                    type: powerbi.visuals.FormattingComponent.ColorPicker,
                                    properties: {
                                        descriptor: { objectName: "locationHeaderSettings", propertyName: "headerFontColor" },
                                        value: { value: this.settings.locationHeaderSettings.headerFontColor }
                                    }
                                }
                            },
                            {
                                uid: "headerBgColor_slice",
                                displayName: "Background Fill",
                                control: {
                                    type: powerbi.visuals.FormattingComponent.ColorPicker,
                                    properties: {
                                        descriptor: { objectName: "locationHeaderSettings", propertyName: "headerBgColor" },
                                        value: { value: this.settings.locationHeaderSettings.headerBgColor }
                                    }
                                }
                            },
                            {
                                uid: "showCaseBadge_slice",
                                displayName: "Show Case Total Badge",
                                control: {
                                    type: powerbi.visuals.FormattingComponent.ToggleSwitch,
                                    properties: {
                                        descriptor: { objectName: "locationHeaderSettings", propertyName: "showCaseBadge" },
                                        value: this.settings.locationHeaderSettings.showCaseBadge
                                    }
                                }
                            },
                            {
                                uid: "badgeBgColor_slice",
                                displayName: "Badge Background Color",
                                control: {
                                    type: powerbi.visuals.FormattingComponent.ColorPicker,
                                    properties: {
                                        descriptor: { objectName: "locationHeaderSettings", propertyName: "badgeBgColor" },
                                        value: { value: this.settings.locationHeaderSettings.badgeBgColor }
                                    }
                                }
                            },
                            {
                                uid: "badgeTextColor_slice",
                                displayName: "Badge Text Color",
                                control: {
                                    type: powerbi.visuals.FormattingComponent.ColorPicker,
                                    properties: {
                                        descriptor: { objectName: "locationHeaderSettings", propertyName: "badgeTextColor" },
                                        value: { value: this.settings.locationHeaderSettings.badgeTextColor }
                                    }
                                }
                            }
                        ]
                    }
                ]
            },
            {
                displayName: "Gantt Bars & Concurrency",
                uid: "barSettingsCard_uid",
                groups: [
                    {
                        displayName: "Appearance & Layout",
                        uid: "barAppearanceGroup_uid",
                        slices: [
                            {
                                uid: "barOpacity_slice",
                                displayName: "Bar Opacity (%)",
                                control: {
                                    type: powerbi.visuals.FormattingComponent.Slider,
                                    properties: {
                                        descriptor: { objectName: "barSettings", propertyName: "barOpacity" },
                                        value: this.settings.barSettings.barOpacity
                                    }
                                }
                            },
                            {
                                uid: "barCornerRadius_slice",
                                displayName: "Corner Radius (px)",
                                control: {
                                    type: powerbi.visuals.FormattingComponent.NumUpDown,
                                    properties: {
                                        descriptor: { objectName: "barSettings", propertyName: "barCornerRadius" },
                                        value: this.settings.barSettings.barCornerRadius
                                    }
                                }
                            },
                            {
                                uid: "showBorder_slice",
                                displayName: "Show Bar Border",
                                control: {
                                    type: powerbi.visuals.FormattingComponent.ToggleSwitch,
                                    properties: {
                                        descriptor: { objectName: "barSettings", propertyName: "showBorder" },
                                        value: this.settings.barSettings.showBorder
                                    }
                                }
                            },
                            {
                                uid: "borderColor_slice",
                                displayName: "Border Color",
                                control: {
                                    type: powerbi.visuals.FormattingComponent.ColorPicker,
                                    properties: {
                                        descriptor: { objectName: "barSettings", propertyName: "borderColor" },
                                        value: { value: this.settings.barSettings.borderColor }
                                    }
                                }
                            },
                            {
                                uid: "borderWidth_slice",
                                displayName: "Border Width (px)",
                                control: {
                                    type: powerbi.visuals.FormattingComponent.NumUpDown,
                                    properties: {
                                        descriptor: { objectName: "barSettings", propertyName: "borderWidth" },
                                        value: this.settings.barSettings.borderWidth
                                    }
                                }
                            },
                            {
                                uid: "showCaseNotches_slice",
                                displayName: "Show Micro Case Ticks",
                                control: {
                                    type: powerbi.visuals.FormattingComponent.ToggleSwitch,
                                    properties: {
                                        descriptor: { objectName: "barSettings", propertyName: "showCaseNotches" },
                                        value: this.settings.barSettings.showCaseNotches
                                    }
                                }
                            },
                            {
                                uid: "overlapMode_slice",
                                displayName: "Concurrent Shifts Mode",
                                control: {
                                    type: powerbi.visuals.FormattingComponent.Dropdown,
                                    properties: {
                                        descriptor: { objectName: "barSettings", propertyName: "overlapMode" },
                                        value: this.settings.barSettings.overlapMode
                                    }
                                }
                            }
                        ]
                    }
                ]
            },
            {
                displayName: "Case Data Labels & Badges",
                uid: "caseLabelCard_uid",
                groups: [
                    {
                        displayName: "Labels",
                        uid: "caseLabelGroup_uid",
                        slices: [
                            {
                                uid: "showCases_slice",
                                displayName: "Show Case Labels",
                                control: {
                                    type: powerbi.visuals.FormattingComponent.ToggleSwitch,
                                    properties: {
                                        descriptor: { objectName: "caseLabelSettings", propertyName: "show" },
                                        value: this.settings.caseLabelSettings.show
                                    }
                                }
                            },
                            {
                                uid: "casePosition_slice",
                                displayName: "Position",
                                control: {
                                    type: powerbi.visuals.FormattingComponent.Dropdown,
                                    properties: {
                                        descriptor: { objectName: "caseLabelSettings", propertyName: "labelPosition" },
                                        value: this.settings.caseLabelSettings.labelPosition
                                    }
                                }
                            },
                            {
                                uid: "caseContent_slice",
                                displayName: "Label Content",
                                control: {
                                    type: powerbi.visuals.FormattingComponent.Dropdown,
                                    properties: {
                                        descriptor: { objectName: "caseLabelSettings", propertyName: "labelContent" },
                                        value: this.settings.caseLabelSettings.labelContent
                                    }
                                }
                            },
                            {
                                uid: "caseFontSize_slice",
                                displayName: "Font Size",
                                control: {
                                    type: powerbi.visuals.FormattingComponent.NumUpDown,
                                    properties: {
                                        descriptor: { objectName: "caseLabelSettings", propertyName: "fontSize" },
                                        value: this.settings.caseLabelSettings.fontSize
                                    }
                                }
                            },
                            {
                                uid: "caseFontColor_slice",
                                displayName: "Font Color",
                                control: {
                                    type: powerbi.visuals.FormattingComponent.ColorPicker,
                                    properties: {
                                        descriptor: { objectName: "caseLabelSettings", propertyName: "fontColor" },
                                        value: { value: this.settings.caseLabelSettings.fontColor }
                                    }
                                }
                            },
                            {
                                uid: "badgeBg_slice",
                                displayName: "Pill Badge Fill",
                                control: {
                                    type: powerbi.visuals.FormattingComponent.ColorPicker,
                                    properties: {
                                        descriptor: { objectName: "caseLabelSettings", propertyName: "badgeBgColor" },
                                        value: { value: this.settings.caseLabelSettings.badgeBgColor }
                                    }
                                }
                            },
                            {
                                uid: "casePrefix_slice",
                                displayName: "Prefix",
                                control: {
                                    type: powerbi.visuals.FormattingComponent.Dropdown,
                                    properties: {
                                        descriptor: { objectName: "caseLabelSettings", propertyName: "prefix" },
                                        value: this.settings.caseLabelSettings.prefix
                                    }
                                }
                            },
                            {
                                uid: "caseSuffix_slice",
                                displayName: "Suffix",
                                control: {
                                    type: powerbi.visuals.FormattingComponent.Dropdown,
                                    properties: {
                                        descriptor: { objectName: "caseLabelSettings", propertyName: "suffix" },
                                        value: this.settings.caseLabelSettings.suffix
                                    }
                                }
                            }
                        ]
                    }
                ]
            },
            {
                displayName: "Timeline & Time Window",
                uid: "timeAxisCard_uid",
                groups: [
                    {
                        displayName: "Scale & Gridlines",
                        uid: "timeAxisGroup_uid",
                        slices: [
                            {
                                uid: "timeFormat_slice",
                                displayName: "Time Axis Format",
                                control: {
                                    type: powerbi.visuals.FormattingComponent.Dropdown,
                                    properties: {
                                        descriptor: { objectName: "timeAxisSettings", propertyName: "timeFormat" },
                                        value: this.settings.timeAxisSettings.timeFormat
                                    }
                                }
                            },
                            {
                                uid: "clampDay_slice",
                                displayName: "Clamp to Daily Surgical Hours",
                                control: {
                                    type: powerbi.visuals.FormattingComponent.ToggleSwitch,
                                    properties: {
                                        descriptor: { objectName: "timeAxisSettings", propertyName: "clampToSurgicalDay" },
                                        value: this.settings.timeAxisSettings.clampToSurgicalDay
                                    }
                                }
                            },
                            {
                                uid: "dayStart_slice",
                                displayName: "Day Start Hour (0-23)",
                                control: {
                                    type: powerbi.visuals.FormattingComponent.NumUpDown,
                                    properties: {
                                        descriptor: { objectName: "timeAxisSettings", propertyName: "dayStartHour" },
                                        value: this.settings.timeAxisSettings.dayStartHour
                                    }
                                }
                            },
                            {
                                uid: "dayEnd_slice",
                                displayName: "Day End Hour (1-24)",
                                control: {
                                    type: powerbi.visuals.FormattingComponent.NumUpDown,
                                    properties: {
                                        descriptor: { objectName: "timeAxisSettings", propertyName: "dayEndHour" },
                                        value: this.settings.timeAxisSettings.dayEndHour
                                    }
                                }
                            },
                            {
                                uid: "axFontSize_slice",
                                displayName: "Font Size",
                                control: {
                                    type: powerbi.visuals.FormattingComponent.NumUpDown,
                                    properties: {
                                        descriptor: { objectName: "timeAxisSettings", propertyName: "fontSize" },
                                        value: this.settings.timeAxisSettings.fontSize
                                    }
                                }
                            },
                            {
                                uid: "axFontColor_slice",
                                displayName: "Font Color",
                                control: {
                                    type: powerbi.visuals.FormattingComponent.ColorPicker,
                                    properties: {
                                        descriptor: { objectName: "timeAxisSettings", propertyName: "fontColor" },
                                        value: { value: this.settings.timeAxisSettings.fontColor }
                                    }
                                }
                            },
                            {
                                uid: "axGrid_slice",
                                displayName: "Show Vertical Gridlines",
                                control: {
                                    type: powerbi.visuals.FormattingComponent.ToggleSwitch,
                                    properties: {
                                        descriptor: { objectName: "timeAxisSettings", propertyName: "showGridlines" },
                                        value: this.settings.timeAxisSettings.showGridlines
                                    }
                                }
                            },
                            {
                                uid: "axGridColor_slice",
                                displayName: "Gridline Color",
                                control: {
                                    type: powerbi.visuals.FormattingComponent.ColorPicker,
                                    properties: {
                                        descriptor: { objectName: "timeAxisSettings", propertyName: "gridlineColor" },
                                        value: { value: this.settings.timeAxisSettings.gridlineColor }
                                    }
                                }
                            }
                        ]
                    }
                ]
            },
            {
                displayName: "Shift & Time Reference Lines",
                uid: "referenceLineCard_uid",
                groups: [
                    {
                        displayName: "Reference Markers",
                        uid: "referenceLineGroup_uid",
                        slices: [
                            {
                                uid: "showNow_slice",
                                displayName: "Show 'Now' Indicator",
                                control: {
                                    type: powerbi.visuals.FormattingComponent.ToggleSwitch,
                                    properties: {
                                        descriptor: { objectName: "referenceLineSettings", propertyName: "showNowLine" },
                                        value: this.settings.referenceLineSettings.showNowLine
                                    }
                                }
                            },
                            {
                                uid: "nowColor_slice",
                                displayName: "Now Indicator Color",
                                control: {
                                    type: powerbi.visuals.FormattingComponent.ColorPicker,
                                    properties: {
                                        descriptor: { objectName: "referenceLineSettings", propertyName: "nowLineColor" },
                                        value: { value: this.settings.referenceLineSettings.nowLineColor }
                                    }
                                }
                            },
                            {
                                uid: "showShiftChange_slice",
                                displayName: "Turnover Lines (07:00/15:00/19:00)",
                                control: {
                                    type: powerbi.visuals.FormattingComponent.ToggleSwitch,
                                    properties: {
                                        descriptor: { objectName: "referenceLineSettings", propertyName: "showShiftChangeLines" },
                                        value: this.settings.referenceLineSettings.showShiftChangeLines
                                    }
                                }
                            },
                            {
                                uid: "shiftChangeColor_slice",
                                displayName: "Turnover Line Color",
                                control: {
                                    type: powerbi.visuals.FormattingComponent.ColorPicker,
                                    properties: {
                                        descriptor: { objectName: "referenceLineSettings", propertyName: "shiftChangeColor" },
                                        value: { value: this.settings.referenceLineSettings.shiftChangeColor }
                                    }
                                }
                            }
                        ]
                    }
                ]
            },
            {
                displayName: "Overtime & Shift Alerts",
                uid: "overtimeCard_uid",
                groups: [
                    {
                        displayName: "Alert Rules",
                        uid: "overtimeGroup_uid",
                        slices: [
                            {
                                uid: "hlLong_slice",
                                displayName: "Flag Extended Shifts",
                                control: {
                                    type: powerbi.visuals.FormattingComponent.ToggleSwitch,
                                    properties: {
                                        descriptor: { objectName: "overtimeSettings", propertyName: "highlightLongShifts" },
                                        value: this.settings.overtimeSettings.highlightLongShifts
                                    }
                                }
                            },
                            {
                                uid: "threshold_slice",
                                displayName: "Shift Threshold (Hours)",
                                control: {
                                    type: powerbi.visuals.FormattingComponent.NumUpDown,
                                    properties: {
                                        descriptor: { objectName: "overtimeSettings", propertyName: "thresholdHours" },
                                        value: this.settings.overtimeSettings.thresholdHours
                                    }
                                }
                            },
                            {
                                uid: "warnColor_slice",
                                displayName: "Alert Border Color",
                                control: {
                                    type: powerbi.visuals.FormattingComponent.ColorPicker,
                                    properties: {
                                        descriptor: { objectName: "overtimeSettings", propertyName: "warningBorderColor" },
                                        value: { value: this.settings.overtimeSettings.warningBorderColor }
                                    }
                                }
                            }
                        ]
                    }
                ]
            },
            {
                displayName: "Sorting & Order",
                uid: "sortCard_uid",
                groups: [
                    {
                        displayName: "Hierarchy Sorting",
                        uid: "sortGroup_uid",
                        slices: [
                            {
                                uid: "locSort_slice",
                                displayName: "Location Order",
                                control: {
                                    type: powerbi.visuals.FormattingComponent.Dropdown,
                                    properties: {
                                        descriptor: { objectName: "sortSettings", propertyName: "locationSort" },
                                        value: this.settings.sortSettings.locationSort
                                    }
                                }
                            },
                            {
                                uid: "provSort_slice",
                                displayName: "Provider Order",
                                control: {
                                    type: powerbi.visuals.FormattingComponent.Dropdown,
                                    properties: {
                                        descriptor: { objectName: "sortSettings", propertyName: "providerSort" },
                                        value: this.settings.sortSettings.providerSort
                                    }
                                }
                            }
                        ]
                    }
                ]
            }
        ];

        return { cards };
    }

    public enumerateObjectInstances(options: EnumerateVisualObjectInstancesOptions): VisualObjectInstance[] | VisualObjectInstanceEnumerationObject {
        const instances: VisualObjectInstance[] = [];

        switch (options.objectName) {
            case "colorPaletteSettings":
                instances.push({
                    objectName: "colorPaletteSettings",
                    selector: (null as any),
                    properties: { palette: this.settings.colorPaletteSettings.palette }
                });
                break;
            case "legendSettings":
                instances.push({
                    objectName: "legendSettings",
                    selector: (null as any),
                    properties: {
                        show: this.settings.legendSettings.show,
                        position: this.settings.legendSettings.position,
                        fontSize: this.settings.legendSettings.fontSize,
                        fontColor: { solid: { color: this.settings.legendSettings.fontColor } }
                    }
                });
                break;
            case "swimlaneSettings":
                instances.push({
                    objectName: "swimlaneSettings",
                    selector: (null as any),
                    properties: {
                        rowHeight: this.settings.swimlaneSettings.rowHeight,
                        barHeightPercent: this.settings.swimlaneSettings.barHeightPercent,
                        providerColumnWidth: this.settings.swimlaneSettings.providerColumnWidth,
                        providerFontSize: this.settings.swimlaneSettings.providerFontSize,
                        providerFontColor: { solid: { color: this.settings.swimlaneSettings.providerFontColor } },
                        showZebraStriping: this.settings.swimlaneSettings.showZebraStriping,
                        evenRowColor: { solid: { color: this.settings.swimlaneSettings.evenRowColor } },
                        oddRowColor: { solid: { color: this.settings.swimlaneSettings.oddRowColor } },
                        showRowBorders: this.settings.swimlaneSettings.showRowBorders,
                        rowBorderColor: { solid: { color: this.settings.swimlaneSettings.rowBorderColor } }
                    }
                });
                break;
            case "locationHeaderSettings":
                instances.push({
                    objectName: "locationHeaderSettings",
                    selector: (null as any),
                    properties: {
                        showHeader: this.settings.locationHeaderSettings.showHeader,
                        headerHeight: this.settings.locationHeaderSettings.headerHeight,
                        headerFontSize: this.settings.locationHeaderSettings.headerFontSize,
                        headerFontColor: { solid: { color: this.settings.locationHeaderSettings.headerFontColor } },
                        headerBgColor: { solid: { color: this.settings.locationHeaderSettings.headerBgColor } },
                        showCaseBadge: this.settings.locationHeaderSettings.showCaseBadge,
                        badgeBgColor: { solid: { color: this.settings.locationHeaderSettings.badgeBgColor } },
                        badgeTextColor: { solid: { color: this.settings.locationHeaderSettings.badgeTextColor } }
                    }
                });
                break;
            case "barSettings":
                instances.push({
                    objectName: "barSettings",
                    selector: (null as any),
                    properties: {
                        barOpacity: this.settings.barSettings.barOpacity,
                        barCornerRadius: this.settings.barSettings.barCornerRadius,
                        showBorder: this.settings.barSettings.showBorder,
                        borderColor: { solid: { color: this.settings.barSettings.borderColor } },
                        borderWidth: this.settings.barSettings.borderWidth,
                        showCaseNotches: this.settings.barSettings.showCaseNotches,
                        overlapMode: this.settings.barSettings.overlapMode
                    }
                });
                break;
            case "caseLabelSettings":
                instances.push({
                    objectName: "caseLabelSettings",
                    selector: (null as any),
                    properties: {
                        show: this.settings.caseLabelSettings.show,
                        labelPosition: this.settings.caseLabelSettings.labelPosition,
                        labelContent: this.settings.caseLabelSettings.labelContent,
                        fontSize: this.settings.caseLabelSettings.fontSize,
                        fontColor: { solid: { color: this.settings.caseLabelSettings.fontColor } },
                        badgeBgColor: { solid: { color: this.settings.caseLabelSettings.badgeBgColor } },
                        prefix: this.settings.caseLabelSettings.prefix,
                        suffix: this.settings.caseLabelSettings.suffix
                    }
                });
                break;
            case "timeAxisSettings":
                instances.push({
                    objectName: "timeAxisSettings",
                    selector: (null as any),
                    properties: {
                        timeFormat: this.settings.timeAxisSettings.timeFormat,
                        clampToSurgicalDay: this.settings.timeAxisSettings.clampToSurgicalDay,
                        dayStartHour: this.settings.timeAxisSettings.dayStartHour,
                        dayEndHour: this.settings.timeAxisSettings.dayEndHour,
                        fontSize: this.settings.timeAxisSettings.fontSize,
                        fontColor: { solid: { color: this.settings.timeAxisSettings.fontColor } },
                        showGridlines: this.settings.timeAxisSettings.showGridlines,
                        gridlineColor: { solid: { color: this.settings.timeAxisSettings.gridlineColor } }
                    }
                });
                break;
            case "referenceLineSettings":
                instances.push({
                    objectName: "referenceLineSettings",
                    selector: (null as any),
                    properties: {
                        showNowLine: this.settings.referenceLineSettings.showNowLine,
                        nowLineColor: { solid: { color: this.settings.referenceLineSettings.nowLineColor } },
                        showShiftChangeLines: this.settings.referenceLineSettings.showShiftChangeLines,
                        shiftChangeColor: { solid: { color: this.settings.referenceLineSettings.shiftChangeColor } }
                    }
                });
                break;
            case "overtimeSettings":
                instances.push({
                    objectName: "overtimeSettings",
                    selector: (null as any),
                    properties: {
                        highlightLongShifts: this.settings.overtimeSettings.highlightLongShifts,
                        thresholdHours: this.settings.overtimeSettings.thresholdHours,
                        warningBorderColor: { solid: { color: this.settings.overtimeSettings.warningBorderColor } }
                    }
                });
                break;
            case "sortSettings":
                instances.push({
                    objectName: "sortSettings",
                    selector: (null as any),
                    properties: {
                        locationSort: this.settings.sortSettings.locationSort,
                        providerSort: this.settings.sortSettings.providerSort
                    }
                });
                break;
        }

        return instances;
    }
}