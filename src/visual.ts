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

interface ScheduleItem {
    id: string;
    groupName: string;
    resourceName: string;
    category: string;
    startDate: Date;
    endDate: Date;
    primaryMetric: number;
    secondaryMetric: number;
    subRowIndex: number;
    selectionId: powerbi.visuals.ISelectionId;
}

interface ResourceRow {
    resourceName: string;
    items: ScheduleItem[];
    maxSubRows: number;
    totalMetric: number;
    earliestStart: Date;
}

interface GroupSwimlane {
    groupName: string;
    resources: ResourceRow[];
    totalMetric: number;
}

interface VisualSettings {
    colorPaletteSettings: { palette: string };
    legendSettings: { show: boolean; position: string; fontSize: number; fontColor: string };
    swimlaneSettings: {
        rowHeight: number;
        barHeightPercent: number;
        resourceColumnWidth: number;
        resourceFontSize: number;
        resourceFontColor: string;
        showZebraStriping: boolean;
        evenRowColor: string;
        oddRowColor: string;
        showRowBorders: boolean;
        rowBorderColor: string;
    };
    groupHeaderSettings: {
        showHeader: boolean;
        headerHeight: number;
        headerFontSize: number;
        headerFontColor: string;
        headerBgColor: string;
        showMetricBadge: boolean;
        badgeBgColor: string;
        badgeTextColor: string;
    };
    barSettings: {
        barOpacity: number;
        barCornerRadius: number;
        showBorder: boolean;
        borderColor: string;
        borderWidth: number;
        showMetricNotches: boolean;
        overlapMode: string;
    };
    labelSettings: {
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
        clampToDailyWindow: boolean;
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
        showMilestoneLines: boolean;
        milestoneLineColor: string;
    };
    thresholdAlertSettings: {
        highlightLongIntervals: boolean;
        thresholdHours: number;
        warningBorderColor: string;
    };
    sortSettings: {
        groupSort: string;
        resourceSort: string;
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
                resourceColumnWidth: 160,
                resourceFontSize: 10,
                resourceFontColor: "#333333",
                showZebraStriping: true,
                evenRowColor: "#FFFFFF",
                oddRowColor: "#F8FAFC",
                showRowBorders: true,
                rowBorderColor: "#E2E8F0"
            },
            groupHeaderSettings: {
                showHeader: true,
                headerHeight: 28,
                headerFontSize: 11,
                headerFontColor: "#0F172A",
                headerBgColor: "#E2E8F0",
                showMetricBadge: true,
                badgeBgColor: "#3B82F6",
                badgeTextColor: "#FFFFFF"
            },
            barSettings: {
                barOpacity: 92,
                barCornerRadius: 4,
                showBorder: true,
                borderColor: "#FFFFFF",
                borderWidth: 1.5,
                showMetricNotches: true,
                overlapMode: "stagger"
            },
            labelSettings: {
                show: true,
                labelPosition: "pillBadge",
                labelContent: "metricOnly",
                fontSize: 9,
                fontColor: "#FFFFFF",
                badgeBgColor: "rgba(15, 23, 42, 0.75)",
                prefix: "none",
                suffix: " cases"
            },
            timeAxisSettings: {
                timeFormat: "auto",
                clampToDailyWindow: false,
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
                showMilestoneLines: true,
                milestoneLineColor: "#94A3B8"
            },
            thresholdAlertSettings: {
                highlightLongIntervals: true,
                thresholdHours: 10,
                warningBorderColor: "#DC2626"
            },
            sortSettings: {
                groupSort: "natural",
                resourceSort: "natural"
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
        const gh = objects["groupHeaderSettings"];
        const bar = objects["barSettings"];
        const lbl = objects["labelSettings"];
        const ax = objects["timeAxisSettings"];
        const ref = objects["referenceLineSettings"];
        const alrt = objects["thresholdAlertSettings"];
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
                resourceColumnWidth: sl?.["resourceColumnWidth"] !== undefined ? Number(sl["resourceColumnWidth"]) : defaults.swimlaneSettings.resourceColumnWidth,
                resourceFontSize: sl?.["resourceFontSize"] !== undefined ? Number(sl["resourceFontSize"]) : defaults.swimlaneSettings.resourceFontSize,
                resourceFontColor: (sl?.["resourceFontColor"] as { solid?: { color?: string } })?.solid?.color ?? defaults.swimlaneSettings.resourceFontColor,
                showZebraStriping: sl?.["showZebraStriping"] !== undefined ? Boolean(sl["showZebraStriping"]) : defaults.swimlaneSettings.showZebraStriping,
                evenRowColor: (sl?.["evenRowColor"] as { solid?: { color?: string } })?.solid?.color ?? defaults.swimlaneSettings.evenRowColor,
                oddRowColor: (sl?.["oddRowColor"] as { solid?: { color?: string } })?.solid?.color ?? defaults.swimlaneSettings.oddRowColor,
                showRowBorders: sl?.["showRowBorders"] !== undefined ? Boolean(sl["showRowBorders"]) : defaults.swimlaneSettings.showRowBorders,
                rowBorderColor: (sl?.["rowBorderColor"] as { solid?: { color?: string } })?.solid?.color ?? defaults.swimlaneSettings.rowBorderColor
            },
            groupHeaderSettings: {
                showHeader: gh?.["showHeader"] !== undefined ? Boolean(gh["showHeader"]) : defaults.groupHeaderSettings.showHeader,
                headerHeight: gh?.["headerHeight"] !== undefined ? Number(gh["headerHeight"]) : defaults.groupHeaderSettings.headerHeight,
                headerFontSize: gh?.["headerFontSize"] !== undefined ? Number(gh["headerFontSize"]) : defaults.groupHeaderSettings.headerFontSize,
                headerFontColor: (gh?.["headerFontColor"] as { solid?: { color?: string } })?.solid?.color ?? defaults.groupHeaderSettings.headerFontColor,
                headerBgColor: (gh?.["headerBgColor"] as { solid?: { color?: string } })?.solid?.color ?? defaults.groupHeaderSettings.headerBgColor,
                showMetricBadge: gh?.["showMetricBadge"] !== undefined ? Boolean(gh["showMetricBadge"]) : defaults.groupHeaderSettings.showMetricBadge,
                badgeBgColor: (gh?.["badgeBgColor"] as { solid?: { color?: string } })?.solid?.color ?? defaults.groupHeaderSettings.badgeBgColor,
                badgeTextColor: (gh?.["badgeTextColor"] as { solid?: { color?: string } })?.solid?.color ?? defaults.groupHeaderSettings.badgeTextColor
            },
            barSettings: {
                barOpacity: bar?.["barOpacity"] !== undefined ? Number(bar["barOpacity"]) : defaults.barSettings.barOpacity,
                barCornerRadius: bar?.["barCornerRadius"] !== undefined ? Number(bar["barCornerRadius"]) : defaults.barSettings.barCornerRadius,
                showBorder: bar?.["showBorder"] !== undefined ? Boolean(bar["showBorder"]) : defaults.barSettings.showBorder,
                borderColor: (bar?.["borderColor"] as { solid?: { color?: string } })?.solid?.color ?? defaults.barSettings.borderColor,
                borderWidth: bar?.["borderWidth"] !== undefined ? Number(bar["borderWidth"]) : defaults.barSettings.borderWidth,
                showMetricNotches: bar?.["showMetricNotches"] !== undefined ? Boolean(bar["showMetricNotches"]) : defaults.barSettings.showMetricNotches,
                overlapMode: String(bar?.["overlapMode"] ?? defaults.barSettings.overlapMode)
            },
            labelSettings: {
                show: lbl?.["show"] !== undefined ? Boolean(lbl["show"]) : defaults.labelSettings.show,
                labelPosition: String(lbl?.["labelPosition"] ?? defaults.labelSettings.labelPosition),
                labelContent: String(lbl?.["labelContent"] ?? defaults.labelSettings.labelContent),
                fontSize: lbl?.["fontSize"] !== undefined ? Number(lbl["fontSize"]) : defaults.labelSettings.fontSize,
                fontColor: (lbl?.["fontColor"] as { solid?: { color?: string } })?.solid?.color ?? defaults.labelSettings.fontColor,
                badgeBgColor: (lbl?.["badgeBgColor"] as { solid?: { color?: string } })?.solid?.color ?? defaults.labelSettings.badgeBgColor,
                prefix: String(lbl?.["prefix"] ?? defaults.labelSettings.prefix),
                suffix: String(lbl?.["suffix"] ?? defaults.labelSettings.suffix)
            },
            timeAxisSettings: {
                timeFormat: String(ax?.["timeFormat"] ?? defaults.timeAxisSettings.timeFormat),
                clampToDailyWindow: ax?.["clampToDailyWindow"] !== undefined ? Boolean(ax["clampToDailyWindow"]) : defaults.timeAxisSettings.clampToDailyWindow,
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
                showMilestoneLines: ref?.["showMilestoneLines"] !== undefined ? Boolean(ref["showMilestoneLines"]) : defaults.referenceLineSettings.showMilestoneLines,
                milestoneLineColor: (ref?.["milestoneLineColor"] as { solid?: { color?: string } })?.solid?.color ?? defaults.referenceLineSettings.milestoneLineColor
            },
            thresholdAlertSettings: {
                highlightLongIntervals: alrt?.["highlightLongIntervals"] !== undefined ? Boolean(alrt["highlightLongIntervals"]) : defaults.thresholdAlertSettings.highlightLongIntervals,
                thresholdHours: alrt?.["thresholdHours"] !== undefined ? Number(alrt["thresholdHours"]) : defaults.thresholdAlertSettings.thresholdHours,
                warningBorderColor: (alrt?.["warningBorderColor"] as { solid?: { color?: string } })?.solid?.color ?? defaults.thresholdAlertSettings.warningBorderColor
            },
            sortSettings: {
                groupSort: String(srt?.["groupSort"] ?? defaults.sortSettings.groupSort),
                resourceSort: String(srt?.["resourceSort"] ?? defaults.sortSettings.resourceSort)
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

        const grpCol = categories.find(c => c.source.roles && c.source.roles["group"]) || categories[0];
        const resCol = categories.find(c => c.source.roles && c.source.roles["item"]) || categories[1];
        const catCol = categories.find(c => c.source.roles && c.source.roles["category"]) || categories[2];
        const startCol = categories.find(c => c.source.roles && c.source.roles["startDate"]) || categories[3];
        const endCol = categories.find(c => c.source.roles && c.source.roles["endDate"]) || categories[4];

        const primaryCol = values.find(v => v.source.roles && v.source.roles["primaryMetric"]) || values[0];
        const secondaryCol = values.find(v => v.source.roles && v.source.roles["secondaryMetric"]) || values[1];

        const rowCount = grpCol.values.length;
        const rawItems: ScheduleItem[] = [];
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
                .withCategory(resCol, i)
                .withCategory(catCol, i)
                .createSelectionId();

            rawItems.push({
                id: `${i}`,
                groupName: String(grpCol.values[i] ?? "Primary Group"),
                resourceName: String(resCol.values[i] ?? "Unassigned"),
                category: String(catCol.values[i] ?? "Standard"),
                startDate,
                endDate,
                primaryMetric: Number(primaryCol?.values[i]) || 0,
                secondaryMetric: Number(secondaryCol?.values[i]) || 0,
                subRowIndex: 0,
                selectionId
            });
        }

        if (rawItems.length === 0) return;

        // Daily Time Window Clamping (e.g. 07:00 to 19:00)
        if (this.settings.timeAxisSettings.clampToDailyWindow) {
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

        const uniqueCategories = Array.from(new Set(rawItems.map(s => s.category)));
        const colorScale = this.getColorScale(uniqueCategories);

        // Group Hierarchy: Group -> Resource with Concurrency Overlap Management
        const groupMap = new Map<string, Map<string, ScheduleItem[]>>();
        rawItems.forEach(s => {
            if (!groupMap.has(s.groupName)) groupMap.set(s.groupName, new Map<string, ScheduleItem[]>());
            const rMap = groupMap.get(s.groupName)!;
            if (!rMap.has(s.resourceName)) rMap.set(s.resourceName, []);
            rMap.get(s.resourceName)!.push(s);
        });

        const groups: GroupSwimlane[] = [];
        groupMap.forEach((rMap, grpName) => {
            const resourceRows: ResourceRow[] = [];
            let grpMetric = 0;

            rMap.forEach((itemList, rName) => {
                itemList.sort((a, b) => a.startDate.getTime() - b.startDate.getTime());

                // Stagger algorithm for concurrent overlapping intervals
                let maxSubRows = 1;
                if (this.settings.barSettings.overlapMode === "stagger") {
                    const lanes: Date[] = [];
                    itemList.forEach(item => {
                        let assignedLane = -1;
                        for (let l = 0; l < lanes.length; l++) {
                            if (item.startDate >= lanes[l]) {
                                assignedLane = l;
                                lanes[l] = item.endDate;
                                break;
                            }
                        }
                        if (assignedLane === -1) {
                            assignedLane = lanes.length;
                            lanes.push(item.endDate);
                        }
                        item.subRowIndex = assignedLane;
                    });
                    maxSubRows = Math.max(1, lanes.length);
                }

                let rMetric = 0;
                let earliest = itemList[0].startDate;
                itemList.forEach(item => {
                    rMetric += item.primaryMetric;
                    if (item.startDate < earliest) earliest = item.startDate;
                });
                grpMetric += rMetric;

                resourceRows.push({
                    resourceName: rName,
                    items: itemList,
                    maxSubRows,
                    totalMetric: rMetric,
                    earliestStart: earliest
                });
            });

            // Resource Sorting inside Group
            const rSort = this.settings.sortSettings.resourceSort;
            if (rSort === "asc") resourceRows.sort((a, b) => a.resourceName.localeCompare(b.resourceName));
            else if (rSort === "desc") resourceRows.sort((a, b) => b.resourceName.localeCompare(a.resourceName));
            else if (rSort === "metricDesc") resourceRows.sort((a, b) => b.totalMetric - a.totalMetric);
            else if (rSort === "metricAsc") resourceRows.sort((a, b) => a.totalMetric - b.totalMetric);
            else if (rSort === "startAsc") resourceRows.sort((a, b) => a.earliestStart.getTime() - b.earliestStart.getTime());

            groups.push({ groupName: grpName, resources: resourceRows, totalMetric: grpMetric });
        });

        // Group Swimlane Sorting
        const gSort = this.settings.sortSettings.groupSort;
        if (gSort === "asc") groups.sort((a, b) => a.groupName.localeCompare(b.groupName));
        else if (gSort === "desc") groups.sort((a, b) => b.groupName.localeCompare(a.groupName));
        else if (gSort === "metricDesc") groups.sort((a, b) => b.totalMetric - a.totalMetric);
        else if (gSort === "metricAsc") groups.sort((a, b) => a.totalMetric - b.totalMetric);

        // Legend Rendering
        let legendMarginTop = 10;
        if (this.settings.legendSettings.show && uniqueCategories.length > 0) {
            legendMarginTop = this.renderLegend(uniqueCategories, colorScale, viewportWidth);
        }

        const resColWidth = this.settings.swimlaneSettings.resourceColumnWidth;
        const rightMargin = 25;
        const timelineX = resColWidth + 12;
        const timelineWidth = Math.max(100, viewportWidth - timelineX - rightMargin);

        const timeScale = d3.scaleTime()
            .domain([minDate, maxDate])
            .range([0, timelineWidth]);

        const baseRowHeight = this.settings.swimlaneSettings.rowHeight;
        const headerHeight = this.settings.groupHeaderSettings.headerHeight;
        const timeAxisHeight = 42;

        let totalHeight = timeAxisHeight;
        groups.forEach(g => {
            if (this.settings.groupHeaderSettings.showHeader) totalHeight += headerHeight;
            g.resources.forEach(r => {
                totalHeight += baseRowHeight * r.maxSubRows;
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

        // Render Groups, Swimlanes & Interval Bars
        let currentY = timeAxisHeight;
        const prefix = cleanAffix(this.settings.labelSettings.prefix);
        const suffix = cleanAffix(this.settings.labelSettings.suffix);

        groups.forEach(grp => {
            // Group Header Banner
            if (this.settings.groupHeaderSettings.showHeader) {
                const groupHeader = this.chartContainer.append("g")
                    .attr("class", "group-header")
                    .attr("transform", `translate(0, ${currentY})`);

                groupHeader.append("rect")
                    .attr("x", 0)
                    .attr("y", 0)
                    .attr("width", viewportWidth - rightMargin + 8)
                    .attr("height", headerHeight)
                    .attr("fill", this.settings.groupHeaderSettings.headerBgColor)
                    .attr("rx", 3);

                groupHeader.append("text")
                    .attr("x", 12)
                    .attr("y", headerHeight / 2 + 4)
                    .attr("font-size", `${this.settings.groupHeaderSettings.headerFontSize}px`)
                    .attr("font-weight", "bold")
                    .attr("fill", this.settings.groupHeaderSettings.headerFontColor)
                    .text(grp.groupName);

                if (this.settings.groupHeaderSettings.showMetricBadge) {
                    const badgeText = `${grp.totalMetric} Total`;
                    const badgeWidth = Math.max(70, badgeText.length * 7 + 16);
                    const badgeX = timelineX;

                    groupHeader.append("rect")
                        .attr("x", badgeX)
                        .attr("y", 4)
                        .attr("width", badgeWidth)
                        .attr("height", headerHeight - 8)
                        .attr("rx", (headerHeight - 8) / 2)
                        .attr("fill", this.settings.groupHeaderSettings.badgeBgColor);

                    groupHeader.append("text")
                        .attr("x", badgeX + badgeWidth / 2)
                        .attr("y", headerHeight / 2 + 3.5)
                        .attr("text-anchor", "middle")
                        .attr("font-size", "9px")
                        .attr("font-weight", "bold")
                        .attr("fill", this.settings.groupHeaderSettings.badgeTextColor)
                        .text(badgeText);
                }

                currentY += headerHeight;
            }

            // Resource Swimlanes
            grp.resources.forEach((res, rIdx) => {
                const computedRowHeight = baseRowHeight * res.maxSubRows;
                const rowGroup = this.chartContainer.append("g")
                    .attr("class", "resource-row")
                    .attr("transform", `translate(0, ${currentY})`);

                // Alternating row background
                const rowBgColor = rIdx % 2 === 0 ? this.settings.swimlaneSettings.evenRowColor : this.settings.swimlaneSettings.oddRowColor;
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

                // Resource Name Label
                rowGroup.append("text")
                    .attr("x", 12)
                    .attr("y", computedRowHeight / 2 + 4)
                    .attr("font-size", `${this.settings.swimlaneSettings.resourceFontSize}px`)
                    .attr("font-weight", "600")
                    .attr("fill", this.settings.swimlaneSettings.resourceFontColor)
                    .text(res.resourceName);

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

                // Render Interval Bars
                const subSlotHeight = computedRowHeight / res.maxSubRows;
                const barH = subSlotHeight * (this.settings.swimlaneSettings.barHeightPercent / 100);
                const barPadY = (subSlotHeight - barH) / 2;

                res.items.forEach(item => {
                    const x0 = timelineX + timeScale(item.startDate);
                    const x1 = timelineX + timeScale(item.endDate);
                    const barW = Math.max(5, x1 - x0);
                    const barY = item.subRowIndex * subSlotHeight + barPadY;

                    const durationHrs = (item.endDate.getTime() - item.startDate.getTime()) / 3600000;
                    const isAlert = this.settings.thresholdAlertSettings.highlightLongIntervals && (durationHrs >= this.settings.thresholdAlertSettings.thresholdHours);
                    const pacing = item.primaryMetric > 0 ? (item.primaryMetric / durationHrs).toFixed(1) : "0";

                    const tooltipItems: VisualTooltipDataItem[] = [
                        { displayName: "Group", value: item.groupName },
                        { displayName: "Resource", value: item.resourceName },
                        { displayName: "Category", value: item.category },
                        { displayName: "Start", value: d3.timeFormat("%m/%d %H:%M")(item.startDate) },
                        { displayName: "End", value: d3.timeFormat("%m/%d %H:%M")(item.endDate) },
                        { displayName: "Duration", value: `${durationHrs.toFixed(2)} hrs` },
                        { displayName: "Primary Metric", value: `${item.primaryMetric}` },
                        { displayName: "Rate / Pacing", value: `${pacing}/hr` },
                        ...(item.secondaryMetric > 0 ? [{ displayName: "Secondary Metric", value: d3.format(",.2f")(item.secondaryMetric) }] : [])
                    ];

                    const strokeColor = isAlert ? this.settings.thresholdAlertSettings.warningBorderColor : this.settings.barSettings.borderColor;
                    const strokeWidth = isAlert ? 2.5 : (this.settings.barSettings.showBorder ? this.settings.barSettings.borderWidth : 0);

                    const barRect = rowGroup.append("rect")
                        .attr("class", "gantt-bar")
                        .attr("x", x0)
                        .attr("y", barY)
                        .attr("width", barW)
                        .attr("height", barH)
                        .attr("rx", this.settings.barSettings.barCornerRadius)
                        .attr("fill", colorScale(item.category))
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
                            this.selectionManager.select(item.selectionId, event.ctrlKey || event.metaKey);
                        });

                    // Micro Metric Notches along bottom edge of the bar
                    if (this.settings.barSettings.showMetricNotches && item.primaryMetric > 0 && barW > 25) {
                        const notchCount = Math.min(item.primaryMetric, 15);
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

                    // Data Labels & Floating Badges
                    if (this.settings.labelSettings.show && item.primaryMetric > 0) {
                        let content = `${prefix}${item.primaryMetric}${suffix}`;
                        if (this.settings.labelSettings.labelContent === "metricAndDuration") {
                            content = `${item.primaryMetric} (${durationHrs.toFixed(1)}h)`;
                        } else if (this.settings.labelSettings.labelContent === "durationOnly") {
                            content = `${durationHrs.toFixed(1)}h`;
                        }

                        const pos = this.settings.labelSettings.labelPosition;
                        const fontSize = this.settings.labelSettings.fontSize;

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
                                    .attr("fill", this.settings.labelSettings.badgeBgColor)
                                    .attr("pointer-events", "none");

                                rowGroup.append("text")
                                    .attr("x", badgeX + badgeW / 2)
                                    .attr("y", badgeY + badgeH / 2 + 3)
                                    .attr("text-anchor", "middle")
                                    .attr("font-size", `${fontSize}px`)
                                    .attr("font-weight", "bold")
                                    .attr("fill", this.settings.labelSettings.fontColor)
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
                                    .attr("fill", this.settings.labelSettings.fontColor)
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

        // Milestone and Live Time Reference Lines
        const plotBottomY = currentY;
        if (this.settings.referenceLineSettings.showMilestoneLines) {
            const milestoneHours = [7, 15, 19];
            const startDay = new Date(minDate);
            startDay.setHours(0, 0, 0, 0);

            for (let dayOffset = 0; dayOffset <= Math.ceil((maxDate.getTime() - minDate.getTime()) / 86400000); dayOffset++) {
                milestoneHours.forEach(hr => {
                    const lineTime = new Date(startDay.getTime() + dayOffset * 86400000 + hr * 3600000);
                    if (lineTime >= minDate && lineTime <= maxDate) {
                        const rx = timelineX + timeScale(lineTime);
                        this.chartContainer.append("line")
                            .attr("x1", rx)
                            .attr("x2", rx)
                            .attr("y1", timeAxisHeight)
                            .attr("y2", plotBottomY)
                            .attr("stroke", this.settings.referenceLineSettings.milestoneLineColor)
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
                displayName: "Swimlanes & Layout",
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
                                uid: "resourceColWidth_slice",
                                displayName: "Resource Column Width (px)",
                                control: {
                                    type: powerbi.visuals.FormattingComponent.NumUpDown,
                                    properties: {
                                        descriptor: { objectName: "swimlaneSettings", propertyName: "resourceColumnWidth" },
                                        value: this.settings.swimlaneSettings.resourceColumnWidth
                                    }
                                }
                            },
                            {
                                uid: "resourceFontSize_slice",
                                displayName: "Resource Font Size",
                                control: {
                                    type: powerbi.visuals.FormattingComponent.NumUpDown,
                                    properties: {
                                        descriptor: { objectName: "swimlaneSettings", propertyName: "resourceFontSize" },
                                        value: this.settings.swimlaneSettings.resourceFontSize
                                    }
                                }
                            },
                            {
                                uid: "resourceFontColor_slice",
                                displayName: "Resource Font Color",
                                control: {
                                    type: powerbi.visuals.FormattingComponent.ColorPicker,
                                    properties: {
                                        descriptor: { objectName: "swimlaneSettings", propertyName: "resourceFontColor" },
                                        value: { value: this.settings.swimlaneSettings.resourceFontColor }
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
                displayName: "Group Headers",
                uid: "groupHeaderCard_uid",
                groups: [
                    {
                        displayName: "Banner Appearance",
                        uid: "groupHeaderGroup_uid",
                        slices: [
                            {
                                uid: "showGrpHeader_slice",
                                displayName: "Show Group Header",
                                control: {
                                    type: powerbi.visuals.FormattingComponent.ToggleSwitch,
                                    properties: {
                                        descriptor: { objectName: "groupHeaderSettings", propertyName: "showHeader" },
                                        value: this.settings.groupHeaderSettings.showHeader
                                    }
                                }
                            },
                            {
                                uid: "headerHeight_slice",
                                displayName: "Banner Height (px)",
                                control: {
                                    type: powerbi.visuals.FormattingComponent.NumUpDown,
                                    properties: {
                                        descriptor: { objectName: "groupHeaderSettings", propertyName: "headerHeight" },
                                        value: this.settings.groupHeaderSettings.headerHeight
                                    }
                                }
                            },
                            {
                                uid: "headerFontSize_slice",
                                displayName: "Header Font Size",
                                control: {
                                    type: powerbi.visuals.FormattingComponent.NumUpDown,
                                    properties: {
                                        descriptor: { objectName: "groupHeaderSettings", propertyName: "headerFontSize" },
                                        value: this.settings.groupHeaderSettings.headerFontSize
                                    }
                                }
                            },
                            {
                                uid: "headerFontColor_slice",
                                displayName: "Header Font Color",
                                control: {
                                    type: powerbi.visuals.FormattingComponent.ColorPicker,
                                    properties: {
                                        descriptor: { objectName: "groupHeaderSettings", propertyName: "headerFontColor" },
                                        value: { value: this.settings.groupHeaderSettings.headerFontColor }
                                    }
                                }
                            },
                            {
                                uid: "headerBgColor_slice",
                                displayName: "Background Fill",
                                control: {
                                    type: powerbi.visuals.FormattingComponent.ColorPicker,
                                    properties: {
                                        descriptor: { objectName: "groupHeaderSettings", propertyName: "headerBgColor" },
                                        value: { value: this.settings.groupHeaderSettings.headerBgColor }
                                    }
                                }
                            },
                            {
                                uid: "showMetricBadge_slice",
                                displayName: "Show Metric Total Badge",
                                control: {
                                    type: powerbi.visuals.FormattingComponent.ToggleSwitch,
                                    properties: {
                                        descriptor: { objectName: "groupHeaderSettings", propertyName: "showMetricBadge" },
                                        value: this.settings.groupHeaderSettings.showMetricBadge
                                    }
                                }
                            },
                            {
                                uid: "badgeBgColor_slice",
                                displayName: "Badge Background Color",
                                control: {
                                    type: powerbi.visuals.FormattingComponent.ColorPicker,
                                    properties: {
                                        descriptor: { objectName: "groupHeaderSettings", propertyName: "badgeBgColor" },
                                        value: { value: this.settings.groupHeaderSettings.badgeBgColor }
                                    }
                                }
                            },
                            {
                                uid: "badgeTextColor_slice",
                                displayName: "Badge Text Color",
                                control: {
                                    type: powerbi.visuals.FormattingComponent.ColorPicker,
                                    properties: {
                                        descriptor: { objectName: "groupHeaderSettings", propertyName: "badgeTextColor" },
                                        value: { value: this.settings.groupHeaderSettings.badgeTextColor }
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
                                uid: "showMetricNotches_slice",
                                displayName: "Show Micro Metric Ticks",
                                control: {
                                    type: powerbi.visuals.FormattingComponent.ToggleSwitch,
                                    properties: {
                                        descriptor: { objectName: "barSettings", propertyName: "showMetricNotches" },
                                        value: this.settings.barSettings.showMetricNotches
                                    }
                                }
                            },
                            {
                                uid: "overlapMode_slice",
                                displayName: "Concurrent Overlap Mode",
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
                displayName: "Data Labels & Badges",
                uid: "labelCard_uid",
                groups: [
                    {
                        displayName: "Labels",
                        uid: "labelGroup_uid",
                        slices: [
                            {
                                uid: "showLabels_slice",
                                displayName: "Show Data Labels",
                                control: {
                                    type: powerbi.visuals.FormattingComponent.ToggleSwitch,
                                    properties: {
                                        descriptor: { objectName: "labelSettings", propertyName: "show" },
                                        value: this.settings.labelSettings.show
                                    }
                                }
                            },
                            {
                                uid: "labelPosition_slice",
                                displayName: "Position",
                                control: {
                                    type: powerbi.visuals.FormattingComponent.Dropdown,
                                    properties: {
                                        descriptor: { objectName: "labelSettings", propertyName: "labelPosition" },
                                        value: this.settings.labelSettings.labelPosition
                                    }
                                }
                            },
                            {
                                uid: "labelContent_slice",
                                displayName: "Label Content",
                                control: {
                                    type: powerbi.visuals.FormattingComponent.Dropdown,
                                    properties: {
                                        descriptor: { objectName: "labelSettings", propertyName: "labelContent" },
                                        value: this.settings.labelSettings.labelContent
                                    }
                                }
                            },
                            {
                                uid: "labelFontSize_slice",
                                displayName: "Font Size",
                                control: {
                                    type: powerbi.visuals.FormattingComponent.NumUpDown,
                                    properties: {
                                        descriptor: { objectName: "labelSettings", propertyName: "fontSize" },
                                        value: this.settings.labelSettings.fontSize
                                    }
                                }
                            },
                            {
                                uid: "labelFontColor_slice",
                                displayName: "Font Color",
                                control: {
                                    type: powerbi.visuals.FormattingComponent.ColorPicker,
                                    properties: {
                                        descriptor: { objectName: "labelSettings", propertyName: "fontColor" },
                                        value: { value: this.settings.labelSettings.fontColor }
                                    }
                                }
                            },
                            {
                                uid: "badgeBg_slice",
                                displayName: "Pill Badge Fill",
                                control: {
                                    type: powerbi.visuals.FormattingComponent.ColorPicker,
                                    properties: {
                                        descriptor: { objectName: "labelSettings", propertyName: "badgeBgColor" },
                                        value: { value: this.settings.labelSettings.badgeBgColor }
                                    }
                                }
                            },
                            {
                                uid: "labelPrefix_slice",
                                displayName: "Prefix",
                                control: {
                                    type: powerbi.visuals.FormattingComponent.Dropdown,
                                    properties: {
                                        descriptor: { objectName: "labelSettings", propertyName: "prefix" },
                                        value: this.settings.labelSettings.prefix
                                    }
                                }
                            },
                            {
                                uid: "labelSuffix_slice",
                                displayName: "Suffix",
                                control: {
                                    type: powerbi.visuals.FormattingComponent.Dropdown,
                                    properties: {
                                        descriptor: { objectName: "labelSettings", propertyName: "suffix" },
                                        value: this.settings.labelSettings.suffix
                                    }
                                }
                            }
                        ]
                    }
                ]
            },
            {
                displayName: "Timeline & Window Clamping",
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
                                displayName: "Clamp to Daily Window",
                                control: {
                                    type: powerbi.visuals.FormattingComponent.ToggleSwitch,
                                    properties: {
                                        descriptor: { objectName: "timeAxisSettings", propertyName: "clampToDailyWindow" },
                                        value: this.settings.timeAxisSettings.clampToDailyWindow
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
                displayName: "Reference Lines & Markers",
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
                                displayName: "Current Time Color",
                                control: {
                                    type: powerbi.visuals.FormattingComponent.ColorPicker,
                                    properties: {
                                        descriptor: { objectName: "referenceLineSettings", propertyName: "nowLineColor" },
                                        value: { value: this.settings.referenceLineSettings.nowLineColor }
                                    }
                                }
                            },
                            {
                                uid: "showMilestones_slice",
                                displayName: "Show Interval Markers",
                                control: {
                                    type: powerbi.visuals.FormattingComponent.ToggleSwitch,
                                    properties: {
                                        descriptor: { objectName: "referenceLineSettings", propertyName: "showMilestoneLines" },
                                        value: this.settings.referenceLineSettings.showMilestoneLines
                                    }
                                }
                            },
                            {
                                uid: "milestoneColor_slice",
                                displayName: "Interval Marker Color",
                                control: {
                                    type: powerbi.visuals.FormattingComponent.ColorPicker,
                                    properties: {
                                        descriptor: { objectName: "referenceLineSettings", propertyName: "milestoneLineColor" },
                                        value: { value: this.settings.referenceLineSettings.milestoneLineColor }
                                    }
                                }
                            }
                        ]
                    }
                ]
            },
            {
                displayName: "Duration & Threshold Alerts",
                uid: "thresholdAlertCard_uid",
                groups: [
                    {
                        displayName: "Alert Rules",
                        uid: "thresholdAlertGroup_uid",
                        slices: [
                            {
                                uid: "hlLong_slice",
                                displayName: "Flag Extended Durations",
                                control: {
                                    type: powerbi.visuals.FormattingComponent.ToggleSwitch,
                                    properties: {
                                        descriptor: { objectName: "thresholdAlertSettings", propertyName: "highlightLongIntervals" },
                                        value: this.settings.thresholdAlertSettings.highlightLongIntervals
                                    }
                                }
                            },
                            {
                                uid: "threshold_slice",
                                displayName: "Duration Threshold (Hours)",
                                control: {
                                    type: powerbi.visuals.FormattingComponent.NumUpDown,
                                    properties: {
                                        descriptor: { objectName: "thresholdAlertSettings", propertyName: "thresholdHours" },
                                        value: this.settings.thresholdAlertSettings.thresholdHours
                                    }
                                }
                            },
                            {
                                uid: "warnColor_slice",
                                displayName: "Alert Border Color",
                                control: {
                                    type: powerbi.visuals.FormattingComponent.ColorPicker,
                                    properties: {
                                        descriptor: { objectName: "thresholdAlertSettings", propertyName: "warningBorderColor" },
                                        value: { value: this.settings.thresholdAlertSettings.warningBorderColor }
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
                                uid: "grpSort_slice",
                                displayName: "Group Order",
                                control: {
                                    type: powerbi.visuals.FormattingComponent.Dropdown,
                                    properties: {
                                        descriptor: { objectName: "sortSettings", propertyName: "groupSort" },
                                        value: this.settings.sortSettings.groupSort
                                    }
                                }
                            },
                            {
                                uid: "resSort_slice",
                                displayName: "Resource / Item Order",
                                control: {
                                    type: powerbi.visuals.FormattingComponent.Dropdown,
                                    properties: {
                                        descriptor: { objectName: "sortSettings", propertyName: "resourceSort" },
                                        value: this.settings.sortSettings.resourceSort
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
                        resourceColumnWidth: this.settings.swimlaneSettings.resourceColumnWidth,
                        resourceFontSize: this.settings.swimlaneSettings.resourceFontSize,
                        resourceFontColor: { solid: { color: this.settings.swimlaneSettings.resourceFontColor } },
                        showZebraStriping: this.settings.swimlaneSettings.showZebraStriping,
                        evenRowColor: { solid: { color: this.settings.swimlaneSettings.evenRowColor } },
                        oddRowColor: { solid: { color: this.settings.swimlaneSettings.oddRowColor } },
                        showRowBorders: this.settings.swimlaneSettings.showRowBorders,
                        rowBorderColor: { solid: { color: this.settings.swimlaneSettings.rowBorderColor } }
                    }
                });
                break;
            case "groupHeaderSettings":
                instances.push({
                    objectName: "groupHeaderSettings",
                    selector: (null as any),
                    properties: {
                        showHeader: this.settings.groupHeaderSettings.showHeader,
                        headerHeight: this.settings.groupHeaderSettings.headerHeight,
                        headerFontSize: this.settings.groupHeaderSettings.headerFontSize,
                        headerFontColor: { solid: { color: this.settings.groupHeaderSettings.headerFontColor } },
                        headerBgColor: { solid: { color: this.settings.groupHeaderSettings.headerBgColor } },
                        showMetricBadge: this.settings.groupHeaderSettings.showMetricBadge,
                        badgeBgColor: { solid: { color: this.settings.groupHeaderSettings.badgeBgColor } },
                        badgeTextColor: { solid: { color: this.settings.groupHeaderSettings.badgeTextColor } }
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
                        showMetricNotches: this.settings.barSettings.showMetricNotches,
                        overlapMode: this.settings.barSettings.overlapMode
                    }
                });
                break;
            case "labelSettings":
                instances.push({
                    objectName: "labelSettings",
                    selector: (null as any),
                    properties: {
                        show: this.settings.labelSettings.show,
                        labelPosition: this.settings.labelSettings.labelPosition,
                        labelContent: this.settings.labelSettings.labelContent,
                        fontSize: this.settings.labelSettings.fontSize,
                        fontColor: { solid: { color: this.settings.labelSettings.fontColor } },
                        badgeBgColor: { solid: { color: this.settings.labelSettings.badgeBgColor } },
                        prefix: this.settings.labelSettings.prefix,
                        suffix: this.settings.labelSettings.suffix
                    }
                });
                break;
            case "timeAxisSettings":
                instances.push({
                    objectName: "timeAxisSettings",
                    selector: (null as any),
                    properties: {
                        timeFormat: this.settings.timeAxisSettings.timeFormat,
                        clampToDailyWindow: this.settings.timeAxisSettings.clampToDailyWindow,
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
                        showMilestoneLines: this.settings.referenceLineSettings.showMilestoneLines,
                        milestoneLineColor: { solid: { color: this.settings.referenceLineSettings.milestoneLineColor } }
                    }
                });
                break;
            case "thresholdAlertSettings":
                instances.push({
                    objectName: "thresholdAlertSettings",
                    selector: (null as any),
                    properties: {
                        highlightLongIntervals: this.settings.thresholdAlertSettings.highlightLongIntervals,
                        thresholdHours: this.settings.thresholdAlertSettings.thresholdHours,
                        warningBorderColor: { solid: { color: this.settings.thresholdAlertSettings.warningBorderColor } }
                    }
                });
                break;
            case "sortSettings":
                instances.push({
                    objectName: "sortSettings",
                    selector: (null as any),
                    properties: {
                        groupSort: this.settings.sortSettings.groupSort,
                        resourceSort: this.settings.sortSettings.resourceSort
                    }
                });
                break;
        }

        return instances;
    }
}