﻿module powerbi.visuals.experimental {

    export module ScatterPrototype {
        export interface ScatterDataPointModel {
            x: number;
            y: number;
            size: number;
        }

        export interface ScatterDataPointViewModel {
            x: number;
            y: number;
            radius: number;
            fill: string;
        }

        export interface ScatterViewModel {
            dataPoints: ScatterDataPointViewModel[];
            boundingBox: BoundingBox;
        }

        export interface ScatterDataModel {
            axes: any[];
            legend: any;
            dataPoints: ScatterDataPointModel[];
            //scatterPlotArea: ScatterPlotAreaViewModel;
        }

        export class ScatterVisual implements IVisual {
            private element: JQuery;
            private svg: D3.Selection;
            private initOptions: VisualInitOptions;

            private viewport: IViewport;
            
            public init(options: VisualInitOptions) {
                this.initOptions = options;

                let element = this.element = options.element;

                this.svg = d3.select(element.get(0)).append('svg');
            }

            //private static forEachDataPoint(callback: () => void) {
            //    for (let categoryIdx = 0, ilen = categoryValues.length; categoryIdx < ilen; categoryIdx++) {
            //        let categoryValue = categoryValues[categoryIdx];

            //        for (let seriesIdx = 0, len = grouped.length; seriesIdx < len; seriesIdx++) {
            //            let grouping = grouped[seriesIdx];
            //        }
            //    }
            //}

            public layout(model: ScatterDataModel, boundingBox: BoundingBox): ScatterViewModel {

                let dataPoints: ScatterDataPointViewModel[] = [];

                for (let i = 0; i < model.dataPoints.length; i++) {
                    //let dataPoint = model.dataPoints[i];

                    //let colorHelper = new ColorHelper(colorPalette, scatterChartProps.dataPoint.fill, defaultDataPointColor);

                    //let color: string;
                    //if (hasDynamicSeries) {
                    //    color = colorHelper.getColorForSeriesValue(grouping.objects, dataValues.identityFields, grouping.name);
                    //}
                    //else {
                    //    // If we have no Size measure then use a blank query name
                    //    let measureSource = (measureSize != null)
                    //        ? measureSize.source.queryName
                    //        : '';

                    //    color = colorHelper.getColorForMeasure(categoryObjects && categoryObjects[categoryIdx], measureSource);
                    //}

                    dataPoints.push(<ScatterDataPointViewModel> {
                        x: (Math.random() * boundingBox.width) + boundingBox.left,
                        y: (Math.random() * boundingBox.height) + boundingBox.top,
                        fill: "red",
                        radius: (Math.random() * 22.123),
                    });
                }

                return {
                    dataPoints: dataPoints,
                    boundingBox: boundingBox,
                };
            }

            public convert(dataView: DataView, colorPalette: IDataColorPalette, defaultDataPointColor?: string): ScatterDataModel {
                return new ScatterChartDataConverter(dataView).convert(colorPalette, defaultDataPointColor);
            }

            public render(options: RenderOptions<ScatterViewModel>) {
                let plot = options.drawingSurface.select('.plot');

                if (plot.size() === 0) {
                    plot = options.drawingSurface.append('g').classed('plot', true);
                }

                let bbox = options.viewModel.boundingBox;
                plot.style('transform', SVGUtil.translate(bbox.left, bbox.top));

                let selection = plot.selectAll('.dataPoint').data(options.viewModel.dataPoints);
                selection.enter()
                    .append('circle')
                    .classed('dataPoint', true);

                selection.attr({
                    'cx': (d: ScatterDataPointViewModel) => d.x,
                    'cy': (d: ScatterDataPointViewModel) => d.y,
                    'fill': (d: ScatterDataPointViewModel) => d.fill,
                    'r': (d: ScatterDataPointViewModel) => d.radius,
                });

                selection.exit();

                DebugHelper.drawRect(options.drawingSurface, bbox, "green", "plot");
            }

            //public update(options: VisualUpdateOptions) {
            //    let dataView = options.dataViews[0];

            //    let plotAreaDataModel = PerfectScatter.converter(dataView);
            //    //TODO: fix null params
            //    let legendDataModel = Legend.converter(dataView.categorical.values, null, "", null);
            //    let axesDataModels = 

            //    let plotAreaBoundingBox = PerfectScatter.getPreferredLayout(plotAreaDataModel);
            //    let legendBoundingBox = Legend.getPreferredLayout(legendDataModel);
            //    //...

            //    this.reconcileLayout(/*...*/);

            //    let legendViewModel = Legend.layout(legenDataModel, finalLegendBoundingBox);
            //    let plotAreaViewModel = PerfectScatter.layout(plotAreaDataModel, finalPlotAreaBoundingBox);

            //    Legend.render(legendViewModel);
            //    //AxesHelper.render(axes[]);
            //    PerfectScatter.render({viewModel: viewModel});
            //}

            public update(options: VisualUpdateOptions) {
                let dataView = options.dataViews[0];

                this.viewport = options.viewport;

                // --- Create data models ---
                let scatterDataModel = this.convert(dataView, this.initOptions.style.colorPalette.dataColors);
                //TODO: fix null params
                let legend = new Legend(dataView, null, "", null);
                let axes = new CartesianAxes(dataView);

                // --- Layout ---
                let visualBoundingBox: BoundingBox = {
                    top: 0,
                    left: 0,
                    height: this.viewport.height,
                    width: this.viewport.width,
                };
                let layoutManager = new DockLayoutManager(visualBoundingBox);

                // --- Legend ---
                let legendPosition = DockPosition.Left;  // TODO: get position from legend?
                let legendBoundingBox = layoutManager.measure(legend, legendPosition);
                let axesBoundingBox = layoutManager.measure(axes, DockPosition.Fill);

                // TODO:
                //let plotBoundingBox = axes.getPlotArea();

                // Get view models
                let viewModel = this.layout(scatterDataModel, visualBoundingBox);
                let legendViewModel = legend.layout(legendBoundingBox);
                let axesViewModels = axes.layout(axesBoundingBox);

                // Render view models
                let svg = this.svg;
                svg.attr({
                    width: this.viewport.width,
                    height: this.viewport.height,
                });

                this.render({ viewModel: viewModel, drawingSurface: svg });
                legend.render({ viewModel: legendViewModel, drawingSurface: svg });
                axes.render({ viewModel: axesViewModels, drawingSurface: svg });
            }

            public onResizing(viewport: IViewport) { /*NOT NEEDED*/ }

            public onDataChanged(options: VisualDataChangedOptions) {
                this.update(<VisualUpdateOptions> {
                    dataViews: options.dataViews,
                    suppressAnimations: options.suppressAnimations,
                    viewport: this.viewport,
                });
            }

            public destroy() { }
        }

        class ScatterMetadata {
            public xIndex: number;
            public yIndex: number;
            public sizeIndex: number;

            constructor(grouped: DataViewValueColumnGroup[]) {
                let xIndex = DataRoleHelper.getMeasureIndexOfRole(grouped, 'X');
                let yIndex = DataRoleHelper.getMeasureIndexOfRole(grouped, 'Y');
                let sizeIndex = DataRoleHelper.getMeasureIndexOfRole(grouped, 'Size');

                let xCol: DataViewMetadataColumn;
                let yCol: DataViewMetadataColumn;
                let sizeCol: DataViewMetadataColumn;
                let xAxisLabel = "";
                let yAxisLabel = "";

                if (grouped && grouped.length) {
                    let firstGroup = grouped[0],
                        measureCount = firstGroup.values.length;

                    if (!(xIndex >= 0))
                        xIndex = ScatterMetadata.getDefaultMeasureIndex(measureCount, yIndex, sizeIndex);
                    if (!(yIndex >= 0))
                        yIndex = ScatterMetadata.getDefaultMeasureIndex(measureCount, xIndex, sizeIndex);
                    if (!(sizeIndex >= 0))
                        sizeIndex = ScatterMetadata.getDefaultMeasureIndex(measureCount, xIndex, yIndex);

                    if (xIndex >= 0) {
                        xCol = firstGroup.values[xIndex].source;
                        xAxisLabel = firstGroup.values[xIndex].source.displayName;
                    }
                    if (yIndex >= 0) {
                        yCol = firstGroup.values[yIndex].source;
                        yAxisLabel = firstGroup.values[yIndex].source.displayName;
                    }
                    if (sizeIndex >= 0) {
                        sizeCol = firstGroup.values[sizeIndex].source;
                    }
                }

                this.xIndex = xIndex;
                this.yIndex = yIndex;
                this.sizeIndex = sizeIndex;
            }

            private static getDefaultMeasureIndex(count: number, usedIndex: number, usedIndex2: number): number {
                for (let i = 0; i < count; i++) {
                    if (i !== usedIndex && i !== usedIndex2)
                        return i;
                }
            }
        }

        class ScatterChartDataConverter {
            private dataView: DataView;
            private dataViewHelper: DataViewHelper;

            constructor(dataView: DataView) {
                this.dataView = dataView;
                this.dataViewHelper = new DataViewHelper(dataView);
            }

            public convert(colorPalette: IDataColorPalette, defaultDataPointColor?: string): ScatterDataModel {
                let categoryValues = this.dataViewHelper.categoryValues;
                let categoryObjects = this.dataViewHelper.categoryObjects;
                let seriesColumns = this.dataViewHelper.seriesColumns;

                let scatterMetadata = new ScatterMetadata(seriesColumns);

                let dataPoints: ScatterDataPointModel[] = [];

                let colorHelper = new ColorHelper(colorPalette, scatterChartProps.dataPoint.fill, defaultDataPointColor);
				
                let categoryFormatter = categoryValues.length > 0
                    ? valueFormatter.create({ format: valueFormatter.getFormatString(this.dataViewHelper.categoryColumn.source, scatterChartProps.general.formatString), value: categoryValues[0], value2: categoryValues[categoryValues.length - 1] })
                    : valueFormatter.createDefaultFormatter(null);

                for (let categoryIdx = 0, catLen = categoryValues.length; categoryIdx < catLen; categoryIdx++) {
                    let categoryValue = categoryValues[categoryIdx];

                    for (let seriesIdx = 0, seriesLen = seriesColumns.length; seriesIdx < seriesLen; seriesIdx++) {
                        let seriesColumn = seriesColumns[seriesIdx];
                        let seriesValues = seriesColumn.values;

                        let measureX = ScatterChartDataConverter.getMeasureColumn(scatterMetadata.xIndex, seriesValues);
                        let measureY = ScatterChartDataConverter.getMeasureColumn(scatterMetadata.yIndex, seriesValues);
                        let measureSize = ScatterChartDataConverter.getMeasureColumn(scatterMetadata.sizeIndex, seriesValues);

                        let xVal = ScatterChartDataConverter.getMeasureValue(measureX, categoryIdx);
                        let yVal = ScatterChartDataConverter.getMeasureValue(measureY, categoryIdx, 0);
                        let size = ScatterChartDataConverter.getMeasureValue(measureSize, categoryIdx);

                        let hasNullValue = (xVal == null) || (yVal == null);
                        if (hasNullValue)
                            continue;

                        let color: string;
                        if (this.dataViewHelper.hasDynamicSeries) {
                            color = colorHelper.getColorForSeriesValue(seriesColumn.objects, this.dataViewHelper.seriesIdentity, seriesColumn.name);
                        }
                        else {
                            // If we have no Size measure then use a blank query name
                            let measureSource = (measureSize != null)
                                ? measureSize.source.queryName
                                : '';

                            color = colorHelper.getColorForMeasure(categoryObjects && categoryObjects[categoryIdx], measureSource);
                        }
						
                        let identity = SelectionIdBuilder.builder()
                            .withCategory(this.dataViewHelper.categoryColumn, categoryIdx)
                            .withSeries(this.dataView.categorical.values, seriesColumn)
                            .createSelectionId();

						// TODO: tooltips

                        let dataPoint: ScatterDataPointModel = {
                            x: xVal,
                            y: yVal,
                            //radius: { sizeMeasure: measureSize, index: categoryIdx },
                            size: size,
                            fill: color,
                            category: categoryFormatter.format(categoryValue),
                            selected: false,
                            identity: identity,
                            //tooltipInfo: tooltipInfo,
                            //labelFill: labelSettings.labelColor,
                        };

                        dataPoints.push(dataPoint);
                    }
                }

                return <ScatterDataModel> {
                    dataPoints: dataPoints,
                };
            }

            private static getMeasureColumn(measureIndex: number, seriesValues: DataViewValueColumn[]): DataViewValueColumn {
                if (measureIndex >= 0)
                    return seriesValues[measureIndex];

                return null;
            }

            private static getMeasureValue(measureColumn: DataViewValueColumn, index: number, defaultValue: any = null): any {
                return measureColumn && measureColumn.values ? measureColumn.values[index] : defaultValue;
            }
        }

        module BubbleHelpers {
            let AreaOf300By300Chart = 90000;
            let MinSizeRange = 200;
            let MaxSizeRange = 3000;

            export function getBubbleRadius(value: number, valueRange: NumberRange, viewport: IViewport): number {
                let min = valueRange.min || 0;
                min = Math.min(min, 0);

                let max = valueRange.max || 0;
                max = Math.max(max, 0);

                // TODO: calculate this in converter?
                let valueDataRange = {
                    minRange: min,
                    maxRange: max,
                    delta: max - min
                };

                // TODO: calculate this only once per update
                let bubblePixelAreaSizeRange = getBubblePixelAreaSizeRange(viewport, MinSizeRange, MaxSizeRange);

                return projectSizeToPixels(value, valueDataRange, bubblePixelAreaSizeRange) / 2;
            }

            function getBubblePixelAreaSizeRange(viewport: IViewport, minSizeRange: number, maxSizeRange: number): DataRange {
                let ratio = 1.0;
                if (viewport.height > 0 && viewport.width > 0) {
                    let minDimension = Math.min(viewport.height, viewport.width);
                    ratio = (minDimension * minDimension) / AreaOf300By300Chart;
                }

                let minRange = Math.round(minSizeRange * ratio);
                let maxRange = Math.round(maxSizeRange * ratio);
                return {
                    minRange: minRange,
                    maxRange: maxRange,
                    delta: maxRange - minRange
                };
            }

            function projectSizeToPixels(size: number, actualSizeDataRange: DataRange, bubblePixelAreaSizeRange: DataRange): number {
                // Project value on the required range of bubble area sizes
                let projectedSize = bubblePixelAreaSizeRange.maxRange;
                if (actualSizeDataRange.delta !== 0) {
                    let value = Math.min(Math.max(size, actualSizeDataRange.minRange), actualSizeDataRange.maxRange);
                    projectedSize = project(value, actualSizeDataRange, bubblePixelAreaSizeRange);
                }

                projectedSize = Math.sqrt(projectedSize / Math.PI) * 2;

                return Math.round(projectedSize);
            }

            function project(value: number, actualSizeDataRange: DataRange, bubblePixelAreaSizeRange: DataRange): number {
                if (actualSizeDataRange.delta === 0 || bubblePixelAreaSizeRange.delta === 0) {
                    return (rangeContains(actualSizeDataRange, value)) ? bubblePixelAreaSizeRange.minRange : null;
                }

                let relativeX = (value - actualSizeDataRange.minRange) / actualSizeDataRange.delta;
                return bubblePixelAreaSizeRange.minRange + relativeX * bubblePixelAreaSizeRange.delta;
            }

            function rangeContains(range: DataRange, value: number): boolean {
                return range.minRange <= value && value <= range.maxRange;
            }
        }
    }

    export module DebugHelper {
        export function drawRect(drawingSurface: D3.Selection, bbox: BoundingBox, color: string, tag: string): void {
            let placeholder = drawingSurface.select('.debug' + tag);

            if(placeholder.size() === 0) {
                placeholder = drawingSurface.append('rect').classed('debug' + tag, true);
            }

            placeholder.style('transform', SVGUtil.translate(bbox.left, bbox.top));
            placeholder.attr({
                width: bbox.width,
                height: bbox.height,
                "stroke-width": "1px",
                stroke: color,
                fill: "none",
            });
        }
    }
}