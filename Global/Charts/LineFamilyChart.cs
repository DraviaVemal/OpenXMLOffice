using DocumentFormat.OpenXml;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OpenXMLOffice.Global
{
    public class LineFamilyChart : ChartBase
    {
        #region Protected Fields

        protected LineChartSetting LineChartSetting;

        #endregion Protected Fields

        #region Protected Constructors

        protected LineFamilyChart(LineChartSetting LineChartSetting, ChartData[][] DataCols) : base(LineChartSetting)
        {
            this.LineChartSetting = LineChartSetting;
            SetChartPlotArea(CreateChartPlotArea(DataCols));
        }

        #endregion Protected Constructors

        #region Private Methods

        private C.PlotArea CreateChartPlotArea(ChartData[][] DataCols)
        {
            C.PlotArea plotArea = new();
            plotArea.Append(new C.Layout());
            C.LineChart LineChart = new(
                new C.Grouping
                {
                    Val = LineChartSetting.LineChartTypes switch
                    {
                        LineChartTypes.STACKED => C.GroupingValues.Stacked,
                        LineChartTypes.STACKED_MARKER => C.GroupingValues.Stacked,
                        LineChartTypes.PERCENT_STACKED => C.GroupingValues.PercentStacked,
                        LineChartTypes.PERCENT_STACKED_MARKER => C.GroupingValues.PercentStacked,
                        // Clusted
                        _ => C.GroupingValues.Standard,
                    }
                },
                new C.VaryColors { Val = false });
            int seriesIndex = 0;
            CreateDataSeries(DataCols, LineChartSetting.ChartDataSetting)
            .ForEach(Series =>
            {
                C.Marker Marker = new[] { LineChartTypes.CLUSTERED_MARKER, LineChartTypes.STACKED_MARKER, LineChartTypes.PERCENT_STACKED_MARKER }.Contains(LineChartSetting.LineChartTypes) ? new(
                    new C.Symbol { Val = C.MarkerStyleValues.Circle },
                    new C.Size { Val = 5 },
                    new C.ShapeProperties(
                        new A.SolidFill(new A.SchemeColor { Val = new A.SchemeColorValues($"accent{(seriesIndex % 6) + 1}") }),
                        new A.Outline(new A.SolidFill(new A.SchemeColor { Val = new A.SchemeColorValues($"accent{(seriesIndex % 6) + 1}") })),
                        new A.EffectList()
                    )) :
                    new(new C.Symbol()
                    {
                        Val = C.MarkerStyleValues.None
                    });
                LineChart.Append(CreateLineChartSeries(seriesIndex, Series, Marker,
                     GetSolidFill(LineChartSetting.LineChartSeriesSettings
                            .Where(item => item.FillColor != null)
                            .Select(item => item.FillColor!)
                            .ToList(), seriesIndex),
                    GetDataLabels(LineChartSetting, seriesIndex)));
                seriesIndex++;
            });
            C.DataLabels DataLabels = new(
                new C.ShowLegendKey { Val = false },
                new C.ShowValue { Val = false },
                new C.ShowCategoryName { Val = false },
                new C.ShowSeriesName { Val = false },
                new C.ShowPercent { Val = false },
                new C.ShowBubbleSize { Val = false });
            LineChart.Append(DataLabels);
            LineChart.Append(new C.Smooth { Val = false });
            LineChart.Append(new C.AxisId { Val = 1362418656 });
            LineChart.Append(new C.AxisId { Val = 1358349936 });
            plotArea.Append(LineChart);
            plotArea.Append(CreateCategoryAxis(1362418656));
            plotArea.Append(CreateValueAxis(1358349936));
            C.ShapeProperties ShapeProperties = new();
            ShapeProperties.Append(new A.NoFill());
            ShapeProperties.Append(new A.Outline(new A.NoFill()));
            ShapeProperties.Append(new A.EffectList());
            plotArea.Append(ShapeProperties);
            return plotArea;
        }

        private C.DataLabels CreateDataLabel(LineChartDataLabel LineChartDataLabel)
        {
            C.DataLabels DataLabels = new(
                new C.ShowLegendKey { Val = LineChartDataLabel.ShowLegendKey },
                new C.ShowValue { Val = LineChartDataLabel.ShowValue },
                new C.ShowCategoryName { Val = LineChartDataLabel.ShowCategoryName },
                new C.ShowSeriesName { Val = LineChartDataLabel.ShowSeriesName },
                new C.ShowPercent { Val = false },
                new C.ShowBubbleSize { Val = false },
                new C.ShowLeaderLines() { Val = false });
            DataLabels.InsertAt(new C.DataLabelPosition()
            {
                Val = LineChartDataLabel.DataLabelPosition switch
                {
                    LineChartDataLabel.eDataLabelPosition.LEFT => C.DataLabelPositionValues.Left,
                    LineChartDataLabel.eDataLabelPosition.RIGHT => C.DataLabelPositionValues.Right,
                    LineChartDataLabel.eDataLabelPosition.ABOVE => C.DataLabelPositionValues.Top,
                    LineChartDataLabel.eDataLabelPosition.BELOW => C.DataLabelPositionValues.Bottom,
                    //Center
                    _ => C.DataLabelPositionValues.Center,
                }
            }, 0);
            DataLabels.InsertAt(new C.ShapeProperties(new A.NoFill(), new A.Outline(new A.NoFill()), new A.EffectList()), 0);
            A.Paragraph Paragraph = new(new A.ParagraphProperties(new A.DefaultRunProperties(
                new A.SolidFill(new A.SchemeColor(new A.LuminanceModulation() { Val = 75000 }, new A.LuminanceOffset() { Val = 25000 }) { Val = A.SchemeColorValues.Text1 }),
                new A.LatinFont() { Typeface = "+mn-lt" }, new A.EastAsianFont() { Typeface = "+mn-ea" }, new A.ComplexScriptFont() { Typeface = "+mn-cs" })
            {
                FontSize = 1197,
                Bold = false,
                Italic = false,
                Underline = A.TextUnderlineValues.None,
                Strike = A.TextStrikeValues.NoStrike,
                Kerning = 1200,
                Baseline = 0
            }), new A.EndParagraphRunProperties() { Language = "en-US" });
            DataLabels.InsertAt(new C.TextProperties(new A.BodyProperties(new A.ShapeAutoFit())
            {
                Rotation = 0,
                UseParagraphSpacing = true,
                VerticalOverflow = A.TextVerticalOverflowValues.Ellipsis,
                Vertical = A.TextVerticalValues.Horizontal,
                Wrap = A.TextWrappingValues.Square,
                LeftInset = 38100,
                TopInset = 19050,
                RightInset = 38100,
                BottomInset = 19050,
                Anchor = A.TextAnchoringTypeValues.Center,
                AnchorCenter = true
            }, new A.ListStyle(),
           Paragraph), 0);
            return DataLabels;
        }

        private C.LineChartSeries CreateLineChartSeries(int seriesIndex, ChartDataGrouping ChartDataGrouping, C.Marker Marker, A.SolidFill SolidFill, C.DataLabels DataLabels)
        {
            C.LineChartSeries series = new(
                new C.Index { Val = new UInt32Value((uint)seriesIndex) },
                new C.Order { Val = new UInt32Value((uint)seriesIndex) },
                new C.SeriesText(new C.StringReference(new C.Formula(ChartDataGrouping.SeriesHeaderFormula!), AddStringCacheValue(new[] { ChartDataGrouping.SeriesHeaderCells! }))),
                Marker);
            C.ShapeProperties ShapeProperties = new();
            ShapeProperties.Append(new A.Outline(SolidFill, new A.Round()));
            ShapeProperties.Append(new A.EffectList());
            series.Append(DataLabels);
            series.Append(ShapeProperties);
            series.Append(new C.CategoryAxisData(new C.StringReference(new C.Formula(ChartDataGrouping.XaxisFormula!), AddStringCacheValue(ChartDataGrouping.XaxisCells!))));
            series.Append(new C.Values(new C.NumberReference(new C.Formula(ChartDataGrouping.YaxisFormula!), AddNumberCacheValue(ChartDataGrouping.YaxisCells!, null))));
            series.Append(new C.Smooth() { Val = false });
            return series;
        }

        private C.DataLabels GetDataLabels(LineChartSetting LineChartSetting, int index)
        {
            if (index < LineChartSetting.LineChartSeriesSettings.Count)
            {
                return CreateDataLabel(LineChartSetting.LineChartSeriesSettings?[index]?.LineChartDataLabel ?? new LineChartDataLabel());
            }
            return CreateDataLabel(new LineChartDataLabel());
        }

        #endregion Private Methods
    }
}