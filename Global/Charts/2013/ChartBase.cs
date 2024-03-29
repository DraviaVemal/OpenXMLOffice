// Copyright (c) DraviaVemal. Licensed under the MIT License. See License in the project root.

using DocumentFormat.OpenXml;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using C15 = DocumentFormat.OpenXml.Office2013.Drawing.Chart;

namespace OpenXMLOffice.Global_2013;

/// <summary>
/// Chart Base Class Common to all charts. Class is only intended to get created by inherited classes
/// </summary>
public class ChartBase : CommonProperties
{
	internal const int AccentColurCount = 6;
	internal uint CategoryAxisId = 1362418656;
	internal uint ValueAxisId = 1358349936;
	internal const int SecondaryCategoryAxisId = 1615085760;
	internal const int SecondaryValueAxisId = 1474633616;

	/// <summary>
	/// Chart Data Groupings
	/// </summary>
	internal List<ChartDataGrouping> chartDataGroupings = new();

	/// <summary>
	/// Core chart settings common for every possible chart
	/// </summary>
	internal ChartSetting chartSetting;

	private readonly C.Chart chart;

	private readonly C.ChartSpace openXMLChartSpace;

	/// <summary>
	/// Chartbase class constructor restricted only for inheritance use
	/// </summary>
	/// <param name="chartSetting">
	/// </param>
	internal ChartBase(ChartSetting chartSetting)
	{
		CategoryAxisId = chartSetting.categoryAxisId ?? CategoryAxisId;
		ValueAxisId = chartSetting.valueAxisId ?? ValueAxisId;
		this.chartSetting = chartSetting;
		openXMLChartSpace = CreateChartSpace();
		chart = CreateChart();
		GetChartSpace().Append(chart);
		GetChartSpace().Append(new C.ExternalData(
			new C.AutoUpdate() { Val = false })
		{ Id = "rId1" });
	}

	/// <summary>
	/// Get OpenXML ChartSpace
	/// </summary>
	public virtual C.ChartSpace GetChartSpace()
	{
		return openXMLChartSpace;
	}

	/// <summary>
	/// Create Bubble Size Axis for the chart
	/// </summary>
	internal static C.BubbleSize CreateBubbleSizeAxisData(string formula, ChartData[] cells)
	{
		if (cells.All(v => v.dataType != DataType.NUMBER))
		{
			Console.WriteLine($"Object Details Value : {cells.FirstOrDefault(v => v.dataType != DataType.NUMBER)?.value} is not numaric");
			Console.WriteLine($"Object Details Number Format : {cells.FirstOrDefault(v => v.dataType != DataType.NUMBER)?.numberFormat}");
			Console.WriteLine($"Object Details Data Type : {cells.FirstOrDefault(v => v.dataType != DataType.NUMBER)?.dataType}");
			throw new ArgumentException($"Bubble Size Data Should Be numaric.");
		}
		return new(new C.NumberReference(new C.Formula(formula), AddNumberCacheValue(cells)));
	}

	/// <summary>
	/// Create Category Axis for the chart
	/// </summary>
	/// <param name="categoryAxisSetting">
	/// </param>
	/// <returns>
	/// </returns>
	internal C.CategoryAxis CreateCategoryAxis(CategoryAxisSetting categoryAxisSetting)
	{
		C.CategoryAxis CategoryAxis = new(
			new C.AxisId { Val = categoryAxisSetting.id },
			new C.Scaling(new C.Orientation { Val = categoryAxisSetting.invertOrder ? C.OrientationValues.MaxMin : C.OrientationValues.MinMax }),
			new C.Delete { Val = !categoryAxisSetting.isVisible },
			new C.AxisPosition
			{
				Val = categoryAxisSetting.axisPosition switch
				{
					AxisPosition.LEFT => C.AxisPositionValues.Left,
					AxisPosition.RIGHT => C.AxisPositionValues.Right,
					AxisPosition.TOP => C.AxisPositionValues.Top,
					_ => C.AxisPositionValues.Bottom
				}
			},
			new C.MajorTickMark { Val = C.TickMarkValues.None },
			new C.MinorTickMark { Val = C.TickMarkValues.None },
			new C.TickLabelPosition { Val = C.TickLabelPositionValues.NextTo });
		if (categoryAxisSetting.isVisible)
		{
			if (chartSetting.chartGridLinesOptions.isMajorCategoryLinesEnabled)
			{
				CategoryAxis.Append(CreateMajorGridLine());
			}
			if (chartSetting.chartGridLinesOptions.isMinorCategoryLinesEnabled)
			{
				CategoryAxis.Append(CreateMinorGridLine());
			}
			CategoryAxis.Append(CreateChartShapeProperties());
			SolidFillModel solidFillModel = new()
			{
				schemeColorModel = new()
				{
					themeColorValues = ThemeColorValues.TEXT_1,
					luminanceModulation = 65000,
					luminanceOffset = 35000
				}
			};
			if (categoryAxisSetting.fontColor != null)
			{
				solidFillModel.hexColor = categoryAxisSetting.fontColor;
				solidFillModel.schemeColorModel = null;
			}
			CategoryAxis.Append(CreateChartTextProperties(new()
			{
				drawingBodyProperties = new(),
				drawingParagraph = new()
				{
					paragraphPropertiesModel = new()
					{
						defaultRunProperties = new()
						{
							solidFill = solidFillModel,
							fontSize = ConverterUtils.FontSizeToFontSize(categoryAxisSetting.fontSize),
							isBold = categoryAxisSetting.isBold,
							isItalic = categoryAxisSetting.isItalic,
							underline = categoryAxisSetting.underLineValues,
							strike = categoryAxisSetting.strikeValues,
							baseline = 0,
						}
					}
				}
			}));
		}
		CategoryAxis.Append(
			new C.CrossingAxis { Val = categoryAxisSetting.crossAxisId },
			new C.Crosses { Val = C.CrossesValues.AutoZero },
			new C.AutoLabeled { Val = true },
			new C.LabelAlignment { Val = C.LabelAlignmentValues.Center },
			new C.LabelOffset { Val = 100 },
			new C.NoMultiLevelLabels { Val = false });
		return CategoryAxis;
	}

	/// <summary>
	/// Create Chart Shape Properties for the chart
	/// </summary>
	internal static C.Layout CreateLayout(LayoutModel? layoutModel = null)
	{
		if (layoutModel == null)
		{
			return new();
		}
		double x = layoutModel.x;
		double y = layoutModel.y;
		double width = layoutModel.width;
		double height = layoutModel.height;
		if (x < 0 || x > 1 || width < 0 || width > 1 || x + width < 0 || x + width > 1)
		{
			throw new ArgumentException("Layout value is not within acceptable range. X and Width values should be between 0 and 1, and their sum should be between 0 and 1.");
		}

		if (y < 0 || y > 1 || height < 0 || height > 1 || y + height < 0 || y + height > 1)
		{
			throw new ArgumentException("Layout value is not within acceptable range. Y and Height values should be between 0 and 1, and their sum should be between 0 and 1.");
		}

		return new(
			new C.ManualLayout(
				new C.LayoutTarget { Val = C.LayoutTargetValues.Inner },
				new C.LeftMode { Val = C.LayoutModeValues.Edge },
				new C.TopMode { Val = C.LayoutModeValues.Edge },
				new C.Left { Val = x },
				new C.Top { Val = y },
				new C.Width { Val = width },
				new C.Height { Val = height }
			));
	}


	/// <summary>
	/// Create Category Axis Data for the chart
	/// </summary>
	/// <param name="formula">
	/// </param>
	/// <param name="cells">
	/// </param>
	/// <returns>
	/// </returns>
	internal static C.CategoryAxisData CreateCategoryAxisData(string formula, ChartData[] cells)
	{
		if (cells.All(v => v.dataType == DataType.NUMBER))
		{
			return new(new C.NumberReference(new C.Formula(formula), AddNumberCacheValue(cells)));
		}
		else
		{
			return new(new C.StringReference(new C.Formula(formula), AddStringCacheValue(cells)));
		}
	}

	/// <summary>
	/// Create Data Labels for the chart
	/// </summary>
	internal C.DataLabels CreateDataLabels(ChartDataLabel chartDataLabel, int? dataLabelCount = 0)
	{
		C.Extension extension = new(
				new C15.ShowDataLabelsRange() { Val = chartDataLabel.showValueFromColumn },
				new C15.ShowLeaderLines() { Val = false }
			)
		{ Uri = "{CE6537A1-D6FC-4f65-9D91-7224C49458BB}" };
		if (chartDataLabel.showValueFromColumn)
		{
			extension.InsertAt(new C15.DataLabelFieldTable(), 0);
		}
		C.ExtensionList extensionList = new(extension);
		C.DataLabels dataLabels = new();
		if (chartDataLabel.showValueFromColumn)
		{
			for (int i = 0; i < dataLabelCount; i++)
			{
				A.Paragraph Paragraph = new(CreateField("CELLRANGE", "[CELLRANGE]"));
				if (chartDataLabel.showSeriesName)
				{
					Paragraph.Append(CreateDrawingRun(new() { text = chartDataLabel.separator }));
					Paragraph.Append(CreateField("SERIESNAME", "[SERIES NAME]"));
				}
				if (chartDataLabel.showCategoryName)
				{
					Paragraph.Append(CreateDrawingRun(new() { text = chartDataLabel.separator }));
					Paragraph.Append(CreateField("CATEGORYNAME", "[CATEGORY NAME]"));
				}
				if (chartDataLabel.showValue)
				{
					Paragraph.Append(CreateDrawingRun(new() { text = chartDataLabel.separator }));
					Paragraph.Append(CreateField("VALUE", "[VALUE]"));
				}
				Paragraph.Append(new A.EndParagraphRunProperties { Language = "en-IN" });
				dataLabels.Append(new C.DataLabel(
					new C.Index() { Val = (uint)i },
					new C.SeriesText(
						new C.RichText(
							new A.BodyProperties(),
							new A.ListStyle(),
							Paragraph
						)
					),
					new C.ShowLegendKey { Val = chartDataLabel.showLegendKey },
					new C.ShowValue { Val = chartDataLabel.showValue },
					new C.ShowCategoryName { Val = chartDataLabel.showCategoryName },
					new C.ShowSeriesName { Val = chartDataLabel.showSeriesName },
					new C.ShowPercent() { Val = true },
					new C.ShowBubbleSize() { Val = true },
					new C.Separator(chartDataLabel.separator),
					(OpenXmlElement)extensionList.Clone()
				));
			}
		}
		dataLabels.Append(new C.ShowLegendKey { Val = chartDataLabel.showLegendKey },
			new C.ShowValue { Val = chartDataLabel.showValue },
			new C.ShowCategoryName { Val = chartDataLabel.showCategoryName },
			new C.ShowSeriesName { Val = chartDataLabel.showSeriesName },
			new C.ShowPercent { Val = false },
			new C.ShowBubbleSize() { Val = true },
			new C.Separator(chartDataLabel.separator),
			new C.ShowLeaderLines() { Val = false },
			(OpenXmlElement)extensionList.Clone());
		dataLabels.Append(CreateChartShapeProperties());
		SolidFillModel solidFillModel = new()
		{
			schemeColorModel = new()
			{
				themeColorValues = ThemeColorValues.TEXT_1,
				luminanceModulation = 65000,
				luminanceOffset = 35000
			}
		};
		if (chartDataLabel.fontColor != null)
		{
			solidFillModel.hexColor = chartDataLabel.fontColor;
			solidFillModel.schemeColorModel = null;
		}
		dataLabels.Append(CreateChartTextProperties(new()
		{
			drawingBodyProperties = new()
			{
				rotation = 0,
				anchorCenter = true,
				anchor = TextAnchoringValues.CENTER,
				bottomInset = 19050,
				leftInset = 38100,
				rightInset = 38100,
				topInset = 19050,
				useParagraphSpacing = true,
				vertical = TextVerticalAlignmentValues.HORIZONTAL,
				verticalOverflow = TextVerticalOverflowValues.ELLIPSIS,
				wrap = TextWrappingValues.SQUARE,
			},
			drawingParagraph = new()
			{
				paragraphPropertiesModel = new()
				{
					defaultRunProperties = new()
					{
						solidFill = solidFillModel,
						complexScriptFont = "+mn-cs",
						eastAsianFont = "+mn-ea",
						latinFont = "+mn-lt",
						fontSize = ConverterUtils.FontSizeToFontSize(chartDataLabel.fontSize),
						isBold = chartDataLabel.isBold,
						isItalic = chartDataLabel.isItalic,
						underline = chartDataLabel.underLineValues,
						strike = chartDataLabel.strikeValues,
						kerning = 1200,
						baseline = 0,
					}
				}
			}
		}));
		return dataLabels;
	}

	/// <summary>
	/// Create Data Labels Range for the chart.Used in value from Column
	/// </summary>
	internal static C15.DataLabelsRange CreateDataLabelsRange(string formula, ChartData[] cells)
	{
		return new(new C15.Formula(formula), AddDataLabelCacheValue(cells));
	}

	/// <summary>
	/// Create Data Series for the chart
	/// </summary>
	internal List<ChartDataGrouping> CreateDataSeries(ChartData[][] dataCols, ChartDataSetting chartDataSetting)
	{
		List<uint> seriesColumns = new();
		for (uint col = chartDataSetting.chartDataColumnStart + 1; col <= (chartDataSetting.chartDataColumnEnd == 0 ? dataCols.Length - 1 : chartDataSetting.chartDataColumnEnd); col++)
		{
			seriesColumns.Add(col);
		}
		if ((chartDataSetting.chartDataRowEnd == 0 ? dataCols[0].Length : chartDataSetting.chartDataRowEnd) - chartDataSetting.chartDataRowStart < 1 || (chartDataSetting.chartDataColumnEnd == 0 ? dataCols.Length : chartDataSetting.chartDataColumnEnd) - chartDataSetting.chartDataColumnStart < 1)
		{
			throw new ArgumentException("Data Series Invalid Range");
		}
		for (int i = 0; i < seriesColumns.Count; i++)
		{
			uint column = seriesColumns[i];
			List<ChartData> xAxisCells = ((ChartData[]?)dataCols[chartDataSetting.chartDataColumnStart].Clone()!).Skip((int)chartDataSetting.chartDataRowStart + 1).Take((chartDataSetting.chartDataRowEnd == 0 ? dataCols[0].Length : (int)chartDataSetting.chartDataRowEnd) - (int)chartDataSetting.chartDataRowStart).ToList();
			List<ChartData> yAxisCells = ((ChartData[]?)dataCols[column].Clone()!).Skip((int)chartDataSetting.chartDataRowStart + 1).Take((chartDataSetting.chartDataRowEnd == 0 ? dataCols[0].Length : (int)chartDataSetting.chartDataRowEnd) - (int)chartDataSetting.chartDataRowStart).ToList();
			ChartDataGrouping chartDataGrouping = new()
			{
				id = i,
				seriesHeaderFormula = $"Sheet1!${ConverterUtils.ConvertIntToColumnName((int)column + 1)}${chartDataSetting.chartDataRowStart + 1}",
				seriesHeaderCells = ((ChartData[]?)dataCols[column].Clone()!)[chartDataSetting.chartDataRowStart],
				xAxisFormula = $"Sheet1!${ConverterUtils.ConvertIntToColumnName((int)chartDataSetting.chartDataColumnStart + 1)}${chartDataSetting.chartDataRowStart + 2}:${ConverterUtils.ConvertIntToColumnName((int)chartDataSetting.chartDataColumnStart + 1)}${chartDataSetting.chartDataRowStart + xAxisCells.Count + 1}",
				xAxisCells = xAxisCells.ToArray(),
				yAxisFormula = $"Sheet1!${ConverterUtils.ConvertIntToColumnName((int)column + 1)}${chartDataSetting.chartDataRowStart + 2}:${ConverterUtils.ConvertIntToColumnName((int)column + 1)}${chartDataSetting.chartDataRowStart + yAxisCells.Count + 1}",
				yAxisCells = yAxisCells.ToArray(),
			};
			if (chartDataSetting.is3Ddata)
			{
				i++;
				column = seriesColumns[i];
				List<ChartData> zAxisCells = ((ChartData[]?)dataCols[column].Clone()!).Skip((int)chartDataSetting.chartDataRowStart + 1).Take((chartDataSetting.chartDataRowEnd == 0 ? dataCols[0].Length : (int)chartDataSetting.chartDataRowEnd) - (int)chartDataSetting.chartDataRowStart).ToList();
				chartDataGrouping.zAxisFormula = $"Sheet1!${ConverterUtils.ConvertIntToColumnName((int)column + 1)}${chartDataSetting.chartDataRowStart + 2}:${ConverterUtils.ConvertIntToColumnName((int)column + 1)}${chartDataSetting.chartDataRowStart + zAxisCells.Count + 1}";
				chartDataGrouping.zAxisCells = zAxisCells.ToArray();
			}
			if (chartDataSetting.valueFromColumn.TryGetValue(column, out uint DataValueColumn))
			{
				List<ChartData> dataLabelCells = ((ChartData[]?)dataCols[DataValueColumn].Clone()!).Skip((int)chartDataSetting.chartDataRowStart).Take((chartDataSetting.chartDataRowEnd == 0 ? dataCols[0].Length : (int)chartDataSetting.chartDataRowEnd) - (int)chartDataSetting.chartDataRowStart).ToList();
				chartDataGrouping.dataLabelFormula = $"Sheet1!${ConverterUtils.ConvertIntToColumnName((int)DataValueColumn + 1)}${chartDataSetting.chartDataRowStart + 2}:${ConverterUtils.ConvertIntToColumnName((int)DataValueColumn + 1)}${chartDataSetting.chartDataRowStart + dataLabelCells.Count}";
				chartDataGrouping.dataLabelCells = dataLabelCells.ToArray();
			}
			chartDataGroupings.Add(chartDataGrouping);
		}
		return chartDataGroupings;
	}

	/// <summary>
	/// Create Series Text for the chart
	/// </summary>
	internal static C.SeriesText CreateSeriesText(string formula, ChartData[] cells)
	{
		return new(new C.StringReference(new C.Formula(formula), AddStringCacheValue(cells)));
	}

	/// <summary>
	/// Create Value Axis for the chart
	/// </summary>
	internal C.ValueAxis CreateValueAxis(ValueAxisSetting valueAxisSetting)
	{
		C.ValueAxis valueAxis = new(
			new C.AxisId { Val = valueAxisSetting.id },
			new C.Scaling(new C.Orientation { Val = valueAxisSetting.invertOrder ? C.OrientationValues.MaxMin : C.OrientationValues.MinMax }),
			new C.Delete { Val = !valueAxisSetting.isVisible },
			new C.AxisPosition
			{
				Val = valueAxisSetting.axisPosition switch
				{
					AxisPosition.LEFT => C.AxisPositionValues.Left,
					AxisPosition.RIGHT => C.AxisPositionValues.Right,
					AxisPosition.TOP => C.AxisPositionValues.Top,
					_ => C.AxisPositionValues.Bottom
				}
			});
		if (chartSetting.chartGridLinesOptions.isMajorValueLinesEnabled)
		{
			valueAxis.Append(CreateMajorGridLine());
		}
		if (chartSetting.chartGridLinesOptions.isMinorValueLinesEnabled)
		{
			valueAxis.Append(CreateMinorGridLine());
		}
		valueAxis.Append(
			new C.NumberingFormat { FormatCode = "General", SourceLinked = true },
			new C.MajorTickMark { Val = valueAxisSetting.majorTickMark },
			new C.MinorTickMark { Val = valueAxisSetting.minorTickMark },
			new C.TickLabelPosition { Val = C.TickLabelPositionValues.NextTo });
		valueAxis.Append(CreateChartShapeProperties());
		SolidFillModel solidFillModel = new()
		{
			schemeColorModel = new()
			{
				themeColorValues = ThemeColorValues.TEXT_1,
				luminanceModulation = 65000,
				luminanceOffset = 35000
			}
		};
		if (valueAxisSetting.fontColor != null)
		{
			solidFillModel.hexColor = valueAxisSetting.fontColor;
			solidFillModel.schemeColorModel = null;
		}
		valueAxis.Append(CreateChartTextProperties(new()
		{
			drawingBodyProperties = new(),
			drawingParagraph = new()
			{
				paragraphPropertiesModel = new()
				{
					defaultRunProperties = new()
					{
						solidFill = solidFillModel,
						fontSize = ConverterUtils.FontSizeToFontSize(valueAxisSetting.fontSize),
						isBold = valueAxisSetting.isBold,
						isItalic = valueAxisSetting.isItalic,
						underline = valueAxisSetting.underLineValues,
						strike = valueAxisSetting.strikeValues,
						baseline = 0
					}
				}
			}
		}));
		valueAxis.Append(
			new C.CrossingAxis { Val = valueAxisSetting.crossAxisId },
			new C.Crosses { Val = valueAxisSetting.crosses },
			new C.CrossBetween { Val = C.CrossBetweenValues.Between });
		return valueAxis;
	}

	/// <summary>
	/// Create Value Axis Data for the chart
	/// </summary>
	internal static C.Values CreateValueAxisData(string formula, ChartData[] cells)
	{
		if (cells.All(v => v.dataType != DataType.NUMBER))
		{
			Console.WriteLine($"Object Details Value : {cells.FirstOrDefault(v => v.dataType != DataType.NUMBER)?.value} is not numaric");
			Console.WriteLine($"Object Details Number Format : {cells.FirstOrDefault(v => v.dataType != DataType.NUMBER)?.numberFormat}");
			Console.WriteLine($"Object Details Data Type : {cells.FirstOrDefault(v => v.dataType != DataType.NUMBER)?.dataType}");

			throw new ArgumentException($"Value Axis Data Should Be numaric.");
		}
		return new(new C.NumberReference(new C.Formula(formula), AddNumberCacheValue(cells)));
	}

	/// <summary>
	/// Create X Axis Data for the chart
	/// </summary>
	internal static C.XValues CreateXValueAxisData(string formula, ChartData[] cells)
	{
		if (cells.All(v => v.dataType != DataType.NUMBER))
		{
			Console.WriteLine($"Object Details Value : {cells.FirstOrDefault(v => v.dataType != DataType.NUMBER)?.value} is not numaric");
			Console.WriteLine($"Object Details Number Format : {cells.FirstOrDefault(v => v.dataType != DataType.NUMBER)?.numberFormat}");
			Console.WriteLine($"Object Details Data Type : {cells.FirstOrDefault(v => v.dataType != DataType.NUMBER)?.dataType}");

			throw new ArgumentException($"X Axis Data Should Be numaric.");
		}
		return new(new C.NumberReference(new C.Formula(formula), AddNumberCacheValue(cells)));
	}

	/// <summary>
	/// Create Y Axis Data for the chart
	/// </summary>
	internal static C.YValues CreateYValueAxisData(string formula, ChartData[] cells)
	{
		if (cells.All(v => v.dataType != DataType.NUMBER))
		{
			Console.WriteLine($"Object Details Value : {cells.FirstOrDefault(v => v.dataType != DataType.NUMBER)?.value} is not numaric");
			Console.WriteLine($"Object Details Number Format : {cells.FirstOrDefault(v => v.dataType != DataType.NUMBER)?.numberFormat}");
			Console.WriteLine($"Object Details Data Type : {cells.FirstOrDefault(v => v.dataType != DataType.NUMBER)?.dataType}");
			throw new ArgumentException($"Y Axis Data Should be numaric.");
		}
		return new(new C.NumberReference(new C.Formula(formula), AddNumberCacheValue(cells)));
	}

	/// <summary>
	/// Set chart plot area
	/// </summary>
	internal void SetChartPlotArea(C.PlotArea plotArea)
	{
		chart.PlotArea = plotArea;
	}

	/// <summary>
	///
	/// </summary>
	internal C.PlotArea? GetPlotArea()
	{
		return chart.PlotArea;
	}

	private static C15.DataLabelsRangeChache AddDataLabelCacheValue(ChartData[] cells)
	{
		try
		{
			C15.DataLabelsRangeChache dataLabelsRangeChache = new()
			{
				PointCount = new C.PointCount()
				{
					Val = (UInt32Value)(uint)cells.Length
				},
			};
			int count = 0;
			foreach (ChartData Cell in cells)
			{
				C.StringPoint stringPoint = new()
				{
					Index = (UInt32Value)(uint)count,
				};
				stringPoint.AppendChild(new C.NumericValue(Cell.value ?? ""));
				dataLabelsRangeChache.AppendChild(stringPoint);
				++count;
			}
			return dataLabelsRangeChache;
		}
		catch
		{
			throw new Exception("Chart. Data Label Ref Error");
		}
	}

	private static C.NumberingCache AddNumberCacheValue(ChartData[] cells)
	{
		try
		{
			C.NumberingCache numberingCache = new()
			{
				PointCount = new C.PointCount()
				{
					Val = (UInt32Value)(uint)cells.Length
				},
			};
			int count = 0;
			foreach (ChartData Cell in cells)
			{
				C.NumericPoint numericPoint = new()
				{
					Index = (UInt32Value)(uint)count,
					FormatCode = Cell.numberFormat,
				};
				numericPoint.AppendChild(new C.NumericValue(Cell.value ?? ""));
				numberingCache.AppendChild(numericPoint);
				++count;
			}
			return numberingCache;
		}
		catch
		{
			throw new Exception("Chart. Numeric Ref Error");
		}
	}

	private static C.StringCache AddStringCacheValue(ChartData[] cells)
	{
		try
		{
			C.StringCache stringCache = new()
			{
				PointCount = new C.PointCount()
				{
					Val = (UInt32Value)(uint)cells.Length
				},
			};
			int count = 0;
			foreach (ChartData Cell in cells)
			{
				C.StringPoint stringPoint = new()
				{
					Index = (UInt32Value)(uint)count
				};
				stringPoint.AppendChild(new C.NumericValue(Cell.value ?? ""));
				stringCache.AppendChild(stringPoint);
				++count;
			}
			return stringCache;
		}
		catch
		{
			throw new Exception("Chart. String Ref Error");
		}
	}

	private C.Chart CreateChart()
	{
		C.Chart chart = new()
		{
			PlotVisibleOnly = new C.PlotVisibleOnly()
			{
				Val = true
			},
			AutoTitleDeleted = new C.AutoTitleDeleted()
			{
				Val = true
			},
			DisplayBlanksAs = new C.DisplayBlanksAs()
			{
				Val = C.DisplayBlanksAsValues.Gap
			},
			ShowDataLabelsOverMaximum = new C.ShowDataLabelsOverMaximum()
			{
				Val = false
			}
		};
		if (chartSetting.chartLegendOptions.isEnableLegend)
		{
			chart.Legend = CreateChartLegend(chartSetting.chartLegendOptions);
		}
		if (chartSetting.titleOptions != null)
		{
			chart.Title = CreateTitle(chartSetting.titleOptions);
		}
		return chart;
	}

	private C.Legend CreateChartLegend(ChartLegendOptions chartLegendOptions)
	{
		C.Legend legend = new();
		if (chartLegendOptions.manualLayout != null)
		{
			legend.Append(CreateLayout(chartLegendOptions.manualLayout));
		}
		legend.Append(new C.LegendPosition()
		{
			Val = chartLegendOptions.legendPosition switch
			{
				ChartLegendOptions.LegendPositionValues.TOP_RIGHT => C.LegendPositionValues.TopRight,
				ChartLegendOptions.LegendPositionValues.TOP => C.LegendPositionValues.Top,
				ChartLegendOptions.LegendPositionValues.BOTTOM => C.LegendPositionValues.Bottom,
				ChartLegendOptions.LegendPositionValues.LEFT => C.LegendPositionValues.Left,
				ChartLegendOptions.LegendPositionValues.RIGHT => C.LegendPositionValues.Right,
				_ => C.LegendPositionValues.Bottom
			}
		});
		legend.Append(new C.Overlay { Val = chartLegendOptions.isLegendChartOverLap });
		legend.Append(CreateChartShapeProperties());
		SolidFillModel solidFillModel = new()
		{
			schemeColorModel = new()
			{
				themeColorValues = ThemeColorValues.TEXT_1,
				luminanceModulation = 65000,
				luminanceOffset = 35000
			}
		};
		if (chartLegendOptions.fontColor != null)
		{
			solidFillModel.hexColor = chartLegendOptions.fontColor;
			solidFillModel.schemeColorModel = null;
		}
		legend.Append(CreateChartTextProperties(new()
		{
			drawingBodyProperties = new()
			{
				rotation = 0,
				useParagraphSpacing = true,
				verticalOverflow = TextVerticalOverflowValues.ELLIPSIS,
				vertical = TextVerticalAlignmentValues.HORIZONTAL,
				wrap = TextWrappingValues.SQUARE,
				anchor = TextAnchoringValues.CENTER,
				anchorCenter = true,
			},
			drawingParagraph = new()
			{
				paragraphPropertiesModel = new()
				{
					defaultRunProperties = new()
					{
						solidFill = solidFillModel,
						complexScriptFont = "+mn-cs",
						eastAsianFont = "+mn-ea",
						latinFont = "+mn-lt",
						fontSize = ConverterUtils.FontSizeToFontSize(chartLegendOptions.fontSize),
						isBold = chartLegendOptions.isBold,
						isItalic = chartLegendOptions.isItalic,
						underline = chartLegendOptions.underLineValues,
						strike = chartLegendOptions.strikeValues,
						kerning = 1200,
						baseline = 0,
					}
				}
			}
		}));
		return legend;
	}

	private static C.ChartSpace CreateChartSpace()
	{
		C.ChartSpace chartSpace = new();
		chartSpace.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
		chartSpace.AddNamespaceDeclaration("c", "http://schemas.openxmlformats.org/drawingml/2006/chart");
		chartSpace.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
		chartSpace.AddNamespaceDeclaration("c15", "http://schemas.microsoft.com/office/drawing/2012/chart");
		chartSpace.RoundedCorners = new C.RoundedCorners()
		{
			Val = false
		};
		chartSpace.Date1904 = new C.Date1904()
		{
			Val = false
		};
		chartSpace.EditingLanguage = new C.EditingLanguage()
		{
			Val = "en-IN"
		};
		return chartSpace;
	}

	private static A.Field CreateField(string type, string text)
	{
		return new A.Field(
			new A.RunProperties() { Language = "en-IN" },
			new A.ParagraphProperties(),
			new A.Text()
			{
				Text = text
			}
		)
		{ Type = type, Id = GeneratorUtils.GenerateNewGUID() };
	}

	private C.MajorGridlines CreateMajorGridLine()
	{
		return new(CreateChartShapeProperties(new()
		{
			outline = new OutlineModel()
			{
				solidFill = new SolidFillModel()
				{
					schemeColorModel = new SchemeColorModel()
					{
						themeColorValues = ThemeColorValues.TEXT_1,
						luminanceModulation = 15000,
						luminanceOffset = 85000
					}
				},
				width = 9525,
				outlineCapTypeValues = OutlineCapTypeValues.FLAT,
				outlineLineTypeValues = OutlineLineTypeValues.SINGEL,
				outlineAlignmentValues = OutlineAlignmentValues.CENTER
			}
		}));
	}

	private C.MinorGridlines CreateMinorGridLine()
	{
		return new(CreateChartShapeProperties(new()
		{
			outline = new OutlineModel()
			{
				solidFill = new SolidFillModel()
				{
					schemeColorModel = new SchemeColorModel()
					{
						themeColorValues = ThemeColorValues.TEXT_1,
						luminanceModulation = 5000,
						luminanceOffset = 95000
					}
				},
				width = 9525,
				outlineCapTypeValues = OutlineCapTypeValues.FLAT,
				outlineLineTypeValues = OutlineLineTypeValues.SINGEL,
				outlineAlignmentValues = OutlineAlignmentValues.CENTER
			}
		}));
	}

	private C.Title CreateTitle(ChartTitleModel titleModel)
	{
		SolidFillModel solidFillModel = new()
		{
			schemeColorModel = new()
			{
				themeColorValues = ThemeColorValues.TEXT_1
			}
		};
		if (titleModel.fontColor != null)
		{
			solidFillModel.hexColor = titleModel.fontColor;
			solidFillModel.schemeColorModel = null;
		}
		C.Title title = new(new C.ChartText(CreateChartRichText(new()
		{
			drawingBodyProperties = new()
			{
				anchor = TextAnchoringValues.CENTER,
				anchorCenter = true,
				useParagraphSpacing = true,
				vertical = TextVerticalAlignmentValues.HORIZONTAL,
				verticalOverflow = TextVerticalOverflowValues.ELLIPSIS,
				wrap = TextWrappingValues.SQUARE,
				rotation = 0,
			},
			drawingParagraph = new()
			{
				paragraphPropertiesModel = new(),
				drawingRun = new()
				{
					text = titleModel.title,
					drawingRunProperties = new()
					{
						solidFill = solidFillModel,
						fontSize = titleModel.fontSize,
						isBold = titleModel.isBold,
						isItalic = titleModel.isItalic,
						underline = titleModel.underLineValues,
					}
				}
			}
		})));
		title.Append(new C.Overlay { Val = false });
		title.Append(CreateChartShapeProperties());
		return title;
	}

	/// <summary>
	///
	/// </summary>
	internal C.Marker CreateMarker(MarkerModel marketModel)
	{
		C.Marker marker = new()
		{
			Symbol = new() { Val = MarkerModel.GetMarkerStyleValues(marketModel.markerShapeValues) },
		};
		if (marketModel.markerShapeValues != MarkerModel.MarkerShapeValues.NONE)
		{
			marker.Size = new() { Val = (ByteValue)marketModel.size };
			marker.Append(CreateChartShapeProperties(marketModel.shapeProperties));
		}
		return marker;
	}

}
