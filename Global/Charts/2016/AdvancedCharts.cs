// Copyright (c) DraviaVemal. Licensed under the MIT License. See License in the project root.

using OpenXMLOffice.Global_2013;
using CX = DocumentFormat.OpenXml.Office2016.Drawing.ChartDrawing;

namespace OpenXMLOffice.Global_2016
{
	/// <summary>
	///
	/// </summary>
	public class AdvanceCharts : ChartBase
	{
		private readonly CX.ChartSpace chartSpace = new();
		/// <summary>
		///
		/// </summary>
		internal AdvanceCharts(ChartSetting chartSetting) : base(chartSetting) { }

		/// <summary>
		///
		/// </summary>
		/// <returns></returns>
		public CX.ChartSpace GetExtendedChartSpace()
		{
			return chartSpace;
		}
		/// <summary>
		///
		/// </summary>
		private CX.StringDimension CreateStringDimension(string formula, ChartData[] cells)
		{
			CX.StringDimension stringDimension = new() { Type = CX.StringDimensionType.Cat };
			stringDimension.Append(new CX.Formula(formula));
			CX.StringLevel stringLevel = new() { PtCount = (uint)cells.Length };
			uint index = 0;
			cells.ToList().ForEach(cell =>
			{
				stringLevel.Append(new CX.ChartStringValue(cell.value!) { Index = index });
				index++;
			});
			stringDimension.Append(stringLevel);
			return stringDimension;
		}

		/// <summary>
		///
		/// </summary>
		private CX.NumericDimension CreateNumberDimension(string formula, ChartData[] cells)
		{
			CX.NumericDimension numericDimension = new() { Type = CX.NumericDimensionType.Val };
			numericDimension.Append(new CX.Formula(formula));
			CX.NumericLevel numericLevel = new() { PtCount = (uint)cells.Length, FormatCode = cells[0].numberFormat };
			uint index = 0;
			cells.ToList().ForEach(cell =>
			{
				numericLevel.Append(new CX.NumericValue(cell.value!) { Idx = index });
				index++;
			});
			numericDimension.Append(numericLevel);
			return numericDimension;
		}

		private CX.Series CreateSeries(ChartDataGrouping dataSeries)
		{
			CX.Series series = new()
			{
				LayoutId = CX.SeriesLayout.Waterfall,
				UniqueId = "{BCAD149B-F3BE-45E7-9A19-6B9C31CD6306}"
			};
			series.Append(new CX.Text(
				new CX.TextData(
					new CX.Formula(dataSeries.seriesHeaderFormula!),
					new CX.VXsdstring(dataSeries.seriesHeaderCells!.value!)
				)
			));
			series.Append(new CX.DataLabels(
				new CX.DataLabelVisibilities()
				{
					SeriesName = false,
					CategoryName = false,
					Value = true
				}
			)
			{ Pos = CX.DataLabelPos.OutEnd });
			series.Append(new CX.DataId() { Val = 0 });
			series.Append(new CX.SeriesLayoutProperties(
				new CX.Subtotals(
					new CX.UnsignedIntegerType() { Val = 0 },
					new CX.UnsignedIntegerType() { Val = 4 },
					new CX.UnsignedIntegerType() { Val = 7 }
				)
			));
			return series;
		}

		/// <summary>
		///
		/// </summary>
		private CX.Data CreateData(ChartDataGrouping dataSeries)
		{
			CX.Data data = new()
			{
				Id = (uint)dataSeries.id
			};
			data.Append(CreateStringDimension(dataSeries.xAxisFormula!, dataSeries.xAxisCells!));
			data.Append(CreateNumberDimension(dataSeries.yAxisFormula!, dataSeries.yAxisCells!));
			return data;
		}
		/// <summary>
		///
		/// </summary>
		internal CX.ChartData CreateChartData(List<ChartDataGrouping> chartDataGroupings)
		{
			CX.ChartData chartData = new()
			{
				ExternalData = new()
				{
					Id = "rId1",
					AutoUpdate = true
				}
			};
			chartDataGroupings.ForEach(dataSeries =>
			{
				chartData.Append(CreateData(dataSeries));
			});
			return chartData;
		}
		/// <summary>
		///
		/// </summary>
		internal CX.Chart CreateChart(List<ChartDataGrouping> chartDataGroupings)
		{
			CX.PlotAreaRegion plotAreaRegion = new();
			chartDataGroupings.Take(1).ToList().ForEach(dataSeries =>
			{
				plotAreaRegion.Append(CreateSeries(dataSeries));
			});
			CX.PlotArea plotArea = new()
			{
				PlotAreaRegion = plotAreaRegion
			};
			CX.Chart chart = new()
			{
				PlotArea = plotArea,
				Legend = new CX.Legend()
				{
					Pos = CX.SidePos.T,
					Align = CX.PosAlign.Ctr,
					Overlay = false
				},
				ChartTitle = new CX.ChartTitle()
				{
					Pos = CX.SidePos.T,
					Align = CX.PosAlign.Ctr,
					Overlay = false
				},
			};
			plotArea.Append(new CX.Axis(
				new CX.CategoryAxisScaling() { GapWidth = "0.5" },
				new CX.TickLabels()
			)
			{ Id = 0 });
			plotArea.Append(new CX.Axis(
				new CX.ValueAxisScaling(),
				new CX.MajorGridlinesGridlines(),
				new CX.TickLabels()
			)
			{ Id = 1 });
			return chart;
		}
	}
}
