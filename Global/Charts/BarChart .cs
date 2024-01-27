// Copyright (c) DraviaVemal. Licensed under the MIT License. See License in the project root.

using CS = DocumentFormat.OpenXml.Office2013.Drawing.ChartStyle;

namespace OpenXMLOffice.Global_2013
{
	/// <summary>
	/// Represents the settings for a bar chart.
	/// </summary>
	public class BarChart : BarFamilyChart
	{        /// <summary>
			 /// Create Bar Chart with provided settings
			 /// </summary>
			 /// <param name="BarChartSetting">
			 /// </param>
			 /// <param name="DataCols">
			 /// </param>
		public BarChart(BarChartSetting BarChartSetting, ChartData[][] DataCols) : base(BarChartSetting, DataCols)
		{
		}

		/// <summary>
		/// Get Chart Style
		/// </summary>
		/// <returns>
		/// </returns>
		public static CS.ChartStyle GetChartStyle()
		{
			return CreateChartStyles();
		}

		/// <summary>
		/// Get Color Style
		/// </summary>
		/// <returns>
		/// </returns>
		public static CS.ColorStyle GetColorStyle()
		{
			return CreateColorStyles();
		}


	}
}
