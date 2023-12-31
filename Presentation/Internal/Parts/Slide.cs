/*
* Copyright (c) DraviaVemal. All Rights Reserved. Licensed under the MIT License.
* See License in the project root for license information.
*/

using DocumentFormat.OpenXml.Packaging;
using OpenXMLOffice.Global;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace OpenXMLOffice.Presentation
{
    public class Slide
    {
        #region Private Fields

        private readonly P.Slide OpenXMLSlide = new();

        #endregion Private Fields

        #region Internal Constructors

        internal Slide(P.Slide? OpenXMLSlide = null)
        {
            if (OpenXMLSlide != null)
            {
                this.OpenXMLSlide = OpenXMLSlide;
            }
            else
            {
                CommonSlideData commonSlideData = new(PresentationConstants.CommonSlideDataType.SLIDE, PresentationConstants.SlideLayoutType.BLANK);
                this.OpenXMLSlide.CommonSlideData = commonSlideData.GetCommonSlideData();
                this.OpenXMLSlide.ColorMapOverride = new P.ColorMapOverride()
                {
                    MasterColorMapping = new A.MasterColorMapping()
                };
                this.OpenXMLSlide.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
                this.OpenXMLSlide.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            }
        }

        #endregion Internal Constructors

        #region Public Methods

        public Chart AddChart(Excel.DataCell[][] DataCells, AreaChartSetting AreaChartSetting)
        {
            Chart Chart = new(this, DataCells, AreaChartSetting);
            P.GraphicFrame GraphicFrame = Chart.GetChartGraphicFrame();
            GetSlide().CommonSlideData!.ShapeTree!.Append(GraphicFrame);
            return Chart;
        }

        public Chart AddChart(Excel.DataCell[][] DataCells, BarChartSetting BarChartSetting)
        {
            Chart Chart = new(this, DataCells, BarChartSetting);
            P.GraphicFrame GraphicFrame = Chart.GetChartGraphicFrame();
            GetSlide().CommonSlideData!.ShapeTree!.Append(GraphicFrame);
            return Chart;
        }

        public Chart AddChart(Excel.DataCell[][] DataCells, ColumnChartSetting ColumnChartSetting)
        {
            Chart Chart = new(this, DataCells, ColumnChartSetting);
            P.GraphicFrame GraphicFrame = Chart.GetChartGraphicFrame();
            GetSlide().CommonSlideData!.ShapeTree!.Append(GraphicFrame);
            return Chart;
        }

        public Chart AddChart(Excel.DataCell[][] DataCells, LineChartSetting LineChartSetting)
        {
            Chart Chart = new(this, DataCells, LineChartSetting);
            P.GraphicFrame GraphicFrame = Chart.GetChartGraphicFrame();
            GetSlide().CommonSlideData!.ShapeTree!.Append(GraphicFrame);
            return Chart;
        }

        public Chart AddChart(Excel.DataCell[][] DataCells, PieChartSetting PieChartSetting)
        {
            Chart Chart = new(this, DataCells, PieChartSetting);
            P.GraphicFrame GraphicFrame = Chart.GetChartGraphicFrame();
            GetSlide().CommonSlideData!.ShapeTree!.Append(GraphicFrame);
            return Chart;
        }

        public Chart AddChart(Excel.DataCell[][] DataCells, ScatterChartSetting ScatterChartSetting)
        {
            Chart Chart = new(this, DataCells, ScatterChartSetting);
            P.GraphicFrame GraphicFrame = Chart.GetChartGraphicFrame();
            GetSlide().CommonSlideData!.ShapeTree!.Append(GraphicFrame);
            return Chart;
        }

        public Picture AddPicture(string FilePath, PictureSetting PictureSetting)
        {
            Picture Picture = new(FilePath, this, PictureSetting);
            GetSlide().CommonSlideData!.ShapeTree!.Append(Picture.GetPicture());
            return Picture;
        }

        public Picture AddPicture(Stream Stream, PictureSetting PictureSetting)
        {
            Picture Picture = new(Stream, this, PictureSetting);
            GetSlide().CommonSlideData!.ShapeTree!.Append(Picture.GetPicture());
            return Picture;
        }

        public Table AddTable(TableRow[] DataCells, TableSetting TableSetting)
        {
            Table Table = new(DataCells, TableSetting);
            P.GraphicFrame GraphicFrame = Table.GetTableGraphicFrame();
            GetSlide().CommonSlideData!.ShapeTree!.Append(GraphicFrame);
            return Table;
        }

        public IEnumerable<Shape> FindShapeByText(string searchText)
        {
            IEnumerable<P.Shape> searchResults = GetCommonSlideData().ShapeTree!.Elements<P.Shape>().Where(shape =>
            {
                return shape.InnerText == searchText;
            });
            return searchResults.Select(shape =>
            {
                return new Shape(shape);
            });
        }

        #endregion Public Methods

        #region Internal Methods

        internal string GetNextSlideRelationId()
        {
            return string.Format("rId{0}", GetSlidePart().Parts.Count() + 1);
        }

        internal P.Slide GetSlide()
        {
            return OpenXMLSlide;
        }

        internal SlidePart GetSlidePart()
        {
            return OpenXMLSlide.SlidePart!;
        }

        #endregion Internal Methods

        #region Private Methods

        private P.CommonSlideData GetCommonSlideData()
        {
            return OpenXMLSlide.CommonSlideData!;
        }

        #endregion Private Methods
    }
}