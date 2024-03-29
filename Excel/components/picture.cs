// Copyright (c) DraviaVemal. Licensed under the MIT License. See License in the project root.

using OpenXMLOffice.Global_2013;
using DocumentFormat.OpenXml.Packaging;
using A = DocumentFormat.OpenXml.Drawing;
using X = DocumentFormat.OpenXml.Spreadsheet;
using XDR = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace OpenXMLOffice.Excel_2013
{
	/// <summary>
	/// Excel Picture
	/// </summary>
	public class Picture
	{
		private readonly ExcelPictureSetting excelPictureSetting;
		private readonly Worksheet currentWorksheet;
		/// <summary>
		/// Initializes a new instance of the <see cref="Picture"/> class.
		/// </summary>
		public Picture(Worksheet worksheet, string filePath, ExcelPictureSetting excelPictureSetting)
		{
			this.excelPictureSetting = excelPictureSetting;
			currentWorksheet = worksheet;
			string embedId = GetNextSlideRelationId();
			GetDrawingsPart().WorksheetDrawing.Append(CreateTwoCellAnchor(embedId));
			ImagePart imagePart = GetDrawingsPart().AddNewPart<ImagePart>(excelPictureSetting.imageType switch
			{
				ImageType.PNG => "image/png",
				ImageType.GIF => "image/gif",
				ImageType.TIFF => "image/tiff",
				_ => "image/jpeg"
			}, embedId);
			imagePart.FeedData(new FileStream(filePath, FileMode.Open, FileAccess.Read));
			CreateTwoCellAnchor(embedId);
			worksheet.GetWorksheet().Append(new X.Drawing() { Id = GetDrawingsPart().GetIdOfPart(imagePart) });
		}

		/// <summary>
		/// Initializes a new instance of the <see cref="Picture"/> class.
		/// </summary>
		public Picture(Worksheet worksheet, Stream stream, ExcelPictureSetting excelPictureSetting)
		{
			this.excelPictureSetting = excelPictureSetting;
			currentWorksheet = worksheet;
			string embedId = currentWorksheet.GetNextSlideRelationId();
			GetDrawingsPart().WorksheetDrawing.Append(CreateTwoCellAnchor(embedId));
			ImagePart imagePart = GetDrawingsPart().AddNewPart<ImagePart>(excelPictureSetting.imageType switch
			{
				ImageType.PNG => "image/png",
				ImageType.GIF => "image/gif",
				ImageType.TIFF => "image/tiff",
				_ => "image/jpeg"
			}, embedId);
			imagePart.FeedData(stream);
			CreateTwoCellAnchor(embedId);
			worksheet.GetWorksheet().Append(new X.Drawing() { Id = currentWorksheet.GetWorksheetPart().GetIdOfPart(GetDrawingsPart()) });
		}

		private DrawingsPart GetDrawingsPart()
		{
			if (currentWorksheet.GetWorksheetPart().DrawingsPart == null)
			{
				currentWorksheet.GetWorksheetPart().AddNewPart<DrawingsPart>(currentWorksheet.GetNextSlideRelationId());
				currentWorksheet.GetWorksheetPart().Worksheet.Save();
				currentWorksheet.GetWorksheetPart().DrawingsPart!.WorksheetDrawing ??= new();
			}
			return currentWorksheet.GetWorksheetPart().DrawingsPart!;
		}

		internal string GetNextSlideRelationId()
		{
			return string.Format("rId{0}", GetDrawingsPart().Parts.Count() + 1);
		}

		internal XDR.TwoCellAnchor CreateTwoCellAnchor(string embedId)
		{
			return new(CreatePicture(embedId), new XDR.ClientData())
			{
				EditAs = XDR.EditAsValues.OneCell,
				FromMarker = new()
				{
					ColumnId = new XDR.ColumnId(excelPictureSetting.fromCol.ToString()),
					ColumnOffset = new XDR.ColumnOffset(excelPictureSetting.fromColOff.ToString()),
					RowId = new XDR.RowId(excelPictureSetting.fromRow.ToString()),
					RowOffset = new XDR.RowOffset(excelPictureSetting.fromRowOff.ToString())
				},
				ToMarker = new()
				{
					ColumnId = new XDR.ColumnId(excelPictureSetting.toCol.ToString()),
					ColumnOffset = new XDR.ColumnOffset(excelPictureSetting.toColOff.ToString()),
					RowId = new XDR.RowId(excelPictureSetting.toRow.ToString()),
					RowOffset = new XDR.RowOffset(excelPictureSetting.toRowOff.ToString())
				},
			};
		}

		internal XDR.Picture CreatePicture(string embedId)
		{
			return new()
			{
				NonVisualPictureProperties = new()
				{
					NonVisualDrawingProperties = new()
					{
						Id = 2U,
						Name = "Picture 1"
					},
					NonVisualPictureDrawingProperties = new()
					{
						PictureLocks = new() { NoChangeAspect = true }
					}
				},
				BlipFill = new(new A.Stretch(new A.FillRectangle()))
				{
					Blip = new()
					{
						Embed = embedId
					}
				},
				ShapeProperties = new(new A.PresetGeometry(new A.AdjustValueList())
				{
					Preset = A.ShapeTypeValues.Rectangle
				})
			};
		}
	}

}
