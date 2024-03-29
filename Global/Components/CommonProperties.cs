// Copyright (c) DraviaVemal. Licensed under the MIT License. See License in the project root.

using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OpenXMLOffice.Global_2013
{
	/// <summary>
	/// Common Properties organised in one place to get inherited by child classes
	/// </summary>
	public class CommonProperties
	{
		/// <summary>
		///
		/// </summary>
		public static A.TextAlignmentTypeValues GetTextAlignmentValue(HorizontalAlignmentValues horizontalAlignmentValues)
		{
			return horizontalAlignmentValues switch
			{
				HorizontalAlignmentValues.RIGHT => A.TextAlignmentTypeValues.Right,
				HorizontalAlignmentValues.JUSTIFY => A.TextAlignmentTypeValues.Justified,
				HorizontalAlignmentValues.CENTER => A.TextAlignmentTypeValues.Center,
				_ => A.TextAlignmentTypeValues.Left
			};
		}
		/// <summary>
		///
		/// </summary>
		public static A.CompoundLineValues GetBorderStyleValue(BorderStyleValues borderStyle)
		{
			return borderStyle switch
			{
				BorderStyleValues.DOUBLE => A.CompoundLineValues.Double,
				BorderStyleValues.TRIPLE => A.CompoundLineValues.Triple,
				BorderStyleValues.THICK_THIN => A.CompoundLineValues.ThickThin,
				BorderStyleValues.THIN_THICK => A.CompoundLineValues.ThinThick,
				_ => A.CompoundLineValues.Single,
			};
		}
		/// <summary>
		///
		/// </summary>
		public static A.PresetLineDashValues GetDashStyleValue(DrawingPresetLineDashValues dashStyle)
		{
			return dashStyle switch
			{
				DrawingPresetLineDashValues.DASH => A.PresetLineDashValues.Dash,
				DrawingPresetLineDashValues.DASH_DOT => A.PresetLineDashValues.DashDot,
				DrawingPresetLineDashValues.DOT => A.PresetLineDashValues.Dot,
				DrawingPresetLineDashValues.LARGE_DASH => A.PresetLineDashValues.LargeDash,
				DrawingPresetLineDashValues.LARGE_DASH_DOT => A.PresetLineDashValues.LargeDashDot,
				DrawingPresetLineDashValues.LARGE_DASH_DOT_DOT => A.PresetLineDashValues.LargeDashDotDot,
				DrawingPresetLineDashValues.SYSTEM_DASH => A.PresetLineDashValues.SystemDash,
				DrawingPresetLineDashValues.SYSTEM_DASH_DOT => A.PresetLineDashValues.SystemDashDot,
				DrawingPresetLineDashValues.SYSTEM_DASH_DOT_DOT => A.PresetLineDashValues.SystemDashDotDot,
				DrawingPresetLineDashValues.SYSTEM_DOT => A.PresetLineDashValues.SystemDot,
				_ => A.PresetLineDashValues.Solid,
			};
		}

		/// <summary>
		///
		/// </summary>
		public static A.LineEndValues GetEndArrowValue(DrawingEndArrowValues endArrowValues)
		{
			return endArrowValues switch
			{
				DrawingEndArrowValues.ARROW => A.LineEndValues.Arrow,
				DrawingEndArrowValues.DIAMOND => A.LineEndValues.Diamond,
				DrawingEndArrowValues.OVAL => A.LineEndValues.Oval,
				DrawingEndArrowValues.STEALTH => A.LineEndValues.Stealth,
				DrawingEndArrowValues.TRIANGLE => A.LineEndValues.Triangle,
				_ => A.LineEndValues.None,
			};
		}

		/// <summary>
		///
		/// </summary>
		public static A.LineEndLengthValues GetLineEndLengthValue(LineWidthValues lineEndWidth)
		{
			return lineEndWidth switch
			{
				LineWidthValues.LARGE => A.LineEndLengthValues.Large,
				LineWidthValues.MEDIUM => A.LineEndLengthValues.Medium,
				_ => A.LineEndLengthValues.Small
			};
		}

		/// <summary>
		///
		/// </summary>
		public static A.LineEndWidthValues GetLineEndWidthValue(LineWidthValues lineEndWidth)
		{
			return lineEndWidth switch
			{
				LineWidthValues.LARGE => A.LineEndWidthValues.Large,
				LineWidthValues.MEDIUM => A.LineEndWidthValues.Medium,
				_ => A.LineEndWidthValues.Small
			};
		}

		/// <summary>
		///
		/// </summary>
		public static A.LineEndLengthValues GetLineStartLengthValue(LineWidthValues lineStartWidth)
		{
			return lineStartWidth switch
			{
				LineWidthValues.LARGE => A.LineEndLengthValues.Large,
				LineWidthValues.MEDIUM => A.LineEndLengthValues.Medium,
				_ => A.LineEndLengthValues.Small
			};
		}

		/// <summary>
		///
		/// </summary>
		public static A.LineEndValues GetBeginArrowValue(DrawingBeginArrowValues beginArrowValues)
		{
			return beginArrowValues switch
			{
				DrawingBeginArrowValues.ARROW => A.LineEndValues.Arrow,
				DrawingBeginArrowValues.DIAMOND => A.LineEndValues.Diamond,
				DrawingBeginArrowValues.OVAL => A.LineEndValues.Oval,
				DrawingBeginArrowValues.STEALTH => A.LineEndValues.Stealth,
				DrawingBeginArrowValues.TRIANGLE => A.LineEndValues.Triangle,
				_ => A.LineEndValues.None,
			};
		}

		/// <summary>
		///
		/// </summary>
		public static A.LineEndWidthValues GetLineStartWidthValue(LineWidthValues lineStartWidth)
		{
			return lineStartWidth switch
			{
				LineWidthValues.LARGE => A.LineEndWidthValues.Large,
				LineWidthValues.MEDIUM => A.LineEndWidthValues.Medium,
				_ => A.LineEndWidthValues.Small
			};
		}

		internal static A.TextAnchoringTypeValues GetAnchorValues(TextAnchoringValues textAnchoring)
		{
			return textAnchoring switch
			{
				TextAnchoringValues.BOTTOM => A.TextAnchoringTypeValues.Bottom,
				TextAnchoringValues.CENTER => A.TextAnchoringTypeValues.Center,
				_ => A.TextAnchoringTypeValues.Top
			};
		}

		internal static A.TextVerticalValues GetTextVerticalAlignmentValues(TextVerticalAlignmentValues textVerticalAlignment)
		{
			return textVerticalAlignment switch
			{
				TextVerticalAlignmentValues.EAST_ASIAN_VERTICAL => A.TextVerticalValues.EastAsianVetical,
				TextVerticalAlignmentValues.HORIZONTAL => A.TextVerticalValues.Horizontal,
				TextVerticalAlignmentValues.MONGOLIAN_VERTICAL => A.TextVerticalValues.MongolianVertical,
				TextVerticalAlignmentValues.VERTICAL => A.TextVerticalValues.Vertical,
				TextVerticalAlignmentValues.VERTICAL_270 => A.TextVerticalValues.Vertical270,
				TextVerticalAlignmentValues.WORD_ART_LEFT_TO_RIGHT => A.TextVerticalValues.WordArtVertical,
				_ => A.TextVerticalValues.WordArtVertical
			};
		}

		internal static A.TextVerticalOverflowValues GetTextVerticalOverflowValues(TextVerticalOverflowValues textVerticalOverflow)
		{
			return textVerticalOverflow switch
			{
				TextVerticalOverflowValues.CLIP => A.TextVerticalOverflowValues.Clip,
				TextVerticalOverflowValues.ELLIPSIS => A.TextVerticalOverflowValues.Ellipsis,
				_ => A.TextVerticalOverflowValues.Overflow
			};
		}

		internal static A.TextWrappingValues GetWrapingValues(TextWrappingValues textWrapping)
		{
			return textWrapping switch
			{
				TextWrappingValues.NONE => A.TextWrappingValues.None,
				_ => A.TextWrappingValues.Square
			};
		}
		internal static A.TextStrikeValues GetTextStrikeValues(StrikeValues strikeValues)
		{
			return strikeValues switch
			{
				StrikeValues.SINGLE_STRIKE => A.TextStrikeValues.SingleStrike,
				StrikeValues.DOUBLE_STRIKE => A.TextStrikeValues.DoubleStrike,
				_ => A.TextStrikeValues.NoStrike
			};
		}
		internal static string GetSchemeColorValuesText(ThemeColorValues themeColorValues)
		{
			return themeColorValues switch
			{
				ThemeColorValues.ACCENT_1 => "accent1",
				ThemeColorValues.ACCENT_2 => "accent2",
				ThemeColorValues.ACCENT_3 => "accent3",
				ThemeColorValues.ACCENT_4 => "accent4",
				ThemeColorValues.ACCENT_5 => "accent5",
				ThemeColorValues.ACCENT_6 => "accent6",
				ThemeColorValues.DARK_1 => "dk1",
				ThemeColorValues.DARK_2 => "dk2",
				ThemeColorValues.BACKGROUND_1 => "bg1",
				ThemeColorValues.BACKGROUND_2 => "bg2",
				ThemeColorValues.LIGHT_1 => "lt1",
				ThemeColorValues.LIGHT_2 => "lt2",
				ThemeColorValues.TEXT_1 => "tx1",
				ThemeColorValues.TEXT_2 => "tx2",
				ThemeColorValues.HYPERLINK => "hlink",
				ThemeColorValues.FOLLOW_HYPERLINK => "folHlink",
				_ => "phClr"
			};
		}
		internal static A.PenAlignmentValues GetLineAlignmentValues(OutlineAlignmentValues outlineAlignmentValues)
		{
			return outlineAlignmentValues switch
			{
				OutlineAlignmentValues.CENTER => A.PenAlignmentValues.Center,
				_ => A.PenAlignmentValues.Insert
			};
		}

		internal static A.LineCapValues GetLineCapValues(OutlineCapTypeValues outlineCapTypeValues)
		{
			return outlineCapTypeValues switch
			{
				OutlineCapTypeValues.SQUARE => A.LineCapValues.Square,
				OutlineCapTypeValues.ROUND => A.LineCapValues.Round,
				_ => A.LineCapValues.Flat
			};
		}

		internal static A.CompoundLineValues GetLineTypeValues(OutlineLineTypeValues outlineLineTypeValues)
		{
			return outlineLineTypeValues switch
			{
				OutlineLineTypeValues.DOUBLE => A.CompoundLineValues.Double,
				OutlineLineTypeValues.TRIPLE => A.CompoundLineValues.Triple,
				OutlineLineTypeValues.THICK_THIN => A.CompoundLineValues.ThickThin,
				OutlineLineTypeValues.THIN_THICK => A.CompoundLineValues.ThinThick,
				_ => A.CompoundLineValues.Single
			};
		}

		internal static A.TextUnderlineValues GetTextUnderlineValues(UnderLineValues runPropertiesUnderLineValues)
		{
			return runPropertiesUnderLineValues switch
			{
				UnderLineValues.DASH => A.TextUnderlineValues.Dash,
				UnderLineValues.DASH_HEAVY => A.TextUnderlineValues.DashHeavy,
				UnderLineValues.DASH_LONG => A.TextUnderlineValues.DashLong,
				UnderLineValues.DASH_LONG_HEAVY => A.TextUnderlineValues.DashLongHeavy,
				UnderLineValues.DOT_DASH => A.TextUnderlineValues.DotDash,
				UnderLineValues.DOT_DASH_HEAVY => A.TextUnderlineValues.DotDashHeavy,
				UnderLineValues.DOT_DOT_DASH => A.TextUnderlineValues.DotDotDash,
				UnderLineValues.DOT_DOT_DASH_HEAVY => A.TextUnderlineValues.DotDotDashHeavy,
				UnderLineValues.DOTTED => A.TextUnderlineValues.Dotted,
				UnderLineValues.DOUBLE => A.TextUnderlineValues.Double,
				UnderLineValues.HEAVY => A.TextUnderlineValues.Heavy,
				UnderLineValues.HEAVY_DOTTED => A.TextUnderlineValues.HeavyDotted,
				UnderLineValues.SINGLE => A.TextUnderlineValues.Single,
				UnderLineValues.WAVY => A.TextUnderlineValues.Wavy,
				UnderLineValues.WAVY_DOUBLE => A.TextUnderlineValues.WavyDouble,
				UnderLineValues.WAVY_HEAVY => A.TextUnderlineValues.WavyHeavy,
				UnderLineValues.WORDS => A.TextUnderlineValues.Words,
				_ => A.TextUnderlineValues.None
			};
		}


		/// <summary>
		/// Class is only for inheritance purposes.
		/// </summary>
		protected CommonProperties() { }

		/// <summary>
		/// Create Soild Fill XML Property
		/// </summary>
		protected static A.SolidFill CreateSolidFill(SolidFillModel solidFillModel)
		{
			if (solidFillModel.hexColor == null && solidFillModel.schemeColorModel == null)
			{
				throw new ArgumentException("Solid Fill Color Error");
			}
			if (solidFillModel.hexColor != null)
			{
				A.RgbColorModelHex rgbColorModelHex = new() { Val = solidFillModel.hexColor };
				if (solidFillModel.transparency != null)
				{
					rgbColorModelHex.Append(new A.Alpha() { Val = 100000 - (solidFillModel.transparency * 1000) });
				}
				return new A.SolidFill() { RgbColorModelHex = rgbColorModelHex };
			}
			else
			{
				A.SchemeColor schemeColor = new()
				{ Val = new A.SchemeColorValues(GetSchemeColorValuesText(solidFillModel.schemeColorModel!.themeColorValues)) };
				if (solidFillModel.transparency != null)
				{
					schemeColor.Append(new A.Alpha() { Val = 100000 - (solidFillModel.transparency * 1000) });
				}
				if (solidFillModel.schemeColorModel.tint != null)
				{
					schemeColor.Append(new A.Tint() { Val = solidFillModel.schemeColorModel.tint });
				}
				if (solidFillModel.schemeColorModel.shade != null)
				{
					schemeColor.Append(new A.Shade() { Val = solidFillModel.schemeColorModel.shade });
				}
				if (solidFillModel.schemeColorModel.saturationModulation != null)
				{
					schemeColor.Append(new A.SaturationModulation() { Val = solidFillModel.schemeColorModel.saturationModulation });
				}
				if (solidFillModel.schemeColorModel.saturationOffset != null)
				{
					schemeColor.Append(new A.SaturationOffset() { Val = solidFillModel.schemeColorModel.saturationOffset });
				}
				if (solidFillModel.schemeColorModel.luminanceModulation != null)
				{
					schemeColor.Append(new A.LuminanceModulation() { Val = solidFillModel.schemeColorModel.luminanceModulation });
				}
				if (solidFillModel.schemeColorModel.luminanceOffset != null)
				{
					schemeColor.Append(new A.LuminanceOffset() { Val = solidFillModel.schemeColorModel.luminanceOffset });
				}
				return new A.SolidFill(schemeColor);
			}
		}

		/// <summary>
		/// Create Shape Properties With Default Settings
		/// </summary>
		/// <returns></returns>
		protected C.ShapeProperties CreateChartShapeProperties()
		{
			return CreateChartShapeProperties(new());
		}

		/// <summary>
		/// Create Shape Properties
		/// </summary>
		/// <returns></returns>
		protected C.ShapeProperties CreateChartShapeProperties(ShapePropertiesModel shapePropertiesModel)
		{
			C.ShapeProperties shapeProperties = new();
			if (shapePropertiesModel.solidFill != null)
			{
				shapeProperties.Append(CreateSolidFill(shapePropertiesModel.solidFill));
			}
			else
			{
				shapeProperties.Append(new A.NoFill());
			}
			shapeProperties.Append(CreateOutline(shapePropertiesModel.outline));
			if (shapePropertiesModel.effectList != null)
			{
				shapeProperties.Append(CreateEffectList(shapePropertiesModel.effectList));
			}
			else
			{
				shapeProperties.Append(new A.EffectList());
			}
			return shapeProperties;
		}

		/// <summary>
		/// Create Effect List
		/// </summary>
		protected static A.EffectList CreateEffectList(EffectListModel effectListModel)
		{
			return new();
		}

		/// <summary>
		/// Create Outline
		/// </summary>
		protected static A.Outline CreateOutline(OutlineModel outlineModel)
		{
			A.Outline outline = new();
			if (outlineModel.width != null)
			{
				outline.Width = outlineModel.width;
			}
			if (outlineModel.outlineCapTypeValues != null)
			{
				outline.CapType = GetLineCapValues((OutlineCapTypeValues)outlineModel.outlineCapTypeValues);
			}
			if (outlineModel.outlineLineTypeValues != null)
			{
				outline.CompoundLineType = GetLineTypeValues((OutlineLineTypeValues)outlineModel.outlineLineTypeValues);
			}
			if (outlineModel.outlineAlignmentValues != null)
			{
				outline.Alignment = GetLineAlignmentValues((OutlineAlignmentValues)outlineModel.outlineAlignmentValues);
			}
			if (outlineModel.solidFill != null)
			{
				outline.Append(CreateSolidFill(outlineModel.solidFill));
				outline.Append(new A.Round());
			}
			else
			{
				outline.Append(new A.NoFill());
			}
			if (outlineModel.dashType != null)
			{
				outline.Append(new A.PresetDash { Val = GetDashStyleValue((DrawingPresetLineDashValues)outlineModel.dashType) });
			}
			A.HeadEnd headEnd = new() { Type = GetBeginArrowValue(outlineModel.beginArrowValues) };
			if (outlineModel.lineStartWidth != null)
			{
				headEnd.Width = GetLineStartWidthValue((LineWidthValues)outlineModel.lineStartWidth);
				headEnd.Length = GetLineStartLengthValue((LineWidthValues)outlineModel.lineStartWidth);
			}
			outline.Append(headEnd);
			A.TailEnd tailEnd = new() { Type = GetEndArrowValue(outlineModel.endArrowValues) };
			if (outlineModel.lineEndWidth != null)
			{
				tailEnd.Width = GetLineEndWidthValue((LineWidthValues)outlineModel.lineEndWidth);
				tailEnd.Length = GetLineEndLengthValue((LineWidthValues)outlineModel.lineEndWidth);
			}
			outline.Append(tailEnd);
			return outline;
		}

		/// <summary>
		/// Create Default Run Properties
		/// </summary>
		protected static A.DefaultRunProperties CreateDefaultRunProperties()
		{
			return CreateDefaultRunProperties(new());
		}


		/// <summary>
		///     Create Default Run Properties
		/// </summary>
		protected static A.DefaultRunProperties CreateDefaultRunProperties(DefaultRunPropertiesModel defaultRunPropertiesModel)
		{
			A.DefaultRunProperties defaultRunProperties = new();
			if (defaultRunPropertiesModel.solidFill != null)
			{
				defaultRunProperties.Append(CreateSolidFill(defaultRunPropertiesModel.solidFill));
			}
			if (defaultRunPropertiesModel.latinFont != null)
			{
				defaultRunProperties.Append(new A.LatinFont { Typeface = defaultRunPropertiesModel.latinFont });
			}
			if (defaultRunPropertiesModel.eastAsianFont != null)
			{
				defaultRunProperties.Append(new A.EastAsianFont { Typeface = defaultRunPropertiesModel.eastAsianFont });
			}
			if (defaultRunPropertiesModel.complexScriptFont != null)
			{
				defaultRunProperties.Append(new A.ComplexScriptFont { Typeface = defaultRunPropertiesModel.complexScriptFont });
			}
			if (defaultRunPropertiesModel.fontSize != null)
			{
				defaultRunProperties.FontSize = defaultRunPropertiesModel.fontSize;
			}
			if (defaultRunPropertiesModel.isBold != null)
			{
				defaultRunProperties.Bold = defaultRunPropertiesModel.isBold;
			}
			if (defaultRunPropertiesModel.isItalic != null)
			{
				defaultRunProperties.Italic = defaultRunPropertiesModel.isItalic;
			}
			if (defaultRunPropertiesModel.underline != null)
			{
				defaultRunProperties.Underline = GetTextUnderlineValues((UnderLineValues)defaultRunPropertiesModel.underline);
			}
			if (defaultRunPropertiesModel.strike != null)
			{
				defaultRunProperties.Strike = GetTextStrikeValues((StrikeValues)defaultRunPropertiesModel.strike);
			}
			if (defaultRunPropertiesModel.kerning != null)
			{
				defaultRunProperties.Kerning = defaultRunPropertiesModel.kerning;
			}
			if (defaultRunPropertiesModel.baseline != null)
			{
				defaultRunProperties.Baseline = defaultRunPropertiesModel.baseline;
			}
			return defaultRunProperties;
		}

		/// <summary>
		///
		/// </summary>
		protected A.Paragraph CreateDrawingParagraph(DrawingParagraphModel drawingParagraphModel)
		{
			A.Paragraph paragraph = new();
			if (drawingParagraphModel.paragraphPropertiesModel != null)
			{
				paragraph.Append(CreateDrawingParagraphProperties(drawingParagraphModel.paragraphPropertiesModel));
			}
			if (drawingParagraphModel.drawingRun != null)
			{
				paragraph.Append(CreateDrawingRun(drawingParagraphModel.drawingRun));
			}
			else
			{
				if (drawingParagraphModel.paragraphPropertiesModel != null)
				{
					paragraph.Append(new A.EndParagraphRunProperties() { Language = "en-IN" });
				}
			}
			return paragraph;
		}

		/// <summary>
		///
		/// </summary>
		private A.ParagraphProperties CreateDrawingParagraphProperties(ParagraphPropertiesModel paragraphPropertiesModel)
		{
			A.ParagraphProperties paragraphProperties = new();
			if (paragraphPropertiesModel.defaultRunProperties != null)
			{
				paragraphProperties.Append(CreateDefaultRunProperties(paragraphPropertiesModel.defaultRunProperties));
			}
			if (paragraphPropertiesModel.horizontalAlignment != null)
			{
				paragraphProperties.Alignment = GetTextAlignmentValue((HorizontalAlignmentValues)paragraphPropertiesModel.horizontalAlignment);
			}
			return paragraphProperties;
		}

		/// <summary>
		///
		/// </summary>
		protected static A.ListStyle CreateDrawingListStyle()
		{
			return new();
		}

		/// <summary>
		///     Create Chart Text Properties
		/// </summary>
		protected C.TextProperties CreateChartTextProperties(ChartTextPropertiesModel chartTextPropertiesModel)
		{
			C.TextProperties textProperties = new();
			if (chartTextPropertiesModel.drawingBodyProperties != null)
			{
				textProperties.Append(CreateDrawingBodyProperties(chartTextPropertiesModel.drawingBodyProperties));
			}
			textProperties.Append(CreateDrawingListStyle());
			if (chartTextPropertiesModel.drawingParagraph != null)
			{
				textProperties.Append(CreateDrawingParagraph(chartTextPropertiesModel.drawingParagraph));
			}
			return textProperties;
		}
		/// <summary>
		///
		/// </summary>
		protected C.RichText CreateChartRichText(ChartTextPropertiesModel chartTextPropertiesModel)
		{
			C.RichText richText = new();
			if (chartTextPropertiesModel.drawingBodyProperties != null)
			{
				richText.Append(CreateDrawingBodyProperties(chartTextPropertiesModel.drawingBodyProperties));
			}
			richText.Append(CreateDrawingListStyle());
			if (chartTextPropertiesModel.drawingParagraph != null)
			{
				richText.Append(CreateDrawingParagraph(chartTextPropertiesModel.drawingParagraph));
			}
			return richText;
		}
		/// <summary>
		///
		/// </summary>
		protected static A.Run CreateDrawingRun(DrawingRunModel drawingRunModel)
		{
			A.Run run = new(CreateDrawingRunProperties(drawingRunModel.drawingRunProperties));
			if (drawingRunModel.text != null)
			{
				run.Append(new A.Text(drawingRunModel.text));
			}
			if (drawingRunModel.textHightlight != null)
			{
				run.Append(new A.Highlight(new A.RgbColorModelHex { Val = drawingRunModel.textHightlight }));
			}
			return run;
		}

		/// <summary>
		///
		/// </summary>
		protected static A.RunProperties CreateDrawingRunProperties(DrawingRunPropertiesModel drawingRunPropertiesModel)
		{
			A.RunProperties runProperties = new()
			{
				FontSize = ConverterUtils.FontSizeToFontSize(drawingRunPropertiesModel.fontSize),
				Bold = drawingRunPropertiesModel.isBold,
				Italic = drawingRunPropertiesModel.isItalic,
				Dirty = false
			};
			if (drawingRunPropertiesModel.solidFill != null)
			{
				runProperties.Append(CreateSolidFill(drawingRunPropertiesModel.solidFill));
			}
			if (drawingRunPropertiesModel.fontFamily != null)
			{
				runProperties.Append(new A.LatinFont { Typeface = drawingRunPropertiesModel.fontFamily });
			}
			if (drawingRunPropertiesModel.fontFamily != null)
			{
				runProperties.Append(new A.EastAsianFont { Typeface = drawingRunPropertiesModel.fontFamily });
			}
			if (drawingRunPropertiesModel.fontFamily != null)
			{
				runProperties.Append(new A.ComplexScriptFont { Typeface = drawingRunPropertiesModel.fontFamily });
			}
			if (drawingRunPropertiesModel.underline != null)
			{
				runProperties.Underline = GetTextUnderlineValues((UnderLineValues)drawingRunPropertiesModel.underline);
			}
			return runProperties;
		}
		/// <summary>
		///    Create Drawing Body Properties
		/// </summary>
		/// <param name="drawingBodyPropertiesModel"></param>
		/// <returns></returns>
		private static A.BodyProperties CreateDrawingBodyProperties(DrawingBodyPropertiesModel drawingBodyPropertiesModel)
		{
			A.BodyProperties bodyProperties = new(new A.ShapeAutoFit())
			{
				Rotation = drawingBodyPropertiesModel.rotation
			};
			if (drawingBodyPropertiesModel.leftInset != null)
			{
				bodyProperties.LeftInset = drawingBodyPropertiesModel.leftInset;
			}
			if (drawingBodyPropertiesModel.topInset != null)
			{
				bodyProperties.TopInset = drawingBodyPropertiesModel.topInset;
			}
			if (drawingBodyPropertiesModel.rightInset != null)
			{
				bodyProperties.RightInset = drawingBodyPropertiesModel.rightInset;
			}
			if (drawingBodyPropertiesModel.bottomInset != null)
			{
				bodyProperties.BottomInset = drawingBodyPropertiesModel.bottomInset;
			}
			if (drawingBodyPropertiesModel.useParagraphSpacing != null)
			{
				bodyProperties.UseParagraphSpacing = drawingBodyPropertiesModel.useParagraphSpacing;
			}
			if (drawingBodyPropertiesModel.verticalOverflow != null)
			{
				bodyProperties.VerticalOverflow = GetTextVerticalOverflowValues((TextVerticalOverflowValues)drawingBodyPropertiesModel.verticalOverflow);
			}
			if (drawingBodyPropertiesModel.vertical != null)
			{
				bodyProperties.Vertical = GetTextVerticalAlignmentValues((TextVerticalAlignmentValues)drawingBodyPropertiesModel.vertical);
			}
			if (drawingBodyPropertiesModel.wrap != null)
			{
				bodyProperties.Wrap = GetWrapingValues((TextWrappingValues)drawingBodyPropertiesModel.wrap);
			}
			if (drawingBodyPropertiesModel.anchor != null)
			{
				bodyProperties.Anchor = GetAnchorValues((TextAnchoringValues)drawingBodyPropertiesModel.anchor);
			}
			return bodyProperties;
		}


	}
}
