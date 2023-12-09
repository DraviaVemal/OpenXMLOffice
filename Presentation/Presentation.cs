using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace OpenXMLOffice.Presentation
{
    internal class Presentation : PresentationCore
    {
        private readonly PresentationProperties presentationProperties;
        private readonly PresentationDocument presentationDocument;
        private readonly PresentationInfo presentationInfo = new();
        private ExtendedFilePropertiesPart? extendedFilePropertiesPart;
        private PresentationPart? presentationPart;
        private SlideMasterPart? slideMasterPart;
        private SlideLayoutPart? slideLayoutPart;
        public Presentation(string filePath, bool isEditable, PresentationProperties? presentationProperties = null, bool autosave = true)
        {
            presentationInfo.FilePath = filePath;
            this.presentationProperties = presentationProperties ?? new();
            FileStream reader = new(filePath, FileMode.Open);
            MemoryStream memoryStream = new();
            reader.CopyTo(memoryStream);
            reader.Close();
            presentationDocument = PresentationDocument.Open(memoryStream, isEditable, new OpenSettings()
            {
                AutoSave = autosave
            });
            if (isEditable)
            {
                presentationInfo.IsExistingFile = true;
            }
            else
            {
                presentationInfo.IsEditable = false;
            }
        }
        public Presentation(string filePath, PresentationProperties? presentationProperties = null, PresentationDocumentType presentationDocumentType = PresentationDocumentType.Presentation, bool autosave = true)
        {
            presentationInfo.FilePath = filePath;
            this.presentationProperties = presentationProperties ?? new();
            presentationDocument = PresentationDocument.Create(filePath, presentationDocumentType, autosave);
            PreparePresentation(this.presentationProperties);
        }

        public Presentation(Stream stream, PresentationProperties? presentationProperties = null, PresentationDocumentType presentationDocumentType = PresentationDocumentType.Presentation, bool autosave = true)
        {
            this.presentationProperties = presentationProperties ?? new();
            presentationDocument = PresentationDocument.Create(stream, presentationDocumentType, autosave);
            PreparePresentation(this.presentationProperties);
        }

        private P.DefaultTextStyle CreateDefaultTextStyle()
        {
            P.DefaultTextStyle defaultTextStyle = new();
            A.DefaultParagraphProperties defaultParagraphProperties = new();
            A.DefaultRunProperties defaultRunProperties = new() { Language = "en-US" };
            defaultParagraphProperties.Append(defaultRunProperties);
            defaultTextStyle.Append(defaultParagraphProperties);
            A.Level1ParagraphProperties levelParagraphProperties = new()
            {
                Alignment = A.TextAlignmentTypeValues.Left,
                DefaultTabSize = 914400,
                EastAsianLineBreak = true,
                LatinLineBreak = false,
                LeftMargin = 457200,
                RightToLeft = false
            };
            A.DefaultRunProperties levelRunProperties = new()
            {
                Kerning = 1200,
                FontSize = 1800
            };
            A.SolidFill solidFill = new();
            A.SchemeColor schemeColor = new() { Val = A.SchemeColorValues.Text1 };
            solidFill.Append(schemeColor);
            levelRunProperties.Append(solidFill);
            A.LatinFont latinTypeface = new() { Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianTypeface = new() { Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptTypeface = new() { Typeface = "+mn-cs" };
            levelRunProperties.Append(latinTypeface, eastAsianTypeface, complexScriptTypeface);
            levelParagraphProperties.Append(levelRunProperties);
            defaultTextStyle.Append(levelParagraphProperties);
            return defaultTextStyle;
        }

        private void PreparePresentation(PresentationProperties? powerPointProperties)
        {
            SlideMaster slideMaster = new();
            SlideLayout slideLayout = new();
            presentationPart = presentationDocument.PresentationPart ?? presentationDocument.AddPresentationPart();
            presentationPart.Presentation ??= new P.Presentation();
            presentationPart.Presentation.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
            presentationPart.Presentation.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            presentationPart.Presentation.AddNamespaceDeclaration("p", "http://schemas.openxmlformats.org/presentationml/2006/main");
            if (presentationPart.Presentation.GetFirstChild<P.SlideMasterIdList>() == null)
            {
                presentationPart.Presentation.AppendChild(new P.SlideMasterIdList());
            }
            if (presentationPart.Presentation.SlideIdList == null)
            {
                presentationPart.Presentation.AppendChild(new P.SlideIdList());
            }
            if (presentationPart.Presentation.GetFirstChild<P.SlideSize>() == null)
            {
                presentationPart.Presentation.AppendChild(new P.SlideSize() { Cx = 12192000, Cy = 6858000 });
            }
            if (presentationPart.Presentation.GetFirstChild<P.NotesSize>() == null)
            {
                presentationPart.Presentation.AppendChild(new P.NotesSize() { Cx = 6858000, Cy = 9144000 });
            }

            if (presentationPart.Presentation.GetFirstChild<P.DefaultTextStyle>() == null)
            {
                presentationPart.Presentation.AppendChild(CreateDefaultTextStyle());
            }
            if (presentationPart.ViewPropertiesPart == null)
            {
                ViewProperties viewProperties = new();
                ViewPropertiesPart viewPropertiesPart = presentationPart.AddNewPart<ViewPropertiesPart>(string.Format("rId{0}", presentationPart.Parts.Count() + 1));
                viewPropertiesPart.ViewProperties = viewProperties.GetViewProperties();
                viewPropertiesPart.ViewProperties.Save();
            }
            if (presentationPart.PresentationPropertiesPart == null)
            {
                PresentationPropertiesPart presentationPropertiesPart = presentationPart.AddNewPart<PresentationPropertiesPart>(string.Format("rId{0}", presentationPart.Parts.Count() + 1));
                presentationPropertiesPart.PresentationProperties ??= new P.PresentationProperties();
                presentationPropertiesPart.PresentationProperties.Save();
            }
            if (!presentationPart.SlideMasterParts.Any())
            {
                slideMasterPart = presentationPart.AddNewPart<SlideMasterPart>(string.Format("rId{0}", presentationPart.Parts.Count() + 1));
                slideMasterPart.SlideMaster = slideMaster.GetSlideMaster();
                P.SlideMasterIdList slideMasterIdList = presentationPart.Presentation.SlideMasterIdList!;
                P.SlideMasterId slideMasterId = new() { Id = (uint)(2147483647 + slideMasterIdList.Count() + 1), RelationshipId = presentationPart.GetIdOfPart(slideMasterPart) };
                slideMasterIdList.Append(slideMasterId);
                slideLayoutPart = slideMasterPart.AddNewPart<SlideLayoutPart>(string.Format("rId{0}", presentationPart.Parts.Count() + 1));
                slideMaster.AddSlideLayoutIdToList(slideMasterPart.GetIdOfPart(slideLayoutPart));
                slideLayoutPart.SlideLayout = slideLayout.GetSlideLayout();
                slideLayout.UpdateRelationship(slideMasterPart, presentationPart.GetIdOfPart(slideMasterPart));
                slideLayoutPart.SlideLayout.Save();
                slideMasterPart.SlideMaster.Save();
            }
            if (presentationDocument.ExtendedFilePropertiesPart == null)
            {
                extendedFilePropertiesPart = presentationDocument.AddExtendedFilePropertiesPart();
                extendedFilePropertiesPart.Properties ??= new DocumentFormat.OpenXml.ExtendedProperties.Properties();
                extendedFilePropertiesPart.Properties.Save();
            }
            if (presentationPart.TableStylesPart == null)
            {
                TableStylesPart tableStylesPart = presentationPart.AddNewPart<TableStylesPart>(string.Format("rId{0}", presentationPart.Parts.Count() + 1));
                tableStylesPart.TableStyleList ??= new A.TableStyleList()
                {
                    Default = string.Format("{{{0}}}", Guid.NewGuid().ToString("D").ToUpper())
                };
                tableStylesPart.TableStyleList.Save();
            }
            if (presentationPart.ThemePart == null)
            {
                presentationPart.AddNewPart<ThemePart>(string.Format("rId{0}", presentationPart.Parts.Count() + 1));

            }
            Theme theme = new(powerPointProperties?.Theme);
            presentationPart.ThemePart!.Theme = theme.GetTheme();
            slideMaster.UpdateRelationship(presentationPart.ThemePart, presentationPart.GetIdOfPart(presentationPart.ThemePart));
            presentationPart.Presentation.Save();
        }

        public void AddBlankSlide()
        {
            SlidePart slidePart = presentationPart!.AddNewPart<SlidePart>(string.Format("rId{0}", presentationPart.Parts.Count() + 1));
            Slide slide = new();
            slidePart.Slide = slide.GetSlide();
            slidePart.AddPart(slideLayoutPart!);
            P.SlideIdList slideIdList = presentationPart.Presentation.SlideIdList!;
            P.SlideId slideId = new() { Id = (uint)(255 + slideIdList.Count() + 1), RelationshipId = presentationPart.GetIdOfPart(slidePart) };
            slideIdList.Append(slideId);
        }

        public void Save()
        {
            if (presentationInfo.FilePath == null)
            {
                throw new FieldAccessException("Data Is in File Stream Use SaveAs to Target save file");
            }
            if (presentationInfo.IsEditable)
            {
                presentationDocument.Clone(presentationInfo.FilePath).Dispose();
            }
            presentationDocument.Dispose();
        }

        public void SaveAs(string filePath)
        {
            presentationDocument.Clone(filePath).Dispose();
            presentationDocument.Dispose();
        }
    }
}