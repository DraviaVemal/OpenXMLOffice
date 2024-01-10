---
layout:
  title:
    visible: true
  description:
    visible: false
  tableOfContents:
    visible: true
  outline:
    visible: true
  pagination:
    visible: true
---

# PowerPoint

The `Powerpoint` class, a core component of the `OpenXMLOffice.Presentation` library, empowers developers to create, open, and manipulate PowerPoint (.pptx) files with ease. Whether generating new presentations or working with existing ones, this class provides a simple yet powerful interface for efficient content manipulation. Once modifications are complete, users can effortlessly save the updated presentation.

### Usage, Options and Examples

Create or open a pptx file from path

```csharp
public static CreateNew(){
    PowerPoint powerPoint = new(string.Format("../../test-{0}.pptx", DateTime.Now.ToString("yyyy-MM-dd-HH-mm-ss")), null);
    powerPoint.Save();
}

public static OpenExisting(){
    PowerPoint powerPoint = new(string.Format("../../test-{0}.pptx", DateTime.Now.ToString("yyyy-MM-dd-HH-mm-ss")),true, null);
    powerPoint.SaveAs("/NewPath/file.pptx");
}
```

Create or open a pptx object using a stream

```csharp
public static CreateUsingStream(Stream stream){
    PowerPoint powerPoint = new(stream, null);
    powerPoint.Save();
}
```

Sample using most of the exposed functions

```csharp
public static CreateNew(){
    PowerPoint powerPoint = new(string.Format("../../test-{0}.pptx", DateTime.Now.ToString("yyyy-MM-dd-HH-mm-ss")), null);
    // Add Blank Slide To the Blank Presentation
    // Return Slide Object that can be used to do slide level operation
    Slide slide = powerPoint.AddSlide(PresentationConstants.SlideLayoutType.BLANK);
    powerPoint.AddSlide(PresentationConstants.SlideLayoutType.BLANK);
    Slide slide1 = powerPoint.GetSlideByIndex(1);
    // Move the Slide Order
    powerPoint.MoveSlideByIndex(1,0);
    // Remove Slide and its content from Presentation
    powerPoint.RemoveSlideByIndex(0);
    // Save the Opened Presentation
    powerPoint.Save();
}
```

### PresentationProperties

```csharp
public class PresentationProperties
    {
        #region Public Fields

        public PresentationSettings Settings = new();

        /// <summary>
        /// TODO : Multi Theme Slide Master Support
        /// </summary>
        public Dictionary<string, PresentationSlideMaster>? SlideMasters;

        public ThemePallet Theme = new();

        #endregion Public Fields
    }
```

Direct Primary Theme for the presentation is done using ThemePallet

### TODO

* Multi Slide Master Support
* Each Slide Master Theme Support
