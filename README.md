# OSCP.SyncfusionPresentation.HtmlConversion.NET
Convert PowerPoint Presentations to/from HTML using the Syncfusion.Presentation.NET library. 

**The initial release is planned for 4 June 2023.**

Example
---
**Namespace**
```csharp
using OSCP.SyncfusionPresentation.HtmlConversion.NET;
```

**Convert PowerPoint to HTML**
```csharp
IPresentation presentation = Presentation.Open(pptxFilePath);
PresentationToHtmlConverter converter = new PresentationToHtmlConverter();
string html = converter.Convert(presentation);
System.IO.File.WriteAllText(htmlFilePath, html);
```

**Convert HTML to PowerPoint**
```csharp
string html = System.IO.File.ReadAllText(htmlFilePath);
HtmlToPresentationConverter htmlToPresentationConverter = new HtmlToPresentationConverter();
presentation = htmlToPresentationConverter.Convert(html);
presentation.Save(pptxFilePath);
presentation.Close();
```

PowerPoint Features Supported
---
| Feature | PowerPoint to HTML | HTML to PowerPoint |
|---------|:------------------:|:------------------:|
| **Slide** |
| Background Image | :ballot_box_with_check: | :ballot_box_with_check: |
| Background Solid Color | :ballot_box_with_check: | :ballot_box_with_check: |
| Background Pattern | :large_blue_diamond: | :large_blue_diamond: |
| Background Gradient | :large_blue_diamond: | :large_blue_diamond: |
| **Paragraph** |
| Bold | :ballot_box_with_check: | :ballot_box_with_check: |
| Italic | :ballot_box_with_check: | :ballot_box_with_check: |
| Underline | :ballot_box_with_check: | :ballot_box_with_check: |
| Text Shadow | :x: | :x: |
| Strikethrough | :ballot_box_with_check: | :ballot_box_with_check: |
| Double Strikethrough | :ballot_box_with_check: | :ballot_box_with_check: |
| Subscript | :ballot_box_with_check: | :ballot_box_with_check: |
| Superscript | :ballot_box_with_check: | :ballot_box_with_check: |
| All Caps | :ballot_box_with_check: | :ballot_box_with_check: |
| Small Caps | :ballot_box_with_check: | :ballot_box_with_check: |
| Highlight | :ballot_box_with_check: | :ballot_box_with_check: |
| Font | :ballot_box_with_check: | :ballot_box_with_check: |
| Font Size | :ballot_box_with_check: | :ballot_box_with_check: |
| Font Color | :ballot_box_with_check: | :ballot_box_with_check: |
| **Ordered List** |
| 1. 2. 3. | :ballot_box_with_check: | :large_blue_diamond: |
| 1) 2) 3) | :x: | :x: |
| I. II. III. | :ballot_box_with_check: | :large_blue_diamond: |
| i. ii. iii. | :ballot_box_with_check: | :large_blue_diamond: |
| A. B. C. | :ballot_box_with_check: | :large_blue_diamond: |
| a. b. c. | :ballot_box_with_check: | :large_blue_diamond: |
| a) b) c) | :x: | :x: |
| **Unordered List** |
| Defaul Bulleted List Options | :ballot_box_with_check: | :large_blue_diamond: |
| Custom Bulleted List Symbols | :white_circle: | :white_circle: |
| **Pictures** | :large_blue_diamond: | :large_blue_diamond: |
| **Tables** | | |
| Header Row | :ballot_box_with_check: | :large_blue_diamond: |
| Banded Rows | :ballot_box_with_check: | :large_blue_diamond: |
| Total Row | :large_blue_diamond: | :large_blue_diamond: |
| Banded Columns | :ballot_box_with_check: | :large_blue_diamond: |
| First Column | :large_blue_diamond: | :large_blue_diamond: |
| Last Column | :large_blue_diamond: | :large_blue_diamond: |
| Shading | :large_blue_diamond: | :large_blue_diamond: |
| Borders | :ballot_box_with_check: | :large_blue_diamond: |
| Built-in Styles | :ballot_box_with_check: | :large_blue_diamond: |
| **Effects** | :white_circle: | :white_circle: |
| **Charts** | :white_circle: | :white_circle: |
| **Master Slides** | :large_blue_diamond: | N/A |
| **Transitions** | :x: | :x: |
| **Animations** | :x: | :x: |
| **Drawing** | :x: | :x: |
| **Comments** | :white_circle: | :white_circle: |

- :white_check_mark: - In current release.
- :ballot_box_with_check: - Development complete for next release.
- :large_blue_diamond: - Development planned for next release.
- :white_circle: - Development expected in future release.
- :x: - Not planned for development.
