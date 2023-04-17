# OSCP.SyncfusionPresentation.HtmlConversion.NET
Convert PowerPoint Presentations to/from HTML using the Syncfusion.Presentation.NET library.

Example
---
**Namespace**
```csharp
using OSCP.SyncfusionPresentation.HtmlConversion.NET;
```

**Convert PowerPoint to HTML**
```csharp
IPresentation presentation = Presentation.Open(filepath);
PresentationToHtmlConverter converter = new PresentationToHtmlConverter();
string html = converter.Convert(presentation);
```

PowerPoint Features Supported In Conversion
---
| Feature | PowerPoint to HTML | HTML to PowerPoint |
|---------|:------------------:|:------------------:|
| **Slide** |
| Background Image | :ballot_box_with_check: | :large_blue_diamond: |
| **Paragraph** |
| Bold | :ballot_box_with_check: | :large_blue_diamond: |
| Italic | :ballot_box_with_check: | :large_blue_diamond: |
| Underline | :ballot_box_with_check: | :large_blue_diamond: |
| Strikethrough | :ballot_box_with_check: | :large_blue_diamond: |
| Double Strikethrough | :ballot_box_with_check: | :large_blue_diamond: |
| Subscript | :ballot_box_with_check: | :large_blue_diamond: |
| Superscript | :ballot_box_with_check: | :large_blue_diamond: |
| All Caps | :ballot_box_with_check: | :large_blue_diamond: |
| Small Caps | :ballot_box_with_check: | :large_blue_diamond: |
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
| **Tables** | :large_blue_diamond: | :large_blue_diamond: |
| **Charts** | :white_circle: | :white_circle: |

- :white_check_mark: - In current release.
- :ballot_box_with_check: - Development complete for next release.
- :ballot_box_with_check: - Development planned for next release.
- :white_circle: - No development planned yet.
- :x: - Not planned for development.
