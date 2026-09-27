using DRIT.Document;
using DRIT.Document.Drawing;
using DRIT.Document.Examples.Shared;

var document = new Document();
document.AddSection().AddParagraph("Confidential report content");

// Configure a reusable text watermark, then attach it to the section.
var watermark = new TextWatermark(document, "CONFIDENTIAL")
{
	Semitransparent = true,
	Rotation = 315
};
document.Sections[0].SetWatermark(watermark);

document.Save(ExampleSupport.OutputPath("Watermarks.docx"));
