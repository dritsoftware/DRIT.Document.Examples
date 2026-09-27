using DRIT.Document;
using DRIT.Document.Examples.Shared;
using DRIT.Document.Export.Image;

var document = new Document();
document.AddSection().AddParagraph("Raster export");
document.Sections[0].Body.AddParagraph("This page is rendered as a PNG image.");

// ImageSaveOptions selects the raster format written by SaveAsImage.
document.SaveAsImage(
	ExampleSupport.OutputPath("ImageExport.png"),
	new ImageSaveOptions(ImageSaveFormat.Png));
