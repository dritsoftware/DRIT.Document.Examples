using DRIT.Document;
using DRIT.Document.Examples.Shared;

var document = new Document();
var heading = document.AddSection().AddParagraph("Project overview");

// StyleName selects a built-in Word style without requiring a style XML object.
heading.ParagraphFormat.StyleName = "Heading1";
document.Sections[0].Body.AddParagraph("This paragraph follows the heading.");

document.Save(ExampleSupport.OutputPath("Styles.docx"));
