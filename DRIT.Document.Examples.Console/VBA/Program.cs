using DRIT.Document;
using DRIT.Document.Examples.Shared;

var document = new Document();

// Add a VBA module and assign its source code before saving as macro-enabled DOCM.
var module = document.VbaProject.Modules.AddModule("ExampleModule");
module.Code = "Sub Hello()\r\n    MsgBox \"Hello from DRIT.Document\"\r\nEnd Sub\r\n";
document.AddSection().AddParagraph("Macro-enabled document");

document.Save(ExampleSupport.OutputPath("VBA.docm"));
