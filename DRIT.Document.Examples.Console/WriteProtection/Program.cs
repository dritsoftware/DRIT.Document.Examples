using DRIT.Document;
using DRIT.Document.Examples.Shared;

var document = new Document();
document.AddSection().AddParagraph("This document is recommended to open as read-only.");

// Write protection recommends read-only opening while retaining editable content.
document.DocumentSettings.WriteProtection.StartEnforcingWriteProtection("example-password");

document.Save(ExampleSupport.OutputPath("WriteProtection.docx"));
