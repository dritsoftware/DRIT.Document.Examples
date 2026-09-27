using DRIT.Document.Examples.Shared;
using System;
using DRIT.Document;
using DRIT.Document.Examples.Shared;

var document = new Document();
document.AddSection().AddParagraph("Welcome to NovaEdge. NovaEdge builds reliable systems.");

// Replace returns the number of matches changed in the document content.
var replacementCount = document.Replace("NovaEdge", "Meridian", StringComparison.Ordinal);
document.Sections[0].Body.AddParagraph($"Replacements: {replacementCount}");

document.Save(ExampleSupport.OutputPath("FindAndReplace.docx"));
