using DRIT.Document;
using DRIT.Document.Examples.Shared;

var document = new Document();
document.AddSection().AddParagraph("Only comments are allowed in this protected example.");

// StartEnforcingProtection writes the editing restriction and its password hash.
document.DocumentSettings.Protection.StartEnforcingProtection(
	DocumentProtectionType.Comments,
	"example-password");

document.Save(ExampleSupport.OutputPath("RestrictEditing.docx"));
