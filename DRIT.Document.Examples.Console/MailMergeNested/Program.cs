using DRIT.Document;
using DRIT.Document.Examples.Shared;

var document = new Document();

// Nested regions expand one collection inside another collection.
document.Sections[0].Body.AddParagraph(
	"{{#Projects}}{{Name}}:\n" +
	"{{#Tasks}}- {{Title}}\n{{/Tasks}}{{/Projects}}");

// Execute the merge with a project containing its own task records.
document.MailMerge.Execute(new
{
	Projects = new[]
	{
		new
		{
			Name = "Migration",
			Tasks = new[] { new { Title = "Plan" }, new { Title = "Ship" } }
		}
	}
});

document.Save(ExampleSupport.OutputPath("MailMergeNested.docx"));
