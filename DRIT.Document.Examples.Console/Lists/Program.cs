using DRIT.Document;
using DRIT.Document.Examples.Shared;
using DRIT.Document.Lists;

var document = new Document();
var first = document.AddSection().AddParagraph("NebulaCore Server");
var second = document.Sections[0].Body.AddParagraph("ArcLight Monitor");

// Share one list definition between paragraphs to keep numbering consistent.
var list = new List(ListTemplate.BulletDefault);
first.ParagraphFormat.ListFormat = new ListFormat { List = list, ListLevelNumber = 0 };
second.ParagraphFormat.ListFormat = new ListFormat { List = list, ListLevelNumber = 0 };

document.Save(ExampleSupport.OutputPath("Lists.docx"));
