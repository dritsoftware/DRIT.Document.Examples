using DRIT.Document.Examples.Shared;
using DRIT.Document;
using DRIT.Document.Examples.Shared;

var document = new Document();
document.AddSection().AddParagraph("Inventory");

// Create a fixed-size table and fill its cells through the row and cell collections.
var table = document.Sections[0].Body.AddTable(3, 3);
table.Rows[0].Cells[0].SetText("SKU");
table.Rows[0].Cells[1].SetText("Product");
table.Rows[0].Cells[2].SetText("Stock");
table.Rows[1].Cells[0].SetText("NE-SRV-001");
table.Rows[1].Cells[1].SetText("NebulaCore Server");
table.Rows[1].Cells[2].SetText("42");
table.Rows[2].Cells[0].SetText("NE-MON-002");
table.Rows[2].Cells[1].SetText("ArcLight Monitor");
table.Rows[2].Cells[2].SetText("135");

document.Save(ExampleSupport.OutputPath("Tables.docx"));
