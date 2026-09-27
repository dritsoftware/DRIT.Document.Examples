using DRIT.Document;
using DRIT.Document.Chart;
using DRIT.Document.Chart.Office;
using DRIT.Document.Examples.Shared;

var document = new Document();
var chart = new LineChart(document)
{
	Grouping = Grouping.Standard,
	DataSource = "Category,Value\nAlpha,10\nBeta,25\nGamma,18"
};
chart.AddLegend();

// Charts are drawing objects, so attach the configured chart to a paragraph.
document.AddSection().AddParagraph("Quarterly results");
document.Sections[0].Body.AddParagraph().AppendChart(chart);
document.Save(ExampleSupport.OutputPath("Charts.docx"));
