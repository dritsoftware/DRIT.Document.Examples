# DRIT.Document

A pure-managed, cross-platform .NET library for creating, reading, and manipulating Word documents — **DOCX**, **DOCM**, **DOC**, **ODT**, and **HTML** — with full support for **tables**, **lists**, **styles**, **charts**, **mail merge**, **track changes**, **form fields**, **footnotes and endnotes**, **comments**, **content controls**, **digital signatures** (XAdES-BES), **AES-256 encryption**, **VBA macros**, and **PDF export**.

DRIT.Document is built around a clean-room **OOXML (ECMA-376), binary DOC, and ODF implementation**: the entire stack — package model, OpenXML reader/writer, binary DOC reader/writer, ODF text pipeline, drawing model, chart hierarchy, mail merge engine, and signing ceremony — is implemented in source, with no third-party Word dependencies.

[![.NET](https://img.shields.io/badge/.NET-net48%20%7C%20netstandard2.0-blue.svg)](https://dotnet.microsoft.com/)
[![NuGet](https://img.shields.io/nuget/v/DRIT.Document.svg)](https://www.nuget.org/packages/DRIT.Document)
[![License: Commercial](https://img.shields.io/badge/license-Commercial-orange.svg)](#license)

## Public examples

Runnable examples are maintained under
[`DRIT.Document.Examples.Console`](DRIT.Document.Examples.Console/README.md).
Each example has its own console project and follows the layout used by
`DRIT.Pdf/GitHub/DRIT.Pdf.Examples.Console`.

---

## Table of Contents

- [What is DRIT.Document?](#what-is-dritdocument)
- [Features](#features)
- [Supported File Formats](#supported-file-formats)
- [Platform Independence](#platform-independence)
- [Get Started](#get-started)
  - [Create a Word Document](#create-a-word-document)
  - [Load and Iterate Document Content](#load-and-iterate-document-content)
  - [Apply Character Formatting](#apply-character-formatting)
  - [Create a Table](#create-a-table)
  - [Create a Chart](#create-a-chart)
  - [Mail Merge](#mail-merge)
  - [Find and Replace](#find-and-replace)
  - [Add Headers and Footers](#add-headers-and-footers)
  - [Track Changes](#track-changes)
  - [Export to PDF](#export-to-pdf)
  - [Encrypt a Document](#encrypt-a-document)
  - [Add a Digital Signature](#add-a-digital-signature)
  - [Create a Macro-Enabled Document](#create-a-macro-enabled-document)
- [Security](#security)
- [Repository Contents](#repository-contents)
- [Building](#building)
- [License](#license)

---

## What is DRIT.Document?

DRIT.Document is a **pure-managed .NET Word document library**. It does not depend on Microsoft Word, Microsoft Office Interop, or any third-party Word library. The entire stack — OOXML package model, OpenXML reader/writer, binary DOC (Word 97-2003) reader/writer, ODF text pipeline, drawing model, chart hierarchy, mail merge engine, and digital-signature ceremony — is implemented in source.

Key positioning:

- **Clean-room implementation.** Written from scratch against the ECMA-376 (Office Open XML), [MS-DOC], and ODF 1.2 specifications. No code copied from Open XML SDK, GemBox.Document, or any other library.
- **Full format coverage.** Read and write DOCX, DOCM, and ODT; read legacy binary DOC with preservation for round-trip fidelity; import and export HTML.
- **Complete document model.** Sections, body, paragraphs, runs, styles, tables with cell merging, lists with multi-level indentation, headers/footers, footnotes/endnotes, comments, fields, form fields, and content controls.
- **Mail merge engine.** Scalar `{{FieldName}}` replacement, `{{#Range}}` block expansion for `IEnumerable` data sources, nested collections, and `FieldMerging` / `ImageFieldMerging` callbacks.
- **Track changes.** `DocumentSettings.TrackRevisions` with a typed `RevisionCollection` and `AcceptAll` / `RejectAll`.
- **Digital signatures.** XAdES-BES signing with commitment type and signer role; append, validate, and remove signatures.
- **Cross-format conversion.** Load DOCX, save as ODT or HTML (and vice-versa); export to PDF with bookmarks and font fallback.
- **Integrated with DRIT.Pdf.** PDF export with document properties and bookmarks.
- **Integrated with DRIT.Drawing.** Advanced graphics, font support, and image format coverage.

## Features

- **Create, load, and save** documents from file or stream (DOCX, DOCM, DOC, ODT, HTML), with load options (password, encryption expectation) and cancellation support.
- **Merge & split** — combine the body content of two documents into one with conflict policies for styles and lists.
- **Cross-format conversion** — load DOCX, save as ODT or HTML (and vice-versa) with round-trip fidelity.
- **Preservation** — unknown and unmapped OpenXml elements are preserved for round-trip fidelity.
- **Text and formatting:** character formatting (font name, size, bold, italic, underline, strikethrough, color, all-caps, superscript, subscript) via `Font`; paragraph formatting (alignment, indentation, line spacing, space before/after, borders, background) via `ParagraphFormat`; built-in and custom styles via `StyleCollection`; rich text across multiple runs.
- **Breaks:** line breaks, column breaks, and page breaks in multi-column sections.
- **Sections:** add, reorder, and remove sections with section breaks; page setup (size, orientation, margins, multi-column layout, line numbering) via `PageSetup`.
- **Headers and footers:** default, first-page-only, and odd/even page headers and footers with page numbering.
- **Lists:** bullet lists and numbered lists with multi-level indentation via `List` and `ListFormat`.
- **Tables:** create tables, set cell text, configure column widths, apply borders, merge cells (horizontal and vertical), and shade header cells.
- **Bookmarks, hyperlinks, and references:** bookmarks; external URL and email hyperlinks; internal document references; footnotes and endnotes via `Document.AddFootnote` / `AddEndnote`; table of contents via `Paragraph.AppendTableOfContents`.
- **Comments and review:** reviewer comments with author, initials, and date; track changes via `DocumentSettings.TrackRevisions` with `RevisionCollection` and `AcceptAll` / `RejectAll`.
- **Fields and forms:** `DocumentFieldInfo` with `FieldInstruction` (AUTHOR, DATE, IF, and more); text, check box, and drop-down form fields; content control round-trip preservation (w:sdt).
- **Pictures, shapes, and charts:** inline pictures via `Shape.SetImage`; drawing shapes with outline, fill, and rotation; text boxes; text watermarks; 17+ chart types (Bar, Line, Pie, Scatter, Bubble, Area, Radar, Stock, Treemap, Sunburst, BoxWhisker, Funnel, Histogram, Map, Pareto, Waterfall, and 3-D variants) with titles, legends, axes, and data labels.
- **Mail merge:** scalar `{{FieldName}}` replacement; `{{#Range}}` block expansion for `IEnumerable` data sources; nested `{{#Collection}}` blocks; `FieldMerging` and `ImageFieldMerging` callbacks; `MailMergeCleanupOptions`.
- **Find and replace:** `Document.Find(text)` with `StringComparison` and `FindReplaceOptions`; `Document.Replace(old, new)` with plain-text, case-insensitive, and regex modes.
- **Restrict editing:** `DocumentSettings.Protection.StartEnforcingProtection` with granular `DocumentProtectionType` and password.
- **Write protection:** `DocumentSettings.WriteProtection.SetPassword` / `Recommended` with SHA-512 hash round-trip.
- **Encryption:** AES-256 (ECMA-376 Agile) DOCX encryption with password-protected opening.
- **Digital signatures:** XAdES-BES with commitment type and signer role; `SaveSigned`, `AppendDigitalSignature`, `ListDigitalSignatures`, and `RemoveDigitalSignatures`.
- **VBA macros:** `VbaProject` with modules (`AddModule` / `AddClass` / `AddDocument`), references, and code read/write; save as `.docm`.
- **Metadata:** built-in summary (Title, Author, Subject, Keywords, Company) and statistics; custom properties; document variables; custom XML parts.
- **Conversion:** PDF (bookmarks, document properties, font fallback) via `DocumentPdfExporter.SaveAsPdf`; HTML (flow-layout and fixed-layout); ODT; plain text via `Document.GetText()`.

## Supported File Formats

| Format | Read | Write | Specification |
|---|:---:|:---:|---|
| **DOCX** | ✅ | ✅ | ECMA-376 (Office Open XML) |
| **DOCM** | ✅ | ✅ | ECMA-376 (Macro-Enabled) |
| **DOC** | ✅ | Preservation-only | [MS-DOC] (Word 97-2003 Binary) |
| **ODT** | ✅ | ✅ | ODF 1.2 (OpenDocument Text) |
| **HTML** | ✅ | ✅ | HyperText Markup Language |
| **PDF** | — | ✅ | PDF 1.5 |

> PDF export is provided via the integrated **DRIT.Pdf** companion.

## Platform Independence

DRIT.Document is implemented in pure managed C# and targets **`net48`** and **`netstandard2.0`**, so it runs on:

- .NET Framework 4.8+
- .NET Standard 2.0 compatible runtimes (.NET Core 2.0+, .NET 5.0+)
- Windows, Linux, and macOS

Its dependencies are **DRIT.Pdf** (PDF export), **DRIT.Drawing** (graphics and font support), **System.Drawing.Common**, and **System.Text.Encoding.CodePages**. No Microsoft Word installation, no Office Interop, and no third-party Word packages are required.

---

## Get Started

The `Document` class is the entry point for creating, loading, and saving Word documents.

### Create a Word Document

```csharp
using DRIT.Document;

// Initialize a new document
var doc = new Document();
var section = doc.AddSection();

// Add content
section.AddParagraph("NovaEdge Technologies — Employee Handbook 2026");
section.AddParagraph("This handbook sets out the company's policies and guidelines.");

// Save the document
doc.Save("output.docx");
```

### Load and Iterate Document Content

```csharp
using DRIT.Document;

// Load an existing document
var doc = new Document("input.docx");

Console.WriteLine($"Paragraphs: {doc.Sections[0].Body.Paragraphs.Count}");

// Iterate all body paragraphs
foreach (var para in doc.Sections[0].Body.Paragraphs)
{
    Console.WriteLine($"[Para] \"{para.GetText()}\"");
    foreach (var run in para.Runs)
    {
        Console.WriteLine($"  [Run] \"{run.Text}\" " +
            $"Bold={run.Font.Bold} Italic={run.Font.Italic} Size={run.Font.Size}");
    }
}

// Full text extraction
Console.WriteLine(doc.GetText());
```

### Apply Character Formatting

```csharp
using DRIT.Document;

var doc = new Document();
var section = doc.AddSection();
var para = section.AddParagraph();

para.AppendRun("NovaEdge Technologies").Font.Bold = true;
para.AppendBreak(BreakType.LineBreak);

var italicRun = para.AppendRun("Employee Handbook 2026");
italicRun.Font.Italic = true;
italicRun.Font.Size = 16;

doc.Save("formatted.docx");
```

### Create a Table

```csharp
using DRIT.Document;
using DRIT.Document.Tables;

var doc = new Document();
var section = doc.AddSection();

// Create a 5-column table
var table = section.Body.AddTable(5, 5);
table.Alignment = TableAlignment.Center;

// Header row
string[] headers = { "SKU", "Product", "Category", "Unit Price", "Stock" };
for (int col = 0; col < headers.Length; col++)
{
    var cell = table.Rows[0].Cells[col];
    cell.SetText(headers[col]);
    cell.Paragraphs[0].Runs[0].Font.Bold = true;
}

// Data rows
string[][] data =
{
    new[] { "NE-SRV-001", "NebulaCore Server", "Servers", "4 899.00", "42" },
    new[] { "NE-MON-002", "ArcLight Monitor", "Displays", "649.00", "135" },
};
for (int r = 0; r < data.Length; r++)
    for (int c = 0; c < data[r].Length; c++)
        table.Rows[r + 1].Cells[c].SetText(data[r][c]);

doc.Save("table.docx");
```

### Create a Chart

```csharp
using DRIT.Document;
using DRIT.Document.Chart;

var doc = new Document();
var section = doc.AddSection();
var para = section.AddParagraph();

// Create a bar chart
var chart = new BarChart(doc);

// Set category labels
chart.Data.Categories.Add("Q1");
chart.Data.Categories.Add("Q2");
chart.Data.Categories.Add("Q3");
chart.Data.Categories.Add("Q4");

// Add series
chart.DataSeries.Add("Orion", "cache:[Q1|Q2|Q3|Q4]", "cache:[4200|5100|4800|6300]");
chart.DataSeries.Add("Vega",  "cache:[Q1|Q2|Q3|Q4]", "cache:[3100|2900|3600|4200]");
chart.DataSeries.Add("Lyra",  "cache:[Q1|Q2|Q3|Q4]", "cache:[1800|2400|3100|3900]");
chart.Data.ApplyCategoriesToSeries();

// Set the chart title and legend
chart.AddTitle();
chart.Title.FormattedText.Text = "Quarterly Revenue by Product";
chart.AddLegend();

// Append the chart to the paragraph
para.AppendChart(chart, 22, 14);

doc.Save("chart.docx");
```

### Mail Merge

```csharp
using DRIT.Document;

// Load a template with merge fields
var doc = new Document("template.docx");

// Perform merge with a dictionary data source
doc.Merge(new Dictionary<string, object>
{
    ["CustomerName"] = "Aiko Tanaka",
    ["ProductName"] = "NebulaCore Server",
    ["OrderDate"] = "2026-05-12",
    ["Quantity"] = "1",
    ["TotalPrice"] = "4 899.00"
});

doc.Save("merged.docx");
```

### Find and Replace

```csharp
using DRIT.Document;

var doc = new Document("input.docx");

// Find all occurrences
var matches = doc.Find("NovaEdge Technologies");
Console.WriteLine($"Found {matches.Count} occurrence(s).");

// Replace across the whole document
int replaced = doc.Replace("NovaEdge Technologies", "NovaEdge Technologies Ltd.");
Console.WriteLine($"Replaced {replaced} occurrence(s).");

doc.Save("updated.docx");
```

### Add Headers and Footers

```csharp
using DRIT.Document;

var doc = new Document();
var section = doc.AddSection();
section.PageSetup.DifferentFirstPageHeaderFooter = true;

// First-page header
var firstHeader = section.AddHeaderFooter(HeaderFooterType.HeaderFirst);
firstHeader.AddParagraph("NovaEdge Technologies — Employee Handbook");

// Default footer with page number
var footer = section.AddHeaderFooter(HeaderFooterType.FooterPrimary);
footer.AddParagraph("Page [page] of [numpages]");

doc.Save("headers.docx");
```

### Track Changes

```csharp
using DRIT.Document;

var doc = new Document("input.docx");

// Enable revision tracking
doc.DocumentSettings.TrackRevisions = true;

// Make an edit — it will be recorded as a revision
doc.Sections[0].Body.AddParagraph("Revised section added during review.");

// Accept or reject all tracked changes
doc.Revisions.AcceptAll();
// doc.Revisions.RejectAll();

doc.Save("tracked.docx");
```

### Export to PDF

```csharp
using DRIT.Document;
using DRIT.Document.Export.Pdf;

var doc = new Document("input.docx");

// Export entire document to PDF
doc.SaveAsPdf("output.pdf", new DocumentPdfSaveOptions
{
    IncludeBookmarks = true,
    CopyDocumentProperties = true
});

// Export specific sections only
doc.SaveAsPdf("sections.pdf", new DocumentPdfSaveOptions
{
    SectionIndexes = new[] { 0, 1 }
});
```

### Encrypt a Document

```csharp
using DRIT.Document;

var doc = new Document();
doc.Sections[0].Body.AddParagraph("Confidential content");

// Save with AES-256 encryption
doc.Save("encrypted.docx", new DocumentSaveOptions
{
    EncryptionMode = DocxEncryptionMode.Ecma376Agile,
    Password = "open-password"
});

// Open with the correct password
var loaded = new Document("encrypted.docx",
    new DocumentLoadOptions { Password = "open-password" });
Console.WriteLine($"Paragraphs: {loaded.Sections[0].Body.Paragraphs.Count}");
```

### Add a Digital Signature

```csharp
using System.Security.Cryptography.X509Certificates;
using DRIT.Document;

var doc = new Document();
doc.Sections[0].Body.AddParagraph("Signed content");

// Sign with a PFX certificate (XAdES-BES)
var certificate = new X509Certificate2("certificate.pfx", "password");
doc.SaveSigned("signed.docx", new DocxDigitalSignatureSaveOptions
{
    Certificate = certificate,
    SignerRole = "Author"
});

// Validate on reload
var signatures = Document.ListDigitalSignatures("signed.docx");
Console.WriteLine($"Signatures: {signatures.Count}");
```

### Create a Macro-Enabled Document

```csharp
using DRIT.Document;

var doc = new Document();
doc.Sections[0].Body.AddParagraph("Macro-enabled document");

// Add a VBA module
var module = doc.VbaProject.Modules.AddModule("HandbookHelpers");
module.Code = @"
Sub InsertDisclaimer()
    Selection.TypeText ""Confidential — NovaEdge Technologies""
End Sub";

// Save as .docm
doc.Save("macros.docm");

// Edit existing code
var loaded = new Document("macros.docm");
var existing = loaded.VbaProject.Modules["HandbookHelpers"];
existing.Code = existing.Code.Replace("InsertDisclaimer", "InsertConfidential");
loaded.Save("macros_edited.docm");
```

---

## Security

| Capability | Implementation |
|---|---|
| AES-256 encryption | `DocumentSaveOptions` with `EncryptionMode.Ecma376Agile` and `Password` |
| Restrict editing | `DocumentSettings.Protection.StartEnforcingProtection(type, password)` |
| Write protection | `DocumentSettings.WriteProtection.SetPassword()` / `Recommended` (SHA-512) |
| XAdES-BES signing | `DocxDigitalSignatureSaveOptions` with `Certificate`, `CommitmentType`, and `SignerRole` |
| Append signatures | `Document.AppendDigitalSignature(path, signature)` |
| Validate signatures | `Document.ListDigitalSignatures(path)` |
| Remove signatures | `Document.RemoveDigitalSignatures(path)` |

---

## Repository Contents

This repository contains **runnable console examples** for DRIT.Document. The core library is closed-source and distributed via NuGet.

| Example | Description |
|---|---|
| `GettingStarted` | Create a document with paragraphs, save, and reload. |
| `Reading` | Load an existing document and iterate paragraphs, runs, and formatting. |
| `CharacterFormatting` | Per-run font properties: bold, italic, underline, size, color, strikethrough, caps. |
| `ParagraphFormatting` | Paragraph-level settings: alignment, line spacing, indentation, and borders. |
| `Styles` | Apply built-in paragraph styles and create custom character styles. |
| `Tables` | Create a table, populate it with data, merge cells, and apply borders. |
| `Lists` | Create bullet lists and numbered lists with multi-level indentation. |
| `Charts` | Create bar, line, and pie charts from in-memory data. |
| `Pictures` | Embed an image file as an inline shape and read its bytes back. |
| `Shapes` | Add drawing shapes with outline color, fill, and rotation. |
| `TextBoxes` | Insert inline text-box shapes with body text. |
| `HeadersAndFooters` | Add default, first-page, and odd/even page headers and footers. |
| `Sections` | Add sections with section breaks and configure page setup. |
| `Bookmarks` | Insert bookmarks into paragraphs and locate them by name. |
| `Hyperlinks` | External URL and email hyperlinks, and internal document references. |
| `Footnotes` | Insert footnotes and endnotes and reference them from paragraph runs. |
| `Comments` | Attach reviewer comments to paragraph ranges with author and date. |
| `TrackChanges` | Enable revision tracking and accept or reject tracked changes. |
| `Fields` | Insert and read fields (AUTHOR, DATE, IF) with field instructions. |
| `FormFields` | Create text, check box, and drop-down form fields and read their values. |
| `ContentControls` | Round-trip preserve structured document tags (w:sdt). |
| `MailMerge` | Perform scalar field, range block, and nested collection merges. |
| `FindAndReplace` | Text search and replacement (plain-text, case-insensitive, and regex). |
| `Watermarks` | Add text watermarks with diagonal layout, color, and semi-transparency. |
| `DocumentProperties` | Read and write built-in summary properties and custom document properties. |
| `VBA` | Create a VBA module, add code, and edit existing code in a macro-enabled document. |
| `Protection` | Restrict editing with a password and apply write protection. |
| `Encryption` | DOCX password encryption (AES-256, ECMA-376 Agile). |
| `DigitalSignature` | Sign a document with a PFX certificate (XAdES-BES) and validate on reload. |
| `ConvertToPdf` | Export a document to PDF with bookmarks and document properties. |
| `ConvertToHtml` | Export a document to flow-layout and fixed-layout HTML. |
| `ConvertToOdt` | Cross-format save to OpenDocument Text. |

## Building

Each example is a standalone console project. Open the individual `.csproj` in Visual Studio 2022 (17.10+) or build from the command line.

The examples reference the **DRIT.Document NuGet package**. The first public snapshot uses `DRIT.Document 26.9.1376`; future compatible updates may use the rolling `26.*` package line. Restore and run a single example:

```bash
cd GettingStarted
dotnet add package DRIT.Document --version 26.9.1376
dotnet run
```

Or build all examples at once (if a solution file is provided):

```bash
dotnet build DRIT.Document.Examples.sln -c Release
```

The examples target `net48` and `net8.0` to match the public example projects.

## License

DRIT.Document is a **closed-source, commercial product** of DR-IT Ltd. The examples in this repository are provided for evaluation and learning. See the [product page](https://www.dritsoftware.com/netdocument) for licensing terms.

[Home](https://www.dritsoftware.com) | [Product Page](https://www.dritsoftware.com/netdocument) | [NuGet](https://www.nuget.org/packages/DRIT.Document)
