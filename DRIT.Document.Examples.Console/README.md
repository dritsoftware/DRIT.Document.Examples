# DRIT.Document Examples

Each example is an independent console project. The layout follows the
`DRIT.Pdf/GitHub/DRIT.Pdf.Examples.Console` convention:

```text
DRIT.Document.Examples.Console/
  GettingStarted/GettingStarted.csproj
  GettingStarted/Program.cs
  ...
```

By default projects consume the released `DRIT.Document` NuGet package. The
first public snapshot uses `DRIT.Document 26.9.1376`; future compatible 26.x
snapshots may use the rolling `26.*` package line. Run an example from the
public repository root with:

```powershell
dotnet run --project .\DRIT.Document.Examples.Console\GettingStarted\GettingStarted.csproj -f net8.0
```

Outputs are written to the workspace `Out/` directory when launched from the
repository root. Input fixtures are created under an example's `In/` directory
when needed. The examples use original NovaEdge, Meridian, and Orion sample
content and do not depend on Microsoft Word.

## Catalog

| # | Example | Focus |
|---:|---|---|
| 1 | `GettingStarted` | Create, save, reload |
| 2 | `CharacterFormatting` | Run-level formatting |
| 3 | `ParagraphFormatting` | Alignment, indents, spacing |
| 4 | `Lists` | List workflow scaffold |
| 5 | `Tables` | Tables and cell text |
| 6 | `FindAndReplace` | Text replacement |
| 7 | `Reading` | Load and extract text |
| 8 | `Styles` | Style workflow scaffold |
| 9 | `HeadersAndFooters` | Header/footer stories |
| 10 | `Sections` | Multiple sections |
| 11 | `PageSetup` | Orientation and margins |
| 12 | `Breaks` | Page breaks |
| 13 | `DocumentProperties` | Built-in metadata |
| 14 | `MailMerge` | Scalar and repeating merges |
| 15 | `MailMergeRanges` | Repeating data regions |
| 16 | `MailMergeNested` | Nested regions |
| 17 | `Templates` | Template token replacement |
| 18 | `Combining` | Append documents |
| 19 | `Charts` | Chart workflow scaffold |
| 20 | `Pictures` | Picture workflow scaffold |
| 21 | `Shapes` | Shape workflow scaffold |
| 22 | `TextBoxes` | Text-box workflow scaffold |
| 23 | `Watermarks` | Watermark workflow scaffold |
| 24 | `BookmarksAndHyperlinks` | Reference workflow scaffold |
| 25 | `FootnotesAndEndnotes` | Footnotes and endnotes |
| 26 | `Comments` | Review workflow scaffold |
| 27 | `Fields` | Field workflow scaffold |
| 28 | `FormFields` | Form workflow scaffold |
| 29 | `ContentControls` | Content-control workflow scaffold |
| 30 | `TableOfContents` | TOC workflow scaffold |
| 31 | `TrackChanges` | Revision settings |
| 32 | `RightToLeftText` | RTL workflow scaffold |
| 33 | `Hyphenation` | Hyphenation workflow scaffold |
| 34 | `ViewOptions` | View and zoom |
| 35 | `UnitConversion` | Unit workflow scaffold |
| 36 | `Encryption` | Encryption workflow scaffold |
| 37 | `DigitalSignature` | Signature workflow scaffold |
| 38 | `WriteProtection` | Write-protection workflow scaffold |
| 39 | `RestrictEditing` | Editing restrictions scaffold |
| 40 | `VBA` | Macro workflow scaffold |
| 41 | `ConvertToPdf` | PDF export |
| 42 | `ConvertToHtml` | HTML export |
| 43 | `ConvertToOdt` | ODT export |
| 44 | `ConvertFromHtml` | HTML import workflow scaffold |
| 45 | `ConvertFromDoc` | Legacy DOC workflow scaffold |
| 46 | `ExtractText` | Plain-text extraction |
| 47 | `MailMergeImages` | Binary image merge |
| 48 | `MailMergeAnalysis` | Read-only template analysis |
| 49 | `MailMergeFormatting` | Culture-aware merge formatting |
| 50 | `DocumentSettings` | Settings round trip |
| 51 | `FootnoteOptions` | Note workflow scaffold |
| 52 | `OfficeMath` | OMML equation |
| 53 | `OleAndVbaReferences` | Macro package workflow scaffold |
| 54 | `TextFormats` | Text-format workflow scaffold |
| 55 | `ImageExport` | PNG raster export |
| 56 | `PdfSecurity` | PDF security workflow scaffold |

The entries marked “workflow scaffold” are runnable, source-compatible starting
points for fixture-dependent APIs whose reproducible input packages are not
checked into this public example tree. Their `Program.cs` entry points are kept
independent so those workflows can grow without changing the repository layout.
