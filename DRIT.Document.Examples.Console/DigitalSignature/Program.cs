using System;
using System.Security.Cryptography.X509Certificates;
using DRIT.Document;
using DRIT.Document.Examples.Shared;

var certificatePath = ExampleSupport.InputPath("signing-certificate.pfx");
if (!System.IO.File.Exists(certificatePath))
{
	throw new InvalidOperationException(
		$"Add a signing certificate at '{certificatePath}' before running this example.");
}

// Load the certificate that will provide the package signature.
using var certificate = new X509Certificate2(certificatePath, "example-password");
var document = new Document();
document.AddSection().AddParagraph("Digitally signed report");

// SaveSigned writes the DOCX package together with an OPC digital signature.
document.SaveSigned(
	ExampleSupport.OutputPath("DigitalSignature.docx"),
	new DocxDigitalSignatureSaveOptions
	{
		Certificate = certificate,
		CommitmentType = "http://schemas.openxmlformats.org/package/2006/digital-signature/commitment-type/proof-of-creation",
		SignerRole = "Author"
	});
