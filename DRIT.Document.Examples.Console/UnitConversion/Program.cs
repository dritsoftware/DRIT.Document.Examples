using System;
using DRIT.Document;
using DRIT.Document.Examples.Shared;
using DRIT.Document.Office;

// Convert a physical measurement explicitly by naming both source and target units.
var points = UnitConverter.Convert(MeasurementUnit.Centimeter, MeasurementUnit.Point, 2.54);
Console.WriteLine($"2.54 cm = {points} points");

// Pixel conversion accepts a DPI so the result is deterministic for the target device.
var pixels = UnitConverter.Convert(MeasurementUnit.Point, MeasurementUnit.Pixel, 72, 96);
Console.WriteLine($"72 points at 96 DPI = {pixels} pixels");
