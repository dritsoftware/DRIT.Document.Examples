using System;
using System.IO;

namespace DRIT.Document.Examples.Shared
{
    internal static class ExampleSupport
    {
        internal static string OutputPath(string fileName)
        {
            var directory = Path.Combine(Environment.CurrentDirectory, "Out");
            Directory.CreateDirectory(directory);
            return Path.Combine(directory, fileName);
        }

        internal static string InputPath(string fileName)
        {
            var directory = Path.Combine(Environment.CurrentDirectory, "In");
            Directory.CreateDirectory(directory);
            return Path.Combine(directory, fileName);
        }

        internal static byte[] OnePixelPng()
        {
            return Convert.FromBase64String(
                "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=");
        }

        internal static string EnsureOnePixelPng()
        {
            var path = InputPath("product-banner.png");
            if (!File.Exists(path))
            {
                File.WriteAllBytes(path, OnePixelPng());
            }

            return path;
        }
    }
}
