using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text.RegularExpressions;
using System.Xml;

namespace EDTRFeldolgozo
{
    static class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("EDTR feldolgozó");

            // Egy parameter legyen, ami egy ZIP file az EDTR-bol
            if (args.Length < 1)
            {
                Console.WriteLine("Használat: EDTRFeldolgozo.exe edtr_csomag_neve.zip");
                Console.WriteLine("Például: EDTRFeldolgozo.exe Koznevelesi_Muvelodesi_es_Ifjusagi_Bizottsag_2020_10_15.zip");
                return;
            }

            string fileName = args[0];

            // Letezik-e a parameterkent megadott allomany
            if (!File.Exists(fileName))
            {
                Console.WriteLine($"Az állomány nem létezik: {fileName}!");
                return;
            }

            // Az allomany zip kiterjesztesu-e
            if (Path.GetExtension(fileName).ToLowerInvariant() != ".zip")
            {
                Console.WriteLine($"Az állomány nem ZIP kiterjesztésű: {fileName}!");
                return;
            }

            Console.WriteLine($"Állomány kitömörítése: {Path.GetFileName(fileName)}...");
            if (!ExtractFile(fileName))
            {
                return;
            }
            Console.WriteLine("\tKitömörítés kész!");

            // Hol fogjuk feldolgozni
            string targetDir = Path.Combine(Path.GetDirectoryName(fileName), Path.GetFileNameWithoutExtension(fileName));

            // EDTR a konyvtar datum reszeben alahuzas helyett kotojelet hasznal
            targetDir = ReplaceLastOccurrence(targetDir, "_", "-");
            targetDir = ReplaceLastOccurrence(targetDir, "_", "-");

            Console.WriteLine("Zárt ülések mellékleteinek törlése...");
            if (!RemoveClosedTopics(targetDir))
            {
                return;
            }
            Console.WriteLine("\tTörlés kész!");

            // HTM file keresese
            string htmFile = Directory.GetFiles(targetDir).FirstOrDefault();
            if (String.IsNullOrEmpty(htmFile))
            {
                Console.WriteLine("Nem találok HTML állományt a csomagban!");
                return;
            }

            Console.WriteLine($"HTML állomány feldolgozása: {Path.GetFileName(htmFile)}");
            if (!ProcessHtmFile(htmFile))
            {
                return;
            }

            Console.WriteLine("HTML feldolgozás kész!");

            Console.WriteLine("PDF állomány eltávolítása...");
            string pdfFile = htmFile.Replace(".htm", ".pdf");
            if(File.Exists(pdfFile))
            {
                File.Delete(pdfFile);
            }

            Console.WriteLine("PDF állomány eltávolítás kész!");

            Console.WriteLine("Folyamat befejezve");
        }

        private static string ReplaceLastOccurrence(string str, string toReplace, string replacement)
        {
            return Regex.Replace(str, $@"^(.*){toReplace}(.*?)$", $"$1{replacement}$2");
        }

        private static bool RemoveClosedTopics(string sessionDir)
        {
            try
            {
                Directory.GetDirectories(sessionDir).Where(x => x.EndsWith("_zart")).ToList().ForEach((dir) =>
                {
                    Directory.Delete(dir, true);
                });
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Probléma a zárt témák eltávolításánál: {ex.Message}");
                return false;
            }

            return true;
        }

        private static bool ProcessHtmFile(string fileName)
        {
            // Valamiert dupla HTML head es body van
            string[] content = File.ReadAllLines(fileName);
            string fixedContent = String.Join("\r\n", content, 8, content.Length - 10)
                .Replace("&nbsp;", "")
                .Replace("<pd4ml:page.break>", "")
                .Replace("$[page]", "");

            XmlDocument xmlDoc = new XmlDocument();

            XmlReaderSettings settings = new XmlReaderSettings { NameTable = new NameTable() };
            XmlNamespaceManager xmlns = new XmlNamespaceManager(settings.NameTable);
            xmlns.AddNamespace("pd4ml", "");
            XmlParserContext context = new XmlParserContext(null, xmlns, "", XmlSpace.Default);
            XmlReader reader = XmlReader.Create(new StringReader(fixedContent), settings, context);
            xmlDoc.Load(reader);

            // Kikeressuk az osszes napirendi pontot
            XmlNodeList napiRendek = xmlDoc.SelectNodes("//table");

            Console.WriteLine("\tNapirendi pontok feldolgozása...");
            string pont;
            string dir;
            string[] files;

            foreach (XmlNode napiRend in napiRendek)
            {
                // "x. napirendi pont" kiturasa es konyvtarnevve alakitasa
                pont = napiRend.PreviousSibling.PreviousSibling.PreviousSibling.PreviousSibling.OuterXml;
                pont = pont.Replace("<div style=\"text-align: right;\">", "")
                    .Replace("</div>", "")
                    .Replace("<b>", "")
                    .Replace("</b>", "")
                    .Replace("\r\n", "")
                    .Replace("\n", "")
                    .Replace(".", "")
                    .Replace(" ", "_")
                    .Trim()
                    .ToLowerInvariant();

                if (pont.Contains("(zárt_ülés)"))
                {
                    pont = pont.Replace("__", "")
                        .Replace("(zárt_ülés)", "");

                    Console.WriteLine($"\t\t{pont} zárt ülés - eltávolítás");

                    napiRend.PreviousSibling.RemoveAll();
                    napiRend.PreviousSibling.PreviousSibling.RemoveAll();
                    napiRend.PreviousSibling.PreviousSibling.PreviousSibling.RemoveAll();
                    napiRend.PreviousSibling.PreviousSibling.PreviousSibling.PreviousSibling.RemoveAll();
                    napiRend.PreviousSibling.PreviousSibling.PreviousSibling.PreviousSibling.PreviousSibling.RemoveAll();
                    napiRend.NextSibling.RemoveAll();

                    napiRend.RemoveAll();
                    continue;
                }

                Console.WriteLine($"\t\t{pont} feldolgozása...");

                // Ez most eleg paraszt megoldas lesz a nulla kezelesere
                dir = Path.Combine(Path.GetDirectoryName(fileName), pont);

                if (Directory.Exists(dir))
                {
                    files = Directory.GetFiles(dir);

                    foreach(string file in files)
                    {
                        InsertLink(xmlDoc, napiRend, pont, file);
                    }
                }
                else
                {
                    dir = Path.Combine(Path.GetDirectoryName(fileName), "0" + pont);

                    if (Directory.Exists(dir))
                    {
                        files = Directory.GetFiles(dir);

                        foreach (string file in files)
                        {
                            InsertLink(xmlDoc, napiRend, pont, file);
                        }
                    }
                    else
                    {
                        Console.WriteLine($"\t\t\tNem találok {pont} könyvtárat!");
                    }
                }

                Console.WriteLine($"\t\t{pont} feldolgozás kész");
            }

            // Nem siman elmentjuk, mert elotte meg kicsit okositunk rajta
            string finalXml = xmlDoc.InnerXml;

            finalXml = "<!DOCTYPE HTML PUBLIC \" -//W3C//DTD HTML 4.0 Transitional//EN\">\r\n" + finalXml;
            finalXml = finalXml.Replace("<!-- Rendezés a viewban --><div></div><div></div><div></div><br /><table></table><br />", "");
            File.WriteAllText(fileName, finalXml);
            Console.WriteLine("\tNapirendi pontok feldolgozás kész");

            return true;
        }

        private static void InsertLink(XmlDocument xmlDoc, XmlNode napiRend, string pont, string fileName)
        {
            Console.WriteLine($"\t\t\tMelléklet beillesztése: {Path.GetFileName(fileName)}");
            XmlElement xmlElement = xmlDoc.CreateElement("a");
            xmlElement.SetAttribute("href", $"{pont}/{Path.GetFileName(fileName)}");
            xmlElement.InnerText = Path.GetFileName(fileName);
            napiRend.NextSibling.AppendChild(xmlElement);
            xmlElement = xmlDoc.CreateElement("br");
            napiRend.NextSibling.AppendChild(xmlElement);
        }

        private static bool ExtractFile(string fileName)
        {
            string targetDir = Path.GetDirectoryName(fileName);

            try
            {
                using (MemoryStream sourceStream = new MemoryStream())
                {
                    using (FileStream sourceZipFile = new FileStream(fileName, FileMode.Open))
                    {
                        sourceZipFile.CopyTo(sourceStream);

                        using (ZipArchive srcArchive = new ZipArchive(sourceStream, ZipArchiveMode.Read))
                        {
                            srcArchive.ExtractToDirectory(targetDir);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Hiba kicsomagolás közben: {ex.Message}");
                return false;
            }

            return true;
        }
    }
}