using Aspose.Words;
using System;
using System.IO;
using System.Reflection;
using Aspose.Words.Tables;
using Newtonsoft.Json;
using PaperSeparator.DataTransfer;
using PaperSeparator.Enum;

namespace PaperSeparator
{
    public static class FunctionBase
    {
        private static void SetLicense()
        {
            License awLic = new License();
            string Key =
            "PExpY2Vuc2U+DQogIDxEYXRhPg0KICAgIDxMaWNlbnNlZFRvPkFzcG9zZSBTY290bGFuZCB" +
            "UZWFtPC9MaWNlbnNlZFRvPg0KICAgIDxFbWFpbFRvPmJpbGx5Lmx1bmRpZUBhc3Bvc2UuY2" +
            "9tPC9FbWFpbFRvPg0KICAgIDxMaWNlbnNlVHlwZT5EZXZlbG9wZXIgT0VNPC9MaWNlbnNlV" +
            "HlwZT4NCiAgICA8TGljZW5zZU5vdGU+TGltaXRlZCB0byAxIGRldmVsb3BlciwgdW5saW1p" +
            "dGVkIHBoeXNpY2FsIGxvY2F0aW9uczwvTGljZW5zZU5vdGU+DQogICAgPE9yZGVySUQ+MTQ" +
            "wNDA4MDUyMzI0PC9PcmRlcklEPg0KICAgIDxVc2VySUQ+OTQyMzY8L1VzZXJJRD4NCiAgIC" +
            "A8T0VNPlRoaXMgaXMgYSByZWRpc3RyaWJ1dGFibGUgbGljZW5zZTwvT0VNPg0KICAgIDxQc" +
            "m9kdWN0cz4NCiAgICAgIDxQcm9kdWN0PkFzcG9zZS5Ub3RhbCBmb3IgLk5FVDwvUHJvZHVj" +
            "dD4NCiAgICA8L1Byb2R1Y3RzPg0KICAgIDxFZGl0aW9uVHlwZT5FbnRlcnByaXNlPC9FZGl" +
            "0aW9uVHlwZT4NCiAgICA8U2VyaWFsTnVtYmVyPjlhNTk1NDdjLTQxZjAtNDI4Yi1iYTcyLT" +
            "djNDM2OGYxNTFkNzwvU2VyaWFsTnVtYmVyPg0KICAgIDxTdWJzY3JpcHRpb25FeHBpcnk+M" +
            "jAxNTEyMzE8L1N1YnNjcmlwdGlvbkV4cGlyeT4NCiAgICA8TGljZW5zZVZlcnNpb24+My4w" +
            "PC9MaWNlbnNlVmVyc2lvbj4NCiAgICA8TGljZW5zZUluc3RydWN0aW9ucz5odHRwOi8vd3d" +
            "3LmFzcG9zZS5jb20vY29ycG9yYXRlL3B1cmNoYXNlL2xpY2Vuc2UtaW5zdHJ1Y3Rpb25zLm" +
            "FzcHg8L0xpY2Vuc2VJbnN0cnVjdGlvbnM+DQogIDwvRGF0YT4NCiAgPFNpZ25hdHVyZT5GT" +
            "zNQSHNibGdEdDhGNTlzTVQxbDFhbXlpOXFrMlY2RThkUWtJUDdMZFRKU3hEaWJORUZ1MXpP" +
            "aW5RYnFGZkt2L3J1dHR2Y3hvUk9rYzF0VWUwRHRPNmNQMVpmNkowVmVtZ1NZOGkvTFpFQ1R" +
            "Hc3pScUpWUVJaME1vVm5CaHVQQUprNWVsaTdmaFZjRjhoV2QzRTRYUTNMemZtSkN1YWoyTk" +
            "V0ZVJpNUhyZmc9PC9TaWduYXR1cmU+DQo8L0xpY2Vuc2U+";
            using (Stream stream = new MemoryStream(Convert.FromBase64String(Key)))
            {
                //awLic.SetLicense(@"D:\VietBank\Portal\Library\Aspose.Words.lic");
                awLic.SetLicense(stream);
            }
        }
        private static Configure LoadConfig(string path)
        {
            Configure config = new Configure();
            using (StreamReader sr = new StreamReader(path))
            {
                string json = sr.ReadToEnd();
                config = JsonConvert.DeserializeObject<Configure>(json);
            }
            return config;
        }

        private static DocumentBuilder CreateDocumentBuilder(Document doc, Configure configure)
        {
            PaperSize paperSize;
            HeightRule heightRule;
            ParagraphAlignment paragraphAlign;
            CellVerticalAlignment cellAlignment;
            switch (configure.PageSetupPaperSize)
            {
                case ConfigureEnum.A4:
                    paperSize = PaperSize.A4;
                    break;
                case ConfigureEnum.Letter:
                    paperSize = PaperSize.Letter;
                    break;
                case ConfigureEnum.A3:
                    paperSize = PaperSize.A3;
                    break;
                case ConfigureEnum.A5:
                    paperSize = PaperSize.A5;
                    break;
                case ConfigureEnum.B4:
                    paperSize = PaperSize.B4;
                    break;
                default:
                    paperSize = PaperSize.B5;
                    break;
            }

            switch (configure.RowFormatHeightRule)
            {
                case ConfigureEnum.Exactly:
                    heightRule = HeightRule.Exactly;
                    break;
                case ConfigureEnum.AtLeast:
                    heightRule = HeightRule.AtLeast;
                    break;
                default:
                    heightRule = HeightRule.Auto;
                    break;
            }

            switch (configure.ParagraphFormatAlignment)
            {
                case ConfigureEnum.Center:
                    paragraphAlign = ParagraphAlignment.Center;
                    break;
                case ConfigureEnum.Justify:
                    paragraphAlign = ParagraphAlignment.Justify;
                    break;
                case ConfigureEnum.Right:
                    paragraphAlign = ParagraphAlignment.Right;
                    break;
                default:
                    paragraphAlign = ParagraphAlignment.Left;
                    break;
            }

            switch (configure.CellFormatVerticalAlignment)
            {
                case ConfigureEnum.Center:
                    cellAlignment = CellVerticalAlignment.Center;
                    break;
                case ConfigureEnum.Bottom:
                    cellAlignment = CellVerticalAlignment.Bottom;
                    break;
                default:
                    cellAlignment = CellVerticalAlignment.Top;
                    break;
            }

            DocumentBuilder builder = new DocumentBuilder(doc)
            {
                PageSetup =
                {
                    Orientation = configure.PageSetupOrientation == ConfigureEnum.Portrait ? Orientation.Portrait : Orientation.Landscape,
                    PaperSize = paperSize,
                    LeftMargin = configure.PageSetupLeftMargin,
                    RightMargin = configure.PageSetupRightMargin,
                    TopMargin = configure.PageSetupTopMargin,
                    BottomMargin = configure.PageSetupBottomMargin
                },
            };

            builder.CellFormat.TopPadding = configure.CellFormatTopPadding;
            builder.CellFormat.BottomPadding = configure.CellFormatBottomPadding;
            builder.CellFormat.LeftPadding = configure.CellFormatLeftPadding;
            builder.CellFormat.RightPadding = configure.CellFormatRightPadding;

            //table.LeftIndent = 0;//20.0;
            builder.RowFormat.Height = configure.RowFormatHeight;
            builder.RowFormat.HeightRule = heightRule;
            builder.ParagraphFormat.Alignment = paragraphAlign;
            builder.CellFormat.VerticalAlignment = cellAlignment;
            //builder.CellFormat.Width = 200.0;
            builder.Font.Size = configure.FontSize;
            builder.Font.Name = configure.FontName;
            builder.Font.Bold = configure.FontBold != ConfigureEnum.False;
            return builder;
        }

        private static void ReStartApplication(bool isSuccess, string error = "")
        {
            if (isSuccess)
                Console.WriteLine("Execute Success.\n");
            else
            {
                Console.WriteLine("Execute fail.\n Look at error below:\n");
                Console.WriteLine(error);
            }
            Console.WriteLine("Do you want to re-start? (Y/N)");
            var info = Console.ReadKey();
            if (info.Key == ConsoleKey.Y)
            {
                string fileName = Assembly.GetExecutingAssembly().Location;
                System.Diagnostics.Process.Start(fileName);
            }
        }

        public static void Separator()
        {
            try
            {
                SetLicense();
                Configure configure = LoadConfig("config.json");
                int column = configure.Column;

                string[] lines = File.ReadAllLines(configure.DataPath);
                int wordCount = lines.Length;
                if (wordCount == 0)
                {
                    Console.WriteLine("Data file is empty. Fill it some data.");
                    return;
                }
                int numRows = wordCount / column;
                int modRows = wordCount % column;

                Document doc = new Document();
                DocumentBuilder builder = CreateDocumentBuilder(doc, configure);
                Table table = builder.StartTable();
                int count = 0;
                for (int i = 0; i < column; i++)
                {
                    builder.InsertCell();
                    builder.Write(lines[count++]);
                }
                builder.EndRow();

                table.AllowAutoFit = configure.TableAllowAutoFit != ConfigureEnum.False;
                for (int i = 0; i < numRows - 1; i++)
                {
                    for (int j = 0; j < column; j++)
                    {
                        builder.InsertCell();
                        builder.Write(lines[count++]);
                    }
                    builder.EndRow();
                }
                for (int i = 0; i < modRows; i++)
                {
                    builder.InsertCell();
                    builder.Write(lines[count++]);
                }
                builder.EndTable();
                string dataDir = configure.SavePath + configure.FileName;
                doc.Save(dataDir);
                ReStartApplication(true);
            }
            catch (Exception e)
            {
                ReStartApplication(false, e.ToString());
                //throw;
            }
        }
    }
}
