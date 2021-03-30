using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace LST_XLS_Tool
{
    static class Program
    {
        const string OriHeader = "ORIGINAL";
        const string TLHeader = "TRANSLATION";

        [STAThread]
        static void Main(string[] args)
        {
            if (args == null || args.Length == 0) {
                Console.WriteLine("Just drag&drop your .lst files or a single .xls file to this exe");
                return;
            }

            Application.EnableVisualStyles();

            bool ToXLSMode = args.Length > 1 || args.First().ToLowerInvariant().EndsWith(".lst");

            if (ToXLSMode)
            {
                var LSTs = new(string Name, string[] Original, string[] Translation)[args.Length];
                for (int i = 0; i < args.Length; i++)
                    LSTs[i] = ReadLST(args[i]);

                Console.WriteLine($"{LSTs.Length} LST(s) Imported.");

                IWorkbook XLS = new XSSFWorkbook(XSSFWorkbookType.XLSX);
                for (int i = 0; i < LSTs.Length; i++)
                    XLS.AppendSheet(LSTs[i].Name, LSTs[i].Original, LSTs[i].Translation);


                SaveFileDialog FD = new SaveFileDialog();
                FD.Title = "Export XLSX Where?";
                FD.Filter = "All XLSX Files|*.xlsx|All Files|*.*";

                if (FD.ShowDialog() != DialogResult.OK) {
                    Console.WriteLine("Operation Aborted.");
                    return;
                }

                Console.WriteLine("Exporting XLSX...");
                using (var Stream = File.Create(FD.FileName)) {
                    XLS.Write(Stream);
                }

                Console.WriteLine("Exported.");
            } else {

                string InpXLS = args.First();
                string OutDir = InpXLS + "~\\";

                int Num = 0;
                while (Directory.Exists(OutDir))
                    OutDir = InpXLS.TrimEnd('/', '\\') + $".{Num++}~\\";


                Console.WriteLine("Exporting To: " + OutDir);

                Directory.CreateDirectory(OutDir);

                IWorkbook XLS = new XSSFWorkbook(InpXLS);

                for (int i = 0; i < XLS.NumberOfSheets; i++) {
                    var Sheet = XLS.GetSheetAt(i);

                    var Data = Sheet.ParseSheet();

                    var OutFile = Path.Combine(OutDir, "Strings-" + Data.Name + ".lst");

                    ExportLST(OutFile, Data.Originals, Data.Translations);
                }

                Console.WriteLine("Exported.");
            }
        }

        static (string Name, string[] Original, string[] Translation) ReadLST(string File) {
            Console.WriteLine("Reading: " + Path.GetFileNameWithoutExtension(File));
            var LST = new LSTParser(File);

            var Entries = LST.GetEntries().ToArray();

            return (LST.Name,
                    Entries.Select(x => x.OriginalFlags.GetFlags() + x.OriginalLine).ToArray(),
                    Entries.Select(x => x.TranslationFlags.GetFlags() + x.TranslationLine).ToArray());
        }

        static void AppendSheet(this IWorkbook Workbook, string Name, string[] Original, string[] Translations, int BeginRow = 0, int BeginCol = 0) {
            Console.WriteLine("Generating Sheet: " + Name);
            var Sheet = Workbook.CreateSheet(Name);
            
            var hRow = Sheet.CreateRow(BeginRow + 0);
            var hCelA = hRow.CreateCell(BeginCol + 0);
            var hCelB = hRow.CreateCell(BeginCol + 1);

            hCelA.SetCellValue(OriHeader);
            hCelB.SetCellValue(TLHeader);

            IFont FontStyle = Workbook.CreateFont();
            FontStyle.Color = IndexedColors.White.Index;
            FontStyle.IsBold = true;

            var CelStyle = Workbook.CreateCellStyle();
            CelStyle.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
            CelStyle.VerticalAlignment = VerticalAlignment.Center;
            CelStyle.FillForegroundColor = IndexedColors.Black.Index;
            CelStyle.FillPattern = FillPattern.SolidForeground;
            CelStyle.FillBackgroundColor = IndexedColors.Black.Index;
            CelStyle.SetFont(FontStyle);

            hCelA.CellStyle = CelStyle;
            hCelB.CellStyle = CelStyle;

            if (Original.Length != Translations.Length)
                throw new Exception("Input Translation Data Count Missmatch");

            for (int i = 0; i < Original.Length; i++) {
                var cRow = Sheet.CreateRow(BeginRow + i + 1);
                var cCelA = cRow.CreateCell(BeginCol + 0);
                var cCelB = cRow.CreateCell(BeginCol + 1);

                cCelA.SetCellValue(Original[i]);
                cCelB.SetCellValue(Translations[i]);
            }

            Sheet.AutoSizeColumn(0);
            Sheet.AutoSizeColumn(1);
        }

        static (string Name, string[] Originals, string[] Translations) ParseSheet(this ISheet Sheet)
        {

            Console.WriteLine("Parsing: " + Sheet.SheetName);

            int RowNum = 0;
            int ColNum = 0;

            var hRow = Sheet.GetRow(0);
            var hCelA = hRow.GetCell(0);
            var hCelB = hRow.GetCell(1);

            if (hCelA.StringCellValue != OriHeader || hCelB.StringCellValue != TLHeader) {
                Console.WriteLine("Finding Translation Table Header...");
                var Pos = Sheet.FindSheetStringTable();
                if (Pos.Row == -1 || Pos.Col == -1)
                    throw new Exception("Failed to find the Translation Data Table in this Sheet");

                RowNum = Pos.Row;
                ColNum = Pos.Col;
            }

            hRow = Sheet.GetRow(RowNum);
            hCelA = hRow.GetCell(ColNum + 0);
            hCelB = hRow.GetCell(ColNum + 1);

            string[] Originals = new string[Sheet.PhysicalNumberOfRows - (RowNum + 1)];
            string[] Translations = new string[Originals.Length];

            if (hCelA.StringCellValue != OriHeader || hCelB.StringCellValue != TLHeader)
                throw new Exception("Invalid Sheet Header");

            for (int i = RowNum + 1; i - (RowNum + 1) < Originals.Length; i++) {
                var Row = Sheet.GetRow(i);

                var Original    = Row.GetCell(ColNum + 0).StringCellValue;
                var Translation = Row.GetCell(ColNum + 1).StringCellValue;

                Originals[i - (RowNum + 1)] = Original.Replace("\n", LSTParser.BreakLine).Replace("\r", LSTParser.ReturnLine);
                Translations[i - (RowNum + 1)] = Translation.Replace("\n", LSTParser.BreakLine).Replace("\r", LSTParser.ReturnLine);
            }

            return (Sheet.SheetName, Originals, Translations);
        }

        static (int Row, int Col) FindSheetStringTable(this ISheet Sheet) {
            for (int y = 0; y < Sheet.PhysicalNumberOfRows; y++) {
                var CRow = Sheet.GetRow(y);
                for (int x = 0; x < CRow.LastCellNum - 1; x++) {
                    var hCelA = CRow.GetCell(x);
                    var hCelB = CRow.GetCell(x + 1);
                    if (hCelA.StringCellValue != OriHeader || hCelB.StringCellValue != TLHeader)
                        continue;
                    return (y, x);
                }
            }
            return (-1, -1);
        }

        static void ExportLST(string SaveAs, string[] Originals, string[] Translations)
        {
            if (Originals.Length != Translations.Length)
                throw new Exception("Input Translation Data Count Missmatch");

            Console.WriteLine("Exporting: " + Path.GetFileNameWithoutExtension(SaveAs));

            using (var Writer = new StreamWriter(File.Create(SaveAs), Encoding.UTF8)) {
                for (int i = 0; i < Originals.Length; i++) {
                    Writer.WriteLine(Originals[i]);
                    Writer.WriteLine(Translations[i]);
                }
            }
        }
    }
}
