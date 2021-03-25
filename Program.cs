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
                    OutDir = args + $".{Num}~\\";


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

        static void AppendSheet(this IWorkbook Workbook, string Name, string[] Original, string[] Translations) {
            Console.WriteLine("Generating Sheet: " + Name);
            var Sheet = Workbook.CreateSheet(Name);
            
            var hRow = Sheet.CreateRow(0);
            var hCelA = hRow.CreateCell(0);
            var hCelB = hRow.CreateCell(1);

            hCelA.SetCellValue("Original");
            hCelB.SetCellValue("Translation");


            var CelStyle = Workbook.CreateCellStyle();
            CelStyle.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
            CelStyle.VerticalAlignment = VerticalAlignment.Center;

            hCelA.CellStyle = CelStyle;
            hCelB.CellStyle = CelStyle;

            if (Original.Length != Translations.Length)
                throw new Exception("Input Translation Data Count Missmatch");

            for (int i = 0; i < Original.Length; i++) {
                var cRow = Sheet.CreateRow(i + 1);
                var cCelA = cRow.CreateCell(0);
                var cCelB = cRow.CreateCell(1);

                cCelA.SetCellValue(Original[i]);
                cCelB.SetCellValue(Translations[i]);
            }

            Sheet.AutoSizeColumn(0);
            Sheet.AutoSizeColumn(1);
        }

        static (string Name, string[] Originals, string[] Translations) ParseSheet(this ISheet Sheet)
        {

            Console.WriteLine("Parsing: " + Sheet.SheetName);

            string[] Originals = new string[Sheet.PhysicalNumberOfRows - 1];
            string[] Translations = new string[Sheet.PhysicalNumberOfRows - 1];

            var hRow = Sheet.GetRow(0);
            var hCelA = hRow.GetCell(0);
            var hCelB = hRow.GetCell(1);

            if (hCelA.StringCellValue != "Original" || hCelB.StringCellValue != "Translation")
                throw new Exception($"Failed to find the header in the {Sheet.SheetName} sheet");

            for (int i = 1; i < Originals.Length; i++) {
                var Row = Sheet.GetRow(i);

                var Original    = Row.GetCell(0).StringCellValue;
                var Translation = Row.GetCell(1).StringCellValue;

                Originals[i - 1] = Original;
                Translations[i - 1] = Translation;
            }

            return (Sheet.SheetName, Originals, Translations);
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
