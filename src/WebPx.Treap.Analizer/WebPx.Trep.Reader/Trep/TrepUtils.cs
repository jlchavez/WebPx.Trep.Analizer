using ClosedXML.Excel;
using Microsoft.VisualBasic.FileIO;
using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Net;
using Humanizer;

namespace WebPx.Trep
{
    public sealed class TrepUtils
    {
        public TrepUtils(Fuente fuente)
        {
            _fuente = fuente;
            cookies = new CookieContainer();
        }

        private Fuente _fuente;
        private CookieContainer cookies;

        private const string PrimeraVuelta = "https://primeraeleccion.trep.gt";
        private const string SegundaVuelta = "https://segundaeleccion.trep.gt";

        public string Host =>
            _fuente switch
            {
                Fuente.PrimeraVuelta => PrimeraVuelta,
                _ => SegundaVuelta
            };

        private const string proofsFilename = "GTM-pruebas.zip";

        private string EnsureFileDownloaded(string uri, string? filename = null)
        {
            filename ??= uri;
            if (File.Exists(filename))
                return filename;
            var sourcePath = $"{Host}/ext/jsonData_gtm2023/1687736870/1688050101/{uri}";
            Console.Write($"Descargando archivo '{uri}'... ");
            var content = new FileDownloader(cookies).DownloadAsync(sourcePath);
            File.WriteAllBytes(filename, content.Result);
            Console.WriteLine("descargado.");
            return filename;
        }

        public void ExtraerFechasAExcel(Eleccion[] target, string outputFilename)
        {
            var sourceFilename = $"GTM-pruebas{(_fuente == Fuente.PrimeraVuelta?"A":"B")}.zip";
            EnsureFileDownloaded(proofsFilename, sourceFilename);

            using var zip = ZipFile.OpenRead(sourceFilename);
            int t = 24585, current = 0;
            //int total = t * target.Length;

            string[] Folders(string? path = null)
            {
                var list = new List<string>();
                foreach (var rootEntry in zip.Entries)
                {
                    var zipPath = rootEntry.FullName;
                    var parts = zipPath.Split('/');
                    if (parts.Length > 1 && !list.Contains(parts[0]))
                        list.Add(parts[0]);
                }

                list.Sort();
                return list.ToArray();
            }
            string[] Files(string? path = null)
            {
                var list = new List<string>();
                var entries = zip.Entries.Where(x => x.FullName.StartsWith(path));
                foreach (var rootEntry in entries)
                {
                    var zipPath = rootEntry.FullName;
                    list.Add(zipPath);
                }
                list.Sort();
                return list.ToArray();
            }

            int doc = 0, mesa = 0;
            bool generateXLS = true;
            IXLWorksheet? mainWorksheet = null;
            int homeLineIndex = 1;

            void AddSheetLine(IXLWorksheet worksheet, ref int lineIndex, object?[]? cellValues = null, int i = 1)
            {
                lineIndex++;
                if (cellValues is not { Length: > 0 })
                    return;
                foreach (var cell in cellValues)
                {
                    if (cell != null)
                        worksheet.Cell(lineIndex, i + 1).Value = cell.ToString();
                    i++;
                }

            }

            void AddHomeLine(object?[]? cellValues = null, int i = 1) => AddSheetLine(mainWorksheet, ref homeLineIndex, cellValues);

            XLWorkbook? workbook = null;

            var columns = new object[] { "Departamento ID", "Departamento", "Municipio #", "Municipio", "Mesa", "Fecha/Hora Publicación", "Origen" };
            var indices = new int[] { 2, 1, 4, 3, 0, -2, -1 };

            if (generateXLS)
            {
                workbook ??= new();
                mainWorksheet = workbook.Worksheets.Add($"Página Principal");
                AddHomeLine();
                AddHomeLine(new object[] { "Fechas de Carga de Actas No. 4 al sistema TREP del TSE Guatemala" });
                AddHomeLine(new object[] { $"Fuente: TREP, TSE {Host}"  });
                AddHomeLine();
                AddHomeLine(new object[] { "Hoja", "Elección" });
                foreach (var index in new int[]{ 2, 3 })
                    mainWorksheet.Range($"B{index}:C{index}").Merge();
            }

            int homeStartIndex = homeLineIndex;
            var lastChar = (char)('A' + columns.Length);

            foreach (var folder in Folders())
            {
                int targetLineIndex = 1;
                doc = int.Parse(folder.Split('-').First());
                if (!target.Contains((Eleccion)(doc - 1)))
                    continue;

                var docName = doc switch
                {
                    1 => "Presidencia",
                    2 => "Diputados Listado Nacional",
                    3 => "Diputados Distritales",
                    4 => "Concejos Municipales",
                    5 => "Parlamento Centroamericano",
                };
                Console.Write($"Procesando Elección {docName}... ");
                AddHomeLine(new object[] { doc, docName });

                IXLWorksheet? worksheet = null;
                if (workbook != null)
                    worksheet = workbook.Worksheets.Add($"Elección {doc}");

                void AddLine(object?[]? cellValues = null, int i = 1) => AddSheetLine(worksheet, ref targetLineIndex, cellValues, i);

                if (worksheet != null)
                {
                    AddLine(new object[] { $"Elecciones para {docName}" });
                    AddLine(new object[] { $"Ronda: {_fuente.Humanize()} "  });
                    AddLine(new object[] { $"Fuente: TREP, TSE {Host}"  });
                    AddLine();
                    foreach (var index in new int[]{ 2, 3, 4 })
                        worksheet.Range($"B{index}:{lastChar}{index}").Merge();
                    AddLine(columns);
                }

                int tableStartIndex = targetLineIndex;

                foreach (var entry in Files(folder))
                {
                    int lineIndex = 0;

                    var sr = new StreamReader(zip.GetEntry(entry).Open());

                    while (sr.Peek() != -1)
                    {
                        var filename = entry;
                        var file = Path.GetFileName(filename);
                        file = Path.GetFileNameWithoutExtension(file);

                        var parser = new TextFieldParser(sr);
                        parser.HasFieldsEnclosedInQuotes = true;
                        parser.SetDelimiters(",");

                        while (!parser.EndOfData)
                        {
                            lineIndex++;
                            if (lineIndex == 5)
                                lineIndex++;

                            var fields = parser.ReadFields();

                            if (lineIndex <= 6)
                                continue;

                            var values = new object[indices.Length];
                            for (var valIndex = 0; valIndex < indices.Length; valIndex++)
                            {
                                var indexValue = indices[valIndex];
                                values[valIndex] = indexValue < 0 ? fields[fields.Length + indexValue] : fields[indexValue];
                            }

                            if (worksheet != null)
                                AddLine(values);

                            current++;
                        }
                    }
                }

                worksheet?.Columns().AdjustToContents();
                worksheet?.Range($"B{tableStartIndex}:{lastChar}{targetLineIndex}").CreateTable($"Table{doc}");

                Console.WriteLine();
            }

            mainWorksheet?.Range($"B{homeStartIndex}:C{homeLineIndex}").CreateTable("HomeIndex");
            mainWorksheet?.Columns().AdjustToContents();

            workbook?.SaveAs(outputFilename);
        }
    }
}