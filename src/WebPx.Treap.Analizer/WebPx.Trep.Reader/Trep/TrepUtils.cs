using ClosedXML.Excel;
using Microsoft.VisualBasic.FileIO;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Net;
using System.Net.Mime;
using System.Text;
using System.Text.Json.Nodes;
using Humanizer;
using System.Drawing.Imaging;
using System.Globalization;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Spreadsheet;

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
            var sourceFilename = DeterminarFuente();

            using var zip = ZipFile.OpenRead(sourceFilename);
            //int t = 24585, current = 0;
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

        private string DeterminarFuente()
        {
            var sourceFilename = $"GTM-pruebas{(_fuente == Fuente.PrimeraVuelta ? "A" : "B")}.zip";
            EnsureFileDownloaded(proofsFilename, sourceFilename);
            return sourceFilename;
        }

        public async Task DescargarActas(Eleccion[] target, string outputFilename)
        {
            var sourceFilename = DeterminarFuente();

            using var zip = ZipFile.OpenRead(sourceFilename);
            int t = 24585/*, current = 0*/;
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

            void DumpData(IXLWorksheet worksheet, int lineIndex, object?[]? cellValues = null, int i = 1)
            {
                if (cellValues is not { Length: > 0 })
                    return;
                foreach (var cell in cellValues)
                {
                    if (cell != null)
                    {
                        var c = worksheet.Cell(lineIndex, i + 1);
                        if (cell is DateTime dt)
                        {
                            c.Value = dt.ToString("dd/MM/yyyy hh:mm:ss t");
                            c.Style.NumberFormat.Format = "dd/mm/yyyy hh:mm:ss AM/PM";
                        }
                        else if (cell is int _int)
                        {
                            c.Value = _int;
                            c.Style.NumberFormat.Format = "0";
                        }
                        else
                            c.Value = cell.ToString();
                    }

                    i++;
                }

            }

            void AddSheetLine(IXLWorksheet worksheet, ref int lineIndex, object?[]? cellValues = null, int i = 1)
            {
                lineIndex++;
                DumpData(worksheet, lineIndex, cellValues, i);
            }

            void AddHomeLine(object?[]? cellValues = null, int i = 1) => AddSheetLine(mainWorksheet, ref homeLineIndex, cellValues);

            XLWorkbook? workbook = null;

            var columns = new object[] { "Departamento ID", "Departamento", "Municipio #", "Municipio", "Centro", "Mesa", "Fecha Tomada", "Fecha Cargado", "Carga", "Fecha/Hora Publicación", "Reloj Toma", "DIF A", "Cargado", "DIF B", "Origen", "Nombre Archivo", "Código Integridad", "SimpleProof" };
            var indices = new int?[] { 2, 1, 4, 3, null, 0, null, null, null, -2, null, null, null, null, -1, null, -3, null };
            var indiceFechaTomada = Array.IndexOf(columns, "Fecha Tomada");
            var indiceFechaCargado = Array.IndexOf(columns, "Fecha Cargado");
            var indiceCarga = Array.IndexOf(columns, "Carga");
            var indiceFechaPublicacion = Array.IndexOf(columns, "Fecha/Hora Publicación");
            var indiceNombreArchivo = Array.IndexOf(columns, "Nombre Archivo");
            var indiceHash = Array.IndexOf(columns, "Código Integridad");
            var indiceSimpleProof = Array.IndexOf(columns, "SimpleProof");
            var colRelojToma = Array.IndexOf(columns, "Reloj Toma");
            var colCargado = Array.IndexOf(columns, "Cargado");
            var colDifA = Array.IndexOf(columns, "DIF A");
            var colDifB = Array.IndexOf(columns, "DIF B");

            if (generateXLS)
            {
                workbook ??= new();
                mainWorksheet = workbook.Worksheets.Add($"Página Principal");
                AddHomeLine(new object[] { $"Elecciones Generales Guatemala 2023 - {_fuente.Humanize()}" });
                AddHomeLine(new object[] { "Descarga de Actas No. 4 del sistema TREP del TSE Guatemala" });
                AddHomeLine(new object[] { $"Fuente: TREP, TSE {Host}"  });
                AddHomeLine();
                AddHomeLine(new object[] { "Hoja", "Elección" });
                foreach (var index in new int[]{ 2, 3, 4 })
                    mainWorksheet.Range($"B{index}:C{index}").Merge();
            }

            int homeStartIndex = homeLineIndex;
            var lastChar = (char)('A' + columns.Length);

            Dictionary<string, JsonNode> nodes = new();

            SemaphoreSlim semaphorePool = new SemaphoreSlim(1), semaphoreExcel = new SemaphoreSlim(1);

            List<Task> _pool = new();

            string? status = null;
            TimeSpan rate = new TimeSpan(0, 0, 0, 1);
            var timer = new System.Timers.Timer(rate);
            int x = Console.CursorLeft, y = Console.CursorTop;
            timer.Elapsed += (_, _) =>
            {
                if (status != null)
                {
                    Console.SetCursorPosition(x, y);
                    Console.Write(new string(' ', status.Length));
                }

                status = $"{mesa} de {t} [{_pool.Count}]";
                Console.SetCursorPosition(x, y);
                Console.Write(status);
            };

            List<string> _existingFiles = new();

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
                
                (x, y) = (Console.CursorLeft, Console.CursorTop);
                AddHomeLine(new object[] { doc, docName });

                timer.Start();

                IXLWorksheet? worksheet = null;
                if (workbook != null)
                    worksheet = workbook.Worksheets.Add($"Elección {doc}");

                void AddLine(object?[]? cellValues = null, int i = 1) => AddSheetLine(worksheet, ref targetLineIndex, cellValues, i);

                if (worksheet != null)
                {
                    AddLine(new object[] { $"Elecciones para {docName}" });
                    AddLine(new object[] { $"Ronda: {_fuente.Humanize()} " });
                    AddLine(new object[] { $"Fuente: TREP, TSE {Host}" });
                    AddLine(new object[] { $"Hora de Cierre", new DateTime(2023, 6, 25, 18, 0, 0, DateTimeKind.Local) });
                    AddLine();
                    foreach (var index in new int[]{ 2, 3, 4 })
                        worksheet.Range($"B{index}:{lastChar}{index}").Merge();
                    AddLine(columns);

                }

                int tableStartIndex = targetLineIndex;
                DateTime? nextUpdate = null;
                int rowIndex = -1;

                foreach (var entry in Files(folder))
                {
                    int lineIndex = 0;

                    var sr = new StreamReader(zip.GetEntry(entry)!.Open());

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

                            Interlocked.Increment(ref rowIndex);

                            #region Custom
                            var codigo = fields[fields.Length - 3];
                            int centro = int.Parse(fields[5]);

                            if (string.IsNullOrEmpty(codigo))
                                continue;

                            FileDownloader? fd = null;

                            string? img = null;
                            var correlativo = 0;
                            var div = 0;
                            while (true)
                            {
                                JsonNode? jsonNode = null;
                                var jsonFilename = $"gtm2023_tc{doc}_e{correlativo}.json";
                                var localJsonFilename = $"trep\\{jsonFilename}";
                                var nodeKey = $"{doc}.{correlativo}";

                                mesa = int.Parse(fields[0]);

                                if (!nodes.ContainsKey(nodeKey))
                                {
                                    if (!File.Exists(localJsonFilename))
                                    {
                                        var path22 = $"/ext/jsonData_gtm2023/1687736870/1688050101/{jsonFilename}";
                                        var content = new FileDownloader(cookies).GetFile(path22);

                                        if (!Directory.Exists("trep"))
                                            Directory.CreateDirectory("trep");

                                        File.WriteAllText(localJsonFilename, content.Result);
                                    }
                                    var jsonString = File.ReadAllText(localJsonFilename);
                                    jsonNode = JsonNode.Parse(jsonString)!;
                                    nodes[nodeKey] = jsonNode;
                                }
                                else
                                    jsonNode = nodes[nodeKey];

                                var divs = jsonNode["divs"]!;
                                var c = divs.AsArray().Count;

                                var divNode = divs[div]!;
                                var secciones = divNode["secciones"]!;
                                var seccion = secciones.AsArray().FirstOrDefault(x => x["seccion"].ToString() == $"{centro}");
                                if (seccion == null)
                                {
                                    div++;
                                    if (div >= c)
                                    {
                                        div = 0;
                                        correlativo++;
                                        continue;
                                    }
                                    continue;
                                }

                                var casillasNode = seccion["casillas"];
                                var mesaNode = casillasNode!.AsArray().FirstOrDefault(x => x["mesa"].ToString() == $"{mesa}");
                                var imgNode = mesaNode["info"]["img"];
                                img = imgNode.ToString();
                                break;
                            }

                            var data = (doc, centro, mesa, codigo, img);

                            var folder2 = $"{doc}\\{centro}\\";
                            var path2 = $"{folder2}centro{centro:0000}.mesa{mesa:00000}.{doc}.jpg";
                            var path3 = $"{folder2}centro{centro:0000}.mesa{mesa:00000}.{doc}.txt";

                            #endregion
                            
                            var values = new object[indices.Length];
                            for (var valIndex = 0; valIndex < indices.Length; valIndex++)
                            {
                                var indexValue = indices[valIndex];
                                if (indexValue is not null)
                                    values[valIndex] = indexValue < 0 ? fields[fields.Length + indexValue.Value] : fields[indexValue.Value];
                            }

                            values[0] = Convert.ToInt32(values[0]);
                            values[2] = Convert.ToInt32(values[2]);
                            values[4] = centro;
                            values[5] = mesa;
                            values[indiceNombreArchivo] = path2;
                            values[indiceSimpleProof] = $"https://verify.simpleproof.com/TSE/P-000024/{values[13]}";

                            semaphorePool.Wait();

                            await ProcessRow( 
                                targetLineIndex + rowIndex + 1, 
                                values, 
                                folder2, 
                                path2, 
                                path3,
                                semaphoreExcel, 
                                _existingFiles,
                                indiceHash,
                                indiceFechaTomada,
                                indiceFechaCargado, 
                                worksheet,
                                colDifA,
                                colDifB,
                                colRelojToma,
                                colCargado,
                                indiceFechaPublicacion,
                                indiceCarga,
                                indiceFechaPublicacion,
                                indiceSimpleProof,
                                DumpData
                                );


                            //lock (_pool)
                            //{
                            //    _pool.Add(job);
                            //    job.ContinueWith(_ =>
                            //    {
                            //        lock (_pool)
                            //            _pool.Remove(job);
                            semaphorePool.Release();
                            //    });
                            //}
                        }
                    }
                }

                Task.WaitAll(_pool.ToArray());
                timer.Stop();

                var endLineIndex = targetLineIndex + rowIndex;

                worksheet?.Columns().AdjustToContents();
                var range1 = $"B{tableStartIndex}:{lastChar}{endLineIndex}";
                var table = worksheet?.Range(range1).CreateTable($"Table{doc}");
                var charRT = (char)('B' + colRelojToma);
                var charC = (char)('B' + colCargado);

                var range = table.Range($"{charRT}{tableStartIndex+1}:{charRT}{endLineIndex}");
                void SetConditionalFill(string formula, string bgColor)
                {
                    var fill = range.AddConditionalFormat().WhenIsTrue(formula).Fill;
                    fill.SetBackgroundColor(XLColor.FromHtml(bgColor));
                    /*fill.SetPatternColor(XLColor.FromHtml("#fff"));*/
                }

                SetConditionalFill($"=IF(${charRT}8=\"Normal\",1,0)", "#088247");
                SetConditionalFill($"=IF(${charRT}8=\"Inválido\",1,0)", "#6425b8");
                SetConditionalFill($"=IF(${charRT}8=\"Atrasado\",1,0)", "#820808");
                SetConditionalFill($"=IF(${charRT}8=\"Adelantado\",1,0)", "#b88725");
                range = table.Range($"{charC}{tableStartIndex+1}:{charC}{endLineIndex}");
                SetConditionalFill($"=IF(${charC}8=\"Normal\",1,0)", "#088247");
                SetConditionalFill($"=IF(${charC}8=\"Adelantado\",1,0)", "#b88725");

                Console.WriteLine();
            }

            mainWorksheet?.Range($"B{homeStartIndex}:C{homeLineIndex}").CreateTable("HomeIndex");
            mainWorksheet?.Columns().AdjustToContents();

            workbook?.SaveAs(outputFilename, new SaveOptions() { EvaluateFormulasBeforeSaving = true });
        }

        async Task ProcessRow(int jobRowIndex, object?[]? values2, string folder3, string targetPath, string targetPath2, SemaphoreSlim semaphoreExcel, List<string> _existingFiles,
            int indiceHash, int indiceFechaTomada, int indiceFechaCargado, IXLWorksheet worksheet, int colDifA, int colDifB, int colRelojToma, int colCargado,
            int incideFechaPublicacion, int indiceCarga, int indiceFechaPublicacion, int indiceSimpleProof, Action<IXLWorksheet,int,object?[]?,int> DumpData
            )
        {
            //Debug.WriteLine($"Task for {jobRowIndex} + {values2[5]}");
            try
            {
                if (!_existingFiles.Contains(folder3))
                {
                    if (!Directory.Exists(folder3))
                        Directory.CreateDirectory(folder3);
                    _existingFiles.Add(folder3);
                }

                DateTime? takenAt = null, receptionDate = null;
                if (File.Exists(targetPath2))
                    receptionDate = DateTime.ParseExact(await File.ReadAllTextAsync(targetPath2), "o", CultureInfo.InvariantCulture);
                else
                {
                    var attestation = await new FileDownloader(cookies).GetAttestation(values2[indiceHash].ToString());
                    await File.WriteAllBytesAsync(targetPath, attestation!.SrcFile!);
                    receptionDate = attestation.ReceptionDate;
                    await File.WriteAllTextAsync(targetPath2, receptionDate!.Value.ToString("o", CultureInfo.InvariantCulture));
                }

                using (FileStream fs = new FileStream(targetPath, FileMode.Open, FileAccess.Read))
                using (Image myImage = Image.FromStream(fs, false, false))
                {
                    try
                    {
                        PropertyItem propItem = myImage.GetPropertyItem(36867);
                        string dateTaken = exifRegex.Replace(Encoding.UTF8.GetString(propItem.Value), "-", 2);
                        takenAt = DateTime.Parse(dateTaken);
                    }
                    catch (Exception a)
                    {
                        takenAt = null;
                    }
                }

                values2[indiceFechaTomada] = takenAt;
                values2[indiceFechaCargado] = receptionDate;

                await semaphoreExcel.WaitAsync();
                try
                {
                    if (worksheet != null)
                    {
                        DumpData(worksheet, jobRowIndex, values2, 1);
                        string difACell = $"{(char)('B' + colDifA)}{jobRowIndex}",
                            difBCell = $"{(char)('B' + colDifB)}{jobRowIndex}",
                            relojTomaCell = $"{(char)('B' + colRelojToma)}{jobRowIndex}",
                            cargadoCell = $"{(char)('B' + colCargado)}{jobRowIndex}",
                            ftCell = $"{(char)('B' + indiceFechaTomada)}{jobRowIndex}",
                            fcCell = $"{(char)('B' + indiceFechaCargado)}{jobRowIndex}",
                            fpCell = $"{(char)('B' + indiceFechaPublicacion)}{jobRowIndex}",
                            cargaCell = $"{(char)('B' + indiceCarga)}{jobRowIndex}";
                        //=SI([@[Fecha Cargado]]>$C$5;"Al Cierre";"Antes")
                        //=SI(ESBLANCO([@[Fecha Tomada]]);"Inválido";SI([@[Fecha Tomada]]<$C$5;"Adelantado";SI([@[Fecha Tomada]]-[@[Fecha Cargado]]<0;"Normal";"Atrasado")))
                        string difAFormula = $"=ABS({fcCell} - {ftCell})";
                        string difBFormula = $"=ABS({fcCell} - {fpCell})";
                        
                        worksheet.Cells(difACell).FormulaA1 = difAFormula;
                        worksheet.Cells(difACell).Style.NumberFormat.Format = "hh:mm:ss";
                        worksheet.Cells(difBCell).FormulaA1 = difAFormula;
                        worksheet.Cells(difBCell).Style.NumberFormat.Format = "hh:mm:ss";
                        
                        worksheet.Cells(relojTomaCell).FormulaA1 = $"=IF(ISBLANK({ftCell}),\"Inválido\",IF({ftCell}<$C$5,\"Adelantado\",IF({ftCell}-{fcCell}<0,\"Normal\",\"Atrasado\")))";
                        worksheet.Cells(cargadoCell).FormulaA1 = $"=IF({fcCell}-{fpCell}<0,\"Adelantado\",\"Normal\")";
                        worksheet.Cells(cargaCell).FormulaA1 = $"=IF({fcCell}>$C$5,\"Antes\",\"Al Cierre\")";

                        var link = values2[indiceSimpleProof];
                        worksheet!.Cell(jobRowIndex, indiceSimpleProof + 2).SetHyperlink(new(link.ToString()));
                    }
                }
                finally
                {
                    semaphoreExcel.Release();
                }
            }
            finally
            {

            }
        }

        private static Regex exifRegex = new Regex(":");
    }
}