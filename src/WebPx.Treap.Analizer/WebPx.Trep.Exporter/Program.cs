using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Math;
using Humanizer;

namespace WebPx.Trep.Exporter
{
    internal class Program
    {
        private static void ScreenHeader()
        {
            Console.Clear();
            Console.WriteLine("WebPx Exportador del T.R.E.P. - Elecciones Guatemala 2023!");
            Console.WriteLine("----------------------------------------------------------");
            Console.WriteLine();
        }

        private static Fuente fuente = Fuente.PrimeraVuelta;
        private static TrepUtils? utils = null;

        static void Main(string[] args)
        {
            var continuar = true;
            ScreenHeader();
            Console.WriteLine("Elija la elección:");
            Console.WriteLine();
            Console.WriteLine("[P]rimera Vuelta");
            Console.WriteLine("[S]egunda Vuelta");
            Console.WriteLine();
            Console.WriteLine("[C]ancelar y salir");
            Console.WriteLine();
            do
            {
                var key = Console.ReadKey(true);
                switch (key.Key)
                {
                    case ConsoleKey.P: (fuente, continuar) = (Fuente.PrimeraVuelta, false); break;
                    case ConsoleKey.S: (fuente, continuar) = (Fuente.SegundaVuelta, false); break;
                    case ConsoleKey.C: return;
                }
            } while (continuar);

            utils = new TrepUtils(fuente);
            bool salir = false, render = true;
            do
            {
                if (render)
                {
                    ScreenHeader();
                    Console.WriteLine($"Elija la operación que desea hacer ({fuente.Humanize()})");
                    Console.WriteLine();
                    Console.WriteLine("[F] - Extraer fecha y hora de publicación de actas");
                    Console.WriteLine();
                    Console.WriteLine("[S] - Salir");
                    Console.WriteLine();
                    render = false;
                }

                var key = Console.ReadKey(true);
                switch (key.Key)
                {
                    case ConsoleKey.F:
                        ExtraerFechaHora();
                        render = true;
                        break;
                    case ConsoleKey.S:
                        salir = true;
                        break;
                }
            } while (!salir);

            Console.WriteLine();
            Console.WriteLine("Gracias por utilizar esta herramienta.");
        }

        private static void ExtraerFechaHora()
        {
            ScreenHeader();
            bool salir = false, render = true, process = false;

            Eleccion[] elecciones = fuente switch
            {
                Fuente.PrimeraVuelta => new[] { Eleccion.Presidencia, Eleccion.DiputadosListadoNacional, Eleccion.DiputadosDistritales, Eleccion.CorporaciónMunicipales, Eleccion.DiputadosParlamentoCentroamericano },
                Fuente.SegundaVuelta => new[] { Eleccion.Presidencia },
                _ => new Eleccion[] { }
            };

            Eleccion[] target = elecciones;
            if (elecciones?.Length > 0)
            {
                do
                {
                    process = false;
                    if (render)
                    {
                        ScreenHeader();
                        Console.WriteLine($"Elija la operación que desea hacer ({fuente.Humanize()})");
                        Console.WriteLine();
                        int i = 1;
                        foreach (var eleccion in elecciones)
                        {
                            Console.WriteLine($"[{i++}] - {eleccion.Humanize()}");
                        }

                        Console.WriteLine("[T] - Todas las anteriores");
                        Console.WriteLine();
                        Console.WriteLine("[C] - Cancelar");
                        Console.WriteLine();
                        render = false;
                    }

                    var key = Console.ReadKey(true);
                    switch (key.Key)
                    {
                        case ConsoleKey.C: return;
                        case ConsoleKey.T: (target, process) = (elecciones, true); break;
                        default:
                            if (!Char.IsNumber(key.KeyChar))
                                break;
                            int valor = key.KeyChar - '0';                       
                            if (valor > 0 && valor <= elecciones.Length)
                                (target, process) = (new[] { elecciones[valor - 1] }, true);
                            break;
                    }

                    if (process)
                    {
                        var date = DateTime.Now.ToString("dd_MM_yyyy_HH_mm_ss_zz");
                        var targetFilename = $"{fuente}_{date}.xlsx";
                        utils!.ExtraerFechasAExcel(target, targetFilename);

                        Console.WriteLine($"Archivo generado: {targetFilename}");
                        Console.WriteLine();
                        Console.Write("Desea abrir el archivo? (S/N) ");
                        var valido = false;
                        do
                        {
                            var key2 = Console.ReadKey(true);
                            switch (key2.Key)
                            {
                                case ConsoleKey.S:
                                    OpenFile(targetFilename);
                                    valido = true; 
                                    break;
                                case ConsoleKey.N:
                                    valido = true;
                                    break;
                            }
                        } while (!valido);
                        Console.WriteLine();
                        salir = true;
                    }
                } while (!salir);
            }
        }

        public static void OpenFile(string path)
        {
            var startInfo = new System.Diagnostics.ProcessStartInfo
            {
                WindowStyle = System.Diagnostics.ProcessWindowStyle.Normal,
                FileName = path,
                RedirectStandardInput = false,
                UseShellExecute = true
            };
        }

        public static void OpenFolder(string path)
        {
            var startInfo = new System.Diagnostics.ProcessStartInfo
            {
                WorkingDirectory = path,
                WindowStyle = System.Diagnostics.ProcessWindowStyle.Normal,
                FileName = "cmd.exe",
                RedirectStandardInput = true,
                UseShellExecute = false
            };
        }
    }
}