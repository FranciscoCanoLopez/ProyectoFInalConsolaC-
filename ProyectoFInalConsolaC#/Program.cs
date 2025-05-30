using Spectre.Console;
using System.Net.Http.Headers;
using System.Text.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ProyectoFInalConsolaC_;
using System.Reflection.PortableExecutable;
using static System.Runtime.InteropServices.JavaScript.JSType;
using System.Globalization;

internal class Program
{
    static List<Jugadores> datosActuales = new List<Jugadores>();

    private static async Task Main(string[] args)
    {
        Console.OutputEncoding = Encoding.UTF8;

        List<CsvRow> datosCSV = new();
        List<Movie> datosAPI = new();


        while (true)
        {
            MostrarMenu();
            var opcion = Console.ReadKey(true).Key;

            if (opcion == ConsoleKey.Escape)
            {
                AnsiConsole.MarkupLine("[bold green]Gracias por usar la aplicación. ¡Hasta luego![/]");
                return;
            }

            if (opcion == ConsoleKey.A)
            {
                datosCSV = CargarDesdeCSV();
                MostrarSubmenuCSV(datosCSV);
            }
            else if (opcion == ConsoleKey.B)
            {
                datosAPI = await CargarDesdeAPI();
                MostrarSubmenuAPI(datosAPI);
            }
            else
            {
                AnsiConsole.MarkupLine("[red]Opción inválida. Intenta nuevamente.[/]");

            }
        }

        static void MostrarMenu()
        {
            AnsiConsole.Clear();
            var titulo = new FigletText("Opciones").Color(Color.Magenta1);
            AnsiConsole.Write(titulo);

            var contenido = new Markup(
                "\n[bold underline white]Seleccione el origen de datos[/]\n\n" +
                "[bold magenta]A[/]: [white]Archivo CSV[/]\n" +
                "[bold magenta]B[/]: [white]Datos desde TheMovieDB API[/]\n" +
                "[bold magenta]ESC[/]: [white]Salir[/]");

            var panel = new Panel(contenido)
            {
                Border = BoxBorder.Rounded,
                Padding = new Padding(2, 1),
                Header = new PanelHeader("[bold purple_2]Menú Principal[/]", Justify.Center),
            };

            AnsiConsole.Write(panel);
        }
        static void MostrarSubmenuCSV(List<CsvRow> datos)
        {
            while (true)
            {
                AnsiConsole.Clear();
                var submenu = new SelectionPrompt<string>()
                    .Title("[bold]Selecciona la vista de los datos CSV:[/]")
                    .AddChoices("Tabla", "Árbol", "Estadísticas", "Volver al menú principal");

                var seleccion = AnsiConsole.Prompt(submenu);

                switch (seleccion)
                {
                    case "Tabla": MostrarTablaCSV(datos); break;
                    case "Árbol": MostrarVistaArbolCSV(datos); break;
                    case "Estadísticas": MostrarEstadisticasCSV(datos); break;
                    case "[green]Exportar[/]": MostrarMenuExportacion(datos); break;
                    case "Volver al menú principal": return;
                }
            }
        }
        static void MostrarSubmenuAPI(List<Movie> datos)
        {
            while (true)
            {
                AnsiConsole.Clear();
                var submenu = new SelectionPrompt<string>()
                    .Title("[bold]Selecciona la vista de los datos API:[/]")
                    .AddChoices("Tabla", "Árbol", "Estadísticas", "Volver al menú principal");

                var seleccion = AnsiConsole.Prompt(submenu);

                switch (seleccion)
                {
                    case "Tabla": MostrarTablaAPI(datos); break;
                    case "Árbol": MostrarArbolAPI(datos); break;
                    case "Estadísticas": MostrarEstadisticasAPI(datos); break;
                    case "[green]Exportar[/]": MostrarMenuExportacionApi(datos); break;
                    case "Volver al menú principal": return;
                }
            }
        }

        static List<CsvRow> CargarDesdeCSV()
        {
            var rutaArchivo = @"C:\Users\volub\Downloads\archive\stats_football_players.csv";
            var filas = new List<CsvRow>();

            if (!File.Exists(rutaArchivo))
            {
                AnsiConsole.MarkupLine($"[red]El archivo '{rutaArchivo}' no existe.[/]");
                return filas;
            }

            try
            {
                using var reader = new StreamReader(rutaArchivo);
                var headers = reader.ReadLine()?.Split(',');

                if (headers == null || headers.Length == 0)
                {
                    AnsiConsole.MarkupLine("[red]No se encontraron encabezados en el archivo.[/]");
                    return filas;
                }

                while (!reader.EndOfStream)
                {
                    var line = reader.ReadLine();
                    if (string.IsNullOrWhiteSpace(line)) continue;

                    var values = line.Split(',');

                    var row = new CsvRow();
                    for (int i = 0; i < headers.Length && i < values.Length; i++)
                    {
                        row.Fields[headers[i].Trim()] = values[i].Trim();
                    }

                    filas.Add(row);
                }
            }
            catch (Exception ex)
            {
                AnsiConsole.MarkupLine($"[red]Error al leer el archivo: {ex.Message}[/]");
            }

            return filas;
        }

        static async Task<List<Movie>> CargarDesdeAPI()
        {
            var httpClient = new HttpClient();
            httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer",
                "eyJhbGciOiJIUzI1NiJ9.eyJhdWQiOiJkM2E5NDU5MmNjM2E5NDkwMWNhYjI2N" +
                "zNmZWFhZGM4MiIsIm5iZiI6MTc0ODI4ODIzMy42NjQsInN1YiI6IjY4MzRjMmU5NjQwZTA1YjQyOGI2" +
                "YmEwMiIsInNjb3BlcyI6WyJhcGlfcmVhZCJdLCJ2ZXJzaW9uIjoxfQ.ArYNRf2GqMDkot9TCd_5323QEN-XHpMVXfp1PxoNu4A");

            var url = "https://api.themoviedb.org/3/movie/popular?language=es-ES&page=1";

            try
            {
                var response = await httpClient.GetStringAsync(url);
                var json = JsonDocument.Parse(response);

                var peliculas = json.RootElement
                    .GetProperty("results")
                    .EnumerateArray()
                    .Select(m => new Movie
                    {
                        Title = m.GetProperty("title").GetString()!,
                        VoteAverage = m.GetProperty("vote_average").GetDecimal(),
                        VoteCount = m.GetProperty("vote_count").GetInt32(),
                        Popularity = m.GetProperty("popularity").GetDecimal()
                    })
                    .ToList();

                return peliculas;
            }
            catch (Exception ex)
            {
                AnsiConsole.MarkupLine($"[red]Error al obtener datos de la API: {ex.Message}[/]");
                return new List<Movie>();
            }
        }

        static void MostrarTablaAPI(List<Movie> movies)
        {
            const int pageSize = 10;
            int page = 0;
            int totalPages = (int)Math.Ceiling((double)movies.Count / pageSize);

            while (true)
            {
                AnsiConsole.Clear();
                var table = new Table().Border(TableBorder.Rounded);
                table.AddColumn("[yellow]Nombre de la Película[/]");

                var paginated = movies
                    .Skip(page * pageSize)
                    .Take(pageSize)
                    .ToList();

                foreach (var movie in paginated)
                {
                    table.AddRow($"[white]{movie.Title}[/]");
                }

                AnsiConsole.Write(table);
                AnsiConsole.MarkupLine($"[grey]Página {page + 1} de {totalPages}[/]");
                AnsiConsole.MarkupLine("[blue]Use N para siguiente página, P para anterior, Esc para salir.[/]");

                var key = Console.ReadKey(true).Key;

                if (key == ConsoleKey.N && page < totalPages - 1)
                    page++;
                else if (key == ConsoleKey.P && page > 0)
                    page--;
                else if (key == ConsoleKey.Escape)
                    break;
            }
        }
        static void MostrarArbolAPI(List<Movie> movies)
        {
            foreach (var movie in movies)
            {
                var tree = new Tree($"[bold yellow]{movie.Title}[/]").Style(Style.Parse("blue"));
                tree.AddNode($"[white]Popularidad[/]: [cyan]{movie.Popularity}[/]");
                tree.AddNode($"[white]Votos promedio[/]: [cyan]{movie.VoteAverage}[/]");
                tree.AddNode($"[white]Número de votos[/]: [cyan]{movie.VoteCount}[/]");
                AnsiConsole.Write(tree);
                AnsiConsole.WriteLine();
            }
            AnsiConsole.MarkupLine("[blue]Presione una tecla para continuar...[/]");
            Console.ReadKey(true);
        }

        static void MostrarEstadisticasAPI(List<Movie> movies)
        {
            int index = 0;

            while (true)
            {
                AnsiConsole.Clear();

                if (movies.Count == 0)
                {
                    AnsiConsole.MarkupLine("[red]No hay datos para mostrar.[/]");
                    return;
                }

                var movie = movies[index];

                var chart = new BarChart()
                    .Width(60)
                    .Label($"[bold yellow]{movie.Title}[/]")
                    .CenterLabel()
                    .AddItem("Popularidad", (float)movie.Popularity, Color.Blue)
                    .AddItem("Voto Promedio", (float)movie.VoteAverage, Color.Green)
                    .AddItem("Votos", movie.VoteCount, Color.Orange1);

                AnsiConsole.Write(chart);

                AnsiConsole.MarkupLine($"\n[grey]Película {index + 1} de {movies.Count}[/]");
                AnsiConsole.MarkupLine("[blue]Use ← → para cambiar película. Esc para volver.[/]");

                var key = Console.ReadKey(true).Key;

                if (key == ConsoleKey.RightArrow)
                    index = (index + 1) % movies.Count;
                else if (key == ConsoleKey.LeftArrow)
                    index = (index - 1 + movies.Count) % movies.Count;
                else if (key == ConsoleKey.Escape)
                    return;
            }
        }

        static string MostrarTablaCSV(List<CsvRow> rows)
        {
            string[] columnasMostrar = { "Season", "League", "Team", "Player", "Nation", "Position", "Age", "Match Played", "Goals", "Assists" };
            var columnasValidas = columnasMostrar.Where(c => rows.Any(r => r.Fields.ContainsKey(c))).ToArray();

            const int pageSize = 10;
            int page = 0;
            int totalPages = (int)Math.Ceiling((double)rows.Count / pageSize);

            while (true)
            {
                AnsiConsole.Clear();

                var opciones = new Panel("[bold magenta]Vistas disponibles:[/] [green]A[/]: Tabla  [green]B[/]: Árbol  [green]C[/]: Estadísticas  [green]ESC[/]: Salir")
                {
                    Border = BoxBorder.Rounded,
                    Padding = new Padding(1, 0),
                };
                AnsiConsole.Write(opciones);

                var table = new Table().Border(TableBorder.Rounded);

                foreach (var col in columnasValidas)
                    table.AddColumn($"[bold yellow]{col}[/]");

                var paginated = rows.Skip(page * pageSize).Take(pageSize);

                foreach (var row in paginated)
                {
                    var cells = columnasValidas.Select(h => $"[white]{(row.Fields.ContainsKey(h) ? row.Fields[h] : "")}[/]").ToArray();
                    table.AddRow(cells);
                }

                AnsiConsole.Write(table);
                AnsiConsole.MarkupLine($"[grey]Página {page + 1} de {totalPages}[/]");
                AnsiConsole.MarkupLine("[blue]Use ↑ ↓ para navegar. A/B/C para cambiar vista. Esc para volver al menú.[/]");

                var key = Console.ReadKey(true).Key;

                if (key == ConsoleKey.DownArrow && page < totalPages - 1)
                    page++;
                else if (key == ConsoleKey.UpArrow && page > 0)
                    page--;
                else if (key == ConsoleKey.B) return "B";
                else if (key == ConsoleKey.C) return "C";
                else if (key == ConsoleKey.A) return "A";
                else if (key == ConsoleKey.Escape) return "Salir";
            }
        }

        static string MostrarVistaArbolCSV(List<CsvRow> rows)
        {
            const int pageSize = 1;
            int page = 0;
            int totalPages = (int)Math.Ceiling((double)rows.Count / pageSize);

            while (true)
            {
                AnsiConsole.Clear();

                var opciones = new Panel("[bold magenta]Vistas disponibles:[/] [green]A[/]: Tabla  [green]B[/]: Árbol  [green]C[/]: Estadísticas  [green]ESC[/]: Salir")
                {
                    Border = BoxBorder.Rounded,
                    Padding = new Padding(1, 0),
                };
                AnsiConsole.Write(opciones);

                var paginated = rows.Skip(page * pageSize).Take(pageSize).ToList();

                foreach (var jugador in paginated)
                {
                    if (!jugador.Fields.ContainsKey("Player")) continue;

                    var root = new Tree($"[bold yellow]{jugador.Fields["Player"]}[/]").Style(Style.Parse("blue"));
                    foreach (var kvp in jugador.Fields)
                    {
                        if (kvp.Key == "Player") continue;
                        root.AddNode($"[white]{kvp.Key}[/]: [cyan]{kvp.Value}[/]");
                    }
                    AnsiConsole.Write(root);
                    AnsiConsole.WriteLine();
                }

                AnsiConsole.MarkupLine($"[grey]Página {page + 1} de {totalPages}[/]");
                AnsiConsole.MarkupLine("[blue]Use ↑ ↓ para navegar. A/B/C para cambiar vista. Esc para volver al menú.[/]");

                var key = Console.ReadKey(true).Key;

                if (key == ConsoleKey.DownArrow && page < totalPages - 1)
                    page++;
                else if (key == ConsoleKey.UpArrow && page > 0)
                    page--;
                else if (key == ConsoleKey.A) return "A";
                else if (key == ConsoleKey.C) return "C";
                else if (key == ConsoleKey.B) return "B";
                else if (key == ConsoleKey.Escape) return "Salir";
            }
        }

        static string MostrarEstadisticasCSV(List<CsvRow> rows)
        {
            int index = 0;

            while (true)
            {
                AnsiConsole.Clear();

                var opciones = new Panel("[bold magenta]Vistas disponibles:[/] [green]A[/]: Tabla  [green]B[/]: Árbol  [green]C[/]: Estadísticas  [green]ESC[/]: Salir")
                {
                    Border = BoxBorder.Rounded,
                    Padding = new Padding(1, 0),
                };
                AnsiConsole.Write(opciones);

                if (rows.Count == 0)
                {
                    AnsiConsole.MarkupLine("[red]No hay datos para mostrar.[/]");
                    return "Salir";
                }

                var jugador = rows[index];

                if (!jugador.Fields.ContainsKey("Player"))
                {
                    index = (index + 1) % rows.Count;
                    continue;
                }

                var nombre = jugador.Fields["Player"];
                int partidos = 0, goles = 0, asistencias = 0, amarillas = 0, rojas = 0;

                int.TryParse(jugador.Fields.GetValueOrDefault("Match Played", "0"), out partidos);
                int.TryParse(jugador.Fields.GetValueOrDefault("Goals", "0"), out goles);
                int.TryParse(jugador.Fields.GetValueOrDefault("Assists", "0"), out asistencias);
                int.TryParse(jugador.Fields.GetValueOrDefault("Yellow Cards", "0"), out amarillas);
                int.TryParse(jugador.Fields.GetValueOrDefault("Red Cards", "0"), out rojas);

                var chart = new BarChart()
                    .Width(60)
                    .Label($"[bold yellow]Estadísticas de {nombre}[/]")
                    .CenterLabel()
                    .AddItem("Partidos", partidos, Color.Blue)
                    .AddItem("Goles", goles, Color.Green)
                    .AddItem("Asistencias", asistencias, Color.Orange1)
                    .AddItem("Amarillas", amarillas, Color.Yellow)
                    .AddItem("Rojas", rojas, Color.Red);


                AnsiConsole.Write(chart);

                AnsiConsole.MarkupLine($"\n[grey]Jugador {index + 1} de {rows.Count}[/]");
                AnsiConsole.MarkupLine("[blue]Use ← → para cambiar jugador. A/B/C para cambiar vista. Esc para volver al menú.[/]");

                var key = Console.ReadKey(true).Key;

                if (key == ConsoleKey.RightArrow)
                    index = (index + 1) % rows.Count;
                else if (key == ConsoleKey.LeftArrow)
                    index = (index - 1 + rows.Count) % rows.Count;
                else if (key == ConsoleKey.A) return "A";
                else if (key == ConsoleKey.B) return "B";
                else if (key == ConsoleKey.C) return "C";
                else if (key == ConsoleKey.Escape) return "Salir";
            }


        }
        static void MostrarMenuExportacion(List<CsvRow> jugadores)
        {
            if (jugadores == null || jugadores.Count == 0)
            {
                AnsiConsole.MarkupLine("[red]No hay datos para exportar.[/]");
                return;
            }

            var formato = AnsiConsole.Prompt(
                new SelectionPrompt<string>()
                    .Title("Selecciona el [green]formato de exportación[/]:")
                    .AddChoices("CSV", "TXT", "XML", "JSON")
            );

            string ruta = AnsiConsole.Ask<string>("Ingresa la ruta y nombre del archivo (sin extensión):");

            try
            {
                switch (formato.ToUpper())
                {
                    case "CSV":
                        ExportarACSV(jugadores, ruta + ".csv");
                        break;
                    case "TXT":
                        ExportarATXT(jugadores, ruta + ".txt");
                        break;
                    case "XML":
                        ExportarAXML(jugadores, ruta + ".xml");
                        break;
                    case "JSON":
                        ExportarAJSON(jugadores, ruta + ".json");
                        break;
                }

                AnsiConsole.MarkupLine($"[green]Datos exportados correctamente como {formato}.[/]");
            }
            catch (Exception ex)
            {
                AnsiConsole.MarkupLine($"[red]Error al exportar: {ex.Message}[/]");
            }
        }
        // Exportar a CSV
        static void ExportarACSV(List<CsvRow> datos, string ruta)
        {
            try
            {
                if (datos.Count == 0)
                {
                    AnsiConsole.MarkupLine("[red]No hay datos para exportar.[/]");
                    return;
                }

                var headers = datos[0].Fields.Keys.ToList();

                using var writer = new StreamWriter(ruta, false, Encoding.UTF8);
                writer.WriteLine(string.Join(",", headers));

                foreach (var fila in datos)
                {
                    var valores = headers.Select(h => fila.Fields.ContainsKey(h) ? fila.Fields[h] : "");
                    writer.WriteLine(string.Join(",", valores));
                }

                AnsiConsole.MarkupLine($"[green]Datos exportados exitosamente a CSV: {ruta}[/]");
            }
            catch (Exception ex)
            {
                AnsiConsole.MarkupLine($"[red]Error al exportar a CSV: {ex.Message}[/]");
            }
        }

        // Exportar a TXT (formato plano)
        static void ExportarATXT(List<CsvRow> datos, string ruta)
        {
            try
            {
                using var writer = new StreamWriter(ruta, false, Encoding.UTF8);
                foreach (var fila in datos)
                {
                    foreach (var kvp in fila.Fields)
                        writer.WriteLine($"{kvp.Key}: {kvp.Value}");
                    writer.WriteLine(new string('-', 40));
                }

                AnsiConsole.MarkupLine($"[green]Datos exportados exitosamente a TXT: {ruta}[/]");
            }
            catch (Exception ex)
            {
                AnsiConsole.MarkupLine($"[red]Error al exportar a TXT: {ex.Message}[/]");
            }
        }

        // Exportar a JSON
        static void ExportarAJSON(List<CsvRow> datos, string ruta)
        {
            try
            {
                var json = JsonSerializer.Serialize(datos.Select(d => d.Fields), new JsonSerializerOptions { WriteIndented = true });
                File.WriteAllText(ruta, json, Encoding.UTF8);
                AnsiConsole.MarkupLine($"[green]Datos exportados exitosamente a JSON: {ruta}[/]");
            }
            catch (Exception ex)
            {
                AnsiConsole.MarkupLine($"[red]Error al exportar a JSON: {ex.Message}[/]");
            }
        }

        // Exportar a XML
        static void ExportarAXML(List<CsvRow> datos, string ruta)
        {
            try
            {
                var xml = new StringBuilder();
                xml.AppendLine("<Jugadores>");

                foreach (var fila in datos)
                {
                    xml.AppendLine("  <Jugador>");
                    foreach (var kvp in fila.Fields)
                        xml.AppendLine($"    <{kvp.Key}>{System.Security.SecurityElement.Escape(kvp.Value)}</{kvp.Key}>");
                    xml.AppendLine("  </Jugador>");
                }

                xml.AppendLine("</Jugadores>");
                File.WriteAllText(ruta, xml.ToString(), Encoding.UTF8);

                AnsiConsole.MarkupLine($"[green]Datos exportados exitosamente a XML: {ruta}[/]");
            }
            catch (Exception ex)
            {
                AnsiConsole.MarkupLine($"[red]Error al exportar a XML: {ex.Message}[/]");
            }
        }
        static List<CsvRow> ConvertMoviesToCsvRows(List<Movie> movies)
        {
            return movies.Select(m => new CsvRow
            {
                Fields = new Dictionary<string, string>
                {
                    ["Title"] = m.Title!,
                    ["VoteAverage"] = m.VoteAverage.ToString(CultureInfo.InvariantCulture),
                    ["VoteCount"] = m.VoteCount.ToString(),
                    ["Popularity"] = m.Popularity.ToString(CultureInfo.InvariantCulture)
                }
            }).ToList();
        }
        static void MostrarMenuExportacionApi(List<Movie> movies)
        {
            var csvRows = ConvertMoviesToCsvRows(movies);

            while (true)
            {
                Console.Clear();
                AnsiConsole.MarkupLine("[bold cyan]Menú de Exportación (Datos desde API)[/]");
                Console.WriteLine("1. Exportar a CSV");
                Console.WriteLine("2. Exportar a JSON");
                Console.WriteLine("3. Exportar a TXT");
                Console.WriteLine("4. Exportar a XML");
                Console.WriteLine("5. Volver al menú principal");

                Console.Write("\nSeleccione una opción: ");
                string opcion = Console.ReadLine()!;

                switch (opcion)
                {
                    case "1":
                        ExportarACSV(csvRows, "export_api.csv");
                        break;
                    case "2":
                        ExportarAJSON(csvRows, "export_api.json");
                        break;
                    case "3":
                        ExportarATXT(csvRows, "export_api.txt");
                        break;
                    case "4":
                        ExportarAXML(csvRows, "export_api.xml");
                        break;
                    case "5":
                        return;
                    default:
                        Console.WriteLine("Opción inválida. Presione una tecla para continuar...");
                        Console.ReadKey();
                        break;
                }

                Console.WriteLine("Exportación realizada. Presione una tecla para continuar...");
                Console.ReadKey();
            }
        }
    }
}