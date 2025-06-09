Imports Spectre.Console
Imports System.Net.Http.Headers
Imports System.Text.Json
Imports System
Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports System.Threading.Tasks
Imports ProyectoFInalConsolaC_
Imports System.Reflection.PortableExecutable
Imports System.Runtime.InteropServices.JavaScript.JSType
Imports System.Globalization
Imports QuestPDF.Fluent
Imports QuestPDF.Helpers
Imports CsvHelper
Imports System.Data.SqlClient
Imports QuestPDF.Infrastructure
Imports MimeKit

Friend Class Program
    Private Shared Async Function Main(args As String()) As System.Threading.Tasks.Task
        System.Console.OutputEncoding = System.Text.Encoding.UTF8
                ''' Cannot convert LocalDeclarationStatementSyntax, System.InvalidCastException: No se puede convertir un objeto de tipo 'Microsoft.CodeAnalysis.VisualBasic.Syntax.EmptyStatementSyntax' al tipo 'Microsoft.CodeAnalysis.VisualBasic.Syntax.ExpressionSyntax'.
'''    en ICSharpCode.CodeConverter.VB.CommonConversions.RemodelVariableDeclaration(VariableDeclarationSyntax declaration)
'''    en ICSharpCode.CodeConverter.VB.MethodBodyExecutableStatementVisitor.VisitLocalDeclarationStatement(LocalDeclarationStatementSyntax node)
'''    en Microsoft.CodeAnalysis.CSharp.CSharpSyntaxVisitor`1.Visit(SyntaxNode node)
'''    en ICSharpCode.CodeConverter.VB.CommentConvertingMethodBodyVisitor.DefaultVisit(SyntaxNode node)
''' 
''' Input:
''' 
'''         System.Collections.Generic.List<ProyectoFInalConsolaC_.CsvRow> datosCSV = new();
''' 
''' 
                ''' Cannot convert LocalDeclarationStatementSyntax, System.InvalidCastException: No se puede convertir un objeto de tipo 'Microsoft.CodeAnalysis.VisualBasic.Syntax.EmptyStatementSyntax' al tipo 'Microsoft.CodeAnalysis.VisualBasic.Syntax.ExpressionSyntax'.
'''    en ICSharpCode.CodeConverter.VB.CommonConversions.RemodelVariableDeclaration(VariableDeclarationSyntax declaration)
'''    en ICSharpCode.CodeConverter.VB.MethodBodyExecutableStatementVisitor.VisitLocalDeclarationStatement(LocalDeclarationStatementSyntax node)
'''    en Microsoft.CodeAnalysis.CSharp.CSharpSyntaxVisitor`1.Visit(SyntaxNode node)
'''    en ICSharpCode.CodeConverter.VB.CommentConvertingMethodBodyVisitor.DefaultVisit(SyntaxNode node)
''' 
''' Input:
'''         System.Collections.Generic.List<ProyectoFInalConsolaC_.Movie> datosAPI = new();
''' 
''' 

        QuestPDF.Settings.License = QuestPDF.Infrastructure.LicenseType.Community

        While True
            Program.MostrarMenu()
            Dim opcion = System.Console.ReadKey(CBool((True))).Key

            If opcion = System.ConsoleKey.Escape Then
                Call Spectre.Console.AnsiConsole.MarkupLine("[bold green]Gracias por usar la aplicación. ¡Hasta luego![/]")
                Return
            End If

            If opcion = System.ConsoleKey.A Then
                datosCSV = Program.CargarDesdeCSV()
                Program.MostrarSubmenuCSV(datosCSV)
            ElseIf opcion = System.ConsoleKey.B Then
                datosAPI = Await Program.CargarDesdeAPI()
                Program.MostrarSubmenuAPI(datosAPI)
            ElseIf opcion = System.ConsoleKey.C Then
                Dim datosDB = Program.CargarDesdeBaseDeDatos()
                Program.MostrarSubmenuBD(datosDB)
            Else
                Call Spectre.Console.AnsiConsole.MarkupLine("[red]Opción inválida. Intenta nuevamente.[/]")
            End If
        End While
    End Function

    Private Shared Sub MostrarMenu()
        Call Spectre.Console.AnsiConsole.Clear()
        Dim titulo = New Spectre.Console.FigletText(CStr(("Opciones"))).Color(Spectre.Console.Color.Magenta1)
        Call Spectre.Console.AnsiConsole.Write(titulo)

        Dim contenido = New Spectre.Console.Markup(Global.Microsoft.VisualBasic.Constants.vbLf & "[bold underline white]Seleccione el origen de datos[/]" & Global.Microsoft.VisualBasic.Constants.vbLf & Global.Microsoft.VisualBasic.Constants.vbLf & "[bold magenta]A[/]: [white]Archivo CSV[/]" & Global.Microsoft.VisualBasic.Constants.vbLf & "[bold magenta]B[/]: [white]Datos desde TheMovieDB API[/]" & Global.Microsoft.VisualBasic.Constants.vbLf & "[bold magenta]C[/]: [white]Datos desde Base de Datos[/]" & Global.Microsoft.VisualBasic.Constants.vbLf & "[bold magenta]ESC[/]: [white]Salir[/]")

        Dim panel = New Spectre.Console.Panel(contenido) With {
    .Border = Spectre.Console.BoxBorder.Rounded,
    .Padding = New Spectre.Console.Padding(2, 1),
    .Header = New Spectre.Console.PanelHeader("[bold purple_2]Menú Principal[/]", Spectre.Console.Justify.Center)
}

        Call Spectre.Console.AnsiConsole.Write(panel)
    End Sub

    Private Shared Sub MostrarSubmenuCSV(datos As System.Collections.Generic.List(Of ProyectoFInalConsolaC_.CsvRow))
        While True
            Call Spectre.Console.AnsiConsole.Clear()
            Dim submenu = New Spectre.Console.SelectionPrompt(Of String)().Title("[bold]Selecciona la vista de los datos CSV:[/]").AddChoices("Tabla", "Árbol", "Estadísticas", "Exportar", "Volver al menú principal")

            Dim seleccion = Spectre.Console.AnsiConsole.Prompt(submenu)

            Select Case seleccion
                Case "Tabla"
                    Program.MostrarTablaCSV(datos)
                Case "Árbol"
                    Program.MostrarVistaArbolCSV(datos)
                Case "Estadísticas"
                    Program.MostrarEstadisticasCSV(datos)
                Case "Exportar"
                    Program.MostrarMenuExportacion(datos)
                Case "Volver al menú principal"
                    Return
            End Select
        End While
    End Sub

    Private Shared Sub MostrarSubmenuAPI(datos As System.Collections.Generic.List(Of ProyectoFInalConsolaC_.Movie))
        While True
            Call Spectre.Console.AnsiConsole.Clear()
            Dim submenu = New Spectre.Console.SelectionPrompt(Of String)().Title("[bold]Selecciona la vista de los datos API:[/]").AddChoices("Tabla", "Árbol", "Estadísticas", "Exportar", "Volver al menú principal")

            Dim seleccion = Spectre.Console.AnsiConsole.Prompt(submenu)

            Select Case seleccion
                Case "Tabla"
                    Program.MostrarTablaAPI(datos)
                Case "Árbol"
                    Program.MostrarArbolAPI(datos)
                Case "Estadísticas"
                    Program.MostrarEstadisticasAPI(datos)
                Case "Exportar"
                    Program.MostrarMenuExportacionApi(datos)
                Case "Volver al menú principal"
                    Return
            End Select
        End While
    End Sub
    Private Shared Function CargarDesdeBaseDeDatos() As System.Collections.Generic.List(Of ProyectoFInalConsolaC_.CsvRow)
        Dim connectionString As String = "Server=localhost\SQLEXPRESS;Database=TopChessPlayer;Trusted_connection=yes;TrustServerCertificate=true"
        Dim query As String = "SELECT TOP 50000 * FROM top_chess_players_aug_2020;"
        Dim filas = New System.Collections.Generic.List(Of ProyectoFInalConsolaC_.CsvRow)()

        Try
            Dim connection = New System.Data.SqlClient.SqlConnection(connectionString)
            connection.Open()

            Dim command = New System.Data.SqlClient.SqlCommand(query, connection)
            Dim reader = command.ExecuteReader()

            While reader.Read()
                Dim row = New ProyectoFInalConsolaC_.CsvRow()

                For i As Integer = 0 To reader.FieldCount - 1
                    Dim columnName As String = reader.GetName(i)
                    Dim value As String = If(reader.IsDBNull(i), "", If(reader.GetValue(CInt((i))).ToString(), ""))
                    row.Fields(columnName) = value
                Next

                filas.Add(row)
            End While

            Call Spectre.Console.AnsiConsole.MarkupLine($"[green]Filas cargadas desde base de datos: {filas.Count}[/]")
        Catch ex As System.Exception
            Call Spectre.Console.AnsiConsole.MarkupLine($"[red]Error al leer datos de SQL Server: {ex.Message}[/]")
        End Try

        Return filas
    End Function
    Private Shared Function CargarDesdeCSV() As System.Collections.Generic.List(Of ProyectoFInalConsolaC_.CsvRow)
        Dim rutaArchivo = "C:\Users\volub\Downloads\archive\stats_football_players.csv"
        Dim filas = New System.Collections.Generic.List(Of ProyectoFInalConsolaC_.CsvRow)()

        If Not System.IO.File.Exists(rutaArchivo) Then
            Call Spectre.Console.AnsiConsole.MarkupLine($"[red]El archivo '{rutaArchivo}' no existe.[/]")
            Return filas
        End If

        Try
            Dim reader = New System.IO.StreamReader(rutaArchivo)
            Dim headers = reader.ReadLine()?.Split(","c)

            If headers Is Nothing OrElse headers.Length = 0 Then
                Call Spectre.Console.AnsiConsole.MarkupLine("[red]No se encontraron encabezados en el archivo.[/]")
                Return filas
            End If

            While Not reader.EndOfStream
                Dim line = reader.ReadLine()
                If String.IsNullOrWhiteSpace(line) Then Continue While

                Dim values = line.Split(","c)

                Dim row = New ProyectoFInalConsolaC_.CsvRow()
                Dim i As Integer = 0

                While i < headers.Length AndAlso i < values.Length
                    row.Fields(headers(CInt((i))).Trim()) = values(CInt((i))).Trim()
                    i += 1
                End While

                filas.Add(row)
            End While
        Catch ex As System.Exception
            Call Spectre.Console.AnsiConsole.MarkupLine($"[red]Error al leer el archivo: {ex.Message}[/]")
        End Try

        Return filas
    End Function

    Private Shared Async Function CargarDesdeAPI(Optional totalMovies As Integer = 1000) As System.Threading.Tasks.Task(Of System.Collections.Generic.List(Of ProyectoFInalConsolaC_.Movie))
        Dim httpClient = New System.Net.Http.HttpClient()
        httpClient.DefaultRequestHeaders.Authorization = New System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", "eyJhbGciOiJIUzI1NiJ9.eyJhdWQiOiJkM2E5NDU5MmNjM2E5NDkwMWNhYjI2NzNmZWFhZGM4MiIsIm5iZiI6MTc0ODI4ODIzMy42NjQsInN1YiI6IjY4MzRjMmU" & "5NjQwZTA1YjQyOGI2YmEwMiIsInNjb3BlcyI6WyJhcGlfcmVhZCJdLCJ2ZXJzaW9uIjoxfQ.ArYNRf2GqMDkot9TCd_5323QEN-XHpMVXfp1PxoNu4A")

        Dim moviesPerPage As Integer = 20
        Dim pagesToRequest As Integer = CInt(System.Math.Ceiling(totalMovies / CDbl(moviesPerPage)))

        Dim allMovies = New System.Collections.Generic.List(Of ProyectoFInalConsolaC_.Movie)()

        For page As Integer = 1 To pagesToRequest
            Dim url = $"https://api.themoviedb.org/3/movie/popular?language=es-ES&page={page}"

            Try
                Dim response = Await httpClient.GetStringAsync(url)
                Dim json = System.Text.Json.JsonDocument.Parse(response)

                Dim peliculas = json.RootElement.GetProperty(CStr(("results"))).EnumerateArray().[Select](Function(m) New ProyectoFInalConsolaC_.Movie With {
.VoteAverage = m.GetProperty(CStr(("vote_average"))).GetDecimal(),
.VoteCount = m.GetProperty(CStr(("vote_count"))).GetInt32(),
.Popularity = m.GetProperty(CStr(("popularity"))).GetDecimal()
}).ToList()

                allMovies.AddRange(peliculas)

                If allMovies.Count >= totalMovies Then Exit For
            Catch ex As System.Exception
                Call Spectre.Console.AnsiConsole.MarkupLine($"[red]Error al obtener datos de la API en página {page}: {ex.Message}[/]")
                Exit For
            End Try
        Next

        Return allMovies.Take(totalMovies).ToList()
    End Function

    Private Shared Sub MostrarSubmenuBD(datos As System.Collections.Generic.List(Of ProyectoFInalConsolaC_.CsvRow))
        While True
            Call Spectre.Console.AnsiConsole.Clear()
            Dim submenu = New Spectre.Console.SelectionPrompt(Of String)().Title("[bold]Selecciona la vista de los datos BD:[/]").AddChoices("Tabla", "Árbol", "Estadísticas", "Exportar", "Volver al menú principal")

            Dim seleccion = Spectre.Console.AnsiConsole.Prompt(submenu)

            Select Case seleccion
                Case "Tabla"
                    Program.MostrarTablaBaseDeDatos(datos)
                Case "Árbol"
                    Program.MostrarArbolBaseDeDatos(datos)
                Case "Estadísticas"
                    Program.MostrarEstadisticasBaseDeDatos(datos)
                Case "Exportar"
                    Program.MostrarMenuExportacion(datos)
                Case "Volver al menú principal"
                    Return
            End Select
        End While
    End Sub
    ' mostrar base de datos
    Private Shared Sub MostrarTablaBaseDeDatos(rows As System.Collections.Generic.List(Of ProyectoFInalConsolaC_.CsvRow))
        Dim columnasMostrar As String() = {"Fideid", "Name", "Federation", "Gender", "Year_of_birth", "Title", "Standard_Rating", "Rapid_rating", "Blitz_rating", "Inactive_flag"}

        Const pageSize As Integer = 10
        Dim page As Integer = 0
        Dim totalPages As Integer = CInt(System.Math.Ceiling(CDbl(rows.Count) / pageSize))

        While True
            Call Spectre.Console.AnsiConsole.Clear()

            Dim table = New Spectre.Console.Table().Border(Spectre.Console.TableBorder.Rounded)
            For Each col In columnasMostrar
                table.AddColumn($"[bold yellow]{col}[/]")
            Next

            Dim paginated = rows.Skip(page * pageSize).Take(pageSize)

            For Each row In paginated
                Dim cells = columnasMostrar.[Select](Function(h) $"[white]{(If(row.Fields.ContainsKey(h), row.Fields(h), ""))}[/]").ToArray()
                table.AddRow(cells)
            Next

            Call Spectre.Console.AnsiConsole.Write(table)
            Call Spectre.Console.AnsiConsole.MarkupLine($"[grey]Página {page + 1} de {totalPages}[/]")
            Call Spectre.Console.AnsiConsole.MarkupLine("[blue]Use ↑ ↓ para navegar páginas, Esc para salir.[/]")

            Dim key = System.Console.ReadKey(CBool((True))).Key

            If key = System.ConsoleKey.DownArrow AndAlso page < totalPages - 1 Then
                page += 1
            ElseIf key = System.ConsoleKey.UpArrow AndAlso page > 0 Then
                page -= 1
            ElseIf key = System.ConsoleKey.Escape Then
                Return
            End If
        End While
    End Sub
    Private Shared Sub MostrarArbolBaseDeDatos(rows As System.Collections.Generic.List(Of ProyectoFInalConsolaC_.CsvRow))
        Const pageSize As Integer = 1
        Dim page As Integer = 0
        Dim totalPages As Integer = CInt(System.Math.Ceiling(CDbl(rows.Count) / pageSize))

        While True
            Call Spectre.Console.AnsiConsole.Clear()

            Dim paginated = rows.Skip(page * pageSize).Take(pageSize).ToList()

            For Each jugador In paginated
                If Not jugador.Fields.ContainsKey("Name") Then Continue For

                Dim root = New Spectre.Console.Tree(CStr(($"[bold yellow]{(jugador.Fields(CStr(("Name"))))}[/]"))).Style(Spectre.Console.Style.Parse("blue"))
                For Each kvp In jugador.Fields
                    If Equals(kvp.Key, "Name") Then Continue For
                    root.AddNode($"[white]{kvp.Key}[/]: [cyan]{kvp.Value}[/]")
                Next
                Call Spectre.Console.AnsiConsole.Write(root)
                Call Spectre.Console.AnsiConsole.WriteLine()
            Next

            Call Spectre.Console.AnsiConsole.MarkupLine($"[grey]Página {page + 1} de {totalPages}[/]")
            Call Spectre.Console.AnsiConsole.MarkupLine("[blue]Use ↑ ↓ para navegar, Esc para volver al menú.[/]")

            Dim key = System.Console.ReadKey(CBool((True))).Key

            If key = System.ConsoleKey.DownArrow AndAlso page < totalPages - 1 Then
                page += 1
            ElseIf key = System.ConsoleKey.UpArrow AndAlso page > 0 Then
                page -= 1
            ElseIf key = System.ConsoleKey.Escape Then
                Exit While
            End If
        End While
    End Sub
    Private Shared Sub MostrarEstadisticasBaseDeDatos(jugadores As System.Collections.Generic.List(Of ProyectoFInalConsolaC_.CsvRow))
        Dim index As Integer = 0

        Dim s As Integer = Nothing, r As Integer = Nothing, b As Integer = Nothing

        While True
            Call Spectre.Console.AnsiConsole.Clear()

            If jugadores.Count = 0 Then
                Call Spectre.Console.AnsiConsole.MarkupLine("[red]No hay datos para mostrar.[/]")
                Return
            End If

            Dim jugador = jugadores(index)
            Dim nombre As String = jugador.Fields.GetValueOrDefault("Name", "Desconocido")
            Dim std As Integer = If(Integer.TryParse(jugador.Fields.GetValueOrDefault("Standard_Rating", "0"), s), s, 0)
            Dim rapid As Integer = If(Integer.TryParse(jugador.Fields.GetValueOrDefault("Rapid_rating", "0"), r), r, 0)
            Dim blitz As Integer = If(Integer.TryParse(jugador.Fields.GetValueOrDefault("Blitz_rating", "0"), b), b, 0)

            Dim chart = Spectre.Console.BarChartExtensions.AddItem(Spectre.Console.BarChartExtensions.AddItem(Spectre.Console.BarChartExtensions.CenterLabel(Spectre.Console.BarChartExtensions.Label(Spectre.Console.BarChartExtensions.Width(New Spectre.Console.BarChart(), CType((60), System.Int32?)), CStr(($"[bold yellow]{(nombre)}[/]")))), CStr(("Standard")), CDbl((std)), Spectre.Console.Color.Green), CStr(("Rapid")), CDbl((rapid)), Spectre.Console.Color.Orange1).AddItem("Blitz", blitz, Spectre.Console.Color.Blue)

            Call Spectre.Console.AnsiConsole.Write(chart)

            Call Spectre.Console.AnsiConsole.MarkupLine($"
[grey]Jugador {index + 1} de {jugadores.Count}[/]")
            Call Spectre.Console.AnsiConsole.MarkupLine("[blue]Use ← → para cambiar jugador. Esc para volver.[/]")

            Dim key = System.Console.ReadKey(CBool((True))).Key

            If key = System.ConsoleKey.RightArrow Then
                index = (index + 1) Mod jugadores.Count
            ElseIf key = System.ConsoleKey.LeftArrow Then
                index = (index - 1 + jugadores.Count) Mod jugadores.Count
            ElseIf key = System.ConsoleKey.Escape Then
                Exit While
            End If
        End While
    End Sub

    ' Mostrar datos API 
    Private Shared Sub MostrarTablaAPI(movies As System.Collections.Generic.List(Of ProyectoFInalConsolaC_.Movie))
        While True
            Const pageSize As Integer = 10
            Dim page As Integer = 0
            Dim totalPages As Integer = CInt(System.Math.Ceiling(CDbl(movies.Count) / pageSize))

            While True
                Call Spectre.Console.AnsiConsole.Clear()
                Dim table = New Spectre.Console.Table().Border(Spectre.Console.TableBorder.Rounded)
                table.AddColumn("[yellow]Nombre de la Película[/]")

                Dim paginated = movies.Skip(page * pageSize).Take(pageSize).ToList()

                For Each movie In paginated
                    table.AddRow($"[white]{movie.Title}[/]")
                Next

                Call Spectre.Console.AnsiConsole.Write(table)
                Call Spectre.Console.AnsiConsole.MarkupLine($"[grey]Página {page + 1} de {totalPages}[/]")
                Call Spectre.Console.AnsiConsole.MarkupLine("[blue]Use N para siguiente página, P para anterior, Esc para salir.[/]")

                Dim key = System.Console.ReadKey(CBool((True))).Key

                If key = System.ConsoleKey.N AndAlso page < totalPages - 1 Then
                    page += 1
                ElseIf key = System.ConsoleKey.P AndAlso page > 0 Then
                    page -= 1
                ElseIf key = System.ConsoleKey.Escape Then
                    Return
                End If
            End While
        End While
    End Sub

    Private Shared Sub MostrarArbolAPI(movies As System.Collections.Generic.List(Of ProyectoFInalConsolaC_.Movie))
        If movies Is Nothing OrElse movies.Count = 0 Then
            Call Spectre.Console.AnsiConsole.MarkupLine("[red]No hay películas para mostrar.[/]")
            Return
        End If

        Dim index As Integer = 0
        Dim key As System.ConsoleKeyInfo

        Do
            System.Console.Clear()
            Dim movie = movies(index)
            Dim tree = New Spectre.Console.Tree(CStr(($"[bold yellow]{(movie.Title)}[/]"))).Style(Spectre.Console.Style.Parse("blue"))
            tree.AddNode($"[white]Popularidad[/]: [cyan]{movie.Popularity}[/]")
            tree.AddNode($"[white]Votos promedio[/]: [cyan]{movie.VoteAverage}[/]")
            tree.AddNode($"[white]Número de votos[/]: [cyan]{movie.VoteCount}[/]")
            Call Spectre.Console.AnsiConsole.Write(tree)
            Call Spectre.Console.AnsiConsole.WriteLine()

            Call Spectre.Console.AnsiConsole.MarkupLine("[blue]Use ← y → para navegar, Esc para salir.[/]")
            key = System.Console.ReadKey(True)

            If key.Key = System.ConsoleKey.RightArrow Then
                index = (index + 10) Mod movies.Count ' siguiente película, vuelve al inicio si llega al final
            ElseIf key.Key = System.ConsoleKey.LeftArrow Then
                index = (index - 10 + movies.Count) Mod movies.Count ' anterior película, vuelve al final si está en 0
            End If

        Loop While key.Key <> System.ConsoleKey.Escape

        System.Console.Clear()
    End Sub

    Private Shared Sub MostrarEstadisticasAPI(movies As System.Collections.Generic.List(Of ProyectoFInalConsolaC_.Movie))
        Dim index As Integer = 0

        While True
            Call Spectre.Console.AnsiConsole.Clear()

            If movies.Count = 0 Then
                Call Spectre.Console.AnsiConsole.MarkupLine("[red]No hay datos para mostrar.[/]")
                Return
            End If

            Dim movie = movies(index)

            Dim chart = Spectre.Console.BarChartExtensions.AddItem(Spectre.Console.BarChartExtensions.AddItem(Spectre.Console.BarChartExtensions.CenterLabel(Spectre.Console.BarChartExtensions.Label(Spectre.Console.BarChartExtensions.Width(New Spectre.Console.BarChart(), CType((60), System.Int32?)), CStr(($"[bold yellow]{(movie.Title)}[/]")))), CStr(("Popularidad")), CSng(movie.Popularity), Spectre.Console.Color.Blue), CStr(("Voto Promedio")), CSng(movie.VoteAverage), Spectre.Console.Color.Green).AddItem("Votos", movie.VoteCount, Spectre.Console.Color.Orange1)

            Call Spectre.Console.AnsiConsole.Write(chart)

            Call Spectre.Console.AnsiConsole.MarkupLine($"
[grey]Película {index + 1} de {movies.Count}[/]")
            Call Spectre.Console.AnsiConsole.MarkupLine("[blue]Use ← → para cambiar película. Esc para volver.[/]")

            Dim key = System.Console.ReadKey(CBool((True))).Key

            If key = System.ConsoleKey.RightArrow Then
                index = (index + 1) Mod movies.Count
            ElseIf key = System.ConsoleKey.LeftArrow Then
                index = (index - 1 + movies.Count) Mod movies.Count
            ElseIf key = System.ConsoleKey.Escape Then
                Return
            End If
        End While
    End Sub

    ' Mostrar CSV
    Private Shared Sub MostrarTablaCSV(rows As System.Collections.Generic.List(Of ProyectoFInalConsolaC_.CsvRow))
        Dim columnasMostrar As String() = {"Season", "League", "Team", "Player", "Nation", "Position", "Age", "Match Played", "Goals", "Assists"}
        Dim columnasValidas = columnasMostrar.Where(Function(c) rows.Any(Function(r) r.Fields.ContainsKey(c))).ToArray()

        While True
            Call Spectre.Console.AnsiConsole.Clear()
            Call Spectre.Console.AnsiConsole.MarkupLine("[bold]Ingrese texto para filtrar (deje vacío para mostrar todo):[/]")
            Dim filtro As String = If(System.Console.ReadLine(), "")

            Dim filasFiltradas = rows.Where(Function(row) columnasValidas.Any(Function(col) row.Fields.ContainsKey(col) AndAlso row.Fields(CStr((col))).Contains(filtro, System.StringComparison.OrdinalIgnoreCase))).ToList()

            If filasFiltradas.Count = 0 Then
                Call Spectre.Console.AnsiConsole.MarkupLine("[red]No se encontraron filas que coincidan con el filtro.[/]")
                Call Spectre.Console.AnsiConsole.MarkupLine("Presione cualquier tecla para intentar otro filtro...")
                System.Console.ReadKey(True)
                Continue While
            End If

            Const pageSize As Integer = 10
            Dim page As Integer = 0
            Dim totalPages As Integer = CInt(System.Math.Ceiling(CDbl(filasFiltradas.Count) / pageSize))

            While True
                Call Spectre.Console.AnsiConsole.Clear()

                Dim table = New Spectre.Console.Table().Border(Spectre.Console.TableBorder.Rounded)
                For Each col In columnasValidas
                    table.AddColumn($"[bold yellow]{col}[/]")
                Next

                Dim paginated = filasFiltradas.Skip(page * pageSize).Take(pageSize)

                For Each row In paginated
                    Dim cells = columnasValidas.[Select](Function(h) $"[white]{(If(row.Fields.ContainsKey(h), row.Fields(h), ""))}[/]").ToArray()
                    table.AddRow(cells)
                Next

                Call Spectre.Console.AnsiConsole.Write(table)
                Call Spectre.Console.AnsiConsole.MarkupLine($"[grey]Página {page + 1} de {totalPages}[/]")
                Call Spectre.Console.AnsiConsole.MarkupLine("[blue]Use ↑ ↓ para navegar páginas, F para nuevo filtro, Esc para salir.[/]")

                Dim key = System.Console.ReadKey(CBool((True))).Key

                If key = System.ConsoleKey.DownArrow AndAlso page < totalPages - 1 Then
                    page += 1
                ElseIf key = System.ConsoleKey.UpArrow AndAlso page > 0 Then
                    page -= 1
                ElseIf key = System.ConsoleKey.F Then
                    Exit While ' Salir para ingresar nuevo filtro
                ElseIf key = System.ConsoleKey.Escape Then
                    Return
                End If
            End While
        End While
    End Sub

    Private Shared Sub MostrarVistaArbolCSV(rows As System.Collections.Generic.List(Of ProyectoFInalConsolaC_.CsvRow))
        Const pageSize As Integer = 1
        Dim page As Integer = 0
        Dim totalPages As Integer = CInt(System.Math.Ceiling(CDbl(rows.Count) / pageSize))

        While True
            Call Spectre.Console.AnsiConsole.Clear()

            Dim paginated = rows.Skip(page * pageSize).Take(pageSize).ToList()

            For Each jugador In paginated
                If Not jugador.Fields.ContainsKey("Player") Then Continue For

                Dim root = New Spectre.Console.Tree(CStr(($"[bold yellow]{(jugador.Fields(CStr(("Player"))))}[/]"))).Style(Spectre.Console.Style.Parse("blue"))
                For Each kvp In jugador.Fields
                    If Equals(kvp.Key, "Player") Then Continue For
                    root.AddNode($"[white]{kvp.Key}[/]: [cyan]{kvp.Value}[/]")
                Next
                Call Spectre.Console.AnsiConsole.Write(root)
                Call Spectre.Console.AnsiConsole.WriteLine()
            Next

            Call Spectre.Console.AnsiConsole.MarkupLine($"[grey]Página {page + 1} de {totalPages}[/]")
            Call Spectre.Console.AnsiConsole.MarkupLine("[blue]Use ↑ ↓ para navegar, Esc para volver al menú.[/]")

            Dim key = System.Console.ReadKey(CBool((True))).Key

            If key = System.ConsoleKey.DownArrow AndAlso page < totalPages - 1 Then
                page += 1
            ElseIf key = System.ConsoleKey.UpArrow AndAlso page > 0 Then
                page -= 1
            ElseIf key = System.ConsoleKey.Escape Then
                Exit While
            End If
        End While
    End Sub

    Private Shared Sub MostrarEstadisticasCSV(rows As System.Collections.Generic.List(Of ProyectoFInalConsolaC_.CsvRow))
        Dim index As Integer = 0

        While True
            Call Spectre.Console.AnsiConsole.Clear()

            If rows.Count = 0 Then
                Call Spectre.Console.AnsiConsole.MarkupLine("[red]No hay datos para mostrar.[/]")
                Return
            End If

            Dim jugador = rows(index)

            If Not jugador.Fields.ContainsKey("Player") Then
                index = (index + 1) Mod rows.Count
                Continue While
            End If

            Dim nombre = jugador.Fields("Player")
            Dim partidos As Integer = 0, goles As Integer = 0, asistencias As Integer = 0, amarillas As Integer = 0, rojas As Integer = 0

            Call Integer.TryParse(jugador.Fields.GetValueOrDefault("Match Played", "0"), partidos)
            Call Integer.TryParse(jugador.Fields.GetValueOrDefault("Goals", "0"), goles)
            Call Integer.TryParse(jugador.Fields.GetValueOrDefault("Assists", "0"), asistencias)
            Call Integer.TryParse(jugador.Fields.GetValueOrDefault("Yellow Cards", "0"), amarillas)
            Call Integer.TryParse(jugador.Fields.GetValueOrDefault("Red Cards", "0"), rojas)

            Dim chart = Spectre.Console.BarChartExtensions.AddItem(Spectre.Console.BarChartExtensions.AddItem(Spectre.Console.BarChartExtensions.AddItem(Spectre.Console.BarChartExtensions.AddItem(Spectre.Console.BarChartExtensions.CenterLabel(Spectre.Console.BarChartExtensions.Label(Spectre.Console.BarChartExtensions.Width(New Spectre.Console.BarChart(), CType((60), System.Int32?)), CStr(($"[bold yellow]Estadísticas de {(nombre)}[/]")))), CStr(("Partidos")), CDbl((partidos)), Spectre.Console.Color.Blue), CStr(("Goles")), CDbl((goles)), Spectre.Console.Color.Green), CStr(("Asistencias")), CDbl((asistencias)), Spectre.Console.Color.Orange1), CStr(("Amarillas")), CDbl((amarillas)), Spectre.Console.Color.Yellow).AddItem("Rojas", rojas, Spectre.Console.Color.Red)

            Call Spectre.Console.AnsiConsole.Write(chart)

            Call Spectre.Console.AnsiConsole.MarkupLine($"
[grey]Jugador {index + 1} de {rows.Count}[/]")
            Call Spectre.Console.AnsiConsole.MarkupLine("[blue]Use ← → para cambiar jugador. Esc para volver al menú.[/]")

            Dim key = System.Console.ReadKey(CBool((True))).Key

            If key = System.ConsoleKey.RightArrow Then
                index = (index + 1) Mod rows.Count
            ElseIf key = System.ConsoleKey.LeftArrow Then
                index = (index - 1 + rows.Count) Mod rows.Count
            ElseIf key = System.ConsoleKey.Escape Then
                Exit While
            End If
        End While
    End Sub

    Private Shared Sub MostrarMenuExportacion(jugadores As System.Collections.Generic.List(Of ProyectoFInalConsolaC_.CsvRow))
        If jugadores Is Nothing OrElse jugadores.Count = 0 Then
            Call Spectre.Console.AnsiConsole.MarkupLine("[red]No hay datos para exportar.[/]")
            System.Console.ReadKey(True)
            Return
        End If

        Dim formato = Spectre.Console.AnsiConsole.Prompt(New Spectre.Console.SelectionPrompt(Of String)().Title("Selecciona el [green]formato de exportación[/]:").AddChoices("CSV", "TXT", "XML", "JSON", "PDF"))

        Dim ruta As String = Spectre.Console.AnsiConsole.Ask(Of String)("Ingresa la ruta y nombre del archivo (sin extensión):")

        Try
            Select Case formato.ToUpper()
                Case "CSV"
                    Program.ExportarACSV(jugadores, ruta & ".csv")
                Case "TXT"
                    Program.ExportarATXT(jugadores, ruta & ".txt")
                Case "XML"
                    Program.ExportarAXML(jugadores, ruta & ".xml")
                Case "JSON"
                    Program.ExportarAJSON(jugadores, ruta & ".json")
                Case "PDF"
                    Program.ExportarAPDF(jugadores, ruta & ".pdf")
            End Select

            Call Spectre.Console.AnsiConsole.MarkupLine($"[green]Datos exportados correctamente como {formato}.[/]")
        Catch ex As System.Exception
            Call Spectre.Console.AnsiConsole.MarkupLine($"[red]Error al exportar: {ex.Message}[/]")
        End Try
        System.Console.ReadKey(True)
    End Sub

    Private Shared Function ConvertMoviesToCsvRows(movies As System.Collections.Generic.List(Of ProyectoFInalConsolaC_.Movie)) As System.Collections.Generic.List(Of ProyectoFInalConsolaC_.CsvRow)
        Return movies.[Select](Function(m) New ProyectoFInalConsolaC_.CsvRow With {
                .Fields = New System.Collections.Generic.Dictionary(Of String, String) From {
        {"VoteAverage", m.VoteAverage.ToString(System.Globalization.CultureInfo.InvariantCulture)},
        {"VoteCount", m.VoteCount.ToString()},
        {"Popularity", m.Popularity.ToString(System.Globalization.CultureInfo.InvariantCulture)}
    }
}).ToList()
    End Function

    Private Shared Sub MostrarMenuExportacionApi(movies As System.Collections.Generic.List(Of ProyectoFInalConsolaC_.Movie))
        Dim csvRows = Program.ConvertMoviesToCsvRows(movies)

        While True
            System.Console.Clear()
            Call Spectre.Console.AnsiConsole.MarkupLine("[bold cyan]Menú de Exportación (Datos desde API)[/]")
            System.Console.WriteLine("1. Exportar a CSV")
            System.Console.WriteLine("2. Exportar a JSON")
            System.Console.WriteLine("3. Exportar a TXT")
            System.Console.WriteLine("4. Exportar a XML")
            System.Console.WriteLine("5. Exportar a PDF")
            System.Console.WriteLine("6. Volver al menú principal")

            System.Console.Write(Global.Microsoft.VisualBasic.Constants.vbLf & "Seleccione una opción: ")
                        ''' Cannot convert LocalDeclarationStatementSyntax, System.InvalidCastException: No se puede convertir un objeto de tipo 'Microsoft.CodeAnalysis.VisualBasic.Syntax.EmptyStatementSyntax' al tipo 'Microsoft.CodeAnalysis.VisualBasic.Syntax.ExpressionSyntax'.
'''    en ICSharpCode.CodeConverter.VB.CommonConversions.RemodelVariableDeclaration(VariableDeclarationSyntax declaration)
'''    en ICSharpCode.CodeConverter.VB.MethodBodyExecutableStatementVisitor.VisitLocalDeclarationStatement(LocalDeclarationStatementSyntax node)
'''    en Microsoft.CodeAnalysis.CSharp.CSharpSyntaxVisitor`1.Visit(SyntaxNode node)
'''    en ICSharpCode.CodeConverter.VB.CommentConvertingMethodBodyVisitor.DefaultVisit(SyntaxNode node)
''' 
''' Input:
'''             string opcion = System.Console.ReadLine()!;
''' 
''' 

            Select Case opcion
                Case "1"
                    Program.ExportarACSV(csvRows, "export_api.csv")
                Case "2"
                    Program.ExportarAJSON(csvRows, "export_api.json")
                Case "3"
                    Program.ExportarATXT(csvRows, "export_api.txt")
                Case "4"
                    Program.ExportarAXML(csvRows, "export_api.xml")
                Case "5"
                    Program.ExportarAPDF(csvRows, "export_api.pdf")
                Case "6"
                    Return
                Case Else
                    System.Console.WriteLine("Opción inválida. Presione una tecla para continuar...")
                    System.Console.ReadKey()
            End Select

            System.Console.WriteLine("Exportación realizada. Presione una tecla para continuar...")
            System.Console.ReadKey()
        End While
    End Sub

    Private Shared Sub ExportarACSV(datos As System.Collections.Generic.List(Of ProyectoFInalConsolaC_.CsvRow), ruta As String)
        Try
            If datos.Count = 0 Then
                Call Spectre.Console.AnsiConsole.MarkupLine("[red]No hay datos para exportar.[/]")
                System.Console.ReadKey(True)
                Return
            End If

            Dim headers = datos(CInt((0))).Fields.Keys.ToList()

            Dim writer = New System.IO.StreamWriter(ruta, False, System.Text.Encoding.UTF8)
            writer.WriteLine(String.Join(",", headers))

            For Each fila In datos
                Dim valores = headers.[Select](Function(h) If(fila.Fields.ContainsKey(h), fila.Fields(h), ""))
                writer.WriteLine(String.Join(",", valores))
            Next

            Call Spectre.Console.AnsiConsole.MarkupLine($"[green]Datos exportados exitosamente a CSV: {ruta}[/]")
            Program.PreguntarYEnviarCorreo(ruta)
        Catch ex As System.Exception
            Call Spectre.Console.AnsiConsole.MarkupLine($"[red]Error al exportar a CSV: {ex.Message}[/]")
        End Try
        System.Console.ReadKey(True)
    End Sub

    Private Shared Sub ExportarATXT(datos As System.Collections.Generic.List(Of ProyectoFInalConsolaC_.CsvRow), ruta As String)
        Try
            Dim writer = New System.IO.StreamWriter(ruta, False, System.Text.Encoding.UTF8)
            For Each fila In datos
                For Each kvp In fila.Fields
                    writer.WriteLine($"{kvp.Key}: {kvp.Value}")
                Next
                writer.WriteLine(New String("-"c, 40))
            Next

            Call Spectre.Console.AnsiConsole.MarkupLine($"[green]Datos exportados exitosamente a TXT: {ruta}[/]")
            Program.PreguntarYEnviarCorreo(ruta)
        Catch ex As System.Exception
            Call Spectre.Console.AnsiConsole.MarkupLine($"[red]Error al exportar a TXT: {ex.Message}[/]")
        End Try
        System.Console.ReadKey(True)
    End Sub

    Private Shared Sub ExportarAJSON(datos As System.Collections.Generic.List(Of ProyectoFInalConsolaC_.CsvRow), ruta As String)
        Try
            Dim json = System.Text.Json.JsonSerializer.Serialize(datos.[Select](Function(d) d.Fields), New System.Text.Json.JsonSerializerOptions With {
                .WriteIndented = True
            })
            Call System.IO.File.WriteAllText(ruta, json, System.Text.Encoding.UTF8)
            Call Spectre.Console.AnsiConsole.MarkupLine($"[green]Datos exportados exitosamente a JSON: {ruta}[/]")
            Program.PreguntarYEnviarCorreo(ruta)
        Catch ex As System.Exception
            Call Spectre.Console.AnsiConsole.MarkupLine($"[red]Error al exportar a JSON: {ex.Message}[/]")
        End Try
        System.Console.ReadKey(True)
    End Sub

    Private Shared Sub ExportarAXML(datos As System.Collections.Generic.List(Of ProyectoFInalConsolaC_.CsvRow), ruta As String)
        Try
            Dim xml = New System.Text.StringBuilder()
            xml.AppendLine("<Jugadores>")

            For Each fila In datos
                xml.AppendLine("  <Jugador>")
                For Each kvp In fila.Fields
                    xml.AppendLine($"    <{kvp.Key}>{System.Security.SecurityElement.Escape(kvp.Value)}</{kvp.Key}>")
                Next
                xml.AppendLine("  </Jugador>")
            Next

            xml.AppendLine("</Jugadores>")
            Call System.IO.File.WriteAllText(ruta, xml.ToString(), System.Text.Encoding.UTF8)

            Call Spectre.Console.AnsiConsole.MarkupLine($"[green]Datos exportados exitosamente a XML: {ruta}[/]")
            Program.PreguntarYEnviarCorreo(ruta)
        Catch ex As System.Exception
            Call Spectre.Console.AnsiConsole.MarkupLine($"[red]Error al exportar a XML: {ex.Message}[/]")
        End Try
        System.Console.ReadKey(True)
    End Sub
    ' Función para preguntar si desea enviar por correo
    Private Shared Sub PreguntarYEnviarCorreo(rutaArchivo As String)
        Call Spectre.Console.AnsiConsole.MarkupLine(Global.Microsoft.VisualBasic.Constants.vbLf & "¿Deseas enviar el archivo exportado por correo electrónico? (s/n)")
        Dim respuesta = System.Console.ReadLine()?.Trim().ToLower()

        If Equals(respuesta, "s") OrElse Equals(respuesta, "si") Then
            Call Spectre.Console.AnsiConsole.MarkupLine("Ingresa el correo electrónico destinatario:")
            Dim destinatario = System.Console.ReadLine()?.Trim()

            If Not String.IsNullOrEmpty(destinatario) Then
                Try
                    Dim mensaje = New MimeKit.MimeMessage()
                    mensaje.From.Add(New MimeKit.MailboxAddress("App Estadística", "elcomparosh97@gmail.com"))
                    mensaje.[To].Add(MimeKit.MailboxAddress.Parse(destinatario))
                    mensaje.Subject = "Archivo exportado"

                    Dim builder = New MimeKit.BodyBuilder With {
    .TextBody = "Se adjunta el archivo exportado desde la aplicación."
}

                    builder.Attachments.Add(rutaArchivo)
                    mensaje.Body = builder.ToMessageBody()

                    Dim cliente = New MailKit.Net.Smtp.SmtpClient()
                    cliente.Connect("smtp.gmail.com", 587, MailKit.Security.SecureSocketOptions.StartTls)

                    ' Autenticación del remitente
                    cliente.Authenticate("elcomparosh97@gmail.com", "kecp sfqy aosh jpjj") ' Usa clave de aplicación, no tu contraseña

                    cliente.Send(mensaje)
                    cliente.Disconnect(True)

                    Call Spectre.Console.AnsiConsole.MarkupLine($"[green]Archivo enviado exitosamente a {destinatario}.[/]")
                Catch ex As System.Exception
                    Call Spectre.Console.AnsiConsole.MarkupLine($"[red]Error al enviar correo: {ex.Message}[/]")
                End Try
            Else
                Call Spectre.Console.AnsiConsole.MarkupLine("[red]Correo inválido. No se envió el archivo.[/]")
            End If
        Else
            Call Spectre.Console.AnsiConsole.MarkupLine("No se envió el archivo por correo.")
        End If

        System.Console.ReadKey(True)
    End Sub
    Private Shared Sub ExportarAPDF(datos As System.Collections.Generic.List(Of ProyectoFInalConsolaC_.CsvRow), ruta As String)
        Dim value As System.String = Nothing
        Try
            If datos Is Nothing OrElse datos.Count = 0 Then
                Call Spectre.Console.AnsiConsole.MarkupLine("[red]No hay datos para exportar.[/]")
                System.Console.ReadKey(True)
                Return
            End If

            Dim headers = datos(CInt((0))).Fields.Keys.ToList()

            Dim document = QuestPDF.Fluent.Document.Create(Sub(container)
                                                               Dim value As System.String = Nothing
                                                               container.Page(Sub(page)
                                                                                  page.Margin(20)
                                                                                  page.Header().Text("Reporte de Jugadores").SemiBold().FontSize(20).FontColor(QuestPDF.Helpers.Colors.Blue.Medium)
                                                                                  Dim value As System.String = Nothing

                                                                                  page.Content().Table(Sub(table)
                                                                                                           ' Define columns count
                                                                                                           table.ColumnsDefinition(Sub(columns)
                                                                                                                                       For Each __ In headers
                                                                                                                                           columns.RelativeColumn()
                                                                                                                                       Next
                                                                                                                                   End Sub)

                                                                                                           ' Header row
                                                                                                           table.Header(Sub(header)
                                                                                                                            For Each headerName In headers
                                                                                                                                QuestPDF.Fluent.PaddingExtensions.Padding(QuestPDF.Fluent.ElementExtensions.Background(header.Cell(), CType((QuestPDF.Helpers.Colors.Grey.Lighten2), QuestPDF.Infrastructure.Color)), CSng((5))).Text(headerName).SemiBold()
                                                                                                                            Next
                                                                                                                        End Sub)
                                                                                                           Dim value As System.String = Nothing

                                                                                                           ' Data rows
                                                                                                           For Each fila In datos
                                                                                                               For Each headerName In headers
                                                                                                                   fila.Fields.TryGetValue(headerName, value)
                                                                                                                   QuestPDF.Fluent.PaddingExtensions.Padding(QuestPDF.Fluent.BorderExtensions.BorderColor(QuestPDF.Fluent.BorderExtensions.BorderBottom(table.Cell(), CSng((1))), CType((QuestPDF.Helpers.Colors.Grey.Lighten2), QuestPDF.Infrastructure.Color)), CSng((5))).Text(If(value, ""))
                                                                                                               Next
                                                                                                           Next
                                                                                                       End Sub)

                                                                                  QuestPDF.Fluent.AlignmentExtensions.AlignCenter(page.Footer()).Text(Sub(x)
                                                                                                                                                          x.Span("Página ")
                                                                                                                                                          x.CurrentPageNumber()
                                                                                                                                                          x.Span(" de ")
                                                                                                                                                          x.TotalPages()
                                                                                                                                                      End Sub)
                                                                              End Sub)
                                                           End Sub)

            document.GeneratePdf(ruta)

            Call Spectre.Console.AnsiConsole.MarkupLine($"[green]Datos exportados exitosamente a PDF: {ruta}[/]")
            Program.PreguntarYEnviarCorreo(ruta)
        Catch ex As System.Exception
            Call Spectre.Console.AnsiConsole.MarkupLine($"[red]Error al exportar a PDF: {ex.Message}[/]")
            System.Console.ReadKey(True)
        End Try
    End Sub
End Class
