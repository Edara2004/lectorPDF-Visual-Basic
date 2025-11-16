Option Strict On
Option Explicit On

Imports System.IO
Imports OfficeOpenXml
Imports System.Globalization

Public Class ExcelLoader
    ' Clase encargada de leer el archivo Excel y convertir cada fila
    ' en un objeto Conexion. No hace nada de la UI, solo lectura y formato.
    Public Sub New()
        ' EPPlus requiere indicar el modo de licencia en tiempo de ejecución.
        ' Para este trabajo académico se usa NonCommercial.
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial
    End Sub

    Public Function Load(filePath As String) As List(Of Conexion)
        ' Validaciones básicas: la ruta no debe estar vacía
        If String.IsNullOrWhiteSpace(filePath) Then
            Throw New ArgumentException("Ruta de archivo vacía.")
        End If

        If Not File.Exists(filePath) Then
            Throw New FileNotFoundException("El archivo no existe.", filePath)
        End If

        Dim result As New List(Of Conexion)()

        ' Abrimos el archivo Excel con EPPlus (se cierra al salir del Using)
        Using package As New ExcelPackage(New FileInfo(filePath))
            If package.Workbook.Worksheets.Count = 0 Then
                Throw New InvalidDataException("El archivo Excel no contiene hojas.")
            End If

            Dim ws = package.Workbook.Worksheets(0)

            If ws.Dimension Is Nothing Then
                Throw New InvalidDataException("La hoja seleccionada est� vac�a.")
            End If

            ' Determinamos el rango usado en la hoja
            Dim lastRow As Integer = ws.Dimension.End.Row
            Dim lastCol As Integer = ws.Dimension.End.Column

            If lastRow <= 1 Then
                Throw New InvalidDataException("El archivo Excel no contiene datos (solo encabezado o vac�o).")
            End If

            ' Mapear columnas (soporta variaciones de nombres)
            Dim colNombre As Integer = -1
            Dim colDia As Integer = -1
            Dim colHoraEntrada As Integer = -1
            Dim colHoraSalida As Integer = -1

            ' Recorremos la primera fila para encontrar qué columna corresponde
            ' a cada dato. Hacemos el match de forma tolerante (variantes).
            For col As Integer = 1 To lastCol
                Dim raw = If(ws.Cells(1, col).Value, String.Empty).ToString()
                Dim hdr = NormalizeHeader(raw)

                If hdr = "nombre" OrElse hdr = "nombrecompleto" OrElse hdr = "name" Then
                    colNombre = col
                ElseIf hdr = "dia" OrElse hdr = "fecha" OrElse hdr = "date" Then
                    colDia = col
                ElseIf hdr = "horadeconexion" OrElse hdr = "horaentrada" OrElse hdr = "entrada" OrElse hdr = "hora_de_conexion" OrElse hdr = "timein" Then
                    colHoraEntrada = col
                ElseIf hdr = "horadedesconexion" OrElse hdr = "horasalida" OrElse hdr = "salida" OrElse hdr = "hora_de_desconexion" OrElse hdr = "timeout" Then
                    colHoraSalida = col
                End If
            Next

            If colNombre = -1 OrElse colDia = -1 OrElse colHoraEntrada = -1 OrElse colHoraSalida = -1 Then
                Throw New InvalidDataException("El archivo no contiene todas las columnas requeridas: Nombre, Día, Hora de conexión, Hora de desconexión.")
            End If

            ' Leemos cada fila (empezando en 2 porque la 1 es el encabezado)
            For row As Integer = 2 To lastRow
                Dim rawNombre = ws.Cells(row, colNombre).Value
                Dim rawDia = ws.Cells(row, colDia).Value
                Dim rawEntrada = ws.Cells(row, colHoraEntrada).Value
                Dim rawSalida = ws.Cells(row, colHoraSalida).Value

                Dim nombreStr As String = If(rawNombre IsNot Nothing, rawNombre.ToString().Trim(), String.Empty)
                Dim diaStr As String = FormatCellAsDateOrText(rawDia)
                Dim entradaStr As String = FormatCellAsTimeOrText(rawEntrada)
                Dim salidaStr As String = FormatCellAsTimeOrText(rawSalida)

                ' Ignorar filas que estén totalmente vacías
                If String.IsNullOrWhiteSpace(nombreStr) AndAlso String.IsNullOrWhiteSpace(diaStr) AndAlso String.IsNullOrWhiteSpace(entradaStr) AndAlso String.IsNullOrWhiteSpace(salidaStr) Then
                    Continue For
                End If

                ' Creamos el objeto Conexion y lo añadimos a la lista
                result.Add(New Conexion(nombreStr, diaStr, entradaStr, salidaStr))
            Next
        End Using

        Return result
    End Function

    Private Function NormalizeHeader(s As String) As String
        If s Is Nothing Then Return String.Empty
        Dim t = s.Trim().ToLowerInvariant()
        t = t.Replace(" " , "")
        t = t.Replace("á", "a").Replace("é", "e").Replace("í", "i").Replace("ó", "o").Replace("ú", "u")
        t = t.Replace("Á", "a").Replace("É", "e").Replace("Í", "i").Replace("Ó", "o").Replace("Ú", "u")
        t = t.Replace("ñ", "n").Replace("Ñ", "n")
        t = t.Replace(":", "").Replace("_", "").Replace("-", "")
        Return t
    End Function

    Private Function FormatCellAsDateOrText(value As Object) As String
        If value Is Nothing Then Return String.Empty
        If TypeOf value Is DateTime Then
            Return CType(value, DateTime).ToString("yyyy-MM-dd", CultureInfo.InvariantCulture)
        End If
        Dim s = value.ToString().Trim()
        ' intentar parsear
        Dim dt As DateTime
        If DateTime.TryParse(s, CultureInfo.CurrentCulture, DateTimeStyles.None, dt) Then
            Return dt.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture)
        End If
        Return s
    End Function

    Private Function FormatCellAsTimeOrText(value As Object) As String
        If value Is Nothing Then Return String.Empty
        If TypeOf value Is DateTime Then
            Return CType(value, DateTime).ToString("HH:mm", CultureInfo.InvariantCulture)
        End If
        Dim s = value.ToString().Trim()
        ' intentar parsear hora
        Dim dt As DateTime
        If DateTime.TryParse(s, CultureInfo.CurrentCulture, DateTimeStyles.NoCurrentDateDefault, dt) Then
            Return dt.ToString("HH:mm", CultureInfo.InvariantCulture)
        End If
        Return s
    End Function
End Class