Option Strict On
Option Explicit On

Imports System

Public Module App
    Public Sub Run()
        ' Punto central de la aplicación: pide la ruta y muestra la tabla
        Dim loader As New ExcelLoader()

        Do
            Console.WriteLine()
            Console.Write("Ruta del archivo Excel (o Enter para salir): ")
            Dim path = Console.ReadLine().Trim()

            ' Si el usuario no escribe nada, salimos
            If String.IsNullOrEmpty(path) Then
                Exit Do
            End If

            Try
                ' Intentamos cargar el archivo y mostrar los registros
                Dim lista = loader.Load(path)
                If lista.Count = 0 Then
                    Console.WriteLine("No se encontraron registros.")
                Else
                    PrintTable(lista)
                End If
            Catch ex As Exception
                ' Mostramos el error en pantalla y el usuario puede intentar otra vez
                Console.WriteLine("Error: " & ex.Message)
            End Try

            Console.WriteLine()
            Console.Write("¿Desea cargar otro archivo? (S para sí/N para no): ")
            Dim respuesta = Console.ReadLine().Trim().ToUpperInvariant()
            If Not respuesta.Equals("S") Then Exit Do
        Loop
    End Sub

    Private Sub PrintTable(items As List(Of Conexion))
        ' Encabezado
        Dim c1 = 30 ' Nombre
        Dim c2 = 12 ' Dia
        Dim c3 = 12 ' Hora entrada
        Dim c4 = 12 ' Hora salida

        Console.WriteLine()
        ' Imprimimos un encabezado simple y una línea separadora
        Console.WriteLine($"{PadRight("Nombre", c1)}{PadRight("Día", c2)}{PadRight("Entrada", c3)}{PadRight("Salida", c4)}")
        Console.WriteLine(New String("-"c, c1 + c2 + c3 + c4))

        ' Recorremos los objetos y los mostramos en formato de columnas
        For Each it In items
            Console.WriteLine($"{PadRight(it.Nombre, c1)}{PadRight(it.Dia, c2)}{PadRight(it.HoraEntrada, c3)}{PadRight(it.HoraSalida, c4)}")
        Next
    End Sub

    Private Function PadRight(s As String, totalWidth As Integer) As String
        If s Is Nothing Then s = String.Empty
        If s.Length >= totalWidth Then
            Return s.Substring(0, totalWidth - 1) & " "
        End If
        Return s.PadRight(totalWidth)
    End Function
End Module