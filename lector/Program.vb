Option Strict On
Option Explicit On

Module Program
    Sub Main(args As String())
        ' Entrada del programa. Solo llamamos al módulo App para ejecutar la
        ' lógica principal. Si hay un error raro, lo mostramos en pantalla.
        Try
            App.Run()
        Catch ex As Exception
            Console.WriteLine("Error inesperado: " & ex.Message)
        End Try
    End Sub
End Module
