Option Strict On
Option Explicit On

Public Class Conexion
    ' Campos privados (encapsulamiento) - se acceden mediante propiedades
    Private _nombre As String
    Private _dia As String
    Private _horaEntrada As String
    Private _horaSalida As String

    Public Sub New(nombre As String, dia As String, horaEntrada As String, horaSalida As String)
        ' Constructor simple: guarda los valores recibidos
        ' Si viene Nothing, lo convertimos a cadena vacía para evitar errores
        _nombre = If(nombre, String.Empty)
        _dia = If(dia, String.Empty)
        _horaEntrada = If(horaEntrada, String.Empty)
        _horaSalida = If(horaSalida, String.Empty)
    End Sub

    Public Property Nombre As String
        Get
            Return _nombre
        End Get
        Set(value As String)
            ' Usamos la propiedad para controlar el acceso al campo privado
            _nombre = If(value, String.Empty)
        End Set
    End Property

    Public Property Dia As String
        Get
            Return _dia
        End Get
        Set(value As String)
            _dia = If(value, String.Empty)
        End Set
    End Property

    Public Property HoraEntrada As String
        Get
            Return _horaEntrada
        End Get
        Set(value As String)
            _horaEntrada = If(value, String.Empty)
        End Set
    End Property

    Public Property HoraSalida As String
        Get
            Return _horaSalida
        End Get
        Set(value As String)
            _horaSalida = If(value, String.Empty)
        End Set
    End Property
End Class