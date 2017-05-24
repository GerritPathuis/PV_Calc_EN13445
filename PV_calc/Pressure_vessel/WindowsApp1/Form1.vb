Imports System.IO
Imports System.Text
Imports System.Math


Public Class Form1
    Public _fy As Double        'Yield stress [N/mm2]
    Public _Ym As Double        'Safety factor [-]
    Public _E_mod As Double     'Elasticity [N/mm2]
    Public _G As Double         'Shear modulus


    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        TextBox1.Text =
        "Based on " & vbCrLf &
        "EN 13445" & vbCrLf &
        "Buckling Strength of plated structures" & vbCrLf &
        "https://rules.dnvgl.com/servicedocuments/dnv" & vbCrLf & vbCrLf &
        "See also" & vbCrLf &
        "http://www.steelconstruction.info/Stiffeners"
    End Sub

End Class
