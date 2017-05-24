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

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click, NumericUpDown18.ValueChanged, NumericUpDown15.ValueChanged, NumericUpDown14.ValueChanged, NumericUpDown13.ValueChanged
        Dim shell_wall, nozzle_wall As Double

        shell_wall = (NumericUpDown15.Value - NumericUpDown18.Value) / 2
        nozzle_wall = (NumericUpDown14.Value - NumericUpDown13.Value) / 2

        NumericUpDown16.Value = shell_wall
        NumericUpDown12.Value = nozzle_wall
    End Sub
End Class
