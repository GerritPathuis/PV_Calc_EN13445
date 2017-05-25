Imports System.IO
Imports System.Text
Imports System.Math


Public Class Form1
    Public _fy As Double        'Yield stress [N/mm2]
    Public _Ym As Double        'Safety factor [-]
    Public _E_mod As Double     'Elasticity [N/mm2]
    Public _G As Double         'Shear modulus

    Dim separators() As String = {";"}
    'Chapter 6, Max allowed values for pressure parts
    Public Shared chap6() As String = {
   "Chap 6.2, Steel, safety, rupture < 30%; 1.5",
   "Chap 6.4, Austenitic steel, rupture 30-35%; 1.5",
   "Chap 6.5, Austenitic steel, rupture >35%; 3.0",
   "Chap 6.6, Cast steel; 1.9"}

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Dim words() As String

        TextBox1.Text =
        "Based on " & vbCrLf &
        "EN 13445" & vbCrLf &
        "Buckling Strength of plated structures" & vbCrLf &
        "https://rules.dnvgl.com/servicedocuments/dnv" & vbCrLf & vbCrLf &
        "See also" & vbCrLf &
        "http://www.steelconstruction.info/Stiffeners"

        ComboBox1.Items.Clear()
        For hh = 0 To (chap6.Length - 1)  'Fill combobox3 arrangment data
            words = chap6(hh).Split(separators, StringSplitOptions.None)
            ComboBox1.Items.Add(words(0))
        Next hh
        '----------------- prevent out of bounds------------------
        ComboBox1.SelectedIndex = CInt(IIf(ComboBox1.Items.Count > 0, 1, -1)) 'Select ..
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click, NumericUpDown18.ValueChanged, NumericUpDown15.ValueChanged, NumericUpDown14.ValueChanged, NumericUpDown13.ValueChanged
        Dim shell_wall, nozzle_wall As Double

        shell_wall = (NumericUpDown15.Value - NumericUpDown18.Value) / 2
        nozzle_wall = (NumericUpDown14.Value - NumericUpDown13.Value) / 2

        NumericUpDown16.Value = shell_wall
        NumericUpDown12.Value = nozzle_wall
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If (ComboBox1.SelectedIndex > -1) Then      'Prevent exceptions
            Dim words() As String = chap6(ComboBox1.SelectedIndex).Split(separators, StringSplitOptions.None)
            ' TextBox33.Text = LTrim(words(6))     'Density steel
        End If
    End Sub
End Class
