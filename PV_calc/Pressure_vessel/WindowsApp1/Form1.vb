Imports System.IO
Imports System.Text
Imports System.Math
Imports System.Globalization
Imports System.Threading

Public Class Form1
    Public _P As Double         'Calculation pressure [Mpa]
    Public _fs As Double        'Allowable stress [N/mm2]


    Dim separators() As String = {";"}

    'Chapter 6, Max allowed values for pressure parts
    Public Shared chap6() As String = {
   "Chap 6.2, Steel, safety, rupture < 30%; 1.5",
   "Chap 6.4, Austenitic steel, rupture 30-35%; 1.5",
   "Chap 6.5, Austenitic steel, rupture >35%; 3.0",
   "Chap 6.6, Cast steel; 1.9"}

    Public Shared joint_eff() As String = {"  0.7", "  0.85", "  1.0"}

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim words() As String

        Thread.CurrentThread.CurrentCulture = New CultureInfo("en-US")
        Thread.CurrentThread.CurrentUICulture = New CultureInfo("en-US")

        TextBox1.Text =
        "Based on " & vbCrLf &
        "EN 13445" & vbCrLf &
        "Buckling Strength of plated structures" & vbCrLf &
        "https://rules.dnvgl.com/servicedocuments/dnv" & vbCrLf & vbCrLf &
        "See also" & vbCrLf &
        "http://www.steelconstruction.info/Stiffeners"

        ComboBox1.Items.Clear()
        For hh = 0 To (chap6.Length - 1)  'Fill combobox1 materials
            words = chap6(hh).Split(separators, StringSplitOptions.None)
            ComboBox1.Items.Add(words(0))
        Next hh

        ComboBox2.Items.Clear()
        For hh = 0 To (joint_eff.Length - 1)   'Fill combobox2 joint efficiency
            ComboBox2.Items.Add(joint_eff(hh))
        Next hh

        '----------------- prevent out of bounds------------------
        ComboBox1.SelectedIndex = CInt(IIf(ComboBox1.Items.Count > 0, 1, -1)) 'Select ..
        ComboBox2.SelectedIndex = CInt(IIf(ComboBox2.Items.Count > 0, 1, -1)) 'Select ..

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click, NumericUpDown14.ValueChanged, NumericUpDown13.ValueChanged
        Dim shell_wall, nozzle_wall As Double

        shell_wall = (NumericUpDown15.Value - NumericUpDown18.Value) / 2
        nozzle_wall = (NumericUpDown14.Value - NumericUpDown13.Value) / 2

        NumericUpDown16.Value = shell_wall
        NumericUpDown12.Value = nozzle_wall
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click, ComboBox1.TextChanged, CheckBox1.CheckedChanged
        Dim sf As Double

        _P = NumericUpDown4.Value       'Calculation pressure [MPa]
        If (ComboBox1.SelectedIndex > -1) Then          'Prevent exceptions
            Dim words() As String = chap6(ComboBox1.SelectedIndex).Split(separators, StringSplitOptions.None)
            Double.TryParse(words(1), sf)               'Safety factor
            TextBox4.Text = sf.ToString                 'Safety factor
            NumericUpDown7.Value = NumericUpDown10.Value / sf
            If CheckBox1.Checked Then NumericUpDown7.Value *= 0.9   'PED cat IV
            _fs = NumericUpDown7.Value      'allowable stress
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click, NumericUpDown18.ValueChanged, NumericUpDown15.ValueChanged, TabPage3.Enter, ComboBox2.SelectedIndexChanged
        Dim De, Di, Dm, ea, z_joint, e_wall, Pmax As Double

        If (ComboBox2.SelectedIndex > -1) Then          'Prevent exceptions
            Double.TryParse(joint_eff(ComboBox2.SelectedIndex), z_joint)      'Joint efficiency
        End If

        De = NumericUpDown15.Value  'OD
        Di = NumericUpDown18.Value  'ID
        Dm = (De + Di) / 2          'Average Dia
        ea = (De - Di) / 2          'Wall thicknes
        NumericUpDown16.Value = ea

        Pmax = 2 * _fs * z_joint * ea / Dm               'Max pressure equation 7.4.3 


        e_wall = _P * De / (2 * _fs * z_joint + _P)     'equation 7.4.2 Required wall thickness



        TextBox2.Text = Round(e_wall, 4).ToString  'required wall [mm]
        TextBox6.Text = _P.ToString
        TextBox7.Text = _fs.ToString
        TextBox8.Text = Round(Pmax, 2).ToString
    End Sub

    Private Sub GroupBox3_Enter(sender As Object, e As EventArgs) Handles GroupBox3.Enter

    End Sub
End Class
