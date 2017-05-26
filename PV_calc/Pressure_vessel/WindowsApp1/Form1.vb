Imports System.IO
Imports System.Text
Imports System.Math
Imports System.Globalization
Imports System.Threading

Public Class Form1
    Public _P As Double         'Calculation pressure [Mpa]
    Public _fs As Double        'Allowable stress shell [N/mm2]
    Public _fp As Double        'Allowable stress reinforcement [N/mm2]

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
        ComboBox2.SelectedIndex = CInt(IIf(ComboBox2.Items.Count > 0, 0, -1)) 'Select ..

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click, NumericUpDown14.ValueChanged, TabPage4.Enter, NumericUpDown12.ValueChanged
        Dim shell_wall, nozzle_wall As Double
        Dim fob, fop As Double
        Dim Afs, Afw As Double
        Dim lso, eas, ris As Double   'Max length shell contibuting to reinforcement
        Dim lbo, eab, rib As Double   'Max length nozzle contibuting to reinforcement
        Dim nozzle_OD, nozzle_ID, Shell_OD, Shell_ID As Double
        Dim Aps, Apb, Afb, Afp, Ap_phi As Double
        Dim eq_left, eq_right, eq_ratio As Double
        Dim Ln, Ln1, Ln2 As Double
        Dim D_small_opening As Double
        Dim small_opening As Boolean

        Shell_OD = NumericUpDown15.Value
        Shell_ID = NumericUpDown18.Value
        shell_wall = NumericUpDown16.Value

        nozzle_OD = NumericUpDown14.Value
        If nozzle_OD >= Shell_OD Then 'Nozzle can not be bigger then shell
            nozzle_OD = Shell_OD
            NumericUpDown14.Value = nozzle_OD
        End If

        nozzle_wall = NumericUpDown12.Value
        nozzle_ID = nozzle_OD - 2 * nozzle_wall
        If nozzle_ID < 10 Then nozzle_ID = 10
        NumericUpDown13.Value = nozzle_ID

        '--------- Small opening 9.5.2.2
        D_small_opening = 0.15 * Sqrt((Shell_ID + shell_wall) * shell_wall)
        Label77.Text = "D= " & D_small_opening.ToString("0.0") & " [mm]"

        '------- reinforment materials is identical to shell material----
        fob = _fs
        fop = _fs

        '---------formula 9.5-2 Shell-----
        eas = shell_wall
        ris = Shell_ID / 2
        lso = Sqrt((2 * ris + eas) + eas)

        '---------formula 9.5-2 Nozzle-----
        eab = nozzle_wall
        rib = nozzle_ID / 2
        lbo = Sqrt((2 * rib + eab) + eab)

        '-------- formula 9.5-7---------------
        Afw = 0                                     'Weld area in neglected
        Afb = nozzle_wall * (lbo + shell_wall)      'Nozzle wall
        Afp = 0                                     'reinforcement ring NOT present
        Afs = lso * shell_wall                      'Shell wall area
        Aps = Shell_ID / 2 * (lso + nozzle_OD / 2)  'Shell area
        Apb = nozzle_ID / 2 * (lbo + shell_wall)    'Nozzle
        Ap_phi = 0                                  'Oblique nozzles

        eq_left = (Afs + Afw) * (_fs - 0.5 * _P)
        eq_left += Afp * (fop - 0.5 * _P)
        eq_left += Afb * (fob - 0.5 * _P)

        eq_right = _P * (Aps + Apb + 0.5 * Ap_phi)
        eq_ratio = eq_left / eq_right

        Ln1 = 0.5 * nozzle_OD + 2 * nozzle_wall  'Equation 9.4-4
        Ln2 = 0.5 * nozzle_OD + 40
        Ln = IIf(Ln1 > Ln2, Ln1, Ln2)

        '----- present--------
        TextBox9.Text = Afs.ToString("0")     'Shell area
        TextBox10.Text = Afw.ToString("0")    'Weld area
        TextBox11.Text = Afb.ToString("0")
        TextBox12.Text = Aps.ToString("0")
        TextBox13.Text = Apb.ToString("0")
        TextBox14.Text = lso.ToString("0.0")
        TextBox15.Text = lbo.ToString("0.0")

        TextBox16.Text = eq_left.ToString("0")
        TextBox17.Text = eq_right.ToString("0")
        TextBox18.Text = eq_ratio.ToString("0.00")
        TextBox19.Text = Ln.ToString("0")
        TextBox20.Text = (Ln - nozzle_OD / 2).ToString("0")
        '----------- check--------
        TextBox16.BackColor = IIf(eq_left < eq_right, Color.Red, Color.LightGreen)
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click, ComboBox1.TextChanged, CheckBox1.CheckedChanged
        Dim sf As Double

        _P = NumericUpDown4.Value                       'Calculation pressure [MPa]
        If (ComboBox1.SelectedIndex > -1) Then          'Prevent exceptions
            Dim words() As String = chap6(ComboBox1.SelectedIndex).Split(separators, StringSplitOptions.None)
            Double.TryParse(words(1), sf)               'Safety factor
            TextBox4.Text = sf.ToString                 'Safety factor
            NumericUpDown7.Value = NumericUpDown10.Value / sf
            If CheckBox1.Checked Then NumericUpDown7.Value *= 0.9   'PED cat IV
            _fs = NumericUpDown7.Value      'allowable stress
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click, NumericUpDown15.ValueChanged, TabPage3.Enter, ComboBox2.SelectedIndexChanged, NumericUpDown16.ValueChanged
        Dim De, Di, Dm, ea, z_joint, e_wall, Pmax, valid_check As Double

        If (ComboBox2.SelectedIndex > -1) Then          'Prevent exceptions
            Double.TryParse(joint_eff(ComboBox2.SelectedIndex), z_joint)      'Joint efficiency
        End If

        De = NumericUpDown15.Value  'OD
        ea = NumericUpDown16.Value  'Wall thicknes
        Di = De - 2 * ea            'ID
        NumericUpDown18.Value = Di

        Dm = (De + Di) / 2          'Average Dia
        Pmax = 2 * _fs * z_joint * ea / Dm               'Max pressure equation 7.4.3 

        e_wall = _P * De / (2 * _fs * z_joint + _P)     'equation 7.4.2 Required wall thickness
        valid_check = Round(e_wall / De, 4)

        '--------- present results--------
        TextBox2.Text = Round(e_wall, 4).ToString  'required wall [mm]
        TextBox5.Text = valid_check.ToString
        TextBox6.Text = _P.ToString
        TextBox7.Text = _fs.ToString
        TextBox8.Text = Round(Pmax, 2).ToString

        '---------- Check-----
        TextBox5.BackColor = IIf(valid_check > 0.16, Color.Red, Color.LightGreen)
    End Sub


End Class
