Imports System.IO
Imports System.Text
Imports System.Math
Imports System.Globalization
Imports System.Threading

'-------------------------------------------------------
'Pressure vessel calculation according to EN 13445
'Unfired pressure vessels part 3
'-------------------------------------------------------
Public Class Form1
    Public _P As Double         'Calculation pressure [Mpa]
    Public _fs As Double        'Allowable stress shell [N/mm2]
    Public _fp As Double        'Allowable stress reinforcement [N/mm2]

    Public _De As Double        'Outside diameter shell
    Public _Di As Double        'Inside diameter shell
    Public _ecs As Double       'Shell thickness

    Public _deb As Double       'Outside diameter nozzle fitted in shell
    Public _dib As Double       'Inside diameter nozzle fitted in shell
    Public _eb As Double        'Effective thickness nozzle thickness

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
        "Unfired pressure vessels" & vbCrLf &
        "Part 3 Design (issue 3:2016)"

        TextBox22.Text =
        "Important note" & vbCrLf &
        "The yield strength is detremined from the mill certificate" & vbCrLf &
        "given at the operating temperature. Different for piping and sheet steel" & vbCrLf &
        "Safety factors follow the Eurocode"


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

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click, NumericUpDown14.ValueChanged, TabPage4.Enter, NumericUpDown12.ValueChanged, NumericUpDown13.ValueChanged, NumericUpDown1.ValueChanged
        Calc_nozzle_fig949()
    End Sub
    Private Sub Calc_nozzle_fig949()
        Dim nozzle_wall As Double
        Dim fob, fop As Double
        Dim Afs, Afw As Double
        Dim ls As Double
        Dim lso, eas, ris As Double   'Max length shell contibuting to reinforcement
        Dim lbo, eab, rib As Double   'Max length nozzle contibuting to reinforcement

        Dim Aps, Apb, Afb, Afp, Ap_phi As Double
        Dim eq_left, eq_right, eq_ratio As Double
        Dim Ln, Ln1, Ln2 As Double
        Dim D_small_opening As Double
        Dim W_min, W_min1, W_min2 As Double

        ls = NumericUpDown1.Value           'Distance shell-edge-opening to discontinuity
        _De = NumericUpDown15.Value         'Shell OD 
        _Di = NumericUpDown18.Value         'Shell ID 
        _eb = NumericUpDown16.Value         'Shell Wall 
        _deb = NumericUpDown14.Value        'Outside diameter nozzle fitted in shell
        eas = _eb                           'Shell Analysis thickness of shell wall 

        If _deb >= _De Then       'Nozzle dia can not be bigger then shell
            _deb = _De
            NumericUpDown14.Value = _deb
        End If

        nozzle_wall = NumericUpDown12.Value
        _dib = _deb - 2 * nozzle_wall
        If _dib < 10 Then _dib = 10
        NumericUpDown13.Value = _dib

        '--------- Small opening 9.5.2.2
        D_small_opening = 0.15 * Sqrt((_Di + _eb) * _eb)
        Label77.Text = "D= " & D_small_opening.ToString("0.0") & " [mm]"

        '------- reinforment materials is identical to shell material----
        fob = _fs
        fop = _fs

        '---------formula 9.5-2 Nozzle-----
        eab = nozzle_wall
        rib = _dib / 2
        lbo = Sqrt((2 * rib + eab) + eab)

        '----Chapter 9.5.2.4.4.3 Nozzle in cylindrical shell
        Dim a, ls_min, ecs As Double
        ecs = eas                                   'assumed shell thickness for calculation
        a = _deb / 2                                'equation 9.5-90
        ris = (_De / 2) - eas                       'equation 9.5-91
        lso = Sqrt(((_De - 2 * eas) + ecs) * ecs)   'equation 9.5-92
        ls_min = IIf(lso < ls, lso, ls)             'equation 9.5-93
        Aps = ris * (ls_min + a)                    'equation 9.5-94

        '----------------------- formula 9.5-7 -----------------------
        'Af = Stress loaded cross-sectional area effective as reinforcement
        Afw = 0                                     'Weld area in neglected
        Afb = nozzle_wall * (lbo + _eb)      'Nozzle wall
        Afp = 0                                     'reinforcement ring NOT present
        Afs = lso * _eb                      'Shell wall area

        'Ap = Pressure loaded area. 
        Apb = _dib / 2 * (lbo + _eb)         'Nozzle Pressure loaded area
        Ap_phi = 0                                  'Oblique nozzles

        eq_left = (Afs + Afw) * (_fs - 0.5 * _P)
        eq_left += Afp * (fop - 0.5 * _P)
        eq_left += Afb * (fob - 0.5 * _P)

        eq_right = _P * (Aps + Apb + 0.5 * Ap_phi)
        eq_ratio = eq_left / eq_right

        '9.4.8 Minimum Distance between nozle and shell butt-weld
        Ln1 = 0.5 * _deb + 2 * nozzle_wall       'Equation 9.4-4
        Ln2 = 0.5 * _deb + 40
        Ln = IIf(Ln1 > Ln2, Ln1, Ln2)

        'Figure 9.7-5          
        W_min1 = 0.2 * Sqrt((2 * _Di * 0.5 + ecs) * ecs)    'equation 9.7-5
        W_min2 = 3 * eas                                    'equation 9.7-5
        W_min = IIf(W_min1 > W_min2, W_min1, W_min2)        'Find biggest

        '----- present--------
        TextBox9.Text = Afs.ToString("0")       'Shell area reinforcement [mm2]
        TextBox10.Text = Afw.ToString("0")      'Weld area reinforcement [mm2]
        TextBox11.Text = Afb.ToString("0")      'reinforcement [mm2]
        TextBox12.Text = Aps.ToString("0")    'Pressure loaded area [mm2]
        TextBox13.Text = Apb.ToString("0")      'Pressure loaded area [mm2]
        TextBox14.Text = lso.ToString("0.0")
        TextBox15.Text = lbo.ToString("0.0")

        TextBox16.Text = eq_left.ToString("0")
        TextBox17.Text = eq_right.ToString("0")
        TextBox18.Text = eq_ratio.ToString("0.00")
        TextBox19.Text = W_min.ToString("0")
        TextBox20.Text = (Ln - _deb / 2).ToString("0")
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
        Calc_shell742()
    End Sub
    '7.4.2 Cylindrical shells 
    Private Sub Calc_shell742()
        Dim De, Di, Dm, ea, z_joint, e_wall, Pmax, valid_check As Double

        If (ComboBox2.SelectedIndex > -1) Then          'Prevent exceptions
            Double.TryParse(joint_eff(ComboBox2.SelectedIndex), z_joint)      'Joint efficiency
        End If

        De = NumericUpDown15.Value  'OD
        ea = NumericUpDown16.Value  'Wall thicknes
        Di = De - 2 * ea            'ID
        NumericUpDown18.Value = Di

        Dm = (De + Di) / 2                          'Average Dia
        Pmax = 2 * _fs * z_joint * ea / Dm          'Max pressure equation 7.4.3 

        e_wall = _P * De / (2 * _fs * z_joint + _P) 'equation 7.4.2 Required wall thickness
        valid_check = Round(e_wall / De, 4)

        '--------- present results--------
        TextBox2.Text = Round(e_wall, 4).ToString   'required wall [mm]
        TextBox5.Text = valid_check.ToString
        TextBox6.Text = _P.ToString
        TextBox7.Text = _fs.ToString
        TextBox8.Text = Round(Pmax, 2).ToString

        '---------- Check-----
        TextBox5.BackColor = IIf(valid_check > 0.16, Color.Red, Color.LightGreen)
    End Sub
    'Chapter 15 rectangle shell
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Dim P, a, ee, L, L1, σmc, σmd, σmb As Double

        P = NumericUpDown15.Value
        a = NumericUpDown15.Value
        ee = NumericUpDown15.Value
        L = NumericUpDown15.Value
        L1 = NumericUpDown15.Value

        'At C
        σmc = P * (a + L) / ee                '(15.5.1.2-1) 

        'At D
        σmd = σmc

        'At B
        σmb = P / ee * (a + Sqrt(L ^ 2 + L1 ^ 2))   '(15.5.1.2-3) 

    End Sub
End Class
