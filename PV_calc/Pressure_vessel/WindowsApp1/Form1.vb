Imports System.IO
Imports System.Text
Imports System.Math
Imports System.Globalization
Imports System.Threading
Imports Word = Microsoft.Office.Interop.Word

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

    '----------- directory's-----------
    Dim dirpath_Eng As String = "N:\Engineering\VBasic\PV_calc_input\"
    Dim dirpath_Rap As String = "N:\Engineering\VBasic\PV_calc_rapport_copy\"
    Dim dirpath_Home As String = "C:\Temp\"

    'Chapter 6, Max allowed values for pressure parts
    Public Shared chap6() As String = {
   "Chap 6.2, Steel, safety, rupture < 30%; 1.5",
   "Chap 6.4, Austenitic steel, rupture 30-35%; 1.5",
   "Chap 6.5, Austenitic steel, rupture >35%; 3.0",
   "Chap 6.6, Cast steel; 1.9"}

    'EN13455-3 ANNEX H
    Public Shared gaskets() As String = {
   "Rubber without fabric < 75 IRH;0.50;0000;0",
   "Rubber without fabric > 75 IRH;1.00;1.40;0",
   "Asbestos with binder 3.2mm    ;2.00;11.0;0",
   "Asbestos with binder 1.6mm    ;2.75;25.5;0",
   "Asbestos with binder 0.8mm    ;3.50;44.8;0",
   "Spiral wound asbestos filled  ;2.50;69.0;0",
   "Currogated copper or brass    ;3.00;31.0;0",
   "Grooved soft aluminium        ;3.25;37.9;0",
   "Rubber o-ring < 75 IRH        ;0.25;0.70;0",
   "Rubber o-ring > 75 IRH        ;0.25;1.40;0",
   "Square Rubber ring < 75 IRH   ;0.25;1.00;0",
   "Square Rubber ring > 75 IRH   ;0.25;2.80;0",
   "NO seal                       ;0;0;0"}

    'EN 10028-2 for steel
    'EN 10028-7 for stainless steel
    Public Shared steel() As String = {
   "Material-------;50c;100;150;200;250;300;350;400;450;500;550;remarks--",
   "1.0425 (P265GH);265;241;223;205;188;173;160;150;  0;  0;  0; max 400c",
   "1.0473 (P355GH);343;323;299;275;252;232;214;202;  0;  0;  0; max 400c",
   "1.4301 (304)   ;190;157;142;127;118;110;104; 98; 95; 92; 90; max 550c",
   "1.4307 (304L)  ;180;147;132;118;108;100; 94; 89; 85; 81; 80; max 550c",
   "1.4401 (316)   ;204;177;162;147;137;127;120;115;112;110;108; max 550c",
   "1.4404 (316L)  ;200;166;152;137;127;118;113;108;103;100; 98; max 550c"}

    'EN 1993-1-8 Bolts (Eurocode 3)
    Public Shared Bolt() As String = {
   "Bolt class-----;  ultimate;yield 0.2",
   "Bolt class 4.6 ;  400;240",
   "Bolt class 5.6 ;  500;300",
   "Bolt class 8.8 ;  800;640",
   "Bolt class 10.9; 1000;900",
   "Bolt class A2-60; 600;210",
   "Bolt class A2-70; 700;450",
   "Bolt class A2-80; 800;600",
   "Bolt class A4-60; 600;210",
   "Bolt class A4-70; 700;450",
   "Bolt class A4-80; 800;600"}

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
        For hh = 0 To (chap6.Length - 1)  'Fill combobox 
            words = chap6(hh).Split(separators, StringSplitOptions.None)
            ComboBox1.Items.Add(words(0))
        Next hh

        ComboBox2.Items.Clear()
        For hh = 0 To (joint_eff.Length - 1)   'Fill combobox joint efficiency
            ComboBox2.Items.Add(joint_eff(hh))
            ComboBox3.Items.Add(joint_eff(hh))
        Next hh

        ComboBox4.Items.Clear()
        For hh = 0 To (gaskets.Length - 1)  'Fill combobox gasket materials
            words = gaskets(hh).Split(separators, StringSplitOptions.None)
            ComboBox4.Items.Add(words(0))
        Next hh

        ComboBox5.Items.Clear()
        For hh = 1 To (steel.Length - 1)  'Fill combobox steel
            words = steel(hh).Split(separators, StringSplitOptions.None)
            ComboBox5.Items.Add(words(0))
        Next hh

        ComboBox6.Items.Clear()
        For hh = 1 To (Bolt.Length - 1)  'Fill combobox steel
            words = Bolt(hh).Split(separators, StringSplitOptions.None)
            ComboBox6.Items.Add(words(0))
        Next hh
        '----------------- prevent out of bounds------------------
        ComboBox1.SelectedIndex = CInt(IIf(ComboBox1.Items.Count > 0, 1, -1)) 'Select ..
        ComboBox2.SelectedIndex = CInt(IIf(ComboBox2.Items.Count > 0, 0, -1)) 'Select ..
        ComboBox3.SelectedIndex = CInt(IIf(ComboBox3.Items.Count > 0, 0, -1)) 'Select ..
        ComboBox4.SelectedIndex = CInt(IIf(ComboBox4.Items.Count > 0, 1, -1)) 'Select ..
        ComboBox5.SelectedIndex = CInt(IIf(ComboBox5.Items.Count > 0, 0, -1)) 'Select ..
        ComboBox6.SelectedIndex = CInt(IIf(ComboBox6.Items.Count > 0, 0, -1)) 'Select ..
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
        TextBox12.Text = Aps.ToString("0")      'Pressure loaded area [mm2]
        TextBox13.Text = Apb.ToString("0")      'Pressure loaded area [mm2]
        TextBox14.Text = lso.ToString("0.0")
        TextBox15.Text = lbo.ToString("0.0")

        TextBox16.Text = eq_left.ToString("0")
        TextBox17.Text = eq_right.ToString("0")
        TextBox18.Text = eq_ratio.ToString("0.00")
        TextBox19.Text = W_min.ToString("0")
        TextBox20.Text = (Ln - _deb / 2).ToString("0")
        '----------- checks--------
        TextBox16.BackColor = IIf(eq_left < eq_right, Color.Red, Color.LightGreen)
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click, ComboBox1.TextChanged, CheckBox1.CheckedChanged, NumericUpDown5.ValueChanged, NumericUpDown4.ValueChanged, NumericUpDown10.ValueChanged, ComboBox5.SelectedIndexChanged
        Design_stress()
    End Sub
    Private Sub Design_stress()
        Dim sf, temp As Double
        Dim words() As String
        Dim y50, y100, y150, y200, y250, y300, y350, y400 As Double

        If (ComboBox5.SelectedIndex > -1) Then          'Prevent exceptions
            words = steel(ComboBox5.SelectedIndex + 1).Split(separators, StringSplitOptions.None)
            TextBox3.Text = words(1)

            TextBox104.Text = words(2)
            TextBox105.Text = words(3)
            TextBox106.Text = words(4)
            TextBox107.Text = words(5)
            TextBox108.Text = words(6)
            TextBox109.Text = words(7)
            TextBox110.Text = words(8)
            Double.TryParse(words(1), y50)
            Double.TryParse(words(2), y100)
            Double.TryParse(words(3), y150)
            Double.TryParse(words(4), y200)
            Double.TryParse(words(5), y250)
            Double.TryParse(words(6), y300)
            Double.TryParse(words(7), y350)
            Double.TryParse(words(8), y400)

            temp = NumericUpDown5.Value
            Select Case True
                Case 50 > temp
                    NumericUpDown10.Value = y50
                Case 100 > temp
                    NumericUpDown10.Value = y100
                Case 150 > temp
                    NumericUpDown10.Value = y150
                Case 200 > temp
                    NumericUpDown10.Value = y200
                Case 250 > temp
                    NumericUpDown10.Value = y250
                Case 300 > temp
                    NumericUpDown10.Value = y300
                Case 350 > temp
                    NumericUpDown10.Value = y350
                Case 400 > temp
                    NumericUpDown10.Value = y400
                Case temp > 400
                    MessageBox.Show("Problem temp too high")
            End Select
        End If

        _P = NumericUpDown4.Value                       'Calculation pressure [MPa=N/mm2]
        If (ComboBox1.SelectedIndex > -1) Then          'Prevent exceptions
            words = chap6(ComboBox1.SelectedIndex).Split(separators, StringSplitOptions.None)
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
        TextBox53.Text = (_P * 10).ToString
        TextBox7.Text = _fs.ToString
        TextBox8.Text = Round(Pmax, 2).ToString

        '---------- Check-----
        TextBox5.BackColor = IIf(valid_check > 0.16, Color.Red, Color.LightGreen)
    End Sub
    'Chapter 15 rectangle shell
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click, NumericUpDown9.ValueChanged, NumericUpDown8.ValueChanged, NumericUpDown17.ValueChanged, NumericUpDown11.ValueChanged, TabPage2.Enter
        Dim a, ee, L, l1, I1 As Double
        Dim σmD, σmC, σmB, σmA, σmBC As Double  'membrane stress
        Dim σbD, σbC, σbB, σbA, σbBC As Double  'bending stress
        Dim σTD, σTC, σTB, σTA, σTBC As Double  'Total stress
        Dim α3, φ, θ, K3, Ma As Double

        _P = NumericUpDown4.Value   'Pressure
        a = NumericUpDown17.Value   'Inside corner radius
        ee = NumericUpDown11.Value  'wall thickness
        L = NumericUpDown8.Value    'Lenght
        l1 = NumericUpDown9.Value   'Lenght


        '-------- membrane stress----------------
        σmC = _P * (a + L) / ee                     'At C (eq 15.5.1.2-1) 
        σmD = σmC                                   'At D
        σmB = _P * (a + l1) / ee                    'At B (eq 15.5.1.2-2) 
        σmA = σmB                                   'At A
        σmBC = (_P / ee) * (a + Sqrt(L ^ 2 + l1 ^ 2)) 'At corner (eq 15.5.1.2-3) 

        '------- bending stress----------------
        I1 = ee ^ 3 / 12    '(eq 15.5.1.2-4) second moment of area
        α3 = L / l1         '(eq 15.5.1.2-14) factor
        φ = a / l1          '(eq 15.5.1.2-15) angular indication

        K3 = 6 * φ ^ 2 * α3                 '(eq 15.5.1.2-12)
        K3 -= 3 * PI * φ ^ 2
        K3 += 6 * φ ^ 2
        K3 += α3 ^ 3
        K3 += 3 * α3 ^ 2
        K3 -= 6 * φ
        K3 -= 2
        K3 += 1.5 * PI * α3 ^ 2 * φ
        K3 += 6 * φ * α3
        K3 *= l1 ^ 2
        K3 /= 3 * (2 * α3 + PI * φ + 2)     'factor unreinforced vessel

        Ma = _P * K3                   'bending moment middle of side


        '------- bend in the corner -----
        θ = Atan(l1 / L)                    '(eq 15.5.1.2-10) max value

        σbBC = 2 * a * (L * Cos(θ) - l1 * (1 - Sin(θ)))
        σbBC += L ^ 2
        σbBC *= _P
        σbBC += 2 * Ma
        σbBC *= ee / (4 * I1)               '(eq 15.5.1.2-9)

        '------- bend at B -------
        σbB += (ee / (4 * I1)) * (2 * Ma + _P * L ^ 2)    '(eq 15.5.1.2-8)

        '------- bend at A -------
        σbA += Ma * ee / (2 * I1)           '(eq 15.5.1.2-7)

        '------- bend at D -------
        σbD = _P * (2 * a * L - 2 * a * l1 + L ^ 2 - l1 ^ 2)
        σbD += 2 * Ma
        σbD *= ee / (4 * I1)                '(eq 15.5.1.2-6)

        '------- bend at C -------
        σbC = _P * (2 * a * L - 2 * a * l1 + L ^ 2)
        σbC += 2 * Ma
        σbC *= ee / (4 * I1)                '(eq 15.5.1.2-5)

        '----- total stress---
        σTA = Abs(σmA) + Abs(σbA)
        σTB = Abs(σmB) + Abs(σbB)
        σTC = Abs(σmC) + Abs(σbC)
        σTD = Abs(σmD) + Abs(σbD)
        σTBC = Abs(σmBC) + Abs(σbBC)

        '----- pressure -----
        TextBox25.Text = _P.ToString
        '---- membrane stress ---
        TextBox23.Text = σmA.ToString("0.0")
        TextBox24.Text = σmB.ToString("0.0")
        TextBox28.Text = σmC.ToString("0.0")
        TextBox29.Text = σmD.ToString("0.0")
        TextBox30.Text = σmBC.ToString("0.0")   'Bend
        '---- bending stress ---
        TextBox31.Text = σbA.ToString("0.0")
        TextBox32.Text = σbB.ToString("0.0")
        TextBox33.Text = σbC.ToString("0.0")
        TextBox34.Text = σbD.ToString("0.0")
        TextBox35.Text = σbBC.ToString("0.0")   'Bend

        '---- total stress ---
        TextBox26.Text = σTA.ToString("0.0")
        TextBox27.Text = σTB.ToString("0.0")
        TextBox36.Text = σTC.ToString("0.0")
        TextBox37.Text = σTD.ToString("0.0")
        TextBox38.Text = σTBC.ToString("0.0")   'Bend

        '---- check '(eq 15.5.3-2)---
        TextBox26.BackColor = IIf(σTA > 1.5 * _fs, Color.Red, Color.LightGreen)
        TextBox27.BackColor = IIf(σTB > 1.5 * _fs, Color.Red, Color.LightGreen)
        TextBox36.BackColor = IIf(σTC > 1.5 * _fs, Color.Red, Color.LightGreen)
        TextBox37.BackColor = IIf(σTD > 1.5 * _fs, Color.Red, Color.LightGreen)
        TextBox38.BackColor = IIf(σTBC > 1.5 * _fs, Color.Red, Color.LightGreen)

        '----- vessel size
        TextBox40.Text = (L + a) * 2.ToString("0.0")
        TextBox41.Text = (l1 + a) * 2.ToString("0.0")

    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click, NumericUpDown3.ValueChanged, NumericUpDown2.ValueChanged, GroupBox11.Enter, ComboBox3.SelectedIndexChanged
        ' Design_stress()
        Calc_kloepper()
    End Sub

    Private Sub Calc_kloepper()
        Dim De, Di, Dm, Wall, r_knuckle As Double
        Dim Es, Ey, Eb As Double    'Wall thicknes
        Dim E_kloepper As Double    'Required Wall thicknes
        Dim R_central As Double     'is inside spherical radius of central part kloepper
        Dim z_joint As Double
        Dim β As Double 'Figure 7.5-1 or 7.5.3.5, replacing e by ey

        De = NumericUpDown3.Value   'Outside dia shell
        Wall = NumericUpDown2.Value 'Wall shell
        Di = De - 2 * Wall          'Inside dia shell
        Dm = De - Wall              'Mean dia shell
        r_knuckle = 0.1 * De        'radius knuckle
        R_central = De              'radius central part head


        If (ComboBox3.SelectedIndex > -1) Then          'Prevent exceptions
            Double.TryParse(joint_eff(ComboBox3.SelectedIndex), z_joint)  'Joint efficiency
        End If

        'Wall thickness  membrane stress in central part
        Es = _P * R_central / (2 * _fs * z_joint - 0.5 * _P)

        'Wall knuckle to avoid axisymmetric yielding
        β = Calc_kloepper_beta()
        Ey = β * _P * (0.75 * R_central + 0.2 * Di) / _fs

        'Wall thickness  knuckle to avoid plastic buckling
        Eb = (Di / r_knuckle) ^ 0.825
        Eb *= (_P / (111 * _fs))
        Eb = Eb ^ (1 / 1.5)
        Eb *= (0.75 * R_central + 0.2 * Di)

        'Find the thickest wall
        E_kloepper = 0
        If Es > E_kloepper Then E_kloepper = Es
        If Ey > E_kloepper Then E_kloepper = Ey
        If Eb > E_kloepper Then E_kloepper = Eb

        TextBox45.Text = _P.ToString("0.00")
        TextBox51.Text = (_P * 10).ToString("0.0")
        TextBox44.Text = Di.ToString("0")
        TextBox50.Text = Dm.ToString("0")
        TextBox49.Text = r_knuckle.ToString("0")
        TextBox42.Text = R_central.ToString("0")
        TextBox43.Text = _fs.ToString("0")

        TextBox39.Text = Es.ToString("0.0")
        TextBox46.Text = Ey.ToString("0.0")
        TextBox47.Text = Eb.ToString("0.0")
        TextBox48.Text = E_kloepper.ToString("0.0")
        TextBox52.Text = β.ToString("0.0")

        'Checks
        TextBox48.BackColor = IIf(E_kloepper > Wall, Color.Red, Color.LightGreen)

    End Sub
    Private Function Calc_kloepper_beta()
        '7.5.3.5 Formulae for calculation of factor E
        Dim k_e, k_Di, k_De, k_R, k_rr, k_N As Double
        Dim X, Y, Z, β, β006, β01, β02 As Double

        k_e = NumericUpDown2.Value              'Wall tckickness
        k_De = NumericUpDown3.Value             'Outside doameter
        Double.TryParse(TextBox44.Text, k_Di)   'Inside diamter
        Double.TryParse(TextBox42.Text, k_R)    'Radius central part
        Double.TryParse(TextBox49.Text, k_rr)   'Radius knuckle

        Y = k_e / k_R                           '(7.5-9) 
        If Y > 0.04 Then Y = 0.04

        Z = Log10(1 / Y)                        '(7.5-10) 
        X = k_rr / k_Di                         '(7.5-11) 
        k_N = 1.006 - (1 / (6.2 + (90 * Y) ^ 4)) '(7.5-12) 

        β006 = k_N * (-0.3635 * Z ^ 3 + 2.2124 * Z ^ 2 - 3.2937 * Z + 1.8873)   '(7.5-13) 
        β01 = k_N * (-0.1833 * Z ^ 3 + 1.0383 * Z ^ 2 - 1.2943 * Z + 0.837)     '(7.5-15) 
        β02 = 0.95 * (0.56 - 1.94 * Y - 82.5 * Y ^ 2)     '(7.5-17) 
        If β02 < 0.5 Then β02 = 0.5

        Select Case (X)
            Case < 0.1
                β = 25 * ((0.1 - X) * β006 + (X - 0.06) * β01)
            Case < 0.2
                β = 10 * ((0.2 - X) * β01 + (X - 0.1) * β02)
        End Select
        Return (β)
    End Function

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click, NumericUpDown20.ValueChanged, GroupBox12.Enter, TabPage7.Enter, NumericUpDown21.ValueChanged
        'Design_stress()
        Calc_flat_end()
    End Sub
    Private Sub Calc_flat_end()
        'Flat ends welded  directly to the shell 10.4.4 
        Dim De, Di As Double
        Dim E_flat1 As Double       'End plate thicknes
        Dim E_flat2 As Double       'End plate thicknes
        Dim E_flat As Double        'Required end plate thicknes
        Dim es As Double            'Wall shell
        Dim A1, B1 As Double
        Dim C1a, C1b, C1 As Double  ' Equation (10.4-4) 
        Dim C2 As Double            ' Equation (10.4-6)
        Dim C2_temp1, C2_temp2 As Double
        Dim g, H, J, U, f1 As Double
        Dim A, B, F, G_ As Double
        Dim a_, b_, c_, N, Q, K, S As Double
        Dim ν As Double = 0.303     'Poisson 's ratio mild steel 

        De = NumericUpDown20.Value   'Outside dia shell
        es = NumericUpDown21.Value  'Shell Thickness
        Di = De - 2 * es            'Inside dia shell

        '------------- Calc C1 (10.4-4)---------------------

        B1 = 1                          '(10.4-6)
        B1 -= (3 * _fs / _P) * (es / (Di + es)) ^ 2
        B1 += 3 / 16 * (Di / (Di + es)) ^ 4 * (_P / _fs)
        B1 -= 3 * ((2 * Di + es) * es ^ 2) / (4 * (Di + es) ^ 3)

        A1 = B1 * (1 - B1 * es / (2 * (Di + es)))   '(10.4-5) 


        C1a = 0.40825 * A1 * (Di + es) / Di          '(10.4-4) 
        C1b = 0.299 * (1 + 1.7 * es / Di)

        C1 = IIf(C1a > C1b, C1a, C1b)   'Find biggest

        '-----------Calc C2 method 10.4-5 ---------- 

        g = Di / (Di + es)                  '(10.4-16)

        H = (12 * (1 - ν ^ 2)) ^ 0.25       '(10.4-17) 
        H *= Sqrt(es / (Di + es))

        J = 3 * _fs / _P                     '(10.4-18) 
        J -= Di ^ 2 / (4 * (Di + es) * es)
        J -= 1

        U = 2 * (2 - ν * g)
        U /= Sqrt(3 * (1 - ν ^ 2))              '(10.4-19) 

        f1 = 2 * g ^ 2 - g ^ 4                  '(10.4-20) 

        A = 3 * U * Di / (4 * es)               '(10.4-21) 
        A -= 2 * J
        A *= (1 + ν)
        A *= (1 + (1 - ν) * es / (Di + es))

        B = 3 * U * Di / (8 * es)               '(10.4 - 22)
        B -= J
        B *= H ^ 2
        B -= (3 / 2 * (2 - ν * g) * g)
        B *= H

        F = 3 / 8 * U * g                       '(10.4-23) 
        F += 3 / 16 * f1 * (Di + es) / es
        F -= (2 * J * es / (Di + es))
        F *= H ^ 2
        F -= (3 * (2 - ν * g) * g * es / (Di + es))

        G_ = 3 / 8 * f1                         '(10.4-24) 
        G_ -= 2 * J * (es / (Di + es)) ^ 2
        G_ *= H

        a_ = B / A                              '(10.4-25) 
        b_ = F / A                              '(10.4-26) 
        c_ = G_ / A                             '(10.4-27) 

        N = b_ / 3                              '(10.4-28) 
        N -= (a_ ^ 2 / 9)

        Q = c_ / 2                              '(10.4-29)
        Q -= (a_ * b_ / 6)
        Q += (a_ ^ 3 / 27)

        K = N ^ 3 / Q ^ 2                       '(10.4-30)

        If Q >= 0 Then
            S = (Q * (1 + Sqrt(1 + K))) ^ (1 / 3)
        Else
            S = -(Abs(Q) * (1 + Sqrt(1 + K))) ^ (1 / 3)
        End If

        C2_temp1 = (Di + es) * (N / S - S - a_ / 3)
        C2_temp2 = Di * Sqrt(_P / _fs)
        C2 = C2_temp1 / C2_temp2
        If C2 <= 0.3 Then C2 = 0.3
        '-------------------------------------------------

        E_flat1 = C1 * Di * Sqrt(_P / _fs)
        E_flat2 = C2 * Di * Sqrt(_P / _fs)
        E_flat = IIf(E_flat1 > E_flat2, E_flat1, E_flat2)   'The biggest

        TextBox56.Text = _P.ToString("0.00")
        TextBox55.Text = (_P * 10).ToString("0.0")
        TextBox60.Text = Di.ToString("0")

        TextBox57.Text = C1.ToString("0.000")
        TextBox54.Text = C2.ToString("0.000")

        TextBox65.Text = E_flat1.ToString("0")
        TextBox64.Text = E_flat2.ToString("0")
        TextBox58.Text = E_flat.ToString("0")
        TextBox62.Text = _fs.ToString("0")

        '----Chart determine C1 (10.4-4)
        TextBox59.Text = (es / Di).ToString("0.000")
        TextBox61.Text = (_P / _fs).ToString("0.000")

        '----Chart determine C2 (10.4-5)
        TextBox59.Text = (es / Di).ToString("0.000")
        TextBox61.Text = (_P / _fs).ToString("0.000")

        'Checks
        TextBox57.BackColor = IIf(C1 > 0.29 And C1 < 0.42, Color.LightGreen, Color.Red)
        TextBox54.BackColor = IIf(C2 >= 0.3 And C2 < 1.0, Color.LightGreen, Color.Red)

    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click, TabPage8.Enter, NumericUpDown22.ValueChanged
        '10.5.3 Flat end with a full-face gasket 
        Dim e_flange, dia_bolt As Double

        dia_bolt = NumericUpDown22.Value
        e_flange = 0.41 * dia_bolt * Sqrt(_P / _fs)

        TextBox68.Text = _P.ToString("0.0")
        TextBox67.Text = (_P * 10).ToString("0.0")
        TextBox63.Text = _fs.ToString("0.0")
        TextBox71.Text = e_flange.ToString("0.0")
        TextBox74.Text = (e_flange * 0.8).ToString("0.0")
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click, TabPage9.Enter, NumericUpDown24.ValueChanged, NumericUpDown23.ValueChanged, NumericUpDown31.ValueChanged, NumericUpDown30.ValueChanged, NumericUpDown29.ValueChanged, NumericUpDown28.ValueChanged, NumericUpDown34.ValueChanged, NumericUpDown32.ValueChanged, ComboBox4.SelectedIndexChanged, NumericUpDown27.ValueChanged, NumericUpDown26.ValueChanged, NumericUpDown25.ValueChanged, ComboBox6.SelectedIndexChanged
        Calc_flange_Moments()
    End Sub

    Sub Calc_flange_Moments()
        Dim words() As String
        Dim e, G, gt, HG, H, B, C, A As Double
        Dim db, dn, n As Double
        Dim fB As Double
        Dim W, w_, b_gasket, b0_gasket, m As Double
        Dim y, Wa, Wop As Double
        Dim AB_min1, AB_min2, AB_min, Dia_bolt As Double

        Dim HD, HT As Double
        Dim hD_, hG_, hT_ As Double
        Dim Ma, Mop As Double
        Dim G0 As Double
        Dim βT, βU, βY As Double

        Dim CF, δb, K, I0 As Double
        Dim M1, M2, σθ As Double
        Dim temp As Double

        If (ComboBox4.SelectedIndex > -1) Then          'Prevent exceptions
            words = gaskets(ComboBox4.SelectedIndex).Split(separators, StringSplitOptions.None)
            Double.TryParse(words(1), NumericUpDown30.Value)    'Gasket factor m
            Double.TryParse(words(2), NumericUpDown31.Value)    'Gasket seat pressure y
        End If

        If (ComboBox6.SelectedIndex > -1) Then          'Prevent exceptions
            words = Bolt(ComboBox6.SelectedIndex + 1).Split(separators, StringSplitOptions.None)

            Double.TryParse(words(1), temp)    'Bolt stress
            NumericUpDown27.Value = temp / 3
        End If

        A = NumericUpDown34.Value           'OD Flange 
        n = NumericUpDown25.Value           'Is No Bolts  
        C = NumericUpDown28.Value           'Bolt circle
        w_ = NumericUpDown24.Value          'Gasket width
        B = NumericUpDown23.Value           'ID Flange
        gt = NumericUpDown29.Value          'OD Gasket
        y = NumericUpDown31.Value           'Min seat stress
        m = NumericUpDown30.Value           'Gasket factor
        db = NumericUpDown26.Value          'Dia bolt
        dn = db                             'Dia bolt nominal (niet af)
        fB = NumericUpDown27.Value          'Bolt design stress at oper-temp (Rp02/3)
        e = NumericUpDown32.Value           'Slip on flangethickness

        '------------- bolting ------------
        b0_gasket = w_ / 2                  '(11.5.2)

        '------b_gasket = effective gasket width---------
        If b0_gasket <= 6.3 Then
            b_gasket = b0_gasket
            G = gt
        Else
            b_gasket = 2.52 * Sqrt(b0_gasket)  '(11.5-4)
            G = gt - 2 * b_gasket
        End If

        H = PI / 4 * (G ^ 2 * _P)              '(11.5-5) Hydrostatic end force
        HG = 2 * PI * G * b_gasket * m * _P    '(11.5-6)
        Wa = PI * b0_gasket * G * y            '(11.5-7) Min req. bolt load
        Wop = H + HG                           '(11.5-8)

        AB_min1 = Wa / fB
        AB_min2 = Wop / fB
        AB_min = IIf(AB_min1 < AB_min2, AB_min2, AB_min1)   'Take biggest

        '---- dia bolt---
        Dia_bolt = Sqrt((AB_min / n) * 4 / PI)

        '------------- Stepped Flange moment (11.5.3)------------

        HD = PI / 4 * B * 2 * _P    'Hydrostatic force via shell
        HT = H - HD                 'Hydrostatic force via flange face

        hD_ = (C - B) / 2                   '(11.5-13)
        hG_ = (C - G) / 2                   '(11.5-14)
        hT_ = (2 * C - B - G) / 4           '(11.5-15)

        W = 0.5 * (db + dn) * fB            '(11.5-16)

        Ma = W * hG_                        '(11.5-17)
        Mop = HD * hD_ + HT * hT_ + HG * hG_ '(11.5-18)

        '------------- Flange stresses and stress limit (11.5.4)------------
        δb = C * PI / n               '[mm] Distance adjacent bolts
        CF = Sqrt(δb / (2 * db + 6 * e / (m + 0.5)))
        K = A / B                               '(11.5-21)
        G0 = e
        I0 = Sqrt(B * G0)                       '(11.5-22)

        βT = (K ^ 2 * (1 + 8.55246 * Log10(K))) - 1
        βT /= (1.0472 + 1.9448 * K ^ 2) * (K - 1)

        βU = (K ^ 2 * (1 + 8.55246 * Log10(K))) - 1
        βU /= 1.36136 * (K ^ 2 - 1) * (K - 1)

        βY = 1 / (K - 1)
        βY *= 0.66845 + 5.7169 * (K ^ 2 * Log10(K)) / (K ^ 2 - 1)

        M1 = Ma * CF / B                '(11.5-26)
        M2 = Mop * CF / B               '(11.5-27)

        '----------- Loose flange method ----
        σθ = βY * M2 / e ^ 2            '(11.5-35)

        TextBox77.Text = (_P * 10).ToString("0.0")
        TextBox78.Text = _P.ToString("0.0")
        TextBox76.Text = _fs.ToString("0")
        TextBox79.Text = H.ToString("0.0")
        TextBox81.Text = b_gasket.ToString("0.0")
        TextBox93.Text = G.ToString("0.0")
        TextBox84.Text = dn.ToString("0.0")             '[mm]
        TextBox94.Text = AB_min.ToString("0")         '[mm2] required bolt area
        TextBox95.Text = Dia_bolt.ToString("0.0")       '[mm] calculated req. dia bolt

        TextBox85.Text = (H / 1000).ToString("0.0")      '[kN]
        TextBox87.Text = (HG / 1000).ToString("0.0")     '[kN]
        TextBox82.Text = (Wa / 1000).ToString("0.0")     '[kN]
        TextBox83.Text = (Wop / 1000).ToString("0.0")    '[kN]

        TextBox79.Text = (HD / 1000).ToString("0.0")     '[kN]
        TextBox75.Text = (HT / 1000).ToString("0.0")     '[kN]

        TextBox80.Text = hD_.ToString("0.0")        '[mm]
        TextBox86.Text = hG_.ToString("0.0")        '[mm]
        TextBox88.Text = hT_.ToString("0.0")        '[mm]
        TextBox89.Text = W.ToString("0.0")          '[mm]

        TextBox91.Text = (Ma * 10 ^ -3).ToString("0")   '[mm]
        TextBox90.Text = (Mop * 10 ^ -3).ToString("0")  '[mm]
        TextBox92.Text = CF.ToString("0.0")         '[-]
        TextBox98.Text = βT.ToString("0.00")        '[-]
        TextBox97.Text = βU.ToString("0.00")        '[-]
        TextBox96.Text = βY.ToString("0.00")        '[-]
        TextBox99.Text = K.ToString("0.00")         '[-]

        TextBox100.Text = (M1 * 10 ^ -3).ToString("0.00")    '[-] Moment assembly
        TextBox101.Text = (M2 * 10 ^ -3).ToString("0.00")    '[-] Moment operating
        TextBox102.Text = σθ.ToString("0")       '[N/mm2]Tangential flange stress
        TextBox103.Text = δb.ToString("0")       '[mm] Distance bolts

        '-------------- checks --------------------
        NumericUpDown28.BackColor = IIf(C <= B, Color.Red, Color.Yellow)    'Bolt diameter
        NumericUpDown34.BackColor = IIf(A <= C, Color.Red, Color.Yellow)    'Flange OD
        NumericUpDown24.BackColor = IIf(w_ > (A - B) / 2, Color.Red, Color.Yellow)    'Gasket width
        NumericUpDown26.BackColor = IIf(Dia_bolt > db, Color.Red, Color.Yellow)    'Bolt dia
        TextBox102.BackColor = IIf(σθ > _fs, Color.Red, Color.Yellow)    'Flange stress
    End Sub

    Private Sub PictureBox8_Click(sender As Object, e As EventArgs) Handles PictureBox8.Click
        Form2.Show()
    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        Dim oWord As Word.Application ' = Nothing

        Dim oDoc As Word.Document
        Dim oTable As Word.Table
        Dim oPara1, oPara2 As Word.Paragraph
        Dim row, font_sizze As Integer
        Dim ufilename As String
        ufilename = "PV_calc_" & TextBox7.Text & "_" & TextBox8.Text & "_" & DateTime.Now.ToString("yyyy_MM_dd") & ".docx"

        Try
            oWord = New Word.Application()

            'Start Word and open the document template. 
            font_sizze = 9
            oWord = CType(CreateObject("Word.Application"), Word.Application)
            oWord.Visible = True
            oDoc = oWord.Documents.Add

            oDoc.PageSetup.TopMargin = 35
            oDoc.PageSetup.BottomMargin = 15

            'Insert a paragraph at the beginning of the document. 
            oPara1 = oDoc.Content.Paragraphs.Add

            oPara1.Range.Text = "VTK Engineering"
            oPara1.Range.Font.Name = "Arial"
            oPara1.Range.Font.Size = font_sizze + 3
            oPara1.Range.Font.Bold = CInt(True)
            oPara1.Format.SpaceAfter = 1                '24 pt spacing after paragraph. 
            oPara1.Range.InsertParagraphAfter()

            oPara2 = oDoc.Content.Paragraphs.Add(oDoc.Bookmarks.Item("\endofdoc").Range)
            oPara2.Range.Font.Size = font_sizze + 1
            oPara2.Format.SpaceAfter = 1
            oPara2.Range.Font.Bold = CInt(False)
            oPara2.Range.Text = "Pressure Vessel calculation acc. EN13445" & vbCrLf
            oPara2.Range.InsertParagraphAfter()

            '----------------------------------------------
            'Insert a table, fill it with data and change the column widths.
            oTable = oDoc.Tables.Add(oDoc.Bookmarks.Item("\endofdoc").Range, 6, 2)
            oTable.Range.ParagraphFormat.SpaceAfter = 1
            oTable.Range.Font.Size = font_sizze
            oTable.Range.Font.Bold = CInt(False)
            oTable.Rows.Item(1).Range.Font.Bold = CInt(True)

            row = 1
            oTable.Cell(row, 1).Range.Text = "Project Name"
            oTable.Cell(row, 2).Range.Text = TextBox69.Text
            row += 1
            oTable.Cell(row, 1).Range.Text = "Project number "
            oTable.Cell(row, 2).Range.Text = TextBox66.Text
            row += 1
            oTable.Cell(row, 1).Range.Text = "Pressure vessel id"
            oTable.Cell(row, 2).Range.Text = TextBox70.Text
            row += 1
            oTable.Cell(row, 1).Range.Text = "Author"
            oTable.Cell(row, 2).Range.Text = Environment.UserName
            row += 1
            oTable.Cell(row, 1).Range.Text = "Date"
            oTable.Cell(row, 2).Range.Text = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")
            row += 1
            oTable.Cell(row, 1).Range.Text = "File name"
            oTable.Cell(row, 2).Range.Text = ufilename

            oTable.Columns(1).Width = oWord.InchesToPoints(2.5)   'Change width of columns 1 & 2.
            oTable.Columns(2).Width = oWord.InchesToPoints(4)

            oTable.Rows.Item(1).Range.Font.Bold = CInt(True)
            oDoc.Bookmarks.Item("\endofdoc").Range.InsertParagraphAfter()

            '----------------------------------------------
            'Insert a 18 (row) x 3 table (column), fill it with data and change the column widths.
            oTable = oDoc.Tables.Add(oDoc.Bookmarks.Item("\endofdoc").Range, 12, 3)
            oTable.Range.ParagraphFormat.SpaceAfter = 1
            oTable.Range.Font.Size = font_sizze
            oTable.Range.Font.Bold = CInt(False)
            oTable.Rows.Item(1).Range.Font.Bold = CInt(True)
            oTable.Rows.Item(1).Range.Font.Size = font_sizze + 2
            row = 1
            oTable.Cell(row, 1).Range.Text = "Flange Input Data"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Internal Pressure"
            oTable.Cell(row, 2).Range.Text = TextBox77.Text
            oTable.Cell(row, 3).Range.Text = "[bar]"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Flange design stress"
            oTable.Cell(row, 2).Range.Text = TextBox76.Text
            oTable.Cell(row, 3).Range.Text = "[N/mm2]"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Flange OD"
            oTable.Cell(row, 2).Range.Text = NumericUpDown34.Value.ToString("0")
            oTable.Cell(row, 3).Range.Text = "[mm]"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Flange ID"
            oTable.Cell(row, 2).Range.Text = NumericUpDown23.Value.ToString("0")
            oTable.Cell(row, 3).Range.Text = "[mm]"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Bolt circle diameter"
            oTable.Cell(row, 2).Range.Text = NumericUpDown28.Value.ToString("0")
            oTable.Cell(row, 3).Range.Text = "[mm]"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Slip on flange thicknes"
            oTable.Cell(row, 2).Range.Text = NumericUpDown32.Value.ToString("0")
            oTable.Cell(row, 3).Range.Text = "[mm]"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Flange face Gasket width"
            oTable.Cell(row, 2).Range.Text = NumericUpDown24.Value.ToString("0")
            oTable.Cell(row, 3).Range.Text = "[mm]"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Gasket OD"
            oTable.Cell(row, 2).Range.Text = NumericUpDown29.Value.ToString("0")
            oTable.Cell(row, 3).Range.Text = "[mm]"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Gasket material"
            oTable.Cell(row, 2).Range.Text = ComboBox4.Text.Substring(0, 24)
            oTable.Cell(row, 3).Range.Text = ""
            row += 1
            oTable.Cell(row, 1).Range.Text = "Gasket factor m"
            oTable.Cell(row, 2).Range.Text = NumericUpDown30.Value.ToString("0.0")
            oTable.Cell(row, 3).Range.Text = "[-]"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Gasket factor y"
            oTable.Cell(row, 2).Range.Text = NumericUpDown31.Value.ToString("0.0")
            oTable.Cell(row, 3).Range.Text = "[-]"

            oTable.Columns(1).Width = oWord.InchesToPoints(2.5)   'Change width of columns 1 & 2.
            oTable.Columns(2).Width = oWord.InchesToPoints(1.5)
            oTable.Columns(3).Width = oWord.InchesToPoints(1.3)
            oDoc.Bookmarks.Item("\endofdoc").Range.InsertParagraphAfter()

            '---------"Bolting 11.4.1"---------------------------------------
            oTable = oDoc.Tables.Add(oDoc.Bookmarks.Item("\endofdoc").Range, 5, 3)
            oTable.Range.ParagraphFormat.SpaceAfter = 1
            oTable.Range.Font.Size = font_sizze
            oTable.Range.Font.Bold = CInt(False)
            oTable.Rows.Item(1).Range.Font.Bold = CInt(True)
            oTable.Rows.Item(1).Range.Font.Size = font_sizze + 2
            row = 1
            oTable.Cell(row, 1).Range.Text = "Bolting 11.4.1"

            row += 1
            oTable.Cell(row, 1).Range.Text = "Number of Bolts"
            oTable.Cell(row, 2).Range.Text = NumericUpDown25.Value.ToString("0")
            oTable.Cell(row, 3).Range.Text = "[-]"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Bolt diameter"
            oTable.Cell(row, 2).Range.Text = NumericUpDown26.Value.ToString("0")
            oTable.Cell(row, 3).Range.Text = "[mm]"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Bolting class"
            oTable.Cell(row, 2).Range.Text = ComboBox6.Text
            oTable.Cell(row, 3).Range.Text = "[-]"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Bolt design stress"
            oTable.Cell(row, 2).Range.Text = NumericUpDown27.Value.ToString("0")
            oTable.Cell(row, 3).Range.Text = "[N/mm2]"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Bolt required dia"
            oTable.Cell(row, 2).Range.Text = TextBox95.Text
            oTable.Cell(row, 3).Range.Text = "[mm]"


            oTable.Columns(1).Width = oWord.InchesToPoints(2.5)   'Change width of columns 1 & 2.
            oTable.Columns(2).Width = oWord.InchesToPoints(1.5)
            oTable.Columns(3).Width = oWord.InchesToPoints(1.3)
            oDoc.Bookmarks.Item("\endofdoc").Range.InsertParagraphAfter()

            '---------"Bolt Loads and area's 11.5.2"---------------------------------------
            oTable = oDoc.Tables.Add(oDoc.Bookmarks.Item("\endofdoc").Range, 7, 3)
            oTable.Range.ParagraphFormat.SpaceAfter = 1
            oTable.Range.Font.Size = font_sizze
            oTable.Range.Font.Bold = CInt(False)
            oTable.Rows.Item(1).Range.Font.Bold = CInt(True)
            oTable.Rows.Item(1).Range.Font.Size = font_sizze + 2
            row = 1
            oTable.Cell(row, 1).Range.Text = "Bolt Loads and area's 11.5.2"

            row += 1
            oTable.Cell(row, 1).Range.Text = "Effective Gasket width"
            oTable.Cell(row, 2).Range.Text = TextBox81.Text
            oTable.Cell(row, 3).Range.Text = "[N]"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Diameter Gasket load reaction"
            oTable.Cell(row, 2).Range.Text = TextBox93.Text
            oTable.Cell(row, 3).Range.Text = "[N]"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Total Hydrostatic force"
            oTable.Cell(row, 2).Range.Text = TextBox85.Text
            oTable.Cell(row, 3).Range.Text = "[mm]"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Compression load on Gasket"
            oTable.Cell(row, 2).Range.Text = TextBox87.Text
            oTable.Cell(row, 3).Range.Text = "[kN]"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Min bolt load for Assembly"
            oTable.Cell(row, 2).Range.Text = TextBox82.Text
            oTable.Cell(row, 3).Range.Text = "[kN]"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Min bolt load for Operating"
            oTable.Cell(row, 2).Range.Text = TextBox83.Text
            oTable.Cell(row, 3).Range.Text = "[kN]"

            oTable.Columns(1).Width = oWord.InchesToPoints(2.5)   'Change width of columns 1 & 2.
            oTable.Columns(2).Width = oWord.InchesToPoints(1.5)
            oTable.Columns(3).Width = oWord.InchesToPoints(1.3)
            oDoc.Bookmarks.Item("\endofdoc").Range.InsertParagraphAfter()

            '---------"Flange moments 11.5.3"---------------------------------------
            oTable = oDoc.Tables.Add(oDoc.Bookmarks.Item("\endofdoc").Range, 9, 3)
            oTable.Range.ParagraphFormat.SpaceAfter = 1
            oTable.Range.Font.Size = font_sizze
            oTable.Range.Font.Bold = CInt(False)
            oTable.Rows.Item(1).Range.Font.Bold = CInt(True)
            oTable.Rows.Item(1).Range.Font.Size = font_sizze + 2
            row = 1
            oTable.Cell(row, 1).Range.Text = "Flange moments 11.5.3"

            row += 1
            oTable.Cell(row, 1).Range.Text = "Hydrostic force via shell"
            oTable.Cell(row, 2).Range.Text = TextBox79.Text
            oTable.Cell(row, 3).Range.Text = "[N]"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Hydrostic force flange face"
            oTable.Cell(row, 2).Range.Text = TextBox75.Text
            oTable.Cell(row, 3).Range.Text = "[N]"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Radial distance bolt circle"
            oTable.Cell(row, 2).Range.Text = TextBox80.Text
            oTable.Cell(row, 3).Range.Text = "[mm]"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Rad dist Bolt circle-HD"
            oTable.Cell(row, 2).Range.Text = TextBox86.Text
            oTable.Cell(row, 3).Range.Text = "[mm]"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Dist gasket load to bolt circle"
            oTable.Cell(row, 2).Range.Text = TextBox88.Text
            oTable.Cell(row, 3).Range.Text = "[-]"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Design load assembly"
            oTable.Cell(row, 2).Range.Text = TextBox89.Text
            oTable.Cell(row, 3).Range.Text = "[N]"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Total moment assembly condition"
            oTable.Cell(row, 2).Range.Text = TextBox91.Text
            oTable.Cell(row, 3).Range.Text = "[Nm]"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Total moment operating condition"
            oTable.Cell(row, 2).Range.Text = TextBox90.Text
            oTable.Cell(row, 3).Range.Text = "[Nm]"

            oTable.Columns(1).Width = oWord.InchesToPoints(2.5)   'Change width of columns 1 & 2.
            oTable.Columns(2).Width = oWord.InchesToPoints(1.5)
            oTable.Columns(3).Width = oWord.InchesToPoints(1.3)
            oDoc.Bookmarks.Item("\endofdoc").Range.InsertParagraphAfter()


            '---------"Flange stress 11.5.4"---------------------------------------
            oTable = oDoc.Tables.Add(oDoc.Bookmarks.Item("\endofdoc").Range, 6, 3)
            oTable.Range.ParagraphFormat.SpaceAfter = 1
            oTable.Range.Font.Size = font_sizze
            oTable.Range.Font.Bold = CInt(False)
            oTable.Rows.Item(1).Range.Font.Bold = CInt(True)
            oTable.Rows.Item(1).Range.Font.Size = font_sizze + 2
            row = 1
            oTable.Cell(row, 1).Range.Text = "Flange stress 11.5.4"

            row += 1
            oTable.Cell(row, 1).Range.Text = "Distance bolts"
            oTable.Cell(row, 2).Range.Text = TextBox103.Text
            oTable.Cell(row, 3).Range.Text = "[-]"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Bolt pich factor"
            oTable.Cell(row, 2).Range.Text = TextBox92.Text
            oTable.Cell(row, 3).Range.Text = "[-]"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Moment assembly"
            oTable.Cell(row, 2).Range.Text = TextBox100.Text
            oTable.Cell(row, 3).Range.Text = "[Nm]"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Moment operating"
            oTable.Cell(row, 2).Range.Text = TextBox101.Text
            oTable.Cell(row, 3).Range.Text = "[Nm]"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Tangential Flange Stress"
            oTable.Cell(row, 2).Range.Text = TextBox102.Text
            oTable.Cell(row, 3).Range.Text = "[N/mm2]"

            oTable.Columns(1).Width = oWord.InchesToPoints(2.5)   'Change width of columns 1 & 2.
            oTable.Columns(2).Width = oWord.InchesToPoints(1.5)
            oTable.Columns(3).Width = oWord.InchesToPoints(1.3)
            oDoc.Bookmarks.Item("\endofdoc").Range.InsertParagraphAfter()

            '------------------save picture ---------------- 
            'Chart1.SaveImage("c:\Temp\MainChart.gif", System.Drawing.Imaging.ImageFormat.Gif)
            'oPara4 = oDoc.Content.Paragraphs.Add
            'oPara4.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft
            'oPara4.Range.InlineShapes.AddPicture("c:\Temp\MainChart.gif")
            'oPara4.Range.InlineShapes.Item(1).LockAspectRatio = CType(True, Microsoft.Office.Core.MsoTriState)
            'oPara4.Range.InlineShapes.Item(1).Width = 310
            'oDoc.Bookmarks.Item("\endofdoc").Range.InsertParagraphAfter()


            '--------------Save file word file------------------
            'See https://msdn.microsoft.com/en-us/library/63w57f4b.aspx

            If (Not System.IO.Directory.Exists(dirpath_Eng)) Then System.IO.Directory.CreateDirectory(dirpath_Eng)
            If (Not System.IO.Directory.Exists(dirpath_Rap)) Then System.IO.Directory.CreateDirectory(dirpath_Rap)
            If (Not System.IO.Directory.Exists(dirpath_Home)) Then System.IO.Directory.CreateDirectory(dirpath_Home)

            If Directory.Exists(dirpath_Rap) Then
                oWord.ActiveDocument.SaveAs(dirpath_Rap & ufilename)
            Else
                oWord.ActiveDocument.SaveAs(dirpath_Home & ufilename)
            End If

        Catch ex As Exception
            MessageBox.Show(ufilename & vbCrLf & ex.Message)  ' Show the exception's message.
        End Try
    End Sub
End Class
