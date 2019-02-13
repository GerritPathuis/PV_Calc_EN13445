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
    Public _ν As Double = 0.3    'Poisson ratio for steel
    Public _P As Double         'Calculation pressure [Mpa]
    Public _fs As Double        'Allowable stress shell [N/mm2]
    Public _f02 As Double       'Yield 0.2% stress shell [N/mm2]
    Public _fym As Double       'Allowable stress reinforcement [N/mm2]

    Public _De As Double        'Outside diameter shell
    Public _Di As Double        'Inside diameter shell
    Public _ecs As Double       'Shell thickness

    Public _deb As Double       'Outside diameter nozzle fitted in shell
    Public _dib As Double       'Inside diameter nozzle fitted in shell
    Public _eb As Double        'Effective thickness nozzle thickness
    Public _E As Double         'Modulus of elasticity 
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
   "Material-------;50c;100;150;200;250;300;350;400;450;500;550;remarks--;cs/ss",
   "1.0425 (P265GH);265;241;223;205;188;173;160;150;  0;  0;  0; max 400c;cs",
   "1.0473 (P355GH);343;323;299;275;252;232;214;202;  0;  0;  0; max 400c;cs",
   "1.4301 (304)   ;190;157;142;127;118;110;104; 98; 95; 92; 90; max 550c;ss",
   "1.4307 (304L)  ;180;147;132;118;108;100; 94; 89; 85; 81; 80; max 550c;ss",
   "1.4401 (316)   ;204;177;162;147;137;127;120;115;112;110;108; max 550c;ss",
   "1.4404 (316L)  ;200;166;152;137;127;118;113;108;103;100; 98; max 550c;ss"}

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
        "Part 3 Design (issue 2016)"

        TextBox22.Text =
        "Important note" & vbCrLf &
        "The yield strength follows EN 10028-2:2009 (mild steel)" & vbCrLf &
        "and EN 10028-7:2016 for stainless steel at given temperatureerature." & vbCrLf &
        "Safety factors follow the Eurocode" & vbCrLf & vbCrLf &
        "EN 14460:2006, Explosion resistand design follow EN 13445" & vbCrLf &
        "for Explosion-Pressure-Shock-Resistant design stress multiplied bu 1.5"

        TextBox66.Text = "P" & DateTime.Now.ToString("yy") & ".10"

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

        TextBox66.Text = "P" & Now.ToString("yy") & ".10"
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click, NumericUpDown14.ValueChanged, TabPage4.Enter, NumericUpDown12.ValueChanged, NumericUpDown1.ValueChanged, NumericUpDown41.ValueChanged, NumericUpDown18.ValueChanged
        Calc_nozzle_fig949()
    End Sub
    Private Sub Calc_nozzle_fig949()
        Dim nozzle_wall As Double
        Dim fob, fop As Double
        Dim Afs, Afw As Double
        Dim Ls As Double
        Dim Lso, eas, ris As Double   'Max length shell contibuting to reinforcement
        Dim Lbo, eab, rib As Double   'Max length nozzle contibuting to reinforcement

        Dim Aps As Double       'Figure 9.4-7 — Cylindrical shell with isolated opening and set-on nozzle 
        Dim Apb As Double       'Figure 9.4-7  
        Dim Afb As Double       'Figure 9.4-7  
        Dim Afp As Double       'Figure 9.4-7  
        Dim Ap_phi As Double
        Dim eq_left, eq_right, eq_ratio As Double
        Dim Ln, Ln1, Ln2 As Double
        Dim D_small_opening As Double
        Dim W_min, W_min1, W_min2 As Double
        Dim qz As Double

        Ls = NumericUpDown1.Value           'Actual Distance shell-edge-opening to discontinuity
        _De = NumericUpDown18.Value         'Shell OD 
        _eb = NumericUpDown41.Value         'Shell Wall 
        _Di = _De - 2 * _eb                 'Shell ID 
        _deb = NumericUpDown14.Value        'Outside diameter nozzle fitted in shell
        eas = _eb                           'Shell Analysis thickness of shell wall 

        If _deb >= _De Then       'Nozzle dia can not be bigger then shell diameter
            _deb = _De
            NumericUpDown14.Value = CDec(_deb)
        End If

        nozzle_wall = NumericUpDown12.Value
        _dib = _deb - 2 * nozzle_wall
        If _dib < 10 Then _dib = 10
        TextBox144.Text = _dib.ToString("0.0")

        '--------- Small opening 9.5.2.2
        D_small_opening = 0.15 * Sqrt((_Di + _eb) * _eb)        '(9.5-18) 
        Label77.Text = "D= " & D_small_opening.ToString("0.0") & " [mm]"

        '------- reinforment materials is identical to shell material----
        fob = _fs
        fop = _fs

        '---------formula 9.5-2 Nozzle-----
        eab = nozzle_wall               '(9.5-2) 
        rib = _dib / 2
        Lbo = Sqrt((2 * rib + eab) + eab)

        '----Chapter 9.5.2.4.4.3 Nozzle in cylindrical shell
        Dim a, ls_min, ecs As Double
        ecs = eas                                   'assumed shell thickness for calculation
        a = _deb / 2                                'equation (9.5-90) (page 107)
        ris = (_De / 2) - eas                       'equation (9.5-91)
        Lso = Sqrt(((_De - 2 * eas) + ecs) * ecs)   'equation (9.5-92)
        ls_min = CDbl(IIf(Lso < Ls, Lso, Ls))       'equation (9.5-93)
        Aps = ris * (ls_min + a)                    'equation (9.5-94)

        '--------------- formula (9.5-7) (page 99) --------------------
        'Af = Stress loaded cross-sectional area effective as reinforcement
        Afw = 0                              'Weld area in neglected
        Afb = nozzle_wall * (Lbo + _eb)      'Nozzle wall
        Afp = 0                              'reinforcement ring NOT present
        Afs = Lso * _eb                      'Shell wall area

        'Ap = Pressure loaded area. 
        Apb = _dib / 2 * (Lbo + _eb)         'Nozzle Pressure loaded area
        Ap_phi = 0                           'Oblique nozzles

        eq_left = (Afs + Afw) * (_fs - 0.5 * _P)    'left side(9.5-7)
        eq_left += Afp * (fop - 0.5 * _P)
        eq_left += Afb * (fob - 0.5 * _P)

        eq_right = _P * (Aps + Apb + 0.5 * Ap_phi)  'right side(9.5-7)
        eq_ratio = eq_left / eq_right

        '----- 9.4.8 Minimum Distance between nozle and shell butt-weld
        Ln1 = 0.5 * _deb + 2 * nozzle_wall      'Equation 9.4-4 (page 90)
        Ln2 = 0.5 * _deb + 40
        Ln = CDbl(IIf(Ln1 > Ln2, Ln1, Ln2))     'Find biggest


        qz = Ln - _deb / 2      'Distance shell edge opening-weld

        '-------- Figure 9.7-5          
        W_min1 = 0.2 * Sqrt((2 * _Di * 0.5 + ecs) * ecs)    'equation 9.7-5
        W_min2 = 3 * eas                                    'equation 9.7-5
        W_min = CDbl(IIf(W_min1 > W_min2, W_min1, W_min2))  'Find biggest

        '----- present--------
        TextBox164.Text = Ln1.ToString("0")     'Distance [mm2]
        TextBox168.Text = Ln2.ToString("0")     'Distance [mm2]

        TextBox163.Text = _Di.ToString("0")     'Shell Inside Diameter
        TextBox9.Text = Afs.ToString("0")       'Shell area reinforcement [mm2]
        TextBox10.Text = Afw.ToString("0")      'Weld area reinforcement [mm2]
        TextBox11.Text = Afb.ToString("0")      'reinforcement [mm2]
        TextBox12.Text = Aps.ToString("0")      'Pressure loaded area [mm2]
        TextBox13.Text = Apb.ToString("0")      'Pressure loaded area [mm2]
        TextBox14.Text = Lso.ToString("0.0")
        TextBox15.Text = lbo.ToString("0.0")

        TextBox16.Text = eq_left.ToString("0")
        TextBox17.Text = eq_right.ToString("0")
        TextBox19.Text = W_min.ToString("0")
        TextBox20.Text = Ln.ToString("0")
        TextBox169.Text = qz.ToString("0")

        '----------- checks--------
        TextBox16.BackColor = CType(IIf(eq_left < eq_right, Color.Red, Color.LightGreen), Color)
        TextBox17.BackColor = TextBox16.BackColor
        TextBox169.BackColor = CType(IIf(qz > ls, Color.Red, Color.LightGreen), Color)
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click, ComboBox1.TextChanged, NumericUpDown5.ValueChanged, NumericUpDown4.ValueChanged, ComboBox5.SelectedIndexChanged, RadioButton3.CheckedChanged, RadioButton2.CheckedChanged, RadioButton1.CheckedChanged, RadioButton4.CheckedChanged
        Design_stress()
    End Sub
    Private Sub Design_stress()
        Dim sf As Double
        Dim temperature As Double  'temperature
        Dim words() As String
        Dim y50, y100, y150, y200, y250, y300, y350, y400 As Double
        Dim yield_stress As Double = 99
        Dim design_str As Double = 99
        'Dim Δy As Double
        Dim ΔT As Double


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
            TextBox182.Text = words(13) 'cs or ss
            Double.TryParse(words(1), y50)
            Double.TryParse(words(2), y100)
            Double.TryParse(words(3), y150)
            Double.TryParse(words(4), y200)
            Double.TryParse(words(5), y250)
            Double.TryParse(words(6), y300)
            Double.TryParse(words(7), y350)
            Double.TryParse(words(8), y400)

            temperature = CDbl(NumericUpDown5.Value)
            Select Case True

                Case 50 >= temperature
                    yield_stress = CDec(y50)
                Case 100 >= temperature
                    ΔT = 50 - temperature
                    yield_stress = Calc_design_stress(y50, y100, ΔT)
                Case 150 >= temperature
                    ΔT = 100 - temperature
                    yield_stress = Calc_design_stress(y100, y150, ΔT)
                Case 200 >= temperature
                    ΔT = 150 - temperature
                    yield_stress = Calc_design_stress(y150, y200, ΔT)
                Case 250 >= temperature
                    ΔT = 200 - temperature
                    yield_stress = Calc_design_stress(y200, y250, ΔT)
                Case 300 >= temperature
                    ΔT = 250 - temperature
                    yield_stress = Calc_design_stress(y250, y300, ΔT)
                Case 350 >= temperature
                    ΔT = 300 - temperature
                    yield_stress = Calc_design_stress(y300, y350, ΔT)
                Case 400 >= temperature
                    ΔT = 350 - temperature
                    yield_stress = Calc_design_stress(y350, y400, ΔT)
                Case temperature > 450
                    MessageBox.Show("Problem temperature too high")
            End Select
        End If

        _P = NumericUpDown4.Value                       'Calculation pressure [MPa=N/mm2]

        If (ComboBox1.SelectedIndex > -1) Then          'Prevent exceptions
            words = chap6(ComboBox1.SelectedIndex).Split(separators, StringSplitOptions.None)
            Double.TryParse(words(1), sf)               'Safety factor
            TextBox4.Text = sf.ToString                 'Safety factor
            design_str = CDec(yield_stress / sf)

            Select Case True
                Case RadioButton4.Checked
                    design_str = yield_stress / 1.25  'PED article 3.3 (NO calc required)
                Case RadioButton1.Checked
                    design_str = design_str       'PED I,II,III
                Case RadioButton2.Checked
                    design_str = design_str * 0.9 'PED IV
                Case RadioButton3.Checked
                    design_str = yield_stress   'EN 14460 6.2.1 (Shock resistant)
            End Select

            _fs = design_str                    'allowable stress
            '_fym = yield_stress * 1.5               'sigma02*1.5
            '_f02 = yield_stress                     'Yield 0.2%

            If String.Equals(TextBox182.Text, "cs") Then
                _E = (213.16 - 6.92 * temperature / 10 ^ 2 - 1.824 / 10 ^ 5 * temperature ^ 2) * 1000
            Else
                _E = (201.66 - 8.48 * temperature / 10 ^ 2) * 1000
            End If

            '-------- present -------------
            TextBox131.Text = (_P * 10 ^ 4).ToString    'Calculation pressure [mBar]
            'TextBox133.Text = _fym.ToString             'Safety factor
            TextBox136.Text = _f02.ToString("0")        'Max allowed bend
            TextBox137.Text = _fym.ToString("0")        'Max allowed bend+membrane
            TextBox140.Text = _fym.ToString("0")        'Max allowed bend+membrane

            NumericUpDown7.Value = CDec(design_str)     'Design stress
            TextBox133.Text = yield_stress.ToString("0")  'Yield stress
            TextBox178.Text = _E.ToString("0")          'Max allowed bend+membrane
            TextBox209.Text = _ν.ToString("0.0")        'Poissons rate for steel
        End If
    End Sub
    Public Function Calc_design_stress(stress_A As Double, stress_B As Double, ΔT As Double) As Double
        Dim Δy As Double

        Δy = stress_B - stress_A
        Return (stress_A - (ΔT / 50 * Δy))
    End Function
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click, NumericUpDown15.ValueChanged, TabPage3.Enter, ComboBox2.SelectedIndexChanged, NumericUpDown16.ValueChanged, NumericUpDown13.ValueChanged, NumericUpDown42.ValueChanged
        Calc_cyl_shell742()         'Cylindrical shell
        Calc_conical_shell764()     'Conus shell
        Calc_Junction766()          'Junction large end
    End Sub
    '7.4.2 Cylindrical shells internal pressure
    Private Sub Calc_cyl_shell742()
        Dim De, Di, Dm, ea, z_joint, e_wall, Pmax, valid_check As Double

        If (ComboBox2.SelectedIndex > -1) Then          'Prevent exceptions
            Double.TryParse(joint_eff(ComboBox2.SelectedIndex), z_joint)      'Joint efficiency
        End If

        De = NumericUpDown15.Value  'OD
        ea = NumericUpDown16.Value  'Wall thicknes
        Di = De - 2 * ea            'ID
        TextBox160.Text = Di.ToString("0.0")

        Dm = (De + Di) / 2                          'Average Dia
        Pmax = 2 * _fs * z_joint * ea / Dm          'Max pressure equation 7.4.3 

        e_wall = _P * De / (2 * _fs * z_joint + _P) 'equation 7.4.2 Required wall thickness
        valid_check = Round(e_wall / De, 4)

        '--------- present results--------
        TextBox2.Text = Round(e_wall, 4).ToString   'required wall [mm]
        TextBox5.Text = valid_check.ToString
        TextBox6.Text = _P.ToString("0.00")         '[MPa]
        TextBox53.Text = (_P * 10).ToString("0.00") '[Bar]
        TextBox7.Text = _fs.ToString
        TextBox8.Text = (Pmax * 10).ToString("0.00") '[Bar]

        '---------- Check-----
        TextBox5.BackColor = CType(IIf(valid_check > 0.16, Color.Red, Color.LightGreen), Color)
        TextBox8.BackColor = CType(IIf(Pmax < _P, Color.Red, Color.LightGreen), Color)
    End Sub
    'Chapter 15.5 rectangle shell
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click, NumericUpDown9.ValueChanged, NumericUpDown8.ValueChanged, NumericUpDown17.ValueChanged, NumericUpDown11.ValueChanged, TabPage2.Enter
        Calc_square_15_5()

    End Sub
    Private Sub Calc_square_15_5()
        Dim a, ee, L, L1, I1 As Double
        Dim σmD, σmC, σmB, σmA, σmBC As Double  'membrane stress
        Dim σbD, σbC, σbB, σbA, σbBC As Double  'bending stress
        Dim σTD, σTC, σTB, σTA, σTBC As Double  'Total stress
        Dim α3, φ, θ, K3, Ma As Double

        _P = NumericUpDown4.Value   'Pressure
        a = NumericUpDown17.Value   'Inside corner radius
        ee = NumericUpDown11.Value  'wall thickness
        L = NumericUpDown8.Value    'Lenght
        L1 = NumericUpDown9.Value   'Lenght


        '-------- membrane stress----------------
        σmC = _P * (a + L) / ee                     'At C (eq 15.5.1.2-1) 
        σmD = σmC                                   'At D
        σmB = _P * (a + L1) / ee                    'At B (eq 15.5.1.2-2) 
        σmA = σmB                                   'At A
        σmBC = (_P / ee) * (a + Sqrt(L ^ 2 + L1 ^ 2)) 'At corner (eq 15.5.1.2-3) 

        '------- bending stress----------------
        I1 = ee ^ 3 / 12    '(eq 15.5.1.2-4) second moment of area
        α3 = L / L1         '(eq 15.5.1.2-14) factor
        φ = a / L1          '(eq 15.5.1.2-15) angular indication

        K3 = 6 * φ ^ 2 * α3                 '(eq 15.5.1.2-12)
        K3 -= 3 * PI * φ ^ 2
        K3 += 6 * φ ^ 2
        K3 += α3 ^ 3
        K3 += 3 * α3 ^ 2
        K3 -= 6 * φ
        K3 -= 2
        K3 += 1.5 * PI * α3 ^ 2 * φ
        K3 += 6 * φ * α3
        K3 *= L1 ^ 2
        K3 /= 3 * (2 * α3 + PI * φ + 2)     'factor unreinforced vessel

        Ma = _P * K3                   'bending moment middle of side


        '------- bend in the corner -----
        θ = Atan(L1 / L)                    '(eq 15.5.1.2-10) max value

        σbBC = 2 * a * (L * Cos(θ) - L1 * (1 - Sin(θ)))
        σbBC += L ^ 2
        σbBC *= _P
        σbBC += 2 * Ma
        σbBC *= ee / (4 * I1)               '(eq 15.5.1.2-9)

        '------- bend at B -------
        σbB += (ee / (4 * I1)) * (2 * Ma + _P * L ^ 2)    '(eq 15.5.1.2-8)

        '------- bend at A -------
        σbA += Ma * ee / (2 * I1)           '(eq 15.5.1.2-7)

        '------- bend at D -------
        σbD = _P * (2 * a * L - 2 * a * L1 + L ^ 2 - L1 ^ 2)
        σbD += 2 * Ma
        σbD *= ee / (4 * I1)                '(eq 15.5.1.2-6)

        '------- bend at C -------
        σbC = _P * (2 * a * L - 2 * a * L1 + L ^ 2)
        σbC += 2 * Ma
        σbC *= ee / (4 * I1)                '(eq 15.5.1.2-5)

        '----- total stress---
        σTA = Abs(σmA) + Abs(σbA)
        σTB = Abs(σmB) + Abs(σbB)
        σTC = Abs(σmC) + Abs(σbC)
        σTD = Abs(σmD) + Abs(σbD)
        σTBC = Abs(σmBC) + Abs(σbBC)

        '----- pressure -----
        TextBox25.Text = _P.ToString("0.00")        '[MPa]
        TextBox132.Text = (_P * 10).ToString("0.00")    '[bar]
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
        TextBox26.BackColor = CType(IIf(σTA > 1.5 * _f02, Color.Red, Color.LightGreen), Color)
        TextBox27.BackColor = CType(IIf(σTB > 1.5 * _f02, Color.Red, Color.LightGreen), Color)
        TextBox36.BackColor = CType(IIf(σTC > 1.5 * _f02, Color.Red, Color.LightGreen), Color)
        TextBox37.BackColor = CType(IIf(σTD > 1.5 * _f02, Color.Red, Color.LightGreen), Color)
        TextBox38.BackColor = CType(IIf(σTBC > 1.5 * _f02, Color.Red, Color.LightGreen), Color)

        '----- vessel size
        TextBox40.Text = CType((L + a) * CDbl(2.ToString("0.0")), String)
        TextBox41.Text = CType((L1 + a) * CDbl(2.ToString("0.0")), String)

    End Sub
    Private Sub Calc_rect_reinforced_15_6_2()
        'Reinforcemenis a simple rectangle strip
        Dim h_longe, H_shorte As Double    'Lenght inside vessel 
        Dim tw, hrib, br, e, Q1, Q2, Q, yc, bcw As Double
        Dim τw, τr As Double
        Dim I2_rib, I2_wall As Double
        Dim I_11 As Double                          '2nd moment of area short side plate and reinforcement combined
        Dim I1_rib As Double                        '1st Moment of area
        Dim area_rib, area_wall, area_composed As Double       'Areas
        Dim c_rib, c_wall, c_total As Double        'Centriods
        Dim rib_stab As Double                      'Reinforcement stability
        Dim j As Double 'j is distance weld to neutral axis composite to centroid A'
        Dim τ_long, τ_short As Double               'τ welds
        Dim ε As Double                             'ratio

        e = NumericUpDown38.Value           'Vessel-Plate tickness
        tw = NumericUpDown19.Value          'Reinforcement rib thickness
        hrib = NumericUpDown33.Value        'Reinforcement rib height
        br = NumericUpDown35.Value          'Reinforcement rib distance

        h_longe = NumericUpDown36.Value      'Lenght inside vessel (height or width)
        H_shorte = NumericUpDown35.Value     'Lenght inside vessel (height or width)

        '---- calc centroid composite----
        '--- presure side wall is "position zero"
        area_rib = tw * hrib            'Area reinforcement rib
        area_wall = e * br              'Area vessel wall
        area_composed = area_rib + area_wall

        '--------- yc= Distance inside vessel to neutral axis ---------
        'point zero = inside wall vessel
        c_rib = e + (hrib / 2)
        c_wall = e / 2
        c_total = (area_rib * c_rib) + (area_wall * c_wall)
        yc = c_total / area_composed

        '--------- j is distance weld to neutral axis composite to centroid A'-------
        j = yc - e

        '---- calcu, 1st Moment of area rib --
        I1_rib = 0.25 * tw * hrib ^ 2

        '---- calcu, 2nd Moment of area --
        I2_rib = tw * hrib ^ 3 / 12       '2nd Moment of area
        I2_wall = br * e ^ 3 / 12         '2nd Moment of area

        '---- now Parallel axis theorem ---
        I_11 = I2_wall + area_wall * (yc - e / 2) ^ 2
        I_11 += I2_rib + area_rib * (yc - hrib / 2) ^ 2

        '---- Web thickness page 327 for (C1 type)--
        bcw = e

        '----------- Shear load one side ------------------
        Q1 = _P * br * h_longe / 2 'Side Load 1 (15.6.2.3-2)
        Q2 = _P * br * H_shorte / 2 'Side Load 2 (15.6.2.3-2)   
        Q = CDbl(IIf(Q1 > Q2, Q1, Q2))    'Find biggest Shear load

        τw = Q * area_rib * j / (I_11 * bcw)    'Stress weld (15.6.2.2-1)

        τr = Q / area_rib                       'stress reinforcement(15.6.2.3-1)

        '---------- rib compression stabiliy ------
        rib_stab = hrib / tw                'Table 15.6-1 Sketch C1
        ε = Sqrt(235 / _f02)

        '---------- ΔM pressure loads page 329 ------
        '---------- Chapter 15.6.2.3 ----------
        Dim ΔM_long, ΔM_short As Double
        Dim η, lw As Double
        Dim Ma, Mb, Mc, Md As Double
        Dim S As Double     'First moment of area reinforcement
        Dim biw As Double   'Total weld throat of intermittent weld

        '-------- get data ------------
        Double.TryParse(TextBox146.Text, Ma)    'Formula 15.6.5-3
        Double.TryParse(TextBox147.Text, Mb)    'Formula 15.6.5-5
        Double.TryParse(TextBox148.Text, Mc)    'Formula 15.6.5-7
        Double.TryParse(TextBox149.Text, Md)    'Formula 15.6.5-9
        Double.TryParse(TextBox134.Text, S)     'First moment of area reinforcement

        biw = NumericUpDown39.Value             'Total weld throat of intermittent weld
        lw = NumericUpDown40.Value              'Length intermittent welds Figure 15.6-3

        '----- calc 15.6.2.3 ------------
        'Note I_11 identical to I_21 (reinforcements both side are identical)

        η = h_longe / 2 - 1.5 * lw       'Figure 15.6-3

        ΔM_long = br * _P * (H_shorte ^ 2 / 8 - η ^ 2 / 2)   'Formula 15.6.2.2-3
        ΔM_short = br * _P * (h_longe ^ 2 / 8 - η ^ 2 / 2)   'Formula 15.6.2.2-4

        τ_long = Abs(ΔM_long * S / (biw * lw * I_11))       'Formula 15.6.2.2-2
        τ_short = Abs(ΔM_short * S / (biw * lw * I_11))     'Formula 15.6.2.2-2

        '---------- present results -----------
        TextBox21.Text = area_composed.ToString("0")
        TextBox154.Text = area_rib.ToString("0")
        TextBox113.Text = j.ToString("0.0")
        TextBox114.Text = I_11.ToString("0")
        TextBox115.Text = bcw.ToString("0.00")
        TextBox124.Text = CType(Q1 / CDbl(1000.ToString("0")), String)
        TextBox125.Text = CType(Q2 / CDbl(1000.ToString("0")), String)

        TextBox116.Text = CType(Q / CDbl(1000.ToString("0")), String)
        TextBox117.Text = τw.ToString("0")
        TextBox122.Text = I2_rib.ToString("0")
        TextBox123.Text = I2_wall.ToString("0")

        TextBox126.Text = τr.ToString("0")
        TextBox127.Text = _P.ToString("0.00")

        TextBox128.Text = rib_stab.ToString("0.0")

        TextBox129.Text = h_longe.ToString("0")
        TextBox130.Text = H_shorte.ToString("0")
        TextBox134.Text = I1_rib.ToString("0")
        TextBox141.Text = yc.ToString("0.0")

        TextBox138.Text = η.ToString("0")
        TextBox139.Text = ΔM_long.ToString("0.0")
        TextBox135.Text = ΔM_short.ToString("0.0")

        TextBox142.Text = τ_long.ToString("0.0")
        TextBox143.Text = τ_short.ToString("0.0")

        '----------- check -------------
        TextBox117.BackColor = CType(IIf(τw > _f02, Color.Red, Color.LightGreen), Color)
        TextBox126.BackColor = CType(IIf(τr > _f02, Color.Red, Color.LightGreen), Color)
        TextBox128.BackColor = CType(IIf(rib_stab > 10 * ε, Color.Red, Color.LightGreen), Color)

        TextBox142.BackColor = CType(IIf(τ_long > _f02, Color.Red, Color.LightGreen), Color)
        TextBox143.BackColor = CType(IIf(τ_short > _f02, Color.Red, Color.LightGreen), Color)
    End Sub

    Private Sub Calc_rect_unreinforced_15_6_4()
        '=======15.6.4 Wall stress in unsupported zones=============
        Dim h_long As Double    'long edge
        Dim H_short As Double   'short edge
        Dim e_wall As Double    'Wall thickness
        Dim σm As Double        'Longitudinal membrane stress
        Dim σb As Double        'Longitudinal bending stress 
        Dim C As Double         'Constant
        Dim g, b, ratio As Double

        Double.TryParse(TextBox129.Text, h_long)    'Long side [mm]
        Double.TryParse(TextBox130.Text, H_short)   'Short side [mm]
        e_wall = NumericUpDown38.Value              'Wall [mm]

        σm = (_P * h_long * H_short) / (2 * e_wall * (h_long + H_short))  '(15.6.4-1)

        b = CDbl(IIf(H_short < h_long, H_short, h_long))  'Shorter one
        g = CDbl(IIf(H_short > h_long, H_short, h_long))  'Longer one
        ratio = g / b               'Ratio edge length

        Select Case True
            Case (ratio >= 1 And ratio < 1.2)
                C = 0.3078
            Case (ratio >= 1.2 And ratio < 1.4)
                C = 0.3834
            Case (ratio >= 1.4 And ratio < 1.6)
                C = 0.4356
            Case (ratio >= 1.6 And ratio < 1.8)
                C = 0.468
            Case (ratio >= 1.8 And ratio < 2)
                C = 0.4872
            Case (ratio >= 2 And ratio < 2.15)
                C = 0.4974
            Case (ratio >= 2.15)
                C = 0.5
        End Select

        σb = _P * C * (b / e_wall) ^ 2              '(15.6.4-2)

        TextBox118.Text = σm.ToString("0")    'Longitudinal membrane stress
        TextBox119.Text = σb.ToString("0")    'Longitudinal bending stress 
        TextBox120.Text = C.ToString("0.0000")
        TextBox121.Text = ratio.ToString("0.00")

        '----------- check -------------
        Label370.BackColor = CType(IIf(NumericUpDown37.Value > NumericUpDown36.Value, Color.Red, Color.LightGreen), Color)

        TextBox118.BackColor = CType(IIf(σm > _fs, Color.Red, Color.LightGreen), Color)
        TextBox119.BackColor = CType(IIf(((σm + σb) > 1.5 * _f02), Color.Red, Color.LightGreen), Color)
    End Sub

    Private Sub Calc_rect_reinforced_15_6_5()
        Dim σmD, σmA As Double  'Membrane stress
        Dim h_long As Double    'long edge
        Dim H_short As Double   'short edge
        Dim h1_long As Double   'Pitch neutral lines reinforcements on long sides of vessel
        Dim H1_short As Double  'Pitch neutral lines reinforcements on short sides of vessel
        Dim inside_to_neutral As Double 'Distande
        Dim c_fibre As Double   'Distance neutral axis to outer fibre
        Dim br, hr As Double    'Pressure bearing width [mm]
        Dim e_wall As Double    'Wall thickness
        Dim A1, A2 As Double    'Area composite [mm]
        Dim Ma, Mb, Mc, Md As Double 'Bending moment
        Dim κ, α1 As Double
        Dim σbA, σbB, σbC, σbD As Double
        Dim I11 As Double '2nd moment of area short side plate and reinforcement combined
        Dim I21 As Double '2nd moment of area long side plate and reinforcement combined

        '---------- get data ----------------
        h_long = NumericUpDown36.Value      'Long side [mm]
        H_short = NumericUpDown37.Value     'Short side [mm]
        br = NumericUpDown35.Value          'Pressure bearing width [mm]
        e_wall = NumericUpDown38.Value      'Wall [mm]
        hr = NumericUpDown33.Value          'Hoogte rib [mm]

        '--------- Pitch neutral lines reinforcements -----------
        h1_long = h_long + hr       '[mm] long side vessel
        H1_short = H_short + hr     '[mm] short side vessel

        '-------- reinforcement idetical short and long side ------
        Double.TryParse(TextBox154.Text, A1) 'Area composite [mm]
        A2 = A1   'Area composite [mm]

        '-------- Membrane stress ------------
        σmD = _P * h_long * br / (2 * A1)     'Short side [N/mm]
        σmA = _P * H_short * br / (2 * A2)    'Long side [N/mm]

        '--------- Ratio's ---------------
        α1 = H_short / h_long       'Formula (15.5.2-5)
        κ = α1 * 1      'Formula (15.5.2-4) Reinforment short and long sides are identical

        '------- Ma bending moment (Formula 15.6.5-3)-------------
        Ma = _P * br * h_long ^ 2 / 24
        Ma *= 3 - (2 * (1 + α1 ^ 2 * κ) / (1 + κ))  '

        '------- Mb bending moment (Formula 15.6.5-5) -------------
        Mb = _P * br * h_long ^ 2 / 12
        Mb *= (1 + α1 ^ 2 * κ) / (1 + κ)

        '------- Mc bending moment (Formula 15.6.5-7) -------------
        Mc = _P * br * h_long ^ 2 / 12
        Mc *= (1 + α1 ^ 2 * κ) / (1 + κ)

        '------- Md bending moment (Formula 15.6.5-9) -------------
        Md = _P * br * h_long ^ 2 / 24
        Md *= 3 * α1 ^ 2 - (2 * (1 + α1 ^ 2 * κ) / (1 + κ))

        '------- Bending stress ----
        Double.TryParse(TextBox141.Text, inside_to_neutral)
        c_fibre = hr - inside_to_neutral 'Distance neutral axis to outer fibre
        Double.TryParse(TextBox114.Text, I11)
        I21 = I11

        σbA = Abs(Ma * c_fibre / I21)   '(Formula 15.6.5-4)
        σbB = Abs(Mb * c_fibre / I21)   '(Formula 15.6.5-6)
        σbC = Abs(Mc * c_fibre / I11)   '(Formula 15.6.5-8)
        σbD = Abs(Md * c_fibre / I11)   '(Formula 15.6.5-10)

        '------- present -----------
        TextBox150.Text = σmD.ToString("0.0")
        TextBox151.Text = σmA.ToString("0.0")
        TextBox146.Text = (Ma / 10 ^ 3).ToString("0") '[Nm]
        TextBox147.Text = (Mb / 10 ^ 3).ToString("0") '[Nm]
        TextBox148.Text = (Mc / 10 ^ 3).ToString("0") '[Nm]
        TextBox149.Text = (Md / 10 ^ 3).ToString("0") '[Nm]

        TextBox152.Text = α1.ToString("0.00")
        TextBox153.Text = κ.ToString("0.00")

        TextBox155.Text = σbA.ToString("0")
        TextBox156.Text = σbB.ToString("0")
        TextBox157.Text = σbC.ToString("0")
        TextBox158.Text = σbD.ToString("0")
        TextBox159.Text = c_fibre.ToString("0.0")

        '----------- check -------------
        TextBox150.BackColor = CType(IIf(σmD > _f02, Color.Red, Color.LightGreen), Color)
        TextBox151.BackColor = CType(IIf(σmA > _f02, Color.Red, Color.LightGreen), Color)

        TextBox155.BackColor = CType(IIf(σbA > _f02 * 1.5, Color.Red, Color.LightGreen), Color)
        TextBox156.BackColor = CType(IIf(σbB > _f02 * 1.5, Color.Red, Color.LightGreen), Color)
        TextBox157.BackColor = CType(IIf(σbC > _f02 * 1.5, Color.Red, Color.LightGreen), Color)
        TextBox158.BackColor = CType(IIf(σbD > _f02 * 1.5, Color.Red, Color.LightGreen), Color)
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click, NumericUpDown3.ValueChanged, NumericUpDown2.ValueChanged, GroupBox11.Enter, ComboBox3.SelectedIndexChanged
        Calc_kloepper_7_5_3()
    End Sub

    Private Sub Calc_kloepper_7_5_3()
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
        TextBox48.BackColor = CType(IIf(E_kloepper > Wall, Color.Red, Color.LightGreen), Color)

    End Sub
    Private Function Calc_kloepper_beta() As Double
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

        C1 = CDbl(IIf(C1a > C1b, C1a, C1b))   'Find biggest

        '-----------Calc C2 method 10.4-5 ---------- 

        g = Di / (Di + es)                  '(10.4-16)

        H = (12 * (1 - _ν ^ 2)) ^ 0.25       '(10.4-17) 
        H *= Sqrt(es / (Di + es))

        J = 3 * _fs / _P                     '(10.4-18) 
        J -= Di ^ 2 / (4 * (Di + es) * es)
        J -= 1

        U = 2 * (2 - _ν * g)
        U /= Sqrt(3 * (1 - _ν ^ 2))              '(10.4-19) 

        f1 = 2 * g ^ 2 - g ^ 4                  '(10.4-20) 

        A = 3 * U * Di / (4 * es)               '(10.4-21) 
        A -= 2 * J
        A *= (1 + _ν)
        A *= (1 + (1 - _ν) * es / (Di + es))

        B = 3 * U * Di / (8 * es)               '(10.4 - 22)
        B -= J
        B *= H ^ 2
        B -= (3 / 2 * (2 - _ν * g) * g)
        B *= H

        F = 3 / 8 * U * g                       '(10.4-23) 
        F += 3 / 16 * f1 * (Di + es) / es
        F -= (2 * J * es / (Di + es))
        F *= H ^ 2
        F -= (3 * (2 - _ν * g) * g * es / (Di + es))

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
        E_flat = CDbl(IIf(E_flat1 > E_flat2, E_flat1, E_flat2))   'The biggest

        TextBox56.Text = _P.ToString("0.00")
        TextBox55.Text = (_P * 10).ToString("0.0")
        TextBox60.Text = Di.ToString("0")

        TextBox57.Text = C1.ToString("0.000")
        TextBox54.Text = C2.ToString("0.000")

        TextBox65.Text = E_flat1.ToString("0.0")
        TextBox64.Text = E_flat2.ToString("0.0")
        TextBox58.Text = E_flat.ToString("0.0")
        TextBox62.Text = _fs.ToString("0")

        '----Chart determine C1 (10.4-4)
        TextBox59.Text = (es / Di).ToString("0.000")
        TextBox61.Text = (_P / _fs).ToString("0.000")

        '----Chart determine C2 (10.4-5)
        TextBox59.Text = (es / Di).ToString("0.000")
        TextBox61.Text = (_P / _fs).ToString("0.0000")

        'Checks
        TextBox57.BackColor = CType(IIf(C1 > 0.29 And C1 < 0.42, Color.LightGreen, Color.Red), Color)
        TextBox54.BackColor = CType(IIf(C2 >= 0.3 And C2 < 1.0, Color.LightGreen, Color.Red), Color)

    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click, TabPage8.Enter, NumericUpDown22.ValueChanged, NumericUpDown6.ValueChanged
        '10.5.3 Flat end with a full-face gasket------------ 
        Dim dia_bolt_circle As Double
        Dim d, d_procent, G, Y2, e_pierced As Double
        Dim e_flange As Double
        Dim e1 As Double

        dia_bolt_circle = NumericUpDown22.Value
        d_procent = NumericUpDown6.Value / 100


        e_flange = 0.41 * dia_bolt_circle * Sqrt(_P / _fs)  'Fornula (10.5-7)
        e1 = e_flange * 0.8         'Formula (10.5-8) dikte flange at the bolts

        '10.6.1 Pierced cicular flat ends------------- 
        d = d_procent * dia_bolt_circle 'Chapter 10.6.1.3 max 50%
        G = dia_bolt_circle             'Equation (10.6-4) 
        Y2 = Sqrt(G / (G - d))          'Equation (10.6-4)
        e_pierced = e_flange * Y2       'Equation (10.6-2)


        TextBox68.Text = _P.ToString("0.00")
        TextBox67.Text = (_P * 10).ToString("0.0")
        TextBox63.Text = _fs.ToString("0.0")        'Allowed nominal design stress 
        TextBox71.Text = e_flange.ToString("0.0")
        TextBox74.Text = e1.ToString("0.0")
        TextBox73.Text = d.ToString("0")
        TextBox112.Text = Y2.ToString("0.00")
        TextBox111.Text = e_pierced.ToString("0.0")
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click, TabPage9.Enter, NumericUpDown24.ValueChanged, NumericUpDown23.ValueChanged, NumericUpDown31.ValueChanged, NumericUpDown30.ValueChanged, NumericUpDown29.ValueChanged, NumericUpDown28.ValueChanged, NumericUpDown34.ValueChanged, NumericUpDown32.ValueChanged, ComboBox4.SelectedIndexChanged, NumericUpDown27.ValueChanged, NumericUpDown26.ValueChanged, NumericUpDown25.ValueChanged, ComboBox6.SelectedIndexChanged
        Calc_flange_Moments_11_5_3()
    End Sub

    Sub Calc_flange_Moments_11_5_3()
        Dim words() As String
        Dim e, G, gt, HG, H, B, C, A As Double
        Dim db, dn, n As Double
        Dim fB As Double
        Dim W, w_, b_gasket, b0_gasket, m As Double
        Dim y, Wa, Wop As Double
        Dim AB_min1, AB_min2, AB_min, dia_bolt_circle As Double

        Dim HD, HT As Double
        Dim hD_, hG_, hT_ As Double
        Dim Ma, Mop As Double
        Dim g0_ As Double       'is the thickness of hub at small end
        Dim g1_ As Double       'is the thickness of hub at back of flange;
        Dim βT, βU, βY As Double

        Dim CF, δb, K, I0 As Double
        Dim M1, M2, σθ As Double
        Dim temp As Double

        TextBox237.Clear()

        If (ComboBox4.SelectedIndex > -1) Then          'Prevent exceptions
            words = gaskets(ComboBox4.SelectedIndex).Split(separators, StringSplitOptions.None)
            Double.TryParse(words(1), temp)    'Gasket factor m
            NumericUpDown30.Value = CDec(temp)
            Double.TryParse(words(2), temp)    'Gasket seat pressure y
            NumericUpDown31.Value = CDec(temp)
        End If

        If (ComboBox6.SelectedIndex > -1) Then          'Prevent exceptions
            words = Bolt(ComboBox6.SelectedIndex + 1).Split(separators, StringSplitOptions.None)

            Double.TryParse(words(1), temp)    'Bolt stress
            NumericUpDown27.Value = CDec(temp / 3)
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
        AB_min = CDbl(IIf(AB_min1 < AB_min2, AB_min2, AB_min1))   'Take biggest

        '---- dia bolt---
        dia_bolt_circle = Sqrt((AB_min / n) * 4 / PI)

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
        K = A / B                                   '(11.5-21)
        g0_ = e
        I0 = Sqrt(B * g0_)                           '(11.5-22)

        βT = (K ^ 2 * (1 + 8.55246 * Log10(K))) - 1 '(11.5-23)
        βT /= (1.0472 + 1.9448 * K ^ 2) * (K - 1)

        βU = (K ^ 2 * (1 + 8.55246 * Log10(K))) - 1 '(11.5-24)
        βU /= 1.36136 * (K ^ 2 - 1) * (K - 1)

        βY = 1 / (K - 1)                            '(11.5-24)
        βY *= 0.66845 + 5.7169 * (K ^ 2 * Log10(K)) / (K ^ 2 - 1)

        M1 = Ma * CF / B                '(11.5-26)
        M2 = Mop * CF / B               '(11.5-27)

        TextBox237.AppendText("δb = " & δb.ToString & vbCrLf)
        TextBox237.AppendText("CF= " & CF.ToString & vbCrLf)
        TextBox237.AppendText("K= " & K.ToString & vbCrLf)
        TextBox237.AppendText("g0_= " & g0_.ToString & vbCrLf)
        TextBox237.AppendText("I0 = " & I0.ToString & vbCrLf)
        TextBox237.AppendText("βT= " & βT.ToString & vbCrLf)
        TextBox237.AppendText("βU= " & βU.ToString & vbCrLf)
        TextBox237.AppendText("βY= " & βY.ToString & vbCrLf)
        TextBox237.AppendText("M1= " & M1.ToString & vbCrLf)
        TextBox237.AppendText("M2= " & M2.ToString & vbCrLf)


        '----------- Integral method ---------------
        '---11.5.4.1.2 Coefficients for flange stresses calculations
        'alle coefficienten eindigen met een underscore !!
        Dim A_, C_ As Double

        Dim C1, C2, C3, C4, C5, C6, C7, C8, C9, C10 As Double
        Dim C11, C12, C13, C14, C15, C16, C17, C18, C19, C20 As Double
        Dim C21, C22, C23, C24, C25, C26, C27, C28, C29, C30 As Double
        Dim C31, C32, C33, C34, C35, C36, C37 As Double

        g0_ = 99
        g1_ = 99


        A_ = (g1_ / g0_) - 1                    '(11.5-43)
        C_ = 48 * (1 - _ν ^ 2) * (H / L0) ^ 4   '(11.5-44)

        C1 = 1 / 3 + A_ / 12
        C2 = 5 / 42 + 17 * A_ / 336
        C3 = 1 / 210 + A_ / 360
        C4 = 11 / 360 + 59 * A_ / 5040 + (1 + 3 * A_) / C_
        C5 = 1 / 90 + 5 * A_ / 1008 - (1 + A) ^ 3 / C
        C6 = 1 / 120 + 17 * A / 5040 + 1 / C_
        C7 = (215 / 2772 + 51 * A_ / 1232 + (120 + 225 * A_ + 150 * A_ ^ 2 + 35 * A_ ^ 3)) / (14 * C_)
        C8 = (31 / 6930 + 128 * A_ / 45045 + (66 + 165 * A_ + 132 * A_ ^ 2 + 35 * A ^ 3)) / (77 * C_)
        C9 = (553 / 30240 + 653 * A_ / 73920 + (42 + 198 * A_ + 117 * A_ ^ 2 + 25 * A ^ 3)) / (84 * C_)
        C10 = (29 / 3780 + 3 * A_ / 704 - (42 + 198 * A + 243 * A_ ^ 2 + 91 * A_ ^ 3)) / (84 * C_)
        C11 = (31 / 6048 + 1763 * A_ / 665280 + (42 + 72 * A_ + 45 * A_ ^ 2 + 10 * A_ ^ 3)) / (84 * C_)
        C12 = (1 / 2925 + 71 * A_ / 300300 + (88 + 198 * A_ + 156 * A_ ^ 2 + 42 * A_ ^ 3)) / (385 * C_)
        C13 = (761 / 831600 + 937 * A_ / 1663200 + (2 + 12 * A_ + 17 * A_ ^ 2 + 3 * A ^ 3)) / (70 * C_)
        C14 = (197 / 415600 + 97 * A_ / 554400 + (6 + 18 * A + 15 * A ^ 2 + 4 * A_ ^ 3)) / (210 * C_)
        C15 = (233 / 831600 + 97 * A / 554400 + (6 + 18 * A_ + 15 * A ^ 2 + 4 * A ^ 3)) / (210 * C_)
        C16 = C1 * C7 * C12 + C2 * C8 * C3 + C3 * C8 * C2 - (C3 ^ 2 * C7 + C8 ^ 2 * C1 + C2 ^ 2 * C12)
        C17 = (C4 * C7 * C12 + C2 * C8 * C13 + C3 * C8 * C9 - (C13 * C7 * C3 + C8 ^ 2 * C4 + C12 * C2 * C9)) / C16
        C18 = (C5 * C7 * C12 + C2 * C8 * C14 + C3 * C8 * C10 - (C14 * C7 * C3 + C8 ^ 2 * C5 + C12 * C2 * C10)) / C16
        C19 = (C6 * C7 * C13 + C2 * C8 * C15 + C3 * C8 * C11 - (C15 * C7 * C3 + C8 ^ 2 * C6 + C12 * C2 * C11)) / C16
        C20 = (C1 * C9 * C12 + C4 * C8 * C3 + C3 * C13 * C2 - (C3 ^ 2 * C9 + C13 * C8 * C1 + C12 * C4 * C2)) / C16
        C21 = (C1 * C10 * C12 + C5 * C8 * C3 + C3 * C14 * C2 - (C3 ^ 2 * C10 + C14 * C8 * C1 + C12 * C5 * C2)) / C16
        C22 = (C1 * C11 * C12 + C6 * C8 * C3 + C3 * C15 * C2 - (C3 ^ 2 * C11 + C15 * C8 * C1 + C12 * C6 * C2)) / C16
        C23 = (C1 * C7 * C13 + C2 * C8 * C3 + C4 * C8 * C2 - (C3 * C7 * C4 + C8 * C9 * C1 + C2 ^ 2 * C13)) / C16
        C24 = (C1 * C7 * C14 + C2 * C10 * C3 + C5 * C8 * C2 - (C3 * C7 * C5 + C8 * C10 * C1 + C2 ^ 2 * C14)) / C16
        C25 = (C1 * C7 * C15 + C2 * C11 * C3 + C6 * C8 * C2 - (C3 * C7 * C6 + C8 * C11 * C1 + C2 ^ 2 ^ C15)) / C16
        C26 = -(C_ / 4) ^ 0.25
        C27 = C20 - C17 - 5 / 12 * C17 * C26
        C28 = C22 - C19 - 1 / 12 + C19 * C26    '(11.5-73)
        C29 = -(C_ / 4) ^ 0.5
        C30 = -(C_ / 4) ^ 0.75
        C31 = 3 * A_ / 2 - C17 * C30
        C32 = 1 / 2 - C19 * C30
        C33 = C26 * C32 / 2 + C28 * C31 * C29 - (C30 * C28 / 2 + C31 * C27 * C29)
        C34 = 1 / 12 + C18 - C21 - C18 * C26
        C35 = C18 * C30
        C36 = (C28 * C35 * C29 - C32 * C34 * C29) / C33
        C37 = 99

        TextBox237.AppendText("----- Integral method -----" & vbCrLf)
        Dim βf As Double
        βf = 0.90892   'Cylindrical hub         '(11.5-28)

        TextBox237.AppendText("βf= " & βf.ToString & vbCrLf)

        Dim βv As Double
        βv = 0.550103   'Cylindrical hub    '(11.5-29)
        TextBox237.AppendText("βv= " & βv.ToString & vbCrLf)

        Dim φ As Double
        φ = C36 / (1 + A_)                       '(11.5-30)
        TextBox237.AppendText("A= " & A.ToString & vbCrLf)
        TextBox237.AppendText("C36= " & C36.ToString & vbCrLf)
        TextBox237.AppendText("φ= " & φ.ToString & vbCrLf)

        Dim λ As Double
        λ = e * βf + I0 / (βT * I0)             '(11.5-31)
        λ += e ^ 3 * βv / (βU * I0 * g0_ ^ 2)
        TextBox237.AppendText("λ= " & λ.ToString & vbCrLf)

        '-------- Longitudinal hub stress -----
        Dim σH As Double
        σH = φ * m / (λ * g0_ ^ 2)               '(11.5-32)
        TextBox237.AppendText("σH= " & σH.ToString & vbCrLf)

        '-------- Radial flange stress -----
        Dim σr As Double
        σr = (1.333 * e * βf + I0) * m
        σr /= λ * e ^ 2 * I0                    '(11.5-33)
        TextBox237.AppendText("σr= " & σr.ToString & vbCrLf)

        '---------- Tangential flange stress -----
        σθ = βY * m / e ^ 2                     '(11.5-34)
        σθ -= K ^ 2 + 1 / (K ^ 2 - 1)
        TextBox237.AppendText("σθ= " & σθ.ToString & vbCrLf)

        '==================================================
        '----------- Loose flange method (Consists od 2 parts)  ----
        σθ = βY * M2 / e ^ 2                    '(11.5-35)

        TextBox77.Text = (_P * 10).ToString("0.0")
        TextBox78.Text = _P.ToString("0.00")
        TextBox76.Text = _fs.ToString("0")
        TextBox79.Text = H.ToString("0.0")
        TextBox72.Text = b0_gasket.ToString("0.0")
        TextBox81.Text = b_gasket.ToString("0.0")
        TextBox93.Text = G.ToString("0.0")
        TextBox84.Text = dn.ToString("0.0")             '[mm]
        TextBox94.Text = AB_min.ToString("0")           '[mm2] required bolt area
        TextBox95.Text = dia_bolt_circle.ToString("0.0") '[mm] calculated req. dia bolt

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
        NumericUpDown28.BackColor = CType(IIf(C <= B, Color.Red, Color.Yellow), Color)    'Bolt diameter
        NumericUpDown34.BackColor = CType(IIf(A <= C, Color.Red, Color.Yellow), Color)    'Flange OD
        NumericUpDown24.BackColor = CType(IIf(w_ > (A - B) / 2, Color.Red, Color.Yellow), Color)    'Gasket width
        '----- bolts ---------
        NumericUpDown26.BackColor = CType(IIf(dia_bolt_circle > db, Color.Red, Color.Yellow), Color) 'Bolt dia
        TextBox95.BackColor = CType(IIf(dia_bolt_circle > db, Color.Red, Color.LightGreen), Color)   'Bolt dia
        '------ gasket outside diameter --------
        NumericUpDown29.BackColor = CType(IIf(gt > A, Color.Red, Color.Yellow), Color) 'Bolt dia
        '----- flange stress -----
        TextBox102.BackColor = CType(IIf(σθ > _fs, Color.Red, Color.LightGreen), Color)    'Flange stress
    End Sub
    Private Sub PictureBox8_Click(sender As Object, e As EventArgs) Handles PictureBox8.Click
        Form2.Show()
    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        Dim oWord As Word.Application

        Dim oDoc As Word.Document
        Dim oTable As Word.Table
        Dim oPara1, oPara2 As Word.Paragraph
        Dim row, font_sizze As Integer
        Dim ufilename As String
        ufilename = "PV_calc_" & TextBox66.Text & "_" & TextBox69.Text & "_" & TextBox70.Text & DateTime.Now.ToString("yyyy_MM_dd") & ".docx"

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

            oTable.Columns(1).Width = oWord.InchesToPoints(2.91)   'Change width of columns 1 & 2.
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

            oTable.Columns(1).Width = oWord.InchesToPoints(2.91)   'Change width of columns 1 & 2.
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


            oTable.Columns(1).Width = oWord.InchesToPoints(2.91)   'Change width of columns 1 & 2.
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
            oTable.Cell(row, 3).Range.Text = "[kN]"
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

            oTable.Columns(1).Width = oWord.InchesToPoints(2.91)   'Change width of columns 1 & 2.
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

            oTable.Columns(1).Width = oWord.InchesToPoints(2.91)   'Change width of columns 1 & 2.
            oTable.Columns(2).Width = oWord.InchesToPoints(1.5)
            oTable.Columns(3).Width = oWord.InchesToPoints(1.3)
            oDoc.Bookmarks.Item("\endofdoc").Range.InsertParagraphAfter()


            '---------"Flange stress 11.5.4"---------------------------------------
            oTable = oDoc.Tables.Add(oDoc.Bookmarks.Item("\endofdoc").Range, 8, 3)
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
            oTable.Cell(row, 1).Range.Text = "Flange factors K, βT"
            oTable.Cell(row, 2).Range.Text = TextBox99.Text & " - " & TextBox98.Text
            oTable.Cell(row, 3).Range.Text = "[-][-]"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Flange factors βU, βY"
            oTable.Cell(row, 2).Range.Text = TextBox97.Text & " - " & TextBox96.Text
            oTable.Cell(row, 3).Range.Text = "[-][-]"
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

            oTable.Columns(1).Width = oWord.InchesToPoints(2.91)   'Change width of columns 1 & 2.
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

    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click
        Save_tofile()
    End Sub
    'Save control settings and case_x_conditions to file
    Private Sub Save_tofile()
        Dim temp_string As String
        Dim filename As String = "PV_Calc_" & TextBox66.Text & "_" & TextBox69.Text & "_" & TextBox70.Text & DateTime.Now.ToString("_yyyy_MM_dd") & ".vtk"
        Dim all_num, all_combo, all_check, all_radio As New List(Of Control)
        Dim i As Integer

        If String.IsNullOrEmpty(TextBox8.Text) Then
            TextBox8.Text = "-"
        End If

        temp_string = TextBox66.Text & ";" & TextBox69.Text & ";" & TextBox70.Text & ";"
        temp_string &= vbCrLf & "BREAK" & vbCrLf & ";"

        '-------- find all numeric, combobox, checkbox and radiobutton controls -----------------
        FindControlRecursive(all_num, Me, GetType(NumericUpDown))   'Find the control
        all_num = all_num.OrderBy(Function(x) x.Name).ToList()      'Alphabetical order
        For i = 0 To all_num.Count - 1
            Dim grbx As NumericUpDown = CType(all_num(i), NumericUpDown)
            temp_string &= grbx.Value.ToString & ";"
        Next
        temp_string &= vbCrLf & "BREAK" & vbCrLf & ";"

        '-------- find all combobox controls and save
        FindControlRecursive(all_combo, Me, GetType(ComboBox))      'Find the control
        all_combo = all_combo.OrderBy(Function(x) x.Name).ToList()   'Alphabetical order
        For i = 0 To all_combo.Count - 1
            Dim grbx As ComboBox = CType(all_combo(i), ComboBox)
            temp_string &= grbx.SelectedItem.ToString & ";"
        Next
        temp_string &= vbCrLf & "BREAK" & vbCrLf & ";"

        '-------- find all checkbox controls and save
        FindControlRecursive(all_check, Me, GetType(CheckBox))      'Find the control
        all_check = all_check.OrderBy(Function(x) x.Name).ToList()  'Alphabetical order
        For i = 0 To all_check.Count - 1
            Dim grbx As CheckBox = CType(all_check(i), CheckBox)
            temp_string &= grbx.Checked.ToString & ";"
        Next
        temp_string &= vbCrLf & "BREAK" & vbCrLf & ";"

        '-------- find all radio controls and save
        FindControlRecursive(all_radio, Me, GetType(RadioButton))   'Find the control
        all_radio = all_radio.OrderBy(Function(x) x.Name).ToList()  'Alphabetical order
        For i = 0 To all_radio.Count - 1
            Dim grbx As RadioButton = CType(all_radio(i), RadioButton)
            temp_string &= grbx.Checked.ToString & ";"
        Next
        temp_string &= vbCrLf & "BREAK" & vbCrLf & ";"

        '--------- add notes -----
        temp_string &= TextBox63.Text & ";"

        '---- if path not exist then create one----------
        Try
            If (Not System.IO.Directory.Exists(dirpath_Home)) Then System.IO.Directory.CreateDirectory(dirpath_Home)
            If (Not System.IO.Directory.Exists(dirpath_Eng)) Then System.IO.Directory.CreateDirectory(dirpath_Eng)
            If (Not System.IO.Directory.Exists(dirpath_Rap)) Then System.IO.Directory.CreateDirectory(dirpath_Rap)
        Catch ex As Exception
        End Try

        Try
            If CInt(temp_string.Length.ToString) > 100 Then      'String may be empty
                If Directory.Exists(dirpath_Eng) Then
                    File.WriteAllText(dirpath_Eng & filename, temp_string, Encoding.ASCII)      'used at VTK
                Else
                    File.WriteAllText(dirpath_Home & filename, temp_string, Encoding.ASCII)     'used at home
                End If
            End If
        Catch ex As Exception
            MessageBox.Show("Line 5062, " & ex.Message)  ' Show the exception's message.
        End Try
    End Sub

    'Retrieve control settings and case_x_conditions from file
    'Split the file string into 5 separate strings
    'Each string represents a control type (combobox, checkbox,..)
    'Then split up the secton string into part to read into the parameters
    Private Sub Read_file()
        Dim control_words(), words() As String
        Dim i As Integer
        Dim ttt As Double
        Dim k As Integer = 0
        Dim all_num, all_combo, all_check, all_radio As New List(Of Control)
        Dim separators() As String = {";"}
        Dim separators1() As String = {"BREAK"}

        OpenFileDialog1.FileName = "PV_Calc*"
        If Directory.Exists(dirpath_Eng) Then
            OpenFileDialog1.InitialDirectory = dirpath_Eng  'used at VTK
        Else
            OpenFileDialog1.InitialDirectory = dirpath_Home  'used at home
        End If

        OpenFileDialog1.Title = "Open a Text File"
        OpenFileDialog1.Filter = "VTK Files|*.vtk"

        If OpenFileDialog1.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
            Dim readText As String = File.ReadAllText(OpenFileDialog1.FileName, Encoding.ASCII)
            control_words = readText.Split(separators1, StringSplitOptions.None) 'Split the read file content

            '----- retrieve case condition-----
            words = control_words(0).Split(separators, StringSplitOptions.None) 'Split first line the read file content
            TextBox66.Text = words(0)                  'Project number
            TextBox69.Text = words(1)                  'Project name
            TextBox70.Text = words(2)                  'Vessel ID

            '---------- terugzetten numeric controls -----------------
            FindControlRecursive(all_num, Me, GetType(NumericUpDown))
            all_num = all_num.OrderBy(Function(x) x.Name).ToList()                  'Alphabetical order
            words = control_words(1).Split(separators, StringSplitOptions.None)     'Split the read file content
            For i = 0 To all_num.Count - 1
                Dim grbx As NumericUpDown = CType(all_num(i), NumericUpDown)
                '--- dit deel voorkomt problemen bij het uitbreiden van het aantal numeric controls--
                If (i < words.Length - 1) Then
                    If Not (Double.TryParse(words(i + 1), ttt)) Then MessageBox.Show("Numeric controls conversion problem occured")
                    If ttt <= grbx.Maximum And ttt >= grbx.Minimum Then
                        grbx.Value = CDec(ttt)          'OK
                    Else
                        grbx.Value = grbx.Minimum       'NOK
                        MessageBox.Show("Numeric controls value out of ousode min-max range, Minimum value is used")
                    End If
                Else
                    MessageBox.Show("Warning last Numeric controls not found in file")  'NOK
                End If
            Next

            '---------- terugzetten combobox controls -----------------
            FindControlRecursive(all_combo, Me, GetType(ComboBox))
            all_combo = all_combo.OrderBy(Function(x) x.Name).ToList()                  'Alphabetical order
            words = control_words(2).Split(separators, StringSplitOptions.None) 'Split the read file content
            For i = 0 To all_combo.Count - 1
                Dim grbx As ComboBox = CType(all_combo(i), ComboBox)
                '--- dit deel voorkomt problemen bij het uitbreiden van het aantal checkboxes--
                If (i < words.Length - 1) Then
                    grbx.SelectedItem = words(i + 1)
                Else
                    MessageBox.Show("Warning last combobox not found in file")
                End If
            Next

            '---------- terugzetten checkbox controls -----------------
            FindControlRecursive(all_check, Me, GetType(CheckBox))
            all_check = all_check.OrderBy(Function(x) x.Name).ToList()                  'Alphabetical order
            words = control_words(3).Split(separators, StringSplitOptions.None) 'Split the read file content
            For i = 0 To all_check.Count - 1
                Dim grbx As CheckBox = CType(all_check(i), CheckBox)
                '--- dit deel voorkomt problemen bij het uitbreiden van het aantal checkboxes--
                If (i < words.Length - 1) Then
                    Boolean.TryParse(words(i + 1), grbx.Checked)
                Else
                    MessageBox.Show("Warning last checkbox not found in file")
                End If
            Next

            '---------- terugzetten radiobuttons controls -----------------
            FindControlRecursive(all_radio, Me, GetType(RadioButton))
            all_radio = all_radio.OrderBy(Function(x) x.Name).ToList()                  'Alphabetical order
            words = control_words(4).Split(separators, StringSplitOptions.None) 'Split the read file content
            For i = 0 To all_radio.Count - 1
                Dim grbx As RadioButton = CType(all_radio(i), RadioButton)
                '--- dit deel voorkomt problemen bij het uitbreiden van het aantal radiobuttons--
                If (i < words.Length - 1) Then
                    Boolean.TryParse(words(i + 1), grbx.Checked)
                Else
                    MessageBox.Show("Warning last radiobutton not found in file")
                End If
            Next
            '---------- terugzetten Notes -- ---------------
            If control_words.Count > 5 Then
                words = control_words(5).Split(separators, StringSplitOptions.None) 'Split the read file content
                TextBox63.Clear()
                TextBox63.AppendText(words(1))
            Else
                MessageBox.Show("Warning Notes not found in file")
            End If
        End If
    End Sub

    '----------- Find all controls on form1------
    'Nota Bene, sequence of found control may be differen, List sort is required
    Public Shared Function FindControlRecursive(ByVal list As List(Of Control), ByVal parent As Control, ByVal ctrlType As System.Type) As List(Of Control)
        If parent Is Nothing Then Return list

        If parent.GetType Is ctrlType Then
            list.Add(parent)
        End If
        For Each child As Control In parent.Controls
            FindControlRecursive(list, child, ctrlType)
        Next
        Return list
    End Function

    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click
        Read_file()
    End Sub

    Private Sub Button12_Click(sender As Object, e As EventArgs) Handles Button12.Click, TabPage10.Enter, NumericUpDown38.ValueChanged, NumericUpDown37.ValueChanged, NumericUpDown36.ValueChanged, NumericUpDown35.ValueChanged, NumericUpDown33.ValueChanged, NumericUpDown19.ValueChanged, NumericUpDown40.ValueChanged, NumericUpDown39.ValueChanged
        Calc_rect_reinforced_15_6_2()
        Calc_rect_unreinforced_15_6_4()
        Calc_rect_reinforced_15_6_5()
    End Sub

    Private Sub Button13_Click(sender As Object, e As EventArgs) Handles Button13.Click
        Print_15_6()
    End Sub
    Private Sub Print_15_6()
        Dim oWord As Word.Application ' = Nothing

        Dim oDoc As Word.Document
        Dim oTable As Word.Table
        Dim oPara1, oPara2 As Word.Paragraph
        Dim row, font_sizze As Integer
        Dim ufilename As String
        ufilename = "PV_calc_Chapter 15.6" & TextBox66.Text & "_" & TextBox69.Text & "_" & TextBox70.Text & DateTime.Now.ToString("yyyy_MM_dd") & ".docx"

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
            oPara2.Range.Text = "Rectangle Pressure Vessel Reinforced acc. EN13445 15.6" & vbCrLf
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

            oTable.Columns(1).Width = oWord.InchesToPoints(2.91)   'Change width of columns 1 & 2.
            oTable.Columns(2).Width = oWord.InchesToPoints(4)

            oTable.Rows.Item(1).Range.Font.Bold = CInt(True)
            oDoc.Bookmarks.Item("\endofdoc").Range.InsertParagraphAfter()

            '---------------- Material data --------------------
            'Insert a 18 (row) x 3 table (column), fill it with data and change the column widths.
            oTable = oDoc.Tables.Add(oDoc.Bookmarks.Item("\endofdoc").Range, 4, 3)
            oTable.Range.ParagraphFormat.SpaceAfter = 1
            oTable.Range.Font.Size = font_sizze
            oTable.Range.Font.Bold = CInt(False)
            oTable.Rows.Item(1).Range.Font.Bold = CInt(True)
            oTable.Rows.Item(1).Range.Font.Size = font_sizze + 2
            row = 1
            oTable.Cell(row, 1).Range.Text = "Material Ddata"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Material"
            oTable.Cell(row, 2).Range.Text = ComboBox5.Text
            oTable.Cell(row, 3).Range.Text = "[-]"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Max allowed membrane stress"
            oTable.Cell(row, 2).Range.Text = TextBox136.Text
            oTable.Cell(row, 3).Range.Text = "[N/mm2]"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Max allowed membrane+bend stress"
            oTable.Cell(row, 2).Range.Text = TextBox137.Text
            oTable.Cell(row, 3).Range.Text = "[N/mm2]"

            oTable.Columns(1).Width = oWord.InchesToPoints(2.91)
            oTable.Columns(2).Width = oWord.InchesToPoints(1.5)
            oTable.Columns(3).Width = oWord.InchesToPoints(1.3)
            oDoc.Bookmarks.Item("\endofdoc").Range.InsertParagraphAfter()

            '---------------- Vessel data --------------------
            'Insert a 18 (row) x 3 table (column), fill it with data and change the column widths.
            oTable = oDoc.Tables.Add(oDoc.Bookmarks.Item("\endofdoc").Range, 8, 3)
            oTable.Range.ParagraphFormat.SpaceAfter = 1
            oTable.Range.Font.Size = font_sizze
            oTable.Range.Font.Bold = CInt(False)
            oTable.Rows.Item(1).Range.Font.Bold = CInt(True)
            oTable.Rows.Item(1).Range.Font.Size = font_sizze + 2
            row = 1
            oTable.Cell(row, 1).Range.Text = "Vessel Data"
            row += 1
            oTable.Cell(row, 1).Range.Text = Label327.Text
            oTable.Cell(row, 2).Range.Text = NumericUpDown19.Text
            oTable.Cell(row, 3).Range.Text = "[mm]"
            row += 1
            oTable.Cell(row, 1).Range.Text = Label340.Text
            oTable.Cell(row, 2).Range.Text = NumericUpDown33.Text
            oTable.Cell(row, 3).Range.Text = "[mm2]"
            row += 1
            oTable.Cell(row, 1).Range.Text = Label430.Text
            oTable.Cell(row, 2).Range.Text = TextBox141.Text
            oTable.Cell(row, 3).Range.Text = "[mm]"
            row += 1
            oTable.Cell(row, 1).Range.Text = Label352.Text
            oTable.Cell(row, 2).Range.Text = TextBox114.Text
            oTable.Cell(row, 3).Range.Text = "[mm4]"
            row += 1
            oTable.Cell(row, 1).Range.Text = Label370.Text
            oTable.Cell(row, 2).Range.Text = NumericUpDown36.Value.ToString
            oTable.Cell(row, 3).Range.Text = "[mm]"
            row += 1
            oTable.Cell(row, 1).Range.Text = Label367.Text
            oTable.Cell(row, 2).Range.Text = NumericUpDown37.Value.ToString("0")
            oTable.Cell(row, 3).Range.Text = "[mm]"
            row += 1
            oTable.Cell(row, 1).Range.Text = Label373.Text
            oTable.Cell(row, 2).Range.Text = NumericUpDown38.Value.ToString("0")
            oTable.Cell(row, 3).Range.Text = "[mm]"
            row += 1
            oTable.Cell(row, 1).Range.Text = Label357.Text
            oTable.Cell(row, 2).Range.Text = NumericUpDown35.Value.ToString("0")
            oTable.Cell(row, 3).Range.Text = "[mm]"
            row += 1
            oTable.Cell(row, 1).Range.Text = Label396.Text
            oTable.Cell(row, 2).Range.Text = TextBox127.Text
            oTable.Cell(row, 3).Range.Text = "[N/mm2]"

            oTable.Columns(1).Width = oWord.InchesToPoints(2.91)
            oTable.Columns(2).Width = oWord.InchesToPoints(1.5)
            oTable.Columns(3).Width = oWord.InchesToPoints(1.3)
            oDoc.Bookmarks.Item("\endofdoc").Range.InsertParagraphAfter()

            '---------"Wall stress unsupported zone 15.6.4"---------------------------------------
            oTable = oDoc.Tables.Add(oDoc.Bookmarks.Item("\endofdoc").Range, 5, 3)
            oTable.Range.ParagraphFormat.SpaceAfter = 1
            oTable.Range.Font.Size = font_sizze
            oTable.Range.Font.Bold = CInt(False)
            oTable.Rows.Item(1).Range.Font.Bold = CInt(True)
            oTable.Rows.Item(1).Range.Font.Size = font_sizze + 2
            row = 1
            oTable.Cell(row, 1).Range.Text = "Wall stress unsupported zone 15.6.4"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Inside long side"
            oTable.Cell(row, 2).Range.Text = NumericUpDown36.Value.ToString("0")
            oTable.Cell(row, 3).Range.Text = "[mm]"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Inside short side"
            oTable.Cell(row, 2).Range.Text = NumericUpDown37.Value.ToString("0")
            oTable.Cell(row, 3).Range.Text = "[mm]"
            row += 1
            oTable.Cell(row, 1).Range.Text = Label378.Text
            oTable.Cell(row, 2).Range.Text = TextBox118.Text
            oTable.Cell(row, 3).Range.Text = "[N/mm2]"
            row += 1
            oTable.Cell(row, 1).Range.Text = Label381.Text
            oTable.Cell(row, 2).Range.Text = TextBox119.Text
            oTable.Cell(row, 3).Range.Text = "[N/mm2]"


            oTable.Columns(1).Width = oWord.InchesToPoints(2.91)   'Change width of columns 1 & 2.
            oTable.Columns(2).Width = oWord.InchesToPoints(1.5)
            oTable.Columns(3).Width = oWord.InchesToPoints(1.3)
            oDoc.Bookmarks.Item("\endofdoc").Range.InsertParagraphAfter()

            '---------Chapter 15.6.2.3 Reinforments stitch welded---------------------------------------
            oTable = oDoc.Tables.Add(oDoc.Bookmarks.Item("\endofdoc").Range, 5, 3)
            oTable.Range.ParagraphFormat.SpaceAfter = 1
            oTable.Range.Font.Size = font_sizze
            oTable.Range.Font.Bold = CInt(False)
            oTable.Rows.Item(1).Range.Font.Bold = CInt(True)
            oTable.Rows.Item(1).Range.Font.Size = font_sizze + 2
            row = 1
            oTable.Cell(row, 1).Range.Text = "Reinforments stitch welded 15.6.2.3"

            row += 1
            oTable.Cell(row, 1).Range.Text = Label415.Text
            oTable.Cell(row, 2).Range.Text = NumericUpDown39.Value.ToString("0")
            oTable.Cell(row, 3).Range.Text = "[N]"
            row += 1
            oTable.Cell(row, 1).Range.Text = Label418.Text
            oTable.Cell(row, 2).Range.Text = NumericUpDown40.Value.ToString("0")
            oTable.Cell(row, 3).Range.Text = "[N]"
            row += 1
            oTable.Cell(row, 1).Range.Text = "τ Long side stitch"
            oTable.Cell(row, 2).Range.Text = TextBox142.Text
            oTable.Cell(row, 3).Range.Text = "[N/mm2]"
            row += 1
            oTable.Cell(row, 1).Range.Text = "τ Long side stitch"
            oTable.Cell(row, 2).Range.Text = TextBox143.Text
            oTable.Cell(row, 3).Range.Text = "[N/mm2]"

            oTable.Columns(1).Width = oWord.InchesToPoints(2.91)   'Change width of columns 1 & 2.
            oTable.Columns(2).Width = oWord.InchesToPoints(1.5)
            oTable.Columns(3).Width = oWord.InchesToPoints(1.3)
            oDoc.Bookmarks.Item("\endofdoc").Range.InsertParagraphAfter()

            '---------Chapter 15.6.5 membrane and bending str in transv. section (Figure 15.6-4)---------------------------------------
            oTable = oDoc.Tables.Add(oDoc.Bookmarks.Item("\endofdoc").Range, 7, 3)
            oTable.Range.ParagraphFormat.SpaceAfter = 1
            oTable.Range.Font.Size = font_sizze
            oTable.Range.Font.Bold = CInt(False)
            oTable.Rows.Item(1).Range.Font.Bold = CInt(True)
            oTable.Rows.Item(1).Range.Font.Size = font_sizze + 2
            row = 1
            oTable.Cell(row, 1).Range.Text = "Stress transverse section 11.5.3"

            row += 1
            oTable.Cell(row, 1).Range.Text = "σ Membrane short side"
            oTable.Cell(row, 2).Range.Text = TextBox124.Text
            oTable.Cell(row, 3).Range.Text = "[N/mm2]"
            row += 1
            oTable.Cell(row, 1).Range.Text = "σ Membrane long side"
            oTable.Cell(row, 2).Range.Text = TextBox125.Text
            oTable.Cell(row, 3).Range.Text = "[N/mm2]"
            row += 1
            oTable.Cell(row, 1).Range.Text = "σ Bend A"
            oTable.Cell(row, 2).Range.Text = TextBox155.Text
            oTable.Cell(row, 3).Range.Text = "[N/mm2]"
            row += 1
            oTable.Cell(row, 1).Range.Text = "σ Bend B"
            oTable.Cell(row, 2).Range.Text = TextBox156.Text
            oTable.Cell(row, 3).Range.Text = "[N/mm2]"
            row += 1
            oTable.Cell(row, 1).Range.Text = "σ Bend C"
            oTable.Cell(row, 2).Range.Text = TextBox157.Text
            oTable.Cell(row, 3).Range.Text = "[N/mm2]"
            row += 1
            oTable.Cell(row, 1).Range.Text = "σ Bend D"
            oTable.Cell(row, 2).Range.Text = TextBox158.Text
            oTable.Cell(row, 3).Range.Text = "[N/mm2]"

            oTable.Columns(1).Width = oWord.InchesToPoints(2.91)   'Change width of columns 1 & 2.
            oTable.Columns(2).Width = oWord.InchesToPoints(1.5)
            oTable.Columns(3).Width = oWord.InchesToPoints(1.3)
            oDoc.Bookmarks.Item("\endofdoc").Range.InsertParagraphAfter()


            '---------Chapter 15.6.2.4 Reinforced vessels (C1 reinforment rib)---------------------------------------
            oTable = oDoc.Tables.Add(oDoc.Bookmarks.Item("\endofdoc").Range, 5, 3)
            oTable.Range.ParagraphFormat.SpaceAfter = 1
            oTable.Range.Font.Size = font_sizze
            oTable.Range.Font.Bold = CInt(False)
            oTable.Rows.Item(1).Range.Font.Bold = CInt(True)
            oTable.Rows.Item(1).Range.Font.Size = font_sizze + 2
            row = 1
            oTable.Cell(row, 1).Range.Text = "τ in Reinforcement web 15.6.2.4"
            row += 1
            oTable.Cell(row, 1).Range.Text = Label362.Text
            oTable.Cell(row, 2).Range.Text = TextBox116.Text
            oTable.Cell(row, 3).Range.Text = "[kN]"
            row += 1
            oTable.Cell(row, 1).Range.Text = Label365.Text
            oTable.Cell(row, 2).Range.Text = TextBox117.Text
            oTable.Cell(row, 3).Range.Text = "[N/mm2]"
            row += 1
            oTable.Cell(row, 1).Range.Text = Label393.Text
            oTable.Cell(row, 2).Range.Text = TextBox126.Text
            oTable.Cell(row, 3).Range.Text = "[N/mm2]"
            row += 1

            oTable.Columns(1).Width = oWord.InchesToPoints(2.91)   'Change width of columns 1 & 2.
            oTable.Columns(2).Width = oWord.InchesToPoints(1.5)
            oTable.Columns(3).Width = oWord.InchesToPoints(1.3)
            oDoc.Bookmarks.Item("\endofdoc").Range.InsertParagraphAfter()

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
    '7.6.4 Conical shells 
    Private Sub Calc_conical_shell764()
        Dim α As Double 'Half apex cone
        Dim De, Di, ea, z_joint, e_con, e_cone As Double
        Dim pmaxx As Double 'max pressure
        Dim Dm As Double
        If (ComboBox2.SelectedIndex > -1) Then          'Prevent exceptions
            Double.TryParse(joint_eff(ComboBox2.SelectedIndex), z_joint)      'Joint efficiency
        End If

        De = NumericUpDown15.Value  'OD
        ea = NumericUpDown16.Value  'Wall thicknes
        Di = De - 2 * ea            'ID
        α = NumericUpDown13.Value / 180 * PI        'Half apex in radials
        Dm = (De + Di) / 2                          'Average diameter
        e_cone = NumericUpDown42.Value              'Cone wall

        '----------- cone wall thickness ----------
        e_con = _P * Di / (2 * _fs * z_joint - _P)  'equation (7.6-2) Required wall thickness
        e_con *= 1 / Cos(α)

        '---------- max pressure ---------------
        pmaxx = 2 * _fs * z_joint * e_cone * Cos(α) / Dm   'Max pressure equation (7.6-4) 

        '--------- present results--------
        TextBox145.Text = Round(e_con, 2).ToString   'required cone wall [mm]
        TextBox161.Text = (pmaxx * 10).ToString("0.00") '[MPa]-->[Bar]

        '---------- Check-----
        TextBox161.BackColor = CType(IIf(pmaxx < _P, Color.Red, Color.LightGreen), Color)
    End Sub

    '7.6.6 Junction between the large end of a cone and a cylinder without a knuckle
    Private Sub Calc_Junction766()
        Dim α As Double     'is Half apex cone
        Dim β As Double     'is a factor defined in 7.6.6;
        Dim Dc As Double    'diameter large end cone
        Dim ej As Double    'is a required or analysis thickness at a junction at the large end of a cone
        Dim ej1 As Double

        Dc = NumericUpDown15.Value              'OD cone large end
        α = NumericUpDown13.Value / 180 * PI    'Half apex in radials

        ej = 40  'Initial thickness, Now iterate

        For i = 1 To 1000

            '----------- factor β ---------------------
            β = 1 / 3 * Sqrt(Dc / ej)                '(7.6-11)
            β *= Tan(α) / (1 + 1 / Sqrt(Cos(α)))
            β -= 0.15

            '----------- factor ej ---------------------
            ej1 = _P * Dc * β / (2 * _fs)            '(7.6-12)

            If ej < ej1 Then
                ej *= 1.03
            Else
                ej *= 0.97
            End If
            If Abs(ej - ej1) < 0.01 Then
                i = 1000
                TextBox18.BackColor = Color.LightGreen
            Else
                TextBox18.BackColor = Color.Red
            End If
        Next

        '--------- present results--------
        TextBox236.Text = β.ToString("0.0")     'Factor '(7.6-11)
        TextBox18.Text = ej.ToString("0.000")    'required cone wall [mm]

        '---------- Check-----
        'TextBox161.BackColor = CType(IIf(pmaxx < _P, Color.Red, Color.LightGreen), Color)
    End Sub

    Private Sub PictureBox5_Click(sender As Object, e As EventArgs) Handles PictureBox5.Click
        Form3.Show()
    End Sub

    Private Sub PictureBox17_Click(sender As Object, e As EventArgs) Handles PictureBox17.Click
        Form5.Show()
    End Sub

    Private Sub PictureBox16_Click(sender As Object, e As EventArgs) Handles PictureBox16.Click
        Form4.Show()
    End Sub

    Private Sub Button14_Click(sender As Object, e As EventArgs) Handles Button14.Click, TabPage12.Enter, NumericUpDown48.ValueChanged, NumericUpDown46.ValueChanged, NumericUpDown45.ValueChanged, NumericUpDown44.ValueChanged, NumericUpDown43.ValueChanged
        Calc_Cylinder_vacuum_852()
    End Sub
    Private Sub Calc_Cylinder_vacuum_852()
        'Vacumm cylindrical shell NO stiffeners
        Dim σe As Double
        Dim α As Double 'Half apex cone
        Dim De, Lcyl, h, Lcon, S As Double
        Dim Tolerance As Double
        Dim Pr As Double    'calculated lower bound collapse pressure 
        Dim Py As Double    'pressure at which mean circumferential stress yields
        Dim Pm As Double    'theoretical elastic instability pressure for collapse of a perfect cylindrical
        Dim L As Double     'unsupported length of the shell
        Dim R_ As Double 'mean radius of a cylindrical or spherical shell
        Dim ε As Double     'mean elastic circumferential strain at collapse, see 8.5.2.2
        Dim ncyl As Double  'number of circumferential waves for an unstiffened part of a cylinder, see 8.5.2.2; 
        Dim Z As Double     'Formula (8.5.2-7) 
        Dim ea As Double    'shell wall thickness
        Dim x As Double     'Figure 8.5-5 — Values of Pt/PP versus Pm/PP 
        Dim PrPy As Double  'Figure 8.5-5 — Values of Pt/PP versus Pm/PP 

        '--- get data ----
        De = NumericUpDown43.Value      'OD shell
        Lcyl = NumericUpDown44.Value    'Cylinder length
        h = NumericUpDown45.Value       'Dished head height
        Lcon = NumericUpDown46.Value    'Cone length
        ea = NumericUpDown48.Value      'Shell wall thickness
        α = NumericUpDown57.Value       'Half apex cone
        R_ = De / 2                'Radius shell

        '---- material --------
        S = 1.5         'Safety factor (8.4.4-1) 

        '8.4.3 For shells made in austenitic steel, the nominal elastic limit shall be given by: 
        Double.TryParse(TextBox133.Text, σe)
        If String.Equals(TextBox182.Text, "ss") Then
            σe /= 1.25
        Else
            σe /= 1.0
        End If


        '---- calculated lower bound collapse pressure obtained from Figure 8.5-5 ----------
        If α >= 30 Then
            L = Lcyl + 0.4 * h    '(8.5.2) Unsupported Length
        Else
            L = Lcyl + 0.4 * h + Lcon   '(8.5.3) Unsupported Length
        End If

        '--- pressure at which mean circumferential stress yields
        Py = σe * ea / R_            '(8.5.2-4) 

        Z = PI * R_ / L      '(8.5.2-7) 

        '---------- Find the smallest Pm ----------
        Dim Pm_small As Double = 9999

        For i = 2 To 20
            '--- calculate ε ----
            ε = Calc_ε(i, Z, R_, ea, _ν) '(8.5.2-6) 
            '--- theoretical elastic instability pressure for collapse of a perfect cylindrical
            Pm = _E * ea * ε / R_         '(8.5.2-5)
            If Pm < Pm_small Then
                Pm_small = Pm
                ncyl = i
            End If
        Next

        '--- now return to the smalles found case ----
        ε = Calc_ε(ncyl, Z, R_, ea, _ν) '(8.5.2-6) 
        '--- theoretical elastic instability pressure for collapse of a perfect cylindrical
        Pm = _E * ea * ε / R_         '(8.5.2-5)

        '---------------------------
        x = Pm / Py
        PrPy = -0.0016 * x ^ 4 + 0.031 * x ^ 3 - 0.2225 * x ^ 2 + 0.7227 * x - 0.0288

        Pr = PrPy * Py  'Calculated lower bound collapse pressure obtained from Figure 8.5-5

        '----------- Circularity tolerance  ----------
        Tolerance = 0.005 * Pr / (_P * S) * 100  '[%](8.5.1-1) 

        '--------- present results--------
        TextBox162.Text = (_P * 10).ToString("0.00")    '[MPa]-->[Bar]
        TextBox170.Text = σe.ToString("0.0")            '[N/mm]
        TextBox171.Text = S.ToString("0.0")             '[-]
        TextBox165.Text = Tolerance.ToString("0.00")    '[-]
        TextBox166.Text = (Py * 10).ToString("0.00")    '[MPa]-->[Bar]
        TextBox167.Text = L.ToString("0")               '[mm]
        TextBox172.Text = (Pm * 10).ToString("0.00")    '[MPa]-->[Bar]
        TextBox173.Text = Z.ToString("0.0")             '[-]
        TextBox174.Text = ε.ToString("0.00000")         '[-]
        TextBox175.Text = _ν.ToString("0.0")            '[-]
        TextBox177.Text = _E.ToString("0")     '[-]
        TextBox179.Text = ncyl.ToString("0")            '[-]
        TextBox176.Text = x.ToString("0.0")             '[-]Pm/Py
        TextBox180.Text = PrPy.ToString("0.00")         '[-]Pr/Py
        TextBox181.Text = (Pr * 10).ToString("0.00")    '[MPa]-->[Bar]

        '---------- Check-----
        TextBox166.BackColor = CType(IIf(Py < _P, Color.Red, Color.LightGreen), Color)
        TextBox172.BackColor = CType(IIf(Pm < _P, Color.Red, Color.LightGreen), Color)
        TextBox181.BackColor = CType(IIf(Pr / S < _P, Color.Red, Color.LightGreen), Color)
    End Sub
    Private Function Calc_ε(ncyl As Double, Z As Double, R_ As Double, ea As Double, ν As Double) As Double
        'Chapter External pressure 8.5
        Dim ε As Double
        ε = (ncyl ^ 2 - 1 + Z ^ 2) ^ 2      'Formula(8.5.2-6) 
        ε *= ea ^ 2 / (12 * R_ ^ 2 * (1 - ν ^ 2))
        ε += 1 / (ncyl ^ 2 / Z ^ 2 + 1) ^ 2
        ε *= 1 / (ncyl ^ 2 - 1 + Z ^ 2 / 2)
        Return (ε)
    End Function
    Private Sub Calc_Cylinder_vacuum_853()
        'Vacumm cylindrical shell WITH stiffeners
        Dim σe As Double
        Dim De, Lcyl, h, S As Double
        Dim Tolerance As Double
        Dim Pr As Double    'calculated lower bound collapse pressure 
        Dim Py As Double    'pressure at which mean circumferential stress yields
        Dim L_ As Double    'unsupported length of the shell
        Dim L1s As Double   'Light stiffener distance to weld
        Dim L2s As Double   'Light stiffener distance to next light stiffener
        Dim Ls, Lh As Double  'unsupported length of the shell
        Dim R_ As Double 'mean radius of a cylindrical or spherical shell
        Dim Rs As Double    'radius of the centroid of the stiffener cross-section
        Dim G As Double     'parameter in the interstiffener collapse calculation, see Equation (8.5.3-22); 
        Dim γ As Double     'parameter design of stiffeners, see Equations (8.5.3-16)
        Dim δ As Double     'parameter design of stiffeners, see Equations (8.5.3-19) and (8.5.3-20);
        Dim N_ As Double    'parameter interstiffener collapse calc., see Equation (8.5.3-21) and Table 8.5-2; 
        Dim wi As Double    'total width of stiffener i in contact with the shell, see equation (8.5.3-39) and (see Figures 8.5-14 to 8.5-17)
        Dim w As Double     'flange width as shown in Figure 8.5-9 b) 
        Dim Am As Double    'is the modified area of a stiffener, see Equation (8.5.3-17); 
        Dim B As Double     'parameter in the interstiffener collapse calculation, see Equation (8.5.3-18); 
        Dim Z As Double     'Formula (8.5.2-7)
        Dim ε As Double     'mean elastic circumferential strain at collapse, see 8.5.2.2
        Dim Pm As Double    'theoretical elastic instability pressure for collapse of a perfect cylindrical
        Dim ncyl As Double  'number of circumferential waves for an unstiffened part of a cylinder, see 8.5.2.2; 
        Dim x As Double     'Figure 8.5-5 — Values of Pt/PP versus Pm/PP 
        Dim PrPy As Double  'Figure 8.5-5 — Values of Pt/PP versus Pm/PP 

        Dim As_ As Double    'the cross-sectional area of stiffener; 
        Dim stif_h As Double 'Stiffener height
        Dim stif_w As Double 'Stiffener width
        Dim ea As Double     'shell wall thickness

        '--- get data ----
        De = NumericUpDown51.Value      'OD shell
        Lcyl = NumericUpDown52.Value    'Cylinder length
        h = NumericUpDown50.Value       'Dished head height
        L1s = NumericUpDown49.Value     'Light stiffener to weld
        L2s = NumericUpDown53.Value     'Light stiffener-stiffener distance
        ea = NumericUpDown47.Value      'Shell wall thickness
        stif_h = NumericUpDown55.Value  'Stiffener height
        stif_w = NumericUpDown54.Value  'Stiffener width
        R_ = De / 2                     'Radius cylindrical shell


        '---- material --------
        S = 1.5         'Safety factor (8.4.4-1) 
        '8.4.3 For shells made in austenitic steel, the nominal elastic limit shall be given by: 

        Double.TryParse(TextBox133.Text, σe)
        If String.Equals(TextBox182.Text, "ss") Then
            σe /= 1.25
        Else
            σe /= 1.0
        End If

        '--- pressure at which mean circumferential stress yields

        As_ = stif_h * stif_w   'page 59
        Rs = R_ + (stif_h / 2)  'is the radius of the centroid of the stiffener cross-section

        'total width of stiffener i in contact with the shell, see equation (8.5.3-39) and (see Figures 8.5-14 to 8.5-17)
        wi = stif_w
        'Where flanges act as heavy stiffeners, the shaded area shall be determined as shown in Figure 8.5-9 a).
        'Point 'A’ shall be positioned as shown in Figure 8.5-9 b) and w determined. 
        w = stif_w

        '---- Figure 8.5-6 — Cylinder with light stiffeners -------
        L_ = L2s - stif_w                   'see Figure 8.5-8 
        Ls = (L2s + L2s) / 2                '(8.5.3-7)
        Lh = Lcyl + (0.4 * h) * 2           '(8.5.2-10) 
        δ = 1.28 / (Sqrt(R_ * ea))          '(8.5.3-20) 

        If L_ > 3 * Sqrt(R_ * ea) Then      '(8.5.3-22) 
            G = 0
        Else
            G = (Sinh(δ * L_ / 2) * Cos(δ * L_ / 2))
            G += (Cosh(δ * L_ / 2) * Sin(δ * L_ / 2))
            G *= 2
            G /= (Sinh(δ * L_) + Sin(δ * L_))
        End If

        N_ = Cosh(δ * L_) - Cos(δ * L_)       '(8.5.3-21) 
        N_ /= (Sinh(δ * L_) + Sin(δ * L_))

        B = 2 * ea * N_ / (δ * (Am + w * ea))       '(8.5.3-18) 

        Am = (R_ ^ 2 / Rs ^ 2) * As_             '(8.5.3-17) 

        γ = Am * (1 - _ν / 2)                       '(8.5.3-16)
        γ /= (Am + wi * ea) * (1 + B)

        Py = σe * ea / (R_ * (1 - γ * G))           '(8.5.3-15) 

        '----- calculate Pm according 8.5.2.2 ------
        Z = PI * R_ / L_                         '(8.5.2-7) 

        '---------- Find the smallest Pm ----------
        Dim Pm_small As Double = 9999

        For i = 2 To 20
            '--- calculate ε ----
            ε = Calc_ε(i, Z, R_, ea, _ν)            '(8.5.2-6) 
            '--- theoretical elastic instability pressure for collapse of a perfect cylindrical
            Pm = _E * ea * ε / R_                   '(8.5.2-5)
            If Pm < Pm_small Then
                Pm_small = Pm
                ncyl = i
            End If
        Next

        '--- now return to the smalles found case ----
        ε = Calc_ε(ncyl, Z, R_, ea, _ν) '(8.5.2-6) 
        '--- theoretical elastic instability pressure for collapse of a perfect cylindrical
        Pm = _E * ea * ε / R_         '(8.5.2-5)

        '-----------Figure 8.5-5 — Values of Pr/Py versus Pm/Py ----------------
        x = Pm / Py

        PrPy = -0.0016 * x ^ 4 + 0.031 * x ^ 3 - 0.2225 * x ^ 2 + 0.7227 * x - 0.0288

        Pr = PrPy * Py  'Calculated lower bound collapse pressure obtained from Figure 8.5-5

        '----------- Circularity tolerance  ----------
        Tolerance = 0.005 * Pr / (_P * S) * 100  '[%](8.5.1-1) 

        '--------- present results--------
        TextBox183.Text = (_P * 10).ToString("0.00")    '[MPa]-->[Bar]
        TextBox187.Text = σe.ToString("0.0")            '[N/mm]
        TextBox188.Text = L_.ToString("0")              '[mm] unsupported length
        'TextBox184.Text = Tolerance.ToString("0.00")   '[-]
        TextBox185.Text = _E.ToString("0")              '[-]
        TextBox186.Text = _ν.ToString("0.0")            '[-]
        TextBox182.Text = S.ToString("0.0")             '[-]
        TextBox189.Text = N_.ToString("0.00")           '[-]
        TextBox190.Text = As_.ToString                  '[mm2]
        TextBox191.Text = Am.ToString("0")              '[mm2]modified area of a stiffener
        TextBox192.Text = G.ToString("0.000")           '[-]
        TextBox193.Text = B.ToString("0.0")             '[-]
        TextBox194.Text = γ.ToString("0.00")            '[-]
        TextBox195.Text = δ.ToString("0.000")           '[-]'(8.5.3-20) 
        TextBox196.Text = Rs.ToString                   '[-]
        TextBox197.Text = (Py * 10).ToString("0.00")    '[bar]
        TextBox198.Text = Ls.ToString("0")              '[mm]
        TextBox199.Text = Lh.ToString("0")              '[mm]

        TextBox201.Text = (Pr * 10).ToString("0.00")    '[MPa]-->[Bar]
        TextBox202.Text = PrPy.ToString("0.00")         '[-]Pr/Py
        TextBox203.Text = x.ToString("0.0")             '[-]Pm/Py
        TextBox204.Text = ncyl.ToString("0")            '[-]
        TextBox205.Text = S.ToString("0.0")             '[-]
        TextBox206.Text = ε.ToString("0.00000")         '[-]
        TextBox207.Text = Z.ToString("0.0")             '[-]
        TextBox208.Text = (Pm * 10).ToString("0.00")    '[MPa]-->[Bar]

        '---------- Check-----
        TextBox197.BackColor = CType(IIf(Py < _P, Color.Red, Color.LightGreen), Color)
        TextBox208.BackColor = CType(IIf(Pm < _P, Color.Red, Color.LightGreen), Color)
        TextBox201.BackColor = CType(IIf(Pr / S < _P, Color.Red, Color.LightGreen), Color)
    End Sub

    Private Sub Button15_Click(sender As Object, e As EventArgs) Handles Button15.Click, TabPage13.Enter, NumericUpDown52.ValueChanged, NumericUpDown51.ValueChanged, NumericUpDown50.ValueChanged, NumericUpDown49.ValueChanged, NumericUpDown47.ValueChanged, NumericUpDown53.ValueChanged
        Calc_Cylinder_vacuum_853()
    End Sub

    Private Sub Button16_Click(sender As Object, e As EventArgs) Handles Button16.Click, TabPage14.Enter, NumericUpDown55.ValueChanged, NumericUpDown54.ValueChanged
        '8.5.3.6 Design of light stiffeners 
        Dim pg, pg_small As Double
        Dim n_low As Integer

        Calc_Cylinder_vacuum_853()
        'Iterate and find the smallest result
        pg_small = 9999                 'Initial value
        For i = 2 To 6                  '(8.5.3-41)
            Calc_Light_Stiffeners_8536(i)
            Double.TryParse(TextBox223.Text, pg)
            If pg < pg_small Then
                pg_small = pg
                n_low = i
            End If
        Next
        Calc_Light_Stiffeners_8536(n_low)
    End Sub
    Private Sub Calc_Light_Stiffeners_8536(n_wave As Integer)
        Dim ea As Double    'shell wall thickness
        Dim De As Double    'Outside diameter shell
        Dim Ls As Double    'unsupported length of the shell (8.5.3-7)
        Dim Lh As Double    'Is the distance between heavy stiffeners, see Table (8.5.3-10) 
        Dim Le As Double    'is the effective length of shell acting with a light stiffener, see Equation (8.5.3-34)
        Dim h As Double     'Dished head [mm]
        Dim R_ As Double    'radius of a cylindrical or spherical shell
        Dim Rs As Double    'radius of the centroid of the stiffener cross-section
        Dim Pg As Double    'theoretical elastic instability pressure of a stiffener on a cylinder
        Dim β As Double     'Formula (8.5.3-25) 
        Dim Ie As Double    'second moment of area of the composite cross-section 
        Dim Is_ As Double   'second moment of area of the stiffener cross-section 
        Dim As_ As Double   'cross-sectional area of stiffener
        Dim Ae As Double    'cross-sectional area of stiffener and effective length of shell, see Equation (8.5.3-30)
        Dim λ As Double     'parameter depending on stiffener location, see Equations (8.5.3-28) and (8.5.3-29)
        Dim rh, rw As Double    'Rib dimensions
        Dim x As Double     'Formula  (8.5.3-35) 
        Dim y1, y2, y3 As Double 'Formula (8.5.3.6.3)
        Dim Lcyl As Double  'Cylinder length
        Dim L1s As Double   'Distance light stiff - weld
        Dim L2s As Double   'Distance light stiff - light stiff
        Dim u As Double     'Formula '(8.5.3-36)
        Dim Xe As Double

        TextBox232.Clear()

        '--- get data ----
        Lcyl = NumericUpDown52.Value    'Cylinder length
        h = NumericUpDown50.Value       'Dished head length
        L1s = NumericUpDown49.Value     'Distance rib to weld
        L2s = NumericUpDown53.Value     'Distance rib to rib
        De = NumericUpDown51.Value      'OD shell
        rh = NumericUpDown55.Value      'Rib height
        rw = NumericUpDown54.Value      'Rib width
        ea = NumericUpDown47.Value      'shell wall thickness

        R_ = De / 2
        Double.TryParse(TextBox196.Text, Rs)
        Double.TryParse(TextBox190.Text, As_)
        Double.TryParse(TextBox198.Text, Ls)    '(8.5.3-7)
        Double.TryParse(TextBox199.Text, Lh)    '(8.5.3-10)

        '---- Determine u --------
        u = (Ls / R_) / Sqrt(ea / R_)    '(8.5.3-36) 

        Select Case (True)
            Case (u <= 1)
                y1 = u / (1 / 1.098 + 0.03 * u ^ 3)
                y2 = 0
                y3 = 0.6 * (1 - 0.27 * u) * u ^ 2
            Case (u > 1) And (u < 2.2)
                y1 = u / (1 / 1.098 + 0.03 * u ^ 3)
                y2 = u - 1
                y3 = 0.6 * (1 - 0.27 * u) * u ^ 2
            Case (u >= 2.2) And (u <= 2.9)
                y1 = u / (1 / 1.098 + 0.03 * u ^ 3)
                y2 = 1.2
                y3 = 0.6 * (1 - 0.27 * u) * u ^ 2
            Case (u > 2.9) And (u < 4.1)
                y1 = 1.2 + 1.642 / u
                y2 = 1.2
                y3 = 0.75 + (1 / u)
            Case (u >= 4.1) And (u < 5)
                y1 = 1.556 + 0.183 / u
                y2 = 1.2
                y3 = 0.75 + (1 / u)
            Case (u >= 5)
                y1 = 1.556 + 0.183 / u
                y2 = 1.2
                y3 = 0.65 + (1.5 / u)
        End Select

        '------Is_ is the second moment of area of the stiffener cross-section 
        'about the axis passing through the centroid Parallel to the cylinder axis;
        ' for stability calculation 
        Is_ = (rh * rw ^ 3) / 12        '[mm4]

        '----8.5.3.6.3 Determination of Le ----
        x = n_wave ^ 2 * (ea / R_)          '(8.5.3-35) 

        '----  (8.5.3-34) ------------
        If ((ea / R_) <= 0.0346 And (ea / R_) >= 0.001095) Then
            Le = y1 * Sqrt(ea / R_)
            Le /= Sqrt(y3 * x + Sqrt(1 + y2 * x ^ 2))
            Le *= R_
        Else
            Le = y1 * Sqrt(0.0346)
            Le /= Sqrt(y3 * x + Sqrt(1 + y2 * x ^ 2))
            Le *= R_
        End If

        '------ External stiffeners ---
        λ = -1                          '(8.5.3-29) External stiff
        Ae = As_ + ea * Le              '(8.5.3-30) [mm2]

        β = n_wave ^ 2 - 1 + (0.5 * (PI * R_ / Lh) ^ 2)   '(8.5.3-25)
        β *= (n_wave ^ 2 * (Lh / (PI * R_)) ^ 2 + 1) ^ 2
        β ^= -1

        Xe = 0.5 * ea ^ 2 * Le          '(8.5.3-27)
        Xe += As_ * (ea / 2 + λ * (R_ - Rs))
        Xe /= Ae

        Ie = ea ^ 3 * Le / 3            '(8.5.3-26) 
        Ie += Is_
        Ie += As_ * (ea / 2 * λ * (R_ - Rs)) ^ 2
        Ie -= Ae * Xe ^ 2

        Pg = _E * ea * β / R_       '(8.5.3-24) 
        Pg += (n_wave ^ 2 - 1) * _E * Ie / (R_ ^ 3 * Ls)

        '======================================================
        '----- 8.5.3.6.4 Maximum stresses in the stiffeners ---
        Dim σes As Double
        Dim S As Double = 1.5       'Equation(8.4.4-1) Safety factor
        Dim Sf As Double = 1.33     'Equation (8.5.3-33)cold formed
        Dim Pys As Double           'Equation(8.5.3-38)
        Dim dmax, dmax1, dmax2 As Double
        Dim Am As Double            'Equation (8.5.3-17);
        Dim wi As Double            'Rib width
        Dim δ As Double             'Equation (8.5.3-19);
        Dim Rf As Double            'Figures 8.5-14 to 8.5-17);
        Dim N_ As Double            'Equation (8.5.3-21) or Table 8.5-2;
        Dim σs As Double            'Max stress in stiffener (8.5.3-37)

        '------ allowed stress ------------
        Double.TryParse(TextBox133.Text, σes) '(8.4.2-2)

        '------ get data -----
        Double.TryParse(TextBox195.Text, δ)
        Double.TryParse(TextBox189.Text, N_)
        Double.TryParse(TextBox191.Text, Am)

        wi = rw         'Rib width (8.5.3-39)
        Rf = R_ + rh    'Vessel radius +rib height, is the radius to the part of the stiffener furthest from the shell

        '-------- (8.5.3-40) ---------
        dmax1 = λ * (R_ - Rf) - Xe + ea / 2
        dmax2 = Xe
        dmax = CDbl(IIf(dmax1 > dmax2, dmax1, dmax2))   'Take the biggest

        '-------- (8.5.3-38)---------
        Pys = 1 + Am / (wi * ea + 2 * N_ * ea / δ)
        Pys *= σes * ea * Rf
        Pys /= R_ ^ 2 * (1 - _ν / 2)

        '-------- (8.5.3-37) --------
        σs = S * Sf * (_P * σes / Pys)
        σs += _E * dmax * 0.005 * (n_wave ^ 2 - 1) * _P * S * Sf / (R_ * (Pg - _P * S * Sf))

        TextBox232.Text &= "ea=   " & vbTab & ea.ToString & vbCrLf
        TextBox232.Text &= "Ae=   " & vbTab & Ae.ToString & vbCrLf
        TextBox232.Text &= "Le=   " & vbTab & Le.ToString & vbCrLf
        TextBox232.Text &= "Is=   " & vbTab & Is_.ToString & vbCrLf

        TextBox232.Text &= "R_=   " & vbTab & R_.ToString & vbCrLf
        TextBox232.Text &= "Rs=   " & vbTab & Rs.ToString & vbCrLf
        TextBox232.Text &= "λ=   " & vbTab & λ.ToString & vbCrLf
        TextBox232.Text &= "Xe=   " & vbTab & Xe.ToString & vbCrLf
        '--------------
        TextBox232.Text &= "wi=   " & vbTab & wi.ToString & vbCrLf
        TextBox232.Text &= "rh=   " & vbTab & rh.ToString & vbCrLf
        TextBox232.Text &= "Rf=   " & vbTab & Rf.ToString & vbCrLf
        TextBox232.Text &= "Ie=   " & vbTab & Ie.ToString & vbCrLf
        TextBox232.Text &= "δ=   " & vbTab & δ.ToString & vbCrLf
        TextBox232.Text &= "N_=   " & vbTab & N_.ToString & vbCrLf
        TextBox232.Text &= "dmax1= " & vbTab & dmax1.ToString & vbCrLf
        TextBox232.Text &= "dmax2= " & vbTab & dmax2.ToString & vbCrLf
        TextBox232.Text &= "dmax= " & vbTab & dmax.ToString & vbCrLf
        TextBox232.Text &= "Am=   " & vbTab & Am.ToString & vbCrLf
        TextBox232.Text &= "Pys=  " & vbTab & Pys.ToString & vbCrLf

        TextBox232.Text &= "S=   " & vbTab & S.ToString & vbCrLf
        TextBox232.Text &= "Sf=   " & vbTab & Sf.ToString & vbCrLf
        TextBox232.Text &= "σes=  " & vbTab & σes.ToString & vbCrLf
        TextBox232.Text &= "σs=  " & vbTab & σs.ToString & vbCrLf

        '--------- present results--------
        TextBox200.Text = n_wave.ToString               'No of waves
        TextBox210.Text = (_P * 10).ToString("0.00")    '[MPa]-->[Bar]
        TextBox211.Text = De.ToString("0.0")            'Diameter[mm]
        TextBox215.Text = ea.ToString("0.0")            'Diameter[mm]
        TextBox216.Text = As_.ToString("0.0")
        TextBox217.Text = Rs.ToString("0.0")
        TextBox218.Text = Ie.ToString("0.0")

        TextBox212.Text = Lcyl.ToString("0")
        TextBox213.Text = h.ToString("0")
        TextBox214.Text = β.ToString("E1")
        TextBox219.Text = L1s.ToString("0")
        TextBox220.Text = L2s.ToString("0")

        TextBox221.Text = rh.ToString("0.0")
        TextBox222.Text = rw.ToString("0.0")

        TextBox223.Text = (Pg * 10).ToString("0.00")  '[Mpa]-->[bar]
        TextBox224.Text = Le.ToString("0.0")
        TextBox225.Text = u.ToString("0.00")

        TextBox226.Text = y1.ToString("0.000")
        TextBox227.Text = y2.ToString("0.000")
        TextBox228.Text = y3.ToString("0.000")
        TextBox229.Text = (ea / R_).ToString("0.0000")
        TextBox230.Text = Ls.ToString("0")
        TextBox231.Text = Lh.ToString("0")
        TextBox233.Text = Is_.ToString("0") '[mm]
        TextBox234.Text = Xe.ToString("0.00")
        TextBox235.Text = σs.ToString("0.0")

        TextBox184.Text = ("0.5") 'Tolerance 0.5%

        '---------- Check-----
        TextBox223.BackColor = CType(IIf(Pg < _P, Color.Red, Color.LightGreen), Color)
        TextBox235.BackColor = CType(IIf(σes < σs Or σs <= 0, Color.Red, Color.LightGreen), Color) '(8.5.3-41) 

    End Sub

    Private Sub Button17_Click(sender As Object, e As EventArgs) Handles Button17.Click
        Dim oWord As Word.Application

        Dim oDoc As Word.Document
        Dim oTable As Word.Table
        Dim oPara1, oPara2 As Word.Paragraph
        Dim row, font_sizze As Integer
        Dim ufilename As String
        ufilename = "PV_calc_" & TextBox66.Text & "_" & TextBox69.Text & "_" & TextBox70.Text & DateTime.Now.ToString("yyyy_MM_dd") & ".docx"

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

            oTable.Columns(1).Width = oWord.InchesToPoints(2.91)   'Change width of columns 1 & 2.
            oTable.Columns(2).Width = oWord.InchesToPoints(4)

            oTable.Rows.Item(1).Range.Font.Bold = CInt(True)
            oDoc.Bookmarks.Item("\endofdoc").Range.InsertParagraphAfter()

            '---------------Cylinder shells under External pressure 8.4---------------------
            'Insert a 18 (row) x 3 table (column), fill it with data and change the column widths.
            oTable = oDoc.Tables.Add(oDoc.Bookmarks.Item("\endofdoc").Range, 9, 3)
            oTable.Range.ParagraphFormat.SpaceAfter = 1
            oTable.Range.Font.Size = font_sizze
            oTable.Range.Font.Bold = CInt(False)
            oTable.Rows.Item(1).Range.Font.Bold = CInt(True)
            oTable.Rows.Item(1).Range.Font.Size = font_sizze + 2
            row = 1
            oTable.Cell(row, 1).Range.Text = "Cylinder shells under External pressure 8.4"
            row += 1
            oTable.Cell(row, 1).Range.Text = Label459.Text
            oTable.Cell(row, 2).Range.Text = TextBox162.Text
            oTable.Cell(row, 3).Range.Text = "[barg]"
            row += 1
            oTable.Cell(row, 1).Range.Text = Label467.Text
            oTable.Cell(row, 2).Range.Text = NumericUpDown43.Value.ToString
            oTable.Cell(row, 3).Range.Text = "[mm]"
            row += 1
            oTable.Cell(row, 1).Range.Text = Label469.Text
            oTable.Cell(row, 2).Range.Text = NumericUpDown44.Value.ToString
            oTable.Cell(row, 3).Range.Text = "[mm]"
            row += 1
            oTable.Cell(row, 1).Range.Text = Label477.Text
            oTable.Cell(row, 2).Range.Text = NumericUpDown45.Value.ToString
            oTable.Cell(row, 3).Range.Text = "[mm]"
            row += 1
            oTable.Cell(row, 1).Range.Text = Label480.Text
            oTable.Cell(row, 2).Range.Text = NumericUpDown46.Value.ToString
            oTable.Cell(row, 3).Range.Text = "[mm]"
            row += 1
            oTable.Cell(row, 1).Range.Text = Label600.Text
            oTable.Cell(row, 2).Range.Text = TextBox165.Text
            oTable.Cell(row, 3).Range.Text = "[%]"
            row += 1
            oTable.Cell(row, 1).Range.Text = Label471.Text
            oTable.Cell(row, 2).Range.Text = NumericUpDown57.Value.ToString
            oTable.Cell(row, 3).Range.Text = "[deg]"
            row += 1
            oTable.Cell(row, 1).Range.Text = Label514.Text
            oTable.Cell(row, 2).Range.Text = NumericUpDown48.Value.ToString
            oTable.Cell(row, 3).Range.Text = "[-]"

            oTable.Columns(1).Width = oWord.InchesToPoints(2.91)   'Change width of columns 1 & 2.
            oTable.Columns(2).Width = oWord.InchesToPoints(1.5)
            oTable.Columns(3).Width = oWord.InchesToPoints(1.3)
            oDoc.Bookmarks.Item("\endofdoc").Range.InsertParagraphAfter()

            '---------The nominal elastic limit σ  8.4.3 (page 51)------------------
            oTable = oDoc.Tables.Add(oDoc.Bookmarks.Item("\endofdoc").Range, 4, 3)
            oTable.Range.ParagraphFormat.SpaceAfter = 1
            oTable.Range.Font.Size = font_sizze
            oTable.Range.Font.Bold = CInt(False)
            oTable.Rows.Item(1).Range.Font.Bold = CInt(True)
            oTable.Rows.Item(1).Range.Font.Size = font_sizze + 2
            row = 1
            oTable.Cell(row, 1).Range.Text = "Nominal elastic limit σ  8.4.3 "

            row += 1
            oTable.Cell(row, 1).Range.Text = Label492.Text
            oTable.Cell(row, 2).Range.Text = TextBox170.Text
            oTable.Cell(row, 3).Range.Text = "[N/mm2]"
            row += 1
            oTable.Cell(row, 1).Range.Text = Label508.Text
            oTable.Cell(row, 2).Range.Text = TextBox175.Text
            oTable.Cell(row, 3).Range.Text = "[-]"
            row += 1
            oTable.Cell(row, 1).Range.Text = Label517.Text
            oTable.Cell(row, 2).Range.Text = TextBox177.Text
            oTable.Cell(row, 3).Range.Text = "[N/mm2]"

            oTable.Columns(1).Width = oWord.InchesToPoints(2.91)   'Change width of columns 1 & 2.
            oTable.Columns(2).Width = oWord.InchesToPoints(1.5)
            oTable.Columns(3).Width = oWord.InchesToPoints(1.3)
            oDoc.Bookmarks.Item("\endofdoc").Range.InsertParagraphAfter()

            '---------Cylinder wall thickn. Ext. pressure 8.5.2.2 (NO reinforcements) (page 54)-------
            oTable = oDoc.Tables.Add(oDoc.Bookmarks.Item("\endofdoc").Range, 10, 3)
            oTable.Range.ParagraphFormat.SpaceAfter = 1
            oTable.Range.Font.Size = font_sizze
            oTable.Range.Font.Bold = CInt(False)
            oTable.Rows.Item(1).Range.Font.Bold = CInt(True)
            oTable.Rows.Item(1).Range.Font.Size = font_sizze + 2
            row = 1
            oTable.Cell(row, 1).Range.Text = "Cylinder wall thickness External pressure 8.5.2.2 (NO reinforcements)"

            row += 1
            oTable.Cell(row, 1).Range.Text = Label475.Text
            oTable.Cell(row, 2).Range.Text = TextBox167.Text
            oTable.Cell(row, 3).Range.Text = "[mm]"
            row += 1
            oTable.Cell(row, 1).Range.Text = Label473.Text
            oTable.Cell(row, 2).Range.Text = TextBox166.Text
            oTable.Cell(row, 3).Range.Text = "[bar]"
            row += 1
            oTable.Cell(row, 1).Range.Text = Label502.Text
            oTable.Cell(row, 2).Range.Text = TextBox173.Text
            oTable.Cell(row, 3).Range.Text = "[mm]"
            row += 1
            oTable.Cell(row, 1).Range.Text = Label511.Text
            oTable.Cell(row, 2).Range.Text = TextBox179.Text
            oTable.Cell(row, 3).Range.Text = "[-]"
            row += 1
            oTable.Cell(row, 1).Range.Text = Label505.Text
            oTable.Cell(row, 2).Range.Text = TextBox174.Text
            oTable.Cell(row, 3).Range.Text = "[mm]"
            row += 1
            oTable.Cell(row, 1).Range.Text = Label499.Text
            oTable.Cell(row, 2).Range.Text = TextBox172.Text
            oTable.Cell(row, 3).Range.Text = "[bar]"
            row += 1
            oTable.Cell(row, 1).Range.Text = Label520.Text
            oTable.Cell(row, 2).Range.Text = TextBox176.Text & ",  " & TextBox180.Text
            oTable.Cell(row, 3).Range.Text = "[-]"
            row += 1
            oTable.Cell(row, 1).Range.Text = Label495.Text
            oTable.Cell(row, 2).Range.Text = TextBox171.Text
            oTable.Cell(row, 3).Range.Text = "[-]"
            row += 1
            oTable.Cell(row, 1).Range.Text = Label526.Text
            oTable.Cell(row, 2).Range.Text = TextBox181.Text
            oTable.Cell(row, 3).Range.Text = "[bar]"

            oTable.Columns(1).Width = oWord.InchesToPoints(2.91)   'Change width of columns 1 & 2.
            oTable.Columns(2).Width = oWord.InchesToPoints(1.5)
            oTable.Columns(3).Width = oWord.InchesToPoints(1.3)
            oDoc.Bookmarks.Item("\endofdoc").Range.InsertParagraphAfter()


        Catch ex As Exception
            MessageBox.Show(ufilename & vbCrLf & ex.Message)  ' Show the exception's message.
        End Try
    End Sub

    Private Sub Button18_Click(sender As Object, e As EventArgs) Handles Button18.Click
        'Design with Stiffeners (Vacuum operation)
        Dim oWord As Word.Application

        Dim oDoc As Word.Document
        Dim oTable As Word.Table
        Dim oPara1, oPara2 As Word.Paragraph
        Dim row, font_sizze As Integer
        Dim ufilename As String
        ufilename = "PV_calc_" & TextBox66.Text & "_" & TextBox69.Text & "_" & TextBox70.Text & DateTime.Now.ToString("yyyy_MM_dd") & ".docx"

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

            oTable.Columns(1).Width = oWord.InchesToPoints(2.91)   'Change width of column
            oTable.Columns(2).Width = oWord.InchesToPoints(4)

            oTable.Rows.Item(1).Range.Font.Bold = CInt(True)
            oDoc.Bookmarks.Item("\endofdoc").Range.InsertParagraphAfter()

            '----------Cylinder shells under External pressure 8.4 with stiffeners---------------------
            'Insert a 18 (row) x 3 table (column), fill it with data and change the column widths.
            oTable = oDoc.Tables.Add(oDoc.Bookmarks.Item("\endofdoc").Range, 7, 3)
            oTable.Range.ParagraphFormat.SpaceAfter = 1
            oTable.Range.Font.Size = font_sizze
            oTable.Range.Font.Bold = CInt(False)
            oTable.Rows.Item(1).Range.Font.Bold = CInt(True)
            oTable.Rows.Item(1).Range.Font.Size = font_sizze + 2
            row = 1
            oTable.Cell(row, 1).Range.Text = "Stiffened Cylinder under External pressure 8.4"
            row += 1
            oTable.Cell(row, 1).Range.Text = Label533.Text
            oTable.Cell(row, 2).Range.Text = TextBox183.Text
            oTable.Cell(row, 3).Range.Text = "[barg]"
            row += 1
            oTable.Cell(row, 1).Range.Text = Label538.Text
            oTable.Cell(row, 2).Range.Text = NumericUpDown51.Value.ToString
            oTable.Cell(row, 3).Range.Text = "[mm]"
            row += 1
            oTable.Cell(row, 1).Range.Text = Label540.Text
            oTable.Cell(row, 2).Range.Text = NumericUpDown52.Value.ToString
            oTable.Cell(row, 3).Range.Text = "[mm]"
            row += 1
            oTable.Cell(row, 1).Range.Text = Label530.Text
            oTable.Cell(row, 2).Range.Text = NumericUpDown50.Value.ToString
            oTable.Cell(row, 3).Range.Text = "[mm]"
            row += 1
            oTable.Cell(row, 1).Range.Text = Label542.Text
            oTable.Cell(row, 2).Range.Text = TextBox184.Text
            oTable.Cell(row, 3).Range.Text = "[%]"
            row += 1
            oTable.Cell(row, 1).Range.Text = Label521.Text
            oTable.Cell(row, 2).Range.Text = NumericUpDown47.Value.ToString
            oTable.Cell(row, 3).Range.Text = "[mm]"

            oTable.Columns(1).Width = oWord.InchesToPoints(2.91)   'Change width of columns 1 & 2.
            oTable.Columns(2).Width = oWord.InchesToPoints(1.5)
            oTable.Columns(3).Width = oWord.InchesToPoints(1.3)
            oDoc.Bookmarks.Item("\endofdoc").Range.InsertParagraphAfter()

            '---------Fig 8.5-6 - Cilinder with light stiffeners -----------------
            oTable = oDoc.Tables.Add(oDoc.Bookmarks.Item("\endofdoc").Range, 6, 3)
            oTable.Range.ParagraphFormat.SpaceAfter = 1
            oTable.Range.Font.Size = font_sizze
            oTable.Range.Font.Bold = CInt(False)
            oTable.Rows.Item(1).Range.Font.Bold = CInt(True)
            oTable.Rows.Item(1).Range.Font.Size = font_sizze + 2
            row = 1
            oTable.Cell(row, 1).Range.Text = "Cilinder with light stiffeners"

            row += 1
            oTable.Cell(row, 1).Range.Text = Label527.Text
            oTable.Cell(row, 2).Range.Text = NumericUpDown49.Value.ToString
            oTable.Cell(row, 3).Range.Text = "[mm]"
            row += 1
            oTable.Cell(row, 1).Range.Text = Label586.Text
            oTable.Cell(row, 2).Range.Text = NumericUpDown53.Value.ToString
            oTable.Cell(row, 3).Range.Text = "[mm]"
            row += 1
            oTable.Cell(row, 1).Range.Text = Label587.Text
            oTable.Cell(row, 2).Range.Text = TextBox221.Text & "x" & TextBox222.Text
            oTable.Cell(row, 3).Range.Text = "[mm]"
            row += 1
            oTable.Cell(row, 1).Range.Text = Label591.Text
            oTable.Cell(row, 2).Range.Text = TextBox190.Text
            oTable.Cell(row, 3).Range.Text = "[mm2]"
            row += 1
            oTable.Cell(row, 1).Range.Text = Label594.Text
            oTable.Cell(row, 2).Range.Text = TextBox196.Text
            oTable.Cell(row, 3).Range.Text = "[mm]"

            oTable.Columns(1).Width = oWord.InchesToPoints(2.91)   'Change width of columns 1 & 2.
            oTable.Columns(2).Width = oWord.InchesToPoints(1.5)
            oTable.Columns(3).Width = oWord.InchesToPoints(1.3)
            oDoc.Bookmarks.Item("\endofdoc").Range.InsertParagraphAfter()

            '---------Nominal elastic limit σ  8.4.3-------
            oTable = oDoc.Tables.Add(oDoc.Bookmarks.Item("\endofdoc").Range, 4, 3)
            oTable.Range.ParagraphFormat.SpaceAfter = 1
            oTable.Range.Font.Size = font_sizze
            oTable.Range.Font.Bold = CInt(False)
            oTable.Rows.Item(1).Range.Font.Bold = CInt(True)
            oTable.Rows.Item(1).Range.Font.Size = font_sizze + 2
            row = 1
            oTable.Cell(row, 1).Range.Text = "Nominal elastic limit σ  8.4.3"

            row += 1
            oTable.Cell(row, 1).Range.Text = Label551.Text
            oTable.Cell(row, 2).Range.Text = TextBox187.Text
            oTable.Cell(row, 3).Range.Text = "[N/mm2]"
            row += 1
            oTable.Cell(row, 1).Range.Text = Label548.Text
            oTable.Cell(row, 2).Range.Text = TextBox186.Text
            oTable.Cell(row, 3).Range.Text = "[-]"
            row += 1
            oTable.Cell(row, 1).Range.Text = Label545.Text
            oTable.Cell(row, 2).Range.Text = TextBox185.Text
            oTable.Cell(row, 3).Range.Text = "[N/mm2]"

            oTable.Columns(1).Width = oWord.InchesToPoints(2.91)   'Change width of columns 1 & 2.
            oTable.Columns(2).Width = oWord.InchesToPoints(1.5)
            oTable.Columns(3).Width = oWord.InchesToPoints(1.3)
            oDoc.Bookmarks.Item("\endofdoc").Range.InsertParagraphAfter()

            '--------- Table 8.5-1 Definition of cilinder length-------
            oTable = oDoc.Tables.Add(oDoc.Bookmarks.Item("\endofdoc").Range, 4, 3)
            oTable.Range.ParagraphFormat.SpaceAfter = 1
            oTable.Range.Font.Size = font_sizze
            oTable.Range.Font.Bold = CInt(False)
            oTable.Rows.Item(1).Range.Font.Bold = CInt(True)
            oTable.Rows.Item(1).Range.Font.Size = font_sizze + 2
            row = 1
            oTable.Cell(row, 1).Range.Text = "Definition of cilinder length"

            row += 1
            oTable.Cell(row, 1).Range.Text = Label554.Text
            oTable.Cell(row, 2).Range.Text = TextBox188.Text
            oTable.Cell(row, 3).Range.Text = "[mm]"
            row += 1
            oTable.Cell(row, 1).Range.Text = Label580.Text
            oTable.Cell(row, 2).Range.Text = TextBox198.Text
            oTable.Cell(row, 3).Range.Text = "[mm]"
            row += 1
            oTable.Cell(row, 1).Range.Text = Label583.Text
            oTable.Cell(row, 2).Range.Text = TextBox199.Text
            oTable.Cell(row, 3).Range.Text = "[mm]"

            oTable.Columns(1).Width = oWord.InchesToPoints(2.91)   'Change width of columns 1 & 2.
            oTable.Columns(2).Width = oWord.InchesToPoints(1.5)
            oTable.Columns(3).Width = oWord.InchesToPoints(1.3)
            oDoc.Bookmarks.Item("\endofdoc").Range.InsertParagraphAfter()

            '--------- 8.5.3.4 Interstiffener collapse-------
            oTable = oDoc.Tables.Add(oDoc.Bookmarks.Item("\endofdoc").Range, 5, 3)
            oTable.Range.ParagraphFormat.SpaceAfter = 1
            oTable.Range.Font.Size = font_sizze
            oTable.Range.Font.Bold = CInt(False)
            oTable.Rows.Item(1).Range.Font.Bold = CInt(True)
            oTable.Rows.Item(1).Range.Font.Size = font_sizze + 2
            row = 1
            oTable.Cell(row, 1).Range.Text = "8.5.3.4 Interstiffener collapse"

            row += 1
            oTable.Cell(row, 1).Range.Text = Label577.Text
            oTable.Cell(row, 2).Range.Text = TextBox197.Text
            oTable.Cell(row, 3).Range.Text = "[mm]"
            row += 1
            oTable.Cell(row, 1).Range.Text = Label620.Text
            oTable.Cell(row, 2).Range.Text = TextBox208.Text
            oTable.Cell(row, 3).Range.Text = "[bar]"
            row += 1
            oTable.Cell(row, 1).Range.Text = Label609.Text
            oTable.Cell(row, 2).Range.Text = TextBox205.Text
            oTable.Cell(row, 3).Range.Text = "[-]"
            row += 1
            oTable.Cell(row, 1).Range.Text = Label603.Text
            oTable.Cell(row, 2).Range.Text = TextBox201.Text
            oTable.Cell(row, 3).Range.Text = "[mm]"

            oTable.Columns(1).Width = oWord.InchesToPoints(2.91)   'Change width of columns 1 & 2.
            oTable.Columns(2).Width = oWord.InchesToPoints(1.5)
            oTable.Columns(3).Width = oWord.InchesToPoints(1.3)
            oDoc.Bookmarks.Item("\endofdoc").Range.InsertParagraphAfter()

        Catch ex As Exception
            MessageBox.Show(ufilename & vbCrLf & ex.Message)  ' Show the exception's message.
        End Try
    End Sub
End Class
