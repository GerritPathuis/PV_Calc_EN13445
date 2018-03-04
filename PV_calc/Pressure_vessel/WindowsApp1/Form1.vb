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
            ComboBox3.Items.Add(joint_eff(hh))
        Next hh

        '----------------- prevent out of bounds------------------
        ComboBox1.SelectedIndex = CInt(IIf(ComboBox1.Items.Count > 0, 1, -1)) 'Select ..
        ComboBox2.SelectedIndex = CInt(IIf(ComboBox2.Items.Count > 0, 0, -1)) 'Select ..
        ComboBox3.SelectedIndex = CInt(IIf(ComboBox3.Items.Count > 0, 0, -1)) 'Select ..
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
        '----------- check--------
        TextBox16.BackColor = IIf(eq_left < eq_right, Color.Red, Color.LightGreen)
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click, ComboBox1.TextChanged, CheckBox1.CheckedChanged, NumericUpDown6.ValueChanged, NumericUpDown5.ValueChanged, NumericUpDown4.ValueChanged, NumericUpDown10.ValueChanged
        Design_stress()
    End Sub
    Private Sub Design_stress()
        Dim sf As Double

        _P = NumericUpDown4.Value                       'Calculation pressure [MPa=N/mm2]
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
End Class
