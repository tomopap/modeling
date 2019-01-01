Option Explicit

'Units
Global tera As Double, giga As Double, mega As Double, meg As Double, kilo As Double, milli As Double, micro As Double, nano As Double, pico As Double, femt As Double

'Physical constants
Global kB As Double, Temp As Double, Qe As Double, VTHERMAL As Double, PI As Double
Global Boltz As Double, Charge As Double, Temperature As Double, Vkt As Double


'Matrix parameters
Global lo As Integer, up As Integer
Global continue As Integer
Global itercount As Integer
Global rpvt(1 To 100) As Integer, cpvt(1 To 100) As Integer

'Conductance minimum for Simulation
Global GMIN As Double
Global tol As Double
Dim itermax

'Simple model parameters
Global NT As Double, VTO As Double, ALPHA As Double, BETA As Double, LAMBDA As Double, NG As Double

'Diode model parameters
Dim Isgs As Double, ngs As Double, IGSMAX As Double
Dim Isgd As Double, ngd As Double, IGDMAX As Double

'Hirose model parameters
Dim VTH As Double, GAMMA As Double, AGS As Double, IMAXG0 As Double, IMAXG1 As Double, VGIMAX As Double
Dim BGS As Double, NDR As Double
Dim IMAXD0 As Double, IMAXD1 As Double, VDIMAX As Double, BDS As Double
Dim VSAT0 As Double, VSAT1 As Double, VSAT2 As Double, VVSAT As Double, VVSAT2 As Double
Dim BVSAT As Double, BVSAT2 As Double
Dim IDR0 As Double, IDR1 As Double, Gamma2 As Double, IVGST0 As Double, IMAXG2 As Double
Dim KAPPA As Double, NGR As Double, IMAXD2 As Double
Dim VDIMAX2 As Double, BDS2 As Double, IMAXG3 As Double, IDSSCALING As Double
Dim Rs As Double, Rd As Double, Rg As Double

'EEHEMT model parameters
Dim VTO_EE As Double, GAMMA_EE As Double, VCH_EE As Double, VDELT_EE As Double, VDSO_EE As Double
Dim VSAT_EE As Double, KAPA_EE As Double
Dim VTSO_EE As Double, GMMAX_EE As Double, VGO_EE As Double, VCO_EE As Double, MU_EE As Double
Dim VBA_EE As Double, VBC_EE As Double, DELTGM_EE As Double
Dim ALPHA_EE As Double, ISG_EE As Double, NF_EE As Double, KBK_EE As Double, IDSOC_EE As Double
Dim VBR_EE As Double, NBR_EE As Double, PEFF_EE As Double
Dim UGW_EE As Double, NGF_EE As Double, WG_EE As Double


'Simulation parameters
Dim debugmode As Integer
Dim Ddevice$, VDD00 As Double, VDDA0 As Double, VGmax As Double, VGmin As Double
Dim VDDMAX As Double, VDDMIN As Double, VGGMAX As Double, VGGMIN As Double

'Optimization parameters
Dim prm(100)


Sub IdsVds_make_data()
Dim vx As Double, vy As Double, vz As Double, vd As Double, vg As Double, vs As Double
Dim vds As Double, vgs As Double, vgd As Double
Dim VDD As Double, VGG As Double
Dim vdsmin As Double, vdsmax As Double, dvds As Double, vgsmin As Double, vgsmax As Double, dvgs As Double
Dim idrainx As Double, igatex As Double
Dim xids As Double
Dim level As Integer

Dim iter



    Call Initialize
    Call parameters_set
    
level = 1: ' select FET model

isweep = 21
istep = 10

    
ixmax = isweep
jymax = istep

vdsmin = VDDMIN
vdsmax = VDDMAX
dvds = (vdsmax - vdsmin) / (ixmax - 1)

vgsmin = VGGMIN
vgsmax = VGGMAX
dvgs = (vgsmax - vgsmin) / (jymax - 1)

For j = 1 To jymax
    VGG = dvgs * CDbl(j - 1) + vgsmin
    Sheet1.Cells(1, 2 + j) = VGG
    For i = 1 To ixmax
        Sheet1.Cells(i + 1, 1) = i
        VDD = dvds * CDbl(i - 1) + vdsmin
        
        
        Call MultiNewton(level, vd, vg, vs, idrainx, igatex, VDD, VGG, iter)

        vx = vd - vs
        vy = vg - vs
        vz = vg - vd
        
        'xids = Ids_f(vx, vy) - Igd_f(vx, vy)
        xids = Ids(level, vx, vy, vz) - Igd_f(vx, vy, vz)
        'xids = Ids_f(VDD, VGG) - Igd_f(VDD, VGG)
        'xids = Ids0_s(vx, vy)
        'xids2 = gds_s(vx, vy)
        'xids2 = gm_f(vx, vy) - ggd_f(vx, vy)
        xids2 = dfn7x("Idrain", level, vx, vy, vz, 0.001, 1)
        'xids2 = Igd_f(vx, vy)
        'xids2 = gds_s(vx, vy) - ggd_f(vx, vy)
        'xids = Igs_f(vx, vy)
        'xids = Ivgst(vx, vy)
        'xids2 = Idsgsmax(vx, vy)
        'xids = Idsgs(vx, vy)
        'xids2 = gvsat(vx, vy)
        'xids = Ivdst(vx, vy)
        'xids2 = Idsdsmax(vx, vy)
        'xids = Idsds(vx, vy)
        
        Sheet1.Cells(i + 1, 2) = VDD
        Sheet1.Cells(i + 1, 2 + j) = xids: 'Ids_f(vds, vgs)
        
        Sheet4.Cells(i + 1, 2) = VDD
        Sheet4.Cells(i + 1, 2 + j) = xids2: 'Ids_f(vds, vgs)
        
        'Sheet1.Cells(i + 1, 2 + j) = vx: 'Ids_f(vds, vgs)
        
    Next i
    
Next j

ixmax = istep
jymax = isweep

vdsmin = 0.1 * VDDMAX
vdsmax = VDDMAX
dvds = (vdsmax - vdsmin) / (ixmax - 1)

vgsmin = VGGMIN
vgsmax = VGGMAX
dvgs = (vgsmax - vgsmin) / (jymax - 1)


For i = 1 To ixmax
    VDD = dvds * CDbl(i - 1) + vdsmin
    Sheet1.Cells(1, 14 + i) = VDD
    For j = 1 To jymax

        VGG = dvgs * CDbl(j - 1) + vgsmin
        
        Call MultiNewton(level, vd, vg, vs, idrainx, igatex, VDD, VGG, iter)

        vx = vd - vs
        vy = vg - vs

        'xids = Ids_f(vx, vy) - Igd_f(vx, vy)
        xids = Ids(level, vx, vy, vz) - Igd_f(vx, vy, vz)
        'xids2 = gm_f(vx, vy) - ggd_f(vx, vy)
        xids2 = dfn7y("Ids", level, vx, vy, vz, 0.001, 1)
        'xids = Ids0_f(vx, vy)
        'xids = gds_f(vx, vy)
        'xids2 = gm3_f(vx, vy)
        'xids = Igd_f(vx, vy)
        'xids2 = Igs_f(vx, vy)
        'xids = Ivgst(vx, vy)
        'xids2 = Idsgsmax(vx, vy)
        'xids = Idsgs(vx, vy)
        'xids2 = gvsat(vx, vy)
        'xids = Ivdst(vx, vy)
        'xids2 = Idsdsmax(vx, vy)
        'xids = Idsds(vx, vy)
        
        Sheet1.Cells(j + 1, 14) = VGG
        Sheet1.Cells(j + 1, 14 + i) = xids: 'Ids_f(vds, vgs)
        
        Sheet4.Cells(j + 1, 14) = VGG
        Sheet4.Cells(j + 1, 14 + i) = xids2:

    Next j
Next i



End Sub


Sub Initialize()

    Sheet1.Select
    Sheet1.Cells.Select
    Selection.ClearContents
    Sheet2.Select
    Sheet2.Cells.Select
    Selection.ClearContents
    Sheet2.Range("A1").Select
    Sheet3.Select
    Sheet3.Cells.Select
    Selection.ClearContents
    Sheet3.Range("A1").Select
    Sheet4.Select
    Sheet4.Cells.Select
    Selection.ClearContents
    Sheet4.Range("A1").Select
    Sheet1.Select
    Sheet1.Range("A1").Select

End Sub
'========================================================
'                  Setting parameters
'========================================================

Sub parameters_set()
Dim Icurrent$


Icurrent$ = "": ' Icurrent$ = "Ids", "Igd", "Igs"

Ddevice$ = "fet"



itercount = 1
itermax = 15000

debugmode = 1

tera = 1000000000000#
giga = 1000000000#
mega = 1000000000#
meg = mega
kilo = 1000#
milli = 0.001
micro = 0.000001
nano = 0.000000001
pico = 0.000000001
femt = 0.000000000001

tol = 100# * femt:  'tolerance of convergence @ MultiNewton
GMIN = pico

VDDMAX = 10#
VDDMIN = -1# + 1# * pico
VGGMAX = 0.8
VGGMIN = -3#

Isgs = 10# * pico
ngs = 1.2
IGSMAX = 1# * kilo
Isgd = Isgs
ngd = 8.7
IGDMAX = 1# * kilo

Boltz = 1.380662E-23
Charge = 1.6021892E-19
Temperature = 300#
Vkt = Boltz * Temperature / Charge

kB = 1.38064852E-23
Temp = 300
Qe = 1.60217662E-19
VTHERMAL = kB * Temp / Qe
PI = 3.14159265358979


Rs = 1.02
Rd = 0.9
Rg = 1.6959

'**** Level=1 ****

IDSSCALING = 10.5
IVGST0 = 1#
VTH = -0.1
GAMMA = 0.0001
Gamma2 = 0.005 * milli
AGS = 0.4
NGR = 1#

IMAXG0 = 0.1: ' 0.1
IMAXG1 = 0.091: '0.095 gives plateu of gm profile
IMAXG2 = 1#
VGIMAX = 0.12
BGS = 0.504

NDR = 0.5
IMAXD0 = 0.8
IMAXD1 = 0.1
VDIMAX = 0#
BDS = 0.3

VSAT0 = 0.1
VSAT1 = 2.15
VSAT2 = 0.2
VVSAT = 0#
VVSAT2 = 1.2
BVSAT = 0.5
BVSAT2 = 0.2

IDR0 = 1#
IDR1 = 1#

KAPPA = 1# * micro
IMAXD2 = 1#
VDIMAX2 = VDIMAX
BDS2 = BDS
IMAXG3 = IMAXG1


'**** Level=2 ****
Ddevice$ = "fet"
 ALPHA = 3.262
LAMBDA = 0.0024409
   VTO = -2.2013
 BETA = 0.055687
NT = 1#
NG = 2#
GAMMA = 0.001


'**** Level=3 (EEHEMT1) ****
VTO_EE = -2.5
VGO_EE = -1#
VCO_EE = -0.5
VCH_EE = 0.8:  'VTO < VCO < VCH

GAMMA_EE = 0.0001
VDELT_EE = 0#
VDSO_EE = 20#: ' VDSO > -1/GAMMA + Vdsmax
VSAT_EE = 2#
KAPA_EE = 0.001
PEFF_EE = 190000#
VTSO_EE = -19#

GMMAX_EE = 0.5
MU_EE = 0.01

VBA_EE = 1#: 'VBA > 0
VBC_EE = 0.3: 'VBC > 0

DELTGM_EE = 0.06: ' DELTGM < gmoff/(sqrt(ALPHA^2+(Vb0-Vc0)^2)-ALPHA)
ALPHA_EE = 0.2

ISG_EE = 1E-20
NF_EE = 1.6

KBK_EE = 0.001
IDSOC_EE = 0.1
VBR_EE = 100#
NBR_EE = 2#

UGW_EE = 1#
NGF_EE = 1#
WG_EE = 1#



End Sub

Sub IdsVds_check_functions()

Dim vx As Double, vy As Double, vd As Double, vg As Double, vs As Double
Dim vds As Double, vgs As Double
Dim VDD As Double, VGG As Double
Dim vdsmin As Double, vdsmax As Double, dvds As Double, vgsmin As Double, vgsmax As Double, dvgs As Double
Dim xids As Double

Dim iter
'    Sheet4.Select
'    Sheet4.Cells.Select
'    Selection.ClearContents
'    Sheet4.Range("A1").Select


    Call parameters_set


ixmax = 51
jymax = 10

vdsmin = -1# + 1# * milli
vdsmax = 50#
dvds = (vdsmax - vdsmin) / (ixmax - 1)

vgsmin = -8#
vgsmax = 1#
dvgs = (vgsmax - vgsmin) / (jymax - 1)

For j = 1 To jymax
    VGG = dvgs * CDbl(j - 1) + vgsmin
    Sheet4.Cells(1, 2 + j) = VGG
    For i = 1 To ixmax
        Sheet4.Cells(i + 1, 1) = i
        VDD = dvds * CDbl(i - 1) + vdsmin
        
        
        Call MultiNewton(vd, vg, vs, VDD, VGG, iter)

        vx = vd - vs
        vy = vg - vs
        vz = vg - vd
        
        xids = Ids_f(vx, vy, vz) - Igd_f(vx, vy, vz)
        'xids = Ids_f(VDD, VGG) - Igd_f(VDD, VGG)
        'xids = Ids0_s(vx, vy)
        'xids = gds_s(vx, vy)
        'xids = Igd_f(vx, vy)
        'xids = Igs_f(vx, vy)
        'xids = Ivgst(vx, vy)
        xids = Idsgsmax(VDD, VGG, VGG - VDD)
        'xids = Idsgs(VDD, VGG)
        'xids = gvsat(VDD, VGG)
        'xids = Ivdst(VDD, VGG)
        'xids = Idsdsmax(VDD, VGG)
        'xids = Idsds(VDD, VGG)
        
        Sheet4.Cells(i + 1, 2) = VDD
        Sheet4.Cells(i + 1, 2 + j) = xids: 'Ids_f(vds, vgs)
        'Sheet1.Cells(i + 1, 2 + j) = vx: 'Ids_f(vds, vgs)
        
    Next i
    
Next j

ixmax = 10
jymax = 51

vdsmin = 5
vdsmax = 50#
dvds = (vdsmax - vdsmin) / (ixmax - 1)

vgsmin = -4#
vgsmax = 1#
dvgs = (vgsmax - vgsmin) / (jymax - 1)


For i = 1 To ixmax
    VDD = dvds * CDbl(i - 1) + vdsmin
    Sheet4.Cells(1, 14 + i) = VDD
    For j = 1 To jymax

        VGG = dvgs * CDbl(j - 1) + vgsmin
        
        Call MultiNewton(vd, vg, vs, VDD, VGG, iter)

        vx = vd - vs
        vy = vg - vs

        xids = Ids_f(vx, vy, vz) - Igd_f(vx, vy, vz)
        'xids = Ids_f(VDD, VGG) - Igd_f(VDD, VGG)
        'xids = Ids0_f(vx, vy)
        'xids = gds_f(vx, vy)
        'xids = gm_f(vx, vy)
        'xids = Igd_f(vx, vy)
        'xids = Igs_f(vx, vy)
        'xids = Ivgst(vx, vy)
        xids = Idsgsmax(VDD, VGG, VGG - VDD)
        'xids = Idsgs(VDD, VGG)
        'xids = gvsat(VDD, VGG)
        'xids = Ivdst(VDD, VGG)
        'xids = Idsdsmax(VDD, VGG)
        'xids = Idsds(VDD, VGG)
        
        Sheet4.Cells(j + 1, 14) = VGG
        Sheet4.Cells(j + 1, 14 + i) = xids:

    Next j
Next i




End Sub


Function Ids_f(vdsx As Double, vgsx As Double, vgdx As Double) As Double
Dim vds As Double, vgs As Double, vgd As Double, sidr As Double
Dim t8  As Double, t12 As Double, t18 As Double, t24 As Double, t33 As Double, t42 As Double, t44 As Double, t45 As Double, t48 As Double

vds = vdsx
vgs = vgsx
vgd = vgdx

If vds < 0 Then
    'vgs = vgs - vds
    vgs = vgd
    sidr = -1#
Else
    sidr = 1#
End If

'      t8 = IMAXG1 / (1 + IMAXG2 * expm(-(vgsx - VGIMAX) / BGS))
'      t12 = sidr * (vds) ^ 2
'      t18 = expm((vgs - VTH + Gamma * vds + Gamma2 * t12) / AGS) ^ NGR
'      t24 = sidr * vds
'      t33 = sidr * (IMAXD0 + IMAXD1 / (1 + expm(-(t24 - VDIMAX) / BDS)))
'      t42 = VSAT0 + VSAT1 / (1 + expm(-(vgs - VVSAT) / BVSAT))
'      t44 = (t24 * t42) ^ NDR
'      t45 = sidr * IDR0 * t44
'      t48 = IDR1 * tanh(vds * t42)
'      Ids_f = IDSSCALING *((IMAXG0 + t8) * IVGST0 * t18 / (IMAXG0 + t8 + IVGST0 * t18) * t33 * (t45 + t48) / (t33 + t45 + t48))
      
    t8 = IMAXG1 / (1 + IMAXG2 * expm(-(vgsx - VGIMAX) / BGS))
    t12 = (vds) ^ 2
    t18 = (expm((vgs - VTH + GAMMA * vds + sidr * Gamma2 * t12) / AGS)) ^ NGR
    t24 = sidr * vds
    t33 = sidr * (IMAXD0 + IMAXD1 / (1 + expm(-(t24 - VDIMAX) / BDS)))
    t42 = VSAT0 + VSAT1 / (1 + expm(-(vgsx - VVSAT) / BVSAT))
    t44 = (t24 * t42) ^ NDR
    t45 = sidr * IDR0 * t44
    t48 = IDR1 * tanh(vds * t42)
    Ids_f = IDSSCALING * ((IMAXG0 + t8) * IVGST0 * t18 / (IMAXG0 + t8 + IVGST0 * t18) * t33 * (t45 + t48) / (t33 + t45 + t48)) + KAPPA * vds + GMIN * vds
      
'      t1 = Ivgst(vds, vgs)
'      t2 = Idsgsmax(vds, vgs)
'      t7 = Idsdsmax(vds, vgs)
'      t8 = Ivdst(vds, vgs)
'      Ids_f = IDSSCALING *t1 * t2 / (t1 + t2) * t7 * t8 / (t7 + t8)
      
      
      
End Function
Function gds_f(vdsx As Double, vgsx As Double, vgdx As Double) As Double
Dim vds As Double, vgs As Double, vgd As Double


vds = vdsx
vgs = vgsx
vgd = vgdx


    If vds < 0 Then
        vgs = vgs - vds
        sidr = -1#
    Else
        sidr = 1#
    End If

    t8 = IMAXG1 / (1 + IMAXG2 * expm(-(vgs - VGIMAX) / BGS))
    t9 = IMAXG0 + t8
    t10 = t9 * IVGST0
    t12 = (vds) ^ 2
    t15 = 1 / AGS
    t18 = (expm((vgs - VTH + GAMMA * vds + Gamma2 * t12) * t15)) ^ NGR
    t21 = GAMMA + 2 * Gamma2 * vds
    t25 = IMAXG0 + t8 + IVGST0 * t18
    t26 = 1 / t25
    't29 = sidr * vds
    t29 = sidr
    t31 = 1 / BDS
    t33 = expm(-(t29 - VDIMAX) * t31)
    t34 = 1 + t33
    t37 = IMAXD0 + IMAXD1 / t34
    t40 = 1 / BVSAT
    t42 = expm(-(vds - VVSAT) * t40)
    t43 = 1 + t42
    t45 = VSAT1 / t43
    t46 = VSAT0 + t45
    t48 = (sidr * t29 * t46) ^ NDR
    t49 = sidr * IDR0 * t48
    t51 = tanh(vds * t46)
    t52 = IDR1 * t51
    t53 = t49 + t52
    t55 = sidr * t37
    t56 = t55 + t49 + t52
    t57 = 1 / t56
    t58 = t37 * t53 * t57
    t61 = (IVGST0) ^ 2
    t63 = (t18) ^ 2
    t64 = (t25) ^ 2
    t73 = t18 * t26
    t74 = (sidr) ^ 2
    t77 = (t34) ^ 2
    t78 = 1 / t77
    t85 = t10 * t73
    t90 = (t43) ^ 2
    t93 = 1 / t90 * t40 * t42
    't100 = IDR0 * t48 * NDR * (sidr * t46 + t29 * VSAT1 * t93) / vds / t46
    t100 = IDR0 * t48 * NDR * (sidr * t46 + t29 * VSAT1 * t93) / t46
    t101 = (t51) ^ 2
    t107 = IDR1 * (1 - t101) * (VSAT0 + t45 + vds * VSAT1 * t93)
    t112 = (t56) ^ 2
    gds_f = t10 * t18 * NGR * t21 * t15 * t26 * sidr * t58 - t9 * t61 * t63 / t64 * sidr * t58 * NGR * t21 * t15 + t10 * t73 * t74 * IMAXD1 * t78 * t31 * t33 * t53 * t57 + t85 * t55 * (t100 + t107) * t57 - t85 * t55 * t53 / t112 * (t74 * IMAXD1 * t78 * t31 * t33 + t100 + t107) + KAPPA + GMIN

End Function

Function gm_f(vdsx As Double, vgsx As Double, vgdx As Double) As Double
Dim vds As Double, vgs As Double, vgd As Double


vds = vdsx
vgs = vgsx
vgd = vgdx

    If vds < 0 Then
        vgs = vgs - vds
        sidr = -1#
    Else
        sidr = 1#
    End If

    t2 = 1 / BGS
    t4 = expm(-(vgs - VGIMAX) * t2)
    t6 = 1 + IMAXG2 * t4
    t7 = (t6) ^ 2
    t9 = IMAXG1 / t7
    t15 = (vds) ^ 2
    t18 = 1 / AGS
    t21 = (expm((vgs - VTH + GAMMA * vds + Gamma2 * t15) * t18)) ^ NGR
    t23 = IMAXG1 / t6
    t24 = IVGST0 * t21
    t25 = IMAXG0 + t23 + t24
    t26 = 1 / t25
    t29 = sidr * vds
    t37 = IMAXD0 + IMAXD1 / (1 + expm(-(t29 - VDIMAX) / BDS))
    t46 = VSAT0 + VSAT1 / (1 + expm(-(vds - VVSAT) / BVSAT))
    t48 = (t29 * t46) ^ NDR
    t49 = sidr * IDR0 * t48
    t52 = IDR1 * tanh(vds * t46)
    t53 = t49 + t52
    t55 = sidr * t37
    t57 = 1 / (t55 + t49 + t52)
    t58 = t37 * t53 * t57
    t62 = (IMAXG0 + t23) * IVGST0
    t69 = (t25) ^ 2
    gm_f = t9 * IMAXG2 * t2 * t4 * IVGST0 * t21 * t26 * sidr * t58 + t62 * t21 * NGR * t18 * t26 * sidr * t58 - t62 * t21 / t69 * t55 * t53 * t57 * (t9 * IMAXG2 * t2 * t4 + t24 * NGR * t18)

End Function

Function gm2_f(vdsx As Double, vgsx As Double, vgdx As Double) As Double
Dim vds As Double, vgs As Double, vgd As Double


vds = vdsx
vgs = vgsx
vgd = vgdx

    If vds < 0 Then
        vgs = vgs - vds
        sidr = -1#
    Else
        sidr = 1#
    End If


    t2 = 1 / BGS
    t4 = expm(-(vgs - VGIMAX) * t2)
    t6 = 1 + IMAXG2 * t4
    t7 = (t6) ^ 2
    t10 = IMAXG1 / t7 / t6
    t11 = (IMAXG2) ^ 2
    t13 = (BGS) ^ 2
    t14 = 1 / t13
    t15 = (t4) ^ 2
    t20 = (vds) ^ 2
    t23 = 1 / AGS
    t26 = expm((vgs - VTH + GAMMA * vds + Gamma2 * t20) * t23) ^ NGR
    t28 = IMAXG1 / t6
    t29 = IVGST0 * t26
    t30 = IMAXG0 + t28 + t29
    t31 = 1 / t30
    t34 = sidr * vds
    t42 = IMAXD0 + IMAXD1 / (1 + expm(-(t34 - VDIMAX) / BDS))
    t51 = VSAT0 + VSAT1 / (1 + expm(-(vds - VVSAT) / BVSAT))
    t53 = (t34 * t51) ^ NDR
    t54 = sidr * IDR0 * t53
    t57 = IDR1 * tanh(vds * t51)
    t58 = t54 + t57
    t59 = t42 * t58
    t60 = sidr * t42
    t62 = 1 / (t60 + t54 + t57)
    t63 = t59 * t62
    t64 = t26 * t31 * sidr * t63
    t67 = IMAXG1 / t7
    t68 = t67 * IMAXG2
    t73 = t2 * t4
    t76 = NGR * t23
    t78 = t58 * t62
    t84 = (t30) ^ 2
    t85 = 1 / t84
    t86 = t26 * t85
    t92 = t67 * IMAXG2 * t2 * t4 + t29 * t76
    t98 = (IMAXG0 + t28) * IVGST0
    t99 = (NGR) ^ 2
    t101 = (AGS) ^ 2
    t102 = 1 / t101
    t120 = (t92) ^ 2
    gm2_f = 2 * t10 * t11 * t14 * t15 * IVGST0 * t64 - t68 * t14 * t4 * IVGST0 * t64 + 2 * t68 * t73 * t29 * t76 * t31 * t60 * t78 - 2 * t68 * t73 * IVGST0 * t86 * sidr * t59 * t62 * t92 + t98 * t26 * t99 * t102 * t31 * sidr * t63 - 2 * t98 * t26 * NGR * t23 * t85 * sidr * t42 * t78 * t92 + 2 * t98 * t26 / t84 / t30 * t60 * t78 * t120 - t98 * t86 * t60 * t78 * (2 * t10 * t11 * t14 * t15 - t67 * IMAXG2 * t14 * t4 + t29 * t99 * t102)

End Function

Function gm3_f(vdsx As Double, vgsx As Double, vgdx As Double) As Double
Dim vds As Double, vgs As Double, vgd As Double


vds = vdsx
vgs = vgsx
vgd = vgdx

    If vds < 0 Then
        vgs = vgs - vds
        sidr = -1#
    Else
        sidr = 1#
    End If

    t2 = 1 / BGS
    t4 = expm(-(vgs - VGIMAX) * t2)
    t6 = 1 + IMAXG2 * t4
    t7 = (t6) ^ 2
    t10 = IMAXG1 / t7 / t6
    t11 = (IMAXG2) ^ 2
    t12 = t10 * t11
    t13 = (BGS) ^ 2
    t14 = 1 / t13
    t15 = (t4) ^ 2
    t16 = t14 * t15
    t20 = (vds) ^ 2
    t23 = 1 / AGS
    t26 = (expm((vgs - VTH + GAMMA * vds + Gamma2 * t20) * t23)) ^ NGR
    t28 = IMAXG1 / t6
    t29 = IVGST0 * t26
    t30 = IMAXG0 + t28 + t29
    t31 = (t30) ^ 2
    t32 = 1 / t31
    t33 = t26 * t32
    t34 = t33 * sidr
    t35 = sidr * vds
    t43 = IMAXD0 + IMAXD1 / (1 + expm(-(t35 - VDIMAX) / BDS))
    t52 = VSAT0 + VSAT1 / (1 + expm(-(vds - VVSAT) / BVSAT))
    t54 = (t35 * t52) ^ NDR
    t55 = sidr * IDR0 * t54
    t58 = IDR1 * tanh(vds * t52)
    t59 = t55 + t58
    t60 = t43 * t59
    t61 = sidr * t43
    t63 = 1 / (t61 + t55 + t58)
    t65 = IMAXG1 / t7
    t69 = NGR * t23
    t71 = t65 * IMAXG2 * t2 * t4 + t29 * t69
    t72 = t63 * t71
    t73 = t60 * t72
    t74 = t34 * t73
    t76 = (t7) ^ 2
    t78 = IMAXG1 / t76
    t79 = t11 * IMAXG2
    t82 = 1 / t13 / BGS
    t83 = t15 * t4
    t87 = 1 / t30
    t90 = t60 * t63
    t91 = t26 * t87 * sidr * t90
    t100 = t59 * t63
    t101 = t61 * t100
    t102 = t69 * t87 * t101
    t104 = t65 * IMAXG2
    t105 = t14 * t4
    t116 = t2 * t4
    t118 = t104 * t116 * t29
    t119 = t32 * sidr
    t123 = (NGR) ^ 2
    t124 = (AGS) ^ 2
    t125 = 1 / t124
    t126 = t123 * t125
    t131 = t104 * t116 * IVGST0
    t139 = 2 * t10 * t11 * t14 * t15 - t65 * IMAXG2 * t14 * t4 + t29 * t126
    t145 = 1 / t31 / t30
    t147 = t26 * t145 * sidr
    t148 = (t71) ^ 2
    t154 = (IMAXG0 + t28) * IVGST0
    t165 = t123 * NGR
    t167 = 1 / t124 / AGS
    t176 = t154 * t26 * NGR * t23
    t177 = t119 * t43
    t181 = (t31) ^ 2
    gm3_f = -6 * t12 * t16 * IVGST0 * t74 + 6 * t78 * t79 * t82 * t83 * IVGST0 * t91 - 6 * t12 * t82 * t15 * IVGST0 * t91 + 6 * t12 * t16 * t29 * t102 + 3 * t104 * t105 * IVGST0 * t74 + t104 * t82 * t4 * IVGST0 * t91 - 3 * t104 * t105 * t29 * t102 - 6 * t118 * t69 * t119 * t73 + 3 * t118 * t126 * t87 * t101 - 3 * t131 * t34 * t60 * t63 * t139 + 6 * t131 * t147 * t60 * t63 * t148 - t154 * t33 * t61 * t100 * (6 * t78 * t79 * t82 * t83 - 6 * t10 * t11 * t82 * t15 + t65 * IMAXG2 * t82 * t4 + t29 * t165 * t167) - 3 * t176 * t177 * t100 * t139 - 6 * t154 * t26 / t181 * t61 * t100 * t148 * t71 + 6 * t154 * t147 * t60 * t72 * t139 + t154 * t26 * t165 * t167 * t87 * sidr * t90 + 6 * t176 * t145 * sidr * t43 * t100 * t148 - 3 * t154 * t26 * t123 * t125 * t177 * t100 * t71

End Function

Function gm4_f(vdsx As Double, vgsx As Double, vgdx As Double) As Double
Dim vds As Double, vgs As Double, vgd As Double


vds = vdsx
vgs = vgsx
vgd = vgdx

    If vds < 0 Then
        vgs = vgs - vds
        sidr = -1#
    Else
        sidr = 1#
    End If


    t2 = 1 / BGS
    t4 = expm(-(vgs - VGIMAX) * t2)
    t6 = 1 + IMAXG2 * t4
    t7 = (t6) ^ 2
    t10 = IMAXG1 / t7 / t6
    t11 = (IMAXG2) ^ 2
    t12 = t10 * t11
    t13 = (BGS) ^ 2
    t14 = 1 / t13
    t15 = (t4) ^ 2
    t16 = t14 * t15
    t18 = t12 * t16 * IVGST0
    t20 = (vds) ^ 2
    t23 = 1 / AGS
    t26 = expm((vgs - VTH + GAMMA * vds + Gamma2 * t20) * t23) ^ NGR
    t28 = IMAXG1 / t6
    t29 = IVGST0 * t26
    t30 = IMAXG0 + t28 + t29
    t31 = (t30) ^ 2
    t33 = 1 / t31 / t30
    t34 = t26 * t33
    t35 = t34 * sidr
    t36 = sidr * vds
    t44 = IMAXD0 + IMAXD1 / (1 + expm(-(t36 - VDIMAX) / BDS))
    t53 = VSAT0 + VSAT1 / (1 + expm(-(vds - VVSAT) / BVSAT))
    t55 = (t36 * t53) ^ NDR
    t56 = sidr * IDR0 * t55
    t59 = IDR1 * tanh(vds * t53)
    t60 = t56 + t59
    t61 = t44 * t60
    t62 = sidr * t44
    t64 = 1 / (t62 + t56 + t59)
    t66 = IMAXG1 / t7
    t70 = NGR * t23
    t72 = t66 * IMAXG2 * t2 * t4 + t29 * t70
    t73 = (t72) ^ 2
    t74 = t64 * t73
    t75 = t61 * t74
    t76 = t35 * t75
    t79 = (IMAXG0 + t28) * IVGST0
    t80 = (NGR) ^ 2
    t82 = (AGS) ^ 2
    t83 = 1 / t82
    t85 = t79 * t26 * t80 * t83
    t86 = t33 * sidr
    t87 = t86 * t44
    t88 = t60 * t64
    t92 = (t7) ^ 2
    t94 = IMAXG1 / t92
    t95 = t11 * IMAXG2
    t96 = t94 * t95
    t97 = (t13) ^ 2
    t98 = 1 / t97
    t99 = t15 * t4
    t103 = 1 / t30
    t106 = t61 * t64
    t107 = t26 * t103 * sidr * t106
    t110 = 1 / t13 / BGS
    t111 = t110 * t99
    t114 = 1 / t31
    t115 = t26 * t114
    t116 = t115 * sidr
    t118 = t61 * t64 * t72
    t119 = t116 * t118
    t121 = t66 * IMAXG2
    t122 = t110 * t4
    t126 = t14 * t4
    t128 = t121 * t126 * IVGST0
    t135 = t80 * t83
    t137 = 2 * t10 * t11 * t14 * t15 - t66 * IMAXG2 * t14 * t4 + t29 * t135
    t138 = t64 * t137
    t139 = t61 * t138
    t140 = t116 * t139
    t143 = t121 * t29 * t126
    t145 = t62 * t88
    t146 = t135 * t103 * t145
    t148 = t2 * t4
    t150 = t121 * t148 * t29
    t159 = (t137) ^ 2
    t163 = t80 * NGR
    t166 = 1 / t82 / AGS
    t169 = t114 * sidr
    t170 = t169 * t44
    t174 = (t31) ^ 2
    t175 = 1 / t174
    t177 = t26 * t175 * sidr
    t184 = IMAXG1 / t92 / t6
    t185 = (t11) ^ 2
    t187 = (t15) ^ 2
    t192 = t110 * t15
    t197 = t121 * t148 * IVGST0
    t207 = t163 * t166
    t209 = 6 * t94 * t95 * t110 * t99 - 6 * t10 * t11 * t110 * t15 + t66 * IMAXG2 * t110 * t4 + t29 * t207
    t210 = t64 * t209
    t215 = t70 * t169
    t221 = t70 * t103 * t145
    t223 = 24 * t18 * t76 + 12 * t85 * t87 * t88 * t73 - 36 * t96 * t98 * t99 * IVGST0 * t107 - 24 * t96 * t111 * IVGST0 * t119 - 4 * t121 * t122 * IVGST0 * t119 + 6 * t128 * t140 - 6 * t143 * t146 + 24 * t150 * t70 * t86 * t75 + 24 * t150 * t87 * t88 * t137 * t72 + 6 * t79 * t34 * t62 * t88 * t159 - 4 * t79 * t26 * t163 * t166 * t170 * t88 * t72 - 36 * t79 * t177 * t61 * t74 * t137 + 24 * t184 * t185 * t98 * t187 * IVGST0 * t107 + 24 * t12 * t192 * IVGST0 * t119 - 4 * t197 * t116 * t61 * t210 - 12 * t128 * t76 - 12 * t150 * t215 * t139 - 24 * t12 * t192 * t29 * t221
    t225 = t12 * t16 * t29
    t226 = t215 * t118
    t233 = (t80) ^ 2
    t235 = (t82) ^ 2
    t236 = 1 / t235
    t290 = (t73) ^ 2
    t296 = t79 * t26 * NGR * t23
    t299 = t73 * t72
    t317 = -24 * t225 * t226 - 12 * t18 * t140 + 8 * t79 * t35 * t61 * t210 * t72 + t79 * t26 * t233 * t236 * t103 * sidr * t106 - 6 * t85 * t170 * t88 * t137 - 12 * t150 * t135 * t169 * t118 + 12 * t225 * t146 + 24 * t79 * t26 * t70 * t33 * t62 * t60 * t138 * t72 - t121 * t98 * t4 * IVGST0 * t107 - t79 * t115 * t62 * t88 * (24 * t184 * t185 * t98 * t187 - 36 * t94 * t95 * t98 * t99 + 14 * t10 * t11 * t98 * t15 - t66 * IMAXG2 * t98 * t4 + t29 * t233 * t236) + 24 * t96 * t111 * t29 * t221 + 12 * t143 * t226 + 4 * t121 * t122 * t29 * t221 + 24 * t79 * t26 / t174 / t30 * t62 * t88 * t290 - 24 * t296 * t175 * sidr * t44 * t88 * t299 - 4 * t296 * t170 * t88 * t209 + 14 * t12 * t98 * t15 * IVGST0 * t107 - 24 * t197 * t177 * t61 * t64 * t299 + 4 * t150 * t207 * t103 * t145
    gm4_f = t223 + t317

End Function

Function gm5_f(vdsx As Double, vgsx As Double, vgdx As Double) As Double
Dim vds As Double, vgs As Double, vgd As Double


vds = vdsx
vgs = vgsx
vgd = vgdx

    If vds < 0 Then
        vgs = vgs - vds
        sidr = -1#
    Else
        sidr = 1#
    End If


    t2 = 1 / BGS
    t4 = expm(-(vgs - VGIMAX) * t2)
    t6 = 1 + IMAXG2 * t4
    t8 = IMAXG1 / t6
    t10 = (IMAXG0 + t8) * IVGST0
    t12 = (vds) ^ 2
    t15 = 1 / AGS
    t18 = expm((vgs - VTH + GAMMA * vds + Gamma2 * t12) * t15) ^ NGR
    t19 = IVGST0 * t18
    t20 = IMAXG0 + t8 + t19
    t21 = (t20) ^ 2
    t22 = (t21) ^ 2
    t23 = 1 / t22
    t25 = t18 * t23 * sidr
    t26 = t10 * t25
    t27 = sidr * vds
    t35 = IMAXD0 + IMAXD1 / (1 + expm(-(t27 - VDIMAX) / BDS))
    t44 = VSAT0 + VSAT1 / (1 + expm(-(vds - VVSAT) / BVSAT))
    t46 = (t27 * t44) ^ NDR
    t47 = sidr * IDR0 * t46
    t50 = IDR1 * tanh(vds * t44)
    t51 = t47 + t50
    t52 = t35 * t51
    t53 = sidr * t35
    t55 = 1 / (t53 + t47 + t50)
    t56 = (t6) ^ 2
    t58 = IMAXG1 / t56
    t59 = IMAXG2 * t2
    t62 = NGR * t15
    t64 = t58 * t59 * t4 + t19 * t62
    t65 = (t64) ^ 2
    t66 = t55 * t65
    t67 = (t56) ^ 2
    t69 = IMAXG1 / t67
    t70 = (IMAXG2) ^ 2
    t71 = t70 * IMAXG2
    t72 = (BGS) ^ 2
    t74 = 1 / t72 / BGS
    t76 = (t4) ^ 2
    t77 = t76 * t4
    t82 = IMAXG1 / t56 / t6
    t89 = (NGR) ^ 2
    t90 = t89 * NGR
    t91 = (AGS) ^ 2
    t93 = 1 / t91 / AGS
    t94 = t90 * t93
    t96 = 6 * t69 * t71 * t74 * t77 - 6 * t82 * t70 * t74 * t76 + t58 * IMAXG2 * t74 * t4 + t19 * t94
    t100 = t55 * t64
    t101 = 1 / t72
    t108 = 1 / t91
    t109 = t89 * t108
    t111 = 2 * t82 * t70 * t101 * t76 - t58 * IMAXG2 * t101 * t4 + t19 * t109
    t112 = (t111) ^ 2
    t116 = t82 * t70
    t117 = (t72) ^ 2
    t118 = 1 / t117
    t119 = t118 * t76
    t122 = 1 / t21
    t123 = t18 * t122
    t124 = t123 * sidr
    t125 = t52 * t100
    t126 = t124 * t125
    t130 = 1 / t20
    t132 = t51 * t55
    t133 = t53 * t132
    t134 = t62 * t130 * t133
    t136 = t58 * IMAXG2
    t137 = t2 * t4
    t139 = t136 * t137 * IVGST0
    t142 = IMAXG1 / t67 / t6
    t143 = (t70) ^ 2
    t145 = (t76) ^ 2
    t157 = (t89) ^ 2
    t158 = (t91) ^ 2
    t159 = 1 / t158
    t160 = t157 * t159
    t162 = 24 * t142 * t143 * t118 * t145 - 36 * t69 * t71 * t118 * t77 + 14 * t82 * t70 * t118 * t76 - t58 * IMAXG2 * t118 * t4 + t19 * t160
    t163 = t55 * t162
    t167 = t142 * t143
    t168 = t118 * t145
    t176 = 1 / t21 / t20
    t178 = t18 * t176 * sidr
    t183 = t69 * t71
    t184 = t118 * t77
    t188 = t101 * t76
    t190 = t116 * t188 * t19
    t191 = t176 * sidr
    t193 = t52 * t66
    t194 = t62 * t191 * t193
    t199 = t74 * t76
    t201 = t116 * t199 * IVGST0
    t202 = t55 * t111
    t203 = t52 * t202
    t204 = t124 * t203
    t206 = t74 * t77
    t208 = t183 * t206 * IVGST0
    t211 = t136 * t137 * t19
    t212 = t122 * sidr
    t217 = t116 * t188 * IVGST0
    t218 = t55 * t96
    t219 = t52 * t218
    t220 = t124 * t219
    t222 = t10 * t178
    t226 = t74 * t4
    t228 = t136 * t226 * t19
    t230 = t109 * t130 * t133
    t232 = -60 * t26 * t52 * t66 * t96 - 90 * t26 * t52 * t100 * t112 - 70 * t116 * t119 * IVGST0 * t126 + 70 * t116 * t119 * t19 * t134 - 5 * t139 * t124 * t52 * t163 + 120 * t167 * t168 * t19 * t134 - 120 * t167 * t168 * IVGST0 * t126 + 30 * t139 * t178 * t52 * t55 * t112 - 180 * t183 * t184 * t19 * t134 + 120 * t190 * t194 + 180 * t183 * t184 * IVGST0 * t126 + 60 * t201 * t204 - 60 * t208 * t204 - 20 * t211 * t94 * t212 * t125 - 20 * t217 * t220 + 20 * t222 * t52 * t202 * t96 + 10 * t228 * t230
    t233 = t62 * t212
    t234 = t233 * t125
    t236 = t109 * t212
    t237 = t236 * t125
    t242 = IMAXG1 / t67 / t56
    t243 = t143 * IMAXG2
    t245 = 1 / t117 / BGS
    t247 = t145 * t4
    t262 = t157 * NGR
    t264 = 1 / t158 / AGS
    t271 = t178 * t193
    t275 = t18 * NGR
    t281 = t132 * t64 * t111
    t284 = t101 * t4
    t286 = t136 * t284 * t19
    t288 = t94 * t130 * t133
    t291 = 1 / t22 / t20
    t293 = t18 * t291 * sidr
    t294 = (t65) ^ 2
    t301 = t136 * t284 * IVGST0
    t302 = t65 * t64
    t303 = t55 * t302
    t304 = t52 * t303
    t305 = t25 * t304
    t309 = t116 * t199 * t19
    t311 = t191 * t35
    t312 = t311 * t281
    t314 = t23 * sidr
    t315 = t314 * t35
    t325 = t136 * t226 * IVGST0
    t330 = -20 * t228 * t234 - 60 * t190 * t237 - t10 * t123 * t53 * t132 * (120 * t242 * t243 * t245 * t247 - 240 * t142 * t143 * t245 * t145 + 150 * t69 * t71 * t245 * t77 - 30 * t82 * t70 * t245 * t76 + t58 * IMAXG2 * t245 * t4 + t19 * t262 * t264) + 120 * t208 * t271 + 120 * t58 * t59 * t4 * IVGST0 * t275 * t15 * t176 * t53 * t281 - 10 * t286 * t288 + 120 * t139 * t293 * t52 * t55 * t294 - 120 * t201 * t271 + 60 * t301 * t305 + 20 * t190 * t288 - 60 * t309 * t230 + 120 * t190 * t312 - 180 * t211 * t315 * t132 * t65 * t111 - 120 * t217 * t305 - 20 * t211 * t233 * t219 + 120 * t309 * t234 - 10 * t325 * t204 + 5 * t211 * t160 * t130 * t133
    t335 = t212 * t35
    t340 = t10 * t275 * t15
    t348 = t10 * t18 * t89 * t108
    t357 = t10 * t18 * t90 * t93
    t374 = t52 * t55
    t377 = t118 * t4
    t393 = t183 * t206 * t19
    t399 = t233 * t203
    t401 = -5 * t10 * t18 * t157 * t159 * t335 * t132 * t64 + 120 * t340 * t291 * sidr * t35 * t132 * t294 - 10 * t348 * t335 * t132 * t96 + 30 * t340 * t311 * t132 * t112 + 20 * t357 * t311 * t132 * t65 - 60 * t348 * t315 * t132 * t302 - 10 * t357 * t335 * t132 * t111 + 10 * t222 * t52 * t163 * t64 + t10 * t18 * t262 * t264 * t130 * sidr * t374 - 5 * t136 * t377 * t19 * t134 + 5 * t136 * t377 * IVGST0 * t126 - 30 * t211 * t236 * t203 + 60 * t211 * t109 * t191 * t193 - 120 * t211 * t62 * t314 * t304 + 60 * t393 * t230 + 40 * t211 * t311 * t132 * t96 * t64 - 60 * t190 * t399
    t426 = t18 * t130 * sidr * t374
    t428 = t10 * t18
    t431 = t53 * t51
    t465 = -120 * t393 * t234 + 240 * t10 * t293 * t52 * t303 * t111 - 60 * t286 * t312 + 10 * t301 * t220 + 20 * t325 * t271 + 30 * t286 * t237 + 30 * t286 * t399 - 60 * t286 * t194 - 120 * t10 * t18 / t22 / t21 * t53 * t132 * t294 * t64 + t136 * t245 * t4 * IVGST0 * t426 + 40 * t428 * t62 * t176 * t431 * t218 * t64 - 30 * t116 * t245 * t76 * IVGST0 * t426 - 240 * t167 * t245 * t145 * IVGST0 * t426 + 120 * t242 * t243 * t245 * t247 * IVGST0 * t426 + 150 * t183 * t245 * t77 * IVGST0 * t426 + 60 * t428 * t109 * t176 * t431 * t202 * t64 - 180 * t428 * t62 * t23 * t431 * t66 * t111 - 5 * t340 * t335 * t132 * t162
    gm5_f = t232 + t330 + t401 + t465

End Function


Function Ids_mod(vdsx As Double, vgsx As Double, vgdx As Double) As Double
Dim vds As Double, vgs As Double, vgd As Double


vds = vdsx
vgs = vgsx
vgd = vgdx

If vds < 0 Then
    vgs = vgs - vds
    sidr = -1#
Else
    sidr = 1#
End If

    If vds < 0 Then
        t8 = IMAXG3 / (1# + IMAXG2 * expm(-(vgs - VGIMAX) / BGS))
    Else
        t8 = IMAXG1 / (1# + IMAXG2 * expm(-(vgs - VGIMAX) / BGS))
    End If
    
    t12 = sidr * vds * vds
    t17 = IVGST0 * (expm((vgs - VTH + GAMMA * vds + Gamma2 * t12) / AGS)) ^ NGR
    t23 = sidr * vds
    
    If vds < 0 Then
      t32 = sidr * (IMAXD0 + IMAXD2 / (1# + expm(-(t23 - VDIMAX2) / BDS2)))
    Else
      t32 = sidr * (IMAXD0 + IMAXD1 / (1# + expm(-(t23 - VDIMAX) / BDS)))
    End If
    
    t41 = VSAT0 + VSAT1 / (1# + expm(-(vds - VVSAT) / BVSAT))
    t43 = (t23 * t41) ^ (NDR)
    t44 = sidr * IDR0 * t43
    t47 = IDR1 * tanh(vds * t41)
    Ids_mod = (IMAXG0 + t8) * IVGST0 * t17 / (IMAXG0 + t8 + IVGST0 * t17) * t32 * (t44 + t47) / (t32 + t44 + t47)
      
      
End Function



Function Ivgst(vdsx As Double, vgsx As Double, vgdx As Double) As Double
Dim vds As Double, vgs As Double, vgd As Double


vds = vdsx
vgs = vgsx
vgd = vgdx

If vds < 0 Then
    vgs = vgs - vds
    sidr = -1#
Else
    sidr = 1#
End If

Ivgst = IVGST0 * (expm(((vgs - VTH) + GAMMA * vds + sidr * Gamma2 * (vds) ^ 2) / AGS)) ^ NGR
'Ivgst = IVGST0 * (expm(((vgsx - VTH) + Gamma * vds + Gamma2 * (vds) ^ 2) / AGS)) ^ NGR

End Function

Function Idsgsmax(vdsx As Double, vgsx As Double, vgdx As Double) As Double
Dim vds As Double, vgs As Double, vgd As Double


vds = vdsx
vgs = vgsx
vgd = vgdx

If vds < 0 Then
    vgs = vgs - vds
    sidr = -1#
Else
    sidr = 1#
End If

Idsgsmax = IMAXG0 + IMAXG1 / (1 + IMAXG2 * expm(-(vgs - VGIMAX) / BGS))

End Function

Function Idsgs(vdsx As Double, vgsx As Double, vgdx As Double) As Double
Dim vds As Double, vgs As Double, vgd As Double


vds = vdsx
vgs = vgsx
vgd = vgdx

'If vds < 0 Then
'    vgs = vgs - vds
'End If

Idsgs = Idsgsmax(vds, vgs) * Ivgst(vds, vgs) / (Idsgsmax(vds, vgs) + Ivgst(vds, vgs))

End Function

Function gvsat(vdsx As Double, vgsx As Double, vgdx As Double) As Double
Dim vds As Double, vgs As Double, vgd As Double, vcore4 As Double, vcore2 As Double, icore4 As Double, icore2 As Double


vds = vdsx
vgs = vgsx
vgd = vgdx


'If vds < 0 Then
'    vgs = vgs - vds
'End If

'gvsat = VSAT0 + VSAT1 / (1# + expm(-(vgs - VVSAT) / BVSAT))

vcore4 = ((vgs - VVSAT) / BVSAT) ^ 4
vcore2 = -(vgs - VVSAT2) / BVSAT2
icore4 = expm(vcore4)
icore2 = expm(vcore2)
gvsat = VSAT0 + VSAT1 / (1# + icore4) + VSAT2 / (1 + icore2)


End Function

Function Ivdst(vdsx As Double, vgsx As Double, vgdx As Double) As Double
Dim vds As Double, vgs As Double, vgd As Double


vds = vdsx
vgs = vgsx
vgd = vgdx

If vds < 0 Then
    'vgs = vgs - vds
    sidr = -1#
Else
    sidr = 1#
End If

Ivdst = sidr * IDR0 * (sidr * vds * gvsat(vds, vgs)) ^ NDR + IDR1 * tanh(vds * gvsat(vds, vgs))

End Function

Function Idsdsmax(vdsx As Double, vgsx As Double, vgdx As Double) As Double
Dim vds As Double, vgs As Double, vgd As Double


vds = vdsx
vgs = vgsx
vgd = vgdx

If vds < 0 Then
    'vgs = vgs - vds
    sidr = -1#
Else
    sidr = 1#
End If

Idsdsmax = (IMAXD0 + IMAXD1 / (1 + expm(-((vds - VDIMAX) / BDS))))
'Idsdsmax = (IMAXD0 + IMAXD1 / (1 + expm(-((sidr * vds - VDIMAX) / BDS))))
'Idsdsmax = (IMAXD0 * tanh((vds - VDIMAX) / BDS))
'Idsdsmax = sidr * IMAXD0


End Function

Function Idsds(vdsx As Double, vgsx As Double, vgdx As Double)
Dim vds As Double, vgs As Double, vgd As Double

vds = vdsx
vgs = vgsx
vgd = vgdx

If vds < 0 Then
    vgs = vgs - vds
    sidr = -1#
Else
    sidr = 1#
End If

Idsds0 = Idsdsmax(vds, vgs, vgd) * Ivdst(vds, vgs, vgd)
Idsds1 = (Idsdsmax(vds, vgs, vgd) + Ivdst(vds, vgs, vgd))
Idsds = Idsds0 / Idsds1

End Function


Function Ids_org(vdsx As Double, vgsx As Double, vgdx As Double) As Double
Dim vds As Double, vgs As Double, vgd As Double
vds = vdsx
vgs = vgsx
vgd = vgdx

If vds < 0 Then
    vgs = vgdx
    vgd = vgsx
End If


    Ivgst = expm((vgs - prm(1) + prm(2) * vds + prm(19) * vds * vds) / prm(3))
    Idsgsmax = prm(4) + prm(5) / (1 + expm(-(vgs - prm(6)) / prm(7)))
    Idsgs = Idsgsmax * Ivgst / (Idsgsmax + Ivgst)
    
    gvsat = prm(13) + prm(14) / (1 + expm(-(vds - prm(15)) / prm(16)))
    Ivdst = prm(17) * (vds * gvsat) ^ prm(8) + prm(18) * tanh(vds * gvsat)
    Idsdsmax = prm(9) + prm(10) / (1 + expm(-(vds - prm(11)) / prm(12)))
    Idsds = Idsdsmax * Ivdst / (Idsdsmax + Ivdst)
    
    Ids_org = Idsgs * Idsds
    

End Function

Function Ids_s(vdsx As Double, vgsx As Double, vgdx As Double) As Double
Dim vds As Double, vgs As Double, vgd As Double
vds = vdsx
vgs = vgsx
vgd = vgdx

    If vds < 0 Then
        vgs = vgdx
        vgd = vgsx
    End If
    
    vgst = vgs - VTH
    
    If vgst < 0 Then
        t5 = 0#
    Else
        t5 = (vgst) ^ 2
    End If
    Ids_s = Ids0_s(vds, vgs, vgd) * t5 + GMIN * vds
    
End Function

Function Ids0_s(vdsx As Double, vgsx As Double, vgdx As Double) As Double
Dim vds As Double, vgs As Double, vgd As Double
vds = vdsx
vgs = vgsx
vgd = vgdx

    
    If vds < 0 Then
        vgs = vgdx
        vgd = vgsx
    End If
    
    Ids0_s = tanh(vds / VSAT0)

End Function

Function gds_s(vdsx As Double, vgsx As Double, vgdx As Double) As Double
Dim vds As Double, vgs As Double, vgd As Double
vds = vdsx
vgs = vgsx
vgd = vgdx

    If vds < 0 Then
        vgs = vgdx
        vgd = vgsx
    End If
    
    vgst = vgs - VTH
    
    If vgst < 0 Then
        vgst = 0#
    End If
    
    t1 = 1 / VSAT0
    t4 = (tanh(vds * t1)) ^ 2
    t8 = (vgst) ^ 2
    gds_s = (1 - t4) * t1 * t8 + GMIN


End Function

Function gm_s(vdsx As Double, vgsx As Double, vgdx As Double) As Double
Dim vds As Double, vgs As Double, vgd As Double
vds = vdsx
vgs = vgsx
vgd = vgdx

    If vds < 0 Then
        vgs = vgdx
        vgd = vgsx
    End If
    
    vgst = vgs - VTH
    
    If vgst < 0 Then
        vgst = 0#
    End If
    
    t4 = vgst
    gm_s = 2 * tanh(vds / VSAT0) * t4

End Function

'=============================================================================
'
'                                Test routine
'
'=============================================================================
Sub testroutine()
Dim vx As Double, vy As Double, vd As Double, vg As Double, vs As Double
Dim vds As Double, vgs As Double
Dim VDD As Double, VGG As Double
Dim vdsmin As Double, vdsmax As Double, dvds As Double, vgsmin As Double, vgsmax As Double, dvgs As Double
Dim idrainx As Double, igatex As Double
Dim xids As Double
Dim level As Integer

Dim iter



    Call Initialize
    Call parameters_set
    
    Sheet5.Select
    Sheet5.Cells.Select
    Selection.ClearContents
    Sheet5.Range("A1").Select
    
level = 1: ' select FET model(1: Hirose, 2: Simple, 3: EEHEMT1)

isweep = 51
istep = 10

    
ixmax = isweep
jymax = istep

vdsmin = VDDMIN
vdsmax = VDDMAX
dvds = (vdsmax - vdsmin) / (ixmax - 1)

vgsmin = VGGMIN
vgsmax = VGGMAX
dvgs = (vgsmax - vgsmin) / (jymax - 1)

For j = 1 To jymax
    VGG = dvgs * CDbl(j - 1) + vgsmin
    Sheet5.Cells(1, 2 + j) = VGG
    For i = 1 To ixmax
        Sheet5.Cells(i + 1, 1) = i
        VDD = dvds * CDbl(i - 1) + vdsmin
        
        Call MultiNewton(level, vd, vg, vs, idrainx, igatex, VDD, VGG, iter)
        vx = vd - vs
        vy = vg - vs
        vz = vg - vd

        'xids = Ids(level, vx, vy, vz) - Igd(level, vx, vy, vz)
        'xids = Ids(level, VDD, VGG, VGG-VDD) - Igd(level, vx, vy, vz) * 0#
        'xids = Ids_f(VDD, VGG, VGG-VDD) - Igd_f(VDD, VGG, VGG-VDD)
        'xids = Ids0_s(vx, vy, vz)
        'xids2 = gds_s(vx, vy, vz)
        'xids2 = gm_f(vx, vy, vz) - ggd_f(vx, vy, vz)
        'xids = dfn7x("Ids", level, VDD, VGG, VGG-VDD 0.001, 2)
        'xids2 = Igs(level, vx, vy, vz)
        'xids2 = gds_s(vx, vy, vz) - ggd_f(vx, vy, vz)
        'xids = Igs_f(vx, vy, vz)
        'xids = Ivgst(vx, vy, vz)
        'xids2 = Idsgsmax(vx, vy, vz)
        'xids = Idsgs(vx, vy, vz)
        'xids2 = gvsat(vx, vy, vz)
        'xids = Ivdst(vx, vy, vz)
        'xids2 = Idsdsmax(vx, vy, vz)
        'xids = Idsds(vx, vy, vz)
        
        Sheet5.Cells(i + 1, 2) = VDD
        'Sheet5.Cells(i + 1, 2) = vx
        'Sheet5.Cells(i + 1, 2 + j) = xids: 'Ids_f(vds, vgs, vgd)
        Sheet5.Cells(i + 1, 2 + j) = idrainx: 'Ids_f(vds, vgs, vgd)
        
        
    Next i
    
Next j



ixmax = istep
jymax = isweep

vdsmin = 0.1 * VDDMAX
vdsmax = VDDMAX
dvds = (vdsmax - vdsmin) / (ixmax - 1)

vgsmin = VGGMIN
vgsmax = VGGMAX
dvgs = (vgsmax - vgsmin) / (jymax - 1)


For i = 1 To ixmax
    VDD = dvds * CDbl(i - 1) + vdsmin
    Sheet5.Cells(1, 14 + i) = VDD
    For j = 1 To jymax

        VGG = dvgs * CDbl(j - 1) + vgsmin
        
        Call MultiNewton(level, vd, vg, vs, idrainx, igatex, VDD, VGG, iter)

        vx = vd - vs
        vy = vg - vs

        'xids = Ids_f(vx, vy, vz) - Igd_f(vx, vy, vz)
        'xids = Ids(level, vx, vy, vz) - Igd(level, vx, vy, vz)
        'xids2 = gm_f(vx, vy, vz) - ggd_f(vx, vy, vz)
        'xids = dfn7y("Idrain", level, vx, vy, vz, 0.001, 1)
        'xids = dfn7y("Igd", level, vx, vy, vz, 0.001, 1) - dfn7x("Igd", level, vx, vy, vz, 0.001, 1)
        'xids = Igd(level, vx, vy, vz)
        'xids = Ids0_f(vx, vy, vz)
        'xids = gds_f(vx, vy, vz)
        'xids2 = gm3_f(vx, vy, vz)
        'xids = Igd_f(vx, vy, vz)
        'xids2 = Igs_f(vx, vy, vz)
        'xids = Ivgst(vx, vy, vz)
        'xids2 = Idsgsmax(vx, vy, vz)
        'xids = Idsgs(vx, vy, vz)
        'xids2 = gvsat(vx, vy, vz)
        'xids = Ivdst(vx, vy, vz)
        'xids2 = Idsdsmax(vx, vy, vz)
        'xids = Idsds(vx, vy, vz)
        
        Sheet5.Cells(j + 1, 14) = VGG
        'Sheet5.Cells(j + 1, 14 + i) = xids: 'Ids_f(vds, vgs, vgd)
        Sheet5.Cells(j + 1, 14 + i) = idrainx: 'Ids_f(vds, vgs, vgd)
        'Sheet5.Cells(j + 1, 14 + i) = (Ids(level, vx, vy, vz) - Igd(level, vx, vy, vz)):  'Ids_f(vds, vgs, vgd)
        'Sheet5.Cells(j + 1, 14 + i) = vd: 'Ids_f(vds, vgs, vgd)
        

    Next j
Next i

End Sub

'=============================================================================
'
'                           FUNCTION cosh# (x)
'
'=============================================================================
Function cosh(x As Double) As Double
    cosh = (expm(x) + expm(-x)) / 2#
End Function

'=============================================================================
'
'                           FUNCTION sinh# (x)
'
'=============================================================================
Function sinh(x As Double) As Double
    sinh = (expm(x) + expm(-x)) / 2#
End Function

'=============================================================================
'
'                           FUNCTION expm# (x)
'
'=============================================================================
Function expm(x) As Double

        If x > 200# Then
            expm = Exp(200#) / (1# + Exp(200#) / 1E+18)
        Else
            expm = Exp(x) / (1# + Exp(x) / 1E+18)
        End If

End Function

'=============================================================================
'
'                           FUNCTION fermi (x)
'
'=============================================================================
Function fermi(x) As Double

        If x > 200# Then
            fermi = 1# / (1# + Exp(200#))
        Else
            fermi = 1# / (1# + Exp(x))
        End If

End Function


'=============================================================================
'
'                         FUNCTION Idsmodel (x, y)
'
'=============================================================================
Function Idsmodel(ByVal vdsx As Double, ByVal vgsx As Double, ByVal vgdx As Double) As Double
Dim vds As Double, vgs As Double, vgd As Double, x As Double


vds = vdsx
vgs = vgsx
vgd = vgdx


Dim vgst As Double

Select Case Ddevice$
    Case "fet"
        vgst = vgs - VTO + GAMMA * (vds) ^ NT
    
        If vgst <= 0# Then
            Idsmodel = 0# + vds * GMIN
        Else
            Idsmodel = BETA * (1# + LAMBDA * vds) * (vgst ^ NG) * tanh(ALPHA * vds) + vds * GMIN
        End If
    Case "bipolar"
        Idsmodel = BETA * (expm(ALPHA * vds / VTHERMAL / 80#) - 1#) * Abs(2# * vgs) * GAMMA + x * GMIN
End Select

End Function


Function Ids(level, vdsx As Double, vgsx As Double, vgdx As Double) As Double
Dim vds As Double, vgs As Double, vgd As Double

vds = vdsx
vgs = vgsx
vgd = vgdx
    Select Case level
        Case 1
            Ids = Ids_f(vds, vgs, vgd)
        Case 2
            Ids = Idsmodel(vds, vgs, vgd)
        Case 3
            Ids = EEHEMT1_Ids(vds, vgs, vgd)
        Case Else
    End Select

End Function

Function ggdgd(level As Integer, vdsx As Double, vgsx As Double, vgdx As Double) As Double
Dim vds As Double, vgs As Double, vgd As Double

    vds = vdsx
    vgs = vgsx
    vgd = vgdx

    ggdgd = dfn7z("Igd", level, vds, vgs, vgd, 0.001, 1)

End Function

Function ggdgs(level As Integer, vdsx As Double, vgsx As Double, vgdx As Double) As Double
Dim vds As Double, vgs As Double, vgd As Double

    vds = vdsx
    vgs = vgsx
    vgd = vgdx

    ggdgd = dfn7y("Igd", level, vds, vgs, vgd, 0.001, 1)

End Function

Function ggdds(level As Integer, vdsx As Double, vgsx As Double, vgdx As Double) As Double
Dim vds As Double, vgs As Double, vgd As Double

    vds = vdsx
    vgs = vgsx
    vgd = vgdx

    ggdgd = dfn7x("Igd", level, vds, vgs, vgd, 0.001, 1)

End Function

Function ggsgd(level As Integer, vdsx As Double, vgsx As Double, vgdx As Double) As Double
Dim vds As Double, vgs As Double, vgd As Double

    vds = vdsx
    vgs = vgsx
    vgd = vgdx

    ggdgd = dfn7z("Igs", level, vds, vgs, vgd, 0.001, 1)

End Function

Function ggsgs(level As Integer, vdsx As Double, vgsx As Double, vgdx As Double) As Double
Dim vds As Double, vgs As Double, vgd As Double

    vds = vdsx
    vgs = vgsx
    vgd = vgdx

    ggdgd = dfn7y("Igs", level, vds, vgs, vgd, 0.001, 1)

End Function

Function ggsds(level As Integer, vdsx As Double, vgsx As Double, vgdx As Double) As Double
Dim vds As Double, vgs As Double, vgd As Double

    vds = vdsx
    vgs = vgsx
    vgd = vgdx

    ggdgd = dfn7x("Igs", level, vds, vgs, vgd, 0.001, 1)

End Function

Function gdsgd(level As Integer, vdsx As Double, vgsx As Double, vgdx As Double) As Double
Dim vds As Double, vgs As Double, vgd As Double

    vds = vdsx
    vgs = vgsx
    vgd = vgdx

    gdsgd = dfn7z("Ids", level, vds, vgs, vgd, 0.001, 1)

End Function

Function gdsgs(level As Integer, vdsx As Double, vgsx As Double, vgdx As Double) As Double
Dim vds As Double, vgs As Double, vgd As Double

    vds = vdsx
    vgs = vgsx
    vgd = vgdx

    gdsgs = dfn7y("Ids", level, vds, vgs, vgd, 0.001, 1)

End Function

Function gdsds(level As Integer, vdsx As Double, vgsx As Double, vgdx As Double) As Double
Dim vds As Double, vgs As Double, vgd As Double

    vds = vdsx
    vgs = vgsx
    vgd = vgdx

    gdsds = dfn7x("Ids", level, vds, vgs, vgd, 0.001, 1)

End Function


'=============================================================================
'
'                         FUNCTION Igd_f (x, y)
'
'=============================================================================
Function Igd_f(vdsx As Double, vgsx As Double, vgdx As Double) As Double
Dim vds As Double, vgs As Double, vgd As Double, vcore As Double, icore As Double
vds = vdsx
vgs = vgsx
vgd = vgdx
    vcore = (vgd) / (ngd * VTHERMAL)
    icore = Isgd * (expm(vcore) - 1#)
    Igd_f = IGDMAX * icore / (IGDMAX + icore) + (vgd) * GMIN
    
End Function

Function ggd_f(vdsx As Double, vgsx As Double, vgdx As Double) As Double
Dim vds As Double, vgs As Double, vgd As Double
vds = vdsx
vgs = vgsx
vgd = vgdx
    
    vcore = vgd / ngd / VTHERMAL
    ggd_f = Isgd / ngd / VTHERMAL * expm(vcore) + GMIN

End Function

'=============================================================================
'
'                            FUNCTION Igs_f (x, y)
'
'=============================================================================
Function Igs_f(vdsx As Double, vgsx As Double, vgdx As Double) As Double
Dim vds As Double, vgs As Double, vgd As Double, vcore As Double, icore As Double
vds = vdsx
vgs = vgsx
vgd = vgdx

    vcore = vgs / (ngs * VTHERMAL)
    icore = Isgs * (expm(vcore) - 1#)
    Igs_f = icore * IGSMAX / (IGSMAX + icore) + vgs * GMIN
    
End Function

Function ggs_f(vdsx As Double, vgsx As Double, vgdx As Double) As Double
Dim vds As Double, vgs As Double, vgd As Double
vds = vdsx
vgs = vgsx
vgd = vgdx
    
    vcore = vgs / ngs / VTHERMAL
    ggs_f = Isgs / ngs / VTHERMAL * expm(vcore) + GMIN

End Function

Function log10(x As Double) As Double
    log10 = Log(x + 1E-300) / Log(10#)
End Function

Function Igd(level As Integer, vdsx As Double, vgsx As Double, vgdx As Double) As Double
Dim vds As Double, vgs As Double, vgd As Double
vds = vdsx
vgs = vgsx
vgd = vgdx

Select Case level
    Case 1
        Igd = Igd_f(vds, vgs, vgd)
    Case 2
        Igd = Igd_f(vds, vgs, vgd)
    Case 3
        Igd = EEHEMT1_Igd(vds, vgs, vgd)
    Case Else
End Select


End Function

Function Igs(level As Integer, vdsx As Double, vgsx As Double, vgdx As Double) As Double
Dim vds As Double, vgs As Double, vgd As Double
vds = vdsx
vgs = vgsx
vgd = vgdx

Select Case level
    Case 1
        Igs = Igs_f(vds, vgs, vgd)
    Case 2
        Igs = Igs_f(vds, vgs, vgd)
    Case 3
        Igs = EEHEMT1_Igs(vds, vgs, vgd)
    Case Else
End Select


End Function


'==================================matbsD=====================================
'matbsD% takes a matrix in LU form, found by matluD%, and a vector b
'and solves the system Ux=Lb for x. matrices A,b,x are double precision.
'
'Parameters: LU matrix in A, corresponding pivot vectors in rpvt and cpvt,
'            right side in b
'
'Returns: solution in x, b is modified, rest unchanged
'=============================================================================
Function matbsD%(a() As Double, b() As Double, x() As Double)
On Local Error GoTo dbserr: matbsD% = 0
'do row operations on b using the multipliers in L to find Lb

For pvt% = lo To (up - 1)
   c% = cpvt(pvt%)
   For Row% = (pvt% + 1) To up
      r% = rpvt(Row%)
      b(r%) = b(r%) + a(r%, c%) * b(rpvt(pvt%))
   Next Row%
Next pvt%
'backsolve Ux=Lb to find x
For Row% = up To lo Step -1
   c% = cpvt(Row%)
   r% = rpvt(Row%)
   x(c%) = b(r%)
   For col% = (Row% + 1) To up
      x(c%) = x(c%) - a(r%, cpvt(col%)) * x(cpvt(col%))
   Next col%
   x(c%) = x(c%) / a(r%, c%)
Next Row%
dbsexit:
Exit Function
dbserr:
   matbsD% = Err
   Resume dbsexit
End Function

'==================================matluD%====================================
'matluD% does Gaussian elimination with total pivoting to put a square, double
'precision matrix in LU form. The multipliers used in the row operations to
'create zeroes below the main diagonal are saved in the zero spaces.
'
'Parameters: A(n x n) matrix, rpvt(n) and cpvt(n) permutation vectors
'            used to index the row and column pivots
'
'Returns: A in LU form with corresponding pivot vectors; the total number of
'         pivots in count, which is used to find the sign of the determinant.
'=============================================================================
Function matluD%(a() As Double)
On Local Error GoTo dluerr: errcode% = 0
'Checks if A is square, returns error code if not
If Not (lo = LBound(a, 2) And up = UBound(a, 2)) Then Error 198
ReDim rownorm(lo To up) As Double
count = 0                            'initialize count, continue
continue = -1
For Row% = lo To up                  'initialize rpvt and cpvt
   rpvt(Row%) = Row%
   cpvt(Row%) = Row%
   rownorm(Row%) = 0#                'find the row norms of A()
   For col% = lo To up
      rownorm(Row%) = rownorm(Row%) + Abs(a(Row%, col%))
   Next col%
   If rownorm(Row%) = 0# Then        'if any rownorm is zero, the matrix
      continue = 0                   'is singular, set error, exit and
      Error 199                      'do not continue
   End If
Next Row%
For pvt% = lo To (up - 1)
'Find best available pivot
   Max# = 0#                         'checks all values in rows and columns not
   For Row% = pvt% To up             'already used for pivoting and saves the
      r% = rpvt(Row%)                'largest absolute number and its position
      For col% = pvt% To up
         c% = cpvt(col%)
         Temp# = Abs(a(r%, c%)) / rownorm(r%)
         If Temp# > Max# Then
            Max# = Temp#
            bestrow% = Row%          'save the position of new max#
            bestcol% = col%
         End If
      Next col%
   Next Row%
   If Max# = 0# Then                 'if no nonzero number is found, A is
      continue = 0                   'singular, send back error, do not continue
      Error 199
   ElseIf pvt% > 1 Then              'check if drop in pivots is too much
      If Max# < (deps# * oldmax#) Then errcode% = 199
   End If
   oldmax# = Max#
   If rpvt(pvt%) <> rpvt(bestrow%) Then
      count = count + 1                    'if a row or column pivot is
      SWAP rpvt(pvt%), rpvt(bestrow%)      'necessary, count it and permute
   End If                                  'rpvt or cpvt. Note: the rows and
   If cpvt(pvt%) <> cpvt(bestcol%) Then    'columns are not actually switched,
      count = count + 1                    'only the order in which they are
      SWAP cpvt(pvt%), cpvt(bestcol%)      'used.
   End If
'Eliminate all values below the pivot
   rp% = rpvt(pvt%)
   cp% = cpvt(pvt%)
   For Row% = (pvt% + 1) To up
      r% = rpvt(Row%)
      a(r%, cp%) = -a(r%, cp%) / a(rp%, cp%)  'save multipliers
      For col% = (pvt% + 1) To up
         c% = cpvt(col%)                      'complete row operations
         a(r%, c%) = a(r%, c%) + a(r%, cp%) * a(rp%, c%)
      Next col%
   Next Row%
Next pvt%
If a(rpvt(up), cpvt(up)) = 0# Then
   continue = 0                      'if last pivot is zero or pivot drop is
   Error 199                         'too large, A is singular, send back error
ElseIf (Abs(a(rpvt(up), cpvt(up))) / rownorm(rpvt(up))) < (deps# * oldmax#) Then
   errcode% = 199                    'if pivot is not identically zero then
End If                               'continue remains TRUE
If errcode% Then Error errcode%
dluexit:
matluD% = errcode%
Exit Function
dluerr:
   If errcode% < 199 Then continue = 0
   errcode% = Err
   Resume dluexit
End Function

'==================================MatSEqnD%==================================
'MatSEqnD% solves a system of n linear equations, Ax=b, and puts the
'answer in b. A is first put in LU form by matluD%, then matbsD is called
'to solve the system.  matrices A,b are double precision.
'
'Parameters: A(n x n) contains coefficient matrix, b(N) contains the right side
'
'Returns: A in LU form, solution in b
'=============================================================================
Function MatSEqnD%(a() As Double, x() As Double, b() As Double)
On Local Error GoTo dseqnerr: errcode% = 0
    lo = LBound(a, 1)
    up = UBound(a, 1)

    'ReDim x(lo To up) As Double
    'Dim x(up) As Double
    
    'ReDim rpvt(lo To up) As Integer
    'ReDim cpvt(lo To up) As Integer
    
    errcode% = matluD%(a())                      'Get LU matrix
    If Not continue Then Error errcode%
    
    'check dimensions of b
    If (lo <> LBound(b)) Or (up <> UBound(b)) Then Error 197
    
    bserrcode% = matbsD%(a(), b(), x())          'Backsolve system
    
    If bserrcode% Then Error bserrcode%
    
    'For Row% = lo To up
    '   b(Row%) = x(Row%)                         'Put solution in b for return
    'Next Row%
    
    If errcode% Then Error errcode%
dseqnexit:
    Erase rpvt, cpvt
    
    MatSEqnD% = errcode%
    
    Exit Function
    
dseqnerr:
       errcode% = (Err + 5) Mod 200 - 5
       Resume dseqnexit
End Function


Function maxf(x, y)
    If x > y Then
        maxf = x
    Else
        maxf = y
    End If
End Function

Function minf(x, y)
    If x > y Then
        minf = y
    Else
        minf = x
    End If
End Function

'=============================================================================
'
'                 MultiNewton (Multiple Newton Method)
'
'        This subroutine solve two sets of variables by using modified
'     Newton method. In this subroutine, functions which should be solved
'     are defined and is used MatSEqnD() subroutine.
'
'        Arguments: x, y, VDD, VGG, iter
'
'   * fi (i=1,2) are user defined equation which should be 0
'   * v(i) are displacement of voltages
'   * vold and vnew are vectors before and after calculation respectively
'   * gg(i,j) are derivatives of the f(i)
'   * MatSEqnD() is a subroutine to calculate equations
'   * cf is a coefficient which determine the next dv
'=============================================================================
Sub MultiNewton(level As Integer, vd As Double, vg As Double, vs As Double, idrainx As Double, igatex As Double, ByVal VDDx As Double, ByVal VGGx As Double, iter)
Dim v0(1 To 7) As Double, v(1 To 7) As Double, gg0(1 To 7, 1 To 7) As Double, gg(1 To 7, 1 To 7) As Double
Dim xsol(1 To 7) As Double, xsol_old(1 To 7) As Double, dxsol(1 To 7) As Double

Dim vds0 As Double, vgs0 As Double, vgd0 As Double
Dim vds0_old As Double, vgs0_old As Double

Dim drain_current As Double, gate_current As Double
Dim evalfunc1 As Double, evalfunc2 As Double, evalfunc3 As Double, ef(1 To 7) As Double

Dim VDD As Double, VGG As Double
Dim vold As Double, vnew As Double
Dim GGDxx As Double, GMxx As Double, GDSxx As Double, GGSxx As Double
Dim VGMxx0 As Double
Dim cfd As Double, cfg As Double, cff(1 To 7) As Double



iter = 0

VDD = VDDx
VGG = VGGx

xsol(1) = VDD
xsol(2) = VDD
xsol(3) = VGG
xsol(4) = VGG
xsol(5) = 0#
xsol(6) = 0.1
xsol(7) = 0#

cff0 = 0.25
ul = 1#
For i = 1 To 7
    dxsol(i) = 0#
    xsol_old(i) = 0#
    cff(i) = cff0
Next i

'If VDD < 0 Then
'    vds0 = 0# * milli
'Else
'    vds0 = VDD
'End If

vds0 = VDD
vgs0 = VGG
vgd0 = vgs0 - vds0

vds0_old = VDD
vgs0_old = VGG

vds0max = -10000#
vds0min = 10000#
vgs0max = -10000#
vgs0min = 10000#
VGMxx0 = VTH - 100000#

iretry = 1

Do
    If debugmode = 1 Then DoEvents
    
'====================================================================
' This routine is one good tip for improving convergence.
' You will obtaine accurate solutions. But, it is not perfect and
' little bit slow speed. itermax * 0.5 is arbital. You can also
' set itermax * 0.0. This means, you use this routine for all
' bias conditions.

  
'    If Abs(VDD) > 0.01 Then
        If vgs0 <> VGMxx0 Then
            If dfn7y("Ids", level, vds0, vgs0, vgd0, 0.001, 1) < Ids(level, vds0, vgs0, vgd0) / (vgs0 - VGMxx0) Then
                GMxx = Ids(level, vds0, vgs0, vgd0) / (vgs0 - VGMxx0)
            Else
                GMxx = dfn7y("Ids", level, vds0, vgs0, vgd0, 0.001, 1): 'gm_f(vds0, vgs0)
            End If
        Else
            GMxx = dfn7y("Ids", level, vds0, vgs0, vgd0, 0.001, 1): 'gm_f(vds0, vgs0)
        End If
        
        If vds0 <> 0# Then

            If dfn7x("Ids", level, vds0, vgs0, vgd0, 0.001, 1) < Ids(level, vds0, vgs0, vgd0) / vds0 Then
                GDSxx = Ids(level, vds0, vgs0, vgd0) / vds0
                IdsE = 0#
            Else
                GDSxx = dfn7x("Ids", level, vds0, vgs0, vgd0, 0.001, 1)
                'IdsE = Ids(level, vds0, vgs0) - (GDSxx * vds0 + GMxx * vgs0)
                IdsE = Ids(level, vds0, vgs0, vgd0) - (GDSxx * vds0 + GMxx * vgs0)
            End If

        Else
'    Else
            GMxx = dfn7y("Ids", level, vds0, vgs0, vgd0, 0.001, 1): 'gm_f(vds0, vgs0)
            GDSxx = dfn7x("Ids", level, vds0, vgs0, vgd0, 0.001, 1)
            IdsE = Ids(level, vds0, vgs0, vgd0) - (GDSxx * vds0 + GMxx * vgs0)
        End If
'    End If


'====================================================================


' Following conductances and equivalent currents give not enough convergence.
' But convergence speed is very high. Not recommend.

'        GDSX = gds_s(vds0, vgs0)
'        GMX = gm_s(vds0, vgs0)
'        IdsE = Ids_s(vds0, vgs0) - (GDSX * vds0 + GMX * vgs0)
    
    
    
    'IgdE = Igd(level, vds0, vgs0) - ggd_f(vds0, vgs0) * vgd0
    'IgsE = Igs(level, vds0, vgs0) - ggs_f(vds0, vgs0) * vgs0
    'GGDx = dfn7y("Igd", level, vds0, vgs0, 0.001, 1) - dfn7x("Igd", level, vds0, vgs0, 0.001, 1)
    GGDxx = ggd(level, vds0, vgs0, vgd0)
    GGSxx = dfn7y("Igs", level, vds0, vgs0, vgd0, 0.001, 1)
    IgdE = Igd(level, vds0, vgs0, vgd0) - (dfn7z("Igd", level, vds0, vgs0, vgd0, 0.001, 1) * vgd0)
    IgsE = Igs(level, vds0, vgs0, vgd0) - GGSxx * vgs0
    
    
    v(1) = 0#
    v(2) = IgdE - IdsE
    v(3) = 0#
    v(4) = (-IgdE - IgsE)
    v(5) = IgsE + IdsE
    v(6) = VDD
    v(7) = VGG
    
    gg(1, 1) = 1# / Rd
    gg(1, 2) = -1# / Rd
    gg(1, 3) = 0#
    gg(1, 4) = 0#
    gg(1, 5) = 0#
    gg(1, 6) = -1#
    gg(1, 7) = 0#

    gg(2, 1) = -1# / Rd
    'gg(2, 2) = 1# / Rd + gdsx + ggd_f(vds0, vgs0) + GMIN
    gg(2, 2) = 1# / Rd + GDSxx + GGDxx + GMIN
    gg(2, 3) = 0#
    'gg(2, 4) = -ggd_f(vds0, vgs0) - gmx
    gg(2, 4) = -GGDxx + GMxx
    gg(2, 5) = -GDSxx - GMxx - GMIN
    gg(2, 6) = 0#
    gg(2, 7) = 0#

    gg(3, 1) = 0#
    gg(3, 2) = 0#
    gg(3, 3) = 1# / Rg
    gg(3, 4) = -1# / Rg
    gg(3, 5) = 0#
    gg(3, 6) = 0#
    gg(3, 7) = -1#

    gg(4, 1) = 0#
    'gg(4, 2) = -ggd_f(vds0, vgs0)
    gg(4, 2) = -GGDxx
    gg(4, 3) = -1# / Rg
    'gg(4, 4) = 1# / Rg + ggd_f(vds0, vgs0) + ggs_f(vds0, vgs0)
    gg(4, 4) = 1# / Rg + GGDxx + GGSxx
    'gg(4, 5) = -ggs_f(vds0, vgs0)
    gg(4, 5) = -GGSxx
    gg(4, 6) = 0#
    gg(4, 7) = 0#

    gg(5, 1) = 0#
    gg(5, 2) = -GDSxx - GMIN
    gg(5, 3) = 0#
    'gg(5, 4) = -ggs_f(vds0, vgs0) - gmx
    gg(5, 4) = -GGSxx - GMxx
    'gg(5, 5) = 1# / Rs + gdsx + ggs_f(vds0, vgs0) + gmx + GMIN
    gg(5, 5) = 1# / Rs + GDSxx + GGSxx + GMxx + GMIN
    gg(5, 6) = 0#
    gg(5, 7) = 0#

    gg(6, 1) = 1#
    gg(6, 2) = 0#
    gg(6, 3) = 0#
    gg(6, 4) = 0#
    gg(6, 5) = 0#
    gg(6, 6) = 0#
    gg(6, 7) = 0#

    gg(7, 1) = 0#
    gg(7, 2) = 0#
    gg(7, 3) = 1#
    gg(7, 4) = 0#
    gg(7, 5) = 0#
    gg(7, 6) = 0#
    gg(7, 7) = 0#
   
  
'    'v()をv0()へコピー
'    For i = 1 To 7
'        v0(i) = v(i)
'    Next i
   
    'v()をv0()へコピー
    'gg()をgg0()へコピー
    'error by old solution is generated->vold ( dxsol )
    vold = 0#
    For i = 1 To 7
        v0(i) = v(i)
        dxsol(i) = Abs(xsol(i) - xsol_old(i))
        vold = vold + dxsol(i)
        xsol_old(i) = xsol(i)
        
        For j = 1 To 7
            gg0(i, j) = gg(i, j)
        Next j
    Next i

'    vold = 0#
'    For i = 1 To 7
'        dxsol(i) = (xsol(i) - xsol_old(i))
'        vold = vold + dxsol(i)
'    Next i
'
'    For i = 1 To 7
'        xsol_old(i) = xsol(i)
'    Next i
    


    'MatSEqnDを使うと、gg(), v()の中身が変化してしまうので、gg0(), v0()へそれぞれコピーしておく
    'MatSEqnDのggとvをByVal宣言すればこの問題は解決する
    Call MatSEqnD%(gg(), xsol(), v())
    
    'cff array is coeficients for improve convergence. Very important parameters.
    'When gm of transistor is high, cfg must be small values such as 0.1.
    'These parameters will restrict big change of solutions.
    'cfd and cfg will be determind adaptively depending on variation of solutions.
    'Following routine is not adaptive form. Sorry.
    
    vnew = 0#
    For i = 1 To 7
        dxsol(i) = (xsol(i) - xsol_old(i)): ' differences are not modified here.
        vnew = vnew + Abs(dxsol(i))
        xsol(i) = (1# - cff(i)) * xsol_old(i) + cff(i) * xsol(i): 'solution xsol() are updated
    Next i
      
    vds0_old = vds0
    vgs0_old = vgs0
    
    vds0 = xsol(2) - xsol(5)
    vgs0 = xsol(4) - xsol(5)
    vgd0 = vgs0 - vds0
    
    
'    vgs0 = VGG
'    If iter > itermax / 2 Then
'        vds0max = maxf(vds0, vds0max)
'        vds0min = minf(vds0, vds0min)
'        vgs0max = maxf(vgs0, vgs0max)
'        vgs0min = minf(vgs0, vgs0min)
'        'Debug.Print "VDD:"; VDD; " "; "vds0max:"; vds0max; " "; "vds0min:"; vds0min; " "; "VGG:"; VGG; " "; "vgs0max:"; vgs0max; " "; "vgs0min:"; vgs0min
'        vds0 = 0 * vds0max + 1 * (vds0max + vds0min) / 2#
'        vgs0 = 0 * vgs0max + 1 * VGG
'    End If
        
    
    drain_current = xsol(6)
    gate_current = xsol(7)
    
    evalfunc2 = 0
    For i = 1 To 7
        ef(i) = gg0(i, 1) * xsol(1) + gg0(i, 2) * xsol(2) + gg0(i, 3) * xsol(3) + gg0(i, 4) * xsol(4) + gg0(i, 5) * xsol(5) + gg0(i, 6) * xsol(6) + gg0(i, 7) * xsol(7) - v0(i)
        evalfunc2 = evalfunc2 + (ef(i)) ^ 2
    Next i

'    ef1 = gg0(1, 1) * xsol(1) + gg0(1, 2) * xsol(2) + gg0(1, 3) * xsol(3) + gg0(1, 4) * xsol(4) + gg0(1, 5) * xsol(5) + gg0(1, 6) * xsol(6) + gg0(1, 7) * xsol(7) - v0(1)
'    ef2 = gg0(2, 1) * xsol(1) + gg0(2, 2) * xsol(2) + gg0(2, 3) * xsol(3) + gg0(2, 4) * xsol(4) + gg0(2, 5) * xsol(5) + gg0(2, 6) * xsol(6) + gg0(2, 7) * xsol(7) - v0(2)
'    ef3 = gg0(3, 1) * xsol(1) + gg0(3, 2) * xsol(2) + gg0(3, 3) * xsol(3) + gg0(3, 4) * xsol(4) + gg0(3, 5) * xsol(5) + gg0(3, 6) * xsol(6) + gg0(3, 7) * xsol(7) - v0(3)
'    ef4 = gg0(4, 1) * xsol(1) + gg0(4, 2) * xsol(2) + gg0(4, 3) * xsol(3) + gg0(4, 4) * xsol(4) + gg0(4, 5) * xsol(5) + gg0(4, 6) * xsol(6) + gg0(4, 7) * xsol(7) - v0(4)
'    ef5 = gg0(5, 1) * xsol(1) + gg0(5, 2) * xsol(2) + gg0(5, 3) * xsol(3) + gg0(5, 4) * xsol(4) + gg0(5, 5) * xsol(5) + gg0(5, 6) * xsol(6) + gg0(5, 7) * xsol(7) - v0(5)
'    ef6 = gg0(6, 1) * xsol(1) + gg0(6, 2) * xsol(2) + gg0(6, 3) * xsol(3) + gg0(6, 4) * xsol(4) + gg0(6, 5) * xsol(5) + gg0(6, 6) * xsol(6) + gg0(6, 7) * xsol(7) - v0(6)
'    ef7 = gg0(7, 1) * xsol(1) + gg0(7, 2) * xsol(2) + gg0(7, 3) * xsol(3) + gg0(7, 4) * xsol(4) + gg0(7, 5) * xsol(5) + gg0(7, 6) * xsol(6) + gg0(7, 7) * xsol(7) - v0(7)

    evalfunc1 = Abs((vnew) ^ 2 - (vold) ^ 2)
    
'    evalfunc2 = Sqr((ef1) ^ 2 + (ef2) ^ 2 + (ef3) ^ 2 + (ef4) ^ 2 + (ef5) ^ 2 + (ef6) ^ 2 + (ef7) ^ 2)
    
    'Debug.Print v0(1), v0(2), v0(3), v0(4), v0(5), v0(6)
'    For iii = 1 To 7
'        Debug.Print xsol(iii);
'    Next iii
'    Debug.Print ""
    'Debug.Print "iter;"; iter; "evalf:"; evalfunc; " vds0:"; vds0; " Ids_f:"; Ids_f(vds0, vgs0), " vgs0:"; vgs0, " Igs_f:"; Igs_f(vds0, vgs0), " Igd_f:"; Igd_f(vds0, vgs0), " Idrain:"; Idrain, " Igate:"; Igate

    If iter > itermax * 0 Then
        Sheet2.Cells(iter + 1, 1) = iter
        Sheet2.Cells(iter + 1, itercount + 1) = evalfunc1 + evalfunc2
        Sheet3.Cells(iter + 1, 1) = iter
        Sheet3.Cells(1, 2 * (itercount - 1) + 2) = VGG
        'Sheet3.Cells(1, 2 * (itercount - 1) + 3) = vgs0
        Sheet3.Cells(iter + 1, 2 * (itercount - 1) + 2) = vgs0
        'Sheet3.Cells(iter + 1, 3 * (itercount - 1) + 3) = VDD
        Sheet3.Cells(iter + 1, 2 * (itercount - 1) + 3) = vds0
    End If

    'Debug.Print "iter:"; iter, "vold:"; vold, " vnew:"; vnew

    'If vds0 < VDD Then
        If evalfunc1 + evalfunc2 < tol Then
            vx = vds0
            vy = vgs0
            vd = xsol(2)
            vg = xsol(4)
            vs = xsol(5)
            idrainx = xsol(6)
            igatex = xsol(7)
            'Debug.Print evalfunc; " "; xsol(1); " "; xsol(2); " "; xsol(3); " "; xsol(4); " "; xsol(5); " "; xsol(6); " "; xsol(7)
            Exit Do
        End If
    'End If

    If iter > itermax Then
        Debug.Print "Error! @VDD:"; VDD; " VGG:"; VGG
        If iretry > 100 Then
            Debug.Print "No convergence!!"
            Exit Do
        Else
            ul = ul * 1.5
            For i = 1 To 7
                cff(i) = cff(i) / ul
            Next i
            Debug.Print "Rety now("; iretry; ") "; " cff0:"; cff0 / ul
            iretry = iretry + 1
            itermax = 1.5 * itermax
            iter = 0
        End If
    End If
    
    

    iter = iter + 1

Loop

itercount = itercount + 1

Erase v0, v, gg0, gg, xsol, xsol_old, dxsol


End Sub

'=============================================================================
'
'                            FUNCTION tanh# (x)
'
'=============================================================================
Function tanh(x As Double) As Double
    tanh = (expm(x) - expm(-x)) / (expm(x) + expm(-x)) + GMIN * x
End Function


'================ SWAP (a, b) =====================
Function SWAP(a, b)

a0 = a
a = b
b = a0

End Function

'=============================================================================
'
'                            Matrix test
'
'=============================================================================
Sub MatrixTest()
Dim v(1 To 4) As Double, gg(1 To 4, 1 To 4) As Double, xsol(1 To 4) As Double


    gg(1, 1) = 1#
    gg(1, 2) = -1#
    gg(1, 3) = -2#
    gg(1, 4) = 2#
    
    gg(2, 1) = 2#
    gg(2, 2) = -1#
    gg(2, 3) = -3#
    gg(2, 4) = 3#
    
    gg(3, 1) = -1#
    gg(3, 2) = 3#
    gg(3, 3) = 3#
    gg(3, 4) = -2#
    
    gg(4, 1) = 1#
    gg(4, 2) = 2#
    gg(4, 3) = 0#
    gg(4, 4) = -1#
    
    v(1) = 5#
    v(2) = 10#
    v(3) = 2#
    v(4) = -10#
    
    iret% = MatSEqnD%(gg(), xsol(), v())
    
    For i = 1 To 4
        Debug.Print "xsol("; i; ")="; xsol(i)
    Next i
    
    
End Sub

Sub NumericalDerivative()

Dim x As Double, y As Double, h As Double, h0 As Double

Call parameters_set

name$ = "Ids"
level = 2

x = 10#
y = 0.5

    For i = 1 To 8
    
        h = (10) ^ (-i)
        'Debug.Print i; " "; h; " "; (dfnCx(x, y, h, 4) - fn4(x)) / fn4(x) * 100#; " "; (dfn5x(x, y, h, 4) - fn4(x)) / fn4(x) * 100#; " "; (dfn7x(x, y, h, 4) - fn4(x)) / fn4(x) * 100#
        Debug.Print i; " "; h; " "; dfnCy(name$, level, x, y, z, h, 1); " "; dfn5y(name$, level, x, y, z, h, 1); " "; dfn7y(name$, level, x, y, z, h, 1)
    Next i


End Sub


Function fn(name$, level As Integer, vdsx As Double, vgsx As Double, vgdx As Double) As Double
Dim vds As Double, vgs As Double, vgd As Double, msg$
vds = vdsx
vgs = vgsx
vgd = vgdx

    Select Case name$
        Case "Idrain"
            fn = Ids(level, vds, vgs, vgd) - Igd(level, vds, vgs, vgd)
        Case "Ids"
            fn = Ids(level, vds, vgs, vgd)
        Case "Igs"
            fn = Igs(level, vds, vgs, vgd)
        Case "Igd"
            fn = Igd(level, vds, vgs, vgd)
        Case "Igate"
            fn = Igs(level, vds, vgs, vgd) + Igd(level, vds, vgs, vgd)
        Case "exp", "Exp", "EXP"
            fn = expm(vds)
        Case "sin", "Sin", "SIN"
            fn = sin(vds)
        Case "sinc", "Sinc", "SINC"
            fn = sinc(vds)
        Case "sec", "Sec", "SEC"
            fn = sec(vds)
        Case "cos", "Cos", "COS"
            fn = Cos(vds)
        Case "cot", "Cot", "COT"
            fn = cot(vds)
        Case "cosec", "Cosec", "COSEC"
            fn = cosec(vds)
        Case "tan", "Tan", "TAN"
            fn = tan(vds)
        Case "sinh", "Sinh", "SINH"
            fn = sinh(vds)
        Case "cosh", "Cosh", "COSH"
            fn = cosh(vds)
        Case "tanh", "Tanh", "TANH"
            fn = tanh(vds)
        Case "log10", "Log10", "LOG10"
            fn = log10(vds)
        Case "log", "Log", "LOG"
            fn = Log(vds)
        Case "fermi", "Fermi", "FERMI"
            fn = fermi(vds)
        Case "Ids_f"
            fn = Ids_f(vds, vgs, vgd)
        Case "gds_f"
            fn = gds_f(vds, vgs, vgd)
        Case "gm_f"
            fn = gm_f(vds, vgs, vgd)
        Case "gm2_f"
            fn = gm2_f(vds, vgs, vgd)
        Case "gm3_f"
            fn = gm3_f(vds, vgs, vgd)
        Case "gm4_f"
            fn = gm4_f(vds, vgs, vgd)
        Case "gm5_f"
            fn = gm5_f(vds, vgs, vgd)
        Case "Ids_mod"
            fn = Ids_mod(vds, vgs, vgd)
        Case "Ivgst"
            fn = Ivgst(vds, vgs, vgd)
        Case "Idsgsmax"
            fn = Idsgsmax(vds, vgs, vgd)
        Case "Idsgs"
            fn = Idsgs(vds, vgs, vgd)
        Case "gvsat"
            fn = gvsat(vds, vgs, vgd)
        Case "Ivdst"
            fn = Ivdst(vds, vgs, vgd)
        Case "Idsdsmax"
            fn = Idsdsmax(vds, vgs, vgd)
        Case "Idsds"
            fn = Idsds(vds, vgs, vgd)
        Case "Ids_org"
            fn = Ids_org(vds, vgs, vgd)
        Case "Ids_s"
            fn = Ids_s(vds, vgs, vgd)
        Case "Ids0_s"
            fn = Ids0_s(vds, vgs, vgd)
        Case "gds_s"
            fn = gds_s(vds, vgs, vgd)
        Case "gm_s"
            fn = gm_s(vds, vgs, vgd)
        Case "Idsmodel"
            fn = Idsmodel(vds, vgs, vgd)
        Case "ggdgd"
            fn = ggdgd(level, vds, vgs, vgd)
        Case "ggdgs"
            fn = ggdgs(level, vds, vgs, vgd)
        Case "ggdds"
            fn = ggdds(level, vds, vgs, vgd)
        Case "ggsgd"
            fn = ggsgd(level, vds, vgs, vgd)
        Case "ggsgs"
            fn = ggsgs(level, vds, vgs, vgd)
        Case "ggsds"
            fn = ggsds(level, vds, vgs, vgd)
        Case "gdsgd"
            fn = gdsgd(level, vds, vgs, vgd)
        Case "gdsgs"
            fn = gdsgs(level, vds, vgs, vgd)
        Case "gdsds"
            fn = gdsds(level, vds, vgs, vgd)
        Case "Igd_f"
            fn = Igd_f(vds, vgs, vgd)
        Case "ggd_f"
            fn = ggd_f(vds, vgs, vgd)
        Case "Igs_f"
            fn = Igs_f(vds, vgs, vgd)
        Case "ggs_f"
            fn = ggs_f(vds, vgs, vgd)
        Case "EEHEMT1_Ids"
            fn = EEHEMT1_Ids(vds, vgs, vgd)
        Case "EEHEMT1_Idso"
            fn = EEHEMT1_Idso(vds, vgs, vgd)
        Case "EEHEMT1_Idsocomp"
            fn = EEHEMT1_Idsocomp(vds, vgs, vgd)
        Case "EEHEMT1_gdsocomp"
            fn = EEHEMT1_gdsocomp(vds, vgs, vgd)
        Case "EEHEMT1_gmocomp"
            fn = EEHEMT1_gmocomp(vds, vgs, vgd)
        Case "EEHEMT1_gm"
            fn = EEHEMT1_gm(vds, vgs, vgd)
        Case "EEHEMT1_gds"
            fn = EEHEMT1_gds(vds, vgs, vgd)
        Case "EEHEMT1_gmv"
            fn = EEHEMT1_gmv(vds, vgs)
        Case "EEHEMT1_gmm"
            fn = EEHEMT1_gmm(vds, vgs)
        Case "EEHEMT1_gdsv"
            fn = EEHEMT1_gdsv(vds, vgs)
        Case "EEHEMT1_Idsm"
            fn = EEHEMT1_Idsm(vds, vgs)
        Case "EEHEMT1_Idsv"
            fn = EEHEMT1_Idsv(vds, vgs)
        Case "EEHEMT1_Igd"
            fn = EEHEMT1_Igd(vds, vgs, vgd)
        Case "EEHEMT1_Igs"
            fn = EEHEMT1_Igs(vds, vgs, vgd)
        Case "EEHEMT1_Vx"
            fn = EEHEMT1_Vx(vgs, vds)
        Case "ftest"
            fn = ftest(vds, vgs)
        Case Else
            msg$ = "No function was selected." + Chr$(10) + "Program terminated."
            MsgBox (msg$)
            Debug.Print "No function was selected"
            End
    End Select
    
End Function


'==============================================================
'                Numerical derivation function
'                       h: dx, n: order
'                      center differential
'==============================================================

Function dfnCx(name$, level As Integer, x As Double, y As Double, z As Double, h As Double, n As Integer) As Double
Dim i As Integer
    For i = 0 To n
        If i = 0 Then
            dfnCx = fn(name$, level, x, y, z)
        ElseIf i = 1 Then
            dfnCx = (fn(name$, level, x + h, y, z) - fn(name$, level, x - h, y, z)) / (2# * h)
        Else
            dfnCx = (dfnCx(name$, level, x + h, y, z, h, i - 1) - dfnCx(name$, level, x - h, y, z, h, i - 1)) / (2# * h)
        End If
    Next i
    
End Function

Function dfnCy(name$, level As Integer, x As Double, y As Double, z As Double, h As Double, n As Integer) As Double
Dim i As Integer
    For i = 0 To n
        If i = 0 Then
            dfnCy = fn(name$, level, x, y, z)
        ElseIf i = 1 Then
            dfnCy = (fn(name$, level, x, y + h, z) - fn(name$, level, x, y - h, z)) / (2# * h)
        Else
            dfnCy = (dfnCy(name$, level, x, y + h, z, h, i - 1) - dfnCy(name$, level, x, y - h, z, h, i - 1)) / (2# * h)
        End If
    Next i
    
End Function

Function dfnCz(name$, level As Integer, x As Double, y As Double, z As Double, h As Double, n As Integer) As Double
Dim i As Integer
    For i = 0 To n
        If i = 0 Then
            dfnCz = fn(name$, level, x, y, z)
        ElseIf i = 1 Then
            dfnCz = (fn(name$, level, x, y, z + h) - fn(name$, level, x, y, z - h)) / (2# * h)
        Else
            dfnCz = (dfnCz(name$, level, x, y, z + h, h, i - 1) - dfnCz(name$, level, x, y, z - h, h, i - 1)) / (2# * h)
        End If
    Next i
    
End Function


'==============================================================
'                Numerical derivation function
'                       h: dx, n: order
'                     5 points differential
'==============================================================

Function dfn5x(name$, level As Integer, x As Double, y As Double, z As Double, h As Double, n As Integer) As Double
Dim i As Integer
    For i = 0 To n
        If i = 0 Then
            dfn5x = fn(name$, level, x, y, z)
        ElseIf i = 1 Then
            dfn5x = (fn(name$, level, x - 2# * h, y, z) - 8# * fn(name$, level, x - h, y, z) + 8# * fn(name$, level, x + h, y, z) - fn(name$, level, x + 2# * h, y, z)) / (12# * h)
        Else
            dfn5x = (dfn5x(name$, level, x - 2# * h, y, z, h, i - 1) - 8# * dfn5x(name$, level, x - h, y, z, h, i - 1) + 8# * dfn5x(name$, level, x + h, y, z, h, i - 1) - dfn5x(name$, level, x + 2# * h, y, z, h, i - 1)) / (12# * h)
        End If
    Next i
    
End Function

Function dfn5y(name$, level As Integer, x As Double, y As Double, z As Double, h As Double, n As Integer) As Double
Dim i As Integer
    For i = 0 To n
        If i = 0 Then
            dfn5y = fn(name$, level, x, y, z)
        ElseIf i = 1 Then
            dfn5y = (fn(name$, level, x, y - 2# * h, z) - 8# * fn(name$, level, x, y - h, z) + 8# * fn(name$, level, x, y + h, z) - fn(name$, level, x, y + 2# * h, z)) / (12# * h)
        Else
            dfn5y = (dfn5y(name$, level, x, y - 2# * h, z, h, i - 1) - 8# * dfn5y(name$, level, x, y - h, z, h, i - 1) + 8# * dfn5y(name$, level, x, y + h, z, h, i - 1) - dfn5y(name$, level, x, y + 2# * h, z, h, i - 1)) / (12# * h)
        End If
    Next i
    
End Function

Function dfn5z(name$, level As Integer, x As Double, y As Double, h As Double, n As Integer) As Double
Dim i As Integer
    For i = 0 To n
        If i = 0 Then
            dfn5z = fn(name$, level, x, y, z)
        ElseIf i = 1 Then
            dfn5z = (fn(name$, level, x, y, z - 2# * h) - 8# * fn(name$, level, x, y, z - h) + 8# * fn(name$, level, x, y, z + h) - fn(name$, level, x, y, z + 2# * h)) / (12# * h)
        Else
            dfn5z = (dfn5z(name$, level, x, y, z - 2# * h, h, i - 1) - 8# * dfn5z(name$, level, x, y, z - h, h, i - 1) + 8# * dfn5z(name$, level, x, y, z + h, h, i - 1) - dfn5z(name$, level, x, y, z + 2# * h, h, i - 1)) / (12# * h)
        End If
    Next i
    
End Function


'==============================================================
'                Numerical derivation function
'                       h: dx, n: order
'                     7 points differential
'==============================================================

Function dfn7x(name$, level As Integer, x As Double, y As Double, z As Double, h As Double, n As Integer) As Double
Dim i As Integer
    For i = 0 To n
        If i = 0 Then
            dfn7x = fn(name$, level, x, y, z)
        ElseIf i = 1 Then
            dfn7x = (fn(name$, level, x + 3# * h, y, z) - 9# * fn(name$, level, x + 2# * h, y, z) + 45# * fn(name$, level, x + h, y, z) - 45# * fn(name$, level, x - h, y, z) + 9# * fn(name$, level, x - 2# * h, y, z) - fn(name$, level, x - 3# * h, y, z)) / (60# * h)
        Else
            dfn7x = (dfn7x(name$, level, x + 3# * h, y, z, h, i - 1) - 9# * dfn7x(name$, level, x + 2# * h, y, z, h, i - 1) + 45# * dfn7x(name$, level, x + h, y, z, h, i - 1) - 45# * dfn7x(name$, level, x - h, y, z, h, i - 1) + 9# * dfn7x(name$, level, x - 2# * h, y, z, h, i - 1) - dfn7x(name$, level, x - 3# * h, y, z, h, i - 1)) / (60# * h)
        End If
    Next i
    
End Function

Function dfn7y(name$, level As Integer, x As Double, y As Double, z As Double, h As Double, n As Integer) As Double
Dim i As Integer
    For i = 0 To n
        If i = 0 Then
            dfn7y = fn(name$, level, x, y, z)
        ElseIf i = 1 Then
            dfn7y = (fn(name$, level, x, y + 3# * h, z) - 9# * fn(name$, level, x, y + 2# * h, z) + 45# * fn(name$, level, x, y + h, z) - 45# * fn(name$, level, x, y - h, z) + 9# * fn(name$, level, x, y - 2# * h, z) - fn(name$, level, x, y - 3# * h, z)) / (60# * h)
        Else
            dfn7y = (dfn7y(name$, level, x, y + 3# * h, z, h, i - 1) - 9# * dfn7y(name$, level, x, y + 2# * h, z, h, i - 1) + 45# * dfn7y(name$, level, x, y + h, z, h, i - 1) - 45# * dfn7y(name$, level, x, y - h, z, h, i - 1) + 9# * dfn7y(name$, level, x, y - 2# * h, z, h, i - 1) - dfn7y(name$, level, x, y - 3# * h, z, h, i - 1)) / (60# * h)
        End If
    Next i
    
End Function

Function dfn7z(name$, level As Integer, x As Double, y As Double, z As Double, h As Double, n As Integer) As Double
Dim i As Integer
    For i = 0 To n
        If i = 0 Then
            dfn7z = fn(name$, level, x, y, z)
        ElseIf i = 1 Then
            dfn7z = (fn(name$, level, x, y, z + 3# * h) - 9# * fn(name$, level, x, y, z + 2# * h) + 45# * fn(name$, level, x, y, z + h) - 45# * fn(name$, level, x, y, z - h) + 9# * fn(name$, level, x, y, z - 2# * h) - fn(name$, level, x, y, z - 3# * h)) / (60# * h)
        Else
            dfn7z = (dfn7z(name$, level, x, y, z + 3# * h, h, i - 1) - 9# * dfn7z(name$, level, x, y, z + 2# * h, h, i - 1) + 45# * dfn7z(name$, level, x, y, z + h, h, i - 1) - 45# * dfn7z(name$, level, x, y, z - h, h, i - 1) + 9# * dfn7z(name$, level, x, y, z - 2# * h, h, i - 1) - dfn7z(name$, level, x, y, z - 3# * h, h, i - 1)) / (60# * h)
        End If
    Next i
    
End Function


'=======================================================
'                     EEHEMT1 model
'=======================================================
Sub EEHEMT1(ByVal vdsx As Double, ByVal vgsx As Double, vgdx As Double, Ids_eeh As Double, gds_eeh As Double, gm_eeh As Double, Idso_eeh As Double, Idsocomp_eeh As Double, gdsocomp_eeh As Double, gmocomp_eeh As Double)
Dim vds As Double, vgs As Double, vgd As Double
Dim a_eeh As Double, b_eeh As Double, bp1_eeh As Double, gammvdso_eeh As Double, gdsvvb_eeh As Double
Dim gdsx_eeh As Double, gmo_eeh As Double, gmoff_eeh As Double, gmvb_gmoff_eeh As Double
Dim gmvvb_eeh As Double, gmx_eeh As Double, Idsvvb_eeh As Double
Dim Idsx_eeh As Double, Ikt_eeh As Double, kpvds_eeh As Double, pdis_eeh As Double, pdissq_eeh As Double, svb_eeh As Double
Dim thvst_eeh As Double, va_eeh As Double, vb_eeh As Double, vc_eeh As Double, vg_eeh As Double, vgs_va_eeh As Double, vgs_vb_eeh As Double
Dim vsat3_eeh As Double, vt_eeh As Double, vx_eeh As Double, vxgctg_eeh As Double, wg_ugw_eeh As Double, wuikt_eeh As Double
Dim gdso_eeh As Double, msg$
            
            vds = vdsx
            vgs = vgsx
            vgd = vgdx
            
'            If vds < 0 Then
'                vgs = vgdx
'                vgd = vgsx
'            End If
            
            Vkt = Boltz * Temperature / Charge
            
            gammvdso_eeh = GAMMA_EE * (VDSO_EE - (vds)) + 1#
            
            vg_eeh = (VGO_EE - VCH_EE) / (gammvdso_eeh) + VCH_EE
            vt_eeh = (VTO_EE - VCH_EE) / (gammvdso_eeh) + VCH_EE
            vx_eeh = (vgs - VCH_EE) * (gammvdso_eeh)

            If vgs > vg_eeh Then
            
                gmo_eeh = GMMAX_EE * (gammvdso_eeh)
                Idso_eeh = GMMAX_EE * (vx_eeh - (VGO_EE + VTO_EE) / 2# + VCH_EE)
                gdso_eeh = -GMMAX_EE * GAMMA_EE * (vgs - VCH_EE)
                
            ElseIf vgs <= vt_eeh Then
            
                gmo_eeh = 0#
                Idso_eeh = 0#
                gdso_eeh = 0#
                
            ElseIf vgs > vt_eeh And vgs <= vg_eeh Then
            
                gmo_eeh = EEHEMT1_gmm(vgs, vds)
                Idso_eeh = EEHEMT1_Idsm(vgs, vds)
                vxgctg_eeh = (vx_eeh - (VGO_EE - VCH_EE)) / (VTO_EE - VGO_EE)
                gdso_eeh = -GMMAX_EE / 2# * GAMMA_EE * (vgs - VCH_EE) * (Cos(PI * vxgctg_eeh) + 1#)
            Else
                msg$ = "No condition"
                MsgBox (msg$)
            End If

            vc_eeh = VCO_EE + MU_EE * (VDSO_EE - vds)
            vb_eeh = VBC_EE + vc_eeh
            va_eeh = vb_eeh - VBA_EE

            gmoff_eeh = EEHEMT1_gmm(VCO_EE, VDSO_EE)
            gdsvvb_eeh = EEHEMT1_gdsv(vb_eeh, vds)
            Idsvvb_eeh = EEHEMT1_Idsv(vb_eeh, vds)
            gmvvb_eeh = EEHEMT1_gmv(vb_eeh, vds)

            If vgs > vc_eeh Then

                svb_eeh = DELTGM_EE * VBC_EE / Sqr(ALPHA_EE * ALPHA_EE + VBC_EE * VBC_EE)
                gmvb_gmoff_eeh = gmvvb_eeh - gmoff_eeh
                b_eeh = svb_eeh * VBA_EE / gmvb_gmoff_eeh
                bp1_eeh = b_eeh + 1#
                a_eeh = gmvb_gmoff_eeh / (VBA_EE) ^ b_eeh

                If vgs < vb_eeh Then
                
                    gmocomp_eeh = gmo_eeh - EEHEMT1_gmv(vgs, vds)
                    Idsocomp_eeh = Idso_eeh - EEHEMT1_Idsv(vgs, vds)
                    gdsocomp_eeh = gdso_eeh - EEHEMT1_gdsv(vgs, vgs)
                    
                End If
                If vgs >= vb_eeh And b_eeh <> -1 Then
                    
                    vgs_va_eeh = vgs - va_eeh
                    vgs_vb_eeh = vgs - vb_eeh
                    gmocomp_eeh = gmo_eeh - (a_eeh * (vgs_va_eeh) ^ b_eeh + gmoff_eeh)
                    
                    Idsocomp_eeh = Idso_eeh - a_eeh / bp1_eeh * ((vgs_va_eeh) ^ bp1_eeh - (VBA_EE) ^ bp1_eeh) - gmoff_eeh * (vgs_vb_eeh) - Idsvvb_eeh
                    
                    gdsocomp_eeh = gdso_eeh - MU_EE * (a_eeh * (vgs_va_eeh) ^ b_eeh + gmoff_eeh) - gdsvvb_eeh
                    
                ElseIf vgs >= vb_eeh And b_eeh = -1 Then
                
                    gmocomp_eeh = gmo_eeh - (a_eeh * (vgs_va_eeh) ^ b_eeh + gmoff_eeh)
                    Idsocomp_eeh = Idso_eeh - a_eeh * (Log(vgs_va_eeh) - Log(VBA_EE)) - gmoff_eeh * (vgs_vb_eeh) - Idsvvb_eeh
                    gdsocomp_eeh = gdso_eeh - MU_EE * a_eeh / (vgs_va_eeh) - MU_EE * gmoff_eeh - gdsvvb_eeh
                    
                End If
            Else
            
                gmocomp_eeh = gmo_eeh
                Idsocomp_eeh = Idso_eeh
                gdsocomp_eeh = gdso_eeh

            End If

            kpvds_eeh = 1# + KAPA_EE * vds
            vsat3_eeh = 3# * vds / VSAT_EE
            thvst_eeh = tanh(vsat3_eeh)
            
            gmx_eeh = WG_EE / UGW_EE * gmocomp_eeh * (kpvds_eeh) * thvst_eeh
            Ikt_eeh = Idsocomp_eeh * (kpvds_eeh) * thvst_eeh
            wg_ugw_eeh = WG_EE / UGW_EE
            wuikt_eeh = wg_ugw_eeh * Ikt_eeh
            
            If vds < 0 Then
                'Idsx_eeh = -WG_EE / UGW_EE * Abs(Idsocomp_eeh * (1# + KAPA_EE * (vds)) * tanh(3# * vds / VSAT_EE))
                Idsx_eeh = wuikt_eeh
            Else
                Idsx_eeh = wuikt_eeh
            End If
            
            gdsx_eeh = wg_ugw_eeh * ((gdsocomp_eeh * (kpvds_eeh) + Idsocomp_eeh * KAPA_EE) * thvst_eeh + 3# * Idsocomp_eeh * (kpvds_eeh) / VSAT_EE * (1# - (thvst_eeh) ^ 2))

            pdis_eeh = 1# + Abs(Idsx_eeh) * vds / PEFF_EE
            pdissq_eeh = pdis_eeh * pdis_eeh
            
            gm_eeh = gmx_eeh / pdissq_eeh + GMIN
            Ids_eeh = Idsx_eeh / pdis_eeh + GMIN * vds
            gds_eeh = (gdsx_eeh - Idsx_eeh * Idsx_eeh / PEFF_EE) / pdissq_eeh + GMIN



End Sub

Function EEHEMT1_Ids(ByVal vdsx As Double, ByVal vgsx As Double, vgdx As Double) As Double
Dim Ids_eeh As Double, gds_eeh As Double, gm_eeh As Double, Idso_eeh As Double, Idsocomp_eeh As Double, gdsocomp_eeh As Double, gmocomp_eeh As Double

    Call EEHEMT1(vdsx, vgsx, vgdx, Ids_eeh, gds_eeh, gm_eeh, Idso_eeh, Idsocomp_eeh, gdsocomp_eeh, gmocomp_eeh)
    EEHEMT1_Ids = Ids_eeh

End Function

Function EEHEMT1_Idso(ByVal vdsx As Double, ByVal vgsx As Double, vgdx As Double) As Double
Dim Ids_eeh As Double, gds_eeh As Double, gm_eeh As Double, Idso_eeh As Double, Idsocomp_eeh As Double, gdsocomp_eeh As Double, gmocomp_eeh As Double

    Call EEHEMT1(vdsx, vgsx, vgdx, Ids_eeh, gds_eeh, gm_eeh, Idso_eeh, Idsocomp_eeh, gdsocomp_eeh, gmocomp_eeh)
    EEHEMT1_Idso = Idso_eeh

End Function

Function EEHEMT1_Idsocomp(ByVal vdsx As Double, ByVal vgsx As Double, vgdx As Double) As Double
Dim Ids_eeh As Double, gds_eeh As Double, gm_eeh As Double, Idso_eeh As Double, Idsocomp_eeh As Double, gdsocomp_eeh As Double, gmocomp_eeh As Double

    Call EEHEMT1(vdsx, vgsx, vgdx, Ids_eeh, gds_eeh, gm_eeh, Idso_eeh, Idsocomp_eeh, gdsocomp_eeh, gmocomp_eeh)
    EEHEMT1_Idsocomp = Idsocomp_eeh

End Function

Function EEHEMT1_gdsocomp(ByVal vdsx As Double, ByVal vgsx As Double, vgdx As Double) As Double
Dim Ids_eeh As Double, gds_eeh As Double, gm_eeh As Double, Idso_eeh As Double, Idsocomp_eeh As Double, gdsocomp_eeh As Double, gmocomp_eeh As Double

    Call EEHEMT1(vdsx, vgsx, vgdx, Ids_eeh, gds_eeh, gm_eeh, Idso_eeh, Idsocomp_eeh, gdsocomp_eeh, gmocomp_eeh)
    EEHEMT1_gdsocomp = gdsocomp_eeh

End Function

Function EEHEMT1_gmocomp(ByVal vdsx As Double, ByVal vgsx As Double, vgdx As Double) As Double
Dim Ids_eeh As Double, gds_eeh As Double, gm_eeh As Double, Idso_eeh As Double, Idsocomp_eeh As Double, gdsocomp_eeh As Double, gmocomp_eeh As Double

    Call EEHEMT1(vdsx, vgsx, vgdx, Ids_eeh, gds_eeh, gm_eeh, Idso_eeh, Idsocomp_eeh, gdsocomp_eeh, gmocomp_eeh)
    EEHEMT1_gmocomp = gmocomp_eeh

End Function

Function EEHEMT1_gm(ByVal vdsx As Double, ByVal vgsx As Double, vgdx As Double) As Double
Dim Ids_eeh As Double, gds_eeh As Double, gm_eeh As Double, Idso_eeh As Double, Idsocomp_eeh As Double, gdsocomp_eeh As Double, gmocomp_eeh As Double
            
    Call EEHEMT1(vdsx, vgsx, vgdx, Ids_eeh, gds_eeh, gm_eeh, Idso_eeh, Idsocomp_eeh, gdsocomp_eeh, gmocomp_eeh)
    EEHEMT1_gm = gm_eeh

End Function

Function EEHEMT1_gds(ByVal vdsx As Double, ByVal vgsx As Double, vgdx As Double) As Double
Dim Ids_eeh As Double, gds_eeh As Double, gm_eeh As Double, Idso_eeh As Double, Idsocomp_eeh As Double, gdsocomp_eeh As Double, gmocomp_eeh As Double

    Call EEHEMT1(vdsx, vgsx, vgdx, Ids_eeh, gds_eeh, gm_eeh, Idso_eeh, Idsocomp_eeh, gdsocomp_eeh, gmocomp_eeh)
    EEHEMT1_gds = gds_eeh

End Function

Function EEHEMT1_gmv(ByVal v As Double, ByVal vdsx As Double) As Double
Dim vds As Double, vc_eeh As Double

    vds = vdsx
    vc_eeh = VCO_EE + MU_EE * (VDSO_EE - (vds))
    EEHEMT1_gmv = DELTGM_EE * (Sqr(ALPHA_EE * ALPHA_EE + (v - vc_eeh) * (v - vc_eeh)) - ALPHA_EE)
    
End Function

Function EEHEMT1_gmm(ByVal v As Double, ByVal vdsx As Double) As Double
Dim gammvdso_eeh As Double, vg_eeh As Double, vt_eeh As Double, vx_eeh As Double
Dim vds As Double
vds = vdsx

    gammvdso_eeh = GAMMA_EE * (VDSO_EE - vdsx)
    
    vg_eeh = (VGO_EE - VCH_EE) / (1# + gammvdso_eeh) + VCH_EE
    vt_eeh = (VTO_EE - VCH_EE) / (1# + gammvdso_eeh) + VCH_EE
    vx_eeh = (v - VCH_EE) * (1# + gammvdso_eeh)
            
    EEHEMT1_gmm = GMMAX_EE / 2# * (1# + GAMMA_EE * (VDSO_EE - vds)) * (Cos(PI * (vx_eeh - (VGO_EE - VCH_EE)) / (VTO_EE - VGO_EE)) + 1#)
    
End Function


Function EEHEMT1_gdsv(ByVal v As Double, ByVal vds As Double) As Double
Dim vc_eeh As Double, dv_eeh As Double, valp_eeh As Double, vsq_eeh As Double, dv2_eeh As Double
Dim alp2ee_eeh As Double, rvsq_eeh As Double

    vc_eeh = VCO_EE + MU_EE * (VDSO_EE - vds)
    dv_eeh = v - vc_eeh
    valp_eeh = ALPHA_EE * ALPHA_EE + dv_eeh * dv_eeh
    vsq_eeh = Sqr(valp_eeh)
    dv2_eeh = (dv_eeh) ^ 2
    alp2ee_eeh = (ALPHA_EE) ^ 2
    rvsq_eeh = 1 / vsq_eeh
    EEHEMT1_gdsv = DELTGM_EE * MU_EE * (0.5 * (2 * dv2_eeh + alp2ee_eeh) * rvsq_eeh + 0.5 * alp2ee_eeh / (dv_eeh + vsq_eeh) * (1 + dv_eeh * rvsq_eeh) - ALPHA_EE)
    
    
    'EEHEMT1_gdsv = DELTGM_EE * MU_EE * (0.5 * ((2# * dv_eeh * dv_eeh + ALPHA_EE * ALPHA_EE) / vsq_eeh + ALPHA_EE * ALPHA_EE / (dv_eeh + vsq_eeh) * (1# + dv_eeh / vsq_eeh)) - ALPHA_EE)

End Function

Function EEHEMT1_Idsm(ByVal v As Double, ByVal vdsx As Double) As Double
Dim t1 As Double, t4 As Double

    t1 = VTO_EE - VGO_EE
    t4 = EEHEMT1_Vx(v, vdsx)
    EEHEMT1_Idsm = GMMAX_EE * (t1 / PI * sin(PI * (t4 - VGO_EE + VCH_EE) / t1) + t4 - VTO_EE + VCH_EE) / 2
    
    'EEHEMT1_Idsm = GMMAX_EE / 2# * (((VTO_EE - VGO_EE) / PI) * Sin(PI * (EEHEMT1_Vx(v, vdsx) - (VGO_EE - VCH_EE)) / (VTO_EE - VGO_EE)) + EEHEMT1_Vx(v, vdsx) - (VTO_EE - VCH_EE))
End Function

Function EEHEMT1_Idsv(ByVal v As Double, ByVal vdsx As Double) As Double
Dim vc_eeh As Double, dvc_eeh As Double, vsq_eeh As Double, alp2_eeh As Double

    vc_eeh = VCO_EE + MU_EE * (VDSO_EE - Abs(vdsx))
    dvc_eeh = v - vc_eeh
    vsq_eeh = Sqr(ALPHA_EE * ALPHA_EE + dvc_eeh * dvc_eeh)
    alp2_eeh = ALPHA_EE * ALPHA_EE
    EEHEMT1_Idsv = DELTGM_EE * (0.5 * (dvc_eeh * vsq_eeh + alp2_eeh * Log((dvc_eeh + vsq_eeh) / ALPHA_EE)) - ALPHA_EE * dvc_eeh)
    
End Function

Function EEHEMT1_Igd(ByVal vdsx As Double, ByVal vgsx As Double, ByVal vgdx As Double) As Double
Dim vds As Double, vgs As Double, vgd As Double, Ids_eeh As Double

            vgs = vgsx
            vgd = vgdx
            vds = vdsx
            
            Ids_eeh = EEHEMT1_Ids(vdsx, vgsx, vgd)
            
            If -vgd > VBR_EE Then
                EEHEMT1_Igd = -KBK_EE * (1# - Ids_eeh / IDSOC_EE) * (-vgd - VBR_EE) ^ NBR_EE + GMIN * vgd
            Else
                EEHEMT1_Igd = 0# + GMIN * vgd
            End If

End Function

Function EEHEMT1_Igs(ByVal vdsx As Double, ByVal vgsx As Double, ByVal vgdx As Double) As Double
Dim vkne_eeh As Double, exvkne_eeh As Double

    vkne_eeh = vgsx / Vkt / NF_EE
    exvkne_eeh = expm(vkne_eeh) - 1#
    
    EEHEMT1_Igs = ISG_EE * (exvkne_eeh) + GMIN * vgsx

End Function

Function EEHEMT1_Vx(ByVal vj As Double, ByVal vo As Double) As Double

    EEHEMT1_Vx = (vj - VCH_EE) * (1# + GAMMA_EE * (VDSO_EE - (vo)))
    
End Function

Function EEHEMT1_IgdVgd(ByVal vdsx As Double, ByVal vgsx As Double, ByVal vgdx As Double, ByVal vgsx As Double) As Double
Dim vds As Double, vgs As Double, vgd As Double, Ids_eeh As Double

            vgs = vgsx
            vgd = vgdx
            vds = vgsx - vgdx
            
            
            Ids_eeh = EEHEMT1_Ids(vdsx, vgsx)
            
            If -vgd > VBR_EE Then
                EEHEMT1_Igd = -KBK_EE * (1# - Ids_eeh / IDSOC_EE) * (-vgd - VBR_EE) ^ NBR_EE + GMIN * vgd
            Else
                EEHEMT1_Igd = 0# + GMIN * vgd
            End If

End Function

Sub FunctionTest()
Dim x As Double, i As Integer, h As Double, y As Double, z As Double, w As Double
Dim dx As Double, xmin As Double, xmax As Double, ixmax As Integer

    Sheet6.Select
    Sheet6.Cells.Select
    Selection.ClearContents
    Sheet6.Range("A1").Select

    Call parameters_set
    
    h = 0.01
    x = 0#
    xmin = -4#
    xmax = 2#
    ixmax = 200#
    dx = (xmax - xmin) / (ixmax - 1)
    For i = 1 To ixmax
        x = xmin + dx * (i - 1)
'        y = Isgs * (fn("exp", 1, x / Vkt / ngs, h, h) - 1#)
'        z = Isgs * (Exp(x / Vkt / ngs) - 1#): 'dfn7x("exp", 1, 1# * x, h, h, 0.001, 1)
        y = fn("gdsgs", 3, 1, x, 0)
        z = dfn7y("Ids", 3, 1, x, 0, 0.001, 1)
        w = dfn7y("gdsgs", 3, 1, x, 0, 0.001, 1)
        'w = fn("gdsds", 3, x, 1, 0)
        'Debug.Print "x:"; x; " y:"; y
        Sheet6.Cells(i, 1) = x
        Sheet6.Cells(i, 2) = y
        Sheet6.Cells(i, 3) = z
        Sheet6.Cells(i, 4) = w
       
    Next i


End Sub

Sub SingleNewtonTest()
Dim x As Double, i As Integer, j As Integer, h As Double, y As Double, z As Double, w As Double
Dim dx As Double, xmin As Double, xmax As Double, ixmax As Integer
Dim AA As Double, BB As Double, CC As Double, DF As Double, x0 As Double, x0_old As Double, sl As Double
Dim cf As Double, cfmin As Double, cfmax As Double
Dim Rload As Double, xia As Double, level As Integer, VGG As Double, VDDMAX As Double, VDDMIN As Double
Dim hh As Double, DFA As Double
Dim zl As Double, zlp As Double, zp As Double, xp As Double, initer As Integer
Dim xprobe As Double, zpr As Double, zlpr As Double
Dim funcname$, ia As Double, iterymax As Integer, itery As Integer, iter As Integer, iterx As Integer
Dim iterxmax As Integer, islp As Integer, iterini As Integer, Delta As Double
Dim msg$, mode$





    Sheet7.Select
    Sheet7.Cells.Select
    Selection.ClearContents
    Sheet7.Range("A1").Select

    Call parameters_set

    level = 2
    funcname$ = "Idrain"
    mode$ = "vds"
    Rload = 0.1 * kilo
    cfmin = 0.5
    cfmax = 1#
    
    hh = 0.001

Select Case mode$
    Case "vds"
        
        For VGG = -2 To 1 Step 1#
        'VGG = 0.5
        VDDMIN = -8#
        VDDMAX = 100#
        
        x0 = 0#
        x0_old = 10#
        
        ixmax = 101
        
        For ia = VDDMIN To VDDMAX Step (VDDMAX - VDDMIN) / ixmax
            xia = CDbl(ia)
            
            iterymax = 1000
            itermax = 200 + iterymax
            h = 0.01
            x = 0#
            xmin = VDDMIN
            xmax = VDDMAX
            dx = (xmax - xmin) / (ixmax - 1)
            
            'AA = xia / Rload
            BB = xia + GMIN
            CC = 1# / Rload
         
            For i = 1 To ixmax
                x = xmin + dx * (i - 1)
                zl = -CC * (x - BB)
                Sheet7.Cells(i, 1) = x
                Sheet7.Cells(i, 4) = zl
            Next i
            
            
            For i = 1 To ixmax
                x = xmin + dx * (i - 1)
                z = fn(funcname$, level, x, VGG, VGG - x)
                Sheet7.Cells(i, 1) = x
                Sheet7.Cells(i, 2) = z
            Next i
            
            iter = 0
            iterx = 0
            itery = 0
            sl = 1#
            'x0 = xia
            islp = 1
            cf = (cfmax - cfmin) / (1# + (cfmax - cfmin) * Abs(x0_old - x0)) + cfmin
            
            For i = 1 To ixmax
                x = xmin + dx * (i - 1)
                initer = 0
                xp = x + dx
                z = fn(funcname$, level, x, VGG, VGG - x)
                zl = -CC * (x - BB)
                zp = fn(funcname$, level, xp, VGG, VGG - xp)
                zlp = -CC * (xp - BB)
                If zl > z And zlp < zp Then
                    Debug.Print "solve initial data"
                    iterini = 0
                    Do
                        xprobe = (x + xp) / 2#
                        zpr = fn(funcname$, level, xprobe, VGG, VGG - xprobe)
                        zlpr = -CC * (xprobe - BB)
                        If (zlpr - zpr) * (zl - z) < 0 Then
                            xp = xprobe
                            zp = fn(funcname$, level, xp, VGG, VGG - xp)
                            zlp = -CC * (xp - BB)
                        Else
                            x = xprobe
                            z = fn(funcname$, level, x, VGG, VGG - x)
                            zl = -CC * (x - BB)
                        End If
                        If Abs(x - xp) < femt Then
                            x0 = x
                            x0_old = xp
                            Debug.Print "iterini:"; iterini; "x0ini:"; x0
                            Exit For
                        End If
                        If iterini > 10000 Then
                            Exit For
                        End If
                        iterini = iterini + 1
                    Loop
                End If
            Next i
                
                
            
            Do
                z = fn(funcname$, level, x0, VGG, VGG - x0)
                zp = fn(funcname$, level, x0 + hh, VGG, VGG - x0)
                DF = dfn7x(funcname$, level, x0, VGG, VGG - x0, 0.001, 1)
            DoEvents
                For j = 1 To ixmax
                    y = xmin + dx * (j - 1) + GMIN
                    If islp <> 0 Then
                        w = sl * DF * (y - x0) + z
                        islp = 1
                    Else
                        w = sl * z / x0 * y
                    End If
                    Sheet7.Cells(j, 1) = y
                    Sheet7.Cells(j, 3) = w
                Next j
                
                
                If iter > itermax * 0.1 Then
                    
                    'iter = 0
                    If itery < iterymax Then
                        x0 = x0_old - cf * (x0_old - (sl * CC * x0 * BB / (z + sl * CC * x0)))
                        Delta = Abs(x0_old - x0)
                        Debug.Print "VDD:"; xia; "VGG"; VGG; " x0:"; x0; "x0_old:"; x0_old; " delta:"; Delta; " sl:"; sl; " cf:"; cf
                        sl = sl * 1#
                        cf = cf / 1.1
                        If cf < 0.1 Then
                            cf = 0.1
                        End If
                        islp = 0
                        itery = itery + 1
                    Else
                        iter = 0
                        itery = 0
                    End If
                Else
                    x0 = x0_old - cf * (x0_old - (BB * sl * CC + DF * x0 - z) / (sl * CC + DF))
                    Delta = Abs(x0_old - x0)
                    islp = 1
                    cf = cf / 1.01
                End If
                
                If Delta < 1# * micro Then
                    Debug.Print "VDD:"; ia; " "; "iterx:"; iterx; " "; "x0:"; x0; " Delta:"; Delta; " sl:"; sl; " cf:"; cf
                    Sheet7.Cells(1, 5) = x0
                    Sheet7.Cells(1, 6) = fn(funcname$, level, x0, VGG, VGG - x0)
                    Exit Do
                End If
                If iter > itermax Then
                    msg$ = "Not converge!!"
                    MsgBox msg$
                    End
                End If
                
                x0_old = x0
                
                iter = iter + 1
                iterx = iterx + 1
                Sheet7.Cells(1, 5) = x0
                Sheet7.Cells(1, 6) = fn(funcname$, level, x0, VGG, VGG - x0)
                
                Application.Wait [Now() + "0:00:00.03"]
            Loop
                Application.Wait [Now() + "0:00:00.03"]
        Next ia
        
        Next VGG

    Case "vgs"
    
    Case "vgd"
    
    Case Else
End Select


End Sub

Function ftest(x As Double, y As Double) As Double
    
    If x < 0 Then
        ftest = 4# / (1# + expm(((x + 10# * y) / 10#) ^ 4))
    Else
        ftest = 4# / (1# + expm(((x - 10# * y) / 10#) ^ 4))
    End If
End Function

Function square(x As Double, y As Double, z As Double) As Double

    square = (x) ^ 2 + (y) ^ 2 + (z) ^ 2

End Function

Function cubic(x As Double, y As Double, z As Double) As Double

     cubic = (x) ^ 3 + (y) ^ 3 + (z) ^ 3

End Function

Function quad(x As Double, y As Double, z As Double) As Double

    quad = (x) ^ 4 + (y) ^ 4 + (z) ^ 4

End Function

Function sinc(x As Double) As Double

    If x = 0 Then
        sinc = 1#
    Else
        sinc = sin(x) / x
    End If
End Function

Function sec(x As Double) As Double
    sec = 1# / Cos(x)
End Function

Function cosec(x As Double) As Double

    cosec = 1# / sin(x)
    
End Function

Function cot(x As Double) As Double

    cot = Cos(x) / sin(x)
End Function


Sub IdsVgs_LoadLine_test()
Dim vds As Double, vgs As Double, vgd As Double, VDD As Double, VGG As Double
Dim VDDMAX As Double, VDDMIN As Double, VGGMAX As Double, VGGMIN As Double
Dim i As Integer, j As Integer, imax As Integer, jmax As Integer
Dim dVDD As Double, dVGG As Double
'Dim Rd As Double, Rs As Double,
Dim level As Integer
Dim yLp As Double, yIp As Double, yLn As Double, yIn As Double, yLc As Double, yIc As Double, fxn As Double, fxp As Double, fxc As Double
Dim vdsn As Double, vdsp As Double, vdsc As Double
Dim Icurrent$
Dim vgsn As Double, vgsp As Double, vgsc As Double, iter As Integer, itermax As Integer, msg$
Dim dvds As Double, yIc_old As Double
Dim iteroutermax, iterouter
Dim vgsc0 As Double, vgsc_old As Double, cf As Double, cfmin As Double, cfmax As Double
Dim yIgdn As Double, yIgdp As Double, yIgdc As Double, yIgsn As Double, yIgsp As Double, yIgsc As Double
Dim tollocal As Double




    Sheet8.Select
    Sheet8.Cells.Select
    Selection.ClearContents
    Sheet8.Range("A1").Select
    
    Call parameters_set

    cfmin = 0.05
    cfmax = 1#
    cf = 1#
    
    tollocal = milli
    
    VDDMAX = 10# + milli
    VDDMIN = -1# + milli
    
    VGGMAX = 2#
    VGGMIN = -4#
    
'    Rd = 3#
'    Rs = 10# * milli
    
    Icurrent$ = "Ids"
    
    level = 3: ' 1:Hirose, 2:Level=1, 3:EEHEMT1
    
    itermax = 10000
    iteroutermax = 10000
    imax = 21
    jmax = 21
    
    dVDD = (VDDMAX - VDDMIN) / CDbl(jmax - 1)
    dVGG = (VGGMAX - VGGMIN) / CDbl(imax - 1)
    
    
    For j = 1 To jmax: 'Power Supply
        VDD = VDDMIN + dVDD * (j - 1)
        
        For i = 1 To imax: 'Gate Voltage
            VGG = VGGMIN + dVGG * (i - 1)
            vgsc = VGG
            vgsc_old = VGG
            
            yIc_old = 100#
            iterouter = 0
            Do
                'DoEvents
            
                For vds = VDDMIN To VDD Step dVDD
                    
                        vdsn = vds
                        vdsp = vds + dVDD
                        
                        
                        yIgdn = fn("Igd", level, vdsn, vgsc, vgsc - vdsn)
                        yIgsn = fn("Igs", level, vdsn, vgsc, vgsc - vdsn)
                        yLn = -1# / (Rd + Rs) * (vdsn - VDD) + 1# / (Rd + Rs) * (yIgdn * Rd - yIgsn * Rs)
                        yIn = fn(Icurrent$, level, vdsn, vgsc, vgsc - vdsn)
                        
                        yIgdp = fn("Igd", level, vdsp, vgsc, vgsc - vdsp)
                        yIgsp = fn("Igs", level, vdsp, vgsc, vgsc - vdsp)
                        yLp = -1# / (Rd + Rs) * (vdsp - VDD) + 1# / (Rd + Rs) * (yIgdp * Rd - yIgsp * Rs)
                        yIp = fn(Icurrent$, level, vdsp, vgsc, vgsc - vdsp)
                        
                        fxn = yLn - yIn
                        fxp = yLp - yIp
                        
                        
                        If fxn * fxp < 0 Then
                        
                            iter = 0
                            
                            Do
                                'DoEvents
                                
                                yIgdn = fn("Igd", level, vdsn, vgsc, vgsc - vdsn)
                                yIgsn = fn("Igs", level, vdsn, vgsc, vgsc - vdsn)
                                yLn = -1# / (Rd + Rs) * (vdsn - VDD) + 1# / (Rd + Rs) * (yIgdn * Rd - yIgsn * Rs)
                                yIn = fn(Icurrent$, level, vdsn, vgsc, vgsc - vdsn)
                                
                                yIgdp = fn("Igd", level, vdsp, vgsc, vgsc - vdsp)
                                yIgsp = fn("Igs", level, vdsp, vgsc, vgsc - vdsp)
                                yLp = -1# / (Rd + Rs) * (vdsp - VDD) + 1# / (Rd + Rs) * (yIgdp * Rd - yIgsp * Rs)
                                yIp = fn(Icurrent$, level, vdsp, vgsc, vgsc - vdsp)
                                
                                fxn = yLn - yIn
                                fxp = yLp - yIp
                        
                                vdsc = (vdsp + vdsn) / 2#
                                yIgdc = fn("Igd", level, vdsc, vgsc, vgsc - vdsc)
                                yIgsc = fn("Igs", level, vdsc, vgsc, vgsc - vdsc)
                                yLc = -1# / (Rd + Rs) * (vdsc - VDD) + 1# / (Rd + Rs) * (yIgdc * Rd - yIgsc * Rs)
                                yIc = fn(Icurrent$, level, vdsc, vgsc, vgsc - vdsc)
                                
                                fxc = yLc - yIc
                                
                                'Debug.Print "fxc:"; fxc
                                
                                If fxn * fxc < 0 Then
                                    vdsp = vdsc
                                Else
                                    vdsn = vdsc
                                End If
                                
                                If Abs(fxc) < tollocal Then
                                
                                    'Debug.Print " VDD:"; VDD; " VGG:"; VGG; " vdsc:"; vdsc; " Ids:"; yIc
                                    Sheet8.Cells(1, j) = VDD
                                    Sheet8.Cells(i + 1, 1) = VGG
                                    Sheet8.Cells(i + 1, j) = fn("Idrain", level, vdsc, vgsc, vgsc - vdsc)
                                    
                                    vgsc0 = VGG - (yIgsc + yIc) * Rs - (yIgsc + yIgdc) * Rg
                                    vgsc = vgsc_old - cf * (vgsc_old - vgsc0)
                                    
                                    Exit For
                                End If
                                
                                If iter > itermax Then
                                    msg$ = "No convergence @iter"
                                    MsgBox msg$
                                    End
                                End If
                                
                                iter = iter + 1
                                
                            Loop
                            
                        End If
                        
                    Next vds
                    
                    If Abs(yIc - yIc_old) + Abs(vgsc - vgsc_old) < tollocal Then
                        Exit Do
                    Else
                        vgsc_old = vgsc
                        cf = cf * 0.995
                        If cf < cfmin Then
                            cf = cfmin
                        End If
                        'Debug.Print "iterouter:"; iterouter; " cf:"; cf
                    End If
                    
                    If iterouter > iteroutermax Then
                        msg$ = "No convergence @iterouter"
                        MsgBox msg$
                        End
                    End If
                    iterouter = iterouter + 1
                    yIc_old = yIc
                Loop
            
        Next i
        
    Next j



End Sub
