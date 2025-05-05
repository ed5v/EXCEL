VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FORMULARIO 
   Caption         =   "UserForm1"
   ClientHeight    =   5355
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7755
   OleObjectBlob   =   "FORMULARIO.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "FORMULARIO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False






Private Sub ACEPTAR_Click()
On Error Resume Next
If Me.TXBGRAF = "" Then
MsgBox "ASIGNA PUNTOS A CALIBRAR"
Exit Sub

End If

FORMULARIO.Hide
End Sub

Private Sub CANCELAR_Click()

Dim P_1, P_2, P_3 As String
P_1 = "DATOS"
P_2 = "MENU"
P_3 = "CERTIFICADOS"


MsgBox "SE DETUVO LA GENERACIÓN DE CERTIFICADOS"
Worksheets(P_3).Protect "MET2025"

Range("GW4:HC20").value = " "
On Error Resume Next
Selection.ShapeRange(1).Delete
On Error GoTo 0

On Error Resume Next
graf.Delete


'Columns("GS:HH").Hidden = False
Application.ScreenUpdating = True
On Error GoTo 0

Exit Sub
Unload FORMULARIO
End Sub





Private Sub TXBGRAF_Change()
On Error Resume Next
Dim VAL As String
VAL = TXBGRAF.List(TXBGRAF.ListIndex, 0)
If VAL = "" Then
    If TXBGRAF <> "" Then
    MsgBox "NO SE PERMITEN VALORES DIFERENTES"
    TXBGRAF.SetFocus
    TXBGRAF = ""
    End If
End If

If valor = "M_g" Or valor = "M_Kg" Then
    If VAL > 8 Then
    MsgBox ("PARA CERTIFICADOS DE MASA CON MÁS DE 8 PUNTOS SE UTILIZA LA SIGUIENTE FILA A PARTIR DEL PUNTO 3," & _
            " ES NECESARIO COLOCAR ID DE INSTRUMENTO, TIPO DE SERVICIO Y PATRÓN")
    End If

End If

End Sub

Private Sub UserForm_Initialize()

FORMULARIO.Caption = INSTRUMENTO

Me.FRMINSTRUMENTO.Caption = ID

Me.TXBGRAF.AddItem (1)
Me.TXBGRAF.AddItem (2)
Me.TXBGRAF.AddItem (3)
Me.TXBGRAF.AddItem (4)
Me.TXBGRAF.AddItem (5)
Me.TXBGRAF.AddItem (6)
Me.TXBGRAF.AddItem (7)
Me.TXBGRAF.AddItem (8)
Me.TXBGRAF.AddItem (9)
Me.TXBGRAF.AddItem (10)

Me.TXBINSTRUMENTO.value = ".0"
Me.TXBPATRON.value = ".0"

Me.FRMPATRON.Caption = PATRON
Label8.Caption = UNIDADES
Label9.Caption = UNIDADES
Label13.Caption = UNIDADES
Label14.Caption = UNIDADES

'''''''''''''''''FECHA'''''''''''''''
TXBFECHA.value = Format(Range("D1").value, "DD/MMM/YY")

''''''''''''''''''''''''''''''''''''''

If TIPO = "IP" Or TIPO = "CP" Or TIPO = "MB" Or TIPO = "VA" Or TIPO = "IV" Or TIPO = "IA" Then
TXBMIN.Locked = False
TXBMAX.Locked = False
TXBMAX.Visible = True
TXBMIN.Visible = True
Label3.Visible = True
Label5.Visible = True
Label8.Visible = True
Label9.Visible = True

ElseIf TIPO = "IH" Or TIPO = "CH" Then

TXBMIN.Locked = False
TXBMAX.Locked = False
TXBMAX.Visible = True
TXBMIN.Visible = True
TXBMIN.value = 0
TXBMAX.value = 100
Label3.Visible = True
Label5.Visible = True
Label8.Visible = True
Label9.Visible = True

Else
TXBMIN.Locked = True
TXBMAX.Locked = True
TXBMAX.Visible = False
TXBMIN.Visible = False
Label3.Visible = False
Label5.Visible = False
Label8.Visible = False
Label9.Visible = False
End If

End Sub

