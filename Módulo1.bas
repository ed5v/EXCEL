Attribute VB_Name = "Módulo1"
Public FILA, TITULOY, TITULOX, TITULOGRAF, PUNTOSREF, PATRON, ID, MAGNITUD As String
Public INSTRUMENTO, FOLIO, UNIDADES, TIPO, NOMQR, REALIZA, INGENIERO, REVISA As String
Public SEGURIDAD, valor, FECHA_C, ID_TIPO As String
Public DES_PATRON, UBICACION, FIJO_PATRON As String


Public graf As Shape
Public RNg As Range
Public OBQR As Shape
Public COM As Chart
Public P_1, P_2, P_3 As String

Public objWord As Object
Public objDoc As Object
Public strNombreArchivo As String



Sub GenerarCERTIFICADO()

'''''' PENDIENTE RUTA PRUEBAS

Dim CERTIFICADO, REALIZA, FECHA  As String
Dim MAGNITUD_CERTIFICADO, MAGNITUD_PATRON, PROCEDIMIENTO, NORMA As String
Dim TIPO_ERROR, SIGLAS_ERROR As String
Dim objWord As Object
Dim Wordoc As Object
Dim COORDENADAS_MASA As Integer
Dim lookupvalue As Variant, value As Variant, lookupRange As Range
'Dim V_FECHA As Date

''''''''''''''''''''''PERSONAL''''''''''''''''''''''


'''''''''''''''CONTRASEÑA'''''''''''
SEGURIDAD = "MET2025"
''''''''''''''''PESTAÑAS'''''''''''''''''
P_1 = "DATOS"
P_2 = "MENU"
P_3 = "CERTIFICADOS"

Worksheets(P_3).Unprotect "MET2025"


'''''''''''''CONTADOR DE DATOS
NUMDATOS = Sheets(P_3).Range("A" & Rows.Count).End(xlUp).Row - 9
'NUM_ING = Sheets(P_2).Range("A" & Rows.Count).End(xlUp).Row - 9



''''''''''''FOLIO
COD = Sheets(P_3).Range("D6").value
NUM = InStr(COD, "-") - 1
NUM_COD = Len(COD)

If NUM <> -1 Then
RAZ_SOC = Left(COD, NUM)
FOLIO = Right(COD, 4)

ElseIf NUM = -1 Then
RAZ_SOC = "PRO"
FOLIO = Format(Sheets(P_3).Range("D6").value, "0000")

End If




''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''



'LUGAR Y UBICACIÓN


'DETERMINACIÓN DE FOLIOS #'RAZÓN SOCIAL 3 LETRAS' - 'AÑO'-CONSECUTIVO #
If RAZ_SOC <> "PRO" Then

    If RAZ_SOC = "RAZONSOCIAL1" Then
    REVISA = "NOMBRE REVISA 1"
    
    ElseIf RAZ_SOC <> "RAZONSOCIAL2" Then
    REVISA = "NOMBRE REVISA 2"
    End If
    


'LUGAR
'Sheets(P_3).Range("D3").value = "NAVE, ÁREA, EQUIPO,CÓDIGO."

'UBICACIÓN
Sheets(P_3).Range("D4").value = "DIRECCION " & _

End If


'''''''''''''''
'FECHA GENERACIÓN

F_G = Format(Now(), "DD/MMM/YY")
FECHA_GENERACION = Format(F_G, ">")
''''


Application.ScreenUpdating = False

Worksheets(P_3).Unprotect "MET2025"
Columns("GS:HH").Hidden = False

'''''''''''''''''''''''''''''''''''''''''''''''' REGISTRO''''''''''''''''''''''''''''''
On Error Resume Next
Call REGISTRO



SALIDA = ""
Range("GO10").Select
FILA = 10

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''



Dim valor_b As Variant, buscar_R As Range
ING = Range("D5").value 'celda con el valor que buscamos



''''''''''TABLA INGENIEROS

USUARIOS = Sheets(P_2).Range("H3:J16")

V_USUARIO = Application.VLookup(ING, USUARIOS, 2, False)
INGENIERO = Application.VLookup(ING, USUARIOS, 3, False)

If IsError(valor_b) Then
MsgBox "USUARIO NO AUTORIZADO"
Exit Sub

Else

REALIZA = V_USUARIO
End If


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

     While ActiveCell <> ""
     
CERT_:

If Range("GO" & FILA) = "" Then
GoTo FIN
Else
    ID = ActiveSheet.Cells(FILA, 1)
End If

GEN = Application.InputBox("GENERAR CERTIFICADO", "DESEA GENERAR CERTIFICADO PARA:" & ID & "?", "S/N")
 
 If GEN = "N" Or GEN = "n" Then
 FILA = FILA + 1
 
 GoTo CERT_
 
 End If
 
 
 ''''''''''TABLA UBICACION



    Sheets(P_1).Select
NUM_INSTRUMENTOS = Sheets(P_1).Range("A" & Rows.Count).End(xlUp).Row - 1
NUM_INSTRUMENTOS_TOTAL = NUM_INSTRUMENTOS + 1

tbl_LISTA_MAESTRA = Sheets(P_1).Range("A2:AG" & NUM_INSTRUMENTOS_TOTAL)

UBICACION_ID = Application.VLookup(ID, tbl_LISTA_MAESTRA, 25, False)
 

''''''''''TABLA PUNTOS FIJOS



PUNTO_FIJO = Application.VLookup(ID, tbl_LISTA_MAESTRA, 33, False)

If PUNTO_FIJO = "I" Then
FIJO_PATRON = "IBC"

ElseIf PUNTO_FIJO = "P" Then
FIJO_PATRON = "PATRON"
End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sheets(P_3).Select
 
 
 
    valor = ActiveSheet.Cells(FILA, 197).value
    
    '''''''''''''''''''''' MAGNITUD
    
    If valor = "N5-11-ITH-005_T VAISALA" Or valor = "N5-11-ITH-011_T VAISALA" Then
    MAGNITUD_PATRON = "TEMPERATURA"

    
    ElseIf valor = "N5-11-ITH-005_H VAISALA" Or valor = "N5-11-ITH-011_H VAISALA" Then
    MAGNITUD_PATRON = "HUMEDAD"

    End If
    
    If MAGNITUD_PATRON = "TEMPERATURA" Then
        
    MAGNITUD_CERTIFICADO = "TEMPERATURA"
    PROCEDIMIENTO = "P-MET-02"
    NORMA = "NOM-008-SCFI-2002" & Chr(10) & "Publicación Técnica: “Termometría de Resistencia” CENAM (CNM-MET-PT-009) "
    TIPO_E_M = "E.M."
    REPORTE_ERROR = ""
    
    ElseIf MAGNITUD_PATRON = "HUMEDAD" Then
        
    MAGNITUD_CERTIFICADO = "HUMEDAD RELATIVA"
    PROCEDIMIENTO = "P-MET-04"
    NORMA = "Publicación Técnica: “CNM-MET-PT-010 Mediciones de Humedad 1997" & Chr(10) & "Guía Técnica sobre Verificación de Higrómetros"
    TIPO_E_M = "E.M.R."
    REPORTE_ERROR = "El instrumento tiene un error máximo relativo de:  *W5 *ZV2"
    
    End If
    ''''''''''''''''''''''
'/
ID = ActiveSheet.Cells(FILA, 1)
INSTRUMENTO = ActiveSheet.Cells(FILA, 2)


SERVICIO = ActiveSheet.Cells(FILA, 196)
PATRON = ActiveSheet.Cells(FILA, 197)
UNIDADES = ActiveSheet.Cells(FILA, 9)
TIPO = ActiveSheet.Cells(FILA, 3)

''''''magnitud
M_I = 0
MAGNITUD_INSTRUMENTO = Sheets(P_1).Cells(2, 24).Offset(M_I, 0)

If MAGNITUD = "PRESIÓN" Or MAGNITUD = "PRESION" Then
    If UNIDADES = "Mpa" Or UNIDADES = "mpa" Or UNIDADES = "MPA" Then
    MsgBox ("Unidades ingresadas incorrectamente, se modificara por MPa")
    UNIDADES = "MPa"
    ElseIf UNIDADES = "kg/cm2" Or UNIDADES = "Kg/cm2" Or UNIDADES = "KG/CM2" Then
    MsgBox ("Unidades ingresadas incorrectamente, se modificara por kg/cm2")
    UNIDADES = "kg/cm2"
    End If
    
End If

M_I = 1
'''''''''''''''''''''''''''

'''''''''''''''SERVICIO
SERV = ActiveSheet.Cells(FILA, 196)
If SERV = "CALIBRACIÓN" Then
    SERV_CERT = "CC."
ElseIf SERV = "VERIFICACIÓN" Then
    SERV_CERT = "CV."
End If
'''''''''''''''

''''''''''TABLA INGENIEROS
'NUMDATOS = NUMDATOS + 1

NUMDATOSP_1 = Sheets(P_1).Range("A" & Rows.Count).End(xlUp).Row
TABLA_ID = Sheets(P_1).Range("A1:AG" & NUMDATOSP_1)


ID_TIPO = Application.VLookup(ID, TABLA_ID, 33, False)


''''''''''''''''''''''''''''''''''''''''''''''
''''''''''TABLA INSTRUMENTOS
'Call DATOS_PATRON

''''''''''''''''''''''''''''''''''''''''''''''



FORMULARIO.Show

PUNTOSREF = FORMULARIO.TXBGRAF.value
FECHA_C = Format(FORMULARIO.TXBFECHA.value, ">")

ERROR_ = ActiveSheet.Cells(FILA, 199).value


If TIPO = "IP" Or TIPO = "CP" Or TIPO = "MB" Or TIPO = "VA" Then

EMR = Round((ERROR_ * 100) / (FORMULARIO.TXBMAX.value - FORMULARIO.TXBMIN.value), 1)

End If


DGTOPAT = FORMULARIO.TXBPATRON.value
DGTO = FORMULARIO.TXBINSTRUMENTO.value

TITULOY = "ERROR [" & UNIDADES & "]"
TITULOX = "LECTURA [" & UNIDADES & "]"
TITULOGRAF = "ERROR + U [" & UNIDADES & "]"


''''''''''''''''''''''''''''''''''''''----------PATRONES--------'''''''''''''''''''''''''''''''''

On Error GoTo PLANTILLA


    If valor = "N5-11-CT-001 RTD" Then               'Certificado ANTES TIPO k T_MT060001
        
        patharch = ThisWorkbook.Path & "\Certificado N5-11-CT-001.dotx"
        Set objWord = CreateObject("Word.Application")
        objWord.Visible = True
        objWord.documents.Add Template:=patharch, NewTemplate:=False, DocumentType:=0
        'objWord.documents.Open (patharch)
        Set Wordoc = objWord.documents.Open(patharch)
    
       ElseIf valor = "N5-11-CT-002 BAÑORTD" Then               'Certificado antes 1° baño seco N5-11-CT-002 BAÑORTD
  
           patharch = ThisWorkbook.Path & "\CERTIFICADO N5-11-CT-002.dotx"
           Set objWord = CreateObject("Word.Application")
           objWord.Visible = True
           objWord.documents.Add Template:=patharch, NewTemplate:=False, DocumentType:=0
            
            Set Wordoc = objWord.documents.Open(patharch)
          
       ElseIf valor = "N5-11-CT-005 FLUKEPOZOSECO" Then               'Certificado antes 1° baño seco NN5-11-CT-005 FLUKEPOZOSECO
           
           patharch = ThisWorkbook.Path & "\CERTIFICADO N5-11-CT-005.dotx"
           Set objWord = CreateObject("Word.Application")
           objWord.Visible = True
           objWord.documents.Add Template:=patharch, NewTemplate:=False, DocumentType:=0
            
            Set Wordoc = objWord.documents.Open(patharch)
                   
       ElseIf valor = "N5-11-CT-006 FLUKEPOZOSECO" Then
           
           patharch = ThisWorkbook.Path & "\CERTIFICADO N5-11-CT-006.dotx"
           Set objWord = CreateObject("Word.Application")
           objWord.Visible = True
           objWord.documents.Add Template:=patharch, NewTemplate:=False, DocumentType:=0
            
            Set Wordoc = objWord.documents.Open(patharch)
        
         ElseIf valor = "N5-11-CT-007 FLUKEPOZOSECO" Then
           
           patharch = ThisWorkbook.Path & "\CERTIFICADO N5-11-CT-007.dotx"
           Set objWord = CreateObject("Word.Application")
           objWord.Visible = True
           objWord.documents.Add Template:=patharch, NewTemplate:=False, DocumentType:=0
            
            Set Wordoc = objWord.documents.Open(patharch)
   
       ElseIf valor = "N5-11-CT-003 BAÑORTD" Then               'Certificado antes 1° baño seco N5-11-CT-003 BAÑORTD
                    
           patharch = ThisWorkbook.Path & "\CERTIFICADO N5-11-CT-003.dotx"
           Set objWord = CreateObject("Word.Application")
           objWord.Visible = True
           objWord.documents.Add Template:=patharch, NewTemplate:=False, DocumentType:=0
            
            Set Wordoc = objWord.documents.Open(patharch)
 
       ElseIf valor = "N5-11-IA-006 FLUKET" Then               'Certificado NUEVO
                      
           patharch = ThisWorkbook.Path & "\CERTIFICADO N5-11-IA-006.dotx"
           Set objWord = CreateObject("Word.Application")
           objWord.Visible = True
           objWord.documents.Add Template:=patharch, NewTemplate:=False, DocumentType:=0
           
           Set Wordoc = objWord.documents.Open(patharch)
           
   
       ElseIf valor = "N5-11-IA-007 FLUKET" Then
    
           patharch = ThisWorkbook.Path & "\CERTIFICADO N5-11-IA-007.dotx"
           Set objWord = CreateObject("Word.Application")
           objWord.Visible = True
           objWord.documents.Add Template:=patharch, NewTemplate:=False, DocumentType:=0
           
           Set Wordoc = objWord.documents.Open(patharch)
           
               
        ElseIf valor = "N5-11-IT-004 RTD" Then             'Certificado ANTES OMRON IND. J T_MT060004
            
            patharch = ThisWorkbook.Path & "\CERTIFICADO N5-11-IT-004.dotx"
            Set objWord = CreateObject("Word.Application")
            objWord.Visible = True
            objWord.documents.Add Template:=patharch, NewTemplate:=False, DocumentType:=0
            
            Set Wordoc = objWord.documents.Open(patharch)
        
    ElseIf valor = "N5-11-IT-003 MULTITERMOPAR" Then            'Certificado ANTES MULTITERMOPAR J T_MT060003
                
                patharch = ThisWorkbook.Path & "\CERTIFICADO N5-11-IT-003.dotx"
                Set objWord = CreateObject("Word.Application")
                objWord.Visible = True
                objWord.documents.Add Template:=patharch, NewTemplate:=False, DocumentType:=0
                
                Set Wordoc = objWord.documents.Open(patharch)
                
           
        ElseIf valor = "N5-11-IP-004 FLUKEPB" Then         'Certificado ANTES FLUKE P_MT14004
                    
                    patharch = ThisWorkbook.Path & "\CERTIFICADO N5-11-IP-004.dotx"
                    Set objWord = CreateObject("Word.Application")
                    objWord.Visible = True
                    objWord.documents.Add Template:=patharch, NewTemplate:=False, DocumentType:=0
                    
                    Set Wordoc = objWord.documents.Open(patharch)
        
        
        ElseIf valor = "N5-11-IP-005 WIKAMP" Then          'Certificado T_MT14005
                        
                        patharch = ThisWorkbook.Path & "\CERTIFICADO N5-11-IP-005.dotx"
                        Set objWord = CreateObject("Word.Application")
                        objWord.Visible = True
                        objWord.documents.Add Template:=patharch, NewTemplate:=False, DocumentType:=0
                        
                        Set Wordoc = objWord.documents.Open(patharch)
                    
        ElseIf valor = "N5-11-IP-003 WIKAAP" Then           'Certificado P_WIKA ALTA PRESIÓN
                                       
                            patharch = ThisWorkbook.Path & "\CERTIFICADO N5-11-IP-003.dotx"
                            Set objWord = CreateObject("Word.Application")
                            objWord.Visible = True
                            objWord.documents.Add Template:=patharch, NewTemplate:=False, DocumentType:=0
                            
                            Set Wordoc = objWord.documents.Open(patharch)
                      
    ElseIf valor = "N5-11-IP-006 WIKAMP" Then           'Certificado P_WIKA MEDIA PRESIÓN 2 PATRON
                                                        
                            patharch = ThisWorkbook.Path & "\CERTIFICADO N5-11-IP-006.dotx"
                            Set objWord = CreateObject("Word.Application")
                            objWord.Visible = True
                            objWord.documents.Add Template:=patharch, NewTemplate:=False, DocumentType:=0
                            
                            Set Wordoc = objWord.documents.Open(patharch)

                ElseIf valor = "N5-11-IP-002 DWYERPDPa" Then           'Certificado P_Dwyer Pa
                            
                            
                            patharch = ThisWorkbook.Path & "\CERTIFICADO N5-11-IP-002Pa.dotx"
                            Set objWord = CreateObject("Word.Application")
                            objWord.Visible = True
                            objWord.documents.Add Template:=patharch, NewTemplate:=False, DocumentType:=0
                            
                            Set Wordoc = objWord.documents.Open(patharch)
                
                ElseIf valor = "N5-11-IP-002 DWYERPDkPa" Then           'Certificado P_Dwyer kpa

                            
                            patharch = ThisWorkbook.Path & "\CERTIFICADO N5-11-IP-002KPa.dotx"
                            Set objWord = CreateObject("Word.Application")
                            objWord.Visible = True
                            objWord.documents.Add Template:=patharch, NewTemplate:=False, DocumentType:=0
                            
                            Set Wordoc = objWord.documents.Open(patharch)

            ElseIf valor = "N5-11-IP-007 PDPa" Then           'Certificado P_Testo Pa
                                                        
                            patharch = ThisWorkbook.Path & "\CERTIFICADO N5-11-IP-007Pa.dotx"
                            Set objWord = CreateObject("Word.Application")
                            objWord.Visible = True
                            objWord.documents.Add Template:=patharch, NewTemplate:=False, DocumentType:=0
                            
                            Set Wordoc = objWord.documents.Open(patharch)
                            
                 ElseIf valor = "N5-11-IP-008 FLUKEPB" Then           'Certificado P_NUEVO FLUKE
                                                        
                            patharch = ThisWorkbook.Path & "\CERTIFICADO N5-11-IP-008.dotx"
                            Set objWord = CreateObject("Word.Application")
                            objWord.Visible = True
                            objWord.documents.Add Template:=patharch, NewTemplate:=False, DocumentType:=0
                            
                            Set Wordoc = objWord.documents.Open(patharch)
            


               ElseIf valor = "M_Kg" Then           'Certificado M_BASCULAS
                            patharch = ThisWorkbook.Path & "\CERTIFICADO_M kg.dotx"
                            Set objWord = CreateObject("Word.Application")
                            objWord.Visible = True
                            objWord.documents.Add Template:=patharch, NewTemplate:=False, DocumentType:=0
                            
                            Set Wordoc = objWord.documents.Open(patharch)
                            


'///////



Dim dado(0 To 1, 0 To 41) As String '(columna,fila)


dado(0, 0) = "*ZV3"
dado(1, 0) = Format(ActiveSheet.Cells(FILA, 108), "0" & DGTOPAT) '*-*-*
dado(0, 1) = "*ZV4"
dado(1, 1) = Format(ActiveSheet.Cells(FILA, 109), "0" & DGTO) '*-*-*

''''''''''''''''_____________ERROR_4_____________
dado(0, 2) = "*ZV5"
dado(1, 2) = Format(dado(1, 1) - dado(1, 0), "0" & DGTOPAT) '*-*-*

'dado(0, 2) = "*ZV5"
'dado(1, 2) = Format(ActiveSheet.Cells(FILA, 109) - ActiveSheet.Cells(FILA, 108), "0" & DGTOPAT) '*-*-*
'''''''''''''''''''''''''''''''''''''''''''
dado(0, 3) = "*ZV7"
dado(1, 3) = Format(ActiveSheet.Cells(FILA, 110), "0" & DGTOPAT) '*-*-*
dado(0, 4) = "*ZV8"
dado(1, 4) = Format(ActiveSheet.Cells(FILA, 111), "0" & DGTO) '*-*-*

''''''''''''''''_____________ERROR_5_____________
dado(0, 5) = "*ZV9"
dado(1, 5) = Format(dado(1, 4) - dado(1, 3), "0" & DGTOPAT) '*-*-*

'dado(0, 5) = "*ZV9"
'dado(1, 5) = Format(ActiveSheet.Cells(FILA, 111) - ActiveSheet.Cells(FILA, 110), "0" & DGTOPAT) '*-*-*
''''''''''''''''''''''''''''''''''''''''''''''''''''

dado(0, 6) = "*ZA1"
dado(1, 6) = Format(ActiveSheet.Cells(FILA, 112), "0" & DGTOPAT) '*-*-*
dado(0, 7) = "*ZA2"
dado(1, 7) = Format(ActiveSheet.Cells(FILA, 113), "0" & DGTO) '*-*-*

''''''''''''''''_____________ERROR_6_____________
dado(0, 8) = "*ZA3"
dado(1, 8) = Format(dado(1, 7) - dado(1, 6), "0" & DGTOPAT) '*-*-*

'dado(0, 8) = "*ZA3"
'dado(1, 8) = Format(ActiveSheet.Cells(FILA, 113) - ActiveSheet.Cells(FILA, 112), "0" & DGTOPAT) '*-*-*
'''''''''''''''''''''''''''''''''''''''''''''''''''

dado(0, 9) = "*ZA5"
dado(1, 9) = Format(ActiveSheet.Cells(FILA, 114), "0" & DGTOPAT) '*-*-*
dado(0, 10) = "*ZA6"
dado(1, 10) = Format(ActiveSheet.Cells(FILA, 115), "0" & DGTO) '*-*-*

''''''''''''''''_____________ERROR_7_____________
dado(0, 11) = "*ZA7"
dado(1, 11) = Format(dado(1, 10) - dado(1, 9), "0" & DGTOPAT) '*-*-*

'dado(0, 11) = "*ZA7"
'dado(1, 11) = Format(ActiveSheet.Cells(FILA, 115) - ActiveSheet.Cells(FILA, 114), "0" & DGTOPAT) '*-*-*
''''''''''''''''''''''''''''''''''''''''''''''''''

dado(0, 12) = "*ZA9"
dado(1, 12) = Format(ActiveSheet.Cells(FILA, 116), "0" & DGTOPAT) '*-*-*
dado(0, 13) = "*ZB1"
dado(1, 13) = Format(ActiveSheet.Cells(FILA, 117), "0" & DGTO) '*-*-*

''''''''''''''''_____________ERROR_8_____________
dado(0, 14) = "*ZB2"
dado(1, 14) = Format(dado(1, 13) - dado(1, 12), "0" & DGTOPAT) '*-*-*

'dado(0, 14) = "*ZB2"
'dado(1, 14) = Format(ActiveSheet.Cells(FILA, 117) - ActiveSheet.Cells(FILA, 116), "0" & DGTOPAT) '*-*-*
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
dado(0, 15) = "*ZB4"
dado(1, 15) = Format(ActiveSheet.Cells(FILA, 118), "0" & DGTOPAT) '*-*-*
dado(0, 16) = "*ZB5"
dado(1, 16) = Format(ActiveSheet.Cells(FILA, 119), "0" & DGTO) '*-*-*

''''''''''''''''_____________ERROR_9_____________
dado(0, 17) = "*ZB6"
dado(1, 17) = Format(dado(1, 16) - dado(1, 15), "0" & DGTOPAT) '*-*-*

'dado(0, 17) = "*ZB6"
'dado(1, 17) = Format(ActiveSheet.Cells(FILA, 119) - ActiveSheet.Cells(FILA, 118), "0" & DGTOPAT) '*-*-*
'''''''''''''''''''''''''''''''''''''''''''
dado(0, 18) = "*ZB8"
dado(1, 18) = Format(ActiveSheet.Cells(FILA, 120), "0" & DGTOPAT) '*-*-*
dado(0, 19) = "*ZB9"
dado(1, 19) = Format(ActiveSheet.Cells(FILA, 121), "0" & DGTO) '*-*-*

''''''''''''''''_____________ERROR_10_____________
dado(0, 20) = "*ZC1"
dado(1, 20) = Format(dado(1, 19) - dado(1, 18), "0" & DGTOPAT) '*-*-*

'dado(0, 20) = "*ZC1"
'dado(1, 20) = Format(ActiveSheet.Cells(FILA, 121) - ActiveSheet.Cells(FILA, 120), "0" & DGTOPAT) '*-*-*
''''''''''''''''''''''''''''''''''''''''''

dado(0, 21) = "*ZC4"
dado(1, 21) = Format(ActiveSheet.Cells(FILA, 53), "0" & DGTOPAT)
dado(0, 22) = "*ZC5"
dado(1, 22) = Format(ActiveSheet.Cells(FILA, 54), "0" & DGTOPAT)
dado(0, 23) = "*ZC6"
dado(1, 23) = Format(ActiveSheet.Cells(FILA, 55), "0" & DGTOPAT)
dado(0, 24) = "*ZC7"
dado(1, 24) = Format(ActiveSheet.Cells(FILA, 56), "0" & DGTOPAT)
dado(0, 25) = "*ZC8"
dado(1, 25) = Format(ActiveSheet.Cells(FILA, 125), "0" & DGTOPAT) '*-*-*

'//INCERTIDUMBRE MASA
dado(0, 26) = "*ZPM4"
dado(1, 26) = Format(ActiveSheet.Cells(FILA, 187), "0.0E+") '*-*--
dado(0, 27) = "*ZPM5"
dado(1, 27) = Format(ActiveSheet.Cells(FILA, 188), "0.0E+") '*-*--
dado(0, 28) = "*ZPM6"
dado(1, 28) = Format(ActiveSheet.Cells(FILA, 189), "0.0E+") '*-*--
dado(0, 29) = "*ZPM7"
dado(1, 29) = Format(ActiveSheet.Cells(FILA, 190), "0.0E+") '*-*--
dado(0, 30) = "*ZPM8"
dado(1, 30) = Format(ActiveSheet.Cells(FILA, 191), "0.0E+") '*-*--
dado(0, 31) = "*ZPM9"
dado(1, 31) = Format(ActiveSheet.Cells(FILA, 192), "0.0E+") '*-*--
dado(0, 32) = "*ZMP1"
dado(1, 32) = Format(ActiveSheet.Cells(FILA, 193), "0.0E+") '*-*--
dado(0, 33) = "*ZMP2"
dado(1, 33) = Format(ActiveSheet.Cells(FILA, 194), "0.0E+") '*-*--


If PUNTOSREF > 8 Then


dado(0, 34) = "*ZE1"
dado(1, 34) = Format(ActiveSheet.Cells(FILA + 1, 106), "0" & DGTOPAT) '*-*-*
dado(0, 35) = "*ZE2"
dado(1, 35) = Format(ActiveSheet.Cells(FILA + 1, 107), "0" & DGTO) '*-*-*
dado(0, 36) = "*ZE3"
dado(1, 36) = Format(ActiveSheet.Cells(FILA + 1, 107) - ActiveSheet.Cells(FILA + 1, 106), "0" & DGTOPAT) '*-*-*
dado(0, 37) = "*ZE4"
dado(1, 37) = Format(ActiveSheet.Cells(FILA + 1, 108), "0" & DGTOPAT) '*-*-*
dado(0, 38) = "*ZE5"
dado(1, 38) = Format(ActiveSheet.Cells(FILA + 1, 109), "0" & DGTO) '*-*-*
dado(0, 39) = "*ZE6"
dado(1, 39) = Format(ActiveSheet.Cells(FILA + 1, 109) - ActiveSheet.Cells(FILA + 1, 108), "0" & DGTOPAT) '*-*-*



'//INCERTIDUMBRE MASA

dado(0, 40) = "*ZMP3"
dado(1, 40) = Format(ActiveSheet.Cells(FILA + 1, 187), "0.0E+") '*-*--
dado(0, 41) = "*ZMP4"
dado(1, 41) = Format(ActiveSheet.Cells(FILA + 1, 188), "0.0E+") '*-*--


Else

End If



For S = 0 To UBound(dado, 2)
textobuscar = dado(0, S)
objWord.Selection.Move 6, -1
objWord.Selection.Find.Execute FindText:=textobuscar
While objWord.Selection.Find.found = True
objWord.Selection.Text = dado(1, S) 'texto a reemplazar
objWord.Selection.Move 6, -1
objWord.Selection.Find.Execute FindText:=textobuscar
Wend
Next S


Sheets(P_3).Select
'avanza 1 fila hacia abajo
ActiveCell.Offset(1, 0).Select
Sheets(P_3).Select
Range("GN9").Select
                            SALIDA = "Correcto"
                            

                  
            ElseIf valor = "M_g" Then           'Certificado M_gbasculas
                             patharch = ThisWorkbook.Path & "\CERTIFICADO_Mg.dotx"
        Set objWord = CreateObject("Word.Application")
        objWord.Visible = True
        objWord.documents.Add Template:=patharch, NewTemplate:=False, DocumentType:=0
        'objWord.documents.Open (patharch)
        Set Wordoc = objWord.documents.Open(patharch)
                            
                            

'///////



'If PUNTOSREF <= 8 Then

'COORDENADAS_MASA = 33

'ElseIf PUNTOSREF > 8 Then
'COORDENADAS_MASA = 41
'End If



Dim dadu(0 To 1, 0 To 41) As String '(columna,fila)


dadu(0, 0) = "*ZV3"
dadu(1, 0) = Format(ActiveSheet.Cells(FILA, 108), "0" & DGTOPAT) '*-*-*
dadu(0, 1) = "*ZV4"
dadu(1, 1) = Format(ActiveSheet.Cells(FILA, 109), "0" & DGTO) '*-*-*

''''''''''''''''_____________ERROR_4_____________
dado(0, 2) = "*ZV5"
dado(1, 2) = Format(dadu(1, 1) - dadu(1, 0), "0" & DGTOPAT) '*-*-*

'dadu(0, 2) = "*ZV5"
'dadu(1, 2) = Format(ActiveSheet.Cells(FILA, 109) - ActiveSheet.Cells(FILA, 108), "0" & DGTOPAT) '*-*-*
''''''''''''''''''''''''''''''''

dadu(0, 3) = "*ZV7"
dadu(1, 3) = Format(ActiveSheet.Cells(FILA, 110), "0" & DGTOPAT) '*-*-*
dadu(0, 4) = "*ZV8"
dadu(1, 4) = Format(ActiveSheet.Cells(FILA, 111), "0" & DGTO) '*-*-*

''''''''''''''''_____________ERROR_5_____________
dadu(0, 5) = "*ZV9"
dadu(1, 5) = Format(dadu(1, 4) - dadu(1, 3), "0" & DGTOPAT) '*-*-*

'dadu(0, 5) = "*ZV9"
'dadu(1, 5) = Format(ActiveSheet.Cells(FILA, 111) - ActiveSheet.Cells(FILA, 110), "0" & DGTOPAT) '*-*-*
''''''''''''''''''''''''''''''''''''''''

dadu(0, 6) = "*ZA1"
dadu(1, 6) = Format(ActiveSheet.Cells(FILA, 112), "0" & DGTOPAT) '*-*-*
dadu(0, 7) = "*ZA2"
dadu(1, 7) = Format(ActiveSheet.Cells(FILA, 113), "0" & DGTO) '*-*-*

''''''''''''''''_____________ERROR_6_____________
dadu(0, 8) = "*ZA3"
dadu(1, 8) = Format(dadu(1, 7) - dadu(1, 6), "0" & DGTOPAT) '*-*-*

'dadu(0, 8) = "*ZA3"
'dadu(1, 8) = Format(ActiveSheet.Cells(FILA, 113) - ActiveSheet.Cells(FILA, 112), "0" & DGTOPAT) '*-*-*
''''''''''''''''''''''''''''''''''''''''''''

dadu(0, 9) = "*ZA5"
dadu(1, 9) = Format(ActiveSheet.Cells(FILA, 114), "0" & DGTOPAT) '*-*-*
dadu(0, 10) = "*ZA6"
dadu(1, 10) = Format(ActiveSheet.Cells(FILA, 115), "0" & DGTO) '*-*-*

''''''''''''''''_____________ERROR_7_____________
dadu(0, 11) = "*ZA7"
dadu(1, 11) = Format(dadu(1, 10) - dadu(1, 9), "0" & DGTOPAT) '*-*-*

'dadu(0, 11) = "*ZA7"
'dadu(1, 11) = Format(ActiveSheet.Cells(FILA, 115) - ActiveSheet.Cells(FILA, 114), "0" & DGTOPAT) '*-*-*
'''''''''''''''''''''''''''''''''''''''''

dadu(0, 12) = "*ZA9"
dadu(1, 12) = Format(ActiveSheet.Cells(FILA, 116), "0" & DGTOPAT) '*-*-*
dadu(0, 13) = "*ZB1"
dadu(1, 13) = Format(ActiveSheet.Cells(FILA, 117), "0" & DGTO) '*-*-*

''''''''''''''''_____________ERROR_8_____________
dadu(0, 14) = "*ZB2"
dadu(1, 14) = Format(dadu(1, 13) - dadu(1, 12), "0" & DGTOPAT) '*-*-*

'dadu(0, 14) = "*ZB2"
'dadu(1, 14) = Format(ActiveSheet.Cells(FILA, 117) - ActiveSheet.Cells(FILA, 116), "0" & DGTOPAT) '*-*-*
'''''''''''''''''''''''''''''''''''''''''

dadu(0, 15) = "*ZB4"
dadu(1, 15) = Format(ActiveSheet.Cells(FILA, 118), "0" & DGTOPAT) '*-*-*
dadu(0, 16) = "*ZB5"
dadu(1, 16) = Format(ActiveSheet.Cells(FILA, 119), "0" & DGTO) '*-*-*

''''''''''''''''_____________ERROR_9_____________
dadu(0, 17) = "*ZB6"
dadu(1, 17) = Format(dadu(1, 16) - dadu(1, 15), "0" & DGTOPAT) '*-*-*

'dadu(0, 17) = "*ZB6"
'dadu(1, 17) = Format(ActiveSheet.Cells(FILA, 119) - ActiveSheet.Cells(FILA, 118), "0" & DGTOPAT) '*-*-*
'''''''''''''''''''''''''''''''''''''''''''

dadu(0, 18) = "*ZB8"
dadu(1, 18) = Format(ActiveSheet.Cells(FILA, 120), "0" & DGTOPAT) '*-*-*
dadu(0, 19) = "*ZB9"
dadu(1, 19) = Format(ActiveSheet.Cells(FILA, 121), "0" & DGTO) '*-*-*

''''''''''''''''_____________ERROR_10_____________
dadu(0, 20) = "*ZC1"
dadu(1, 20) = Format(dadu(1, 19) - dadu(1, 18), "0" & DTOPAT) '*-*-*

'dadu(0, 20) = "*ZC1"
'dadu(1, 20) = Format(ActiveSheet.Cells(FILA, 121) - ActiveSheet.Cells(FILA, 120), "0" & DGTOPAT) '*-*-*
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

dadu(0, 21) = "*ZC4"
dadu(1, 21) = Format(ActiveSheet.Cells(FILA, 53), "0" & DGTOPAT)
dadu(0, 22) = "*ZC5"
dadu(1, 22) = Format(ActiveSheet.Cells(FILA, 54), "0" & DGTOPAT)
dadu(0, 23) = "*ZC6"
dadu(1, 23) = Format(ActiveSheet.Cells(FILA, 55), "0" & DGTOPAT)
dadu(0, 24) = "*ZC7"
dadu(1, 24) = Format(ActiveSheet.Cells(FILA, 56), "0" & DGTOPAT)
dadu(0, 25) = "*ZC8"
dadu(1, 25) = Format(ActiveSheet.Cells(FILA, 125), "0" & DGTOPAT) '*-*-*

'//INCERTIDUMBRE MASA
dadu(0, 26) = "*ZPM4"
dadu(1, 26) = Format(ActiveSheet.Cells(FILA, 187), "0.0E+") '*-*--
dadu(0, 27) = "*ZPM5"
dadu(1, 27) = Format(ActiveSheet.Cells(FILA, 188), "0.0E+") '*-*--
dadu(0, 28) = "*ZPM6"
dadu(1, 28) = Format(ActiveSheet.Cells(FILA, 189), "0.0E+") '*-*--
dadu(0, 29) = "*ZPM7"
dadu(1, 29) = Format(ActiveSheet.Cells(FILA, 190), "0.0E+") '*-*--
dadu(0, 30) = "*ZPM8"
dadu(1, 30) = Format(ActiveSheet.Cells(FILA, 191), "0.0E+") '*-*--
dadu(0, 31) = "*ZPM9"
dadu(1, 31) = Format(ActiveSheet.Cells(FILA, 192), "0.0E+") '*-*--
dadu(0, 32) = "*ZMP1"
dadu(1, 32) = Format(ActiveSheet.Cells(FILA, 193), "0.0E+") '*-*--
dadu(0, 33) = "*ZMP2"
dadu(1, 33) = Format(ActiveSheet.Cells(FILA, 194), "0.0E+") '*-*--


If PUNTOSREF > 8 Then


dadu(0, 34) = "*ZE1"
dadu(1, 34) = Format(ActiveSheet.Cells(FILA + 1, 106), "0" & DGTOPAT) '*-*-*
dadu(0, 35) = "*ZE2"
dadu(1, 35) = Format(ActiveSheet.Cells(FILA + 1, 107), "0" & DGTO) '*-*-*
dadu(0, 36) = "*ZE3"
dadu(1, 36) = Format(ActiveSheet.Cells(FILA + 1, 107) - ActiveSheet.Cells(FILA + 1, 106), "0" & DGTOPAT) '*-*-*
dadu(0, 37) = "*ZE4"
dadu(1, 37) = Format(ActiveSheet.Cells(FILA + 1, 108), "0" & DGTOPAT) '*-*-*
dadu(0, 38) = "*ZE5"
dadu(1, 38) = Format(ActiveSheet.Cells(FILA + 1, 109), "0" & DGTO) '*-*-*
dadu(0, 39) = "*ZE6"
dadu(1, 39) = Format(ActiveSheet.Cells(FILA + 1, 109) - ActiveSheet.Cells(FILA + 1, 108), "0" & DGTOPAT) '*-*-*



'//INCERTIDUMBRE MASA

dadu(0, 40) = "*ZMP3"
dadu(1, 40) = Format(ActiveSheet.Cells(FILA + 1, 187), "0.0E+") '*-*--
dadu(0, 41) = "*ZMP4"
dadu(1, 41) = Format(ActiveSheet.Cells(FILA + 1, 188), "0.0E+") '*-*--


Else

End If


For S = 0 To UBound(dadu, 2)
textobuscar = dadu(0, S)
objWord.Selection.Move 6, -1
objWord.Selection.Find.Execute FindText:=textobuscar
While objWord.Selection.Find.found = True
objWord.Selection.Text = dadu(1, S) 'texto a reemplazar
objWord.Selection.Move 6, -1
objWord.Selection.Find.Execute FindText:=textobuscar
Wend
Next S


Sheets(P_3).Select
'avanza 1 fila hacia abajo
ActiveCell.Offset(1, 0).Select
Sheets(P_3).Select
Range("GN9").Select
                            SALIDA = "Correcto"
                   

              ElseIf valor = "N5-11-ITH-005_H VAISALA" Or valor = "N5-11-ITH-005_T VAISALA" Then            'Certificado H_MT12001
                                
                                                                
                                patharch = ThisWorkbook.Path & "\CERTIFICADO N5-11-ITH-005.dotx"
                                Set objWord = CreateObject("Word.Application")
                                objWord.Visible = True
                                objWord.documents.Add Template:=patharch, NewTemplate:=False, DocumentType:=0
                                SALIDA = "Correcto"
                                                              
                                Set Wordoc = objWord.documents.Open(patharch)
                                                              
              ElseIf valor = "N5-11-ITH-011_H VAISALA" Or valor = "N5-11-ITH-011_T VAISALA" Then            'Certificado H_MT12001
                                
                             
                                patharch = ThisWorkbook.Path & "\CERTIFICADO N5-11-ITH-011.dotx"
                                Set objWord = CreateObject("Word.Application")
                                objWord.Visible = True
                                objWord.documents.Add Template:=patharch, NewTemplate:=False, DocumentType:=0
                                SALIDA = "Correcto"
                           
                                Set Wordoc = objWord.documents.Open(patharch)
                           
              ElseIf valor = "N5-11-IA-002 FLUKEAMP" Then              'Certificado Electrica Corriente
                                    
  
                                    
                                    patharch = ThisWorkbook.Path & "\CERTIFICADO N5-11-IA-002 A.dotx"
                                    Set objWord = CreateObject("Word.Application")
                                    objWord.Visible = True
                                    objWord.documents.Add Template:=patharch, NewTemplate:=False, DocumentType:=0
                                    SALIDA = "Correcto"
                                    
                                    Set Wordoc = objWord.documents.Open(patharch)
                                                                  
              ElseIf valor = "N5-11-IA-002 V" Then              'Certificado Electrica Corriente
                                        

                                        
                                        patharch = ThisWorkbook.Path & "\Certificado N5-11-IA-002 V.dotx"
                                        Set objWord = CreateObject("Word.Application")
                                        objWord.Visible = True
                                        objWord.documents.Add Template:=patharch, NewTemplate:=False, DocumentType:=0
                                        SALIDA = "Correcto"
                                        
                                        Set Wordoc = objWord.documents.Open(patharch)
                                   
              ElseIf valor = "N5-11-IV-002 FLUKEVOL" Then                 'Certificado Electrica voltaje
                                         
                   
                                         patharch = ThisWorkbook.Path & "\CERTIFICADO N5-11-IV-002.dotx"
                                         Set objWord = CreateObject("Word.Application")
                                         objWord.Visible = True
                                         objWord.documents.Add Template:=patharch, NewTemplate:=False, DocumentType:=0
                                         SALIDA = "Correcto"
                                         
                                         Set Wordoc = objWord.documents.Open(patharch)
                                     
              ElseIf valor = "N5-11-IV-003 FLUKEKVOL" Then                'Certificado Electrica Kv
                                            
                                            
                                            patharch = ThisWorkbook.Path & "\CERTIFICADO N5-11-IV-003.dotx"
                                            Set objWord = CreateObject("Word.Application")
                                            objWord.Visible = True
                                            objWord.documents.Add Template:=patharch, NewTemplate:=False, DocumentType:=0
                                            SALIDA = "Correcto"
                                            
                                            Set Wordoc = objWord.documents.Open(patharch)
                                    
              ElseIf valor = "N5-11-FJ-001 SMC" Then                'Certificado FLUJO
                                            
                                            
                                            patharch = ThisWorkbook.Path & "\CERTIFICADO N5-11-FJ-001.dotx"
                                            Set objWord = CreateObject("Word.Application")
                                            objWord.Visible = True
                                            objWord.documents.Add Template:=patharch, NewTemplate:=False, DocumentType:=0
                                            SALIDA = "Correcto"
                                            
                                            Set Wordoc = objWord.documents.Open(patharch)
                                            
                ElseIf valor = "N5-11-FJ-002" Then                'Certificado FLUJO
                                            
                                            
                                            patharch = ThisWorkbook.Path & "\CERTIFICADO N5-11-FJ-002.dotx"
                                            Set objWord = CreateObject("Word.Application")
                                            objWord.Visible = True
                                            objWord.documents.Add Template:=patharch, NewTemplate:=False, DocumentType:=0
                                            SALIDA = "Correcto"
                                            
                                            Set Wordoc = objWord.documents.Open(patharch)
                
                                     
              ElseIf valor = "N5-11-TC-001 RPM" Or valor = "N5-11-TC-001 Imts" Or valor = "N5-11-TC-001 lHz" Then 'Certificado TACOMETRO
                                            
                                           
                                            patharch = ThisWorkbook.Path & "\CERTIFICADO N5-11-TC-001.dotx"
                                            Set objWord = CreateObject("Word.Application")
                                            objWord.Visible = True
                                            objWord.documents.Add Template:=patharch, NewTemplate:=False, DocumentType:=0
                                            SALIDA = "Correcto"
                                            
                                            Set Wordoc = objWord.documents.Open(patharch)
                                     
              ElseIf valor = "N5-11-IA-005 Sim_mV" Then           'Certificado SIM mV
                                            
                                           patharch = ThisWorkbook.Path & "\CERTIFICADO N5-11-IA-005.dotx"
                                            Set objWord = CreateObject("Word.Application")
                                            objWord.Visible = True
                                            objWord.documents.Add Template:=patharch, NewTemplate:=False, DocumentType:=0
                                            
                                            Set Wordoc = objWord.documents.Open(patharch)
     

                                                                            
Dim dude(0 To 1, 0 To 14) As String '(columna,fila)
dude(0, 0) = "*ZV3"
dude(1, 0) = ActiveSheet.Cells(FILA, 25)
dude(0, 1) = "*ZV4"
dude(1, 1) = ActiveSheet.Cells(FILA, 29)

'dude(0, 2) = "*ZV5"
'dude(1, 2) = ActiveSheet.Cells(FILA, 33)

dude(0, 3) = "*ZV6"
dude(1, 3) = Round((ActiveSheet.Cells(FILA, 29) + ActiveSheet.Cells(FILA, 33)), 3)
dude(0, 4) = "*ZV7"
dude(1, 4) = ActiveSheet.Cells(FILA, 26)
dude(0, 5) = "*ZV8"
dude(1, 5) = ActiveSheet.Cells(FILA, 30)

'dude(0, 6) = "*ZV9"
'dude(1, 6) = ActiveSheet.Cells(FILA, 34)

dude(0, 7) = "*ZA1"
dude(1, 7) = Round((ActiveSheet.Cells(FILA, 30) + ActiveSheet.Cells(FILA, 34)), 3)
dude(0, 8) = "*ZA2"
dude(1, 8) = ActiveSheet.Cells(FILA, 27)

dude(0, 9) = "*ZA3"
dude(1, 9) = ActiveSheet.Cells(FILA, 31)
dude(0, 10) = "*ZA4"
dude(1, 10) = ActiveSheet.Cells(FILA, 35)
dude(0, 11) = "*ZA5"
dude(1, 11) = Round((ActiveSheet.Cells(FILA, 31) + ActiveSheet.Cells(FILA, 35)), 3)
dude(0, 12) = "*ZA6"
dude(1, 12) = ActiveSheet.Cells(FILA, 36)
dude(0, 13) = "*ZA7"
dude(1, 13) = ActiveSheet.Cells(FILA, 37)
dude(0, 14) = "*ZA8"
dude(1, 14) = ActiveSheet.Cells(FILA, 38)
For R = 0 To UBound(dude, 2)
textobuscar = dude(0, R)
objWord.Selection.Move 6, -1
objWord.Selection.Find.Execute FindText:=textobuscar
While objWord.Selection.Find.found = True
objWord.Selection.Text = dude(1, R) 'texto a reemplazar
objWord.Selection.Move 6, -1
objWord.Selection.Find.Execute FindText:=textobuscar
Wend
Next R

'vuelve a la hoja original
Sheets(P_3).Select
'avanza 1 fila hacia abajo
ActiveCell.Offset(1, 0).Select

                            SALIDA = "Correcto"
''-----------------------------------------------------
    
    
    
    End If
    
    
''''''''''--------------------  Envio de Datos a WORD------------------------------
On Error GoTo 0
                                    Dim MAG(0 To 1, 0 To 6) As String
                                    
                                    MAG(0, 0) = "*MAGNITUD*"
                                    MAG(1, 0) = MAGNITUD_CERTIFICADO
                                    MAG(0, 1) = "*PROCEDIMIENTO*"
                                    MAG(1, 1) = PROCEDIMIENTO
                                    MAG(0, 2) = "*NORMA*"
                                    MAG(1, 2) = NORMA
                                    MAG(0, 3) = "*T_EM*"
                                    MAG(1, 3) = TIPO_E_M
                                    MAG(0, 4) = "*ERROR*"
                                    MAG(1, 4) = REPORTE_ERROR
                                    MAG(0, 5) = "*REALIZA*"
                                    MAG(1, 5) = INGENIERO
                                    MAG(0, 6) = "*REVISA*"
                                    MAG(1, 6) = REVISA
                                    
                                    For I = 0 To UBound(MAG, 2)
                                    textobuscar = MAG(0, I)
                                    objWord.Selection.Move 6, -1
                                    objWord.Selection.Find.Execute FindText:=textobuscar
                                    While objWord.Selection.Find.found = True
                                    objWord.Selection.Text = MAG(1, I) 'texto a reemplazar
                                    objWord.Selection.Move 6, -1
                                    objWord.Selection.Find.Execute FindText:=textobuscar
                                    Wend
                                    Next I
    
    
    
Dim dede(0 To 1, 0 To 27) As String '(columna,fila)
dede(0, 0) = "*ZV3"
dede(1, 0) = Format(ActiveSheet.Cells(FILA, 108), "0" & DGTOPAT) '*-*-*
dede(0, 1) = "*ZV4"
dede(1, 1) = Format(ActiveSheet.Cells(FILA, 109), "0" & DGTO) '*-*-*

''''''''''''''''_____________ERROR_4_____________
dede(0, 2) = "*ZV5"
dede(1, 2) = Format(dede(1, 1) - dede(1, 0), "0" & DGTOPAT) '*-*-*

'dede(0, 2) = "*ZV5"
'dede(1, 2) = Format(ActiveSheet.Cells(FILA, 109) - ActiveSheet.Cells(FILA, 108), "0" & DGTOPAT) '*-*-*
'''''''''''''''''''''''''''''''''''''''''''''''''

dede(0, 3) = "*ZV7"
dede(1, 3) = Format(ActiveSheet.Cells(FILA, 110), "0" & DGTOPAT) '*-*-*
dede(0, 4) = "*ZV8"
dede(1, 4) = Format(ActiveSheet.Cells(FILA, 111), "0" & DGTO) '*-*-*

''''''''''''''''_____________ERROR_5_____________

dede(0, 5) = "*ZV9"
dede(1, 5) = Format(dede(1, 4) - dede(1, 3), "0" & DGTOPAT) '*-*-*

'dede(0, 5) = "*ZV9"
'dede(1, 5) = Format(ActiveSheet.Cells(FILA, 111) - ActiveSheet.Cells(FILA, 110), "0" & DGTOPAT) '*-*-*
'''''''''''''''''''''''''''''''''''''''''''

dede(0, 6) = "*ZA1"
dede(1, 6) = Format(ActiveSheet.Cells(FILA, 112), "0" & DGTOPAT) '*-*-*
dede(0, 7) = "*ZA2"
dede(1, 7) = Format(ActiveSheet.Cells(FILA, 113), "0" & DGTO) '*-*-*

''''''''''''''''_____________ERROR_6_____________
dede(0, 8) = "*ZA3"
dede(1, 8) = Format(dede(1, 7) - dede(1, 6), "0" & DGTOPAT) '*-*-*

'dede(0, 8) = "*ZA3"
'dede(1, 8) = Format(ActiveSheet.Cells(FILA, 113) - ActiveSheet.Cells(FILA, 112), "0" & DGTOPAT) '*-*-*
''''''''''''''''''''''''''''''''''''''''
dede(0, 9) = "*ZA5"
dede(1, 9) = Format(ActiveSheet.Cells(FILA, 114), "0" & DGTOPAT) '*-*-*
dede(0, 10) = "*ZA6"
dede(1, 10) = Format(ActiveSheet.Cells(FILA, 115), "0" & DGTO) '*-*-*

''''''''''''''''_____________ERROR_7_____________
dede(0, 11) = "*ZA7"
dede(1, 11) = Format(dede(1, 10) - dede(1, 9), "0" & DGTOPAT) '*-*-*

'dede(0, 11) = "*ZA7"
'dede(1, 11) = Format(ActiveSheet.Cells(FILA, 115) - ActiveSheet.Cells(FILA, 114), "0" & DGTOPAT) '*-*-*
'''''''''''''''''''''''''''''''''''

dede(0, 12) = "*ZA9"
dede(1, 12) = Format(ActiveSheet.Cells(FILA, 116), "0" & DGTOPAT) '*-*-*
dede(0, 13) = "*ZB1"
dede(1, 13) = Format(ActiveSheet.Cells(FILA, 117), "0" & DGTO) '*-*-*

''''''''''''''''_____________ERROR_8_____________
dede(0, 14) = "*ZB2"
dede(1, 14) = Format(dede(1, 13) - dede(1, 12), "0" & DGTOPAT)

'dede(0, 14) = "*ZB2"
'dede(1, 14) = Format(ActiveSheet.Cells(FILA, 117) - ActiveSheet.Cells(FILA, 116), "0" & DGTOPAT)
'''''''''''''''''''''''''''''''''''''''''''''''''

dede(0, 15) = "*ZB4"
dede(1, 15) = Format(ActiveSheet.Cells(FILA, 118), "0" & DGTOPAT)
dede(0, 16) = "*ZB5"
dede(1, 16) = Format(ActiveSheet.Cells(FILA, 119), "0" & DGTO)

''''''''''''''''_____________ERROR_9_____________
dede(0, 17) = "*ZB6"
dede(1, 17) = Format(dede(1, 16) - dede(1, 15), "0" & DGTOPAT) '*-*-*

'dede(0, 17) = "*ZB6"
'dede(1, 17) = Format(ActiveSheet.Cells(FILA, 119) - ActiveSheet.Cells(FILA, 118), "0" & DGTOPAT) '*-*-*
''''''''''''''''''''''''''''''''''''''

dede(0, 18) = "*ZB8"
dede(1, 18) = Format(ActiveSheet.Cells(FILA, 120), "0" & DGTOPAT) '*-*-*
dede(0, 19) = "*ZB9"
dede(1, 19) = Format(ActiveSheet.Cells(FILA, 121), "0" & DGTO) '*-*-*

''''''''''''''''_____________ERROR_10_____________
dede(0, 20) = "*ZC1"
dede(1, 20) = Format(dede(1, 19) - dede(1, 18), "0" & DGTOPAT) '*-*-*

'dede(0, 20) = "*ZC1"
'dede(1, 20) = Format(ActiveSheet.Cells(FILA, 121) - ActiveSheet.Cells(FILA, 120), "0" & DGTOPAT) '*-*-*
''''''''''''''''''''''''''''''''''''''''''''''

dede(0, 21) = "*ZC4"
dede(1, 21) = ActiveSheet.Cells(FILA, 108)  '*-*-*
dede(0, 22) = "*ZC5"
dede(1, 22) = ActiveSheet.Cells(FILA, 110)  '*-*-*
dede(0, 23) = "*ZC6"
dede(1, 23) = ActiveSheet.Cells(FILA, 112)  '*-*-*
dede(0, 24) = "*ZC7"
dede(1, 24) = ActiveSheet.Cells(FILA, 114)  '*-*-*
dede(0, 25) = "*ZC8"
dede(1, 25) = ActiveSheet.Cells(FILA, 116)  '*-*-*
dede(0, 26) = "*ZD3"
dede(1, 26) = ActiveSheet.Cells(FILA, 118)  '*-*-*
dede(0, 27) = "*ZD4"
dede(1, 27) = ActiveSheet.Cells(FILA, 120)  '*-*-*

For N = 0 To UBound(dede, 2)
textobuscar = dede(0, N)
objWord.Selection.Move 6, -1
objWord.Selection.Find.Execute FindText:=textobuscar
While objWord.Selection.Find.found = True
objWord.Selection.Text = dede(1, N) 'texto a reemplazar
objWord.Selection.Move 6, -1
objWord.Selection.Find.Execute FindText:=textobuscar
Wend
Next N
'objWord.Activate
'ActiveSheet.Cells(fila, 188).Value = salida

'fila = fila + 1
'vuelve a la hoja original
Sheets(P_3).Select
'avanza 1 fila hacia abajo
ActiveCell.Offset(1, 0).Select

    
    Dim datos(0 To 1, 0 To 51) As String '(columna,fila)
    datos(0, 0) = "*X1"
    datos(1, 0) = ActiveSheet.Cells(FILA, 99) '(fila,columna)   *-*-*
    
    '''''UBICACIÓN
    datos(0, 1) = "*X2"
    datos(1, 1) = UBICACION_ID
    'datos(1, 1) = ActiveSheet.Cells(FILA, 201)'*-*-*
    'REFERENCIA
'datos(1, 1) = ActiveSheet.Cells(FILA, 100) '*-*-*
    'datos(0, 2) = "*X3"
    'datos(1, 2) = ActiveSheet.Cells(1, 10)
    datos(0, 2) = "*X3"
    datos(1, 2) = FECHA_C
    
''''''''''''''''''''''''''''''''''''''''''' FECHA ''''''''''''''''''''''

'TBL_FECHA = Sheets(P_1).Range("A2:Z45")
Sheets(P_3).Select

'V_FECHA = Application.VLookup(ID, TBL_FECHA, 26, False)


'If IsError(lookupvalue) Then
'MsgBox "USUARIO NO AUTORIZADO"
'Exit Sub

'Else

'REALIZA = V_USUARIO
'End If

    'datos(0, 3) = "*X5"
    'FECHA_C = Application.InputBox("FECHA DE CALIBRACIÓN", "FECHA", ActiveSheet.Cells(1, 10))
    datos(0, 3) = "*X4"
    datos(1, 3) = FECHA_C

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    

    
    'FECHA_GENREACIÓN
    datos(0, 4) = "*X5"
    datos(1, 4) = FECHA_GENERACION
    
    '3 DÍAS HABILES
    'datos(1, 4) = ActiveSheet.Cells(1, 13)
datos(0, 5) = "*X6"
datos(1, 5) = ActiveSheet.Cells(3, 4)
datos(0, 6) = "*X7"
datos(1, 6) = ActiveSheet.Cells(4, 4)
datos(0, 7) = "*X8"
datos(1, 7) = ActiveSheet.Cells(FILA, 2)
datos(0, 8) = "*X9"
datos(1, 8) = ActiveSheet.Cells(FILA, 4)
datos(0, 9) = "*Y1"
datos(1, 9) = ActiveSheet.Cells(FILA, 5)
datos(0, 10) = "*Y2"
datos(1, 10) = ActiveSheet.Cells(FILA, 6)
datos(0, 11) = "*Y3"
datos(1, 11) = ActiveSheet.Cells(FILA, 7)
datos(0, 12) = "*ZV1"
datos(1, 12) = ActiveSheet.Cells(FILA, 8)
datos(0, 13) = "*Y4"
datos(1, 13) = ActiveSheet.Cells(FILA, 1)
datos(0, 14) = "*Y5"
datos(1, 14) = ActiveSheet.Cells(FILA, 11)
datos(0, 15) = "*Y6"
datos(1, 15) = ActiveSheet.Cells(FILA, 12)
datos(0, 16) = "*Y7"
datos(1, 16) = ActiveSheet.Cells(2, 10)
datos(0, 17) = "*Y8"
datos(1, 17) = ActiveSheet.Cells(FILA, 102)  '*-*-*

''''''''''''''''_____________PATRÓN_P1_____________
datos(0, 18) = "*Y9"
datos(1, 18) = Format(ActiveSheet.Cells(FILA, 102), "0" & DGTOPAT)   '*-*-*

datos(0, 19) = "*Z1"
datos(1, 19) = Format(ActiveSheet.Cells(FILA, 103), "0" & DGTO)   '*-*-*

''''''''''''''''_____________ERROR_1_____________
datos(0, 20) = "*Z2"
datos(1, 20) = Format(datos(1, 19) - datos(1, 18), "0" & DGTOPAT)  '*-*-*

'datos(0, 20) = "*Z2"
'datos(1, 20) = Format(ActiveSheet.Cells(FILA, 103) - ActiveSheet.Cells(FILA, 102), "0" & DGTOPAT)  '*-*-*
''''''''''''''''''''''''''''''

datos(0, 21) = "*Z3"
datos(1, 21) = Round(ActiveSheet.Cells(FILA, 185), 2) '*-*-*
datos(0, 22) = "*Z4"
datos(1, 22) = ActiveSheet.Cells(FILA, 104) '*-*-*
datos(0, 23) = "*Z5"
datos(1, 23) = Format(ActiveSheet.Cells(FILA, 104), "0" & DGTOPAT) '*-*-*
datos(0, 24) = "*Z6"
datos(1, 24) = Format(ActiveSheet.Cells(FILA, 105), "0" & DGTO) '*-*-*

''''''''''''''''_____________ERROR_2_____________
datos(0, 25) = "*Z7"
datos(1, 25) = Format(datos(1, 24) - datos(1, 23), "0" & DGTOPAT)


'datos(0, 25) = "*Z7"
'datos(1, 25) = Format(ActiveSheet.Cells(FILA, 105) - ActiveSheet.Cells(FILA, 104), "0" & DGTOPAT) '*-*-*
''''''''''''''''''''''''''''''''''''''

datos(0, 26) = "*Z8"
datos(1, 26) = Round(ActiveSheet.Cells(FILA, 186), 2) '*-*-*-
datos(0, 27) = "*Z9"
datos(1, 27) = ActiveSheet.Cells(FILA, 106) '*-*-*
datos(0, 28) = "*W1"
datos(1, 28) = Format(ActiveSheet.Cells(FILA, 106), "0" & DGTOPAT) '*-*-*
datos(0, 29) = "*W2"
datos(1, 29) = Format(ActiveSheet.Cells(FILA, 107), "0" & DGTO) '*-*-*

''''''''''''''''_____________ERROR_3_____________
datos(0, 30) = "*W3"
datos(1, 30) = Format(datos(1, 29) - datos(1, 27), "0" & DGTOPAT) '*-*-*

'datos(0, 30) = "*W3"
'datos(1, 30) = Format(ActiveSheet.Cells(FILA, 107) - ActiveSheet.Cells(FILA, 106), "0" & DGTOPAT) '*-*-*
''''''''''''''''''''''''''''''''''''

datos(0, 31) = "*W4"
datos(1, 31) = Round(ActiveSheet.Cells(FILA, 187), 2)  '*-*--

''''''''''''''--------------------FORMULARIO------------------'''''''''''''
If TIPO = "IP" Or TIPO = "CP" Or TIPO = "MB" Or TIPO = "VA" Then
datos(0, 32) = "*W5"
datos(1, 32) = EMR
'ElseIf TIPO = "CP" Or TIPO = "CH" Then
'datos(0, 32) = "*W5"
'datos(1, 32) = EMR
Else
datos(0, 32) = "*W5"
datos(1, 32) = Round(ActiveSheet.Cells(FILA, 199), 2) '*-*-*
End If

''''''''''''''--------------------FORMULARIO------------------'''''''''''''

datos(0, 33) = "*W6"
datos(1, 33) = Format(ActiveSheet.Cells(FILA, 95), "0.0")
datos(0, 34) = "*W7"
datos(1, 34) = Format(ActiveSheet.Cells(FILA, 96), "0.0")
datos(0, 35) = "*W8"
datos(1, 35) = Format(ActiveSheet.Cells(FILA, 93), "0.0")
datos(0, 36) = "*W9"
datos(1, 36) = Format(ActiveSheet.Cells(FILA, 94), "0.0")
datos(0, 37) = "*W0"
datos(1, 37) = ActiveSheet.Cells(FILA, 101) '*-*-*
datos(0, 38) = "*ZV2"
datos(1, 38) = ActiveSheet.Cells(FILA, 9)
datos(0, 39) = "*ZC2"
datos(1, 39) = ActiveSheet.Cells(FILA, 196) '*-*-*
datos(0, 40) = "*ZC3"
datos(1, 40) = ActiveSheet.Cells(FILA, 200) '*-*-*-
datos(0, 41) = "*ZC9"
datos(1, 41) = Format(ActiveSheet.Cells(6, 4), "0000")
datos(0, 42) = "*ZD1"
datos(1, 42) = Format(ActiveSheet.Cells(FILA, 97), "0.0")
datos(0, 43) = "*ZD2"
datos(1, 43) = Format(ActiveSheet.Cells(FILA, 98), "0.0")

'''''''''''''''''''''''''_______EMP______'''''''''''''''''''''''''''

datos(0, 44) = "*ZD5"
datos(1, 44) = ActiveSheet.Cells(FILA, 12)
datos(0, 45) = "*ZD6"
datos(1, 45) = Round(ActiveSheet.Cells(FILA, 188), 2)
datos(0, 46) = "*ZD7"
datos(1, 46) = Round(ActiveSheet.Cells(FILA, 189), 2)
datos(0, 47) = "*ZD8"
datos(1, 47) = Round(ActiveSheet.Cells(FILA, 190), 2)
datos(0, 48) = "*ZD9"
datos(1, 48) = Round(ActiveSheet.Cells(FILA, 191), 2)
datos(0, 49) = "*ZP1"
datos(1, 49) = Round(ActiveSheet.Cells(FILA, 192), 2)
datos(0, 50) = "*ZP2"
datos(1, 50) = Round(ActiveSheet.Cells(FILA, 193), 2)
datos(0, 51) = "*ZP3"
datos(1, 51) = Round(ActiveSheet.Cells(FILA, 194), 2)

'//INCERTIDUMBRE MASA
'datos(0, 52) = "*ZP4"
'datos(1, 52) = Format(ActiveSheet.Cells(FILA, 187), "0.0E+") '*-*--

For I = 0 To UBound(datos, 2)
textobuscar = datos(0, I)
objWord.Selection.Move 6, -1
objWord.Selection.Find.Execute FindText:=textobuscar
While objWord.Selection.Find.found = True
objWord.Selection.Text = datos(1, I) 'texto a reemplazar
objWord.Selection.Move 6, -1
objWord.Selection.Find.Execute FindText:=textobuscar
Wend
Next I


'objWord.Activate
ActiveSheet.Cells(FILA, 198).value = SALIDA '*-*-*

'///// QR
Call QR
'///
ActiveSheet.Shapes("NOMQR").CopyPicture
textobuscar = "*ZP5"
objWord.Selection.Move 6, -1
objWord.Selection.Find.Execute FindText:=textobuscar
objWord.Selection.Range.Paste
Selection.ShapeRange(1).Delete

'///////////////// GRÁFICA
Call GRAFICAR
ActiveSheet.Shapes("comportamiento").CopyPicture
textobuscar = "*ZP6"
objWord.Selection.Move 6, -1
objWord.Selection.Find.Execute FindText:=textobuscar
objWord.Selection.Range.Paste

''''''''-------- DATOS U (UNIDADES) -----------


If TIPO = "IP" Or TIPO = "CP" Or TIPO = "MB" Or TIPO = "VA" Then

If UNIDADES = "MPa" Or UNIDADES = "Mpa" Or UNIDADES = "MPA" Or UNIDADES = "mpa" Then
    U_DGT = 5
ElseIf UNIDADES = "kg/cm2" Or UNIDADES = "Kg/cm2" Or UNIDADES = "KG/CM2" Then
    U_DGT = 4
Else
    U_DGT = 3
End If

Dim UNID(0 To 1, 0 To 9) As String '(columna,fila)

On Error Resume Next


'''''''''''''''''''INCERTIDUMBRE'''''''''''''''''

UNID(0, 0) = "*ZQ0"
UNID(1, 0) = Round(ActiveSheet.Cells(10, 208), U_DGT)
UNID(0, 1) = "*ZQ1"
UNID(1, 1) = Round(ActiveSheet.Cells(11, 208), U_DGT)
UNID(0, 2) = "*ZQ2"
UNID(1, 2) = Round(ActiveSheet.Cells(12, 208), U_DGT)
UNID(0, 3) = "*ZQ3"
UNID(1, 3) = Round(ActiveSheet.Cells(13, 208), U_DGT)
UNID(0, 4) = "*ZQ4"
UNID(1, 4) = Round(ActiveSheet.Cells(14, 208), U_DGT)
UNID(0, 5) = "*ZQ5"
UNID(1, 5) = Round(ActiveSheet.Cells(15, 208), U_DGT)
UNID(0, 6) = "*ZQ6"
UNID(1, 6) = Round(ActiveSheet.Cells(16, 208), U_DGT)
UNID(0, 7) = "*ZQ7"
UNID(1, 7) = Round(ActiveSheet.Cells(17, 208), U_DGT)
UNID(0, 8) = "*ZQ8"
UNID(1, 8) = Round(ActiveSheet.Cells(18, 208), U_DGT)
UNID(0, 9) = "*ZQ9"
UNID(1, 9) = Round(ActiveSheet.Cells(19, 208), U_DGT)
On Error GoTo 0

For N = 0 To UBound(UNID, 2)
textobuscar = UNID(0, N)
objWord.Selection.Move 6, -1
objWord.Selection.Find.Execute FindText:=textobuscar
While objWord.Selection.Find.found = True
objWord.Selection.Text = UNID(1, N) 'texto a reemplazar
objWord.Selection.Move 6, -1
objWord.Selection.Find.Execute FindText:=textobuscar
Wend
Next N


End If


'''''''''''Call GUARDAR


strNombreArchivo = ActiveSheet.Cells(FILA, 100).value
REFERENCIA = Replace(strNombreArchivo, "/", ".")


GUARDAR_CERT = ThisWorkbook.Path & "\" & SERV_CERT & REFERENCIA & ".docx"


Wordoc.SaveAs GUARDAR_CERT, _
    FileFormat:=wdFormatDocumentDefault


objWord.Quit


'''''''''''''''''

graf.Delete

Sheets(P_3).Select

        SALIDA = "Correcto"
ActiveSheet.Cells(FILA, 198).value = SALIDA


'''''''''''''''EN MASA SE UTILIZA LA SIGUIENTE FILA PARA LOS
'''''''''''''' DATOS DEL PUNTO 9 Y 10
If valor <> "M_g" Or valor <> "M_Kg" Then

    If PUNTOSREF > 8 Then
    FILA = FILA + 2

    Else
    FILA = FILA + 1
    End If
Else
FILA = FILA + 1

End If


Unload FORMULARIO

'vuelve a la hoja original
Sheets(P_3).Select
ActiveSheet.Cells(FILA, 197).Select

Range("GW4:HC20").value = " "

Range("GT9").value = " "


'''''''''''

Wend

Worksheets(P_3).Protect "MET2025"
Sheets(P_3).Select
Application.ScreenUpdating = True
ActiveSheet.Cells(FILA, 198).value = "REALIZADO"

GoTo FIN

LIMPIAR:

MsgBox "SE DETUVO LA GENERACIÓN DE CERTIFICADOS, POSIBLE ERROR"
Worksheets(P_3).Protect "MET2025"

Range("GW4:HC20").value = " "
On Error Resume Next
Selection.ShapeRange(1).Delete
On Error GoTo 0

On Error Resume Next
graf.Delete
On Error GoTo 0
Columns("GS:HH").Hidden = True
Application.ScreenUpdating = True
Exit Sub

PLANTILLA:
MsgBox "PROBLEMA ENCONTRADO CON LAS PLANTILLAS. VERIRICAR QUE EXISTA LA PLANTILLA PARA EL PATRÓN: " & valor

Worksheets(P_3).Unprotect "MET2025"
Columns("GS:HH").Hidden = True
Worksheets(P_3).Protect "MET2025"
Exit Sub

FIN:
On Error Resume Next


Range("GW4:HC20").value = " "
On Error Resume Next
Selection.ShapeRange(1).Delete
On Error GoTo 0

On Error Resume Next
graf.Delete
On Error GoTo 0

Worksheets(P_3).Unprotect "MET2025"
Columns("GS:HH").Hidden = True


Application.ScreenUpdating = True
MsgBox "CERTIFICADOS REALIZADOS"
On Error GoTo 0

Range("GS10:GS46").value = " "

Worksheets(P_3).Protect "MET2025"

Exit Sub

End Sub

Sub GRAFICAR()
'///////////////// GRÁFICA

'REFERENCIA

Range("GX6").value = "206"
Range("GY6").value = "207"
Range("GZ6").value = "208"
Range("GX8").value = "ERROR VS INCERTUDUMBRE"
Range("GX9").value = "PUNTO REFERENCIA"
Range("GY9").value = "ERROR"
Range("GZ9").value = "INCERTIDUMBRE"

'FILA=10
'''''''''''''''RESOLUCIÓN GRAFICA

DGTOPAT = FORMULARIO.TXBPATRON.value
DGTO = FORMULARIO.TXBINSTRUMENTO.value

''''''''''''''''''''''''''''''''''''''''''

If PATRON = "M_Kg" Or PATRON = "M_g" Then
RESGRAF = ActiveSheet.Cells(FILA, 37).value - ActiveSheet.Cells(FILA, 33).value

Else
RESGRAF = Abs(ActiveSheet.Cells(FILA, 17).value - ActiveSheet.Cells(FILA, 13).value)
End If

For I = 1 To 10
Range("GW9").Offset(I, 0).value = I
Next I

Range("HB5").FormulaLocal = "=MAX(GX10:GX" & 9 + PUNTOSREF & ")+" & RESGRAF & " "
'Range("HA5").FormulaLocal = "=MIN(GX10:GX" & 9 + PUNTOSREF & ") -" & RESGRAF & ""
Range("HA5").FormulaLocal = "=MIN(GX10:GX" & 9 + PUNTOSREF & ")"




    If PATRON = "M_Kg" Or PATRON = "M_g" Then
    
    U = 10
    
    For I = 1 To PUNTOSREF
    
        If I <= 8 Then
        'Format(ActiveSheet.Cells(FILA, 104), "0" & DGTOPAT)
        
        If FIJO_PATRON = "IBC" Then
        PUNTO = Format(ActiveSheet.Cells(FILA, 106).Offset(0, 2 * I - 1).value, "0" & DGTO)
        
        ElseIf FIJO_PATRON = "PATRON" Then
        PUNTO = Format(ActiveSheet.Cells(FILA, 105).Offset(0, 2 * I - 1).value, "0" & DGTOPAT)
        End If
        
        PUNTO_INDICACION = Format(ActiveSheet.Cells(FILA, 105).Offset(0, 2 * I).value, "0" & DGTO)
        PUNTO_REFERENCIA = ActiveSheet.Cells(FILA, 105).Offset(0, 2 * I - 1).value
        
        PUNTOERR = PUNTO_INDICACION - PUNTO_REFERENCIA
        PUNTOU = ActiveSheet.Cells(FILA, 187).Offset(0, I - 1).value
        
        Range("GX9").Offset(I, 0).value = PUNTO
        Range("GY9").Offset(I, 0).value = PUNTOERR
        Range("HA9").Offset(I, 0).value = Format(PUNTOU, "0.0E+")
        Range("GZ9").Offset(I, 0).FormulaLocal = "=SI(O(C" & FILA & "=""IP"", C" & FILA & "=""CP"", C" & FILA & "=""MB"", C" & FILA & "=""VA""),(HA" & U & ")/BUSCARV(I" & FILA & ",EQUIVALENCIAS[#Todo],2,0),HA" & U & ")"
        U = U + 1
        
        Else
        
        filaG = FILA + 1
        '''''''''''''''''''''''''''''''''''''''''
        If FIJO_PATRON = "IBC" Then
        PUNTO = Format(ActiveSheet.Cells(filaG, 106).Offset(0, 2 * I - 17).value, "0" & DGTO)
        
        ElseIf FIJO_PATRON = "PATRON" Then
        PUNTO = Format(ActiveSheet.Cells(filaG, 105).Offset(0, 2 * I - 17).value, "0" & DGTOPAT)
        End If
        
        PUNTO_INDICACION = Format(ActiveSheet.Cells(filaG, 105).Offset(0, 2 * I - 16).value, "0" & DGTO)
        PUNTO_REFERENCIA = Format(ActiveSheet.Cells(filaG, 105).Offset(0, 2 * I - 17).value, "0" & DGTOPAT)
        
        PUNTOERR = PUNTO_INDICACION - PUNTO_REFERENCIA
        
        '''''''''''''''''''''''''''''
        
        
        
        'PUNTO = ActiveSheet.Cells(filaG, 106).Offset(0, 2 * I - 17).value
        'PUNTOERR = ActiveSheet.Cells(filaG, 105).Offset(0, 2 * I - 16).value - ActiveSheet.Cells(filaG, 105).Offset(0, 2 * I - 17).value
        PUNTOU = ActiveSheet.Cells(filaG, 187).Offset(0, I - 9).value
        
        
        Range("GX9").Offset(I, 0).value = PUNTO
        Range("GY9").Offset(I, 0).value = PUNTOERR
        Range("HA9").Offset(I, 0).value = Format(PUNTOU, "0.0E+")
        Range("GZ9").Offset(I, 0).FormulaLocal = "=SI(O(C" & FILA & "=""IP"",C" & filaG & "=""CP"",C" & filaG & "=""MB"",C" & filaG & "=""VA""),(HA" & U & ")/BUSCARV(I" & filaG & ",EQUIVALENCIAS[#Todo],2,0),HA" & U & ")"
        U = U + 1
        
        End If
    
    Next I
    
      Range("GX9").Select
    Tab1 = "=" & ActiveSheet.Name & "!" & Range("GZ10:GZ" & 9 + PUNTOSREF).Address(ReferenceStyle:=xlR1C1)
    
            
             Set graf = ActiveSheet.Shapes.AddChart2(240, xlXYScatterLines)
             graf.Name = "comportamiento"
            Set RNg = Range("HA10:HG19")
            
        ''''''''' grafica
           Set COM = ActiveWorkbook.Sheets(3).ChartObjects("comportamiento").Chart
                 ''''''____________ESCALA
    MAXES = Range("HB5").value
    MINES = Range("HA5").value
             COM.Axes(xlCategory).MaximumScale = MAXES
             COM.Axes(xlCategory).MinimumScale = MINES
             COM.ChartArea.Border.ColorIndex = 2
    'COM.Axes(xlCategory).MajorUnit = 50
    'COM.Axes(xlCategory).MinorUnit = 1
        ''''''''''''''''''''''
            With graf
             .Select
            ActiveChart.SetSourceData Source:=Range("CERTIFICADOS!$GX$10:$GY$" & 9 + PUNTOSREF)
        
         ActiveChart.SeriesCollection(1).ErrorBar _
        Direction:=xlY, Include:=xlErrorBarIncludeBoth, _
    Type:=xlErrorBarTypeCustom, Amount:=Tab1, MinusValues:=Tab1
    
    .Top = RNg(1).Top
    .Left = RNg(1).Left
    .Width = RNg.Width
    .Height = RNg.Height
    .Chart.HasTitle = True
    .Chart.ChartTitle.Text = TITULOGRAF
    
    ActiveChart.SetElement (msoElementLegendNone)
    
      With ActiveChart
      
        .HasAxis(xlValue, xlPrimary) = True
        
        With .Axes(xlValue, xlPrimary)
          .HasTitle = True
          With .AxisTitle
            .Text = TITULOY
    
          End With
        End With
      End With
      
      With ActiveChart.Axes(xlCategory)
     .HasTitle = True
     .AxisTitle.Text = TITULOX
      
      End With
    End With
    Range("GW9").Select
    '''''''''''''''''''
    Application.Wait (Now + TimeValue("00:00:02"))
    
    
    Else
    U = 10
    For I = 1 To PUNTOSREF
        
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        
         If FIJO_PATRON = "IBC" Then
        PUNTO = Format(ActiveSheet.Cells(FILA, 102).Offset(0, 2 * I - 1).value, "0" & DGTO)
        
        ElseIf FIJO_PATRON = "PATRON" Then
        PUNTO = Format(ActiveSheet.Cells(FILA, 101).Offset(0, 2 * I - 1).value, "0" & DGTOPAT)
        End If
        
        PUNTO_INDICACION = Format(ActiveSheet.Cells(FILA, 102).Offset(0, 2 * I - 1).value, "0" & DGTO)
        PUNTO_REFERENCIA = Format(ActiveSheet.Cells(FILA, 101).Offset(0, 2 * I - 1).value, "0" & DGTOPAT)
        
        PUNTOERR = PUNTO_INDICACION - PUNTO_REFERENCIA
        
        '''''''''''''''''''''''''''''''''''''''''''''''''''
        
    
        'If ID_TIPO = "I" Then
        'PUNTO = ActiveSheet.Cells(FILA, 101).Offset(0, 2 * I - 1).value
        'PUNTO = ActiveSheet.Cells(FILA, 101).Offset(0, 2 * I - 1).value
        
        
        'ElseIf ID_TIPO = "P" Then
        'PUNTO = ActiveSheet.Cells(FILA, 102).Offset(0, 2 * I - 1).value
        'End If
        
        'If ID_TIPO = "I" Then
        'PUNTOERR = ActiveSheet.Cells(FILA, 101).Offset(0, 2 * I).value - ActiveSheet.Cells(FILA, 101).Offset(0, 2 * I - 1).value
        
        'ElseIf ID_TIPO = "P" Then
        'PUNTOERR = ActiveSheet.Cells(FILA, 101).Offset(0, 2 * I).value - ActiveSheet.Cells(FILA, 101).Offset(0, 2 * I - 1).value
        'End If
    
    PUNTOU = ActiveSheet.Cells(FILA, 185).Offset(0, I - 1).value
    
    
    ''''''''''''''''''''''PUNTOS GRAF
    
    Range("GX9").Offset(I, 0).value = PUNTO
    Range("GY9").Offset(I, 0).value = PUNTOERR
    Range("HA9").Offset(I, 0).value = Format(PUNTOU, "0.00")
    Range("GZ9").Offset(I, 0).FormulaLocal = "=SI(Y(O(C" & FILA & "=""IP"",C" & FILA & "=""CP"",C" & FILA & "=""MB"",C" & FILA & "=""VA""),I" & FILA & "=""Pa""),HA" & U & ",SI(O(C" & FILA & "=""IP"",C" & FILA & "=""CP"",C" & FILA & "=""MB"",C" & FILA & "=""VA""),(HA" & U & ")/BUSCARV(I" & FILA & ",EQUIVALENCIAS[#Todo],2,0),HA" & U & "))"
    U = U + 1
    Next I
      Range("GX9").Select
    Tab1 = "=" & ActiveSheet.Name & "!" & Range("GZ10:GZ" & 9 + PUNTOSREF).Address(ReferenceStyle:=xlR1C1)
    
            
             Set graf = ActiveSheet.Shapes.AddChart2(240, xlXYScatterLines)
             graf.Name = "comportamiento"
            Set RNg = Range("HA10:HG19")
            
        ''''''''' grafica
           Set COM = ActiveWorkbook.Sheets(3).ChartObjects("comportamiento").Chart
           
                 ''''''____________ESCALA
    MAXES = Range("HB5").value
    MINES = Range("HA5").value
             COM.Axes(xlCategory).MaximumScale = MAXES
             COM.Axes(xlCategory).MinimumScale = MINES
             COM.ChartArea.Border.ColorIndex = 2
    'COM.Axes(xlCategory).MajorUnit = 50
    'COM.Axes(xlCategory).MinorUnit = 1
        ''''''''''''''''''''''
        
            With graf
             .Select
            ActiveChart.SetSourceData Source:=Range("CERTIFICADOS!$GX$10:$GY$" & 9 + PUNTOSREF)
        
         ActiveChart.SeriesCollection(1).ErrorBar _
        Direction:=xlY, Include:=xlErrorBarIncludeBoth, _
    Type:=xlErrorBarTypeCustom, Amount:=Tab1, MinusValues:=Tab1
    
    .Top = RNg(1).Top
    .Left = RNg(1).Left
    .Width = RNg.Width
    .Height = RNg.Height
    .Chart.HasTitle = True
    .Chart.ChartTitle.Text = TITULOGRAF
    
    ActiveChart.SetElement (msoElementLegendNone)
    
      With ActiveChart
      
        .HasAxis(xlValue, xlPrimary) = True
        
        With .Axes(xlValue, xlPrimary)
          .HasTitle = True
          With .AxisTitle
            .Text = TITULOY
    
          End With
        End With
      End With
      
      With ActiveChart.Axes(xlCategory)
     .HasTitle = True
     .AxisTitle.Text = TITULOX
      
      End With
    End With
    Range("GW9").Select
    '''''''''''''''''''
    Application.Wait (Now + TimeValue("00:00:02"))
    
    End If

End Sub

Sub QR()


ID = " " & ActiveSheet.Cells(FILA, 1)
CERTIFICADO = " " & ActiveSheet.Cells(FILA, 99)
FECHA = FECHA_C
'FECHA = " " & Format(ActiveSheet.Cells(1, 4), "dd/mmm/yy")

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''



Dim lookupvalue As Variant, value As Variant, lookupRange As Range
ING = Range("D5").value 'celda con el valor que buscamos

USUARIOS = Sheets(P_2).Range("H3:I16")

V_USUARIO = Application.VLookup(ING, USUARIOS, 2, False)

If IsError(lookupvalue) Then
MsgBox "USUARIO NO AUTORIZADO"
Exit Sub

Else

REALIZA = V_USUARIO
End If



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''



'//////////////////////////////////// QR
On Error Resume Next
Range("GT9").FormulaLocal = "=QRCODE(""ID:" & ID & Chr(10) & "No. Certificado:" & CERTIFICADO & Chr(10) & "Fecha:" & FECHA & Chr(10) & "Realizó:" & REALIZA & """)"

On Error GoTo 0
Sheets(P_3).Select

'Set OBJQROBJ = ActiveWorkbook.Worksheets(P_3).ListObjects("QR_")
'OBJQROBJ.Resize Range("GT9:GU14")
Application.Wait (Now + TimeValue("00:00:02"))
'Range("GT9:GU14").CopyPicture xlScreen, xlPicture

End Sub


Sub REGISTRO()

Dim PUNTO, RUTA_E_U, WORKBOOK_E_U As String
Dim LIBRO, Data As String
Dim ERROR As Double
Dim ID As String


''''''''''''''''PESTAÑAS'''''''''''''''''
P_1 = "DATOS"
P_2 = "MENU"
P_3 = "CERTIFICADOS"
''''''''''''''''''''''''''''''''''''''''

Application.ScreenUpdating = False
Worksheets(P_3).Unprotect "MET2025"

Sheets(P_3).Select
LIBRO = ThisWorkbook.Name
Data = Range("A10:A46").Count - Range("A10:A46").SpecialCells(xlCellTypeBlanks).Count
Path = ThisWorkbook.Path



FECHA_GENERACION = Format(Now(), "DD/MMM/YY")


'FOLIO = Format(Range("D6").Value, "000_00_0000")
FOLIO = Range("D6").value

USUARIO = Range("D5").value

'RUTA_RESPALDO---------------------

RUTA_DATOS = "\\FS-CORPDL-FILE\Grupos\HOJAS DE CALCULO\METROLOGIA\HOJA DE CALCULO\DATOS\"
'RUTA_DATOS = "T:\HOJA DE CALCULO\DATOS\PRUEBAS"


R = 10
I_D = 0

ID = Range("A10").Offset(I_D, 0).value
UNIDADES = Range("I10").Offset(I_D, 0).value


U = 0
E = 0
FECHA = Format(Range("D1"), "DD-MM-YY")
P = 0
PUNTO = 1


Set FSO = CreateObject("SCRIPTING.FILESYSTEMOBJECT")

'' PUNTOS
'' HOJA 1

Set PUNTOS = FSO.CREATETEXTFILE(RUTA_DATOS & "PUNTOS\" & FECHA & "_" & FOLIO & ".TXT")
PUNTOS.WRITE "FECHA,ID,PUNTO,ERROR,INCERTIDUMBRE,E_M_P,UNIDADES,DIVISION_MINIMA" & Chr(10)


'' PUNTOS GENERALES
'' HOJA 2

Set PUNTOS_GENERALES = FSO.CREATETEXTFILE(RUTA_DATOS & "PUNTOS_GENERALES\" & FECHA & "_" & FOLIO & ".TXT")
PUNTOS_GENERALES.WRITE "FECHA,ID,INCERTIDUMBRE,EMP,UNIDADES,DIVISION_MINIMA,PATRON, P1,E1,P2,E2,P3,E3,P4,E4,P5,E5," _
& "P6,E6,P7,E7,P8,E8,P9,E9,P10,E10" & Chr(10)



'' REGISTRO CERTIFICADOS
'' HOJA 3

Set REGISTRO_CERTIFICADOS = FSO.CREATETEXTFILE(RUTA_DATOS & "REGISTRO_CERTIFICADOS\" & FECHA & "_" & FOLIO & ".TXT")
REGISTRO_CERTIFICADOS.WRITE "ID,FECHA,LIBRO,RUTA,PATRON,FECHA GENERACIÓN,USUARIO" & Chr(10)


'Set REGISTRO_CERTIFICADOS = FSO.CREATETEXTFILE(RUTA_DATOS & "PUNTOS_GENERALES\" & ID & "_" & FECHA & "_" & FOLIO & ".TXT")


'For D = 1 To Data
PATRON = Range("GO10").Offset(I_D, 0).value

While PATRON <> ""
D = 1
ID = Range("A10").Offset(I_D, 0).value
UNIDADES = Range("I10").Offset(I_D, 0).value
PATRON = Range("GO10").Offset(I_D, 0).value
RES_INS = Range("K10").Offset(I_D, 0).value

INCERTIDUMBRE = Round(Range("GL10").Offset(I_D, 0).value, 2)
EMP = Range("L10").Offset(I_D, 0).value
E_ = Round(Range("GQ10").Offset(I_D, 0).value, 2)




PUNTOS_GENERALES.WRITE FECHA & "," & ID & "," & INCERTIDUMBRE _
& "," & EMP & "," & UNIDADES & "," & RES_INS & "," & PATRON & ","




REGISTRO_CERTIFICADOS.WRITE ID & "," & FECHA & "," & LIBRO & "," & Path & "," & PATRON & "," & FECHA_GENERACION & "," & USUARIO & Chr(10)


''''''''''MILIVOLTAJE

If PATRON = "N5-11-IA-005 Sim_T" Or PATRON = "N5-11-IA-005 Sim_mV" Then

For I = 1 To 3

PUNTO = Range("M" & R).Offset(0, P).value



INCERTIDUMBRE = Round(Range("GC" & R).Offset(0, U).value, 2)
ERROR = Range("CX" & R).Offset(0, E).value - Range("CY" & R).Offset(0, E).value
P = P + 4
U = U + 1
E = E + 2



PUNTOS.WRITE FECHA & "," & ID & "," & PUNTO & "," & ERROR & "," & INCERTIDUMBRE & "," & EMP & "," & UNIDADES & "," & RES_INS & Chr(10)


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


PUNTOS_GENERALES.WRITE PUNTO & "," & ERROR & ","

Workbooks(LIBRO).Activate
PUNTO = Range("M" & R).Offset(0, P).value


Next I

PUNTOS_GENERALES.WRITE Chr(10)

Else

For U = 0 To 9

PUNTO = Range("M" & R).Offset(0, P).value



INCERTIDUMBRE = Round(Range("GC" & R).Offset(0, U).value, 2)
ERROR = Range("CX" & R).Offset(0, E).value - Range("CY" & R).Offset(0, E).value
P = P + 4
'U = U + 1
E = E + 2



PUNTOS.WRITE FECHA & "," & ID & "," & PUNTO & "," & ERROR & "," & INCERTIDUMBRE & "," & EMP & "," & UNIDADES & "," & RES_INS & Chr(10)
PUNTOS_GENERALES.WRITE PUNTO & "," & ERROR & ","

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Workbooks(LIBRO).Activate
PUNTO = Range("M" & R).Offset(0, P).value


Next U

PUNTOS_GENERALES.WRITE Chr(10)

End If

I_D = I_D + 1
ID = Range("A" & R).Offset(I_D, 0).value

'''''''''''' REINICIAMOS CONTADORES
U = 0
E = 0
FECHA = Range("D1")
P = 0
PUNTO = 1
R = R + 1
PATRON = Range("GO10").Offset(I_D, 0).value

Wend

'Next D





PUNTOS_GENERALES.Close
PUNTOS.Close
REGISTRO_CERTIFICADOS.Close

End Sub

Sub DATOS_PATRON()

'NUMDATOSP_2 = Sheets(P_2).Range("L" & Rows.Count).End(xlUp).Row - 58
'TABLA_INSTRUMENTOS = Sheets(P_2).Range("L3:BA" & NUMDATOSP_2)
'valor = ActiveSheet.Cells(FILA, 197).value


'DES_PATRON = Application.VLookup(valor, TABLA_INSTRUMENTOS, 4, False)
'MARCA
'MODELO
'TIPO
'SERIE
'ALCANCE
'IDENTIFICACION
'RESOLUCION
'TIPO_ERROR
'SIGLAS_ERROR
'NORMA
'INCERTIDUMBRE
'FECHA_CALIBRACION
'CALIBRADO_POR
'INFORME
End Sub
