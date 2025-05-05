Attribute VB_Name = "Módulo3"
Public P_1, P_2, P_3, FOLIO As String
'''''''''Versión 5.1.3 _ 20/NOV/24_ (M:\HOJA DE CALCULO\NO TOCAR\CÓDIGO\DL) (REDONDEO DE RESULTADOS)
Sub BUSCAR_I_D()

Dim CNN As Object
Dim RS As Object
Dim STRCONN As String
Dim STRSQL As String
Dim FOLIO, RAZON_SOCIAL, RUTA, SERVIDOR, CONSULTA_A, RUTA_CONSULTA_ING As String
Dim RUTA_METROLIGA, RUTA_CERTIFICADOS_ING As String
Dim ARCHIVO As String
Dim FILA, FILA_LM As Long


''''''''''''''''PESTAÑAS'''''''''''''''''

P_1 = "DATOS"
P_2 = "MENU"
P_3 = "CERTIFICADOS"
Worksheets(P_3).Protect "MET2025"
Worksheets(P_3).Unprotect "MET2025"

'''''''''''''''''''''''''''''''''''''





Application.ScreenUpdating = False


COD = Sheets(P_3).Range("D6").value
NUM = InStr(COD, "-") - 1
NUM_COD = Len(COD)

If NUM = -1 Then
RAZ_SOC = "PRO"

ElseIf NUM <> -1 Then
RAZ_SOC = Left(COD, NUM)

    If RAZ_SOC = "DLM" Then
   CARPETA_RAZ_SOC = "01_DL_MEDICA"
   SISTEMA_RAZ_SOC = "SISTEMA_DL_MEDICA"
   SIS_RAZ_SOC = "DL_MEDICA"
   
    ElseIf RAZ_SOC = "GIP" Then
   CARPETA_RAZ_SOC = "02_GIP"
   SISTEMA_RAZ_SOC = "SISTEMA_GIP"
   SIS_RAZ_SOC = "GIP"

    ElseIf RAZ_SOC = "DLP" Then
   CARPETA_RAZ_SOC = "03_DLP"
   SISTEMA_RAZ_SOC = "SISTEMA_DLP"
   SIS_RAZ_SOC = "DLP"

    ElseIf RAZ_SOC = "DEN" Then
   CARPETA_RAZ_SOC = "04_DENTILAB"
   SISTEMA_RAZ_SOC = "SISTEMA_DENTILAB"
   SIS_RAZ_SOC = "DEN"
    
    End If
End If

'FOLIO = Right(COD, 4)
'FOLIO = Format(Sheets("MENU").Range("D6").Value, "0000")


'FOLIO

If RAZ_SOC = "PRO" Then
FOLIO = Format(Sheets(P_3).Range("D6").value, "0000")
ElseIf RAZ_SOC <> "PROFI" Then
FOLIO = Sheets(P_3).Range("D6").value
End If

EQUIPO = Environ("UserName")

If EQUIPO = "bmtro06" Then
'NELLY/ EDUARDO
    RUTA_METROLOGIA = "T:\"
    RUTA_CERTIFICADOS_ING = "Y:\"
    
ElseIf EQUIPO = "bmet03" Then
'JUAN/ CHUCHIN
    RUTA_METROLOGIA = "T:\"
    RUTA_CERTIFICADOS_ING = "Z:\"
    
ElseIf EQUIPO = "bmtro05" Then
'WILLY/ LEO
    RUTA_METROLOGIA = "T:\"
    RUTA_CERTIFICADOS_ING = "Z:\"
    
ElseIf EQUIPO = "bmtro02" Then
'EDSON/ ALAN
    RUTA_METROLOGIA = "T:\"
    RUTA_CERTIFICADOS_ING = "Z:\"

ElseIf EQUIPO = "bmtro07" Then
'MIGUE/ LEO
    RUTA_METROLOGIA = "T:\"
    RUTA_CERTIFICADOS_ING = "\\DESDENMET01N05\Users\amet01\Documents\electro04\CERTIFICADOS ING´S\"

ElseIf EQUIPO = "bmet04" Then
'JJ
    RUTA_METROLOGIA = "T:\"
    RUTA_CERTIFICADOS_ING = "D:\"

ElseIf EQUIPO = "cmeto01" Then
'LAU
    'RUTA_METROLOGIA = "T:\"
    RUTA_CERTIFICADOS_ING = "D:\"

ElseIf EQUIPO = "bmtro01" Then
'JOSE
    RUTA_METROLOGIA = "G:\"
    RUTA_CERTIFICADOS_ING = "Z:\"
    
ElseIf EQUIPO = "amet01" Then
'PACHECO
    RUTA_METROLOGIA = "T:\"
    'RUTA_CERTIFICADOS_ING = "Y:\"

Else

End If


''''LIMPIAR REGISTROS


Sheets(P_3).Range("A10:A46").value = ""

        On Error Resume Next
        Range("tbl_LISTA_MAESTRA").Delete
        On Error GoTo 0
        
        
        
        
'''''''''''''''''''''''''''RUTA


RUTA_CERTIFICADOS_ING = "\\DESDENMET01N05\Users\amet01\Documents\electro04\CERTIFICADOS ING´S\"
'RUTA_CERTIFICADOS_ING = RUTA_CERTIFICADOS_ING


RUTA = ThisWorkbook.Path

'FOLIO

If RAZ_SOC = "PRO" Then


CONSULTA_A = " ID, DESCRIPCION, MARCA,MODELO, NO_SERIE, UNIDADES,INTERVALO_DE_USO,D_MINIMA, " & _
             "I_USO_MIN,I_USO_MAX,I_MEDICION_MIN,I_MEDICION_MAX,EMP,SERVICIO,EQUIPO, " & _
             "ESTADO,REF,ULT_ACT,FOLIO,F_SIG,RAZ_SOCIAL,AJUSTE,NO_CERTIFICADO,MAGNITUD, " & _
             "UBICACION, FECHA_SERV, UBICACION_CER, PUNTO_C_1, PUNTO_C_2, PUNTO_C_3, " & _
             "PUNTO_C_4, PUNTO_C_5"
             
RUTA_CONSULTA_ING = "BASE DE DATOS\05_PROFI"

RUTA_ARCHIVO = RUTA_CERTIFICADOS_ING & RUTA_CONSULTA_ING & "\SISTEMA_PROFI_ING'S.accdb"
             

ElseIf RAZ_SOC <> "PROFI" Then
CONSULTA_A = " ID, DESCRIP, MARCA,MODELO, NO_SERIE, UNIDADES,INTERVALO_DE_USO,D_MINIMA, " & _
             "I_USO_MIN,I_USO_MAX,I_MEDICION_MIN,I_MEDICION_MAX,EMP,SERVICIO,EQUIPO, " & _
             "ESTADO,REF,ULT_ACT,FOLIO,F_SIG,RAZ_SOCIAL,AJUSTE,NO_CERTIFICADO,MAGNITUD, " & _
             "UBICACION, FECHA_SERV, UBICACION_CER, PUNTO_C_1, PUNTO_C_2, PUNTO_C_3, " & _
             "PUNTO_C_4, PUNTO_C_5"
             
RUTA_CONSULTA_ING = "BASE DE DATOS\1_DL_MEDICA\"
RUTA_ARCHIVO = RUTA_CERTIFICADOS_ING & RUTA_CONSULTA_ING & "\SISTEMA_DL_MEDICA_ING'S.accdb"

End If

'RUTA_CONSULTA_ING = "RESPALDO\RESPALDO_BD\METROLOGÍA\01_DL_MEDICA\SISTEMA_DL_MEDICA_ING'S.accdb"

''''''''''''RUTA PRUEBAS

'



''''''''''''''PRUEBAS
'RUTA_CONSULTA_ING = "RESPALDO\RESPALDO_BD\METROLOGÍA\" & CARPETA_RAZ_SOC & "\" & SISTEMA_RAZ_SOC & "_ING'S.accdb"


'

'''''conexión
 ' Establecer conexión con base de datos
    Set CNN = CreateObject("ADODB.Connection")
    
    STRCONN = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & RUTA_ARCHIVO
    

On Error GoTo CAMBIAR_RUTA
    CNN.Open STRCONN
On Error GoTo 0
    ' Crear la consulta SQL para extraer los datos
    
'If RAZON_SOCIAL = "DL MÉDICA" Then
    STRSQL = "SELECT" & CONSULTA_A & " FROM QRY_H_C_" & RAZ_SOC & " WHERE FOLIO = '" & FOLIO & "' ORDER BY MAGNITUD, ID;"

'ElseIf RAZON_SOCIAL = "GIP" Then
 '   STRSQL = "SELECT" & CONSULTA_A & "FROM QRY_H_C_GIP WHERE FOLIO = '" & FOLIO & "' ORDER BY MAGNITUD, ID;"

'ElseIf RAZON_SOCIAL = "DLP" Then
    'STRSQL = "SELECT" & CONSULTA_A & "FROM QRY_H_C_DLP WHERE FOLIO = '" & FOLIO & "' ORDER BY MAGNITUD, ID;"

'ElseIf RAZON_SOCIAL = "DENTILAB" Then
 '   STRSQL = "SELECT" & CONSULTA_A & "FROM QRY_H_C_DEN WHERE FOLIO = '" & FOLIO & "' ORDER BY MAGNITUD, ID;"

'End If


 ''''''''''''''''''' Ejecutar la consulta
 
    Set RS = CreateObject("ADODB.Recordset")
    RS.Open STRSQL, CNN
    
    ' Escribir los datos en la Hoja 2
    FILA = 10 ' primera fila para escribir datos
    FILA_LM = 2

    
    Do While Not RS.EOF
        Sheets(P_3).Cells(FILA, 1).value = RS("ID")






'USUARIOS = Sheets(P_2).Range("H3:I16")

'V_USUARIO = Application.VLookup(ING, USUARIOS, 2, False)
        
        
Sheets(P_1).Cells(FILA_LM, 1).value = RS("ID")

I_D_L = Sheets(P_1).Cells(FILA_LM, 1).value


Msg = "¿LOS VALORES EN EL PATRÓN SON LOS PUNTOS FIJOS PARA EL INSTRUMENTO: " & I_D_L & "?"    ' MENSAJE
Style = vbYesNo   ' BOTONES
Title = "VALORES FIJOS"    ' TITULO

        
Response = MsgBox(Msg, Style, Title)

If Response = vbYes Then
    P_FIJO = "P"
Else
    P_FIJO = "I"
End If




If RAZ_SOC = "PRO" Then
Sheets(P_1).Cells(FILA_LM, 2).value = RS("DESCRIPCION")

ElseIf RAZ_SOC <> "PROFI" Then
Sheets(P_1).Cells(FILA_LM, 2).value = RS("DESCRIP")
End If

Sheets(P_1).Cells(FILA_LM, 3).value = RS("MARCA")
Sheets(P_1).Cells(FILA_LM, 4).value = RS("MODELO")
Sheets(P_1).Cells(FILA_LM, 5).value = RS("NO_SERIE")
Sheets(P_1).Cells(FILA_LM, 6).value = RS("UNIDADES")
Sheets(P_1).Cells(FILA_LM, 7).value = RS("INTERVALO_DE_USO")
Sheets(P_1).Cells(FILA_LM, 8).value = RS("D_MINIMA")
Sheets(P_1).Cells(FILA_LM, 9).value = RS("I_USO_MIN")
Sheets(P_1).Cells(FILA_LM, 10).value = RS("I_USO_MAX")
Sheets(P_1).Cells(FILA_LM, 11).value = RS("I_MEDICION_MIN")
Sheets(P_1).Cells(FILA_LM, 12).value = RS("I_MEDICION_MAX")
Sheets(P_1).Cells(FILA_LM, 13).value = RS("EMP")
Sheets(P_1).Cells(FILA_LM, 14).value = RS("SERVICIO")
Sheets(P_1).Cells(FILA_LM, 15).value = RS("EQUIPO")
Sheets(P_1).Cells(FILA_LM, 16).value = RS("ESTADO")
Sheets(P_1).Cells(FILA_LM, 17).value = RS("REF")
Sheets(P_1).Cells(FILA_LM, 18).value = RS("ULT_ACT")
Sheets(P_1).Cells(FILA_LM, 19).value = RS("FOLIO")
Sheets(P_1).Cells(FILA_LM, 20).value = RS("F_SIG")
Sheets(P_1).Cells(FILA_LM, 21).value = RS("RAZ_SOCIAL")
Sheets(P_1).Cells(FILA_LM, 22).value = RS("AJUSTE")
Sheets(P_1).Cells(FILA_LM, 23).value = RS("NO_CERTIFICADO")
Sheets(P_1).Cells(FILA_LM, 24).value = RS("MAGNITUD")
Sheets(P_1).Cells(FILA_LM, 25).value = RS("UBICACION")
Sheets(P_1).Cells(FILA_LM, 26).value = RS("FECHA_SERV")
Sheets(P_1).Cells(FILA_LM, 27).value = RS("UBICACION_CER")
Sheets(P_1).Cells(FILA_LM, 28).value = RS("PUNTO_C_1")
Sheets(P_1).Cells(FILA_LM, 29).value = RS("PUNTO_C_2")
Sheets(P_1).Cells(FILA_LM, 30).value = RS("PUNTO_C_3")
Sheets(P_1).Cells(FILA_LM, 31).value = RS("PUNTO_C_4")
Sheets(P_1).Cells(FILA_LM, 32).value = RS("PUNTO_C_5")
Sheets(P_1).Cells(FILA_LM, 33).value = P_FIJO
        '''''''''''''''''''PUNTOS DE CALIBRACION

P_C_1 = Sheets(P_1).Range("AB2").Offset(FILA_LM - 2, 0).value
If P_C_1 = VACIO Then
P_C_1 = 0
End If

P_C_2 = Sheets(P_1).Range("AC2").Offset(FILA_LM - 2, 0).value
If P_C_2 = VACIO Then
P_C_2 = 0
End If

P_C_3 = Sheets(P_1).Range("AD2").Offset(FILA_LM - 2, 0).value
If P_C_3 = VACIO Then
P_C_3 = 0
End If

P_C_4 = Sheets(P_1).Range("AE2").Offset(FILA_LM - 2, 0).value
If P_C_4 = VACIO Then
P_C_4 = 0
End If

P_C_5 = Sheets(P_1).Range("AF2").Offset(FILA_LM - 2, 0).value
If P_C_5 = VACIO Then
P_C_5 = 0
End If



''''''''''''MAGNITUD
MAGNITUD_ID = Sheets(P_1).Range("X2").Offset(FILA_LM - 2, 0).value

''''''''PUNTO 1

If P_C_1 <> 0 Or P_C_2 <> 0 Or P_C_3 <> 0 Or P_C_4 <> 0 Or P_C_5 <> 0 Then


For PUNTO_C = 0 To 3

If MAGNITUD_ID = "MASA" Then

Sheets(P_3).Range("BI10").Offset(FILA - 10, PUNTO_C).FormulaR1C1 = P_C_1
Sheets(P_3).Range("BM10").Offset(FILA - 10, PUNTO_C).FormulaR1C1 = P_C_2
Sheets(P_3).Range("BQ10").Offset(FILA - 10, PUNTO_C).FormulaR1C1 = P_C_3
Sheets(P_3).Range("BU10").Offset(FILA - 10, PUNTO_C).FormulaR1C1 = P_C_4
Sheets(P_3).Range("BY10").Offset(FILA - 10, PUNTO_C).FormulaR1C1 = P_C_5

Else

    If P_FIJO = "I" Then
    Sheets(P_3).Range("BA10").Offset(FILA - 10, PUNTO_C).FormulaR1C1 = P_C_1
    Sheets(P_3).Range("BE10").Offset(FILA - 10, PUNTO_C).FormulaR1C1 = P_C_2
    Sheets(P_3).Range("BI10").Offset(FILA - 10, PUNTO_C).FormulaR1C1 = P_C_3
    Sheets(P_3).Range("BM10").Offset(FILA - 10, PUNTO_C).FormulaR1C1 = P_C_4
    Sheets(P_3).Range("BQ10").Offset(FILA - 10, PUNTO_C).FormulaR1C1 = P_C_5
    
    ElseIf P_FIJO = "P" Then
    Sheets(P_3).Range("M10").Offset(FILA - 10, PUNTO_C).FormulaR1C1 = P_C_1
    Sheets(P_3).Range("Q10").Offset(FILA - 10, PUNTO_C).FormulaR1C1 = P_C_2
    Sheets(P_3).Range("U10").Offset(FILA - 10, PUNTO_C).FormulaR1C1 = P_C_3
    Sheets(P_3).Range("Y10").Offset(FILA - 10, PUNTO_C).FormulaR1C1 = P_C_4
    Sheets(P_3).Range("AC10").Offset(FILA - 10, PUNTO_C).FormulaR1C1 = P_C_5
    End If
    
End If

Next PUNTO_C

Else
' SIN PUNTOS DETERMINADOS
MsgBox ("INSTRUMENTO " & I_D_L & " NO CUENTA CON PUNTOS DETERMINADOS DE CALIBRACIÓN")
End If

''''''''''''''''''''''''''''''UBICACION INSTRUMENTO

tbl_L_M = Sheets(P_1).Range("A1:AF" & FILA_LM)

UB_INS_N = Application.VLookup(I_D_L, tbl_L_M, 25, False)
'UB_INS_N_REPLACE = Replace(UB_INS_N, "N", "NAVE ", 1, 1)
'UB_INS_N_REPLACE_FORMAT = Format(UB_INS_N_REPLACE, ">")
'UB_INS = UB_INS_N_REPLACE_FORMAT

'''''''''''''''''''''SE REALIZA LA BUSQUEDA EN MÓDULO_1
'Range("GS10").Offset(FILA - 10, 0).FormulaLocal = "=SI(C" & FILA & "<>"""",BUSCARX(A" & FILA & ",tbl_LISTA_MAESTRA[ID]," & _
                                                "tbl_LISTA_MAESTRA[UBICACIÓN],"""",0,1),BUSCARX(A" & FILA & ",tbl_LISTA_MAESTRA[ID]," & _
                                                "tbl_LISTA_MAESTRA[UBICACIÓN],"""",0,-1))"


        FILA = FILA + 1
        FILA_LM = FILA_LM + 1
        RS.MoveNext ' mover al siguiente registro
    Loop
    
    
    ' Cerrar la conexión y liberar memoria
    RS.Close
    CNN.Close
    Set RS = Nothing
    Set CNN = Nothing
    
    
 '''''''''''''''UBICACION'''''''''''''''
UBICACION_CER_N = Sheets(P_1).Range("AA2").value
UBICACION_CER_N_REPLACE = Replace(UBICACION_CER_N, "N", "NAVE ", 1, 1)
UBICACION_CER_N_REPLACE_FORMAT = Format(UBICACION_CER_N_REPLACE, ">")
UBICACION_CER = UBICACION_CER_N_REPLACE_FORMAT

Sheets(P_1).Range("AA2").Copy
Sheets(P_3).Range("D3").FormulaR1C1 = UBICACION_CER
  
If RAZ_SOC = "PRO" Then
UBICACION = "PROFILATEX, S.A. DE C.V."

ElseIf RAZ_SOC <> "PROFI" Then
UBICACION = "Febrero de 1917 s/n Zona Industrial de Chalco, Chalco Edo. de México CP 56600, Tel. 597 56060"

End If

Sheets(P_3).Range("D4").FormulaR1C1 = UBICACION






Sheets(P_3).Range("D3").FormulaR1C1 = UBICACION_CER


Application.ScreenUpdating = True

Worksheets(P_3).Protect "MET2025"
Exit Sub




CAMBIAR_RUTA:
Call CORREGIR_RUTA
Exit Sub

End Sub

Sub CORREGIR_RUTA()


Dim CNN As Object
Dim RS As Object
Dim STRCONN As String
Dim STRSQL As String
Dim FOLIO, RAZON_SOCIAL, RUTA, SERVIDOR, CONSULTA_A, RUTA_CONSULTA_ING As String
Dim RUTA_METROLIGA, RUTA_CERTIFICADOS_ING As String
Dim ARCHIVO As String
Dim FILA, FILA_LM As Long

''''''''''''''''PESTAÑAS'''''''''''''''''

P_1 = "DATOS"
P_2 = "MENU"
P_3 = "CERTIFICADOS"

'''''''''''''''''''''''''''''''''''''





Application.ScreenUpdating = False


COD = Sheets(P_3).Range("D6").value
NUM = InStr(COD, "-") - 1
NUM_COD = Len(COD)

RAZ_SOC = Left(COD, NUM)

'FOLIO = Right(COD, 4)
FOLIO = Range("D6").value

frm_RUTA_DB.Show
RUTA_CORREGIDA = frm_RUTA_DB.TB_RUTA.value

RUTA = ThisWorkbook.Path
CONSULTA_A = " ID,DESCRIP,MARCA,MODELO,NO_SERIE,UNIDADES,INTERVALO_DE_USO,D_MINIMA," & _
            "I_USO_MIN,I_USO_MAX,I_MEDICION_MIN,I_MEDICION_MAX,EMP,SERVICIO,EQUIPO," & _
            "ESTADO,REF,ULT_ACT,FOLIO,F_SIG,RAZ_SOCIAL,AJUSTE,NO_CERTIFICADO,MAGNITUD," & _
            "UBICACION, AREA, FECHA_SERV "
RUTA_CONSULTA_ING = RUTA_CORREGIDA & "\SISTEMA_DL_MEDICA_ING'S.accdb"



'''''conexión
 ' Establecer conexión con base de datos
    Set CNN = CreateObject("ADODB.Connection")
    
    STRCONN = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & RUTA_CERTIFICADOS_ING & RUTA_CONSULTA_ING

    CNN.Open STRCONN

    ' Crear la consulta SQL para extraer los datos
    
'If RAZON_SOCIAL = "DL MÉDICA" Then
    STRSQL = "SELECT" & CONSULTA_A & " FROM QRY_H_C_" & RAZ_SOC & " WHERE FOLIO = '" & FOLIO & "' ORDER BY MAGNITUD, ID;"

'ElseIf RAZON_SOCIAL = "GIP" Then
 '   STRSQL = "SELECT" & CONSULTA_A & "FROM QRY_H_C_GIP WHERE FOLIO = '" & FOLIO & "' ORDER BY MAGNITUD, ID;"

'ElseIf RAZON_SOCIAL = "DLP" Then
    'STRSQL = "SELECT" & CONSULTA_A & "FROM QRY_H_C_DLP WHERE FOLIO = '" & FOLIO & "' ORDER BY MAGNITUD, ID;"

'ElseIf RAZON_SOCIAL = "DENTILAB" Then
 '   STRSQL = "SELECT" & CONSULTA_A & "FROM QRY_H_C_DEN WHERE FOLIO = '" & FOLIO & "' ORDER BY MAGNITUD, ID;"

'End If


 ''''''''''''''''''' Ejecutar la consulta
 
    Set RS = CreateObject("ADODB.Recordset")
    RS.Open STRSQL, CNN
    
    ' Escribir los datos en la Hoja 2
    FILA = 10 ' primera fila para escribir datos
    FILA_LM = 2

    
    Do While Not RS.EOF
        Sheets(P_3).Cells(FILA, 1).value = RS("ID")
Sheets(P_1).Cells(FILA_LM, 1).value = RS("ID")
Sheets(P_1).Cells(FILA_LM, 2).value = RS("DESCRIP")
Sheets(P_1).Cells(FILA_LM, 3).value = RS("MARCA")
Sheets(P_1).Cells(FILA_LM, 4).value = RS("MODELO")
Sheets(P_1).Cells(FILA_LM, 5).value = RS("NO_SERIE")
Sheets(P_1).Cells(FILA_LM, 6).value = RS("UNIDADES")
Sheets(P_1).Cells(FILA_LM, 7).value = RS("INTERVALO_DE_USO")
Sheets(P_1).Cells(FILA_LM, 8).value = RS("D_MINIMA")
Sheets(P_1).Cells(FILA_LM, 9).value = RS("I_USO_MIN")
Sheets(P_1).Cells(FILA_LM, 10).value = RS("I_USO_MAX")
Sheets(P_1).Cells(FILA_LM, 11).value = RS("I_MEDICION_MIN")
Sheets(P_1).Cells(FILA_LM, 12).value = RS("I_MEDICION_MAX")
Sheets(P_1).Cells(FILA_LM, 13).value = RS("EMP")
Sheets(P_1).Cells(FILA_LM, 14).value = RS("SERVICIO")
Sheets(P_1).Cells(FILA_LM, 15).value = RS("EQUIPO")
Sheets(P_1).Cells(FILA_LM, 16).value = RS("ESTADO")
Sheets(P_1).Cells(FILA_LM, 17).value = RS("REF")
Sheets(P_1).Cells(FILA_LM, 18).value = RS("ULT_ACT")
Sheets(P_1).Cells(FILA_LM, 19).value = RS("FOLIO")
Sheets(P_1).Cells(FILA_LM, 20).value = RS("F_SIG")
Sheets(P_1).Cells(FILA_LM, 21).value = RS("RAZ_SOCIAL")
Sheets(P_1).Cells(FILA_LM, 22).value = RS("AJUSTE")
Sheets(P_1).Cells(FILA_LM, 23).value = RS("NO_CERTIFICADO")
Sheets(P_1).Cells(FILA_LM, 24).value = RS("MAGNITUD")
Sheets(P_1).Cells(FILA_LM, 25).value = RS("UBICACION")
Sheets(P_1).Cells(FILA_LM, 26).value = RS("AREA")
Sheets(P_1).Cells(FILA_LM, 27).value = RS("FECHA_SERV")
 
        FILA = FILA + 1
        FILA_LM = FILA_LM + 1
        RS.MoveNext ' mover al siguiente registro
    Loop
    
    ' Cerrar la conexión y liberar memoria
    RS.Close
    CNN.Close
    Set RS = Nothing
    Set CNN = Nothing
    
Application.ScreenUpdating = True
Exit Sub


End Sub
