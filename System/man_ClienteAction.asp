<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001" LCID = 1034%>

<% ' Capa 1 ' %>
<!--#include file="../lib/asp/comunes/odbc/ADOVBS.INC" -->
<!--#include file="arch_utils.asp"-->
<!--#include file="../lib/asp/comunes/odbc/insertUpdate.inc"-->
<!--#include file="../lib/asp/comunes/odbc/ObtenerRecordset.inc"-->

<% ' Capa 2 ' %>
<!--#include file="../lib/asp/comunes/insert/Cliente.inc"-->
<!--#include file="../lib/asp/comunes/recordset/Cliente.inc"-->
<!--#include file="../lib/asp/comunes/general/funciones.inc"-->
<%
Response.CodePage=65001
Response.charset ="utf-8"

AbrirSCG() 

If Request("strFormMode") = "Nuevo" Then
	If Trim(request("COD_CLIENTE")) <> "" Then
		recordset_Cliente Conn, srsRegistro, request("COD_CLIENTE")
		If Not srsRegistro.EOF Then
			Response.Write "<P>Ya existe un registro con el código " &  request("COD_CLIENTE")
			Response.Write "<P>Debe asignarle otro código si desea crear un registro nuevo"
			Response.Write "<FORM><INPUT VALUE=Volver TYPE=BUTTON onClick='javascript:history.back()'></FORM>"
			Response.End
		End If
	End If
End If
'Response.write "hola"
'Response.write "NOMBRE_FANTASIA" & request("NOMBRE_FANTASIA")

Set dicCliente = CreateObject("Scripting.Dictionary")
dicCliente.Add "COD_CLIENTE", request("COD_CLIENTE")
dicCliente.Add "RUT", ValNulo(request("RUT"),"C")
dicCliente.Add "DESCRIPCION", ValNulo(request("DESCRIPCION"),"C")
dicCliente.Add "RAZON_SOCIAL", ValNulo(request("RAZON_SOCIAL"),"C")
dicCliente.Add "NOMBRE_FANTASIA", ValNulo(request("NOMBRE_FANTASIA"),"C")
dicCliente.Add "DIRECCION", ValNulo(request("DIRECCION"),"C")
dicCliente.Add "COMUNA", ValNulo(request("COMUNA"),"C")
dicCliente.Add "FONO_1", ValNulo(request("FONO_1"),"C")
dicCliente.Add "FONO_2", ValNulo(request("FONO_2"),"C")
dicCliente.Add "NOM_CONTACTO", ValNulo(request("NOM_CONTACTO"),"C")
dicCliente.Add "EMAIL_CONTACTO", ValNulo(request("EMAIL_CONTACTO"),"C")
dicCliente.Add "TASA_MAX_CONV", ValNulo(request("TASA_MAX_CONV"),"N")
dicCliente.Add "IC_PORC_CAPITAL", ValNulo(request("IC_PORC_CAPITAL"),"N")
dicCliente.Add "HON_PORC_CAPITAL", ValNulo(request("HON_PORC_CAPITAL"),"N")
dicCliente.Add "PIE_PORC_CAPITAL", ValNulo(request("PIE_PORC_CAPITAL"),"N")
dicCliente.Add "TIPO_CLIENTE", ValNulo(request("TIPO_CLIENTE"),"C")
dicCliente.Add "TIPO_INTERES", ValNulo(request("TIPO_INTERES"),"C")
dicCliente.Add "INTERES_MORA", ValNulo(request("INTERES_MORA"),"N")
dicCliente.Add "EXPIRACION_CONVENIO", ValNulo(request("EXPIRACION_CONVENIO"),"N")
dicCliente.Add "EXPIRACION_ANULACION", ValNulo(request("EXPIRACION_ANULACION"),"N")
dicCliente.Add "GASTOS_OPERACIONALES", ValNulo(request("GASTOS_OPERACIONALES"),"N")
dicCliente.Add "GASTOS_ADMINISTRATIVOS", ValNulo(request("GASTOS_ADMINISTRATIVOS"),"N")
dicCliente.Add "GASTOS_OPERACIONALES_CD", ValNulo(request("GASTOS_OPERACIONALES_CD"),"N")
dicCliente.Add "GASTOS_ADMINISTRATIVOS_CD", ValNulo(request("GASTOS_ADMINISTRATIVOS_CD"),"N")
dicCliente.Add "USA_CUSTODIO", ValNulo(request("USA_CUSTODIO"),"C")
dicCliente.Add "COLOR_CUSTODIO", ValNulo(request("COLOR_CUSTODIO"),"C")
dicCliente.Add "ADIC_1", ValNulo(request("ADIC_1"),"C")
dicCliente.Add "ADIC_2", ValNulo(request("ADIC_2"),"C")
dicCliente.Add "ADIC_3", ValNulo(request("ADIC_3"),"C")
dicCliente.Add "ADIC_4", ValNulo(request("ADIC_4"),"C")
dicCliente.Add "ADIC_5", ValNulo(request("ADIC_5"),"C")
dicCliente.Add "ADIC_91", ValNulo(request("ADIC_91"),"C")
dicCliente.Add "ADIC_92", ValNulo(request("ADIC_92"),"C")
dicCliente.Add "ADIC_93", ValNulo(request("ADIC_93"),"C")
dicCliente.Add "ADIC_94", ValNulo(request("ADIC_94"),"C")
dicCliente.Add "ADIC_95", ValNulo(request("ADIC_95"),"C")
dicCliente.Add "COD_MONEDA", ValNulo(request("COD_MONEDA"),"C")
dicCliente.Add "NOMBRE_CON_PAGARE", ValNulo(request("NOMBRE_CON_PAGARE"),"C")
dicCliente.Add "ACTIVO", ValNulo(request("ACTIVO"),"N")
dicCliente.Add "COD_TIPODOCUMENTO_HON", ValNulo(request("COD_TIPODOCUMENTO_HON"),"C")
dicCliente.Add "MESES_TD_HON", ValNulo(request("MESES_TD_HON"),"N")
dicCliente.Add "ADIC1_DEUDOR", ValNulo(request("ADIC1_DEUDOR"),"C")
dicCliente.Add "ADIC2_DEUDOR", ValNulo(request("ADIC2_DEUDOR"),"C")
dicCliente.Add "ADIC3_DEUDOR", ValNulo(request("ADIC3_DEUDOR"),"C")
dicCliente.Add "NOMBRE_CONV_PAGARE", ValNulo(request("NOMBRE_CONV_PAGARE"),"C")
dicCliente.Add "RETIRO_SABADO", ValNulo(request("RETIRO_SABADO"),"N")
dicCliente.Add "USA_HONORARIOS", ValNulo(request("USA_HONORARIOS"),"N")
dicCliente.Add "FORMULA_HONORARIOS", ValNulo(request("FORMULA_HONORARIOS"),"C")
dicCliente.Add "USA_INTERESES", ValNulo(request("USA_INTERESES"),"N")
dicCliente.Add "FORMULA_INTERESES", ValNulo(request("FORMULA_INTERESES"),"C")
dicCliente.Add "FORMULA_HONORARIOS_FACT", ValNulo(request("FORMULA_HONORARIOS_FACT"),"C")

insert_Cliente Conn, dicCliente

'Response.End

CerrarSCG()
Response.Redirect "man_Cliente.asp"
%>
