<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001" LCID = 1034%>

<% ' Capa 1 ' %>
<!--#include file="../lib/asp/comunes/odbc/ADOVBS.INC" -->
<!--#include file="arch_utils.asp"-->
<!--#include file="../lib/asp/comunes/odbc/insertUpdate.inc"-->
<!--#include file="../lib/asp/comunes/odbc/ObtenerRecordset.inc"-->

<% ' Capa 2 ' %>
<!--#include file="../lib/asp/comunes/insert/Remesa.inc"-->
<!--#include file="../lib/asp/comunes/recordset/Remesa.inc"-->
<!--#include file="../lib/asp/comunes/general/funciones.inc"-->
<%
	Response.CodePage=65001
	Response.charset ="utf-8"
	
AbrirSCG()

If Request("strFormMode") = "Nuevo" Then
	If Trim(request("COD_REMESA")) <> "" AND Trim(request("COD_CLIENTE")) <> "" Then
		recordset_Remesa Conn, srsRegistro, request("COD_REMESA"), request("COD_CLIENTE")
		If Not srsRegistro.EOF Then
			Response.Write "<P>Ya existe un registro con el código de asignacion : " &  request("COD_REMESA")& " y codigo de cliente : " & request("COD_CLIENTE")
			Response.Write "<P>Debe asignarle otro código si desea crear un registro nuevo"
			Response.Write "<FORM><INPUT VALUE=Volver TYPE=BUTTON onClick='javascript:history.back()'></FORM>"
			Response.End
		End If
	End If
End If
'Response.write "hola"
'Response.write "CB_BANCO" & request("CB_BANCO")
'Response.End

Set dicRemesa = CreateObject("Scripting.Dictionary")
dicRemesa.Add "COD_REMESA", request("COD_REMESA")
dicRemesa.Add "COD_CLIENTE", request("COD_CLIENTE")
dicRemesa.Add "NOMBRE", ValNulo(request("NOMBRE"),"C")
dicRemesa.Add "DESCRIPCION", ValNulo(request("DESCRIPCION"),"C")
dicRemesa.Add "FECHA_LLEGADA", ValNulo(request("FECHA_LLEGADA"),"C")
dicRemesa.Add "FECHA_CARGA", ValNulo(request("FECHA_CARGA"),"C")
dicRemesa.Add "ACTIVO", ValNulo(request("ACTIVO"),"N")

insert_Remesa Conn, dicRemesa

CerrarSCG()
Response.Redirect "man_Remesa.asp"
%>
