<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html>
<html lang="es">
<HEAD>
    <meta http-equiv="X-UA-Compatible" content="IE=edge,chrome=1">
    <meta charset="utf-8">
    <link href="../css/normalize.css" rel="stylesheet">
	<!--#include file="sesion.asp"-->
	<!--#include file="arch_utils.asp"-->
	<!--#include file="../lib/comunes/rutinas/funciones.inc" -->
	<!--#include file="../lib/comunes/rutinas/rutinasFecha.inc" -->
	<!--#include file="../lib/asp/comunes/general/rutinasBooleano.inc"-->
	<link href="../css/style.css" rel="stylesheet" type="text/css">
	<link href="../css/style_generales_sistema.css" rel="stylesheet">	
<%

Response.CodePage=65001
Response.charset ="utf-8"


strTipoProceso = request("CB_TIPO_PROCESO")
strAccionRegistro = request("CB_ACCION_REGISTRO")

strGraba = request("strGraba")
strCodCliente=session("ses_codcli")
intCodUsuario = session("session_idusuario")

abrirscg()

strFecha = TraeFechaHoraActual(Conn)

CerrarSCG()
'response.write "strFecha : " & strFecha
'response.write "strAccionRegistro : " & strAccionRegistro
'response.end


 If Trim(strGraba) = "S" Then
	
	If Trim(strTipoProceso) ="6" AND Trim(strAccionRegistro) ="1" Then
		strSql = 			"UPDATE DEUDOR SET ID_FOCO=NULL, ORDEN_FOCO= NULL, FECHA_ESTADO_FOCO=NULL,ID_TIPO_SUB_FOCO=NULL " 
		strSql = strSql & 	"FROM DEUDOR D INNER JOIN CUOTA C ON D.RUT_DEUDOR=C.RUT_DEUDOR AND C.COD_CLIENTE=D.COD_CLIENTE "
		strSql = strSql & 	"INNER JOIN ESTADO_DEUDA ED ON C.ESTADO_DEUDA=ED.CODIGO "
		strSql = strSql & 	"WHERE ED.ACTIVO=1" 
		'Response.WRITE strSql
		
		abrirscg()
		set rsUpdate = Conn.execute(strSql)
		cerrarscg()
		
		strSql2 = 			"DELETE " 
		strSql2 = strSql2 & "FROM DEUDOR_FOCO "
		strSql2 = strSql2 & "WHERE CONVERT(VARCHAR(8),FECHA_CREACION_FOCO,112) =  CONVERT(VARCHAR(8),GETDATE(),112)" 
		'Response.WRITE strSql2
		
		abrirscg()
		set rsUpdate = Conn.execute(strSql2)
		cerrarscg()
		
		intOrdenFoco=0
		
	End If
	
  	intIdDeudor 				= Trim(Request("TX_ID_DEUDOR"))
  	intIdSubFoco 		= Trim(Request("TX_ID_SUB_FOCO"))
	intIdTipoSubFoco 		= Trim(Request("TX_ID_TIPO_SUB_FOCO"))

	vIdDeudor		= split(intIdDeudor,CHR(13))
	vIdSubFoco		= split(intIdSubFoco,CHR(13))
	vIdTipoSubFoco		= split(intIdTipoSubFoco,CHR(13))

	intTamvId 			=ubound(vIdDeudor)

	'Response.WRITE "<br>intIdDeudor=" & intIdDeudor &"<br>intIdSubFoco=" & intIdSubFoco

  		For indice = 0 to intTamvId
		
			intOrdenFoco = intOrdenFoco + 1

  			if intIdSubFoco <> "" and intIdTipoSubFoco <>"" then
  				intIdSubFocoSplit= Replace(vIdSubFoco(indice),CHR(10),"")
				intIdTipoSubFocoSplit = Replace(vIdTipoSubFoco(indice),CHR(10),"")
			end if
			
			
			'---------Generar o Actualiza Focos Call-----------'
			
			
			If Trim(strTipoProceso) = 6 or Trim(strTipoProceso) = 14 AND Trim(strAccionRegistro) ="1" Then
				
				strSql = 		  " UPDATE DEUDOR"
				strSql = strSql & " SET ID_FOCO= " & trim(intIdSubFocoSplit) & ",ID_TIPO_SUB_FOCO= " & trim(intIdTipoSubFocoSplit) & ",ORDEN_FOCO = '" & intOrdenFoco & "',FECHA_ESTADO_FOCO = '" & strFecha & "'"
				strSql = strSql & " FROM DEUDOR D"
				strSql = strSql & " WHERE ID_DEUDOR = " & Replace(trim(vIdDeudor(indice)),CHR(10),"")& " AND " & trim(intIdSubFocoSplit) & " IN (1,2,3)"			

			abrirscg()
			set rsUpdate = Conn.execute(strSql)
			cerrarscg()
			
			End if
			
			If Trim(strTipoProceso) = 15 AND Trim(strAccionRegistro) ="1" Then
				
				strSql = 		  " UPDATE DEUDOR SET ID_FOCO=NULL, ORDEN_FOCO= NULL, FECHA_ESTADO_FOCO=NULL,ID_TIPO_SUB_FOCO=NULL " 
				strSql = strSql & " FROM DEUDOR D"
				strSql = strSql & " WHERE ID_DEUDOR = " & Replace(trim(vIdDeudor(indice)),CHR(10),"")
				
				'Response.WRITE strSql
				'Response.End

			abrirscg()
			set rsUpdate = Conn.execute(strSql)
			cerrarscg()	

				strSql2 = 			"DELETE " 
				strSql2 = strSql2 & "FROM DEUDOR_FOCO "
				strSql2 = strSql2 & "WHERE CONVERT(VARCHAR(8),FECHA_CREACION_FOCO,112) =  CONVERT(VARCHAR(8),GETDATE(),112) AND ID_DEUDOR = " & Replace(trim(vIdDeudor(indice)),CHR(10),"")& " AND ID_SUB_FOCO IN (1,2,3)" 			

				'Response.WRITE strSql2
				'Response.End
				
			abrirscg()
			set rsUpdate = Conn.execute(strSql2)
			cerrarscg()	
			
			End if
			
			
			'---------Generar o Actualiza Focos Canal Alternativo-----------'
			
			
			If (Trim(strTipoProceso) ="6" OR Trim(strTipoProceso) ="14") AND ( Trim(strAccionRegistro) ="2" OR Trim(strAccionRegistro) ="3" OR Trim(strAccionRegistro) ="4" OR Trim(strAccionRegistro) ="5" OR Trim(strAccionRegistro) ="6") Then
			
				strSql = "exec [proc_Insert_Deudor_Foco_Medio_TMP] " &intCodUsuario& "," &intOrdenFoco& "," & Trim(strAccionRegistro) & "," & trim(intIdSubFocoSplit)& "," & trim(intIdTipoSubFocoSplit)& "," & Replace(trim(vIdDeudor(indice)),CHR(10),"")
				
				'Response.Write strSql
				'Response.End
				
				abrirscg()
				set RsInsert = Conn.execute(strSql)
				cerrarscg()
				
			End If					
			
		Next
		
		abrirscg()

		If (Trim(strTipoProceso) ="6" OR Trim(strTipoProceso) ="14") AND Trim(strAccionRegistro) ="1" Then
		
			strSql = "exec proc_Insert_Deudor_Foco_Call"
			set RsInsert = Conn.execute(strSql)
			
		End If					
		
		cerrarscg()
		
		If (Trim(strTipoProceso) ="6" OR Trim(strTipoProceso) ="14") AND ( Trim(strAccionRegistro) ="2" OR Trim(strAccionRegistro) ="3" OR Trim(strAccionRegistro) ="4" OR Trim(strAccionRegistro) ="5" OR Trim(strAccionRegistro) ="6") Then
		
			strSql = "exec [proc_Insert_Deudor_Foco_Medio] "
			
			'Response.Write strSql
			'Response.End
			
			abrirscg()
			set RsInsert = Conn.execute(strSql)
			cerrarscg()
			
		End If	
		
		
		
	%>
	
	<script>
		alert('Proceso realizado correctamente');
	</script>
	<%


  End if

abrirscg()

%>
<title>UTILITARIO</title>

<style type="text/css">
<!--
.Estilo37 {color: #FFFFFF}
-->
</style>
</head>
<body>
<div class="titulo_informe">UTILITARIO POR IDENTIFICADOR</div>
<br>
<table width="90%" align="CENTER" border="0">
   <tr>
    <td valign="top" background="../imagenes/fondo_coventa.jpg">
	<BR>
	<FORM name="datos" method="post">
	
	<table width="40%" border="0" align="center" class="intercalado" style="width:40%;">
		<thead>
			<tr>
				<td>
					<select name="CB_TIPO_PROCESO" style="width:200px;" onChange="CargaRegistros(CB_TIPO_PROCESO.value);">
					<option value="0" <%If strTipoProceso = "0" then response.write "SELECCIONAR"%>>SELECCIONAR</option>
					<option value="6" <%If strTipoProceso = "6" then response.write "SELECTED"%>>GENERAR FOCOS</option>
					<option value="14" <%If strTipoProceso = "14" then response.write "SELECTED"%>>ACTUALIZAR FOCOS</option>
					<option value="15" <%If strTipoProceso = "15" then response.write "SELECTED"%>>ELIMINAR FOCOS</option>
				</td>
				<td>
					<select name="CB_ACCION_REGISTRO" id="CB_ACCION_REGISTRO" style="width:200px;">
				</td>
				<td align="center">
					<input TYPE="BUTTON" class="fondo_boton_100" value="Procesar" name="B1" onClick="envia(CB_TIPO_PROCESO.value,CB_ACCION_REGISTRO.value);return false;">
				</td>
			</tr>
			<tr>
				<td>ID DEUDOR</td>
				<td>ID SUB FOCO / ID MEDIO</td>
				<td>ID TIPO SUB FOCO / ID SCRIPT</td>
			</tr>
			
			<tr>
				<td class=hdr_i>
					<TEXTAREA NAME="TX_ID_DEUDOR" ROWS=30 COLS=15><%=intIdDeudor%></TEXTAREA>
				</td>
				<td colspan=1 class=hdr_i>
					<TEXTAREA NAME="TX_ID_SUB_FOCO" ROWS=30 COLS=15><%=intIdSubFoco%></TEXTAREA>
				</td>
				<td colspan=1 class=hdr_i>
					<TEXTAREA NAME="TX_ID_TIPO_SUB_FOCO" ROWS=30 COLS=15><%=intIdTipoSubFoco%></TEXTAREA>
				</td>
			</tr>
	</table>

	</form>
	
</table>
</body>
</html>

<script language="JavaScript1.2">
function envia(intTipoProceso, intAccionRegistro)	{
		datos.B1.disabled = true;
		var comboBox = document.getElementById('CB_ACCION_REGISTRO');
	
		//alert(intTipoProceso);
		
		if (intTipoProceso == '0')
		{
			alert("Favor seleccione el Tipo de Proceso");
		}
		else if (comboBox.value == '0')
		{
			alert("Favor seleccione la Acción o el Tipo de Registro a modificar");
		}
		else if (intTipoProceso == '6' && intAccionRegistro =='1'){
			if (confirm("*************** ¿ Está seguro de crear nuevos focos Call, esto reemplazará todos los focos Call antiguos asociados a todos los clientes por estos nuevos focos? ***************"))
			{
				document.forms[0].action='UtilitarioFocos.asp?strGraba=S';
				document.forms[0].submit();
			}
		}
		else if (intTipoProceso == '6' && intAccionRegistro =='2'){
			if (confirm("¿ Está seguro de crear nuevos focos email, esto reemplazará todos los focos email antiguos asociados a todos los clientes por estos nuevos focos? "))
			{
				document.forms[0].action='UtilitarioFocos.asp?strGraba=S';
				document.forms[0].submit();
			}
		}
		else if (intTipoProceso == '6' && intAccionRegistro =='3'){
			if (confirm("¿ Está seguro de crear nuevos focos IVR, esto reemplazará todos los focos IVR antiguos asociados a todos los clientes por estos nuevos focos? "))
			{
				document.forms[0].action='UtilitarioFocos.asp?strGraba=S';
				document.forms[0].submit();
			}
		}
		else if (intTipoProceso == '6' && intAccionRegistro =='4'){
			if (confirm("¿ Está seguro de crear nuevos focos SMS, esto reemplazará todos los focos SMS antiguos asociados a todos los clientes por estos nuevos focos? "))
			{
				document.forms[0].action='UtilitarioFocos.asp?strGraba=S';
				document.forms[0].submit();
			}
		}
		else if (intTipoProceso == '6' && intAccionRegistro =='5'){
			if (confirm("¿ Está seguro de crear nuevos focos Carta, esto reemplazará todos los focos Carta antiguos asociados a todos los clientes por estos nuevos focos? "))
			{
				document.forms[0].action='UtilitarioFocos.asp?strGraba=S';
				document.forms[0].submit();
			}
		}
		else if (intTipoProceso == '6' && intAccionRegistro =='6'){
			if (confirm("¿ Está seguro de crear nuevos focos Terreno, esto reemplazará todos los focos Terreno antiguos asociados a todos los clientes por estos nuevos focos? "))
			{
				document.forms[0].action='UtilitarioFocos.asp?strGraba=S';
				document.forms[0].submit();
			}
		}
		else if (intTipoProceso == '14' && intAccionRegistro =='1'){
			if (confirm("¿ Está seguro de Actualizar los focos Call, esto creará nuevos focos Call o modificará los que existen actualmente?"))
			{
				document.forms[0].action='UtilitarioFocos.asp?strGraba=S';
				document.forms[0].submit();
			}
		}
		else if (intTipoProceso == '14' && intAccionRegistro =='2'){
			if (confirm("¿ Está seguro de Actualizar los focos email, esto creará nuevos focos email o modificará los que existen actualmente?"))
			{
				document.forms[0].action='UtilitarioFocos.asp?strGraba=S';
				document.forms[0].submit();
			}
		}
		else if (intTipoProceso == '14' && intAccionRegistro =='3'){
			if (confirm("¿ Está seguro de Actualizar los focos IVR, esto creará nuevos focos IVR o modificará los que existen actualmente?"))
			{
				document.forms[0].action='UtilitarioFocos.asp?strGraba=S';
				document.forms[0].submit();
			}
		}
		else if (intTipoProceso == '14' && intAccionRegistro =='4'){
			if (confirm("¿ Está seguro de Actualizar los focos SMS, esto creará nuevos focos SMS o modificará los que existen actualmente?"))
			{
				document.forms[0].action='UtilitarioFocos.asp?strGraba=S';
				document.forms[0].submit();
			}
		}
		else if (intTipoProceso == '14' && intAccionRegistro =='5'){
			if (confirm("¿ Está seguro de Actualizar los focos Carta, esto creará nuevos focos Carta o modificará los que existen actualmente?"))
			{
				document.forms[0].action='UtilitarioFocos.asp?strGraba=S';
				document.forms[0].submit();
			}
		}
		else if (intTipoProceso == '14' && intAccionRegistro =='6'){
			if (confirm("¿ Está seguro de Actualizar los focos Terreno, esto creará nuevos focos Terreno o modificará los que existen actualmente?"))
			{
				document.forms[0].action='UtilitarioFocos.asp?strGraba=S';
				document.forms[0].submit();
			}
		}
		else if (intTipoProceso == '15' && intAccionRegistro =='1'){
			if (confirm("¿ Está seguro de Eliminar los focos Call asociados a los Id que existen actualmente?"))
			{
				document.forms[0].action='UtilitarioFocos.asp?strGraba=S';
				document.forms[0].submit();
			}
		}
		datos.B1.disabled = false;
}

function CargaRegistros(subCat,registro)
{
//
	var comboBox = document.getElementById('CB_ACCION_REGISTRO');
	comboBox.options.length = 0;

		if (subCat=='6') {
			var newOption = new Option('SELECCIONAR', '0');
			comboBox.options[comboBox.options.length] = newOption;
			var newOption = new Option('CALL', '1');
			comboBox.options[comboBox.options.length] = newOption;
			var newOption = new Option('EMAIL', '2');
			comboBox.options[comboBox.options.length] = newOption;
			var newOption = new Option('IVR', '3');
			comboBox.options[comboBox.options.length] = newOption;
			var newOption = new Option('SMS', '4');
			comboBox.options[comboBox.options.length] = newOption;
			var newOption = new Option('CARTA', '5');
			comboBox.options[comboBox.options.length] = newOption;
			var newOption = new Option('TERRENO', '6');
			comboBox.options[comboBox.options.length] = newOption;
		}
		else if (subCat=='14') {
			var newOption = new Option('SELECCIONAR', '0');
			comboBox.options[comboBox.options.length] = newOption;
			var newOption = new Option('CALL', '1');
			comboBox.options[comboBox.options.length] = newOption;
			var newOption = new Option('EMAIL', '2');
			comboBox.options[comboBox.options.length] = newOption;
			var newOption = new Option('IVR', '3');
			comboBox.options[comboBox.options.length] = newOption;
			var newOption = new Option('SMS', '4');
			comboBox.options[comboBox.options.length] = newOption;
			var newOption = new Option('CARTA', '5');
			comboBox.options[comboBox.options.length] = newOption;
			var newOption = new Option('TERRENO', '6');
			comboBox.options[comboBox.options.length] = newOption;
		}
		else if (subCat=='15') {
			var newOption = new Option('SELECCIONAR', '0');
			comboBox.options[comboBox.options.length] = newOption;
			var newOption = new Option('CALL', '1');
			comboBox.options[comboBox.options.length] = newOption;
		}
		else {
			var newOption = new Option('SELECCIONAR', '0');
			comboBox.options[comboBox.options.length] = newOption;
		}
		ActualizaComboRegistro(comboBox,registro)
}						
						
function InicializaInforme()
{
		var comboBox = document.getElementById('CB_ACCION_REGISTRO');
		comboBox.options.length = 0;
		var newOption = new Option('SELECCIONAR','0');
		comboBox.options[comboBox.options.length] = newOption;
}
function ActualizaComboRegistro(comboRegistro,Registro)
{
		for (var i=0; i< comboRegistro.options.length; i ++)
		{
		if (comboRegistro.options[i].value == Registro)
			comboRegistro.options[i].selected = true;
			}
}

function RefrescaDatos(){
	document.forms[0].submit();
}
<%If strAccionRegistro = "" then%>
InicializaInforme()
<%End If%>

<%If strTipoProceso <> "" then%>
CargaRegistros('<%=strTipoProceso%>','<%=strAccionRegistro%>');
<%End If%>

</script>
