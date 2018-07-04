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

abrirscg()

strFecha = TraeFechaHoraActual(Conn)

CerrarSCG()
'response.write "strFecha : " & strFecha
'response.write "strAccionRegistro : " & strAccionRegistro
'response.end

abrirscg()

		strSql = "SELECT PERFIL_ADM = ISNULL(PERFIL_ADM,0),PERFIL_BACK = ISNULL(PERFIL_BACK,0),PERFIL_SUP = ISNULL(PERFIL_SUP,0),PERFIL_COB = ISNULL(PERFIL_COB,0),PERFIL_EMP = ISNULL(PERFIL_EMP,0),PERFIL_CAJA = ISNULL(PERFIL_CAJA,0)"
		strSql = strSql & " FROM USUARIO U"
		strSql = strSql & " WHERE ID_USUARIO = " & session("session_idusuario")
		
		set RsPer=Conn.execute(strSql)
		If not RsPer.eof then
			intPerfilAdm = RsPer("PERFIL_ADM")
			intPerfilSup = RsPer("PERFIL_SUP")
			intPerfilBack = RsPer("PERFIL_BACK")
			intPerfilCliente= RsPer("PERFIL_EMP")
			
		End if
		RsPer.close
		
		set RsPer=nothing
		if err then
			response.write "ERROR : " & err.description
			response.end()
		end if
		
		'response.write "intPerfilSup : " & intPerfilSup
		'response.end()
			
cerrarscg()


 If Trim(strGraba) = "S" Then
	
	'Setea en NULL el campop ID_FOCO (call) asociado a deudores activos'
	If Trim(strTipoProceso) ="6" AND Trim(strAccionRegistro) ="1" Then
		strSql = 			"UPDATE DEUDOR SET ID_FOCO=NULL, ORDEN_FOCO= NULL, FECHA_ESTADO_FOCO=NULL " 
		strSql = strSql & 	"FROM DEUDOR D INNER JOIN CUOTA C ON D.RUT_DEUDOR=C.RUT_DEUDOR AND C.COD_CLIENTE=D.COD_CLIENTE "
		strSql = strSql & 	"INNER JOIN ESTADO_DEUDA ED ON C.ESTADO_DEUDA=ED.CODIGO "
		strSql = strSql & 	"WHERE ED.ACTIVO=1" 
		'Response.WRITE strSql
		
		abrirscg()
		set rsUpdate = Conn.execute(strSql)
		cerrarscg()
		
		intOrdenFoco=0
		
	End If
	
  	intId 				= Trim(Request("TX_IDENTIFICADOR"))
  	strValorNuevo 		= Trim(Request("TX_NUEVO_DATO"))

	vIdentificador 		= split(intId,CHR(13))
	vValorNuevo 		= split(strValorNuevo,CHR(13))

	intTamvId 			=ubound(vIdentificador)
	intTamvValorNuevo	=ubound(vValorNuevo)

	'Response.WRITE "<br>intId=" & intId &"<br>strValorNuevo=" & strValorNuevo

  		For indice = 0 to intTamvId
		
			intOrdenFoco = intOrdenFoco + 1

  			if strValorNuevo <> "" then
  				strValorNuevo = Replace(vValorNuevo(indice),CHR(10),"")
			end if
			
		'----------Modifica Estados de Deuda-----------'
		
			If Trim(strTipoProceso) = 1 Then

				If Trim(strAccionRegistro) = "REBAJA_PAGOS_EN_CLIENTE" Then
					strSql = "UPDATE CUOTA SET SALDO = 0, FECHA_ESTADO = getdate(),FECHA_PROCESO_ESTADO_DEUDA = GETDATE(), ESTADO_DEUDA = 3, OBSERVACION = 'PAGO EN CLIENTE TOTAL POR " & session("session_idusuario") & "' WHERE ID_CUOTA = " & Replace(trim(vIdentificador(indice)),CHR(10),"") & " AND COD_CLIENTE = '" & trim(strCodCliente) & "' AND ESTADO_DEUDA <> (3)"

				ElseIf Trim(strAccionRegistro) = "REBAJA_PAGOS_EN_EMPRESA" Then
					strSql = "UPDATE CUOTA SET SALDO = 0, FECHA_ESTADO = getdate(),FECHA_PROCESO_ESTADO_DEUDA = GETDATE(), ESTADO_DEUDA = 4, OBSERVACION = 'PAGO EN EMPRESA TOTAL POR " & session("session_idusuario") & "' WHERE ID_CUOTA = " & Replace(trim(vIdentificador(indice)),CHR(10),"") & " AND COD_CLIENTE = '" & trim(strCodCliente) & "' AND ESTADO_DEUDA <> (4)"

				ElseIf Trim(strAccionRegistro) = "REBAJA_PAGOS_CONVENIO" Then
					strSql = "UPDATE CUOTA SET SALDO = 0, FECHA_ESTADO = getdate(),FECHA_PROCESO_ESTADO_DEUDA = GETDATE(), ESTADO_DEUDA = 10, OBSERVACION = 'CONVENIO TOTAL POR " & session("session_idusuario") & "' WHERE ID_CUOTA = " & Replace(trim(vIdentificador(indice)),CHR(10),"") & " AND COD_CLIENTE = '" & trim(strCodCliente) & "' AND ESTADO_DEUDA <> (10)"
					
				ElseIf Trim(strAccionRegistro) = "PENDIENTE" Then
					strSql = "UPDATE CUOTA SET FECHA_ESTADO = getdate(),FECHA_PROCESO_ESTADO_DEUDA = GETDATE(), ESTADO_DEUDA = 12, OBSERVACION = 'PENDIENTE POR " & session("session_idusuario") & "' WHERE ID_CUOTA = " & Replace(trim(vIdentificador(indice)),CHR(10),"") & " AND COD_CLIENTE = '" & trim(strCodCliente) & "' AND ESTADO_DEUDA <> (12)"

				ElseIf Trim(strAccionRegistro) = "ABONO" Then
					strSql = "UPDATE CUOTA SET SALDO = SALDO - " & trim(strValorNuevo) & " , FECHA_ESTADO = getdate(),FECHA_PROCESO_ESTADO_DEUDA = GETDATE(), ESTADO_DEUDA = 7, OBSERVACION = 'ABONO POR " & session("session_idusuario") & "' WHERE ID_CUOTA = " & Replace(trim(vIdentificador(indice)),CHR(10),"") & " AND COD_CLIENTE = '" & trim(strCodCliente) & "'"

				ElseIf Trim(strAccionRegistro) = "RETIRO" Then
					strSql = "UPDATE CUOTA SET ESTADO_DEUDA = 2 , FECHA_ESTADO = getdate(),FECHA_PROCESO_ESTADO_DEUDA = GETDATE(), SALDO = 0, OBSERVACION = 'RETIRADO POR " & session("session_idusuario") & "' WHERE ID_CUOTA = " & Replace(trim(vIdentificador(indice)),CHR(10),"") & " AND COD_CLIENTE = '" & trim(strCodCliente) & "' AND ESTADO_DEUDA <> (2)"

				ElseIf Trim(strAccionRegistro) = "RETIRO_RES" Then
					strSql = "UPDATE CUOTA SET ESTADO_DEUDA = 5 , FECHA_ESTADO = getdate(),FECHA_PROCESO_ESTADO_DEUDA = GETDATE(), SALDO = 0, OBSERVACION = 'RETIRADO POR RESOL. POR " & session("session_idusuario") & "' WHERE ID_CUOTA = " & Replace(trim(vIdentificador(indice)),CHR(10),"") & " AND COD_CLIENTE = '" & trim(strCodCliente) & "' AND ESTADO_DEUDA <> (5)"

				ElseIf Trim(strAccionRegistro) = "NO ASIGNABLE" Then
					strSql = "UPDATE CUOTA SET ESTADO_DEUDA = 13 , FECHA_ESTADO = getdate(),FECHA_PROCESO_ESTADO_DEUDA = GETDATE(), SALDO = 0, OBSERVACION = 'NO ASIGNABLE POR " & session("session_idusuario") & "' WHERE ID_CUOTA = " & Replace(trim(vIdentificador(indice)),CHR(10),"") & " AND COD_CLIENTE = '" & trim(strCodCliente) & "' AND ESTADO_DEUDA <> (13)"

				ElseIf Trim(strAccionRegistro) = "FIN COBRANZA" Then
					strSql = "UPDATE CUOTA SET ESTADO_DEUDA = 14 , FECHA_ESTADO = getdate(),FECHA_PROCESO_ESTADO_DEUDA = GETDATE(), SALDO = 0, OBSERVACION = 'FIN COBRANZA POR " & session("session_idusuario") & "' WHERE ID_CUOTA = " & Replace(trim(vIdentificador(indice)),CHR(10),"") & " AND COD_CLIENTE = '" & trim(strCodCliente) & "' AND ESTADO_DEUDA <> (14)"

				ElseIf Trim(strAccionRegistro) = "ACTIVAR" Then
					strSql = "UPDATE CUOTA"
					strSql = strSql & " SET ESTADO_DEUDA = 1 , FECHA_ESTADO = getdate(),FECHA_PROCESO_ESTADO_DEUDA = GETDATE(), SALDO = VALOR_CUOTA, OBSERVACION = 'VUELTO A ACTIVAR POR " & session("session_idusuario") & "', FECHA_AGEND_ULT_GES = NULL, HORA_AGEND_ULT_GES = NULL, CUOTA.USUARIO_ASIG = DEUDOR.USUARIO_ASIG,CUOTA.FECHA_ASIGNACION = getdate() "
					strSql = strSql & " FROM CUOTA INNER JOIN DEUDOR ON CUOTA.RUT_DEUDOR = DEUDOR.RUT_DEUDOR AND CUOTA.COD_CLIENTE = DEUDOR.COD_CLIENTE"
					strSql = strSql & " WHERE ID_CUOTA = " & Replace(trim(vIdentificador(indice)),CHR(10),"") & " AND CUOTA.COD_CLIENTE = '" & trim(strCodCliente) & "' AND CUOTA.ESTADO_DEUDA <> (1)"

				ElseIf Trim(strAccionRegistro) = "ERROR CARGA" Then
					strSql = "UPDATE CUOTA SET ESTADO_DEUDA = 15 , FECHA_ESTADO = getdate(),FECHA_PROCESO_ESTADO_DEUDA = GETDATE(), SALDO = 0, OBSERVACION = 'ERROR CARGA POR " & session("session_idusuario") & "' WHERE ID_CUOTA = " & Replace(trim(vIdentificador(indice)),CHR(10),"") & " AND COD_CLIENTE = '" & trim(strCodCliente) & "' AND ESTADO_DEUDA <> (15)"
					
				End if
				
		'----------Modifica Deudores------------'	
			ElseIf Trim(strTipoProceso) = 13 Then 
				
				If Trim(strAccionRegistro)= "MODIFICAR_ETAPA_COBRANZA" then
				
				strSql = "UPDATE DEUDOR SET ETAPA_COBRANZA = "& trim(strValorNuevo) & ", FECHA_ESTADO_ETAPA = GETDATE() WHERE ID_DEUDOR = "&Replace(trim(vIdentificador(indice)),CHR(10),"")
				
				End if		
		'----------Modifica Documentos------------'

			ElseIf Trim(strTipoProceso) = 2 Then
			
				If Trim(strAccionRegistro)= "CUSTODIO" then

					If Trim(strValorNuevo) <> "" Then
						strSql = "UPDATE CUOTA SET CUSTODIO = '" & trim(strValorNuevo) & "',FECHA_ESTADO_CUSTODIO = GETDATE(),USUARIO_ESTADO_CUSTODIO = '" & session("session_idusuario") & "' WHERE ID_CUOTA = '" & Replace(trim(vIdentificador(indice)),CHR(10),"") & "'"
					Else
						strSql = "UPDATE CUOTA SET CUSTODIO = NULL,FECHA_ESTADO_CUSTODIO = GETDATE(),USUARIO_ESTADO_CUSTODIO = '" & session("session_idusuario") & "' WHERE ID_CUOTA = '" & Replace(trim(vIdentificador(indice)),CHR(10),"") & "'"
					End If						
														
				ElseIf Trim(strAccionRegistro)= "VALOR_CUOTA" then
				
				strSql = "UPDATE CUOTA SET VALOR_CUOTA = '"& trim(strValorNuevo) & "', SALDO = '"& trim(strValorNuevo) & "' WHERE ID_CUOTA = "&Replace(trim(vIdentificador(indice)),CHR(10),"")

				ElseIf Trim(strAccionRegistro)= "CAMPANA_CLIENTE" then
				
				strSql = "UPDATE CUOTA SET ID_CAMPANA_CLIENTE = "& trim(strValorNuevo) & ", FECHA_CAMPANA_CLIENTE = '"& trim(strFecha) & "' WHERE ID_CUOTA = "&Replace(trim(vIdentificador(indice)),CHR(10),"")

				ElseIf Trim(strAccionRegistro)= "CODIGO_CLIENTE" then
				
				strSql = "UPDATE CUOTA SET CODIGO_CLIENTE = "& trim(strValorNuevo) &" WHERE ID_CUOTA = "&Replace(trim(vIdentificador(indice)),CHR(10),"")

				ElseIf Trim(strAccionRegistro)= "TRAMO_VENCIMIENTO" then
				
				strSql = "UPDATE CUOTA SET ID_SEGMENTO_VENC = "& trim(strValorNuevo) &" WHERE ID_CUOTA = "&Replace(trim(vIdentificador(indice)),CHR(10),"")

				ElseIf Trim(strAccionRegistro)= "TRAMO_MONTO" then
				
				strSql = "UPDATE CUOTA SET ID_SEGMENTO_MONTO = "& trim(strValorNuevo) &" WHERE ID_CUOTA = "&Replace(trim(vIdentificador(indice)),CHR(10),"")

				ElseIf Trim(strAccionRegistro)= "TRAMO_ASIG" then
				
				strSql = "UPDATE CUOTA SET ID_SEGMENTO_ASIG = "& trim(strValorNuevo) &" WHERE ID_CUOTA = "&Replace(trim(vIdentificador(indice)),CHR(10),"")
				
				ElseIf Trim(strAccionRegistro)= "ASIGNACION" then
				
				strSql = "UPDATE CUOTA SET USUARIO_ASIG = "& trim(strValorNuevo) &", FECHA_ASIGNACION = GETDATE(), TIPO_ASIGNACION = 2 WHERE ID_CUOTA = "&Replace(trim(vIdentificador(indice)),CHR(10),"")
				
				Else
				
				strSql = "UPDATE CUOTA SET " & trim(strAccionRegistro) & " = '"& trim(strValorNuevo) & "' WHERE ID_CUOTA = "&Replace(trim(vIdentificador(indice)),CHR(10),"")
				
				End if
			
		'----------Modifica Gestiones------------'
		
			ElseIf Trim(strTipoProceso) = 3 Then

				strSql = "UPDATE GESTIONES "
				strSql = strSql & " SET " & trim(strAccionRegistro) & " = '" & trim(strValorNuevo) & "'"
				strSql = strSql & " FROM GESTIONES G"
				strSql = strSql & " WHERE G.ID_GESTION = " & Replace(trim(vIdentificador(indice)),CHR(10),"")					
					
		'---------Audita Contactabilidad-----------'
			
			ElseIf Trim(strTipoProceso) = 4 Then
			
				If Trim(strAccionRegistro) = "AUDITAR_TELEFONOS" Then
					strSql = "UPDATE DEUDOR_TELEFONO SET ESTADO = " & trim(strValorNuevo) & ", FECHA_REVISION=getdate() WHERE ID_TELEFONO = " & Replace(trim(vIdentificador(indice)),CHR(10),"")

				ElseIf Trim(strAccionRegistro) = "AUDITAR_EMAIL" Then
					strSql = "UPDATE DEUDOR_EMAIL SET ESTADO = " & trim(strValorNuevo) & ", FECHA_REVISION=getdate() WHERE ID_EMAIL = " & Replace(trim(vIdentificador(indice)),CHR(10),"")

				ElseIf Trim(strAccionRegistro) = "AUDITAR_DIRECCIONES" Then
					strSql = "UPDATE DEUDOR_DIRECCION SET ESTADO = " & trim(strValorNuevo) & ", FECHA_REVISION=getdate() WHERE ID_DIRECCION = " & Replace(trim(vIdentificador(indice)),CHR(10),"")

				ElseIf Trim(strAccionRegistro) = "ESTADO_CARGA_TELEFONOS" Then
					strSql = "UPDATE DEUDOR_TELEFONO SET ESTADO_CARGA = " & trim(strValorNuevo) & " WHERE ID_TELEFONO = " & Replace(trim(vIdentificador(indice)),CHR(10),"")

					
				End If
		'---------Eliminar Registros-----------'

			ElseIf Trim(strTipoProceso) = 5 Then
		
				If Trim(strAccionRegistro) = "ELIMINAR_DOCUMENTOS" Then

					strSql = "DELETE FROM GESTIONES_CUOTA WHERE ID_CUOTA = '" & Replace(trim(vIdentificador(indice)),CHR(10),"") & "'"

					abrirscg()
					set rsUpdate = Conn.execute(strSql)
					cerrarscg()

					strSql = "DELETE FROM PRIORIZACIONES_CUOTA WHERE ID_CUOTA = '" & Replace(trim(vIdentificador(indice)),CHR(10),"") & "'"

					abrirscg()
					set rsUpdate = Conn.execute(strSql)
					cerrarscg()

					strSql = "DELETE FROM CUOTA WHERE ID_CUOTA = '" & Replace(trim(vIdentificador(indice)),CHR(10),"") & "'"				

				ElseIf Trim(strAccionRegistro) = "ELIMINAR_GESTIONES" Then

					strSql = "DELETE FROM GESTIONES_CUOTA WHERE ID_GESTION = '" & Replace(trim(vIdentificador(indice)),CHR(10),"") & "'"
					
					abrirscg()
					set rsUpdate = Conn.execute(strSql)
					cerrarscg()

					strSql = "DELETE FROM GESTIONES WHERE ID_GESTION = '" & Replace(trim(vIdentificador(indice)),CHR(10),"") & "'"

				ElseIf Trim(strAccionRegistro) = "ELIMINAR_TELEFONOS" Then
					strSql = "DELETE FROM DEUDOR_TELEFONO WHERE ID_TELEFONO = '" & Replace(trim(vIdentificador(indice)),CHR(10),"") & "'"

				ElseIf Trim(strAccionRegistro) = "ELIMINAR_EMAIL" Then
					strSql = "DELETE FROM DEUDOR_EMAIL WHERE ID_EMAIL = '" & Replace(trim(vIdentificador(indice)),CHR(10),"") & "'"
					
				ElseIf Trim(strAccionRegistro) = "ELIMINAR_DIRECCIONES" Then
					strSql = "DELETE FROM DEUDOR_DIRECCION WHERE ID_DIRECCION = '" & Replace(trim(vIdentificador(indice)),CHR(10),"") & "'"
					
				End if		
				
			End if
			
			'Response.WRITE strSql
			'Response.End

			abrirscg()
			set rsUpdate = Conn.execute(strSql)
			
			If Trim(strAccionRegistro) = "ACTIVAR" Then
				strSql = "EXEC [proc_Asigna_Fec_Agend] 1,'" & trim(strCodCliente) & "'," & Replace(trim(vIdentificador(indice)),CHR(10),"")
				Conn.Execute strSql,64
			End If	
		
		
			cerrarscg()
			
		Next

		abrirscg()
		
		If Trim(strTipoProceso) = 2 AND Trim(strAccionRegistro)= "CUSTODIO" Then
		
			strSql3 = "Exec Proc_Cambia_Custodio_Deudor '" & trim(strCodCliente) & "'," & session("session_idusuario") 
			set rsUpdate = Conn.execute(strSql3)
		End If	

		If (Trim(strTipoProceso) ="6" OR Trim(strTipoProceso) ="14") AND Trim(strAccionRegistro) ="1" Then
		
			strSql = "exec proc_Insert_Deudor_Foco_Call"
			set RsInsert = Conn.execute(strSql)
			
		End If					
		
		cerrarscg()
			
	%>
	
	<script>
		alert('Proceso realizado correctamente');
	</script>
	<%


  End if

abrirscg()
If Trim(intCodUsuario) = "" Then intCodUsuario = session("session_idusuario")

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
	
	<table width="65%" border="0" align="center" class="intercalado" style="width:65%;">
		<thead>
			<tr>
				<td>TIPO PROCESO</td>
				<td colspan=3>ACCIÓN / REGISTRO</td>	
			</tr>
			<tr>
				<td>
					<select name="CB_TIPO_PROCESO" style="width:200px;" onChange="CargaRegistros(CB_TIPO_PROCESO.value);">
					<option value="0" <%If strTipoProceso = "0" then response.write "SELECCIONAR"%>>SELECCIONAR</option>
					<option value="1" <%If strTipoProceso = "1" then response.write "SELECTED"%>>ACTUALIZAR ESTADO DEUDA</option>
					<option value="13" <%If strTipoProceso = "13" then response.write "SELECTED"%>>MODIFICAR DEUDORES</option>
					<option value="2" <%If strTipoProceso = "2" then response.write "SELECTED"%>>MODIFICAR DOCUMENTOS</option>
					<option value="3" <%If strTipoProceso = "3" then response.write "SELECTED"%>>MODIFICAR GESTIONES</option>
					<option value="4" <%If strTipoProceso = "4" then response.write "SELECTED"%>>AUDITAR CONTACTABILIDAD</option>
					<option value="5" <%If strTipoProceso = "5" then response.write "SELECTED"%>>ELIMINAR REGISTROS</option>
				</td>
				<td>
					<select name="CB_ACCION_REGISTRO" id="CB_ACCION_REGISTRO" style="width:200px;">
				</td>
				<td align="center">
					<input TYPE="BUTTON" class="fondo_boton_100" value="Procesar" name="B1" onClick="envia(CB_TIPO_PROCESO.value,CB_ACCION_REGISTRO.value);return false;">
				</td>
			</tr>
			
			<tr>
				<td class=hdr_i>
					IDENTIFICADOR<BR>
					<TEXTAREA NAME="TX_IDENTIFICADOR" ROWS=30 COLS=15><%=intId%></TEXTAREA>
				</td>
				<td colspan=2 class=hdr_i>
					NUEVO DATO<BR>
					<TEXTAREA NAME="TX_NUEVO_DATO" ROWS=30 COLS=80><%=strValorNuevo%></TEXTAREA>
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
		else if (intTipoProceso == '1'){
			if (confirm("¿ Está seguro de actualizar el estado de la Deuda de los id cuota ingresados?"))
			{
				if (confirm("¿ Está REALMENTE seguro de actualizar el estado de la Deuda ? Este proceso afectará varios informes relacionados a este registro"))
				{
					document.forms[0].action='utilitario_2.asp?strGraba=S';
					document.forms[0].submit();
				}
			}
		}
		else if (intTipoProceso == '13'){
			if (confirm("¿ Está seguro de modificar los Deudores Seleccionados ? Este proceso podría afectar la consistencia de la Base y la reportería Asociada"))
			{
				if (confirm("¿ Está REALMENTE seguro de modificar los Deudores Seleccionados?"))
				{
					document.forms[0].action='utilitario_2.asp?strGraba=S';
					document.forms[0].submit();
				}
			}
		}
		else if (intTipoProceso == '2'){
			if (confirm("¿ Está seguro de modificar los documentos ingresados ? Este proceso podría afectar la consistencia de la Base y la reportería Asociada"))
			{
				if (confirm("¿ Está REALMENTE seguro de modificar los documentos ingresados ?"))
				{
					document.forms[0].action='utilitario_2.asp?strGraba=S';
					document.forms[0].submit();
				}
			}
		}
		else if (intTipoProceso == '3'){
			if (confirm("¿ Está seguro de modificar los registros asociados a la Gestión ?"))
			{
				document.forms[0].action='utilitario_2.asp?strGraba=S';
				document.forms[0].submit();
			}
		}
		else if (intTipoProceso == '4'){
			if (confirm("¿ Está seguro de auditar la Contactabilidad del medio seleccionado ?"))
			{
				document.forms[0].action='utilitario_2.asp?strGraba=S';
				document.forms[0].submit();
			}
		}
		else if (intTipoProceso == '5'){
			if (confirm("¿ Está seguro de eliminar los documentos ingresados ?"))
			{
				if (confirm("¿ Está REALMENTE seguro de eliminar los registros seleccionados ? Este proceso es COMPLETAMENTE IRREVERSIBLE"))
				{
					document.forms[0].action='utilitario_2.asp?strGraba=S';
					document.forms[0].submit();
				}
			}
		}
		datos.B1.disabled = false;
}

function CargaRegistros(subCat,registro)
{
//
	var comboBox = document.getElementById('CB_ACCION_REGISTRO');
	comboBox.options.length = 0;

		if (subCat=='1') {
			var newOption = new Option('SELECCIONAR', '0');
			comboBox.options[comboBox.options.length] = newOption;
			var newOption = new Option('ABONO', 'ABONO');
			comboBox.options[comboBox.options.length] = newOption;
			var newOption = new Option('REBAJA PAGOS EN CLIENTE', 'REBAJA_PAGOS_EN_CLIENTE');
			comboBox.options[comboBox.options.length] = newOption;
			var newOption = new Option('REBAJA PAGOS EN EMPRESA', 'REBAJA_PAGOS_EN_EMPRESA');
			comboBox.options[comboBox.options.length] = newOption;
			var newOption = new Option('REBAJA PAGOS CONVENIO', 'REBAJA_PAGOS_CONVENIO');
			comboBox.options[comboBox.options.length] = newOption;
			var newOption = new Option('RETIRO POR CLIENTE', 'RETIRO');
			comboBox.options[comboBox.options.length] = newOption;
			var newOption = new Option('RETIRO POR RESOLUCION', 'RETIRO_RES');
			comboBox.options[comboBox.options.length] = newOption;
			var newOption = new Option('PENDIENTE', 'PENDIENTE');
			comboBox.options[comboBox.options.length] = newOption;
			var newOption = new Option('FIN COBRANZA', 'FIN COBRANZA');
			comboBox.options[comboBox.options.length] = newOption;
			var newOption = new Option('VOLVER A ACTIVAR', 'ACTIVAR');
			comboBox.options[comboBox.options.length] = newOption;
			var newOption = new Option('ERROR CARGA', 'ERROR CARGA');
			comboBox.options[comboBox.options.length] = newOption;
			
		}
		else if (subCat=='2') {
			var newOption = new Option('SELECCIONAR', '0');
			comboBox.options[comboBox.options.length] = newOption;
			var newOption = new Option('CODIGO CLIENTE', 'COD_CLIENTE');
			comboBox.options[comboBox.options.length] = newOption;
			var newOption = new Option('NUMERO DOCUMENTO', 'NRO_DOC');
			comboBox.options[comboBox.options.length] = newOption;
			var newOption = new Option('NUMERO CUOTA', 'NRO_CUOTA');
			comboBox.options[comboBox.options.length] = newOption;
			var newOption = new Option('FECHA VENCIMIENTO', 'FECHA_VENC');			
			comboBox.options[comboBox.options.length] = newOption;
			var newOption = new Option('MONTO ORIGINAL', 'VALOR_CUOTA');					
			comboBox.options[comboBox.options.length] = newOption;			
			var newOption = new Option('FECHA ESTADO', 'FECHA_ESTADO');			
			comboBox.options[comboBox.options.length] = newOption;
			var newOption = new Option('FECHA CARGA', 'FECHA_CREACION');			
			comboBox.options[comboBox.options.length] = newOption;
			var newOption = new Option('GASTOS PROTESTOS', 'GASTOS_PROTESTOS');
			comboBox.options[comboBox.options.length] = newOption;
			var newOption = new Option('TIPO DOCUMENTO', 'TIPO_DOCUMENTO');
			comboBox.options[comboBox.options.length] = newOption;			
			var newOption = new Option('SUCURSAL', 'SUCURSAL');
			comboBox.options[comboBox.options.length] = newOption;			
			var newOption = new Option('NUMERO CLIENTE DOC', 'NRO_CLIENTE_DOC');		
			comboBox.options[comboBox.options.length] = newOption;
			var newOption = new Option('NUMERO CLIENTE DEUDOR', 'NRO_CLIENTE_DEUDOR');
			comboBox.options[comboBox.options.length] = newOption;	
			var newOption = new Option('CODIGO GESTION EXTERNA', 'COD_GESTION_EXTERNA');
			comboBox.options[comboBox.options.length] = newOption;			
			var newOption = new Option('GESTION EXTERNA', 'DES_GESTION_EXTERNA');
			comboBox.options[comboBox.options.length] = newOption;			
			var newOption = new Option('RUT SUBCLIENTE', 'RUT_SUBCLIENTE');
			comboBox.options[comboBox.options.length] = newOption;			
			var newOption = new Option('NOMBRE SUBCLIENTE', 'NOMBRE_SUBCLIENTE');
			comboBox.options[comboBox.options.length] = newOption;			
			var newOption = new Option('CUSTODIO', 'CUSTODIO');
			comboBox.options[comboBox.options.length] = newOption;
			var newOption = new Option('ADIC 1', 'ADIC_1');
			comboBox.options[comboBox.options.length] = newOption;
			var newOption = new Option('ADIC 2', 'ADIC_2');
			comboBox.options[comboBox.options.length] = newOption;
			var newOption = new Option('ADIC 3', 'ADIC_3');
			comboBox.options[comboBox.options.length] = newOption;			
			var newOption = new Option('ADIC 4', 'ADIC_4');
			comboBox.options[comboBox.options.length] = newOption;			
			var newOption = new Option('ADIC 5', 'ADIC_5');
			comboBox.options[comboBox.options.length] = newOption;		
			var newOption = new Option('ADIC 91', 'ADIC_91');
			comboBox.options[comboBox.options.length] = newOption;
			var newOption = new Option('ADIC 92', 'ADIC_92');
			comboBox.options[comboBox.options.length] = newOption;
			var newOption = new Option('ADIC 93', 'ADIC_93');
			comboBox.options[comboBox.options.length] = newOption;
			var newOption = new Option('ADIC 94', 'ADIC_94');
			comboBox.options[comboBox.options.length] = newOption;
			var newOption = new Option('ADIC 95', 'ADIC_95');
			comboBox.options[comboBox.options.length] = newOption;	
			var newOption = new Option('ADIC 96', 'ADIC_96');
			comboBox.options[comboBox.options.length] = newOption;
			var newOption = new Option('ADIC 97', 'ADIC_97');
			comboBox.options[comboBox.options.length] = newOption;
			var newOption = new Option('ADIC 98', 'ADIC_98');
			comboBox.options[comboBox.options.length] = newOption;
			var newOption = new Option('ADIC 99', 'ADIC_99');
			comboBox.options[comboBox.options.length] = newOption;
			var newOption = new Option('ADIC 100', 'ADIC_100');
			comboBox.options[comboBox.options.length] = newOption;	
			var newOption = new Option('CAMPAÑA CLIENTE', 'CAMPANA_CLIENTE');
			comboBox.options[comboBox.options.length] = newOption;
			var newOption = new Option('TRAMO VENCIMIENTO', 'TRAMO_VENCIMIENTO');
			comboBox.options[comboBox.options.length] = newOption;	
			var newOption = new Option('TRAMO MONTO', 'TRAMO_MONTO');
			comboBox.options[comboBox.options.length] = newOption;	
			var newOption = new Option('TRAMO ASIGNACIÓN', 'TRAMO_ASIG');
			comboBox.options[comboBox.options.length] = newOption;			
			var newOption = new Option('ASIGNACIÓN', 'ASIGNACION');
			comboBox.options[comboBox.options.length] = newOption;
			var newOption = new Option('FECHA COMISION', 'FECHA_COMISION');
			comboBox.options[comboBox.options.length] = newOption;				
		}
		else if (subCat=='3') {
			var newOption = new Option('SELECCIONAR', '0');
			comboBox.options[comboBox.options.length] = newOption;
			var newOption = new Option('FECHA INGRESO', 'FECHA_INGRESO');
			comboBox.options[comboBox.options.length] = newOption;
			var newOption = new Option('CODIGO CLIENTE', 'COD_CLIENTE');
			comboBox.options[comboBox.options.length] = newOption;
		}
		else if (subCat=='4') {
			var newOption = new Option('SELECCIONAR', '0');
			comboBox.options[comboBox.options.length] = newOption;
			var newOption = new Option('ESTADO TELEFONO', 'AUDITAR_TELEFONOS');
			comboBox.options[comboBox.options.length] = newOption;
			var newOption = new Option('ESTADO CARGA TELEFONOS', 'ESTADO_CARGA_TELEFONOS');
			comboBox.options[comboBox.options.length] = newOption;
			var newOption = new Option('ESTADO EMAIL', 'AUDITAR_EMAIL');
			comboBox.options[comboBox.options.length] = newOption;
			var newOption = new Option('ESTADO DIRECCIONES', 'AUDITAR_DIRECCIONES');
			comboBox.options[comboBox.options.length] = newOption;

		}
		else if (subCat=='5') {
			var newOption = new Option('SELECCIONAR', '0');
			comboBox.options[comboBox.options.length] = newOption;
			var newOption = new Option('ELIMINAR DOCUMENTOS', 'ELIMINAR_DOCUMENTOS');
			comboBox.options[comboBox.options.length] = newOption;
			var newOption = new Option('ELIMINAR GESTIONES', 'ELIMINAR_GESTIONES');
			comboBox.options[comboBox.options.length] = newOption;
			var newOption = new Option('ELIMINAR TELEFONOS', 'ELIMINAR_TELEFONOS');
			comboBox.options[comboBox.options.length] = newOption;
			var newOption = new Option('ELIMINAR EMAIL', 'ELIMINAR_EMAIL');
			comboBox.options[comboBox.options.length] = newOption;
			var newOption = new Option('ELIMINAR DIRECCIONES', 'ELIMINAR_DIRECCIONES');
			comboBox.options[comboBox.options.length] = newOption;
		}
		else if (subCat=='13') {
			var newOption = new Option('SELECCIONAR', '0');
			comboBox.options[comboBox.options.length] = newOption;
			var newOption = new Option('MODIFICAR ETAPA COBRANZA', 'MODIFICAR_ETAPA_COBRANZA');
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
