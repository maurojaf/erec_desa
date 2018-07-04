<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html>
<html lang="es">
<HEAD>
    <meta http-equiv="X-UA-Compatible" content="IE=edge,chrome=1">
    <meta charset="utf-8">
	<!--#include file="arch_utils.asp"-->
	<!--#include file="../lib/comunes/rutinas/funciones.inc" -->
	<!--#include file="../lib/lib.asp"-->
	<!--#include file="sesion_inicio.asp"-->
	<!--#include file="../lib/asp/comunes/general/SoloNumeros.inc" -->
	<!--#include file="../lib/asp/comunes/general/rutinasBooleano.inc"-->
	<link href="../css/style_generales_sistema.css" rel="stylesheet">	
</head>
<body>
<%

Response.CodePage = 65001
Response.charset="utf-8"

rut 				=request("strRUT_DEUDOR")
strFonoAgestionarO 	=request("strFonoAgestionar")

If session("permite_no_validar_fonos") = "N" Then
	If TraeSiNo(session("perfil_adm"))="Si" or TraeSiNo(session("perfil_sup"))="Si" Then

	Else

		strNoValida = "disabled"
	End If

End If

abrirscg()

	strSql = "SELECT ISNULL(TIPO_SOFTPHONE,2) AS TIPO_SOFTPHONE "
	strSql = strSql & " FROM USUARIO "
	strSql = strSql & " WHERE ID_USUARIO = " & session("session_idusuario")

	set RsCli=Conn.execute(strSql)
	If not RsCli.eof then
		intTipoSoftPhone = RsCli("TIPO_SOFTPHONE")
	End if
	RsCli.close
	set RsCli=nothing

'Response.write "<br>intTipoSoftPhone=" & intTipoSoftPhone
		
cerrarscg()
%>




	  <%
	  abrirscg()
	  	strSql="SELECT DT.RUT_DEUDOR,IdTipoContacto, DIAS_ATENCION,HORA_DESDE, HORA_HASTA, ANEXO, [dbo].[fun_trae_estatus_telefono_solo] ('" & session("ses_codcli") & "', RUT_DEUDOR, ID_TELEFONO) as ANALISIS,"
		strSql = strSql & "ID_TELEFONO,COD_AREA,TELEFONO,CORRELATIVO,ESTADO,FECHA_INGRESO, ISNULL(TELEFONO_DAL,0) AS TELEFONO_DAL, DIAS_ATENCION, NOMBRE_ESTADO = ISNULL(EC.NOMBRE_ESTADO,'NO DEFINIDO') "
		strSql = strSql & "FROM DEUDOR_TELEFONO DT LEFT JOIN ESTADO_CONTACTABILIDAD EC ON DT.ID_ESTADO_CONTACTABILIDAD=EC.ID_ESTADO "
		strSql = strSql & "WHERE RUT_DEUDOR ='" & rut & "' AND ESTADO IN (2) "
		strSql = strSql & "ORDER BY ID_ESTADO_CONTACTABILIDAD ASC, FECHA_CONTACTABILIDAD DESC"
		set rsTel=Conn.execute(strSql)
		if rsTel.eof then
		%>
			<script>
				alert('No existen teléfonos no validos');
				carga_funcion_telefono()
			</script>
		<%
			Response.End
		Else
	  %>
	  <input type="hidden" name="pagina_origen" id="pagina_origen" value="deudor_telefono_nv">
	  &nbsp;
	  <table width="100%" border="0" class="intercalado" style="width:100%;">
	  	<input name="rut" type="hidden" id="rut" value="<%=rut%>">
	  	<thead>
        <tr >
			<td align = "center">TIPO</td>
			<td align = "center">ÁREA</td>
			<td >T&Eacute;LEFONO </td>
			<td align = "center">ANEXO</td>
			<td align = "center">TIPO CONTACTO</td>
			<td align = "center">DIAS ATENCION</td>
			<td colspan=1 align = "center">HORAS ATENCION</td>
			<td align = "center">CONTACTABILIDAD</td>
			<td align = "center">&nbsp;</td>
			<td align="center">ESTADO</td>
			<td align="center">
				<a href="#" onClick="envia('AF');" title="Auditar Fonos"><img src="../imagenes/phone.png" border="0"></a>
				<a href="#" onClick="envia('NF');" title="Nuevo Fono"><img src="../imagenes/phone_add.png" border="0"></a>
				<a href="#" onClick="carga_funcion_telefono();" title="Volver"><img src="../imagenes/arrow_left.png" border="0"></a>
			</td>
        </tr>
    	</thead>
		<%
		sinauditar=0
		novalida=0
		valida=0
		intIdContacto = 1
		Do until rsTel.eof
			FECHA_REVISION=rsTel("FECHA_INGRESO")
			if isNULL(FECHA_REVISION) then
				FECHA_REVISION=""
			end if
			
			strRutDeudor  			=rsTel("RUT_DEUDOR")
			COD_AREA=rsTel("COD_AREA")
			Telefono=rsTel("Telefono")
			correlativo_deudor=rsTel("CORRELATIVO")
			strTelefonoDal=rsTel("TELEFONO_DAL")
			strFonoAgestionar = COD_AREA & "-" & Telefono
			srtAnexo = UCASE(rsTel("ANEXO"))
			Estado=rsTel("Estado")
			if estado="0" then
				strEstadoFono="SIN AUDITAR"
			elseif estado="1" then
				strEstadoFono="VALIDO"
			elseif estado="2" then
				strEstadoFono="NO VALIDO"
			end if

			strAnalisis=Trim(rsTel("ANALISIS"))
			strHoraDesde=Trim(rsTel("HORA_DESDE"))
			strHoraHasta=Trim(rsTel("HORA_HASTA"))
			strDiasAtencion=Trim(rsTel("DIAS_ATENCION"))
			strEstadoContactabilidad =Trim(rsTel("NOMBRE_ESTADO"))

			if srtAnexo ="" then
				strLabelResto ="Sin información"
			else
				strLabelResto =srtAnexo
			end if

		%>
		<input type="hidden" name="correlativo_deudor" id ="correlativo_deudor" value="<%=trim(correlativo_deudor)%>">
        <tr bordercolor="#FFFFFF">

		  <td>

		  <%
		  if COD_AREA="9" then
		  	response.Write("CELULAR")
		  Elseif COD_AREA="0" then
		  	response.Write("SIN ESPECIF.")
		  else
		  	response.Write("RED FIJA")
		  end if

		  %>

          <td title="<%=rsTel("ID_TELEFONO")%>"><div align="CENTER"><%=COD_AREA%></div></td>
		  
		  
		<%If intTipoSoftPhone="1" then%>		  
          <td ><div align="left">
             &nbsp;<a href="SIP:<%=SoloNumeros(strTelefonoDal)%>"><%=Telefono%></a>
            </div>
          </td>
		<%ElseIf intTipoSoftPhone="2" Then%>
          <td ><div align="left">
             &nbsp;<a href="CALLTO:<%=SoloNumeros(strTelefonoDal)%>"><%=Telefono%></a>
            </div>
          </td>		
		 <%End If%>		
		
          	<td title="<%=strLabelResto%>"><div align="CENTER"><input name="TX_ANEXO_<%=correlativo_deudor%>" id="TX_ANEXO_<%=correlativo_deudor%>" type="text" value="<%=srtAnexo%>" size="30" maxlength="30"></td>
			
			<td><select id="cbxTipoContacto_<%=correlativo_deudor%>" name="cbxTipoContacto_<%=correlativo_deudor%>">
				<% if(rsTel("IdTipoContacto") <> "") THEN strSeleccionado = "selected" else strSeleccionado="" end if %>
				<option value="">Seleccione</option>
				
				<% 	strListaTipoContacto = "SELECT IdTipoContacto, Glosa, Descripcion FROM TipoContacto WHERE TipoDatoContacto = 'T' AND CodigoCliente = '"& session("ses_codcli") &"'"					
					set rsListaTipoContacto = Conn.execute(strListaTipoContacto)
					i = 1
					
					Do While Not rsListaTipoContacto.Eof
					if(rsListaTipoContacto("IdTipoContacto") = rsTel("IdTipoContacto")) THEN strSeleccionado = "selected" else strSeleccionado="" end if %>
						<option value="<%=rsListaTipoContacto("IdTipoContacto") %>" <%=strSeleccionado %> title="<%=rsListaTipoContacto("Descripcion") %>">
							<% response.write(i) %> - <%=rsListaTipoContacto("Glosa") %>
						</option>
				<% 	rsListaTipoContacto.movenext
					i = i + 1 
					Loop %>
			</select></td>

			<%

			strChequedLu = ""
			strChequedMa = ""
			strChequedMi = ""
			strChequedJu = ""
			strChequedVi = ""
			strChequedSa = ""

			If instr(strDiasAtencion,"LU") > 0 Then strChequedLu = "CHECKED"
			If instr(strDiasAtencion,"MA") > 0 Then strChequedMa = "CHECKED"
			If instr(strDiasAtencion,"MI") > 0 Then strChequedMi = "CHECKED"
			If instr(strDiasAtencion,"JU") > 0 Then strChequedJu = "CHECKED"
			If instr(strDiasAtencion,"VI") > 0 Then strChequedVi = "CHECKED"
			If instr(strDiasAtencion,"SA") > 0 Then strChequedSa = "CHECKED"
			%>
			<td>
			Lu
			<INPUT TYPE=CHECKBOX NAME="CH_DIAS_<%=correlativo_deudor%>" ID="CH_DIAS_<%=correlativo_deudor%>" value="LU" <%=strChequedLu%>>
			Ma
			<INPUT TYPE=CHECKBOX NAME="CH_DIAS_<%=correlativo_deudor%>" ID="CH_DIAS_<%=correlativo_deudor%>" value="MA" <%=strChequedMa%>>
			Mi
			<INPUT TYPE=CHECKBOX NAME="CH_DIAS_<%=correlativo_deudor%>" ID="CH_DIAS_<%=correlativo_deudor%>" value="MI" <%=strChequedMi%>>
			Ju
			<INPUT TYPE=CHECKBOX NAME="CH_DIAS_<%=correlativo_deudor%>" ID="CH_DIAS_<%=correlativo_deudor%>" value="JU" <%=strChequedJu%>>
			Vi
			<INPUT TYPE=CHECKBOX NAME="CH_DIAS_<%=correlativo_deudor%>" ID="CH_DIAS_<%=correlativo_deudor%>" value="VI" <%=strChequedVi%>>
			Sa
			<INPUT TYPE=CHECKBOX NAME="CH_DIAS_<%=correlativo_deudor%>" ID="CH_DIAS_<%=correlativo_deudor%>" value="SA" <%=strChequedSa%>>
            </td>
          	<td align = "center">
          		Desde
          		<input name="TX_DESDE_<%=correlativo_deudor%>" id="TX_DESDE_<%=correlativo_deudor%>" type="text" value="<%=strHoraDesde%>" size="4" maxlength="5" onChange="return ValidaHora(this,this.value)">
          		Hasta
				<input name="TX_HASTA_<%=correlativo_deudor%>" id="TX_HASTA_<%=correlativo_deudor%>" type="text" value="<%=strHoraHasta%>" size="4" maxlength="5" onChange="return ValidaHora(this,this.value)">
			</td>
			<td align="center"><%=strEstadoContactabilidad%></td>
			
			<td align="center">
			
							<a href="javascript:ventanaGestionesFonos('gestiones_por_telefono.asp?intIdFono=<%=rsTel("ID_TELEFONO")%>&strFonoAgestionar=<%=strFonoAgestionar%>&strRutDeudor=<%=strRutDeudor%>')">
							<img src="../imagenes/icon_gestiones.jpg" border="0">
						</a>
			</td>
			
			<td align="center">
				<div align="center"><span class="Estilo35">
				<input name="radiofon<%=correlativo_deudor%>" id="radiofon<%=correlativo_deudor%>" type="radio" value="1"
				<%if strEstadoFono="VALIDO" then
				Response.Write("checked")
				valida=valida+1
				end if%>>
				VA
				<input name="radiofon<%=correlativo_deudor%>" id="radiofon<%=correlativo_deudor%>" <%=strNoValida%> type="radio" value="2"
				<%if strEstadoFono="NO VALIDO" then
				Response.Write("checked")
				novalida=novalida+1
				end if%>>
				NV
				<input name="radiofon<%=correlativo_deudor%>" id="radiofon<%=correlativo_deudor%>" type="radio" value="0"
				<%if strEstadoFono="SIN AUDITAR" then
				Response.Write("checked")
				sinauditar=sinauditar+1
				end if%>>
				SA
				</span>
				</div>
			</td>
			<td align="center">
				<img src="../imagenes/Agrega_contacto.png" border="0" onclick="modifica_contacto('<%=rut%>','<%=rsTel("ID_TELEFONO")%>')">
			</td>
		</tr>
	<%
		intIdContacto = intIdContacto + 1
	rsTel.movenext
	Loop
	   %>
		<tr class="totales">
			<td colspan="8"><span class="">TOTAL</span></td>
			<td colspan="2"><span class="">NO VÁLIDOS : <%=novalida%></span></td>
			<td colspan="1" align="center">
				<a href="#" onClick="envia('AF');" title="Auditar Fonos"><img src="../imagenes/phone.png" border="0"></a>
				<a href="#" onClick="envia('NF');" title="Nuevo Fono"><img src="../imagenes/phone_add.png" border="0"></a>
				<a href="#" onClick="carga_funcion_telefono();" title="Volver"><img src="../imagenes/arrow_left.png" border="0"></a>
			</td>
		</tr>
		</form>
      </table>

	  <%
		end if
		rsTel.close
		set rsTel=nothing
		cerrarscg()
	  %>

</body>
</html>

