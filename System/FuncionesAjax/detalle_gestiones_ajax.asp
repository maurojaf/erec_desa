<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001" LCID = 1034%>
<!--#include file="../arch_utils.asp"-->
<!--#include file="../../lib/comunes/rutinas/funciones.inc" -->
<!--#include file="../../lib/comunes/rutinas/TraeCampo.inc" -->
<!--#include file="../../lib/asp/comunes/general/rutinasBooleano.inc"-->
<!--#include file="../../lib/lib.asp"-->

<head>
    <style type="text/css">
        .style1
        {
            width: 243px;
        }
        .style2
        {
            width: 33%;
        }
        .style4
        {
            width: 134px;
        }
        .style5
        {
            width: 92px;
        }
        .style6
        {
            width: 202px;
        }
        .style8
        {
            width: 133px;
        }
        .style9
        {
            width: 354px;
        }
    </style>
</head>

<%

Response.CodePage = 65001
Response.charset="utf-8"

    fonoActual = Request("fono_actual")
	strPaginaOrigen = Request("strPaginaOrigen")
	strCodCliente = Request("strCodCliente")

	abrirscg()

		strSql = "SELECT ISNULL(PERFIL_EMP ,0) AS PERFIL_EMP "
		strSql = strSql & " FROM USUARIO "
		strSql = strSql & " WHERE ID_USUARIO = " & session("session_idusuario")
		
		esPerfilEmpresa = false
	
		set RsCli=Conn.execute(strSql)
		If not RsCli.eof then
			esPerfilEmpresa = RsCli("PERFIL_EMP")
		End if
		RsCli.close
		set RsCli=nothing
			
	cerrarscg()


accion_ajax 		=request("accion_ajax")

abrirscg()


 


if trim(accion_ajax)="CARGA_ARCHIVOS_CUOTA" then

	strRutDeudor 		=request.querystring("strRut")
	strCodCliente       =request.querystring("strCodCliente")
	strObservaciones	=request.querystring("TX_OBSERVACIONES")
	IntCuota			=request.querystring("ID_CUOTA")

	sql_carga_cuota     ="exec proc_CARGA_ARCHIVOS_CUOTA '"&trim(strRutDeudor)&"', '"&TRIM(strCodCliente)&"','"&trim(ucase(strObservaciones))&"','"&trim(IntCuota)&"','"&trim(session("session_idusuario"))&"'"
	Conn.execute(sql_carga_cuota)
	if err then
		response.write sql_carga_cuota &" / ERROR :" &err.description
		response.end()
	end if



elseif trim(accion_ajax)="refresca_ubicabilidad" then

	strCodCliente       =request.querystring("strCodCliente")
	strRutDeudor		=request.querystring("rut")

	strDescripcionMedio	= request.querystring("descripcion_medio")

	ssql="EXEC proc_Parametros_Tabla_Cliente '"&TRIM(strRutDeudor)&"','"&TRIM(strCodCliente)&"'"

	'Response.write "<br>------ssql=" & ssql
	set rsCLI=Conn.execute(ssql)
	if not rsCLI.eof then
		strNomFormHon 		= ValNulo(rsCLI("FORMULA_HONORARIOS"),"C")
		strNomFormInt 		= ValNulo(rsCLI("FORMULA_INTERESES"),"C")

		strUsaSubCliente 	= rsCLI("USA_SUBCLIENTE")
		strUsaInteres 		= rsCLI("USA_INTERESES")
		strUsaHonorarios 	= rsCLI("USA_HONORARIOS")
		strUsaProtestos 	= rsCLI("USA_PROTESTOS")


		strRazonSocial 		= rsCLI("RAZON_SOCIAL")
		intRetiroSabado 	= Cint(rsCLI("RETIRO_SABADO"))
		strMsjRetiroSabado 	= ""

		If Trim(intRetiroSabado) = "1" Then
			strMsjRetiroSabado = "sabados,"
		End if

		strUbicFono 		= rsCLI("UBIC_FONO")
		strUbicEmail 		= rsCLI("UBIC_EMAIL")
		strUbicDireccion 	= rsCLI("UBIC_DIRECCION")
	end if


	if trim(strDescripcionMedio)="telefono" then

	 	If strUbicFono = "CONTACTADO" then %>

	   		<img src="../imagenes/mod_telefono_va.png" id="imagen_contacto" onclick="" style="cursor:pointer;" border="0">

	    <% ElseIf strUbicFono = "NO CONTACTADO" then %>

	    	 <img src="../imagenes/mod_telefono_sa.png" id="imagen_contacto" onclick="" style="cursor:pointer;" border="0">

	    <% Else %>

	    	 <img src="../imagenes/mod_telefono_nv.png" id="imagen_contacto" onclick="" style="cursor:pointer;" border="0" >

	    <% End If 

	elseif trim(strDescripcionMedio)="email"   then

		If trim(strUbicEmail) = "CONTACTADO" then %>

        	 <img src="../imagenes/mod_mail_va.png" border="0" id="imagen_email" onclick="" style="cursor:pointer;" >

        <% ElseIf trim(strUbicEmail) = "NO CONTACTADO" then %>

        	 <img src="../imagenes/mod_mail_sa.png" border="0" id="imagen_email" onclick="" style="cursor:pointer;">

        <% Else %>

        	 <img src="../imagenes/mod_mail_nv.png" border="0" id="imagen_email" onclick="" style="cursor:pointer;" >

        <% End If 


    elseif trim(strDescripcionMedio)="direccion" then

		If strUbicDireccion = "CONTACTADO" then %>

        	 <img src="../imagenes/mod_direccion_va.png" border="0"  style="cursor:pointer;" onclick="" id="imagen_direccion">

        <% ElseIf strUbicDireccion = "NO CONTACTADO" then %>

        	 <img src="../imagenes/mod_direccion_sa.png" border="0"  style="cursor:pointer;" onclick="" id="imagen_direccion">


        <% Else %>

        	 <img src="../imagenes/mod_direccion_nv.png" border="0"  onclick="" style="cursor:pointer;" id="imagen_direccion">

        <% End If 

	end if



elseif trim(accion_ajax)="mostrar_todos_cuotas" then
	strRutDeudor		=request.querystring("rut")
	strCodCliente 		=request.querystring("strCodCliente")
	strChTodosCuota		=request.querystring("CH_TODOS_CUOTA")
	IntIdGestion		=request.querystring("ID_GESTION")
	strIDCuotas 		=request.querystring("strIDCuotas")
	pagina_origen 		=request.querystring("pagina_origen")

	ssql="EXEC proc_Parametros_Tabla_Cliente '"&TRIM(strRutDeudor)&"','"&TRIM(strCodCliente)&"'"

	set rsCLI=Conn.execute(ssql)
	if not rsCLI.eof then
		strNomFormHon 		= ValNulo(rsCLI("FORMULA_HONORARIOS"),"C")
		strNomFormInt 		= ValNulo(rsCLI("FORMULA_INTERESES"),"C")

		strUsaSubCliente 	= rsCLI("USA_SUBCLIENTE")
		strUsaInteres 		= rsCLI("USA_INTERESES")
		strUsaHonorarios 	= rsCLI("USA_HONORARIOS")
		strUsaProtestos 	= rsCLI("USA_PROTESTOS")


		nombre_cliente		= rsCLI("RAZON_SOCIAL")
		intRetiroSabado		=Cint(rsCLI("RETIRO_SABADO"))
		strMsjRetiroSabado 	= ""

		If Trim(intRetiroSabado) = "1" Then
			strMsjRetiroSabado 	= "sabados,"
		End if

		strUbicFono 		=rsCLI("UBIC_FONO")
		strUbicEmail 		=rsCLI("UBIC_EMAIL")
		strUbicDireccion 	=rsCLI("UBIC_DIRECCION")

	end if
	
	strCobranza = Request("CB_COBRANZA")

	strSql ="exec proc_Trae_Cuotas_Deudor '"&trim(strCodCliente)&"','"&trim(strRutDeudor)&"','"&trim(strIDCuotas)&"','"&trim(IntIdGestion)&"','"&trim(strNomFormInt)&"', '"&trim(strNomFormHon)&"', '1', '"&trim(strChTodosCuota)&"', '" & strCobranza & "' "

	'response.write strSql&"<br>"
	'response.end

	set rsTemp= Conn.execute(strSql)

	intTasaMensual 		= 2/100
	intTasaDiaria 		= intTasaMensual/30
	intCorrelativo		= 1
	strArrID_CUOTA 		=""
	intTotSelSaldo 		= 0
	intTotSelIntereses 	= 0
	intTotSelProtestos 	= 0
	intTotSelHonorarios = 0
	strDetCuota 		="mas_datos_adicionales.asp"

	strArrConcepto 		= ""
	strArrID_CUOTA 		= ""
	
	strSqlCliente = "SELECT ISNULL(USA_CUSTODIO, 'N') AS USA_CUSTODIO FROM CLIENTE WHERE COD_CLIENTE = '" & strCodCliente & "'"
	
	set rsCliente= Conn.execute(strSqlCliente)
	
	strUsaCustodio = rsCliente("USA_CUSTODIO")

	%>
	<table  border="1" id="table_tablesorter"   class="tablesorter"  <%if trim(pagina_origen)<>"casos_objetados" then%> style="width:100%;" <%else%> style="width:90%;" align="center" <%end if%>bordercolor="#000000" cellSpacing="0" cellPadding="1">
	<thead>
		<tr >
			<%if trim(pagina_origen)<>"casos_objetados" then%>
				<td>&nbsp;</td>
			<%end if%>

			<%If Trim(strUsaSubCliente)="1" Then%>
				<th colspan = "2" >RUT CLIENTE</th>
				<th >NOMBRE CLIENTE</th>
			<%End If%>

			<th >N°DOC</th>
			<th >CUOTA</th>
			<th >FEC.VENC.</th>
			<th >ANT.</th>
			<th >TIPO DOC.</th>
			<th align="center" width="70">CAPITAL</th>
			<%If Trim(strUsaInteres)="1" Then%>
			<th class="cambio_flecha_ordenamiento" align="center" width="70">INTERES</th>
			<%End If%>
			<%If Trim(strUsaProtestos)="1" Then%>
			<th class="cambio_flecha_ordenamiento" align="center" width="80">PROTESTOS</th>
			<%End If%>
			<%If Trim(strUsaHonorarios)="1" Then%>
			<th class="cambio_flecha_ordenamiento" align="center" width="90">HONORARIOS</th>
			<%End If%>
			<th  align="center" width="70">ABONO</th>
			<th  align="center" width="70">SALDO</th>
			<th >FECHA AGEND.</th>
			<% If Trim(strUsaCustodio) = "S" Then %>
			<td class="HeaderWithoutSort">CUSTODIO</td>
			<% End If %>
			<%if trim(pagina_origen)<>"casos_objetados" then%>
				<td>&nbsp;</td>
			<%end if%>
			<td>&nbsp;</td>
			<td>&nbsp;</td>
			<td>&nbsp;</td>
			<td>&nbsp;</td>

		</tr>

	</thead>
	<tbody>
<%
	Do while not rsTemp.eof

		intSaldo 				=  rsTemp("SALDO")
		intValorCuota 			=  rsTemp("VALOR_CUOTA")
		intAbono 				= intValorCuota - intSaldo
		strNroDoc 				= rsTemp("NRO_DOC")
		strNroCuota				= rsTemp("NRO_CUOTA")
		strFechaVenc 			= rsTemp("FECHA_VENC")
		intProrroga 			= rsTemp("PRORROGA")
		strFechaVencOriginal 	= rsTemp("FECHA_VENC_ORIGINAL")
		strTipoDoc 				= rsTemp("TIPO_DOCUMENTO")
		intTipoGestion 			= rsTemp("TIPO_GESTION")
		intVerAgend 			= rsTemp("VER_AGEND")
		intGestionModulos 		= rsTemp("GESTION_MODULOS")
		strFechaAgendUG 		= rsTemp("FECHA_AGEND_ULT_GES")
		strCustodio				= rsTemp("CUSTODIO")

		intAntiguedad 			= ValNulo(rsTemp("ANTIGUEDAD"),"N")

		intIntereses 			= rsTemp("INTERESES")
		intHonorarios 			= rsTemp("HONORARIOS")

		intProtestos 			= ValNulo(rsTemp("GASTOS_PROTESTOS"),"N")

		intTotDoc 				= intSaldo+intIntereses+intProtestos+intHonorarios

		intTotSelSaldo 			= intTotSelSaldo+intSaldo
		intTotSelAbono 			= intTotSelAbono+intAbono
		intTotSelValorCuota 	= intTotSelValorCuota+intValorCuota

		intTotSelIntereses 		= intTotSelIntereses+intIntereses
		intTotSelProtestos 		= intTotSelProtestos+intProtestos
		intTotSelHonorarios 	= intTotSelHonorarios+intHonorarios
		intTotSelDoc 			= intTotSelDoc+intTotDoc

		strArrConcepto 			= strArrConcepto & ";" & "CH_" & rsTemp("ID_CUOTA")
		strArrID_CUOTA 			= strArrID_CUOTA & ";" & rsTemp("ID_CUOTA")

		%>
		<tr class="Estilo34">

			<input name="TX_SALDO_<%=Replace(rsTemp("ID_CUOTA"),"-","_")%>" id="TX_SALDO_<%=Replace(rsTemp("ID_CUOTA"),"-","_")%>" type="hidden" value="<%=intTotDoc%>">
			<input name="TX_CAPITAL_<%=Replace(rsTemp("ID_CUOTA"),"-","_")%>" id="TX_CAPITAL_<%=Replace(rsTemp("ID_CUOTA"),"-","_")%>" type="hidden" value="<%=intValorCuota%>">
			<input name="TX_HONORARIOS_<%=Replace(rsTemp("ID_CUOTA"),"-","_")%>" id="TX_HONORARIOS_<%=Replace(rsTemp("ID_CUOTA"),"-","_")%>" type="hidden" value="<%=intHonorarios%>">
			<input name="TX_INTERESES_<%=Replace(rsTemp("ID_CUOTA"),"-","_")%>" id="TX_INTERESES_<%=Replace(rsTemp("ID_CUOTA"),"-","_")%>" type="hidden" value="<%=intIntereses%>">
			<input name="TX_PROTESTOS_<%=Replace(rsTemp("ID_CUOTA"),"-","_")%>" id="TX_PROTESTOS_<%=Replace(rsTemp("ID_CUOTA"),"-","_")%>" type="hidden" value="<%=intProtestos%>">

			<%if trim(pagina_origen)<>"casos_objetados" then%>
				<TD width="12">				
					<INPUT TYPE="checkbox" <% if esPerfilEmpresa and Trim(strCustodio) <> "LLACRUZ" or not esPerfilEmpresa or Trim(strUsaCustodio) = "N" then %> checked="checked" <% else %> disabled="disabled" <% end if %> NAME="CH_ID_CUOTA" id="CH_ID_CUOTA" value="<%=rsTemp("ID_CUOTA")%>">				
				</TD>
			<%end if%>
			<%If Trim(strUsaSubCliente)="1" Then%>
				<td width="69"><%=rsTemp("RUT_SUBCLIENTE")%></td>
				<td>
					<a href="javascript:ventanaBusqueda('Busqueda.asp?strOrigen=1&TX_RUT_DEUDOR=<%=rsTemp("RUT_DEUDOR")%>&TX_NOMBRE=<%=nombre_deudor%>&TX_RUTSUBCLIENTE=<%=rsTemp("RUT_SUBCLIENTE")%>&TX_NOMBRE_SUBCLIENTE=<%=rsTemp("NOMBRE_SUBCLIENTE")%>')">
					<img src="../imagenes/buscar.png" border="0"></a></td>
				<td title="<%=rsTemp("NOMBRE_SUBCLIENTE")%>">
					<%=Mid(rsTemp("NOMBRE_SUBCLIENTE"),1,35)%>
				</td>
			<%End If%>

			<td><%=strNroDoc%></td>
			<td><%=strNroCuota%></td>

			<%If intProrroga = "0" Then%>
				<td><%=strFechaVenc%></td>
			<%Else%>
				<td bgcolor="#ff6666" title="Vencimiento Original: <%=strFechaVencOriginal%>">
				<%=strFechaVenc%></td>
			<%End If%>


			<td><%=intAntiguedad%></td>
			<td><%=strTipoDoc%></td>

			<td ALIGN="RIGHT"><SPAN style="display:none;"><%=intValorCuota%></SPAN><%=FN(intValorCuota,0)%></td>
			
			<%If Trim(strUsaInteres)="1" Then%>
				<td ALIGN="RIGHT">
					<SPAN style="display:none;"><%=intIntereses%></SPAN>
					<%=FN(intIntereses,0)%></td>
			<%End If%>
			<%If Trim(strUsaProtestos)="1" Then%>
				<td ALIGN="RIGHT"><SPAN style="display:none;"><%=intProtestos%></SPAN><%=FN(intProtestos,0)%></td>
			<%End If%>
			<%If Trim(strUsaHonorarios)="1" Then%>
				<td ALIGN="RIGHT"><SPAN style="display:none;"><%=intHonorarios%></SPAN><%=FN(intHonorarios,0)%></td>
			<%End If%>

			<td ALIGN="RIGHT"><SPAN style="display:none;"><%=intAbono%></SPAN><%=FN(intAbono,0)%></td>
			<td ALIGN="RIGHT"><SPAN style="display:none;"><%=intTotDoc%></SPAN><%=FN(intTotDoc,0)%></td>
		<td ALIGN="RIGHT"><%=strFechaAgendUG%></td>
			<% If Trim(strUsaCustodio) = "S" Then %>
				<td align="center">
				<% If Trim(strCustodio) <> "LLACRUZ" Then%>
					<div align="left"><img src="../imagenes/bolita7x8.jpg" border="0">&nbsp;<%=strCustodio%></div>
				<% Else %>
					<%=strCustodio%>
				<% End If%>
				</td>
			<% End If%>
			<%if trim(pagina_origen)<>"casos_objetados" then%>
			<td align="CENTER">
				<%
					intEstadoNR = ValNulo(rsTemp("NOTIFICACION_RECEPCIONADA"),"N")
					intEstadoFR = ValNulo(rsTemp("FACTURA_RECEPCIONADA"),"N")
				If (intEstadoNR = 0) OR (intEstadoFR = 0) Then
					strImagenGest = "audita_rojo.png"

				ElseIf (intEstadoNR = 2) OR (intEstadoFR = 2) Then
					strImagenGest = "audita_ama.png"

				Else
					strImagenGest = "audita_verde.png"
				End If
				%>

				<A HREF="#" onClick="AuditarDoc(<%=rsTemp("ID_CUOTA")%>)";>
				<img src="../imagenes/<%=strImagenGest%>" border="0"></A>
			</td>
			<%end if%>
			<td ALIGN="CENTER">
				<a href="javascript:ventanaGestionesPorDoc('gestiones_por_documento.asp?intID_CUOTA=<%=trim(rsTemp("ID_CUOTA"))%>&strCodCliente=<%=strCodCliente%>&strNroDoc=<%=trim(rsTemp("NRO_DOC"))%>')">
				<img src="../imagenes/icon_gestiones.jpg" border="0"></a>
			</td>

			<td>
				<a href="javascript:ventanaMas('<%=strDetCuota%>?ID_CUOTA=<%=trim(rsTemp("ID_CUOTA"))%>&cliente=<%=strCodCliente%>&strRUT_DEUDOR=<%=trim(rsTemp("RUT_DEUDOR"))%>&strNroDoc=<%=trim(rsTemp("NRO_DOC"))%>&strNroCuota=<%=rsTemp("NRO_CUOTA")%>')"><img src="../imagenes/Carpeta3.png" border="0"></a>
			</td>
			<td>
				<%IF trim(rsTemp("CANTIDAD_DOCUMENTOS"))>0 then%>
					<img src="../Imagenes/48px-Crystal_Clear_filesystem_folder_green.png" width="20" height="20" style="cursor:pointer;" alt="Historial documentos adjuntos" onclick="bt_ver_historial('<%=trim(rsTemp("ID_CUOTA"))%>')">
				<%else%>
					<img src="../Imagenes/48px-Crystal_Clear_filesystem_folder_red.png" width="20" height="20" style="cursor:pointer;" alt="Sin Historial documentos adjuntos" onclick="bt_ver_historial('<%=trim(rsTemp("ID_CUOTA"))%>')">
				<%end if%>
			</td>

			<td align="center">
				<%
				dtmFechaEstado 		= rsTemp("FECHA_ESTADO")
				dtmFechaCreacion 	= rsTemp("FECHA_CREACION")

				intIdUltGest 		= rsTemp("ID_ULT_GEST")

				dtmFechaIngresoUG 	= rsTemp("FECHA_INGRESO_UG")
				strCodUltgest 		= rsTemp("COD_ULT_GEST")

				strImagenGest1 		=""

				If (intVerAgend = 1 and ValNulo(rsTemp("DIFERENCIA"),"N") > 0) Then
					If (datevalue(dtmFechaIngresoUG) < datevalue(dtmFechaEstado)) and intGestionModulos = 3 Then
						''La fecha de ingreso de ultima gestion del documento (fun_trae_Ultima_Gestion_cuota_tit)es menor a la fecha de estado
						strImagenGest1 = "GestionarRoj.png"
					Else
						strImagenGest1 = "GestionarAzu.PNG"
					End If
				ElseIf (intTipoGestion = 1 or intTipoGestion = 2 ) Then
					'' Define VER AGEND en tabla GESTIONES_TIPO_GESTION igual a "0" y tipo de gestion compormiso pago o ruta
					strImagenGest1 = "NoGestionarAma.PNG"
				ElseIf intVerAgend = 0 or intTipoGestion = 3 or intTipoGestion = 4 Then
					'' Define VER AGEND en tabla GESTIONES_TIPO_GESTION igual a "0" dado a que gestión no se debe gestionar por el cobrador
					strImagenGest1 = "NoGestionarRojo.PNG"
				End If

				If strImagenGest1 <> "" Then %>
					<img src="../Imagenes/<%=strImagenGest1%>" border="0">
				<% Else %>
					&nbsp;
				<% End If %>
			</td>

		</tr>

		
		<%
		intCorrelativo = intCorrelativo + 1
		rsTemp.movenext
		loop

		vArrConcepto 		= split(strArrConcepto,";")
		vArrID_CUOTA 		= split(strArrID_CUOTA,";")
		intTamvConcepto 	= ubound(vArrConcepto)
		strArrID_CUOTA 		= Mid(strArrID_CUOTA,2,len(strArrID_CUOTA))
	%>
		</tbody>
		<thead class="totales">
		<tr class="Estilo34" height="22">
			<%If Trim(strUsaSubCliente)="1" Then
			 	strColspan = "colspan= 9"
			Else
				 strColspan = "colspan= 6"
			End If%>

			<td <%=strColspan%> >&nbsp;&nbsp;&nbsp;&nbsp;Totales :</td>
			<td ALIGN="RIGHT"><%=FN(intTotSelValorCuota,0)%></td>
			<%If Trim(strUsaInteres)="1" Then%>
				<td ALIGN="RIGHT"><%=FN(intTotSelIntereses,0)%></td>
			<%End If%>
			<%If Trim(strUsaProtestos)="1" Then%>
				<td ALIGN="RIGHT"><%=FN(intTotSelProtestos,0)%></td>
			<%End If%>
			<%If Trim(strUsaHonorarios)="1" Then%>
				<td ALIGN="RIGHT"><%=FN(intTotSelHonorarios,0)%></td>
			<%End If%>


			<td ALIGN="RIGHT"><%=FN(intTotSelAbono,0)%></td>
			<td ALIGN="RIGHT"><%=FN(intTotSelDoc,0)%></td>
			<% If Trim(strUsaCustodio) = "S" Then %>
			<td>&nbsp;</td>
			<% End If %>
			<td>&nbsp;</td>
			<%if trim(pagina_origen)<>"casos_objetados" then%>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
			<%end if%>
			
				<td Colspan="3" Rowspan="2" align="center"><input type="button" id="ButtonPlanPago" name="ButtonPlanPago" value="Plan de Pago" class="fondo_boton_100" /></td>
			
			</tr>

			<%if trim(pagina_origen)<>"casos_objetados" then%>
			<tr class="Estilo34" height="25">

			<td <%=strColspan%>>&nbsp;&nbsp;&nbsp;&nbsp;Totales Seleccionados:</td>
			<td ALIGN="RIGHT"><span id="span_TX_CAPITAL" style="font-weight:bold;">0</span>
				<INPUT TYPE="hidden" NAME="TX_CAPITAL" ID="TX_CAPITAL" DISABLED style="text-align:right;width:90" size="10" size="10" onkeyup="format(this)" onchange="format(this)">
			</td>



			<% If Trim(strUsaInteres)="1" Then%>
				<td ALIGN="RIGHT"><INPUT TYPE="TEXT" NAME="TX_INTERESES" ID="TX_INTERESES" DISABLED style="text-align:right;width:90" size="10" size="10" onkeyup="format(this)" onchange="format(this)"></td>
			<% Else%>
				<INPUT TYPE="hidden" NAME="TX_INTERESES">
			<% End If%>

			<% If Trim(strUsaProtestos)="1" Then%>
				<td ALIGN="RIGHT"><INPUT TYPE="TEXT" NAME="TX_PROTESTOS" ID="TX_PROTESTOS" DISABLED style="text-align:right;width:90" size="10" size="10" onkeyup="format(this)" onchange="format(this)"></td>
			<% Else%>
				<INPUT TYPE="hidden" NAME="TX_PROTESTOS">
			<% End If%>

			<% If Trim(strUsaHonorarios)="1" Then%>
				<td ALIGN="RIGHT"><INPUT TYPE="TEXT" NAME="TX_HONORARIOS" ID="TX_HONORARIOS" DISABLED style="text-align:right;width:90" size="10" onkeyup="format(this)" onchange="format(this)"></td>
			<% Else%>
				<INPUT TYPE="hidden" NAME="TX_HONORARIOS">
			<% End If%>



			<td>&nbsp;</td>
			<td ALIGN="RIGHT" ><span  id="span_TX_SALDO" style="font-weight:bold;">0</span>
				<INPUT TYPE="hidden" ID="TX_SALDO" NAME="TX_SALDO" DISABLED style="text-align:right;width:90" size="10" size="10" onkeyup="format(this)" onchange="format(this)">
			</td>
			<% If Trim(strUsaCustodio) = "S" Then %>
			<td>&nbsp;</td>
			<% End If %>
			<td>&nbsp;</td>
			<td>&nbsp;</td>
			<td>&nbsp;</td>

		</tr>
		<%end if%>
		</thead>
		<INPUT TYPE="hidden" NAME="strArrID_CUOTA" VALUE="<%=strArrID_CUOTA%>">
			
	</table>

<script type="text/javascript">
    $('input[id="CH_ID_CUOTA"]').change(function(){

        var contac_TX_CAPITAL	 	="#TX_CAPITAL_"+$(this).val()
        var contac_TX_INTERESES	 	="#TX_INTERESES_"+$(this).val()
        var contac_TX_HONORARIOS 	="#TX_HONORARIOS_"+$(this).val()
        var contac_TX_PROTESTOS  	="#TX_PROTESTOS_"+$(this).val()
        var contac_TX_SALDO	 	 	="#TX_SALDO_"+$(this).val()
        var TX_MONTO_CANCELADO 		=$('#TX_MONTO_CANCELADO').val()

        if($(this).is(':checked')){					

            $('#TX_CAPITAL').val(eval($('#TX_CAPITAL').val()) + eval($(contac_TX_CAPITAL).val()))
            $('#TX_INTERESES').val(eval($('#TX_INTERESES').val()) + eval($(contac_TX_INTERESES).val()))
            $('#TX_HONORARIOS').val(eval($('#TX_HONORARIOS').val()) + eval($(contac_TX_HONORARIOS).val()))
            $('#TX_PROTESTOS').val(eval($('#TX_PROTESTOS').val()) + eval($(contac_TX_PROTESTOS).val()))
            $('#TX_SALDO').val(eval($('#TX_SALDO').val()) + eval($(contac_TX_SALDO).val()))		

            $('#span_TX_SALDO').text($('#TX_SALDO').val())
            $('#span_TX_CAPITAL').text($('#TX_CAPITAL').val())
				
            if(TX_MONTO_CANCELADO!=null){
                $('#TX_MONTO_CANCELADO').val($('#TX_SALDO').val())
            }
        }else{

            $('#TX_CAPITAL').val(eval($('#TX_CAPITAL').val()) - eval($(contac_TX_CAPITAL).val()))
            $('#TX_INTERESES').val(eval($('#TX_INTERESES').val()) - eval($(contac_TX_INTERESES).val()))
            $('#TX_HONORARIOS').val(eval($('#TX_HONORARIOS').val()) - eval($(contac_TX_HONORARIOS).val()))
            $('#TX_PROTESTOS').val(eval($('#TX_PROTESTOS').val()) - eval($(contac_TX_PROTESTOS).val()))
            $('#TX_SALDO').val(eval($('#TX_SALDO').val()) - eval($(contac_TX_SALDO).val()))		

            $('#span_TX_SALDO').text($('#TX_SALDO').val())
            $('#span_TX_CAPITAL').text($('#TX_CAPITAL').val())

            if(TX_MONTO_CANCELADO!=null){
                $('#TX_MONTO_CANCELADO').val($('#TX_SALDO').val())
            }
        }
    })

    marcar_boxes();
		
		function marcar_boxes(){
			var TX_MONTO_CANCELADO =	$('#TX_MONTO_CANCELADO').val()
			/*
			datos.TX_CAPITAL.value = 0;
			datos.TX_INTERESES.value = 0;
			datos.TX_PROTESTOS.value = 0;
			datos.TX_HONORARIOS.value = 0;
			datos.TX_SALDO.value = 0;
			*/
			$("#TX_CAPITAL").val(0)
			$("#TX_INTERESES").val(0)
			$("#TX_HONORARIOS").val(0)
			$("#TX_PROTESTOS").val(0)
			$("#TX_SALDO").val(0)

			if(TX_MONTO_CANCELADO!=null){
				$('#TX_MONTO_CANCELADO').val(0)
			}


			$('#span_TX_SALDO').text(0)
			$('#span_TX_CAPITAL').text(0)	

			$('input[id="CH_ID_CUOTA"]').each(function(){
				
				if (!$(this).is(':disabled') && $(this).is(':visible')) {
				
					$(this).attr('checked', true);

					var contac_TX_CAPITAL	 ="#TX_CAPITAL_"+$(this).val()
					var contac_TX_INTERESES	 ="#TX_INTERESES_"+$(this).val()
					var contac_TX_HONORARIOS ="#TX_HONORARIOS_"+$(this).val()
					var contac_TX_PROTESTOS  ="#TX_PROTESTOS_"+$(this).val()
					var contac_TX_SALDO	 	 ="#TX_SALDO_"+$(this).val()

					$('#TX_CAPITAL').val(eval($('#TX_CAPITAL').val()) + eval($(contac_TX_CAPITAL).val()))
					$('#TX_INTERESES').val(eval($('#TX_INTERESES').val()) + eval($(contac_TX_INTERESES).val()))
					$('#TX_HONORARIOS').val(eval($('#TX_HONORARIOS').val()) + eval($(contac_TX_HONORARIOS).val()))
					$('#TX_PROTESTOS').val(eval($('#TX_PROTESTOS').val()) + eval($(contac_TX_PROTESTOS).val()))
					$('#TX_SALDO').val(eval($('#TX_SALDO').val()) + eval($(contac_TX_SALDO).val()))		

					$('#span_TX_SALDO').text($('#TX_SALDO').val())
					$('#span_TX_CAPITAL').text($('#TX_CAPITAL').val())
					

					if(TX_MONTO_CANCELADO!=null){
						$('#TX_MONTO_CANCELADO').val($('#TX_SALDO').val())
					}
					
				}

			})


		}

		function desmarcar_boxes(){
			var TX_MONTO_CANCELADO =	$('#TX_MONTO_CANCELADO').val()
			
			$("#TX_CAPITAL").val(0)
			$("#TX_INTERESES").val(0)
			$("#TX_HONORARIOS").val(0)
			$("#TX_PROTESTOS").val(0)
			$("#TX_SALDO").val(0)

			if(TX_MONTO_CANCELADO!=null){
				$('#TX_MONTO_CANCELADO').val(0)
			}


			$('#span_TX_SALDO').text(0)
			$('#span_TX_CAPITAL').text(0)	

			$('input[id="CH_ID_CUOTA"]').each(function(){	
				$(this).removeAttr('checked');	
			})

		}
		
		$("#ButtonPlanPago").click(function(){
		
			var anyCuotaChecked = false;
			
			paramsCuotaChecked = '';
		
			$("input[name='CH_ID_CUOTA']").each(function(){
			
				if ($(this).is(":checked")) {
				
					paramsCuotaChecked = paramsCuotaChecked + '&CH_' + $(this).val() + '=on';
				
				}
			
				anyCuotaChecked = anyCuotaChecked || $(this).is(":checked");
			
			});
			
			if (!anyCuotaChecked) {
			
				alert("Indique Cuota a Cancelar")
				return false;
			
			}
		
			var parametros = "alea=" + Math.random();
			
			$('#PlanPagoIngresoGestion').load("PlanPagoIngresoGestion.asp", parametros, function(responseText, textStatus, jqXHR){
			
				$("#PP_TX_CAPITAL").val($("#span_TX_CAPITAL").text());
				
				$("#PP_TX_HONORARIOS").val($("#TX_HONORARIOS").val());
				
				$("#PP_TX_TOTALDEUDA").val(parseInt($("#PP_TX_CAPITAL").val()) + parseInt($("#PP_TX_HONORARIOS").val()));
				
				$("#PP_TX_CAPITAL").val(FormatearNumero($("#PP_TX_CAPITAL").val()));
				
				$("#PP_TX_HONORARIOS").val(FormatearNumero($("#PP_TX_HONORARIOS").val()));
				
				$("#PP_TX_TOTALDEUDA").val(FormatearNumero($("#PP_TX_TOTALDEUDA").val()));
				
				$("#PP_TX_TOTALDEUDA_DESC").val($("#PP_TX_TOTALDEUDA").val());
				
				CalculateCapitalPercentageAndRefreshAgreement();
				
				$("#ButtonGenerarPlanPago").click(function(){
				
					window.open("plan_pago_convenio.asp?Origen=ingreso_gestion&CB_SEDE=LLACRUZ&CB_TIPO=RC" + paramsCuotaChecked + 
								'&PP_TX_CAPITAL=' + $("#PP_TX_CAPITAL").val() +
								'&PP_TX_INTERES=' + $("#PP_TX_INTERES").val() + 
								'&hdintIndemComp=0' + 
								'&PP_TX_GASTOSPROTESTOS=' + $("#PP_TX_GASTOSPROTESTOS").val() + 
								'&PP_TX_HONORARIOS=' + $("#PP_TX_HONORARIOS").val() + 
								'&hdintGastos=0' + 
								'&desc_capital=' + $("#desc_capital").val() + 
								'&desc_indemComp=' + $("#desc_indemComp").val() + 
								'&desc_honorarios=' + $("#desc_honorarios").val() + 
								'&desc_gastos=' + $("#desc_gastos").val() + 
								'&desc_interes=' + $("#desc_interes").val() + 
								'&desc_protestos=' + 
								'&pie=' + $("#pie").val() + 
								'&cuotas=' + ($("#cuotas").val() == '-' ? "1" : $("#cuotas").val()) + 
								'&PP_TX_DIAPAGO=' + $("#PP_TX_DIAPAGO").val());
				
				});
			
			});
		
			$('#PlanPagoIngresoGestion').dialog({
				show:"blind",
				hide:"explode",
				width:950,
				height:400,
				modal:true,
				title: 'Plan de Pago'
			});
		
		})		

    //marcar_boxes();

	</script>	

	<%



elseif trim(accion_ajax)="muestra_cajas_tipo_gestion" then

	cmbcat 			=request.querystring("cmbcat")
	cmbsubcat 		=request.querystring("cmbsubcat")
	cmbgest 		=request.querystring("cmbgest")
	strRutDeudor	=request.querystring("rut")
	strCodCliente 	=Request("strCodCliente")
	IntSaldo 		=request.querystring("TX_SALDO")
	'response.write cmbcat&"<br>"&cmbsubcat&"<br>"&cmbgest
    
	sql_tipo_gestion ="SELECT Cod_Cliente, COD_CATEGORIA, COD_SUB_CATEGORIA, COD_GESTION, "
	sql_tipo_gestion = sql_tipo_gestion & " TIPO_GESTION, Descripcion, MEDIO_ASOCIADO, Obligatoriedad "
	sql_tipo_gestion = sql_tipo_gestion & "FROM GESTIONES_TIPO_GESTION "
	sql_tipo_gestion = sql_tipo_gestion & "where cod_categoria = "&trim(cmbcat)&" and cod_sub_categoria="&trim(cmbsubcat)&" and cod_gestion=" & trim(cmbgest)
	sql_tipo_gestion = sql_tipo_gestion & " and Cod_Cliente= " & trim(strCodCliente)

	
	set rs_tipo_gestion = conn.execute(sql_tipo_gestion)

	if not rs_tipo_gestion.eof then

		strTipoGestion 		=rs_tipo_gestion("TIPO_GESTION")
		strDescripcion 		=rs_tipo_gestion("Descripcion")	
		IndMedioAsociado 	=rs_tipo_gestion("MEDIO_ASOCIADO")
		Obligatoriedad 		=rs_tipo_gestion("Obligatoriedad")

	end if

	'response.write "TIPO_GESTION :"&TIPO_GESTION &"<br>"
	'response.write "MEDIO_ASOCIADO :"&MEDIO_ASOCIADO &"<br>"

	AbrirScg1()	
	strSql = "SELECT TOP 1 GESTIONES.HORA_DESDE,GESTIONES.HORA_HASTA,FORMA_PAGO,UPPER(ISNULL(UPPER(RE.NOMBRE+' '+RE.UBICACION), upper(REPLACE(isnull(dd.CALLE,'') + ' ' + isnull(dd.NUMERO,'') + ' ' + isnull(dd.RESTO,'') + ' ' + isnull(dd.COMUNA,''),'  ',' ') ))) LUGAR_PAGO, ISNULL(DOC_GESTION,'') AS DOC_GESTION FROM GESTIONES "
	
	strSql = strSql & " LEFT JOIN FORMA_RECAUDACION RE ON RE.ID_FORMA_RECAUDACION= GESTIONES.ID_FORMA_RECAUDACION "
	strSql = strSql & " LEFT JOIN DEUDOR_DIRECCION DD ON DD.ID_DIRECCION= GESTIONES.ID_DIRECCION_COBRO_DEUDOR "

	strSql = strSql & " WHERE GESTIONES.COD_CLIENTE = '" & strCodCliente & "'"
	strSql = strSql & " AND GESTIONES.RUT_DEUDOR = '" & strRutDeudor & "'"
	strSql = strSql & " AND CAST(COD_CATEGORIA AS VARCHAR(2)) + '-' + CAST(COD_SUB_CATEGORIA AS VARCHAR(2)) + '-' + CAST(COD_GESTION AS VARCHAR(2)) IN (SELECT CAST(COD_CATEGORIA AS VARCHAR(2)) + '-' + CAST(COD_SUB_CATEGORIA AS VARCHAR(2)) + '-' + CAST(COD_GESTION AS VARCHAR(2)) FROM GESTIONES_TIPO_GESTION "

	strSql = strSql & " WHERE GESTION_MODULOS in(11,1) AND COD_CLIENTE = '" & strCodCliente & "')"
	strSql = strSql & " ORDER BY GESTIONES.FECHA_INGRESO DESC, GESTIONES.CORRELATIVO DESC"
   
   
   '    response.write strSql

	set rsPrevia=Conn1.execute(strSql)



	If not rsPrevia.eof Then
		strHoraDesde = rsPrevia("HORA_DESDE")
		strHoraHasta = rsPrevia("HORA_HASTA")
		strFormaPago = rsPrevia("FORMA_PAGO")
		strLugarPago = rsPrevia("LUGAR_PAGO")
		strDocgestion = rsPrevia("DOC_GESTION")

		vArrDocgestion = split(strDocgestion,"-")
	Else
		strHoraDesde = ""
		strHoraHasta = ""
		strFormaPago = ""
		strLugarPago = ""
		strDocgestion = ""
		vArrDocgestion = ""
		strSinGestionEsp= "1"
	End If

    'response.write ("asd --" & strLugarPago)
	
	ssql="EXEC proc_Parametros_Tabla_Cliente '"&TRIM(strRUT_DEUDOR)&"','"&TRIM(strCodCliente)&"'"

	set rsCLI=Conn1.execute(ssql)

     
	if not rsCLI.eof then
		strNomFormHon 		= ValNulo(rsCLI("FORMULA_HONORARIOS"),"C")
		strNomFormInt 		= ValNulo(rsCLI("FORMULA_INTERESES"),"C")

		strUsaSubCliente 	= rsCLI("USA_SUBCLIENTE")
		strUsaInteres 		= rsCLI("USA_INTERESES")
		strUsaHonorarios 	= rsCLI("USA_HONORARIOS")
		strUsaProtestos 	= rsCLI("USA_PROTESTOS")


		nombre_cliente		= rsCLI("RAZON_SOCIAL")
		intRetiroSabado		=Cint(rsCLI("RETIRO_SABADO"))
		strMsjRetiroSabado 	= ""

		If Trim(intRetiroSabado) = "1" Then
			strMsjRetiroSabado = "sabados,"
		End if

		strUbicFono 		=rsCLI("UBIC_FONO")
		strUbicEmail 		=rsCLI("UBIC_EMAIL")
		strUbicDireccion 	=rsCLI("UBIC_DIRECCION")
	end if

	strSql = "[dbo].[sp_genera_dias_inahabiles] "&trim(year(date()))&", " & trim(intRetiroSabado)


	set rsFeriados=Conn1.execute(strSql)



	strArrFeriados=""
	Do While not rsFeriados.eof
		strArrFeriados = strArrFeriados & "'" & rsFeriados("FECHA") & "',"
		rsFeriados.movenext
	Loop
	strArrFeriados = Mid(strArrFeriados,1,len(strArrFeriados)-1)

	CerrarScg1()

    
  


	if trim(strTipoGestion)="1" then
%>
		<div name="divCompPago" id="divCompPago" >
		<div class="subtitulo_informe">COMPROMISO DE PAGO</div>

		<table width="100%" class="estilo_columnas">
		<thead>
		<tr>
			 <td class="style2">FECHA COMPROMISO</td>
			 <td class="style9">FORMA DE PAGO</td>
			 <td class="style8">LUGAR DE PAGO</td>
			 <td width="33%" style="width: 16%">MONTO COMPROMISO</td>
		</tr>
		</thead>
		<tr>
		<td class="style2">
			<input name="TX_FECHA_COMPROMISO" id="TX_FECHA_COMPROMISO" readonly type="text" size="10" maxlength="10" onblur="ValidaDifFechas();">
		</td>
		<td class="style9">
			<select name="CB_FORMA_PAGO" id="CB_FORMA_PAGO" onChange="LugarPago();">
			<option value="">SELECCIONE</option>
			<%
			AbrirSCG1()
			ssql="SELECT * FROM FORMA_PAGO_CLIENTE WHERE ID_FORMA_PAGO NOT IN ('AB') AND COD_CLIENTE='" & strCodCliente & "'"
			set rsCLI=Conn1.execute(ssql)
			if not rsCLI.eof then
				do until rsCLI.eof
				%>
				<option value="<%=rsCLI("ID_FORMA_PAGO")%>"
				<%if Trim(strFormaPago)=Trim(rsCLI("ID_FORMA_PAGO")) then
					response.Write("Selected")
				end if%>
				><%=ucase(rsCLI("DESC_FORMA_PAGO"))%></option>

				<%rsCLI.movenext
				loop
			end if
			rsCLI.close
			set rsCLI=nothing
			CerrarSCG1()
			%>

			</select>

		</td>
		<td id="td_CB_ID_DIRECCION_COBRO_DEUDOR">
				<select name="CB_ID_DIRECCION_COBRO_DEUDOR1" 
                id="CB_ID_DIRECCION_COBRO_DEUDOR" onclick="return CB_ID_DIRECCION_COBRO_DEUDOR_onclick()">
				
                
                <option value="">SELECCIONE</option>
				<%
				AbrirSCG1()
                strSql ="select top 0 1"

                IF  strFormaPago = "TR" or  strFormaPago ="DP" THEN '' TRANSFERENCIA/DEPOSITO
                     strSql = " SELECT 2 AS TIPO, NOMBRE + ' ' + UBICACION AS LUGAR_PAGO , ORDEN , " & _
                                              " ID_FORMA_RECAUDACION ID, 'FORMA_RECAUDACION' TIPO " & _
                                              " FROM FORMA_RECAUDACION " & _
                                              " WHERE COD_CLIENTE = '" & TRIM(strCodCliente) & "' AND TIPO  = 'DEPOSITO' " & _
                                              " ORDER BY ORDEN ASC"
                END IF 

                IF  strFormaPago = "CF" or  strFormaPago ="CD"  or  strFormaPago ="EF" or strFormaPago = "CF" or strFormaPago = "TC" or strFormaPago = "TD" THEN  '' CHEQUE A FECHA/CHEQUE AL DIA/EFECTIVO/CF

                        strSql = "SELECT 1 as TIPO, REPLACE(isnull(CALLE,'') + ' ' + isnull(NUMERO,'') + ' ' + isnull(RESTO,'') + ' ' + isnull(COMUNA,''),'  ',' ')  as LUGAR_PAGO , " & _
                                 " 0 AS ORDEN, id_direccion ID , 'DIRECCION' tipo " & _
                                 " FROM DEUDOR_DIRECCION WHERE RUT_DEUDOR = '" & TRIM(strRutDeudor) & "' AND ESTADO <> 2 "
                    
				        strSql = strSql & " UNION"

                        strSql = strSql  & " SELECT 2 AS TIPO, NOMBRE + ' ' + UBICACION AS LUGAR_PAGO , ORDEN , " & _
                                      "ID_FORMA_RECAUDACION ID, 'FORMA_RECAUDACION' TIPO " & _
                                      " FROM FORMA_RECAUDACION " & _
                                      " WHERE COD_CLIENTE = '" & TRIM(strCodCliente) & "' AND TIPO  = 'PRESENCIAL' " & _
                                      " ORDER BY ORDEN ASC"
                    END IF 


                    IF  strFormaPago = "VV" or  strFormaPago ="PG"  or  strFormaPago ="LT" THEN '' VALE VISTA/PAGARE/LETRA

                                            strSql = "SELECT 1 as TIPO, REPLACE(isnull(CALLE,'') + ' ' + isnull(NUMERO,'') + ' ' + isnull(RESTO,'') + ' ' + isnull(COMUNA,''),'  ',' ')   as LUGAR_PAGO , " & _
                                                     " 0 AS ORDEN, id_direccion ID , 'DIRECCION' tipo " & _
                                                     " FROM DEUDOR_DIRECCION WHERE RUT_DEUDOR = '" & TRIM(strRutDeudor) & "' AND ESTADO <> 2"
                    
				                            strSql = strSql & " UNION"

                                            strSql = strSql  & " SELECT 2 AS TIPO, NOMBRE + ' ' + UBICACION AS LUGAR_PAGO , ORDEN , " & _
                                                          "ID_FORMA_RECAUDACION ID, 'FORMA_RECAUDACION' TIPO " & _
                                                          " FROM FORMA_RECAUDACION " & _
                                                          " WHERE COD_CLIENTE = '" & TRIM(strCodCliente) & "' AND TIPO  in ('PRESENCIAL','RETIRO') " & _
                                                          " ORDER BY ORDEN ASC"
                    END IF

				'strSql = "SELECT 1 as TIPO, REPLACE(CALLE + ' ' + NUMERO + ' ' + RESTO + ' ' + COMUNA,'  ',' ') as LUGAR_PAGO , 0 AS ORDEN, id_direccion ID , 'DIRECCION' tipo FROM DEUDOR_DIRECCION WHERE RUT_DEUDOR = '" & TRIM(strRutDeudor) & "' AND ESTADO <> 2"

				'strSql = strSql & " UNION"


				'strSql = strSql & " SELECT 2 as TIPO, NOMBRE + ' ' + UBICACION as LUGAR_PAGO , ORDEN, ID_FORMA_RECAUDACION ID, 'FORMA_RECAUDACION' tipo FROM FORMA_RECAUDACION WHERE COD_CLIENTE = '" & TRIM(strCodCliente) & "' ORDER BY ORDEN ASC"

				set rsDIR=Conn1.execute(strSql)
				do until rsDIR.eof
						direccion = rsDIR("LUGAR_PAGO")
						%>
						<option value="<%=rsDIR("ID")&"-"&rsDIR("TIPO")%>"
						<%if Trim(strLugarPago)=Trim(direccion) then
							response.Write("Selected")
						end if%>
						><%=direccion%></option>
						<%
						rsDIR.movenext
					loop
					rsDIR.close
					set rsDIR=nothing
				CerrarSCG1()
				%>
				</select>
		</td>
		<td>
			<input name="TX_MONTO_CANCELADO2" id="TX_MONTO_CANCELADO" 
                onblur="valida_entero(this,this.value)"  type="text" size="12" maxlength="15" 
                value="<%=IntSaldo%>"></td>
		</tr>
		</table>
		<table width="100%" class="estilo_columnas">
		<thead>
		</thead>
		<tr >
			<td id="td_CB_ID_DIRECCION_COBRO_DEUDOR">
				&nbsp;</td>
		</tr>
		</table>

		</div>

<%
	elseIf trim(strTipoGestion)="2" then 
%>
		<div name="divCompPagoRuta" id="divCompPagoRuta">
		<div class="subtitulo_informe">COMPROMISO DE PAGO RUTA</div>

		  <table width="100%" class="estilo_columnas">
		  	<thead>
		    <tr bordercolor="#999999" class="Estilo13" >
			     <td class="style1">FECHA COMPROMISO</td>
			     <td>H.DESDE</td>
			     <td class="style5">H.HASTA</td>
			     <td class="style4">FORMA DE PAGO</td>
			     <td class="style6">LUGAR DE PAGO</td>
			     <td>MONTO COMPROMISO</td>
		    </tr>
			</thead>

		   	<tr> 
				<td class="style1">
					<input name="TX_FECHA_COMPROMISO" id="TX_FECHA_COMPROMISO" type="text" size="10" maxlength="10" onBlur="ValidaDifFechas();" READONLY >
				</td>
				<td>
					<input name="TX_HORA_DESDE" type="text" id="TX_HORA_DESDE" value="<%=strHoraDesde%>" size="4" maxlength="5" onChange="return ValidaHora(this,this.value)">
				</td>
				<td class="style5">
					<input name="TX_HORA_HASTA" type="text" id="TX_HORA_HASTA" value="<%=strHoraHasta%>" size="4" maxlength="5"  onChange="return ValidaHora(this,this.value)">
				</td>
				<td class="style4">
					<select name="CB_FORMA_PAGO" id="CB_FORMA_PAGO" onChange="LugarPago();" >
					<option value="">SELECCIONE</option>
					<%
						AbrirSCG1()
						ssql="SELECT * FROM FORMA_PAGO_CLIENTE WHERE ID_FORMA_PAGO NOT IN ('AB') AND COD_CLIENTE='" & strCodCliente & "'"
						set rsCLI = Conn1.execute(ssql)
                        
						if not rsCLI.eof then
							do until rsCLI.eof
							%>
							<option value="<%=rsCLI("ID_FORMA_PAGO")%>"
							<%if Trim(strFormaPago)=Trim(rsCLI("ID_FORMA_PAGO")) then
								response.Write("Selected")
							end if%>
							><%=ucase(rsCLI("DESC_FORMA_PAGO"))%></option>

							<%rsCLI.movenext
							loop
						end if
						rsCLI.close
						set rsCLI=nothing
						CerrarSCG1()
					%>

					</select>

				</td>
				<td id="td_CB_ID_DIRECCION_COBRO_DEUDOR">
					&nbsp;<select name="CB_ID_DIRECCION_COBRO_DEUDOR2" 
                        id="CB_ID_DIRECCION_COBRO_DEUDOR">
					<option value="">SELECCIONE</option>
					<%
                    AbrirSCG1()
                      strSql ="select top 0 1"
   
                     IF  strFormaPago = "TR" or  strFormaPago ="DP" THEN '' TRANSFERENCIA/DEPOSITO
                     strSql = " SELECT 2 AS TIPO, NOMBRE + ' ' + UBICACION AS LUGAR_PAGO , ORDEN , " & _
                                              " ID_FORMA_RECAUDACION ID, 'FORMA_RECAUDACION' TIPO " & _
                                              " FROM FORMA_RECAUDACION " & _
                                              " WHERE COD_CLIENTE = '" & TRIM(strCodCliente) & "' AND TIPO  = 'DEPOSITO' " & _
                                              " ORDER BY ORDEN ASC"
                    END IF 

                   IF  strFormaPago = "CF" or  strFormaPago ="CD"  or  strFormaPago ="EF" or strFormaPago = "CF" or strFormaPago = "TC" or strFormaPago = "TD" THEN  '' CHEQUE A FECHA/CHEQUE AL DIA/EFECTIVO/CF

                        strSql = "SELECT 1 as TIPO, REPLACE(isnull(CALLE,'') + ' ' + isnull(NUMERO,'') + ' ' + isnull(RESTO,'') + ' ' + isnull(COMUNA,''),'  ',' ')  as LUGAR_PAGO , " & _
                                 " 0 AS ORDEN, id_direccion ID , 'DIRECCION' tipo " & _
                                 " FROM DEUDOR_DIRECCION WHERE RUT_DEUDOR = '" & TRIM(strRutDeudor) & "' AND ESTADO <> 2 "
                    
				        strSql = strSql & " UNION"

                        strSql = strSql  & " SELECT 2 AS TIPO, NOMBRE + ' ' + UBICACION AS LUGAR_PAGO , ORDEN , " & _
                                      "ID_FORMA_RECAUDACION ID, 'FORMA_RECAUDACION' TIPO " & _
                                      " FROM FORMA_RECAUDACION " & _
                                      " WHERE COD_CLIENTE = '" & TRIM(strCodCliente) & "' AND TIPO  = 'PRESENCIAL' " & _
                                      " ORDER BY ORDEN ASC"
                    END IF 


                    IF  strFormaPago = "VV" or  strFormaPago ="PG"  or  strFormaPago ="LT" THEN '' VALE VISTA/PAGARE/LETRA

                                            strSql = "SELECT 1 as TIPO, REPLACE(isnull(CALLE,'') + ' ' + isnull(NUMERO,'') + ' ' + isnull(RESTO,'') + ' ' + isnull(COMUNA,''),'  ',' ')  as LUGAR_PAGO , " & _
                                                     " 0 AS ORDEN, id_direccion ID , 'DIRECCION' tipo " & _
                                                     " FROM DEUDOR_DIRECCION WHERE RUT_DEUDOR = '" & TRIM(strRutDeudor) & "' AND ESTADO <> 2"
                    
				                            strSql = strSql & " UNION"

                                            strSql = strSql  & " SELECT 2 AS TIPO, NOMBRE + ' ' + UBICACION AS LUGAR_PAGO , ORDEN , " & _
                                                          "ID_FORMA_RECAUDACION ID, 'FORMA_RECAUDACION' TIPO " & _
                                                          " FROM FORMA_RECAUDACION " & _
                                                          " WHERE COD_CLIENTE = '" & TRIM(strCodCliente) & "' AND TIPO  in ('PRESENCIAL','RETIRO') " & _
                                                          " ORDER BY ORDEN ASC"
                    END IF 


					'strSql = "SELECT 1 as TIPO, REPLACE(CALLE + ' ' + NUMERO + ' ' + RESTO + ' ' + COMUNA,'  ',' ') as LUGAR_PAGO , 0 AS ORDEN, id_direccion ID , 'DIRECCION' tipo FROM DEUDOR_DIRECCION WHERE RUT_DEUDOR = '" & strRutDeudor & "' AND ESTADO <> 2"

					'strSql = strSql & " UNION"

					'strSql = strSql & " SELECT 2 as TIPO, NOMBRE + ' ' + UBICACION as LUGAR_PAGO , ORDEN, ID_FORMA_RECAUDACION ID, 'FORMA_RECAUDACION' tipo FROM FORMA_RECAUDACION WHERE COD_CLIENTE = '" & TRIM(strCodCliente) & "' ORDER BY ORDEN ASC"
                    'response.write(strSql)
                    'response.end

					set rsDIR=Conn1.execute(strSql)

					do until rsDIR.eof
                        
						direccion = rsDIR("LUGAR_PAGO")
                        'response.write(strLugarPago & "<br>")
                        'response.write(direccion & "<sdf>")
						%>
						<option value="<%=rsDIR("ID")&"-"&rsDIR("TIPO")%>"
						<%if Trim(strLugarPago)=Trim(direccion) then
							response.Write("Selected")
						end if%>
						><%=direccion%></option>
						<%
						rsDIR.movenext
					loop
					rsDIR.close
					set rsDIR=nothing
					CerrarSCG1()
					%>
					</select>
                </td>
				<td >
					&nbsp;<input name="TX_MONTO_CANCELADO1" id="TX_MONTO_CANCELADO" 
                        onblur="valida_entero(this,this.value)" type="text" size="12" maxlength="15" 
                        value="<%=IntSaldo%>"></td>
			</tr>

			<table width="100%" class="estilo_columnas">
			<thead>
			<tr>
				 <td colspan = "4">DOCUMENTOS NECESARIOS</td>
			</tr>
			</thead>
			<tr >
				<td colspan = "3" height="25" >
				<%
					AbrirSCG1()

					if strSinGestionEsp= "1" Then
						intTamvDocGestion = -1
					Else
						intTamvDocGestion = ubound(vArrDocgestion)
					End If

					'Response.write "intTamvDocGestion=" & intTamvDocGestion
					'Response.End

					strSql = "SELECT * FROM TIPO_DOCUMENTOS_GESTION WHERE COD_CLIENTE = '" & strCodCliente & "'"
					set rsDoc=Conn1.execute(strSql)
					Do until rsDoc.eof
						strCheckDocGest = ""
						For I=0 To intTamvDocGestion
							If Trim(vArrDocgestion(I)) = Trim(rsDoc("NOM_DOCUMENTO")) Then
								strCheckDocGest = "CHECKED"
								exit for
							End If
						Next
				%>
						<INPUT TYPE="checkbox" NAME="CK_DOC_GESTION" ID="CK_DOC_GESTION" value="<%=rsDoc("NOM_DOCUMENTO")%>" <%=strCheckDocGest%>>
						&nbsp;<%=rsDoc("NOM_DOCUMENTO")%>

				 <%
				 	rsDoc.movenext
				 	Loop
				 	CerrarSCG1()
				 %>
				 </td>
			</tr>
			<thead>
			<tr>			
		    	<td width="50%" >MÁS DOCUMENTOS NECESARIOS</td>
		     	<td width="17%" >FONO COBRO</td>
			 	<td width="33%" >CONTACTO COBRO</td>
				
		   	</tr>
		   </thead>
		   	 <tr>
				<td>
					<input name="TX_DOC_GESTION_NECESARIOS" type="text" id="TX_DOC_GESTION_NECESARIOS" size="44" maxlength="200">
				</td>

				<td id="td_CB_ID_FONO_COBRO"> 
					<select name="CB_ID_FONO_COBRO" id="CB_ID_FONO_COBRO" onchange="set_CB_ID_CONTACTO_FONO_COBRO(this.value);">

					<option value="">SELECCIONE</option>
					<%if fono_con="0" or fono_con="" then%>
					  <%
						AbrirSCG1()
						ssql_ = "SELECT ID_TELEFONO, TELEFONO, COD_AREA FROM DEUDOR_TELEFONO WHERE RUT_DEUDOR='" & strRutDeudor & "' AND ESTADO <> 2"
						set rsFON=Conn1.execute(ssql_)
						Do until rsFON.eof
							strFonoCB = rsFON("COD_AREA") & "-" & rsFON("Telefono")
							strSel=""
							if strFonoCB = strFonoAgend Then strSel = "SELECTED"
							%>
							<option value="<%=rsFON("ID_TELEFONO")%>" <%=strSel%>><%=rsFON("COD_AREA")%>-<%=rsFON("Telefono")%></option>
							<%
								rsFON.movenext
						Loop
						rsFON.close
						set rsFON=nothing
						CerrarSCG1()
					 %>
					<%else%>
						<option value="<%=fono_con%>"><%=area_con%>-<%=fono_con%></option>
					<%end if %>
					</select>
				</td>

				<td id="td_CB_ID_CONTACTO_FONO_COBRO">
					<select name="CB_ID_CONTACTO_FONO_COBRO" id="CB_ID_CONTACTO_FONO_COBRO" >
					<option value="">SELECCIONE</option>
					</select>
				</td>
			 </tr>

			</table>

		     
			 <input name="validar_fono" type="hidden" value="0"> <!-- invalidar=1, validar=2 nada=0-->
			 <input name="rut" id="rut" type="hidden" value="<%=strRutDeudor%>">

			 <input name="strMasTelefonos" type="hidden" value="<%=strMasTelefonos%>">
			 <input name="strMasDirecciones" type="hidden" value="<%=strMasDirecciones%>">
			 <input name="strMasEmail" type="hidden" value="<%=strMasEmail%>">

		</div>


<%


	elseIf trim(strTipoGestion)="3" then

%>
		<div name="divNormalizacion" id="divNormalizacion" >
		<div class="subtitulo_informe">> GESTION DE NORMALIZACIÓN</div>

		<table width="100%" class="estilo_columnas">
		<thead>
		<tr >
			 <td width="16%">FECHA PAGO</td>
			 <td width="17%">FORMA PAGO</td>
			 <td width="34%">LUGAR PAGO</td>
			 <td>NRO.COMPROBANTE</td>
			 <td>MONTO PAGO</td>
			 <td>ENVIO HRD</td>
		</tr>
		</thead>

		<tr >
			<td>
				<input name="TX_FECHA_PAGO" readonly id="TX_FECHA_PAGO" type="text" size="10" maxlength="10">
			</td>
			<td>
				<select name="CB_FORMA_PAGO" id="CB_FORMA_PAGO">
				<option value="">SELECCIONE</option>
				<%
				AbrirSCG1()
					ssql="SELECT * FROM FORMA_PAGO_CLIENTE WHERE ID_FORMA_PAGO NOT IN ('AB') AND COD_CLIENTE='" & strCodCliente & "'"
					set rsCLI=Conn1.execute(ssql)
					if not rsCLI.eof then
						do until rsCLI.eof
						%>
						<option value="<%=rsCLI("ID_FORMA_PAGO")%>"
						<%if Trim(strCodCliente)=Trim(rsCLI("DESC_FORMA_PAGO")) then
							response.Write("Selected")
						end if%>
						><%=ucase(rsCLI("DESC_FORMA_PAGO"))%></option>

						<%rsCLI.movenext
						loop
					end if
					rsCLI.close
					set rsCLI=nothing
				CerrarSCG1()
				%>
				</select>
			</td>

			<td id="td_CB_ID_DIRECCION_COBRO_DEUDOR">
				<select name="CB_ID_DIRECCION_COBRO_DEUDOR" id="CB_ID_DIRECCION_COBRO_DEUDOR">
				<option value="">SELECCIONE</option>
				<%
				AbrirSCG1()

				strSql = "SELECT 1 as TIPO, REPLACE(isnull(CALLE,'') + ' ' + isnull(NUMERO,'') + ' ' + isnull(RESTO,'') + ' ' + isnull(COMUNA,''),'  ',' ')  as LUGAR_PAGO , 0 AS ORDEN, id_direccion ID , 'DIRECCION' tipo FROM DEUDOR_DIRECCION WHERE RUT_DEUDOR = '" & strRutDeudor & "' AND ESTADO <> 2"
				strSql = strSql & " UNION"

				strSql = strSql & " SELECT 2 as TIPO, NOMBRE + ' ' + UBICACION as LUGAR_PAGO , ORDEN, ID_FORMA_RECAUDACION ID, 'FORMA_RECAUDACION' tipo FROM FORMA_RECAUDACION WHERE COD_CLIENTE = '" & TRIM(strCodCliente) & "' ORDER BY ORDEN ASC"

				set rsDIR=Conn1.execute(strSql)
				do until rsDIR.eof
					direccion = rsDIR("LUGAR_PAGO")
					%>
					<option value="<%=rsDIR("ID")&"-"&rsDIR("TIPO")%>"
					<%if Trim(strLugarPago)=Trim(direccion) then
						response.Write("Selected")
					end if%>
					><%=direccion%></option>
					<%
				rsDIR.movenext
				loop

				rsDIR.close
				set rsDIR=nothing
				CerrarSCG1()
				%>
				</select>
			</td>

			</td>
			<td><input name="TX_NRO_DOC_PAGO" type="text" id="TX_NRO_DOC_PAGO" size="10" maxlength="10"></td>
			<td><input name="TX_MONTO_CANCELADO" id="TX_MONTO_CANCELADO" onblur="valida_entero(this,this.value)" type="text" size="12" maxlength="15" value="<%=IntSaldo%>">

			<td>
				<select name="CB_ENVIO_HRD" id="CB_ENVIO_HRD">
				<option value="">SELECCIONE</option>
				<option value="1">SI</option>
				<option value="0">NO</option>
				</select>

			</td>

		</tr>
		</table>
		</div>

<%

	elseIf trim(strTipoGestion)="4" then
%>
		<div name="divNormalizacion1" id="divNormalizacion1" >
		<div class="subtitulo_informe">GESTION DE OBJECIÓN</div>

		<table width="100%" class="estilo_columnas">
		<thead>
		<tr >
			 <td width="16%">FECHA GESTION</td>
			 <td width="17%">MOTIVO</td>
			 <td width="34%">LUGAR GESTION</td>
			 <td>NRO.COMPROBANTE</td>
			 <td>MONTO ASOCIADO</td>
			 <td>ENVIO HRD</td>
		</tr>
		</thead>
		<tr>
			<td>
				<input name="TX_FECHA_PAGO" id="TX_FECHA_PAGO" readonly type="text" size="10" maxlength="10">
			</td>
			<td>
				<select name="CB_FORMA_PAGO" id="CB_FORMA_PAGO">
				<option value="">SELECCIONE</option>
				<%
				AbrirSCG1()
					ssql="SELECT * FROM FORMA_NORMALIZACION WHERE COD_CLIENTE = '" & strCodCliente & "'"
					set rsFormN=Conn1.execute(ssql)
					if not rsFormN.eof then
						do until rsFormN.eof
						%>
						<option value="<%=rsFormN("ID_FORMA_NORM")%>"><%=ucase(rsFormN("FORMA_NORM"))%></option>

						<%rsFormN.movenext
						loop
					end if
					rsFormN.close
					set rsFormN=nothing
				CerrarSCG1()
				%>

				</select>
			</td>
			<td id="td_CB_ID_DIRECCION_COBRO_DEUDOR">
				<select name="CB_ID_DIRECCION_COBRO_DEUDOR" id="CB_ID_DIRECCION_COBRO_DEUDOR">
					<option value="">SELECCIONE</option>
					<%
					AbrirSCG1()

					strSql = "SELECT 1 as TIPO, REPLACE(isnull(CALLE,'') + ' ' + isnull(NUMERO,'') + ' ' + isnull(RESTO,'') + ' ' + isnull(COMUNA,''),'  ',' ') as LUGAR_PAGO , 0 AS ORDEN, id_direccion ID , 'DIRECCION' tipo FROM DEUDOR_DIRECCION WHERE RUT_DEUDOR = '" & strRutDeudor & "' AND ESTADO <> 2"
					strSql = strSql & " UNION"

					strSql = strSql & " SELECT 2 as TIPO, NOMBRE + ' ' + UBICACION as LUGAR_PAGO , ORDEN, ID_FORMA_RECAUDACION ID, 'FORMA_RECAUDACION' tipo FROM FORMA_RECAUDACION WHERE COD_CLIENTE = '" & TRIM(strCodCliente) & "' ORDER BY ORDEN ASC"

					set rsDIR=Conn1.execute(strSql)
					do until rsDIR.eof
						direccion = rsDIR("LUGAR_PAGO")
						%>
						<option value="<%=rsDIR("ID")&"-"&rsDIR("TIPO")%>"
						<%if Trim(strLugarPago)=Trim(direccion) then
							response.Write("Selected")
						end if%>
						><%=direccion%></option>
						<%
						rsDIR.movenext
					loop
					rsDIR.close
					set rsDIR=nothing
					CerrarSCG1()
					%>
				</select>
			</td>

			</td>
			<td><input name="TX_NRO_DOC_PAGO" type="text" id="TX_NRO_DOC_PAGO" size="10" maxlength="10"></td>
			<td><input name="TX_MONTO_CANCELADO" onblur="valida_entero(this,this.value)" id="TX_MONTO_CANCELADO"  type="text" size="12" maxlength="15" value="<%=IntSaldo%>"></td>

			<td>
				<select name="CB_ENVIO_HRD" id="CB_ENVIO_HRD">
					<option value="">SELECCIONE</option>
					<option value="1">SI</option>
					<option value="0">NO</option>
				</select>
			</td>
		</tr>
		</table>
		</div>


<%


	elseIf trim(strTipoGestion)="5" then

%>
		<div name="divGestionTerreno" id="divGestionTerreno">
			<div class="subtitulo_informe">GESTIÓN TERRENO</div>

			<table width="100%" class="estilo_columnas">
			<thead>
				<tr >
					<td>FECHA GESTIÓN</td>
					<td>H.DESDE</td>
					<td>H.HASTA</td>
					<td width="33%">LUGAR GESTION</td>
				</tr>
			</thead>

			<tr>
				<td>
					<input name="TX_FECHA_COMPROMISO" id="TX_FECHA_COMPROMISO" type="text" size="10" maxlength="10" readonly onBlur="ValidaDifFechas();">
				</td>
				<td>
					<input name="TX_HORA_DESDE" type="text" id="TX_HORA_DESDE" value="<%=strHoraDesde%>" size="4" maxlength="5" onChange="return ValidaHora(this,this.value)">
				</td>
				<td>
					<input name="TX_HORA_HASTA" type="text" id="TX_HORA_HASTA" value="<%=strHoraHasta%>" size="4" maxlength="5"  onChange="return ValidaHora(this,this.value)">
				</td>
				<td id="td_CB_ID_DIRECCION_COBRO_DEUDOR">
					<select name="CB_ID_DIRECCION_COBRO_DEUDOR" id="CB_ID_DIRECCION_COBRO_DEUDOR">
					<option value="">SELECCIONE</option>
					<%
					AbrirSCG1()
					strSql = "SELECT 1 as TIPO, REPLACE(isnull(CALLE,'') + ' ' + isnull(NUMERO,'') + ' ' + isnull(RESTO,'') + ' ' + isnull(COMUNA,''),'  ',' ')  as LUGAR_PAGO , 0 AS ORDEN, id_direccion ID , 'DIRECCION' tipo FROM DEUDOR_DIRECCION WHERE RUT_DEUDOR = '" & strRutDeudor & "' AND ESTADO <> 2"

					set rsDIR=Conn1.execute(strSql)
					do until rsDIR.eof
						direccion = rsDIR("LUGAR_PAGO")
						%>
						<option value="<%=rsDIR("ID")&"-"&rsDIR("TIPO")%>"
						<%if Trim(strLugarPago)=Trim(direccion) then
							response.Write("Selected")
						end if%>
						><%=direccion%></option>
						<%
						rsDIR.movenext
					loop
					rsDIR.close
					set rsDIR=nothing
					CerrarSCG1()
					%>
					</select>
				</td>
			</tr>
			</table>
			<table width="100%" class="estilo_columnas">
			<thead>	
			<tr>
				<td colspan = "4">DOCUMENTOS NECESARIOS</td>
			</tr>
			</thead>
			<tr>
			<td colspan = "3" height="25">
				<%
				AbrirSCG1()

				if strSinGestionEsp= "1" Then
					intTamvDocGestion = -1
				Else
					intTamvDocGestion = ubound(vArrDocgestion)
				End If

				'Response.write "intTamvDocGestion=" & intTamvDocGestion
				'Response.End

				strSql = "SELECT * FROM TIPO_DOCUMENTOS_GESTION WHERE COD_CLIENTE = '" & strCodCliente & "'"
				set rsDoc=Conn1.execute(strSql)
				Do until rsDoc.eof
					strCheckDocGest = ""
					For I=0 To intTamvDocGestion
						If Trim(vArrDocgestion(I)) = Trim(rsDoc("NOM_DOCUMENTO")) Then
							strCheckDocGest = "CHECKED"
							exit for
						End If
					Next
				%>
					<INPUT TYPE="checkbox" NAME="CK_DOC_GESTION" ID="CK_DOC_GESTION" value="<%=rsDoc("NOM_DOCUMENTO")%>" <%=strCheckDocGest%>>
					&nbsp;<%=rsDoc("NOM_DOCUMENTO")%>

				<%
				rsDoc.movenext
				Loop
				CerrarSCG1()
				%>
				</td>
			</tr>
			<thead>
			<tr>			
				<td width="50%">MAS DOCUMENTOS NECESARIOS</td>
				<td width="17%" >
					FONO COBRO&nbsp;<abbr title="Actualizar información de fono de cobro"></abbr>
				</td>
				<td width="33%">CONTACTO COBRO</td>
			</tr>
			</thead>
			<tr>
				<td>
					<input NAME="TX_DOC_GESTION_NECESARIOS" type="text" id="TX_DOC_GESTION_NECESARIOS" size="44" maxlength="200">
				</td>
				<td id="td_CB_ID_FONO_COBRO">
					<select name="CB_ID_FONO_COBRO" id="CB_ID_FONO_COBRO" onchange="set_CB_ID_CONTACTO_FONO_COBRO(this.value);">
					<option value="">SELECCIONE</option>
					<%if fono_con="0" or fono_con="" then%>
					  <%
						AbrirSCG1()
						ssql_ = "SELECT ID_TELEFONO, TELEFONO, COD_AREA FROM DEUDOR_TELEFONO WHERE RUT_DEUDOR='" & strRutDeudor & "' AND ESTADO <> 2"
						set rsFON=Conn1.execute(ssql_)
						Do until rsFON.eof
							strFonoCB = rsFON("COD_AREA") & "-" & rsFON("Telefono")
							strSel=""
							if strFonoCB = strFonoAgend Then strSel = "SELECTED"
							%>
							<option value="<%=rsFON("ID_TELEFONO")%>" <%=strSel%>><%=rsFON("COD_AREA")%>-<%=rsFON("Telefono")%></option>
							<%
								rsFON.movenext
						Loop
						rsFON.close
						set rsFON=nothing
						CerrarSCG1()
					 %>
					<%else%>
						<option value="<%=fono_con%>"><%=area_con%>-<%=fono_con%></option>
					<%end if %>
					</select>
				</td>
				<td id="td_CB_ID_CONTACTO_FONO_COBRO">
					<select name="CB_ID_CONTACTO_FONO_COBRO" id="CB_ID_CONTACTO_FONO_COBRO">
					<option value="">SELECCIONE</option>
					</select>
				</td>
			</tr>
			</table>

		</div>


<%
	end if
%>
	<div name="divObsGestion" id="divObsGestion">

	<table width="100%" BORDER="0" class="estilo_columnas">
		<thead>
		<tr>
			<td width="680">OBSERVACIONES (Max. 600 Caract.)</td>
			<td >SCRIPT OBSERVACIÓN</td>
		</tr>
		</thead>
		<tr>
		   	<td align="LEFT">
				<TEXTAREA NAME="TX_OBSERVACIONES" placeholder="Ingresa observación" ID="TX_OBSERVACIONES" style="font-size:14px;" ROWS="4" COLS="100"></TEXTAREA>
		  	</td>
		  	<TD align="left">
		  		<%
		  			SQL_SCRIPT ="SELECT COD_CLIENTE, ORDEN, COD_GESTION, NOM_GESTION, SCRIPT_GESTION, RUTA_IMAGEN "
					SQL_SCRIPT = SQL_SCRIPT & " FROM SCRIPT_GESTION "
					SQL_SCRIPT = SQL_SCRIPT & " WHERE COD_CLIENTE =  " & TRIM(strCodCliente)
					set rs_script = conn.execute(SQL_SCRIPT)

		  		%>
				<table>
				<tr>
					<td align="center" width="80"> 
						<IMG style="cursor:pointer;" SRC="../Imagenes/48px-Crystal_Clear_mimetype_document2.png" width="20" height="20" onclick="bt_script_observacion('borrar')">
						<br>
						Borrar obs.
						&nbsp;
					</td>					
					<%if not rs_script.eof then
						do while not rs_script.eof%>
							<td align="center" width="80"> 
								<IMG style="cursor:pointer;" SRC="<%=TRIM(rs_script("RUTA_IMAGEN"))%>" width="25" height="25" onclick="bt_script_observacion('<%=rs_script("SCRIPT_GESTION")%>')">
								<br>
								<%=trim(rs_script("NOM_GESTION"))%>
								&nbsp;
							</td>
						<%rs_script.movenext
						loop
					end if%>
				</tr>
				</table>

		  	</TD>
		</tr>

	</table>

	</div>



<!-- ####***********************************************************************************#### -->
<!-- ####																					#### -->
<!-- ####  								AGENDAMIENTO										#### -->
<!-- ####  																					#### -->
<!-- ####***********************************************************************************#### -->

	<div name="divAgend" id="divAgend">
		<table width="100%">
			<tr>
				<td height="20" ALIGN=LEFT class="subtitulo_informe">
					> AGENDAMIENTO
				</td>
			</tr>
		</table>
		<table width="100%" class="estilo_columnas">
			<thead>
			<tr>
				<td>FECHA</td>
				<td >HORA</td>
				<%if trim(IndMedioAsociado)="1" then%>

					<td>FONO AGENDAMIENTO
						&nbsp;
						<abbr title="Actualizar información de fono de agendamiento">

						</abbr>
					</td>
					<td>FONO GESTION
						&nbsp;
						<abbr title="Actualizar información de fono de gestion">

						</abbr>
					</td>
					<td colspan="2">CONTACTO GESTIÓN</td>
	
				<%elseif trim(IndMedioAsociado)="2" then%>

					<td>EMAIL AGENDAMIENTO
						&nbsp;
						<abbr title="Actualizar información de fono de agendamiento">

						</abbr>
					</td>
					<td>EMAIL GESTION
						&nbsp;
						<abbr title="Actualizar información de fono de gestion">

						</abbr>
					</td>
					<td colspan="2">CONTACTO GESTIÓN</td>
	
				<%elseif trim(IndMedioAsociado)="3" then%>

					<td>DIRECCIÓN AGENDAMIENTO
						&nbsp;
						<abbr title="Actualizar información de fono de agendamiento">

						</abbr>
					</td>
					<td>DIRECCIÓN GESTION
						&nbsp;
						<abbr title="Actualizar información de fono de gestion">

						</abbr>
					</td>
					<td colspan="2">CONTACTO GESTIÓN</td>

				<%else%>
					<td colspan="6">&nbsp;</td>	

				<%end if%>
				
			</tr>
			</thead>
			<tr>
			  <td width="150">
				<input name="TX_FEC_AGEND" readonly type="text" id="TX_FEC_AGEND" size="10" maxlength="10" onBlur="ValidaDifFechas();">
			 </td>
			 <td width="100">
				<input name="TX_HORAAGEND" type="text" id="TX_HORAAGEND" size="5" maxlength="5" onChange="return ValidaHora(this,this.value)">
			 </td>

			<%if trim(IndMedioAsociado)<>"0" then%>
			
				<td id="td_ID_MEDIO_AGENDAMIENTO">
					<SELECT NAME="CB_ID_MEDIO_AGENDAMIENTO" id="CB_ID_MEDIO_AGENDAMIENTO">
						<OPTION VALUE="" >SELECCIONE</OPTION>
						<%AbrirSCG1()
						if trim(IndMedioAsociado)="1" then
						
							if fonoActual="" then 
								strEstadosCombo = "(0,1)"
							else 
								strEstadosCombo = "(0,1,2)"
							end if
							
							ssql_ = "SELECT ID_TELEFONO, TELEFONO,COD_AREA FROM DEUDOR_TELEFONO WHERE RUT_DEUDOR='" & strRutDeudor & "' AND ESTADO IN " & strEstadosCombo
							set rsFON=Conn1.execute(ssql_)
							Do until rsFON.eof
								if(rsFON("COD_AREA") = "0") then
									strFonoCB = rsFON("Telefono")
								else
									strFonoCB = rsFON("COD_AREA") & "-" & rsFON("Telefono")
								end if
								strSel=""
								if strFonoCB = strFonoAsociado Then strSel = "SELECTED"

                                if Replace(strFonoCB, "-", "") = fonoActual Then strSel = "SELECTED"
								%>
								<option value="<%=rsFON("ID_TELEFONO")%>" <%=strSel%>><%=strFonoCB%></option>
								<%
							rsFON.movenext
							Loop
							rsFON.close
							set rsFON=nothing
						
						elseif trim(IndMedioAsociado)="2" then
						
							ssql_ = "SELECT ID_EMAIL, UPPER(EMAIL) AS EMAIL FROM DEUDOR_EMAIL WHERE RUT_DEUDOR = '" & strRutDeudor & "' AND ESTADO <>2"
							set rsEmail=Conn1.execute(ssql_)
							Do until rsEmail.eof
								strEmailCB = rsEmail("EMAIL")
								strSel=""
								if strEmailCB = strEmailAgestionar Then strSel = "SELECTED"
								%>
								<option value="<%=rsEmail("ID_EMAIL")%>" <%=strSel%>><%=rsEmail("EMAIL")%></option>
								<%
									rsEmail.movenext
							Loop
							rsEmail.close
							set rsEmail=nothing

						elseif trim(IndMedioAsociado)="3" then

							strSql = "SELECT 1 as TIPO, REPLACE(CALLE + ' ' + NUMERO + ' ' + RESTO + ' ' + COMUNA,'  ',' ') as LUGAR_PAGO , 0 AS ORDEN, id_direccion ID , 'DIRECCION' tipo FROM DEUDOR_DIRECCION WHERE RUT_DEUDOR = '" & strRutDeudor & "' AND ESTADO <> 2"

							set rsDIR=Conn1.execute(strSql)
							do until rsDIR.eof
								direccion = rsDIR("LUGAR_PAGO")
								%>
								<option value="<%=rsDIR("ID")%>"
								<%if Trim(strLugarPago)=Trim(direccion) then
									response.Write("Selected")
								end if%>
								><%=direccion%></option>
								<%
								rsDIR.movenext
							loop
							rsDIR.close
							set rsDIR=nothing

						end if
						CerrarSCG1()%>
					</SELECT>
				</td>

				<td id="td_ID_MEDIO_GESTION">
					<select name="CB_ID_MEDIO_GESTION" id="CB_ID_MEDIO_GESTION"  onchange="set_CB_ID_CONTACTO_GESTION(<%=trim(IndMedioAsociado)%>, this.value)">	
					<option value="">SELECCIONE</option>
						<%AbrirSCG1()
						if trim(IndMedioAsociado)="1" then
						
							if fonoActual="" then 
								strEstadosCombo = "(0,1)"
							else 
								strEstadosCombo = "(0,1,2)"
							end if
							
							ssql_ = "SELECT ID_TELEFONO, TELEFONO,COD_AREA FROM DEUDOR_TELEFONO WHERE RUT_DEUDOR='" & strRutDeudor & "' AND ESTADO IN " & strEstadosCombo
							set rsFON=Conn1.execute(ssql_)
							Do until rsFON.eof
								if(rsFON("COD_AREA") = "0") then
									strFonoCB = rsFON("Telefono")
								else
									strFonoCB = rsFON("COD_AREA") & "-" & rsFON("Telefono")
								end if
								strSel=""
								if strFonoCB = strFonoAsociado Then strSel = "SELECTED"

                                if Replace(strFonoCB, "-", "") = fonoActual Then strSel = "SELECTED"
								%>
								<option value="<%=rsFON("ID_TELEFONO")%>" <%=strSel%>><%=strFonoCB%></option>
								<%
							rsFON.movenext
							Loop
							rsFON.close
							set rsFON=nothing
						
						elseif trim(IndMedioAsociado)="2" then

							ssql_ = "SELECT ID_EMAIL, UPPER(EMAIL) AS EMAIL FROM DEUDOR_EMAIL WHERE RUT_DEUDOR = '" & strRutDeudor & "' AND ESTADO <> 2"
							set rsEmail=Conn1.execute(ssql_)
							Do until rsEmail.eof
								strEmailCB = rsEmail("EMAIL")
								strSel=""
								if strEmailCB = strEmailAgestionar Then strSel = "SELECTED"
								%>
								<option value="<%=rsEmail("ID_EMAIL")%>" <%=strSel%>><%=rsEmail("EMAIL")%></option>
								<%
									rsEmail.movenext
							Loop
							rsEmail.close
							set rsEmail=nothing

						elseif trim(IndMedioAsociado)="3" then

							strSql = "SELECT 1 as TIPO, REPLACE(CALLE + ' ' + NUMERO + ' ' + RESTO + ' ' + COMUNA,'  ',' ') as LUGAR_PAGO , 0 AS ORDEN, id_direccion ID , 'DIRECCION' tipo FROM DEUDOR_DIRECCION WHERE RUT_DEUDOR = '" & strRutDeudor & "' AND ESTADO <> 2"

							set rsDIR=Conn1.execute(strSql)
							do until rsDIR.eof
								direccion = rsDIR("LUGAR_PAGO")
								%>
								<option value="<%=rsDIR("ID")%>"
								<%if Trim(strLugarPago)=Trim(direccion) then
									response.Write("Selected")
								end if%>
								><%=direccion%></option>
								<%
								rsDIR.movenext
							loop
							rsDIR.close
							set rsDIR=nothing

						end if
						CerrarSCG1()%>
					</select>
					<script>
						$("#CB_ID_MEDIO_GESTION").change()
					</script>
				</td>
				<td id="td_CB_ID_CONTACTO_GESTION">
					<select name="CB_ID_CONTACTO_GESTION" id="CB_ID_CONTACTO_GESTION">
						<option value="">SELECCIONE</option>
					</select>
				</td>

			<%else%>
				<td colspan="4"></td>				
			<%end if%>

			<td align="right">
				<input name="ingresar" id="ingresar" class="fondo_boton_130" type="button" onClick="ingreso_nueva_gestion();" value="Ingresar Gestión">
			</td>

		</table>
	</div>

	<input type="HIDDEN" name="TIPO_GESTION" 	id="TIPO_GESTION" 	value="<%=trim(strTipoGestion)%>">
	<input type="HIDDEN" name="MEDIO_ASOCIADO" 	id="MEDIO_ASOCIADO" value="<%=trim(IndMedioAsociado)%>">
	<input type="HIDDEN" name="strArrFeriados" 	id="strArrFeriados" 	value="<%=strArrFeriados%>">
	<input type="HIDDEN" name="Obligatoriedad" 	id="Obligatoriedad" 	value="<%=Obligatoriedad%>">

	<SCRIPT TYPE="text/javascript">

		var RangeDates = [<%=strArrFeriados%>];

		var RangeDatesIsDisable = true;
		function DisableDays(date) {
		    var isd = RangeDatesIsDisable;
		    var rd = RangeDates;
		    var d = date.getDate();
		    var m = date.getMonth();
		    var y = date.getFullYear();
		    for (i = 0; i < rd.length; i++) {
		        var ds = rd[i].split(',');
		        var di, df;
		        var m1, d1, y1, m2, d2, y2;

		        if (ds.length == 1) {
		            di = ds[0].split('/');
		            d1 = parseInt(di[0]);
		            m1 = parseInt(di[1]);
		            y1 = parseInt(di[2]);
		            if (y1 == y && m1 == (m + 1) && d1 == d) return [!isd];
		            
		        } else if (ds.length > 1) {
		            di = ds[0].split('/');
		            df = ds[1].split('/');
		            d1 = parseInt(di[0]);
		            m1 = parseInt(di[1]);
		            y1 = parseInt(di[2]);
		            d2 = parseInt(df[0]);
		            m2 = parseInt(df[1]);
		            y2 = parseInt(df[2]);

		            if (y1 >= y || y <= y2) {
		                
		                if ((m + 1) >= m1 && (m + 1) <= m2) {
		                    if (m1 == m2) {
		                        if (d >= d1 && d <= d2) return [!isd];
		                    } else if (m1 == (m + 1)) {
		                        if (d >= d1) return [!isd];
		                    } else if (m2 == (m + 1)) {
		                        if (d <= d2) return [!isd];
		                    } else return [!isd];
		                }
		            }

		        }
		    }
		    return [isd];
		}



		function ValidaDifFechas() {
			var TX_FECHA_COMPROMISO 	=$('#TX_FECHA_COMPROMISO').val()
			var TX_FEC_AGEND 			=$('#TX_FEC_AGEND').val()

			if(TX_FEC_AGEND != '' && TX_FEC_AGEND != null && TX_FECHA_COMPROMISO != '' && TX_FECHA_COMPROMISO != null){
				var diferencia=DiferenciaEntreFechas(TX_FECHA_COMPROMISO,TX_FEC_AGEND)

				if (diferencia >= 0) {

				}else{
					alert('La fecha de agendamiento debe ser menor o igual a la fecha de compromiso de pago.')
					$('#TX_FEC_AGEND').val("")
				}
			}
		}


		function DiferenciaEntreFechas (CadenaFecha1, CadenaFecha2) {
		   fecha_hoy = getCurrentDate() //hoy

		   //Obtiene dia, mes y año
		   var fecha1 = new fecha( CadenaFecha1 )
		   var fecha2 = new fecha( CadenaFecha2)

		   //Obtiene objetos Date
		   var miFecha1 = new Date( fecha1.anio, fecha1.mes, fecha1.dia )
		   var miFecha2 = new Date( fecha2.anio, fecha2.mes, fecha2.dia )

		   //Resta fechas y redondea
		   var diferencia = miFecha1.getTime() - miFecha2.getTime()
		   var dias = Math.floor(diferencia / (1000 * 60 * 60 * 24))
		   var segundos = Math.floor(diferencia / 1000)
		   //alert ('La diferencia es de ' + dias + ' dias,\no ' + segundos + ' segundos.')

		   return dias //false
		}



		//---------------------------------------------------------------------
		function fecha( cadena ) {
		   //Separador para la introduccion de las fechas
		   var separador = "/"
		   //Separa por dia, mes y año
			   if ( cadena.indexOf( separador ) != -1 ) {
			        var POSI_1 = 0
			        var POSI_2 = cadena.indexOf( separador, POSI_1 + 1 )
			        var POSI_3 = cadena.indexOf( separador, POSI_2 + 1 )
			        this.dia = cadena.substring( POSI_1, POSI_2 )
			        this.mes = cadena.substring( POSI_2 + 1, POSI_3 )
			        this.anio = cadena.substring( POSI_3 + 1, cadena.length )
			   } else {
			        this.dia = 0
			        this.mes = 0
			        this.anio = 0
			   }
		}
			
function CB_ID_DIRECCION_COBRO_DEUDOR_onclick() {

}

    </SCRIPT>

<%elseif trim(accion_ajax)="ingreso_gestion" then


	
	strRutDeudor				=request.querystring("rut")	
	strCodCliente 				=request.querystring("strCodCliente")
	strGestionTipoGestion 		=request.querystring("strGestionTipoGestion")
	intIdCampana 				=request.querystring("intIdCampana")
	IntTipoGestion				=request.querystring("IntTipoGestion")
	strCuotasDeudor				=request.querystring("cuotas_deudor")
	strObservaciones			=request.querystring("strObservaciones")

	dtmFechaCompromiso 			=request.querystring("dtmFechaCompromiso")
	strNroDocPago 				=request.querystring("strNroDocPago")
	dtmFechaPago				=request.querystring("dtmFechaPago")
	strHoraDesde				=request.querystring("strHoraDesde")
	strHoraHasta				=request.querystring("strHoraHasta")
	strFormaPago				=request.querystring("strFormaPago")
	strDocGestion 				=request.querystring("strDocGestion") 
	strDocGestionNecesarios		=request.querystring("strDocGestionNecesarios")
	IntMontoCancelado 			=request.querystring("IntMontoCancelado")
	strEnvioHdr					=request.querystring("strEnvioHdr")
	intIdFonoCobro				=request.querystring("intIdFonoCobro")
	intIdContactoFonoCobro 		=request.querystring("intIdContactoFonoCobro")
	intIdDireccionCobroDeudor 	=request.querystring("intIdDireccionCobroDeudor")

	dtmFecAgend 				=request.querystring("dtmFecAgend")
	strHoraAgend				=request.querystring("strHoraAgend")
	intIdMedioAgendamiento		=request.querystring("intIdMedioAgendamiento")
	intIdMedioGestion			=request.querystring("intIdMedioGestion")
	intIdContactoGestion		=request.querystring("intIdContactoGestion")

	intMedioAsociado			=request.querystring("intMedioAsociado")
	intIdSegmentoVenc			=request.querystring("intIdSegmentoVenc")
	intIdSegmentoMonto			=request.querystring("intIdSegmentoMonto")
	intIdSegmentoAsig			=request.querystring("intIdSegmentoAsig")
	
	'response.write "ID_DIRECCION_COBRO_DEUDOR : "& ID_DIRECCION_COBRO_DEUDOR &"<br>"
	'response.end

	if trim(strObservaciones)="undefined" OR trim(strObservaciones)="" Then
		strObservaciones 	="NULL"
	else
		strObservaciones 	="'"&strObservaciones&"'"
	end if

	if trim(dtmFechaCompromiso)="undefined" OR trim(dtmFechaCompromiso)="" Then
		dtmFechaCompromiso 	="NULL"
	else
		dtmFechaCompromiso 	="'"&dtmFechaCompromiso&"'"
	end if

	if trim(strNroDocPago)="undefined" OR trim(strNroDocPago)="" Then
		strNroDocPago 	="NULL"
	else
		strNroDocPago 	="'"&strNroDocPago&"'"
	end if

	if trim(dtmFechaPago)="undefined" OR trim(dtmFechaPago)="" Then
		dtmFechaPago 	="NULL"
	else
		dtmFechaPago 	="'"&dtmFechaPago&"'"
	end if

	if trim(strHoraDesde)="undefined" OR trim(strHoraDesde)="" Then
		strHoraDesde 	="NULL"
	else
		strHoraDesde 	="'"&strHoraDesde&"'"
	end if

	if trim(strHoraHasta)="undefined" OR trim(strHoraHasta)="" Then
		strHoraHasta 	="NULL"
	else
		strHoraHasta 	="'"&strHoraHasta&"'"
	end if

	if trim(strFormaPago)="undefined" OR trim(strFormaPago)="" Then
		strFormaPago 	="NULL"
	else
		strFormaPago 	="'"&strFormaPago&"'"
	end if
	
	if trim(strDocGestion)="undefined" OR trim(strDocGestion)="" Then

		if trim(strDocGestionNecesarios)="undefined" OR trim(strDocGestionNecesarios)="" Then
			strDocGestion ="NULL"	
		
		else
			strDocGestion = "'"&strDocGestionNecesarios&"......'"		

		end if

	else

		if trim(strDocGestionNecesarios)<>"undefined" OR trim(strDocGestionNecesarios)<>"" Then
			strDocGestion = "'"&strDocGestion &" / "&strDocGestionNecesarios&"'"
		
		else
			strDocGestion ="'"&strDocGestion&"'"		

		end if

	end if


	if trim(IntMontoCancelado)="undefined" OR trim(IntMontoCancelado)="" Then
		IntMontoCancelado 	="0"
	end if

	if trim(strEnvioHdr)="undefined" OR trim(strEnvioHdr)="" Then
		strEnvioHdr 	="NULL"
	end if

	if trim(intIdFonoCobro)="undefined" OR trim(intIdFonoCobro)="" Then
		intIdFonoCobro 	="NULL"
	end if

	if trim(intIdContactoFonoCobro)="undefined" OR trim(intIdContactoFonoCobro)="" Then
		intIdContactoFonoCobro 	="NULL"
	end if

	if trim(intIdDireccionCobroDeudor)="undefined" OR trim(intIdDireccionCobroDeudor)="" Then
		intIdDireccionCobroDeudor 	="NULL"
		intFormaRecaudacion			="NULL"
	else

		concat_intIdDireccionCobroDeudor 	= split(intIdDireccionCobroDeudor,"-")
		ID		=concat_intIdDireccionCobroDeudor(0)
		TIPO 	=concat_intIdDireccionCobroDeudor(1)

		IF trim(TIPO)="FORMA_RECAUDACION" Then
			intIdDireccionCobroDeudor 	="NULL"
			intFormaRecaudacion 		=ID
		ELSE
			intIdDireccionCobroDeudor 	=ID
			intFormaRecaudacion 		="NULL"
		END IF


	end if

	if trim(dtmFecAgend)="undefined" OR trim(dtmFecAgend)="" Then
		dtmFecAgend 	="NULL"
	else
		dtmFecAgend 	="'"&dtmFecAgend&"'"
	end if

	if trim(strHoraAgend)="undefined" OR trim(strHoraAgend)="" Then
		strHoraAgend 	="NULL"
	else
		strHoraAgend 	="'"&strHoraAgend&"'"
	end if

	if trim(intIdMedioAgendamiento)="undefined" OR trim(intIdMedioAgendamiento)="" Then
		intIdMedioAgendamiento 	="NULL"
	end if

	if trim(intIdMedioGestion)="undefined" OR trim(intIdMedioGestion)="" Then
		intIdMedioGestion 	="NULL"
	end if

	if trim(intIdContactoGestion)="undefined" OR trim(intIdContactoGestion)="" Then
		intIdContactoGestion 	="NULL"
	end if
	
	if fonoActual <> "" and pagina_origen <> "agendamiento_tactico" then 
		intProvieneDiscador = 1
	else intProvieneDiscador = "NULL"
	end if	

 	if trim(intIdSegmentoVenc)="undefined" OR trim(intIdSegmentoVenc)="" Then
		intIdSegmentoVenc 	="NULL"
	end if

 	if trim(intIdSegmentoMonto)="undefined" OR trim(intIdSegmentoMonto)="" Then
		intIdSegmentoMonto 	="NULL"
	end if

 	if trim(intIdSegmentoAsig)="undefined" OR trim(intIdSegmentoAsig)="" Then
		intIdSegmentoAsig 	="NULL"
	end if
	
        dim IntMontoCanceladoFormat  
        IntMontoCanceladoFormat  = FormatNumber(IntMontoCancelado,0)

	sql_insert_cuota ="EXEC proc_Ingreso_Gestion '" & strRutDeudor & "',"
	sql_insert_cuota = sql_insert_cuota &	"'" & strCodCliente & "',"
	sql_insert_cuota = sql_insert_cuota &	"'" & strGestionTipoGestion & "',"
	sql_insert_cuota = sql_insert_cuota & 	"'" & session("session_idusuario") & "',"
	sql_insert_cuota = sql_insert_cuota &	dtmFechaCompromiso & ","
	sql_insert_cuota = sql_insert_cuota &	strNroDocPago & ","
	sql_insert_cuota = sql_insert_cuota &   dtmFechaPago & ","
	sql_insert_cuota = sql_insert_cuota &   UCASE(strObservaciones) & ","
	sql_insert_cuota = sql_insert_cuota &   dtmFecAgend & ","
	sql_insert_cuota = sql_insert_cuota &  	strHoraAgend & ","

	sql_insert_cuota = sql_insert_cuota & 	intIdCampana & ","
	sql_insert_cuota = sql_insert_cuota & 	strFormaPago & ","
	sql_insert_cuota = sql_insert_cuota & 	strHoraDesde  & ","
	sql_insert_cuota = sql_insert_cuota & 	strHoraHasta  & ","
	sql_insert_cuota = sql_insert_cuota & 	strDocGestion & ","
	sql_insert_cuota = sql_insert_cuota & 	IntMontoCancelado & "," 

	sql_insert_cuota = sql_insert_cuota & 	strEnvioHdr & ","
	sql_insert_cuota = sql_insert_cuota & 	intIdFonoCobro & ","
	sql_insert_cuota = sql_insert_cuota & 	intIdContactoFonoCobro & ","
	sql_insert_cuota = sql_insert_cuota & 	trim(intIdMedioGestion) & ","
	sql_insert_cuota = sql_insert_cuota & 	trim(intIdContactoGestion)&","
	sql_insert_cuota = sql_insert_cuota & 	trim(intIdDireccionCobroDeudor)&","
	sql_insert_cuota = sql_insert_cuota & 	TRIM(intFormaRecaudacion)&","
	sql_insert_cuota = sql_insert_cuota & 	TRIM(intIdMedioAgendamiento)&","
	sql_insert_cuota = sql_insert_cuota & 	"'" & strCuotasDeudor &"',"
	sql_insert_cuota = sql_insert_cuota & 	intMedioAsociado & ","
	sql_insert_cuota = sql_insert_cuota & 	"'" & IntMontoCanceladoFormat &"',"
	sql_insert_cuota = sql_insert_cuota & 	intProvieneDiscador & ","
	sql_insert_cuota = sql_insert_cuota & 	intIdSegmentoVenc & ","
	sql_insert_cuota = sql_insert_cuota & 	intIdSegmentoMonto & ","
	sql_insert_cuota = sql_insert_cuota & 	intIdSegmentoAsig
	
	'response.write sql_insert_cuota
	'response.end()

	set rs_insert_cuota = conn.execute(sql_insert_cuota)
	if not rs_insert_cuota. eof then
		intIdGestion = rs_insert_cuota("ID_GESTION")

	else
		intIdGestion =0
	end if

'**/Consulta la prioridad calculada de la cuota/**'


	strSql = "SELECT GESTION_MODULOS, PRIORIDAD FROM GESTIONES_TIPO_GESTION "
	strSql= strSql & " WHERE COD_CLIENTE = " & strCodCliente & " AND COD_CATEGORIA  = CONVERT(int, SUBSTRING('"&strGestionTipoGestion&"',1,1)) AND COD_SUB_CATEGORIA  = CONVERT(int, SUBSTRING('"&strGestionTipoGestion&"',3,1)) AND COD_GESTION  = CONVERT(int, SUBSTRING('"&strGestionTipoGestion&"',5,1)) "
	'response.write strSql&"<br>"
	set rsGesTipoGes = Conn.execute(strSql)
	If Not rsGesTipoGes.eof Then
		intGestionModulos 	= rsGesTipoGes("GESTION_MODULOS")
		intPrioridadGestion = Cdbl(rsGesTipoGes("PRIORIDAD"))
	Else
		intGestionModulos 	= 0
		intPrioridadGestion = 0
	End If


	strSql = "SELECT CAST(ISNULL(PRIORIDAD_CUOTA,11) AS NUMERIC(4,1)) AS PRIORIDAD_CUOTA, PRIORIDAD_CUOTA_CAL = CASE"
	strSql = strSql & "    WHEN [dbo].[fun_FonosDias] (RUT_DEUDOR,2) >= 1 AND [dbo].[fun_dias_atencion_telefonica] (RUT_DEUDOR, GETDATE(),0)>=1 AND CUOTA.COD_CLIENTE = 1100"
	strSql = strSql & "    THEN 4"
	strSql = strSql & "    WHEN CAST(GETDATE() - FECHA_VENC  as int) >= 30 AND CUOTA.COD_CLIENTE = 1100"
	strSql = strSql & "    THEN 5"
	strSql = strSql & "    WHEN CAST(GETDATE() - FECHA_VENC  as int) >= 10 AND CUOTA.COD_CLIENTE = 1100"
	strSql = strSql & "    THEN 6"
	strSql = strSql & "    WHEN SALDO >=100000000 AND CUOTA.COD_CLIENTE = 1100"
	strSql = strSql & "    THEN 7"
	strSql = strSql & "    WHEN CAST(GETDATE() - FECHA_VENC  as int) >= 0 AND CUOTA.COD_CLIENTE = 1100"
	strSql = strSql & "    THEN 9"
	strSql = strSql & "    WHEN CUOTA.ID_ULT_GEST IS NULL AND CUOTA.COD_CLIENTE = 1100"
	strSql = strSql & "    THEN 10"
	strSql = strSql & "    WHEN CUOTA.COD_CLIENTE = 1100"
	strSql = strSql & "    THEN 11"
	strSql = strSql & "    ELSE 100"
	strSql = strSql & " END"
	strSql = strSql & " FROM dbo.CUOTA INNER JOIN dbo.ESTADO_DEUDA ON dbo.ESTADO_DEUDA.CODIGO = cuota.ESTADO_DEUDA"
	strSql = strSql & " WHERE ESTADO_DEUDA.ACTIVO = 1 AND ID_CUOTA IN (" & strCuotasDeudor & ")"

	'Response.write strSql &"<br>"

	set rsTmp = Conn.execute(strSql)
	If Not rsTmp.eof Then
		intPrioridadCal 	= rsTmp("PRIORIDAD_CUOTA_CAL")
		intPrioridadCuota 	= Cdbl(rsTmp("PRIORIDAD_CUOTA"))
	End if



	If (intPrioridadCal <= intPrioridadGestion AND intPrioridadGestion <> 8) or (intPrioridadCuota >= 8 AND intPrioridadGestion = 8) Then
		intPrioridadFinal = intPrioridadCal
	ElseIf intPrioridadCuota < 8 AND intPrioridadGestion = 8 Then
	   	intPrioridadFinal = intPrioridadCuota
	Else
	   	intPrioridadFinal = intPrioridadGestion
	End If

	If ((intPrioridadCuota  > 5) and UCASE(Request("CH_PRIORITARIA")) = "ON") Then
		intPrioridadFinal = "2.2"
	End If



	strSql_p = "UPDATE CUOTA SET PRIORIDAD_CUOTA = " & Replace(intPrioridadFinal,",",".")
	strSql_p = strSql_p & " WHERE ID_CUOTA in ("& strCuotasDeudor & ")"
	'Response.write "<br>" & strSql
	Conn.execute(strSql_p)


'**/Redirige según variable a distintas partes del sistema al ingresar gestión/**'

	strSql = " SELECT DIRECCION_RETIRO = (CASE WHEN (ID_DIRECCION_COBRO_DEUDOR IS NOT NULL OR FR.TIPO = 'RETIRO') THEN 1 ELSE 0 END),"
	strSql = strSql & " CONFIRMA_CP = ISNULL(CONFIRMA_CP,0)"

	strSql = strSql & " FROM GESTIONES G	INNER JOIN GESTIONES_TIPO_GESTION GTG ON G.COD_CATEGORIA = GTG.COD_CATEGORIA "
	strSql = strSql & " 							   								 AND G.COD_SUB_CATEGORIA = GTG.COD_SUB_CATEGORIA "
	strSql = strSql & " 							   								 AND G.COD_GESTION = GTG.COD_GESTION "
	strSql = strSql & " 							   								 AND G.COD_CLIENTE = GTG.COD_CLIENTE "
	strSql = strSql & " 					LEFT JOIN FORMA_RECAUDACION FR ON G.ID_FORMA_RECAUDACION = FR.ID_FORMA_RECAUDACION "
	
	strSql = strSql & " WHERE ID_GESTION = " & intIdGestion
	'Response.write "<br><br>strSql=" & strSql

	set rsInf = Conn.execute(strSql )
	
	DireccionRetiro = rsInf("DIRECCION_RETIRO")
	GestionConfirmaRuta = rsInf("CONFIRMA_CP")
	
	'Response.write "<br><br>DireccionRetiro=" & DireccionRetiro
	'response.end

	strSql = "SELECT TOTAL_CUOTAS = COUNT(ID_CUOTA)"
	strSql = strSql & " FROM CUOTA C INNER JOIN ESTADO_DEUDA ED ON C.ESTADO_DEUDA = ED.CODIGO"
	strSql = strSql & " WHERE C.RUT_DEUDOR='" & strRutDeudor & "' AND C.COD_CLIENTE='" & strCodCliente & "'"
	strSql = strSql & " AND ED.ACTIVO = 1"
	strSql = strSql & " AND COD_ULT_GEST IN (SELECT cast(COD_CATEGORIA as varchar(2))+ '*' + cast(COD_SUB_CATEGORIA as varchar(2)) + '*' + cast(COD_GESTION as varchar(2))"
	strSql = strSql & " FROM GESTIONES_TIPO_GESTION WHERE VER_AGEND = 1 AND COD_CLIENTE ='" & strCodCliente & "')"
	strSql = strSql & " AND DATEDIFF(MINUTE,(ISNULL(FECHA_AGEND_ULT_GES,GETDATE()-200) + convert(varchar(10),ISNULL(HORA_AGEND_ULT_GES,'00:00'),108)),GETDATE()) >= 0"
	
	'Response.write "<br><br>strSql=" & strSql
	set rsValida = Conn.execute(strSql )
	
	TotalCuotasPendienteGestion = rsValida("TOTAL_CUOTAS")

	If Trim(DireccionRetiro) = "1" and Trim(GestionConfirmaRuta) = "1" and Trim(dtmFechaCompromiso) <> "" and Trim(dtmFechaCompromiso) <> "01/01/1900" and Trim(dtmFechaCompromiso) <> "NULL" Then
	
		pagina_redireccionamiento ="confirmar_cp"
		
	ElseIf TotalCuotasPendienteGestion = 0 then

		pagina_redireccionamiento ="principal"

	End If

    'response.write TotalCuotasPendienteGestion
	'Response.Redirect strArchivoAsp & "&rut=" & strRutDeudor & "&cliente=" & strCodCliente

	'response.write strArchivoAsp&"<br>"

%>
	<input type="hidden" name="pagina_redireccionamiento" id="pagina_redireccionamiento" value="<%=pagina_redireccionamiento%>">
	<input type="hidden" name="intIdGestion" id="intIdGestion" value="<%=intIdGestion%>">
	

<%elseif trim(accion_ajax)="agendamiento_gestion_sin_contacto" then
	
	strRutDeudor				=request.querystring("rut")	
	strCodCliente      		    =request.querystring("strCodCliente")
	dtmFecAgend 				=request.querystring("dtmFecAgend")
	strHoraAgend				=request.querystring("strHoraAgend")

	if trim(dtmFecAgend)="undefined" OR trim(dtmFecAgend)="" Then
		dtmFecAgend 	="NULL"
	else
		dtmFecAgend 	="'"&dtmFecAgend&"'"
	end if

	if trim(strHoraAgend)="undefined" OR trim(strHoraAgend)="" Then
		strHoraAgend 	="NULL"
	else
		strHoraAgend 	="'"&strHoraAgend&"'"
	end if

	sql_insert_gestion_sc ="EXEC proc_Ingreso_Gestion_Sin_Contacto '" & strRutDeudor & "',"
	sql_insert_gestion_sc = sql_insert_gestion_sc &	"'" & strCodCliente & "',"
	sql_insert_gestion_sc = sql_insert_gestion_sc &   dtmFecAgend & ","
	sql_insert_gestion_sc = sql_insert_gestion_sc &   strHoraAgend
	
	'response.write sql_insert_gestion_sc
	'response.end()

	set sql_insert_gestion_sc = conn.execute(sql_insert_gestion_sc)

elseif trim(accion_ajax)="refresca_historial" then


	strRutDeudor	= request.querystring("rut")
	strCodCliente   = request.querystring("strCodCliente")
	inicio 			= request.querystring("inicio")
	finales 		= request.querystring("finales")
	strFiltro 		= request.querystring("CB_FILTRO")
    id_usuario      = session("session_idusuario")'request.querystring("id_usuario")

    strSql="select EsInterno,EsExterno,PuedenEscucharMisGrabaciones,PuedoEscucharGrabaciones from usuario where id_usuario = " & id_usuario

    'response.write("id_usuario--" & id_usuario)

    set RsUser=Conn.execute(strSql)

    			If not RsUser.eof then
	    Usr_EsInterno                       = RsUser("EsInterno")    
        Usr_EsExterno                       = RsUser("EsExterno")    
        Usr_PuedenEscucharMisGrabaciones    = RsUser("PuedenEscucharMisGrabaciones")    
        Usr_PuedoEscucharGrabaciones        = RsUser("PuedoEscucharGrabaciones")   

	End if
    
    RsUser.close
	set RsUser=nothing	

         


    strSql=""
	strSql="SELECT MAX( CAST((CONVERT(VARCHAR(10),G.FECHA_INGRESO,103)+' '+G.HORA_INGRESO) AS DATETIME)) AS MAX_FECHA_GES_TIT"
	strSql=strSql + " FROM GESTIONES G INNER JOIN GESTIONES_TIPO_CATEGORIA GTC ON G.COD_CATEGORIA = GTC.COD_CATEGORIA "
	strSql=strSql + " 				 INNER JOIN GESTIONES_TIPO_SUBCATEGORIA GTSC ON G.COD_CATEGORIA = GTSC.COD_CATEGORIA" 
	strSql=strSql + " 																   AND G.COD_SUB_CATEGORIA = GTSC.COD_SUB_CATEGORIA "
	strSql=strSql + " 				 INNER JOIN GESTIONES_TIPO_GESTION GTG ON G.COD_CATEGORIA = GTG.COD_CATEGORIA "
	strSql=strSql + " 														   AND G.COD_SUB_CATEGORIA = GTG.COD_SUB_CATEGORIA"
	strSql=strSql + " 														   AND G.COD_GESTION = GTG.COD_GESTION"
	strSql=strSql + " 														   AND GTG.COD_CLIENTE = '" & strCodCliente & "'"
	strSql=strSql + " 				 INNER JOIN GESTIONES_CUOTA GC ON G.ID_GESTION = GC.ID_GESTION "
	strSql=strSql + " 				 INNER JOIN CUOTA C ON C.ID_CUOTA = GC.ID_CUOTA AND GC.ID_GESTION = G.ID_GESTION "
	strSql=strSql + " 										AND C.COD_CLIENTE = '" & strCodCliente & "'"
	strSql=strSql + " 				 INNER JOIN ESTADO_DEUDA ED ON C.ESTADO_DEUDA=ED.CODIGO"

	strSql=strSql + " WHERE G.COD_CLIENTE = '" & strCodCliente & "' AND G.RUT_DEUDOR = '" & strRutDeudor & "' AND ACTIVO=1 AND ISNULL(GTG.PRIORIDAD_GTIT,0)=1"
	

	set RsFec=Conn.execute(strSql)
	If not RsFec.eof then
		dtmMaxFecTitular = RsFec("MAX_FECHA_GES_TIT")
	End if
	RsFec.close
	set RsFec=nothing		
	


	strSql = "SELECT * FROM ( "
	strSql = strSql & "SELECT row_number() over (order by CAST(CONVERT(VARCHAR(10), FECHA_INGRESO_GESTION,103)+' '+CONVERT(VARCHAR(5),convert(datetime, HORA_INGRESO), 108) AS DATETIME) DESC , ID_GESTION DESC) as numero_fila "
	strSql = strSql & " ,COD_CLIENTE "
	strSql = strSql & " ,ID_GESTION "
	strSql = strSql & " ,FECHA_INGRESO_GESTION "
	strSql = strSql & " ,RUT_DEUDOR "
	strSql = strSql & " ,NOMBRE_DEUDOR "
	strSql = strSql & " ,TOTAL_DOC "
	strSql = strSql & " ,TOTAL_DOC_ACTIVO "
	strSql = strSql & " ,DM "
	strSql = strSql & " ,FECHA_GESTION "
	strSql = strSql & " ,HORA_INGRESO "
	strSql = strSql & " ,FECHA_COMPROMISO "
	strSql = strSql & " ,FECHA_AGENDAMIENTO "
	strSql = strSql & " ,HORA_AGENDAMIENTO "
	strSql = strSql & " ,MONTO_GESTION "
	strSql = strSql & " ,FORMA_NORMALIZACION "
	strSql = strSql & " ,LUGAR_GESTION "
	strSql = strSql & " ,NRO_DOC_PAGO "
	strSql = strSql & " ,OBSERVACIONES_CAMPO "
	strSql = strSql & " ,OBSERVACIONES "
	strSql = strSql & " ,NOM_MEDIO_GESTION "
	strSql = strSql & " ,NOM_MEDIO_AGEND "
	strSql = strSql & " ,NOM_CONTACTO_GESTION "
	strSql = strSql & " ,MEDIO_ASOCIADO "
	strSql = strSql & " ,ES_GESTION_CALL "
	strSql = strSql & " ,TIPO_MEDIO_GESTION "
	strSql = strSql & " ,ES_G_EFE "
	strSql = strSql & " ,ES_G_TIT "
	strSql = strSql & " ,ES_G_NC "
	strSql = strSql & " ,IND "
	strSql = strSql & " ,Gestion_Modulos "
	strSql = strSql & " ,COD_CATEGORIA "
	strSql = strSql & " ,COD_SUB_CATEGORIA "
	strSql = strSql & " ,COD_GESTION "
	strSql = strSql & " ,LOGIN "
	strSql = strSql & " ,COMUNICA "
	strSql = strSql & " ,PRIORIDAD_GTEL "
	strSql = strSql & " ,PRIORIDAD_GMAIL "
	strSql = strSql & " ,PRIORIDAD_GDIR "	
	strSql = strSql & " ,DESCRIPCION "
	strSql = strSql & " ,ANEXO "
	strSql = strSql & " ,CUOTAS_ACTIVAS "
	strSql = strSql & " ,CUOTAS_CANCELADAS "
	strSql = strSql & " ,CUOTAS_RETIRADAS "
	strSql = strSql & " ,CUOTAS_NO_ASIGNABLES "			
	strSql = strSql & " ,DOC_A_CONFIRMAR, id_usuario ,EsInterno,EsExterno,PuedenEscucharMisGrabaciones,PuedoEscucharGrabaciones,Ubicabilidad "			
    strSql = strSql & " ,DATENAME(weekday,FECHA_INGRESO_GESTION) + ', ' + SUBSTRING(convert(varchar,CONVERT(datetime,FECHA_INGRESO_GESTION),105),1,2) + ' de ' + DATENAME(month,FECHA_INGRESO_GESTION)    Fecha_Tootltip"
	strSql = strSql & " FROM VIEW_HISTORIAL_GESTIONES VV "
	strSql = strSql & " WHERE VV.COD_CLIENTE = '"&TRIM(strCodCliente)&"' AND VV.RUT_DEUDOR = '"&TRIM(strRutDeudor)&"' " 	

	if Trim(strFiltro) = "EFECTIVAS ACTIVAS" Then 'Todas las vigente menos las no comunica inferiores a la ultima fecha de gestion efectiva.
		strSql=strSql & " AND NOT (VV.PRIORIDAD_GTEL = '1' AND VV.PRIORIDAD_GEFE = '0' AND CAST((CONVERT(VARCHAR(10),VV.FECHA_INGRESO_GESTION,103)+' '+CONVERT(VARCHAR(5),convert(datetime,VV.HORA_INGRESO), 108) ) AS DATETIME) < CAST('" & dtmMaxFecTitular & "' AS DATETIME))"
		strSql=strSql & " AND NOT (VV.PRIORIDAD_GTEL = '1' AND VV.PRIORIDAD_GEFE = '0' AND CAST((CONVERT(VARCHAR(10),VV.FECHA_INGRESO_GESTION,103)+' '+VV.HORA_INGRESO) AS DATETIME) < CAST('" & dtmMaxFecTitular & "' AS DATETIME))"
		strSql=strSql & " AND VV.TOTAL_DOC_ACTIVO >0 "
		strSql=strSql & " AND (VV.GESTION_MASIVA='N' OR VV.GESTION_MODULOS IN (26,27,28,32))"
	End If

	if Trim(strFiltro) = "EFECTIVAS" Then 'Todas las gestiones menos las no comunica
		strSql=strSql & " AND NOT (VV.PRIORIDAD_GTEL = '1' AND VV.PRIORIDAD_GEFE = '0')"
		strSql=strSql & " AND NOT (VV.PRIORIDAD_GTEL = '1' AND VV.PRIORIDAD_GEFE = '0') " 
		strSql=strSql & " AND (VV.GESTION_MASIVA='N' OR VV.GESTION_MODULOS IN (26,27,28,32))"
	End If

	if Trim(strFiltro) = "ACTIVAS CALL" Then 'Todas las gestiones menos las no comunica
		strSql=strSql & " AND VV.TOTAL_DOC_ACTIVO >0 AND VV.GESTION_MASIVA='N'"
	End If	
	
	if Trim(strFiltro) = "ACTIVAS MASIVO" Then 'Todas las gestiones menos las no comunica
		strSql=strSql & " AND VV.TOTAL_DOC_ACTIVO >0 AND VV.GESTION_MASIVA='S' AND VV.COD_SUB_CATEGORIA NOT IN (6,8)"
	End If	
	
	strSql = strSql & " ) a WHERE a.numero_fila BETWEEN "&inicio&" AND "&finales

    strSql = strSql & " ORDER BY CAST(CONVERT(VARCHAR(10), FECHA_INGRESO_GESTION,103)+' '+CONVERT(VARCHAR(5),convert(datetime, HORA_INGRESO), 108) AS DATETIME) DESC, ID_GESTION DESC"
	'strSql = strSql & " ORDER BY a.ID_GESTION DESC"

	'Response.write "<br>strSql==" & strSql
	'RESPONSE.End()

	set rsDET = conn.execute(strSql)
	%>
	<table class="" style="width:100%;" border="0" cellSpacing="0" cellPadding="0">
	<tbody>
	<%
	if not rsDET.eof then
	intCorr = 1
	do while not rsDET.eof
		numero_fila 			=rsDET("numero_fila")
		strCodCliente			=rsDET("COD_CLIENTE")
		intIdGestion			=rsDET("ID_GESTION")
		dtmFechaIngresoGestion	=rsDET("FECHA_INGRESO_GESTION")
		strRutDeudor			=rsDET("RUT_DEUDOR")
		strNombreDeudor			=rsDET("NOMBRE_DEUDOR")
		intTotalDoc				=rsDET("TOTAL_DOC")
		intTotalDocActivo		=rsDET("TOTAL_DOC_ACTIVO")
		intDiasMora				=rsDET("DM")
		dtmFechaGestion			=rsDET("FECHA_GESTION")
		strHoraIngreso			=rsDET("HORA_INGRESO")
		dtmFechaCompromiso		=rsDET("FECHA_COMPROMISO")
		dtmFechaAgendamiento	=rsDET("FECHA_AGENDAMIENTO")
		strHoraAgendamiento		=rsDET("HORA_AGENDAMIENTO")
		intMontoGestion			=rsDET("MONTO_GESTION")
		strFormaNormalizacion	=rsDET("FORMA_NORMALIZACION")
		strLugarGestion			=rsDET("LUGAR_GESTION")
		strNroDocPago			=rsDET("NRO_DOC_PAGO")
		strObservacionesCampo	=rsDET("OBSERVACIONES_CAMPO")
		strObsdervaciones		=rsDET("OBSERVACIONES")
		intNomMedioGestion		=rsDET("NOM_MEDIO_GESTION")
		strNomMedioAgend		=rsDET("NOM_MEDIO_AGEND")
		strNomContactoGestion	=rsDET("NOM_CONTACTO_GESTION")
		intMedioAsociado		=rsDET("MEDIO_ASOCIADO")
		intTipoMedioGestion		=rsDET("TIPO_MEDIO_GESTION")
		intES_G_EFE				=rsDET("ES_G_EFE")
		intES_G_TIT				=rsDET("ES_G_TIT")
		intES_G_NC				=rsDET("ES_G_NC")
		intInd					=rsDET("IND")
		intEsGestionCall		=rsDET("ES_GESTION_CALL")
		intGestionMOdulos		=rsDET("Gestion_Modulos")
		intCodCategoria			=rsDET("COD_CATEGORIA")
		intCodSubCategoria		=rsDET("COD_SUB_CATEGORIA")
		intCodGestion			=rsDET("COD_GESTION")
		strLogin				=rsDET("LOGIN")
		strComunica				=rsDET("COMUNICA")
		intPrioridadGtel		=rsDET("PRIORIDAD_GTEL")
		intPrioridadGmail		=rsDET("PRIORIDAD_GMAIL")
		intPrioridadGdir		=rsDET("PRIORIDAD_GDIR")
		Obs 					= UCASE(Trim(rsDET("OBSERVACIONES")))
		sessiontrCodGestion 	= rsDET("COD_CATEGORIA") & rsDET("COD_SUB_CATEGORIA") & rsDET("COD_GESTION")
		strNomGestion 			= rsDET("DESCRIPCION")
		intGestionComunica 		= rsDET("COMUNICA")
		intGestionGtel 			= rsDET("PRIORIDAD_GTEL")
		intGestionGmail 		= rsDET("PRIORIDAD_GMAIL")
		intGestionGdir	 		= rsDET("PRIORIDAD_GDIR")
		strTipoGestion 			= rsDET("GESTION_MODULOS")
		strLoginCobrador 		= rsDET("LOGIN")
		strAnexo 				= rsDET("ANEXO")
		intCuotasActivas		= rsDET("CUOTAS_ACTIVAS")
		intCuotasCanceladas		= rsDET("CUOTAS_CANCELADAS")
		intCuotasRetiradas		= rsDET("CUOTAS_RETIRADAS")
		intCuotasNoAsignables	= rsDET("CUOTAS_NO_ASIGNABLES")
		intDocConfirmar 		= rsDET("DOC_A_CONFIRMAR")
		intIDUsuario 			= rsDET("id_usuario")
        
        Gst_EsInterno                       = rsDET("EsInterno")    
        Gst_EsExterno                       = rsDET("EsExterno")    
        Gst_PuedenEscucharMisGrabaciones    = rsDET("PuedenEscucharMisGrabaciones")    
        Gst_PuedoEscucharGrabaciones        = rsDET("PuedoEscucharGrabaciones")   
        Ubicabilidad                        = rsDET("Ubicabilidad")
        Fecha_Tootltip                      = rsDET("Fecha_Tootltip")

        

		If Trim(intCuotasActivas) <> "" Then
			strTextoDocAct 		= "Doc.Asociados : " & intCuotasActivas & "<BR>"
		End If

		If Trim(intCuotasCanceladas) <> "" Then
			strTextoDocPag 		= "Doc.Cancelados : " & intCuotasCanceladas & "<BR>"
		End If

		If Trim(intCuotasRetiradas) <> "" Then
			strTextoDocRet 		= "Doc.Desasignados : " & intCuotasRetiradas & "<BR>"
		End If

		If Trim(intCuotasNoAsignables) <> "" Then
			strTextoDocNoAsig 	= "Doc.No Asignable : " & intCuotasNoAsignables & "<BR>"
		End If

		strTextoDoc = strTextoDocAct & strTextoDocPag & strTextoDocRet & strTextoDocNoAsig

        dim titulo 
        titulo  =""
        If TRIM(intDocConfirmar) = 1 or TRIM(intDocConfirmar) = 2 or TRIM(intDocConfirmar) = 3 or TRIM(intDocConfirmar) = 4 or TRIM(intDocConfirmar) = 5 Then  
            titulo  ="Confirmar / desconfirmar compromiso"
        end if 

       

	%>
	<tr bordercolor="#FFFFFF" class="td_hover">
		
		<td align="left" width="10" title="<%=rsDET("ID_GESTION")%>"><%=numero_fila%></td>
        <td width="20" title="<%=titulo%>">
			<%If TRIM(intDocConfirmar) = 1 Then%>
                <A HREF="#" onClick="ConfirmarCP(<%=rsDet("ID_GESTION")%>,'<%=dtmFecCompromiso%>','<%=strCodGestion%>','False')";><img src="../imagenes/icon_cruz_roja.jpg" border="0"></A>
			<%Elseif trim(intDocConfirmar) = 2 Then%>
				<A HREF="#" onClick="ConfirmarCP(<%=rsDet("ID_GESTION")%>,'<%=dtmFecCompromiso%>','<%=strCodGestion%>','False')";><img src="../imagenes/IconoVer.jpg" border="0"></A>
            <%Elseif trim(intDocConfirmar) = 3 Then%>
				<A HREF="#" onClick="ConfirmarCP(<%=rsDet("ID_GESTION")%>,'<%=dtmFecCompromiso%>','<%=strCodGestion%>','True')";><img src="../imagenes/icon_amarillo.jpg" border="0"></A>
			<%Elseif trim(intDocConfirmar) = 4 Then%>
				<A HREF="#" onClick="ConfirmarCP(<%=rsDet("ID_GESTION")%>,'<%=dtmFecCompromiso%>','<%=strCodGestion%>','True')";><img src="../imagenes/bt_confirmar.jpg" border="0"></A>
			<%Elseif trim(intDocConfirmar) = 5 Then%>
				<img src="../imagenes/bt_confirmar.jpg" border="0">
			<%End If%>          	
		</td>

		<td align="left" width="70" title="<%=Fecha_Tootltip%>" class=""><%=rsDET("FECHA_INGRESO_GESTION")%></td>
		<td align="left" width="50" class=""><%=rsDET("HORA_INGRESO")%></td>
		<td align="left" width="350" class=""><%=strNomGestion%></td>
		<td align="left" width="60" class=""><%=rsDET("FECHA_COMPROMISO")%></td>
		<td align="left" width="60" class=""><%=rsDET("FECHA_AGENDAMIENTO")%></td>
		<td align="left" width="50" class=""><%=rsDET("HORA_AGENDAMIENTO")%></td>
		<td align="left" width="350" class="" title="<%=Obs%>"><%=Mid(Obs,1,45)%></td>



		<%
		if trim(rsDET("NOM_MEDIO_GESTION")) <> "" AND NOT ISNULL(rsDET("NOM_MEDIO_GESTION")) then
			NOM_MEDIO_GESTION =trim(rsDET("NOM_MEDIO_GESTION"))
		else
			NOM_MEDIO_GESTION ="SIN MEDIO ASOCIADO"
		end if

		if trim(rsDET("NOM_CONTACTO_GESTION")) <> "" AND NOT ISNULL(rsDET("NOM_CONTACTO_GESTION")) then
			NOM_CONTACTO_GESTION =trim(rsDET("NOM_CONTACTO_GESTION"))
		else
			NOM_CONTACTO_GESTION ="SIN CONTACTO ASOCIADO"
		end if

		If trim(rsDET("TIPO_MEDIO_GESTION")) = "2" Then%>
			<td width="80" align="center" class="" title="<%=NOM_MEDIO_GESTION%>">
				<img src="../imagenes/Arroa.png" border="0">
			</td>

		<%ElseIf trim(rsDET("TIPO_MEDIO_GESTION")) = "1" Then%>
			<td WIDTH="80" class="" align="center"><%=rsDET("NOM_MEDIO_GESTION")%></td>

		<%ElseIf trim(rsDET("TIPO_MEDIO_GESTION")) = "3" Then%>
			<td WIDTH="80" class="" align="center">
			 	<img src="../imagenes/mod_direccion_va.png" title="<%=NOM_MEDIO_GESTION%>">
			</td>

		<%Else%>
			<td WIDTH="80"  class="">&nbsp;</td>
		<%End If%>
		
		<%If intGestionComunica = 0 AND intGestionGmail = 1 Then%>
			<td width="20" class="" align="center" title="<%=NOM_CONTACTO_GESTION%>">
				<img src="../imagenes/Contacto.rojo.png" border="0">
			</td>
		<%ElseIf intGestionComunica = 0 AND intGestionGtel = 1 Then%>
			<td width="20" class="" align="center" title="<%=NOM_CONTACTO_GESTION%>">
				<img src="../imagenes/Contacto.rojo.png" border="0">
			</td>
		<%ElseIf intGestionComunica = 0 AND intGestionGdir = 1 Then%>
			<td width="20" class="" align="center" title="<%=NOM_CONTACTO_GESTION%>">
				<img src="../imagenes/Contacto.rojo.png" border="0">
			</td>
		<%ElseIf intES_G_TIT = 1 AND intGestionGtel = 1 Then%>
			<td width="20" class="" align="center" title="<%=NOM_CONTACTO_GESTION%>">
				<img src="../imagenes/Contacto.azul.png" border="0">
			</td>
		<%ElseIf intES_G_EFE = 1 AND intGestionGtel = 1 Then%>
			<td width="20" class="" align="center" title="<%=NOM_CONTACTO_GESTION%>">
				<img src="../imagenes/Contacto.amarillo.png" border="0">
			</td>
		<%ElseIf intGestionComunica = 1 AND intGestionGmail = 1 Then%>
			<td width="20" class="" align="center" title="<%=NOM_CONTACTO_GESTION%>">
				<img src="../imagenes/Contacto.azul.png" border="0">
			</td>
		<%ElseIf intGestionComunica = 1 AND intGestionGdir = 1 Then%>
			<td width="20" class="" align="center" title="<%=NOM_CONTACTO_GESTION%>">
				<img src="../imagenes/Contacto.azul.png" border="0">
			</td>
		<%Else%>
			<td>&nbsp;</td>
		<%End If%>

		<%If trim(rsDET("TIPO_MEDIO_GESTION")) = "1" Then%>
			<td width="20" align="center" class="" title="<%=strNomMedioAgend%>">
				<img src="../imagenes/mod_telefono_va.png" border="0">
			</td>
		<%Else%>
			<td WIDTH="20"  class="">&nbsp;</td>
		<%End If%>
		
		<td align="left" width="80" class=""><%=UCASE(strLoginCobrador)%></td>
		<td width="20"class="" title="<%=strTextoDoc%>">
			<img src="../imagenes/carpeta1.png" border="0" onclick="trae_cuotas_por_gestion('<%=trim(rsDet("ID_GESTION"))%>')">
		</td>

        <%'response.write("--->" & Gst_PuedenEscucharMisGrabaciones) %>

        <%if Usr_PuedoEscucharGrabaciones = "Verdadero" then         %>
            <%if Gst_PuedenEscucharMisGrabaciones = "Verdadero" then  %>
                 <%if Ubicabilidad = "1" or Ubicabilidad = "2" or Ubicabilidad = "3"   then  %>
                    <% 'if ((TraeSiNo(session("perfil_sup"))="Si" or  TraeSiNo(session("perfil_adm"))="Si") and TraeSiNo(session("perfil_emp"))<>"Si") and intGestionGtel = 1 and strAnexo <> "" Then %>
			                <td width="20" class="">
				                <A HREF="#" onClick="TraerGrabacion('<%="0" & Replace(NOM_MEDIO_GESTION,"-","")%>','<%=rsDET("FECHA_INGRESO_GESTION")%>','<%=rsDET("HORA_INGRESO")%>','<%=intIDUsuario%>','<%=strAnexo%>');return false;";>
				                <img src="../imagenes/sound.png" border="0">
			                </A>
			                </td>
                 <%else%>	
                            <td width="20">&nbsp;</td>
                 <% End if%>
		    <%else%>	
			        <td width="20">&nbsp;</td>
		    <%End if%>

        <%else%>	
			<td width="20">&nbsp;</td>
        <%End if %>
	</tr>

	 <%
	 
	 strTextoDocAct = ""
	 strTextoDocPag = ""
	 strTextoDocRet = ""
	 strTextoDocNoAsig = ""
	 
	 intCorr = intCorr +1
	 response.Flush()
	 rsDET.movenext
	 Loop
	 
	if trim(numero_fila)=finales then
		inicio 		= numero_fila
		finales 	= cint(numero_fila) + 25
	%>		
		<tr style="border:0">
		<td colspan="14" align="center" id="refreso_mas_registros_<%=finales%>">

			<div class="mas_registros fondo_boton_130" onclick="bt_mostrar_mas_registros(<%=inicio%>,<%=finales%>)">Más registros</div>

		</td>
		</tr>

	<%end if	

	else%>
	
		<tr bgcolor="<%=strbgcolor%>" class="Estilo8" style="border:1px solid">																					
			<td colspan="11" align = "center"><h3>No Existen Gestiones Asociadas al Deudor</h3></td>	
		</tr>

	<%end if

	rsDET.close
	set rsDET=nothing

	%>

	</tbody>
	</table>

<%
elseIf trim(accion_ajax)="trae_cuotas_por_gestion" then
	IntIdGestion 		= request.querystring("ID_GESTION")
	strRUT_DEUDOR	=request.querystring("strRUT_DEUDOR")	
	strCodCliente   =request.querystring("strCodCliente")

	ssql="EXEC proc_Parametros_Tabla_Cliente '"&TRIM(strRUT_DEUDOR)&"','"&TRIM(strCodCliente)&"'"

	set rsCLI=Conn.execute(ssql)
	if not rsCLI.eof then
		strNomFormHon 		= ValNulo(rsCLI("FORMULA_HONORARIOS"),"C")
		strNomFormInt 		= ValNulo(rsCLI("FORMULA_INTERESES"),"C")

		strUsaSubCliente 	= rsCLI("USA_SUBCLIENTE")
		strUsaInteres 		= rsCLI("USA_INTERESES")
		strUsaHonorarios 	= rsCLI("USA_HONORARIOS")
		strUsaProtestos 	= rsCLI("USA_PROTESTOS")


		nombre_cliente		= rsCLI("RAZON_SOCIAL")
		intRetiroSabado		=Cint(rsCLI("RETIRO_SABADO"))
		strMsjRetiroSabado 	= ""

		If Trim(intRetiroSabado) = "1" Then
			strMsjRetiroSabado = "sabados,"
		End if

		strUbicFono 		=rsCLI("UBIC_FONO")
		strUbicEmail 		=rsCLI("UBIC_EMAIL")
		strUbicDireccion 	=rsCLI("UBIC_DIRECCION")
	end if

	strSql ="exec proc_Trae_Cuotas_Deudor '"&trim(strCodCliente)&"','"&trim(strRUT_DEUDOR)&"','','"&trim(IntIdGestion)&"','"&trim(strNomFormInt)&"', '"&trim(strNomFormHon)&"', '1', '"&trim(CH_TODOS_CUOTA)&"', '' "

	set rsTemp= Conn.execute(strSql)

	intTasaMensual 		= 2/100
	intTasaDiaria 		= intTasaMensual/30
	intCorrelativo		= 1
	strArrID_CUOTA 		=""
	intTotSelSaldo 		= 0
	intTotSelIntereses 	= 0
	intTotSelProtestos 	= 0
	intTotSelHonorarios = 0
	strDetCuota 		="mas_datos_adicionales.asp"

	strArrConcepto 		= ""
	strArrID_CUOTA 		= ""

	%>
	<table  border="1" id="table_tablesorter"  style="width:95%;" bordercolor="#000000" cellSpacing="0" cellPadding="1">
	<thead>
		<tr class="Estilo34">
			<td>&nbsp;</td>

			<%If Trim(strUsaSubCliente)="1" Then%>
				<th colspan = "2" >RUT CLIENTE</th>
				<th>NOMBRE CLIENTE</th>
			<%End If%>

			<th >N°DOC</th>
			<th >CUOTA</th>
			<th >FEC.VENC.</th>
			<th >ANT.</th>
			<th >TIPO DOC.</th>
			<th  align="center" width="70">CAPITAL</th>
			<%If Trim(strUsaInteres)="1" Then%>
			<th  align="center" width="70">INTERES</th>
			<%End If%>
			<%If Trim(strUsaProtestos)="1" Then%>
			<th  align="center" width="80">PROTESTOS</th>
			<%End If%>
			<%If Trim(strUsaHonorarios)="1" Then%>
			<th  align="center" width="90">HONORARIOS</th>
			<%End If%>
			<th  align="center" width="70">ABONO</th>
			<th  align="center" width="70">SALDO</th>
			<th >FECHA AGEND.</th>
			<td>&nbsp;</td>
			<td>&nbsp;</td>
			<td>&nbsp;</td>
			<td>&nbsp;</td>

		</tr>

	</thead>
	<tbody>
<%
	Do while not rsTemp.eof

		intSaldo 				=  rsTemp("SALDO")
		intValorCuota 			=  rsTemp("VALOR_CUOTA")
		intAbono 				= intValorCuota - intSaldo
		strNroDoc 				= rsTemp("NRO_DOC")
		strNroCuota				= rsTemp("NRO_CUOTA")
		strFechaVenc 			= rsTemp("FECHA_VENC")
		intProrroga 			= rsTemp("PRORROGA")
		strFechaVencOriginal 	= rsTemp("FECHA_VENC_ORIGINAL")
		strTipoDoc 				= rsTemp("TIPO_DOCUMENTO")
		intTipoGestion 			= rsTemp("TIPO_GESTION")
		intVerAgend 			= rsTemp("VER_AGEND")
		intGestionModulos 		= rsTemp("GESTION_MODULOS")
		strFechaAgendUG 		= rsTemp("FECHA_AGEND_ULT_GES")

		intAntiguedad 			= ValNulo(rsTemp("ANTIGUEDAD"),"N")

		intIntereses 			= rsTemp("INTERESES")
		intHonorarios 			= rsTemp("HONORARIOS")

		intProtestos 			= ValNulo(rsTemp("GASTOS_PROTESTOS"),"N")

		intTotDoc 				= intSaldo+intIntereses+intProtestos+intHonorarios

		intTotSelSaldo 			= intTotSelSaldo+intSaldo
		intTotSelAbono 			= intTotSelAbono+intAbono
		intTotSelValorCuota 	= intTotSelValorCuota+intValorCuota

		intTotSelIntereses 		= intTotSelIntereses+intIntereses
		intTotSelProtestos 		= intTotSelProtestos+intProtestos
		intTotSelHonorarios 	= intTotSelHonorarios+intHonorarios
		intTotSelDoc 			= intTotSelDoc+intTotDoc

		strArrConcepto 			= strArrConcepto & ";" & "CH_" & rsTemp("ID_CUOTA")
		strArrID_CUOTA 			= strArrID_CUOTA & ";" & rsTemp("ID_CUOTA")

		%>
		<tr class="Estilo34">

			<input name="TX_SALDO_<%=Replace(rsTemp("ID_CUOTA"),"-","_")%>" id="TX_SALDO_<%=Replace(rsTemp("ID_CUOTA"),"-","_")%>" type="hidden" value="<%=intTotDoc%>">
			<input name="TX_CAPITAL_<%=Replace(rsTemp("ID_CUOTA"),"-","_")%>" id="TX_CAPITAL_<%=Replace(rsTemp("ID_CUOTA"),"-","_")%>" type="hidden" value="<%=intValorCuota%>">
			<input name="TX_HONORARIOS_<%=Replace(rsTemp("ID_CUOTA"),"-","_")%>" id="TX_HONORARIOS_<%=Replace(rsTemp("ID_CUOTA"),"-","_")%>" type="hidden" value="<%=intHonorarios%>">
			<input name="TX_INTERESES_<%=Replace(rsTemp("ID_CUOTA"),"-","_")%>" id="TX_INTERESES_<%=Replace(rsTemp("ID_CUOTA"),"-","_")%>" type="hidden" value="<%=intIntereses%>">
			<input name="TX_PROTESTOS_<%=Replace(rsTemp("ID_CUOTA"),"-","_")%>" id="TX_PROTESTOS_<%=Replace(rsTemp("ID_CUOTA"),"-","_")%>" type="hidden" value="<%=intProtestos%>">


			<TD width="12">

				<INPUT TYPE="checkbox" checked="checked" NAME="CH_ID_CUOTA" id="CH_ID_CUOTA" value="<%=rsTemp("ID_CUOTA")%>">

			</TD>

			<%If Trim(strUsaSubCliente)="1" Then%>
				<td width="69"><%=rsTemp("RUT_SUBCLIENTE")%></td>
				<td>
					<a href="javascript:ventanaBusqueda('Busqueda.asp?strOrigen=1&TX_RUT_DEUDOR=<%=rsTemp("RUT_DEUDOR")%>&TX_NOMBRE=<%=nombre_deudor%>&TX_RUTSUBCLIENTE=<%=rsTemp("RUT_SUBCLIENTE")%>&TX_NOMBRE_SUBCLIENTE=<%=rsTemp("NOMBRE_SUBCLIENTE")%>')">
					<img src="../imagenes/buscar.png" border="0"></a></td>
				<td title="<%=rsTemp("NOMBRE_SUBCLIENTE")%>">
					<%=Mid(rsTemp("NOMBRE_SUBCLIENTE"),1,35)%>
				</td>
			<%End If%>

			<td><%=strNroDoc%></td>
			<td><%=strNroCuota%></td>

			<%If intProrroga = "0" Then%>
				<td><%=strFechaVenc%></td>
			<%Else%>
				<td bgcolor="#ff6666" title="Vencimiento Original: <%=strFechaVencOriginal%>">
				<%=strFechaVenc%></td>
			<%End If%>


			<td><%=intAntiguedad%></td>
			<td><%=strTipoDoc%></td>

			<td ALIGN="RIGHT"><SPAN style="display:none;"><%=intValorCuota%></SPAN><%=FN(intValorCuota,0)%></td>
			
			<%If Trim(strUsaInteres)="1" Then%>
				<td ALIGN="RIGHT">
					<SPAN style="display:none;"><%=intIntereses%></SPAN>
					<%=FN(intIntereses,0)%></td>
			<%End If%>
			<%If Trim(strUsaProtestos)="1" Then%>
				<td ALIGN="RIGHT"><SPAN style="display:none;"><%=intProtestos%></SPAN><%=FN(intProtestos,0)%></td>
			<%End If%>
			<%If Trim(strUsaHonorarios)="1" Then%>
			<td ALIGN="RIGHT"><SPAN style="display:none;"><%=intHonorarios%></SPAN><%=FN(intHonorarios,0)%></td>
			<%End If%>

			<td ALIGN="RIGHT"><SPAN style="display:none;"><%=intAbono%></SPAN><%=FN(intAbono,0)%></td>
			<td ALIGN="RIGHT"><SPAN style="display:none;"><%=intTotDoc%></SPAN><%=FN(intTotDoc,0)%></td>
		<td ALIGN="RIGHT"><%=strFechaAgendUG%></td>

			<td align="CENTER">
				<%
					intEstadoNR = ValNulo(rsTemp("NOTIFICACION_RECEPCIONADA"),"N")
					intEstadoFR = ValNulo(rsTemp("FACTURA_RECEPCIONADA"),"N")
				If (intEstadoNR = 0) OR (intEstadoFR = 0) Then
					strImagenGest = "audita_rojo.png"

				ElseIf (intEstadoNR = 2) OR (intEstadoFR = 2) Then
					strImagenGest = "audita_ama.png"

				Else
					strImagenGest = "audita_verde.png"
				End If
				%>

				<A HREF="#" onClick="AuditarDoc(<%=rsTemp("ID_CUOTA")%>)";>
				<img src="../imagenes/<%=strImagenGest%>" border="0"></A>
			</td>

			<td ALIGN="CENTER">
				<a href="javascript:ventanaGestionesPorDoc('gestiones_por_documento.asp?intID_CUOTA=<%=trim(rsTemp("ID_CUOTA"))%>&strCodCliente=<%=strCodCliente%>&strNroDoc=<%=trim(rsTemp("NRO_DOC"))%>')">
				<img src="../imagenes/icon_gestiones.jpg" border="0"></a>
			</td>

			<td>
				<a href="javascript:ventanaMas('<%=strDetCuota%>?ID_CUOTA=<%=trim(rsTemp("ID_CUOTA"))%>&cliente=<%=strCodCliente%>&strRUT_DEUDOR=<%=trim(rsTemp("RUT_DEUDOR"))%>&strNroDoc=<%=trim(rsTemp("NRO_DOC"))%>&strNroCuota=<%=rsTemp("NRO_CUOTA")%>')"><img src="../imagenes/Carpeta3.png" border="0"></a>
			</td>
			<td align="center">
				<%
				dtmFechaEstado 		= rsTemp("FECHA_ESTADO")
				dtmFechaCreacion 	= rsTemp("FECHA_CREACION")

				intIdUltGest 		= rsTemp("ID_ULT_GEST")

				dtmFechaIngresoUG 	= rsTemp("FECHA_INGRESO_UG")
				strCodUltgest 		= rsTemp("COD_ULT_GEST")

				strImagenGest1 		=""

				If (intVerAgend = 1 and ValNulo(rsTemp("DIFERENCIA"),"N") > 0) Then
					If (datevalue(dtmFechaIngresoUG) < datevalue(dtmFechaEstado)) and intGestionModulos = 3 Then
						''La fecha de ingreso de ultima gestion del documento (fun_trae_Ultima_Gestion_cuota_tit)es menor a la fecha de estado
						strImagenGest1 = "GestionarRoj.png"
					Else
						strImagenGest1 = "GestionarAzu.PNG"
					End If
				ElseIf (intTipoGestion = 1 or intTipoGestion = 2 ) Then
					'' Define VER AGEND en tabla GESTIONES_TIPO_GESTION igual a "0" y tipo de gestion compormiso pago o ruta
					strImagenGest1 = "NoGestionarAma.PNG"
				ElseIf intVerAgend = 0 or intTipoGestion = 3 or intTipoGestion = 4 Then
					'' Define VER AGEND en tabla GESTIONES_TIPO_GESTION igual a "0" dado a que gestión no se debe gestionar por el cobrador
					strImagenGest1 = "NoGestionarRojo.PNG"
				End If

				If strImagenGest1 <> "" Then %>
					<img src="../Imagenes/<%=strImagenGest1%>" border="0">
				<% Else %>
					&nbsp;
				<% End If %>
			</td>

		</tr>

		
		<%
		intCorrelativo = intCorrelativo + 1
		rsTemp.movenext
		loop

		vArrConcepto 		= split(strArrConcepto,";")
		vArrID_CUOTA 		= split(strArrID_CUOTA,";")
		intTamvConcepto 	= ubound(vArrConcepto)
		strArrID_CUOTA 		= Mid(strArrID_CUOTA,2,len(strArrID_CUOTA))
	%>
		</tbody>
		<thead class="totales">
		<tr class="Estilo34" height="22">
			<%If Trim(strUsaSubCliente)="1" Then
			 	strColspan = "colspan= 9"
			Else
				 strColspan = "colspan= 6"
			End If%>

			<td <%=strColspan%> >&nbsp;&nbsp;&nbsp;&nbsp;Totales :</td>
			<td ALIGN="RIGHT"><%=FN(intTotSelValorCuota,0)%></td>
			<%If Trim(strUsaInteres)="1" Then%>
				<td ALIGN="RIGHT"><%=FN(intTotSelIntereses,0)%></td>
			<%End If%>
			<%If Trim(strUsaProtestos)="1" Then%>
				<td ALIGN="RIGHT"><%=FN(intTotSelProtestos,0)%></td>
			<%End If%>
			<%If Trim(strUsaHonorarios)="1" Then%>
				<td ALIGN="RIGHT"><%=FN(intTotSelHonorarios,0)%></td>
			<%End If%>


			<td ALIGN="RIGHT"><%=FN(intTotSelAbono,0)%></td>
				<td ALIGN="RIGHT"><%=FN(intTotSelDoc,0)%></td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
			</tr>


			<tr class="Estilo34" height="25">

			<td <%=strColspan%>>&nbsp;&nbsp;&nbsp;&nbsp;Totales Seleccionados:</td>
			<td ALIGN="RIGHT"><span id="span_TX_CAPITAL" style="font-weight:bold;">0</span>
				<INPUT TYPE="hidden" NAME="TX_CAPITAL" ID="TX_CAPITAL" DISABLED style="text-align:right;width:90" size="10" size="10" onkeyup="format(this)" onchange="format(this)">
			</td>



			<% If Trim(strUsaInteres)="1" Then%>
				<td ALIGN="RIGHT"><INPUT TYPE="TEXT" NAME="TX_INTERESES" ID="TX_INTERESES" DISABLED style="text-align:right;width:90" size="10" size="10" onkeyup="format(this)" onchange="format(this)"></td>
			<% Else%>
				<INPUT TYPE="hidden" NAME="TX_INTERESES">
			<% End If%>

			<% If Trim(strUsaProtestos)="1" Then%>
				<td ALIGN="RIGHT"><INPUT TYPE="TEXT" NAME="TX_PROTESTOS" ID="TX_PROTESTOS" DISABLED style="text-align:right;width:90" size="10" size="10" onkeyup="format(this)" onchange="format(this)"></td>
			<% Else%>
				<INPUT TYPE="hidden" NAME="TX_PROTESTOS">
			<% End If%>

			<% If Trim(strUsaHonorarios)="1" Then%>
				<td ALIGN="RIGHT"><INPUT TYPE="TEXT" NAME="TX_HONORARIOS" ID="TX_HONORARIOS" DISABLED style="text-align:right;width:90" size="10" onkeyup="format(this)" onchange="format(this)"></td>
			<% Else%>
				<INPUT TYPE="hidden" NAME="TX_HONORARIOS">
			<% End If%>



			<td>&nbsp;</td>
			<td ALIGN="RIGHT" ><span  id="span_TX_SALDO" style="font-weight:bold;">0</span>
				<INPUT TYPE="hidden" ID="TX_SALDO" NAME="TX_SALDO" DISABLED style="text-align:right;width:90" size="10" size="10" onkeyup="format(this)" onchange="format(this)">
			</td>
			<td>&nbsp;</td>
			<td>&nbsp;</td>
			<td>&nbsp;</td>
			<td>&nbsp;</td>
			<td>&nbsp;</td>
		</tr>
		</thead>
		<INPUT TYPE="hidden" NAME="strArrID_CUOTA" VALUE="<%=strArrID_CUOTA%>">
			
	</table>



<%	
elseif trim(accion_ajax)="refresca_subcategoria" then%>
	<select name="cmbsubcat" id="cmbsubcat" onChange="cargagest(this.value,cmbcat.value);"  >
		  <option value="">SELECCIONE</option>
	</select>
<%
elseif trim(accion_ajax)="refresca_gestion" then%>
	<select name="cmbgest" id="cmbgest" onChange="cajas_tipo_gestion();" >
		<option value="">SELECCIONE</option>
	</select>
<%

end if





if trim(accion_ajax)="refresca_lugar_pago" then 

'abrirscg()

forma_Pago      = request("forma_Pago")
strCodCliente   = request("strCodCliente")
strRutDeudor    = request("rut")
strSql = ""

        IF  forma_Pago = "TR" or  FORMA_PAGO ="DP" THEN '' TRANSFERENCIA/DEPOSITO
                            strSql = strSql & " SELECT 2 AS TIPO, NOMBRE + ' ' + UBICACION AS LUGAR_PAGO , ORDEN , " & _
                                              "ID_FORMA_RECAUDACION ID, 'FORMA_RECAUDACION' TIPO " & _
                                              " FROM FORMA_RECAUDACION " & _
                                              " WHERE COD_CLIENTE = '" & TRIM(STRCODCLIENTE) & "' AND TIPO  = 'DEPOSITO' " & _
                                              " ORDER BY ORDEN ASC"
        END IF 


        IF  forma_Pago = "CF" or  FORMA_PAGO ="CD"  or  FORMA_PAGO ="EF" or FORMA_PAGO = "CF" or  FORMA_PAGO ="TC" or  FORMA_PAGO ="TD" or  FORMA_PAGO ="TRE" THEN  '' CHEQUE A FECHA/CHEQUE AL DIA/EFECTIVO/CF

                        strSql = "SELECT 1 as TIPO, REPLACE(CALLE + ' ' + NUMERO + ' ' + RESTO + ' ' + COMUNA,'  ',' ') as LUGAR_PAGO , " & _
                                 " 0 AS ORDEN, id_direccion ID , 'DIRECCION' tipo " & _
                                 " FROM DEUDOR_DIRECCION WHERE RUT_DEUDOR = '" & TRIM(strRutDeudor) & "' AND ESTADO <> 2 "
                    
				        strSql = strSql & " UNION"

                        strSql = strSql  & " SELECT 2 AS TIPO, NOMBRE + ' ' + UBICACION AS LUGAR_PAGO , ORDEN , " & _
                                      "ID_FORMA_RECAUDACION ID, 'FORMA_RECAUDACION' TIPO " & _
                                      " FROM FORMA_RECAUDACION " & _
                                      " WHERE COD_CLIENTE = '" & TRIM(STRCODCLIENTE) & "' AND TIPO  = 'PRESENCIAL' " & _
                                      " ORDER BY ORDEN ASC"
        END IF 
        
        IF  forma_Pago = "VV" or  FORMA_PAGO ="PG"  or  FORMA_PAGO ="LT" THEN '' VALE VISTA/PAGARE/LETRA

                                            strSql = "SELECT 1 as TIPO, REPLACE(CALLE + ' ' + NUMERO + ' ' + RESTO + ' ' + COMUNA,'  ',' ') as LUGAR_PAGO , " & _
                                                     " 0 AS ORDEN, id_direccion ID , 'DIRECCION' tipo " & _
                                                     " FROM DEUDOR_DIRECCION WHERE RUT_DEUDOR = '" & TRIM(strRutDeudor) & "' AND ESTADO <> 2"
                    
				                            strSql = strSql & " UNION"

                                            strSql = strSql  & " SELECT 2 AS TIPO, NOMBRE + ' ' + UBICACION AS LUGAR_PAGO , ORDEN , " & _
                                                          "ID_FORMA_RECAUDACION ID, 'FORMA_RECAUDACION' TIPO " & _
                                                          " FROM FORMA_RECAUDACION " & _
                                                          " WHERE COD_CLIENTE = '" & TRIM(STRCODCLIENTE) & "' AND TIPO  in ('PRESENCIAL','RETIRO') " & _
                                                          " ORDER BY ORDEN ASC"
          END IF 


		         On error Resume Next 
				        set rsDIR=Conn.execute(strSql)
                     'response.write(",Error," & strSql )
                     'response.end
                If Err.number<>0 then
                   response.write(",Error," & Err.Description )
                   response.end
               end if 

			   valor = "<select name=""CB_ID_DIRECCION_COBRO_DEUDOR"" id=""CB_ID_DIRECCION_COBRO_DEUDOR""><option value=""0"">SELECCIONE</option>"
			   
            	DO WHILE NOT  rsDIR.eof
			            direccion = rsDIR("LUGAR_PAGO")
				        valor = valor & "<option value='"  &  rsDIR("ID")&"-"&rsDIR("TIPO")  &  "'>" & direccion &  "</option>" 
        				rsDIR.movenext
				loop
               
				Objeto = valor & "</select>"
				response.write (Objeto)
		
               response.end

end if

%>
<script>
   if($("#fonoActual").val()!= "" && $("#strPaginaOrigen").val()== ""){	        
        $("#CB_ID_MEDIO_GESTION").prop('disabled', true);  
	}
</script>