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
	<!--#include file="../lib/comunes/rutinas/TraeCampo.inc" -->
	<!--#include file="../lib/asp/comunes/general/rutinasSCG.inc" -->
	<!--#include file="../lib/asp/comunes/general/RutinasVarias.inc" -->
	<!--#include file="../lib/asp/comunes/general/rutinasBooleano.inc"-->

	<script language="JavaScript" src="../javascripts/cal2.js"></script>
	<script language="JavaScript" src="../javascripts/cal_conf2.js"></script>
	<script language="JavaScript" src="../javascripts/validaciones.js"></script>
	<link href="../css/style.css" rel="stylesheet" type="text/css">
	<link href="../css/style_generales_sistema.css" rel="stylesheet" >	

	<script src="../Componentes/jquery-1.9.2/js/jquery-1.8.3.js"></script>
	<script src="../Componentes/jquery-1.9.2/js/jquery-ui-1.9.2.custom.js"></script>
	<link rel="stylesheet" type="text/css" href="../Componentes/jquery-1.9.2/css/start/jquery-ui-1.9.2.custom.css">


<%
	Response.CodePage=65001
	Response.charset ="utf-8"

	id_gestion 		= request("id_gestion")

	strRutDeudor = request("rut")

	if trim(strRutDeudor) = "" Then
		strRutDeudor = session("session_RUT_DEUDOR") 
	End if

    

	session("session_RUT_DEUDOR") = strRutDeudor
	
	strGraba = request("strGraba")
	strTipo = request("strTipo")
	strCodCliente = session("ses_codcli")

	usuario=session("session_idusuario")

    mostrar = request("mostrar")
    ''''response.Write("asfd" & mostrar)
    
    if  mostrar = "" then   mostrar = true
    
    Dim strEsconder

    if mostrar = "False" then
      strEsconder = " style=display:none;"
    else
      strEsconder = ""
    end if


	AbrirSCG()
	
	strSql = "SELECT  FECHA_COMPROMISO FROM GESTIONES"
	strSql = strSql & " WHERE ID_GESTION = " & id_gestion

	set rsFechacomp=Conn.execute(strSql)
	
	dtmFecCompGest = rsFechacomp("FECHA_COMPROMISO")




	
	''' response.write "strEsconder = " & strEsconder

	If strRutDeudor <> "" then
		strNombreDeudor = TraeNombreDeudor(Conn,strRutDeudor)
	Else
		strNombreDeudor=""
	End if

	If Trim(request("strGraba")) = "SI" Then

		If strTipo = "DP" Then

			strSql = "UPDATE GESTIONES SET OBSERVACIONES_CONF = '" & Mid(Request("TX_OBSERVACIONES"),1,300) & "'"
			strSql = strSql & " WHERE ID_GESTION = " & id_gestion

			'Response.write "<br>strSql" & strSql
			set rsUpdate=Conn.execute(strSql)

			strArchivoAsp = "principal.asp?a=1"

			strSql = "SELECT ID_CUOTA FROM CUOTA WHERE RUT_DEUDOR='" & strRutDeudor & "' AND COD_CLIENTE='" & strCodCliente & "' AND  (DATEDIFF(MINUTE,(ISNULL(FECHA_AGEND_ULT_GES,GETDATE()-200)+ convert(varchar(10),ISNULL(HORA_AGEND_ULT_GES,'00:00'),108)),GETDATE()) >= 0) AND ESTADO_DEUDA IN (SELECT CODIGO FROM ESTADO_DEUDA WHERE ACTIVO = 1) "

			set rsValida = Conn.execute(strSql )

			If Not rsValida.Eof Then
				strArchivoAsp = "detalle_gestiones.asp?strFonoAgestionar=" & telges & "&strContactoSel=" & intIdContacto
			End If

			'Response.write "strArchivoAsp=" & strArchivoAsp

			Response.Redirect strArchivoAsp & "&rut=" & strRutDeudor & "&cliente=" & strCodCliente

		Else

			If strTipo = "C" or strTipo = "D" Then

				strSql = "UPDATE GESTIONES SET OBSERVACIONES_CONF = '" & Mid(Request("TX_OBSERVACIONES"),1,300) & "'"
				strSql = strSql & " WHERE ID_GESTION = " & id_gestion

				'Response.write "<br>strSql" & strSql
				set rsUpdate=Conn.execute(strSql)

			End If


			strSql = "SELECT ID_CUOTA FROM CUOTA WHERE RUT_DEUDOR='" & strRutDeudor & "' AND COD_CLIENTE='" & strCodCliente & "' AND SALDO > 0 AND ID_CUOTA IN (SELECT ID_CUOTA FROM GESTIONES_CUOTA WHERE ID_GESTION = " & id_gestion & ")"
			''Response.write "<BR>strSql=" & strSql
			set rsTemp= Conn.execute(strSql)


			Do until rsTemp.eof

				strObjeto = "CH_" & Replace(Trim(rsTemp("ID_CUOTA")),"-","_")

				If UCASE(Request(strObjeto)) = "ON" Then

					intID_CUOTA = rsTemp("ID_CUOTA")

					If strTipo = "C" Then

						strSql = "UPDATE GESTIONES_CUOTA SET CONFIRMACION_CP = 'S' ,FechaConfirmaCp = getdate() ,UsuarioConfirmaCp = " & usuario
                        strSql = strSql  & " WHERE ID_CUOTA = " & intID_CUOTA & " AND ID_GESTION = " & id_gestion
						Response.write "<br>strSql" & strSql
						set rsUpdate=Conn.execute(strSql)

						strSql = "UPDATE CUOTA SET FECHA_AGEND_ULT_GES = [dbo].[fun_Trae_Proxima_Fecha_Habil] (DATEADD(DAY,2,'" & trim(dtmFecCompGest) & "')), HORA_AGEND_ULT_GES = '11:00' WHERE ID_CUOTA = " & intID_CUOTA 
						'Response.write "<br>strSql" & strSql
						'response.end()
						set rsUpdate=Conn.execute(strSql)

					End If

					If strTipo = "D" Then

						strSql = "UPDATE GESTIONES_CUOTA SET CONFIRMACION_CP = 'N' ,FechaConfirmaCp = getdate() ,UsuarioConfirmaCp = " & usuario 
                        strSql = strSql  & " WHERE ID_CUOTA = " & intID_CUOTA & " AND ID_GESTION = " & id_gestion
						'Response.write "<br>strSql" & strSql
						set rsUpdate=Conn.execute(strSql)

						strSql = "UPDATE CUOTA SET FECHA_AGEND_ULT_GES = [dbo].[fun_Trae_Proxima_Fecha_Habil] (DATEADD(DAY,-2,'" & trim(dtmFecCompGest)& "')), HORA_AGEND_ULT_GES = '11:00' WHERE ID_CUOTA = " & intID_CUOTA
						''Response.write "<br>strSql" & strSql
						set rsUpdate=Conn.execute(strSql)

					End If

				End if

			rsTemp.movenext

			Loop
			rsTemp.close
			set rsTemp=nothing

				strSql = "SELECT COUNT(ID_CUOTA) AS TOTAL_CUOTAS, SUM(ESTADO_CP) AS CUOTAS_CONFIRMADAS FROM (SELECT ID_CUOTA, (CASE WHEN ISNULL(CONFIRMACION_CP,'N')='N' THEN 0 ELSE 1 END) AS ESTADO_CP FROM GESTIONES_CUOTA WHERE ID_GESTION = " & id_gestion & ") AS PP"
				'Response.write "<br>strSql = " & strSql
				set rsInf=Conn.execute(strSql)

				intTotalCuotas = rsInf("TOTAL_CUOTAS")
				intTotalCuotasConf = rsInf("CUOTAS_CONFIRMADAS")


			If strTipo = "C" and intTotalCuotas = intTotalCuotasConf Then

				strSql = "UPDATE GESTIONES SET FECHA_AGENDAMIENTO = [dbo].[fun_Trae_Proxima_Fecha_Habil] (DATEADD(DAY,2,'" & dtmFecCompGest & "')), HORA_AGENDAMIENTO = '11:00'"
				strSql = strSql & " WHERE ID_GESTION = " & id_gestion
				'Response.write "<br>strSql" & strSql
				set rsUpdate=Conn.execute(strSql)

			Else

				strSql = "UPDATE GESTIONES SET FECHA_AGENDAMIENTO = [dbo].[fun_Trae_Proxima_Fecha_Habil] (DATEADD(DAY,-2,'" & dtmFecCompGest & "')), HORA_AGENDAMIENTO = '11:00'"
				strSql = strSql & " WHERE ID_GESTION = " & id_gestion
				set rsUpdate=Conn.execute(strSql)

			End If


			Response.Redirect "detalle_gestiones.asp?strFonoAgestionar=" & telges & "&strContactoSel=" & intIdContacto & "&rut=" & strRutDeudor & "&cliente=" & strCodCliente

			''Response.End

		End If

	End If


	If Trim(request("strAgendar")) = "SI" Then
	
		strSql = "UPDATE GESTIONES SET OBSERVACIONES_CONF = '" & Mid(Request("TX_OBSERVACIONES"),1,300) & "'"
		strSql = strSql & " WHERE ID_GESTION = " & id_gestion

		'Response.write "<br>strSql" & strSql
		set rsUpdate=Conn.execute(strSql)

		dtmFecAgend = Request("TX_FEC_AGEND")
		strHoraAgend = Request("TX_HORA_AGEND")
		If trim(strHoraAgend)="" Then strHoraAgend = "08:00"

		'Response.write "dtmFecAgend=" & dtmFecAgend

		If dtmFecAgend <> "" Then
			dtmFecAgend = dtmFecAgend & " " &  strHoraAgend & ":00"
		End If

		If dtmFecAgend = "" Then
			strSql = "UPDATE DEUDOR SET FECHA_AGEND_CONF = NULL, HORA_AGEND_CONF = NULL"
			strSql = strSql & " WHERE COD_CLIENTE = '" & strCodCliente & "' AND RUT_DEUDOR = '" & strRutDeudor & "'"
		Else
			strSql = "UPDATE DEUDOR SET FECHA_AGEND_CONF = '" & dtmFecAgend & "', HORA_AGEND_CONF = '" & strHoraAgend & "'"
			strSql = strSql & " WHERE COD_CLIENTE = '" & strCodCliente & "' AND RUT_DEUDOR = '" & strRutDeudor & "'"
		End If
		''Response.write "strSql=" & strSql
		set rsUpdate=Conn.execute(strSql)
		%>

		<SCRIPT>
			IrAPrincipal()
		</SCRIPT>
		<%
	End If

	If Trim(request("strLimpiar")) = "SI" Then

		'strSql = "UPDATE DEUDOR SET FECHA_AGEND_CONF = NULL, HORA_AGEND_CONF = NULL"
		'strSql = strSql & " WHERE COD_CLIENTE = '" & strCodCliente & "' AND RUT_DEUDOR = '" & strRutDeudor & "'"

		''Response.write "strSql=" & strSql
		'set rsUpdate=Conn.execute(strSql)

		strSql = "UPDATE GESTIONES SET OBSERVACIONES_CONF = NULL"
		strSql = strSql & " WHERE ID_GESTION = " & id_gestion
		set rsUpdate=Conn.execute(strSql)


		%>

		<SCRIPT>
			IrAPrincipal()
		</SCRIPT>
		<%
	End If


%>
	<style type="text/css">
	<!--
	.Estilo13 {color: #FFFFFF}
	.Estilo27 {color: #FFFFFF}
	.Estilo1 {
		color: #FF0000;
		font-weight: bold;
		font-family: Arial, Helvetica, sans-serif;
	--> }
	    .style3
        {
            width: 34px;
        }
	    .style5
        {
            width: 61px;
        }
        .style6
        {
            width: 90px;
        }
        .style7
        {
            width: 90px;
        }
	</style>

</head>	
<body>
<form name="datos" method="post">
<INPUT TYPE="hidden" NAME="intOrigen" value="<%=intOrigen%>">
<INPUT TYPE="hidden" NAME="strAgendar" value="">

<div class="titulo_informe">Confirmación - Desconfirmación de Compromisos Pagos</div>
<table width="90%" border="0" bordercolor="#999999" align="CENTER">
    <tr>
    <td>

	  <%

	strSql = "SELECT convert(varchar,FECHA_COMPROMISO,105) FECHA_COMPROMISO ,HORA_DESDE ,HORA_HASTA,isnull(OBSERVACIONES_CAMPO,'') OBSERVACIONES_CAMPO ,OBSERVACIONES_CONF FROM GESTIONES WHERE ID_GESTION = " & id_gestion
    ''' response.Write("asd" & strSql)
	set rsGestion=Conn.execute(strSql)
	if Not rsGestion.eof Then
		FECHA_COMPROMISO = rsGestion("FECHA_COMPROMISO")
        HORA_DESDE = rsGestion("HORA_DESDE")
        HORA_HASTA = rsGestion("HORA_HASTA")
        OBSERVACIONES_CAMPO = rsGestion("OBSERVACIONES_CAMPO")
        strObsNueva = rsGestion("OBSERVACIONES_CONF")
	End If


	strSql="SELECT CONVERT(VARCHAR(10),FECHA_AGEND_CONF,103) AS FECHA_AGEND_CONF, HORA_AGEND_CONF FROM DEUDOR WHERE RUT_DEUDOR = '" & strRutDeudor & "' AND COD_CLIENTE='" & strCodCliente & "'"
	set rsDeudor=Conn.execute(strSql)
	If not rsDeudor.eof then
		dtmFechaAgend = Trim(rsDeudor("FECHA_AGEND_CONF"))
		dtmHoraAgend = Trim(rsDeudor("HORA_AGEND_CONF"))
	End If

	%>


	<table width="100%" border="0" bordercolor="#FFFFFF"  class="estilo_columnas">
		<thead>
		<tr >
			<td>RUT DEUDOR</td>
			<td>NOMBRE O RAZON SOCIAL</td>
			<td>FECHA COMPROMISO</td>
			<td>HORARIO DESDE</td>
            <td>HORARIO HASTA</td>
            <td>OBSERVACIÓN GESTIÓN</td>
		</tr>
		</thead>
	      <tr bgcolor="#FFFFFF" class="Estilo8">

			<td>
				<A HREF="principal.asp?TX_RUT=<%=strRutDeudor%>">
				<acronym title="Llevar a pantalla de selección"><%=strRutDeudor%></acronym>
				</A>
			</td>
								
			<td><%=strNombreDeudor%><INPUT TYPE="hidden" NAME="rut" value="<%=strRutDeudor%>"> </td>
            <td><%=FECHA_COMPROMISO%></td>
            <td><%=HORA_DESDE%></td>
            <td><%=HORA_HASTA%></td>
            <td align="center"><img src='../imagenes/priorizar_normal.png' border='0' title="<%=OBSERVACIONES_CAMPO%>"></td>
	      </tr>
    </table>
	</td>
	</tr>

	<tr>
	<td>

	<table width="90%" border="0" ALIGN="left">
	 <tr>
	 	<td height="22">
	 	<font class="subtitulo_informe">> Detalle</font></td>
	</tr>
	</table>

	<table width="100%" border="0" ALIGN="CENTER">
	  <tr>
	    <td>
		<%
		If Trim(strRutDeudor) <> "" then
		abrirscg()

			strSql=" SELECT CUOTA.ID_CUOTA, RUT_SUBCLIENTE, NOMBRE_SUBCLIENTE, DATEDIFF(DAY,FECHA_VENC,GETDATE()) AS ANT_DIAS,"
			strSql = strSql & "	RUT_DEUDOR, IsNull(FECHA_VENC,'01/01/1900') as FECHA_VENC,"
			strSql = strSql & "	IsNull(VALOR_CUOTA,0) as VALOR_CUOTA,IsNull(SALDO,0) as SALDO,isnull(GASTOS_PROTESTOS,0) as GASTOS_PROTESTOS,"
			strSql = strSql & "	IsNull(USUARIO_ASIG,0) as USUARIO_ASIG, ESTADO_DEUDA.DESCRIPCION AS ESTADO_DEUDA,"
			strSql = strSql & "	NRO_CUOTA, SUCURSAL, NRO_DOC,"
			strSql = strSql & "	TIPO_DOCUMENTO.NOM_TIPO_DOCUMENTO, ISNULL(CUSTODIO,'LLACRUZ') AS CUSTODIO,ISNULL(GC.CONFIRMACION_CP,'N') AS CONFIRMACION_CP"
            strSql = strSql & "	,isnull(convert(varchar ,FechaConfirmaCp,105),'---') FechaConfirmaCp, isnull(SUBSTRING(convert(varchar,FechaConfirmaCp, 108),1,5),'') HoraConfirmaCp,isnull(ucp.LOGIN,'---') UsuarioConfirmaCp "
		    strSql = strSql & "	,isnull(Estado_Ruta,'0') Estado_Ruta,isnull(convert(varchar ,Fecha_Estado_Ruta,105),'---') Fecha_Estado_Ruta,isnull(SUBSTRING(convert(varchar,Fecha_Estado_Ruta, 108),1,5),'') HoraEstado_Ruta,isnull(ucr.LOGIN,'---') Usuario_Estado_Ruta"
            strSql = strSql & "	,GC.OBSERVACION_RUTA OBS ,CASE WHEN CUOTA.ID_ULT_GEST_GENERAL =  GC.ID_GESTION THEN 1 ELSE 0 END  VALIDA_GESTION "
			strSql = strSql & "	FROM CUOTA INNER JOIN TIPO_DOCUMENTO ON CUOTA.TIPO_DOCUMENTO = TIPO_DOCUMENTO.COD_TIPO_DOCUMENTO"
			strSql = strSql & "			   INNER JOIN ESTADO_DEUDA ON CUOTA.ESTADO_DEUDA = ESTADO_DEUDA.CODIGO"
			strSql = strSql & "			   INNER JOIN GESTIONES_CUOTA GC ON CUOTA.ID_CUOTA = GC.ID_CUOTA AND GC.ID_GESTION = " & id_gestion
            strSql = strSql & "			   LEFT JOIN USUARIO UCP ON UCP.ID_USUARIO = GC.UsuarioConfirmaCp  "
            strSql = strSql & "			   LEFT JOIN USUARIO UCR ON UCR.ID_USUARIO = GC.Usuario_Estado_Ruta  "

			strSql = strSql & "	WHERE ESTADO_DEUDA.ACTIVO = 1"
			strSql = strSql & "	ORDER BY CUENTA, FECHA_VENC DESC"

			''Response.write "<BR>strSql=" & strSql
			'response.End()
			set rsDET=Conn.execute(strSql)
			if not rsDET.eof then
			%>
			  <table width="100%" border="0" bordercolor="#FFFFFF" class="intercalado" style="width:100%;" id="tbl_Procesa">
			  	<thead>
		        <tr >
                  <td colspan="4" align="center" >CONFIRMACI&OacuteN COMPROMISO</td>
	    		  <td colspan="9" align="center" >DETALLE DE DOCUMENTOS</td>
                  <td colspan="4" align="center" >PROCESO RUTA</td>
		          </tr>

                <tr >
		          <td class="style3" >
		          	<a href="#" onClick="marcar_boxes(true);" <%=strEsconder%> >M</a>&nbsp
	    			<a href="#" onClick="desmarcar_boxes(true);" <%=strEsconder%> >D</a> </td>
	    		  <td align="CENTER">ESTADO</td>
                  <td align="CENTER">FECHA</td>
                  <td align="CENTER">USUARIO</td>
	    		  <td align="CENTER">RUT CLIENTE</td>
	    		  <td align="CENTER">NOMBRE CLIENTE</td>
		          <td align="CENTER">Nº DOC</td>
	    		  <td align="CENTER">CUOTA</td>
		          <td align="CENTER">FEC.VENC.</td>
		          <td align="CENTER">ANT.</td>
		          <td align="CENTER" class="style5">TIPO DOC.</td>
		          <td align="CENTER" class="style6">CAPITAL</td>
		          <td align="CENTER" class="style7">SALDO</td>
                  <td align="CENTER">ESTADO</td>
                  <td align="CENTER">FECHA</td>
                  <td align="CENTER">USUARIO</td>
                  <td align="CENTER">OBS</td>
                   
		          </tr>
		         </thead> 
		         <tbody>
				<%
				strArrConcepto = ""
				strArrID_CUOTA = ""

				Do While Not rsDET.eof

				strConfirmada = "NO CONFIRMADA"
                ImgConfirmda = "../imagenes/icon_amarillo.jpg"
				If Trim(rsDET("CONFIRMACION_CP")) <> "N" Then
					strConfirmada = "CONFIRMADA"
                    ImgConfirmda = "../imagenes/bt_confirmar.jpg"
				End If


                dim Estado_Ruta
                Estado_Ruta =""
                Fecha_Estado_Ruta=""
                Usuario_Estado_Ruta =""
                
                If Trim(rsDET("Estado_Ruta")) = "1" Then
                    Estado_Ruta="PROCESADO"
                elseIf Trim(rsDET("Estado_Ruta")) ="2" Then
                    Estado_Ruta="RECHAZADO"
                else 
                Estado_Ruta ="NO PROCESADO"
                end if 


              dim TitleCompormiso   
              TitleCompormiso   ="" 
              AlingCompormiso ="Center"
              AlingRuta ="Center"

              if rsDET("HoraConfirmaCp") <> "" then 
                TitleCompormiso  ="Hora: " & rsDET("HoraConfirmaCp")
                AlingCompormiso ="Left"
              end if 

              if rsDET("HoraEstado_Ruta") <> "" then 
                TitleRuta  ="Hora: " & rsDET("HoraEstado_Ruta")
                AlingRuta = "Left"
              end if 
                
       	        strArrConcepto = strArrConcepto & ";" & "CH_" & rsDET("ID_CUOTA")
		        strArrID_CUOTA = strArrID_CUOTA & ";" & rsDET("ID_CUOTA")

                
                IF Estado_Ruta = "NO PROCESADO" THEN 
                    TituloObs      = "NO PROCESADO" 
                ELSEIF Estado_Ruta <> "NO PROCESADO" AND  rsDET("obs") = "" THEN 
                    TituloObs      = "PROCESADA SIN INFORMACIÓN" 
                ELSEIF Estado_Ruta <> "NO PROCESADO" AND  rsDET("obs") <> ""  THEN 
                      TituloObs      = rsDET("obs") 
                END IF  
                

                
                'response.Write("asd" & mostrar)

				%>
		        <tr bordercolor="#999999" >
		          <td class="style3">
                  <% if rsDET("VALIDA_GESTION") = 0 or mostrar ="False"  then %>
                  <INPUT TYPE=checkbox NAME="CH_<%=Replace(rsDET("ID_CUOTA"),"-","_")%>" disabled=disabled  title=" No se puede confirmar el compromiso producto a que posee gestión posterior o el compromiso esta vencido" >
                  <% else %>
                    <INPUT TYPE=checkbox NAME="CH_<%=Replace(rsDET("ID_CUOTA"),"-","_")%>"  >
                  <%end if %>
                  
                  </td>
		          <td align="center"> 
                  <img src="<%=ImgConfirmda%>" border="0" title="<%=strConfirmada%>">
                  </td>
                  <td align="<%=AlingCompormiso%>" title="<%=TitleCompormiso%>"><%=rsDET("FechaConfirmaCp")%></td>
                  <td align="<%=AlingCompormiso%>" ><%=rsDET("UsuarioConfirmaCp")%></td>
                  <td><%=rsDET("RUT_SUBCLIENTE")%></td>
		          <td  title="<%=rsDET("NOMBRE_SUBCLIENTE")%>"><%=mid(rsDET("NOMBRE_SUBCLIENTE"),1,30)%></td>
		          <td><div align="right"><%=rsDET("NRO_DOC")%></div></td>
		          <td><div align="right"><%=rsDET("NRO_CUOTA")%></div></td>
		          <td><div align="right"><%=rsDET("FECHA_VENC")%></div></td>
		          <td><div align="right"><%=rsDET("ANT_DIAS")%></div></td>
		          <td align="right" class="style5"><%=rsDET("NOM_TIPO_DOCUMENTO")%></td>
		          <td align="right" class="style6" >$ <%=FN((rsDET("VALOR_CUOTA")),0)%></td>
		          <td align="right" class="style7" >$ <%=FN((rsDET("SALDO")),0)%></td>
                  <td >&nbsp&nbsp<%=Estado_Ruta%></td>
                  <td align="<%=AlingRuta%>" title="<%=TitleRuta%>"><%=rsDET("Fecha_Estado_Ruta")%></td>
                  <td align="<%=AlingRuta%>"><%=rsDET("Usuario_Estado_Ruta")%></td>
                  <td align="center"><img src='../imagenes/priorizar_normal.png' border='0' title="<%=TituloObs%>"></td>
				 </tr>
				 <%

				 rsDET.movenext
				 loop

				vArrConcepto = split(strArrConcepto,";")
				vArrID_CUOTA = split(strArrID_CUOTA,";")
				intTamvConcepto = ubound(vArrConcepto)

				 %>
				</tbody>

	</table>

	<table width="100%" border="0" bordercolor="#FFFFFF">
		<tr class="estilo_columna_individual" >
			<td>OBSERVACIONES (Max. 300 Caract.)</td>
		</tr>

		<tr>
		   	<td align="left">
		 		<TEXTAREA NAME="TX_OBSERVACIONES" ROWS="4" COLS="90"><%=strObsNueva%></TEXTAREA>
		  	</td>
		 </tr>

		<tr>
			<td class="estilo_columna_individual">&nbsp;</td>
		</tr>

	 </table>

			  <%end if
			  rsDET.close
			  set rsDET=nothing
		  Else
		  end if%>

	    </td>
	  </tr>

		<tr>
			<TD>

				<table width="100%" border="0" bordercolor="#FFFFFF">
							<tr bordercolor="#999999" class="Estilo8">
							<td align="left">
								<INPUT TYPE="BUTTON" NAME="Confirmar" class="fondo_boton_100" <%=strEsconder%>  value="Confirmar" onClick="envia('C');" class="Estilo8">
								<INPUT TYPE="BUTTON" NAME="DesConfirmar" class="fondo_boton_100" <%=strEsconder%>  value="DesConfirmar" onClick="envia('D');" class="Estilo8">
								<INPUT TYPE="BUTTON" NAME="Dejar Pendiente" class="fondo_boton_100" <%=strEsconder%>  value="Dejar Pendiente" onClick="envia('DP');" class="Estilo8">
							</td>
							<td align="right">
								<INPUT TYPE="BUTTON" NAME="Ver Gestiones" class="fondo_boton_100" value="Ver Gestiones" onClick="ir_detalle_gestiones();" class="Estilo8">
								<INPUT TYPE="BUTTON" NAME="Limpiar" class="fondo_boton_100"  <%=strEsconder%> value="Limpiar" onClick="LimpiarDatos();" class="Estilo8">
								<INPUT TYPE="BUTTON" NAME="Volver" class="fondo_boton_100" value="Volver" onClick="history.back();" class="Estilo8">
							</td>
							</tr>
				</table>
			</TD>
		</tr>

		<tr>
			<TD ALIGN="CENTER">
				<table width="100%" border="0" bordercolor="#FFFFFF" class="estilo_columnas">
					<thead>
					<tr>
						<td>FECHA AGENDAMIENTO</td>
						<td colspan = "2" align="left">HORA AGEND.</td>
					</tr>
				</thead>
					<tr>
						 <td width = "200" >
							<input name="TX_FEC_AGEND" type="text" id="TX_FEC_AGEND" value="<%=dtmFechaAgend%>" size="10" maxlength="10">
	
						 </td>
						 <td width="100" align="left">
							<input name="TX_HORA_AGEND" type="text" id="TX_HORA_AGEND" value="<%=dtmHoraAgend%>" size="5" maxlength="5" onChange="return ValidaHora(this,this.value)">
						</td>
		
					<% If TraeSiNo(session("perfil_adm"))="Si" or TraeSiNo(session("perfil_sup"))="Si" Then %>
						<td align="left" >
							<INPUT TYPE="BUTTON" class="fondo_boton_100" NAME="BT_AGENDAR" <%=strEsconder%> value="Agendar" onClick="Agendar();" class="Estilo8">
						</td>
					<% Else%>
						<td align="left" >&nbsp;</td>						
					<% End If %>

					</tr>
				</table>
			</TD>
		</tr>
	</table>


	</td>
	</tr>
<td>
</table>

<INPUT TYPE="hidden" NAME="strGraba" value="">
<INPUT TYPE="hidden" NAME="strTipo" value="">
<INPUT TYPE="hidden" NAME="strLimpiar" value="">
</form>
</body>
</html>

<script language="JavaScript" type="text/javascript">
$(document).ready(function(){
	$('#TX_FEC_AGEND').datepicker( {changeMonth: true,changeYear: true, dateFormat: 'dd/mm/yy'});
});

function envia(strTipo){


		if (strTipo == 'C') {
			if (confirm("¿ Está seguro de Confirmar los documentos seleccionados ?"))
				{			
					datos.strGraba.value='SI';
					datos.strTipo.value='C';
					datos.submit();
				}
		}
		if (strTipo == 'D') {
			if (confirm("¿ Está seguro de Desconfirmar los documentos seleccionados ?"))
				{	
					datos.strGraba.value='SI';
					datos.strTipo.value='D';
					datos.submit();
				}					
		}
		if (strTipo == 'DP') {

			datos.strGraba.value='SI';
			datos.strTipo.value='DP';
			datos.submit();
		}

}	
function LimpiarDatos(){
	if (confirm("¿ Está seguro de Borrar la observación ?"))
	{
	datos.strLimpiar.value='SI';
	datos.submit();
	}
}

function marcar_boxes(){
		<% For i=1 TO intTamvConcepto %>
			document.forms[0].<%=vArrConcepto(i)%>.checked=true;
		<% Next %>
}

function desmarcar_boxes(){
		<% For i=1 TO intTamvConcepto %>
			document.forms[0].<%=vArrConcepto(i)%>.checked=false;
		<% Next %>
}

function Agendar() {

	if (confirm("¿ Está seguro de agendar ? "))
	{
			datos.strAgendar.value='SI';
			datos.submit();
	}

}
function ir_detalle_gestiones(){

	datos.action="detalle_gestiones.asp?rut=<%=strRutDeudor%>&cliente=<%=strCodCliente%>";
	datos.submit();
}

  function MarcaChexbok() {
           var chkHeader = document.getElementById("chckHead");
           for (var i = 2; i < document.getElementById('tbl_Procesa').rows.length; i++)
            {
                chk = document.getElementById("ChckRow_" + i)
                chk.checked = chkHeader.checked
           }
        }

</script>