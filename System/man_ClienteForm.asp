<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html>
<html lang="es">
<HEAD>
    <meta http-equiv="X-UA-Compatible" content="IE=edge,chrome=1">
    <meta charset="utf-8">
    <link href="../css/normalize.css" rel="stylesheet">
	<!--#include file="sesion.asp"-->
	<!--#include file="arch_utils.asp"-->
    <!--#include file="../lib/freeaspupload.asp" -->
	<link href="../css/style_generales_sistema.css" rel="stylesheet">

	<style type="text/css" media="screen">
	.hdr_i{ 
		font-family: Tahoma, Helvetica, sans-serif; 
		font-size: 12px; 
		font-weight: bold;   
		background-color: #C9DEF2; 
		text-align: left ; 
		border: 1px solid #fff; 
	}

	.largo_texto input{
		width:200px;
	}
	.largo_texto{
		border-bottom: 1px solid #ccc;
	}
	.largo_texto select{
		width:205px;
	}	
	</style>
<%
	Response.CodePage=65001
	Response.charset ="utf-8"

	sintNuevo 		= request("sintNuevo")
    COD_CLIENTE		= request("COD_CLIENTE")
    archivo 		= Request("archivo")
    accion_archivo	= request("accion_archivo")
	ruta 			= Request("ruta") 

	


    AbrirSCG()


if trim(COD_CLIENTE)<>"" then

	sql_sel = "SELECT  COD_CLIENTE, DESCRIPCION, RAZON_SOCIAL, NOMBRE_FANTASIA, RUT, REP_LEGAL, EMAIL_CONTACTO, ACTIVO, TASA_MAX_CONV "
	sql_sel =  sql_sel & ", IC_PORC_CAPITAL, HON_PORC_CAPITAL, PIE_PORC_CAPITAL, TIPO_INTERES, TIPO_CLIENTE, GASTOS_OPERACIONALES "
	sql_sel =  sql_sel & ", GASTOS_ADMINISTRATIVOS, GASTOS_OPERACIONALES_CD, GASTOS_ADMINISTRATIVOS_CD, ADIC_1, ADIC_2, ADIC_3, ADIC_4 "
	sql_sel =  sql_sel & ", ADIC_5, ADIC_91, ADIC_92, ADIC_93, ADIC_94, ADIC_95, USA_CUSTODIO, COLOR_CUSTODIO, ADIC_96, ADIC_97, ADIC_98 "
	sql_sel =  sql_sel & ", ADIC_99, ADIC_100, NOMBRE_CONV_PAGARE, INTERES_MORA, EXPIRACION_CONVENIO, EXPIRACION_ANULACION, COD_MONEDA "
	sql_sel =  sql_sel & ", COD_TIPODOCUMENTO_HON, MESES_TD_HON, ADIC1_DEUDOR, ADIC2_DEUDOR, ADIC3_DEUDOR, RETIRO_SABADO, COD_ULT_GES "
	sql_sel =  sql_sel & ", OBS_ULT_GES, FORMULA_HONORARIOS, FORMULA_HONORARIOS_FACT, FORMULA_INTERESES, USA_HONORARIOS, USA_INTERESES "
	sql_sel =  sql_sel & ", USA_SUBCLIENTE, USA_REPLEGAL, USA_PROTESTOS, NRO_CLIENTE_DOC, NRO_CLIENTE_DEUDOR, DIRECCION, USA_COB_INTERNA "
	sql_sel =  sql_sel & " FROM  CLIENTE "
	sql_sel =  sql_sel & " WHERE COD_CLIENTE = " & TRIM(COD_CLIENTE)


	set rs_cliente = conn.execute(sql_sel)
    If Not rs_cliente.Eof Then
    	COD_CLIENTE 				=rs_cliente("COD_CLIENTE")
		DESCRIPCION 				=rs_cliente("DESCRIPCION")
		RAZON_SOCIAL 				=rs_cliente("RAZON_SOCIAL")		
		NOMBRE_FANTASIA 			=rs_cliente("NOMBRE_FANTASIA")		
		RUT 						=rs_cliente("RUT")		
		REP_LEGAL 					=rs_cliente("REP_LEGAL")		
		EMAIL_CONTACTO 				=rs_cliente("EMAIL_CONTACTO")		
		ACTIVO 						=rs_cliente("ACTIVO")		
		TASA_MAX_CONV 				=rs_cliente("TASA_MAX_CONV")		
		IC_PORC_CAPITAL 			=rs_cliente("IC_PORC_CAPITAL")		
		HON_PORC_CAPITAL 			=rs_cliente("HON_PORC_CAPITAL")		
		PIE_PORC_CAPITAL 			=rs_cliente("PIE_PORC_CAPITAL")		
		TIPO_INTERES 				=rs_cliente("TIPO_INTERES")		
		TIPO_CLIENTE 				=rs_cliente("TIPO_CLIENTE")		
		GASTOS_OPERACIONALES 		=rs_cliente("GASTOS_OPERACIONALES")		
		GASTOS_ADMINISTRATIVOS 		=rs_cliente("GASTOS_ADMINISTRATIVOS")		
		GASTOS_OPERACIONALES_CD 	=rs_cliente("GASTOS_OPERACIONALES_CD")		
		GASTOS_ADMINISTRATIVOS_CD 	=rs_cliente("GASTOS_ADMINISTRATIVOS_CD")		
		ADIC_1 						=rs_cliente("ADIC_1")		
		ADIC_2 						=rs_cliente("ADIC_2")		
		ADIC_3 						=rs_cliente("ADIC_3")		
		ADIC_4 						=rs_cliente("ADIC_4")		
		ADIC_5 						=rs_cliente("ADIC_5")		
		ADIC_91 					=rs_cliente("ADIC_91")		
		ADIC_92 					=rs_cliente("ADIC_92")		
		ADIC_93 					=rs_cliente("ADIC_93")		
		ADIC_94 					=rs_cliente("ADIC_94")		
		ADIC_95 					=rs_cliente("ADIC_95")		
		USA_CUSTODIO 				=rs_cliente("USA_CUSTODIO")		
		COLOR_CUSTODIO 				=rs_cliente("COLOR_CUSTODIO")		
		ADIC_96 					=rs_cliente("ADIC_96")		
		ADIC_97 					=rs_cliente("ADIC_97")		
		ADIC_98 					=rs_cliente("ADIC_98")		
		ADIC_99 					=rs_cliente("ADIC_99")		
		ADIC_100 					=rs_cliente("ADIC_100")				
		NOMBRE_CONV_PAGARE 			=rs_cliente("NOMBRE_CONV_PAGARE")		
		INTERES_MORA 				=rs_cliente("INTERES_MORA")		
		EXPIRACION_CONVENIO 		=rs_cliente("EXPIRACION_CONVENIO")		
		EXPIRACION_ANULACION 		=rs_cliente("EXPIRACION_ANULACION")		
		COD_MONEDA 					=rs_cliente("COD_MONEDA")		
		COD_TIPODOCUMENTO_HON 		=rs_cliente("COD_TIPODOCUMENTO_HON")		
		MESES_TD_HON 				=rs_cliente("MESES_TD_HON")		
		ADIC1_DEUDOR 				=rs_cliente("ADIC1_DEUDOR")		
		ADIC2_DEUDOR 				=rs_cliente("ADIC2_DEUDOR")		
		ADIC3_DEUDOR 				=rs_cliente("ADIC3_DEUDOR")		
		RETIRO_SABADO 				=rs_cliente("RETIRO_SABADO")		
		COD_ULT_GES 				=rs_cliente("COD_ULT_GES")		
		OBS_ULT_GES 				=rs_cliente("OBS_ULT_GES")		
		FORMULA_HONORARIOS 			=rs_cliente("FORMULA_HONORARIOS")		
		FORMULA_HONORARIOS_FACT 	=rs_cliente("FORMULA_HONORARIOS_FACT")		
		FORMULA_INTERESES 			=rs_cliente("FORMULA_INTERESES")		
		USA_HONORARIOS 				=rs_cliente("USA_HONORARIOS")		
		USA_INTERESES 				=rs_cliente("USA_INTERESES")		
		USA_SUBCLIENTE 				=rs_cliente("USA_SUBCLIENTE")		
		USA_REPLEGAL 				=rs_cliente("USA_REPLEGAL")		
		USA_PROTESTOS 				=rs_cliente("USA_PROTESTOS")		
		NRO_CLIENTE_DOC 			=rs_cliente("NRO_CLIENTE_DOC")		
		NRO_CLIENTE_DEUDOR 			=rs_cliente("NRO_CLIENTE_DEUDOR")			
		DIRECCION					=rs_cliente("DIRECCION")		 
		USA_COB_INTERNA 			=rs_cliente("USA_COB_INTERNA")		

	else

		DESCRIPCION 				=""
		RAZON_SOCIAL 				=""	
		NOMBRE_FANTASIA 			=""	
		RUT 						=""		
		REP_LEGAL 					=""		
		EMAIL_CONTACTO 				=""		
		ACTIVO 						="1"		
		TASA_MAX_CONV 				="0"		
		IC_PORC_CAPITAL 			="0"		
		HON_PORC_CAPITAL 			="0"	
		PIE_PORC_CAPITAL 			="0"	
		TIPO_INTERES 				=""	
		TIPO_CLIENTE 				=""		
		GASTOS_OPERACIONALES 		="0"		
		GASTOS_ADMINISTRATIVOS 		="0"		
		GASTOS_OPERACIONALES_CD 	="0"	
		GASTOS_ADMINISTRATIVOS_CD 	="0"		
		ADIC_1 						=""		
		ADIC_2 						=""	
		ADIC_3 						=""		
		ADIC_4 						=""		
		ADIC_5 						=""		
		ADIC_91 					=""		
		ADIC_92 					=""	
		ADIC_93 					=""		
		ADIC_94 					=""		
		ADIC_95 					=""	
		USA_CUSTODIO 				=""		
		COLOR_CUSTODIO 				=""	
		ADIC_96 					=""		
		ADIC_97 					=""	
		ADIC_98 					=""	
		ADIC_99 					=""	
		ADIC_100 					=""				
		NOMBRE_CONV_PAGARE 			=""		
		INTERES_MORA 				="0"		
		EXPIRACION_CONVENIO 		="0"		
		EXPIRACION_ANULACION 		="0"	
		COD_MONEDA 					=""		
		COD_TIPODOCUMENTO_HON 		=""		
		MESES_TD_HON 				="0"	
		ADIC1_DEUDOR 				=""	
		ADIC2_DEUDOR 				=""	
		ADIC3_DEUDOR 				=""	
		RETIRO_SABADO 				="0"		
		COD_ULT_GES 				=""	
		OBS_ULT_GES 				=""		
		FORMULA_HONORARIOS 			=""		
		FORMULA_HONORARIOS_FACT 	=""		
		FORMULA_INTERESES 			=""		
		USA_HONORARIOS 				="0"		
		USA_INTERESES 				="0"		
		USA_SUBCLIENTE 				=""		
		USA_REPLEGAL 				=""	
		USA_PROTESTOS 				=""	
		NRO_CLIENTE_DOC 			=""	
		NRO_CLIENTE_DEUDOR 			=""			
		DIRECCION					=""		 
		USA_COB_INTERNA 			=""	


    End If

else

		DESCRIPCION 				=""
		RAZON_SOCIAL 				=""	
		NOMBRE_FANTASIA 			=""	
		RUT 						=""		
		REP_LEGAL 					=""		
		EMAIL_CONTACTO 				=""		
		ACTIVO 						="1"		
		TASA_MAX_CONV 				="0"		
		IC_PORC_CAPITAL 			="0"		
		HON_PORC_CAPITAL 			="0"	
		PIE_PORC_CAPITAL 			="0"	
		TIPO_INTERES 				=""	
		TIPO_CLIENTE 				=""		
		GASTOS_OPERACIONALES 		="0"		
		GASTOS_ADMINISTRATIVOS 		="0"		
		GASTOS_OPERACIONALES_CD 	="0"	
		GASTOS_ADMINISTRATIVOS_CD 	="0"		
		ADIC_1 						=""		
		ADIC_2 						=""	
		ADIC_3 						=""		
		ADIC_4 						=""		
		ADIC_5 						=""		
		ADIC_91 					=""		
		ADIC_92 					=""	
		ADIC_93 					=""		
		ADIC_94 					=""		
		ADIC_95 					=""	
		USA_CUSTODIO 				=""		
		COLOR_CUSTODIO 				=""	
		ADIC_96 					=""		
		ADIC_97 					=""	
		ADIC_98 					=""	
		ADIC_99 					=""	
		ADIC_100 					=""				
		NOMBRE_CONV_PAGARE 			=""		
		INTERES_MORA 				="0"		
		EXPIRACION_CONVENIO 		="0"		
		EXPIRACION_ANULACION 		="0"	
		COD_MONEDA 					=""		
		COD_TIPODOCUMENTO_HON 		=""		
		MESES_TD_HON 				="0"	
		ADIC1_DEUDOR 				=""	
		ADIC2_DEUDOR 				=""	
		ADIC3_DEUDOR 				=""	
		RETIRO_SABADO 				="0"		
		COD_ULT_GES 				=""	
		OBS_ULT_GES 				=""		
		FORMULA_HONORARIOS 			=""		
		FORMULA_HONORARIOS_FACT 	=""		
		FORMULA_INTERESES 			=""		
		USA_HONORARIOS 				="0"		
		USA_INTERESES 				="0"		
		USA_SUBCLIENTE 				=""		
		USA_REPLEGAL 				=""	
		USA_PROTESTOS 				=""	
		NRO_CLIENTE_DOC 			=""	
		NRO_CLIENTE_DEUDOR 			=""			
		DIRECCION					=""		 
		USA_COB_INTERNA 			=""	
end if

IntId = COD_CLIENTE 

    Dim DestinationPath
		DestinationPath = Server.mapPath("../Archivo/BibliotecaClientes") & "\" & IntId

	
		Dim uploadsDirVar
		uploadsDirVar = DestinationPath

       
       Dim Upload, fileName, fileSize, ks, i, fileKey, resumen
			
			
		function SaveFiles
           
			Dim Upload, fileName, fileSize, ks, i, fileKey, resumen
			Set Upload = New FreeASPUpload
			Upload.Save(uploadsDirVar)
			If Err.Number <> 0 then Exit function
			SaveFiles = ""
			ks = Upload.UploadedFiles.keys
			if (UBound(ks) <> -1) then
				resumen = "<B>Archivos subidos:</B> "
				for each fileKey in Upload.UploadedFiles.keys
					resumen = resumen & Upload.UploadedFiles(fileKey).FileName & " (" & Upload.UploadedFiles(fileKey).Length & "B) "
				    archivo = Upload.UploadedFiles(fileKey).FileName
				next
             else
			End if

             

		End function

       

	If Request.ServerVariables("REQUEST_METHOD") = "POST" and trim(accion_archivo) <> "descarga" then
  
  		response.write SaveFiles()

		strSql = "EXEC Proc_Audita_Archivo 1, 1, "&trim(session("session_idusuario"))&",null, '"&trim(IntId)&"', '"&trim(archivo)&"', '',0 "
		
		Conn.execute(strSql)



	End if

	if Request.ServerVariables("REQUEST_METHOD") = "POST" and trim(accion_archivo) = "descarga" then
        
        response.write DownloadFile(ruta)
		response.write ruta&""

	End if	

	Set Obj_FSO = createobject("scripting.filesystemobject")


	If not Obj_FSO.FolderExists(Server.mapPath("../Archivo/BibliotecaClientes") & "\" & IntId) = True Then ' verifica la existencia del archivo
		Obj_FSO.CreateFolder(Server.mapPath("../Archivo/BibliotecaClientes") & "\" & IntId) 
	End if


%>


<script src="../Componentes/jquery-1.9.2/js/jquery-1.8.3.js"></script>
</HEAD>

<BODY BGCOLOR='FFFFFF'>

<INPUT TYPE="hidden" NAME="sintNuevo" id="sintNuevo" VALUE="<%= sintNuevo %>">
<div class="titulo_informe">MANTENEDOR DE CLIENTES</div>
<br>
		<table width="90%" border="0" align="center">
			<tr >
				<td class="subtitulo_informe">> INGRESO CLIENTES</td>
			</tr>
		</Table>

		<table width="90%" border="0" CLASS="tabla1" align="center">
			<tr BGCOLOR="#FFFFFF">
				<td class="hdr_i">Codigo</td>
				<td class="td_t largo_texto">
					<%if trim(COD_CLIENTE)<>"" then%>
						<%=trim(COD_CLIENTE)%>
						<input type="hidden" name="COD_CLIENTE" id="COD_CLIENTE" value="<%=trim(COD_CLIENTE)%>">

					<%else%>
						<input type="text" name="COD_CLIENTE" id="COD_CLIENTE" value="<%=trim(COD_CLIENTE)%>">
					<%end if%>
				</td>
		        <td class="hdr_i">Tipo Cliente</td>
				<td class="td_t largo_texto">
					<select name="TIPO_CLIENTE" id="TIPO_CLIENTE" onChange="">
						<option value="JUDICIAL" <%=strTipoClienteJU%>>Judicial</option>
						<option value="EXTRAJUD" <%=strTipoClienteEJ%>>Extra Judicial</option>
					</select>
			</TR>
			<tr BGCOLOR="#FFFFFF">
				<td class="hdr_i">Nombre</td>
				<td class="td_t largo_texto"><input type="text" name="DESCRIPCION" id="DESCRIPCION" value="<%=trim(DESCRIPCION)%>"></td>
				<td class="hdr_i">Razon Social</td>
				<td class="td_t largo_texto"><input type="text" name="RAZON_SOCIAL" id="RAZON_SOCIAL" value="<%=trim(RAZON_SOCIAL)%>"></td>
			</TR>
			<tr BGCOLOR="#FFFFFF">
				<td class="hdr_i">Rut</td>
				<td class="td_t largo_texto"><input type="text" name="RUT" id="RUT" value="<%=trim(RUT)%>"></td>
				<td class="hdr_i">Nombre Fantasia</td>
				<td class="td_t largo_texto"><input type="text" name="NOMBRE_FANTASIA" id="NOMBRE_FANTASIA" value="<%=trim(NOMBRE_FANTASIA)%>"></td>
			</TR>
			<tr BGCOLOR="#FFFFFF">
				<td class="hdr_i">Email Contacto</td>
				<td class="td_t largo_texto"><input type="text" name="EMAIL_CONTACTO" id="EMAIL_CONTACTO" value="<%=trim(EMAIL_CONTACTO)%>"></td>
				<td class="hdr_i">&nbsp;</td>
				<td class="td_t largo_texto">&nbsp;</td>
			</TR>
			<tr BGCOLOR="#FFFFFF">
				<td class="hdr_i">Imdem.Comp.</td>
				<td class="td_t largo_texto"><input type="text" name="IC_PORC_CAPITAL" id="IC_PORC_CAPITAL" value="<%=trim(IC_PORC_CAPITAL)%>"></td>
				<td class="hdr_i">Honorarios</td>
				<td class="td_t largo_texto"><input type="text" name="HON_PORC_CAPITAL" id="HON_PORC_CAPITAL" value="<%=trim(HON_PORC_CAPITAL)%>">(% Capital)</td>
			</TR>
			<tr BGCOLOR="#FFFFFF">
				<td class="hdr_i">Pie Convenio</td>
				<td class="td_t largo_texto"><input type="text" name="PIE_PORC_CAPITAL" id="PIE_PORC_CAPITAL" value="<%=trim(PIE_PORC_CAPITAL)%>">(% Capital)</td>
				<td class="hdr_i">Interes Mora</td>
				<td class="td_t largo_texto" colspan="3"><input type="text" name="INTERES_MORA" id="INTERES_MORA" value="<%=trim(INTERES_MORA)%>">(Mensual)</td>
			</TR>
			<tr BGCOLOR="#FFFFFF">
				<td class="hdr_i">Interes Futuro</td>
				<td class="td_t largo_texto" ><input type="text" name="TASA_MAX_CONV" id="TASA_MAX_CONV" value="<%=trim(TASA_MAX_CONV)%>">(Mensual)</td>
				<td class="hdr_i">Tipo Int.Futuro</td>
				<td class="td_t largo_texto"><input type="text" name="TIPO_INTERES" id="TIPO_INTERES" value="<%=trim(TIPO_INTERES)%>"> (C:Comp., S:Simple)</td>
			</TR>
			<tr BGCOLOR="#FFFFFF">
				<td class="hdr_i">Expiracion Convenio</td>
				<td class="td_t largo_texto" ><input type="text" name="EXPIRACION_CONVENIO" id="EXPIRACION_CONVENIO" value="<%=trim(EXPIRACION_CONVENIO)%>">(Dias)</td>
				<td class="hdr_i">Expiracion Anulacion</td>
				<td class="td_t largo_texto"><input type="text" name="EXPIRACION_ANULACION" id="EXPIRACION_ANULACION" value="<%=trim(EXPIRACION_ANULACION)%>">(Dias)</td>
			</TR>
			<tr BGCOLOR="#FFFFFF">
				<td class="hdr_i">Con Custodio</td>
				<td class="td_t largo_texto"><input type="text" name="USA_CUSTODIO" id="USA_CUSTODIO" value="<%=trim(USA_CUSTODIO)%>">(S:Si, N:No)</td>
				<td class="hdr_i">Color Custodio</td>
				<td class="td_t largo_texto"><input type="text" name="COLOR_CUSTODIO" id="COLOR_CUSTODIO" value="<%=trim(COLOR_CUSTODIO)%>"></td>
			</TR>
			<tr BGCOLOR="#FFFFFF">
				<td class="hdr_i">Gastos Op. S/Dem.</td>
				<td class="td_t largo_texto"><input type="text" name="GASTOS_OPERACIONALES" id="GASTOS_OPERACIONALES" value="<%=trim(GASTOS_OPERACIONALES)%>"></td>
				<td class="hdr_i">Gastos Adm. S/Dem.</td>
				<td class="td_t largo_texto"><input type="text" name="GASTOS_ADMINISTRATIVOS" id="GASTOS_ADMINISTRATIVOS" value="<%=trim(GASTOS_ADMINISTRATIVOS)%>"></td>
			</TR>
			<tr BGCOLOR="#FFFFFF">
				<td class="hdr_i">Gastos Op. C/Dem.</td>
				<td class="td_t largo_texto"><input type="text" name="GASTOS_OPERACIONALES_CD" id="GASTOS_OPERACIONALES_CD" value="<%=trim(GASTOS_OPERACIONALES_CD)%>"></td>
				<td class="hdr_i">Gastos Adm. C/Dem.</td>
				<td class="td_t largo_texto"><input type="text" name="GASTOS_ADMINISTRATIVOS_CD" id="GASTOS_ADMINISTRATIVOS_CD" value="<%=trim(GASTOS_ADMINISTRATIVOS_CD)%>"></td>
			</TR>
			<tr BGCOLOR="#FFFFFF">
				<td class="hdr_i">Adic 1</td>
				<td class="td_t largo_texto"><input type="text" name="ADIC_1" id="ADIC_1" value="<%=trim(ADIC_1)%>"></td>
				<td class="hdr_i">Adic 2</td>
				<td class="td_t largo_texto"><input type="text" name="ADIC_2" id="ADIC_2" value="<%=trim(ADIC_2)%>"></td>
			</TR>
			<tr BGCOLOR="#FFFFFF">
				<td class="hdr_i">Adic 3</td>
				<td class="td_t largo_texto"><input type="text" name="ADIC_3" id="ADIC_3" value="<%=trim(ADIC_3)%>"></td>
				<td class="hdr_i">Adic 4</td>
				<td class="td_t largo_texto"><input type="text" name="ADIC_4" id="ADIC_4" value="<%=trim(ADIC_4)%>"></td>
			</TR>
			<tr BGCOLOR="#FFFFFF">
				<td class="hdr_i">Adic 5</td>
				<td class="td_t largo_texto"><input type="text" name="ADIC_5" id="ADIC_5" value="<%=trim(ADIC_5)%>"></td>
				<td class="hdr_i">Adic 91</td>
				<td class="td_t largo_texto"><input type="text" name="ADIC_91" id="ADIC_91" value="<%=trim(ADIC_91)%>"></td>
			</TR>
			<tr BGCOLOR="#FFFFFF">
				<td class="hdr_i">Adic 92</td>
				<td class="td_t largo_texto"><input type="text" name="ADIC_92" id="ADIC_92" value="<%=trim(ADIC_92)%>"></td>
				<td class="hdr_i">Adic 93</td>
				<td class="td_t largo_texto"><input type="text" name="ADIC_93" id="ADIC_93" value="<%=trim(ADIC_93)%>"></td>
			</TR>
			<tr BGCOLOR="#FFFFFF">
				<td class="hdr_i">Adic 94</td>
				<td class="td_t largo_texto"><input type="text" name="ADIC_94" id="ADIC_94" value="<%=trim(ADIC_94)%>"></td>
				<td class="hdr_i">Adic 95</td>
				<td class="td_t largo_texto"><input type="text" name="ADIC_95" id="ADIC_95" value="<%=trim(ADIC_95)%>"></td>
			</TR>
			<tr BGCOLOR="#FFFFFF">
				<td class="hdr_i">Nombre Conv-Pagare</td>
				<td class="td_t largo_texto"><input type="text" name="NOMBRE_CONV_PAGARE" id="NOMBRE_CONV_PAGARE" value="<%=trim(NOMBRE_CONV_PAGARE)%>"></td>

				<td class="hdr_i">Meses Calc.Hon.</td>
				<td class="td_t largo_texto"><input type="text" name="MESES_TD_HON" id="MESES_TD_HON" value="<%=trim(MESES_TD_HON)%>"></td>
			</TR>
			<tr BGCOLOR="#FFFFFF">
				<td class="hdr_i">Tipo Moneda</td>
				<td class="td_t largo_texto ">
					<SELECT NAME="COD_MONEDA" id="COD_MONEDA" onChange="">

					<%
					strSql="SELECT * FROM MONEDA"
					set rsTemp= Conn.execute(strSql)
					if not rsTemp.eof then
						do until rsTemp.eof%>
						  <option value="<%=rsTemp("COD_MONEDA")%>" <% if Trim(rsTemp("COD_MONEDA")) = intCodMoneda then Response.Write "SELECTED" %> ><%=rsTemp("NOM_MONEDA")%></option>
						  <%
						rsTemp.movenext
						loop
					end if
					rsTemp.close
					set rsTemp=nothing

					%>
					  </SELECT>
				</td>
				<td class="hdr_i">Tipo Doc.Hon.</td>
				<td class="td_t largo_texto">
					<SELECT NAME="COD_TIPODOCUMENTO_HON" id="COD_TIPODOCUMENTO_HON" onChange="">
					<option value="" <% if Trim(intCodTipoDoc) = "" then Response.Write "SELECTED" %> >SIN TIPO DOC</option>

					<%
					strSql="SELECT * FROM TIPO_DOCUMENTO"
					set rsTemp= Conn.execute(strSql)
					if not rsTemp.eof then
					do until rsTemp.eof%>
					  <option value="<%=rsTemp("COD_TIPO_DOCUMENTO")%>" <% if Trim(rsTemp("COD_TIPO_DOCUMENTO")) = intCodTipoDoc then Response.Write "SELECTED" %> ><%=rsTemp("NOM_TIPO_DOCUMENTO")%></option>
					  <%
					rsTemp.movenext
					loop
					end if
					rsTemp.close
					set rsTemp=nothing

					%>
					</SELECT>
				</td>
			</tr>

			<tr BGCOLOR="#FFFFFF">
				<td class="hdr_i">Adic 1 Deudor</td>
				<td class="td_t largo_texto"><input type="text" name="ADIC1_DEUDOR" id="ADIC1_DEUDOR" value="<%=trim(ADIC1_DEUDOR)%>"></td>
				<td class="hdr_i">Adic 2 Deudor</td>
				<td class="td_t largo_texto"><input type="text" name="ADIC2_DEUDOR" id="ADIC2_DEUDOR" value="<%=trim(ADIC2_DEUDOR)%>"></td>
			</TR>
			<tr BGCOLOR="#FFFFFF">
				<td class="hdr_i">Adic 3 Deudor</td>
				<td class="td_t largo_texto"><input type="text" name="ADIC3_DEUDOR" id="ADIC3_DEUDOR" value="<%=trim(ADIC3_DEUDOR)%>"></td>
				<td class="hdr_i">&nbsp;</td>
				<td class="td_t largo_texto">&nbsp;</td>
			</TR>

			<tr BGCOLOR="#FFFFFF">
				<td class="hdr_i">Retiro/Agend Sabado</td>
				<td class="td_t largo_texto"><input type="text" name="RETIRO_SABADO" id="RETIRO_SABADO" value="<%=trim(RETIRO_SABADO)%>"></td>
				<td class="hdr_i">Activo</td>
				<td class="td_t ">
					<input type="radio" name="ACTIVO" id="ACTIVO" <%if trim(ACTIVO)=true then response.write " CHECKED " end if%> value="1">SI
					<input type="radio" name="ACTIVO" id="ACTIVO" <%if trim(ACTIVO)=false then response.write " CHECKED " end if%> value="0">NO

				</td>
			</tr>
			<tr BGCOLOR="#FFFFFF">
				<td class="hdr_i">Usa Honorarios</td>
				<td class="td_t largo_texto"><input type="text" name="USA_HONORARIOS" id="USA_HONORARIOS" value="<%=trim(USA_HONORARIOS)%>"></td>
				<td class="hdr_i">Usa Intereses </td>
				<td class="td_t largo_texto"><input type="text" name="USA_INTERESES" id="USA_INTERESES" value="<%=trim(USA_INTERESES)%>"></td>
			</tr>
			<tr BGCOLOR="#FFFFFF">
				<td class="hdr_i">Formula Intereses</td>
				<td class="td_t largo_texto"><input type="text" name="FORMULA_INTERESES" id="FORMULA_INTERESES" value="<%=trim(FORMULA_INTERESES)%>"></td>
				<td class="hdr_i">Formula Honorarios</td>
				<td class="td_t largo_texto"><input type="text" name="FORMULA_HONORARIOS" id="FORMULA_HONORARIOS" value="<%=trim(FORMULA_HONORARIOS)%>"></td>
			</tr>
			<tr BGCOLOR="#FFFFFF">
				<td class="hdr_i">Formula Honorarios Fact.</td>
				<td class="td_t largo_texto"><input type="text" name="FORMULA_HONORARIOS_FACT" id="FORMULA_HONORARIOS_FACT" value="<%=trim(FORMULA_HONORARIOS_FACT)%>"></td>
				<td class="hdr_i">&nbsp;</td>
				<td class="td_t largo_texto">&nbsp;</td>

			</tr>

		</table>

		<table width="90%" border="0" align="center">
		     <TR>
			  <td align="right"  width="25%">
			   <INPUT TYPE="BUTTON" value="Guardar" class="fondo_boton_100" name="B1" onClick="Continuar()">
			   <input type="BUTTON" value="Terminar" class="fondo_boton_100" name="terminar" onClick="Terminar('man_Cliente.asp');return false;"></TD>
			  </TD>
		    </TR>
		    <%If sintNuevo = 1 Then %>
		    <TR>
		    <TD align="center">
		     <IMG BORDER="0" src="../imagenes/bolita.jpg" WIDTH=10>=Campo requerido
		     </TD>
		    </TR>
			<%End If %>
		</table>

		<div id="guarda_cliente"></div>

	<%if trim(COD_CLIENTE)<>"" then%>
	<FORM name="frmSend" id="frmSend" onSubmit="return enviar(this)"  method="POST" enctype="multipart/form-data" action="man_ClienteForm.asp">

	<table width="90%" border="0" CLASS="tabla1" align="center">
	<tr BGCOLOR="#FFFFFF">
		<td height="45">Archivo: </td>
		<td align=""><input name="File1" id="File1" type="file" VALUE="<%= File1%>" size="40" maxlength="40">&nbsp; <input Name="SubmitButton" Value="Cargar" class="fondo_boton_100" Type="BUTTON" onClick="enviar();"></td>
	</tr>

	<%

		SQL_SEL ="SELECT id_archivo, nombre_archivo, cod_cliente, rut  " & _
				"FROM CARGA_ARCHIVOS " & _
				"WHERE activo =1 AND cod_cliente="&trim(COD_CLIENTE)  & _
				" AND origen = 1 "
		set rs_sql_sel = Conn.execute(SQL_SEL)

		do while not rs_sql_sel.eof
	%>
			<tr BGCOLOR="#FFFFFF">
				<td>
					<a href="#" onclick="bt_descargar('../Archivo/BibliotecaClientes/<%=trim(rs_sql_sel("cod_cliente"))%>/<%=trim(rs_sql_sel("nombre_archivo"))%>')"><%=trim(rs_sql_sel("nombre_archivo"))%></a>	
				</td>
				<td align="center">
					<a onclick="bt_eliminar('<%=trim(rs_sql_sel("cod_cliente"))%>','<%=trim(rs_sql_sel("nombre_archivo"))%>','biblioteca_cliente','<%=trim(rs_sql_sel("id_archivo"))%>')" href="#">Eliminar</a>
				</td>
			</tr>
	<%

		rs_sql_sel.movenext 
		loop

	%>
	</table>
	</FORM>
	<%end if %>
	<div id="verifica_archivo"></div>
	<div id="verifica_cliente"></div>


<script type="text/javascript">
function Continuar(){
	var sintNuevo 					=$('#sintNuevo').val()
	var COD_CLIENTE 				=$('#COD_CLIENTE').val()
	var TIPO_CLIENTE				=$('#TIPO_CLIENTE').val()
	var DESCRIPCION					=$('#DESCRIPCION').val()
	var RAZON_SOCIAL				=$('#RAZON_SOCIAL').val()
	var RUT							=$('#RUT').val()
	var NOMBRE_FANTASIA				=$('#NOMBRE_FANTASIA').val()
	var EMAIL_CONTACTO				=$('#EMAIL_CONTACTO').val()
	var IC_PORC_CAPITAL				=$('#IC_PORC_CAPITAL').val()
	var HON_PORC_CAPITAL			=$('#HON_PORC_CAPITAL').val()
	var PIE_PORC_CAPITAL			=$('#PIE_PORC_CAPITAL').val()
	var INTERES_MORA				=$('#INTERES_MORA').val()
	var TASA_MAX_CONV				=$('#TASA_MAX_CONV').val()
	var TIPO_INTERES				=$('#TIPO_INTERES').val()
	var EXPIRACION_CONVENIO			=$('#EXPIRACION_CONVENIO').val()
	var EXPIRACION_ANULACION		=$('#EXPIRACION_ANULACION').val()
	var USA_CUSTODIO				=$('#USA_CUSTODIO').val()
	var COLOR_CUSTODIO				=$('#COLOR_CUSTODIO').val()
	var GASTOS_OPERACIONALES		=$('#GASTOS_OPERACIONALES').val()
	var GASTOS_ADMINISTRATIVOS		=$('#GASTOS_ADMINISTRATIVOS').val()
	var GASTOS_OPERACIONALES_CD		=$('#GASTOS_OPERACIONALES_CD').val()
	var GASTOS_ADMINISTRATIVOS_CD	=$('#GASTOS_ADMINISTRATIVOS_CD').val()
	var ADIC_1						=$('#ADIC_1').val()
	var ADIC_2						=$('#ADIC_2').val()
	var ADIC_3						=$('#ADIC_3').val()
	var ADIC_4						=$('#ADIC_4').val()
	var ADIC_5						=$('#ADIC_5').val()
	var ADIC_91						=$('#ADIC_91').val()
	var ADIC_92						=$('#ADIC_92').val()
	var ADIC_93						=$('#ADIC_93').val()
	var ADIC_94						=$('#ADIC_94').val()
	var ADIC_95						=$('#ADIC_95').val()
	var NOMBRE_CONV_PAGARE			=$('#NOMBRE_CONV_PAGARE').val()
	var MESES_TD_HON				=$('#MESES_TD_HON').val()
	var COD_MONEDA					=$('#COD_MONEDA').val()
	var COD_TIPODOCUMENTO_HON		=$('#COD_TIPODOCUMENTO_HON').val()
	var ADIC1_DEUDOR				=$('#ADIC1_DEUDOR').val()
	var ADIC2_DEUDOR				=$('#ADIC2_DEUDOR').val()
	var ADIC3_DEUDOR				=$('#ADIC3_DEUDOR').val()
	var RETIRO_SABADO				=$('#RETIRO_SABADO').val()
	var ACTIVO						=$('#ACTIVO').val()
	var USA_HONORARIOS				=$('#USA_HONORARIOS').val()
	var USA_INTERESES				=$('#USA_INTERESES').val()
	var FORMULA_INTERESES			=$('#FORMULA_INTERESES').val()
	var FORMULA_HONORARIOS			=$('#FORMULA_HONORARIOS').val()
	var FORMULA_HONORARIOS_FACT		=$('#FORMULA_HONORARIOS_FACT').val()
	
	if(sintNuevo!="0"){
		var accion_ajax ="guarda_cliente"
	}else{
		var accion_ajax ="update_cliente"
	}

	var criterios ="alea="+Math.random()+"&accion_ajax="+accion_ajax+"&COD_CLIENTE="+COD_CLIENTE+"&TIPO_CLIENTE="+encodeURIComponent(TIPO_CLIENTE)+"&DESCRIPCION="+encodeURIComponent(DESCRIPCION)+"&RAZON_SOCIAL="+encodeURIComponent(RAZON_SOCIAL)+"&RUT="+encodeURIComponent(RUT)+"&NOMBRE_FANTASIA="+encodeURIComponent(NOMBRE_FANTASIA)+"&EMAIL_CONTACTO="+encodeURIComponent(EMAIL_CONTACTO)+"&IC_PORC_CAPITAL="+encodeURIComponent(IC_PORC_CAPITAL)+"&HON_PORC_CAPITAL="+encodeURIComponent(HON_PORC_CAPITAL)+"&PIE_PORC_CAPITAL="+encodeURIComponent(PIE_PORC_CAPITAL)+"&INTERES_MORA="+encodeURIComponent(INTERES_MORA)+"&TASA_MAX_CONV="+encodeURIComponent(TASA_MAX_CONV)+"&TIPO_INTERES="+encodeURIComponent(TIPO_INTERES)+"&EXPIRACION_CONVENIO="+encodeURIComponent(EXPIRACION_CONVENIO)+"&EXPIRACION_ANULACION="+encodeURIComponent(EXPIRACION_ANULACION)+"&USA_CUSTODIO="+encodeURIComponent(USA_CUSTODIO)+"&COLOR_CUSTODIO="+encodeURIComponent(COLOR_CUSTODIO)+"&GASTOS_OPERACIONALES="+encodeURIComponent(GASTOS_OPERACIONALES)+"&GASTOS_ADMINISTRATIVOS="+encodeURIComponent(GASTOS_ADMINISTRATIVOS)+"&GASTOS_OPERACIONALES_CD="+encodeURIComponent(GASTOS_OPERACIONALES_CD)+"&GASTOS_ADMINISTRATIVOS_CD="+encodeURIComponent(GASTOS_ADMINISTRATIVOS_CD)+"&ADIC_1="+encodeURIComponent(ADIC_1)+"&ADIC_2="+encodeURIComponent(ADIC_2)+"&ADIC_3="+encodeURIComponent(ADIC_3)+"&ADIC_4="+encodeURIComponent(ADIC_4)+"&ADIC_5="+encodeURIComponent(ADIC_5)+"&ADIC_91="+encodeURIComponent(ADIC_91)+"&ADIC_92="+encodeURIComponent(ADIC_92)+"&ADIC_93="+encodeURIComponent(ADIC_93)+"&ADIC_94="+encodeURIComponent(ADIC_94)+"&ADIC_95="+encodeURIComponent(ADIC_95)+"&NOMBRE_CONV_PAGARE="+encodeURIComponent(NOMBRE_CONV_PAGARE)+"&MESES_TD_HON="+encodeURIComponent(MESES_TD_HON)+"&COD_MONEDA="+encodeURIComponent(COD_MONEDA)+"&COD_TIPODOCUMENTO_HON="+encodeURIComponent(COD_TIPODOCUMENTO_HON)+"&ADIC1_DEUDOR="+encodeURIComponent(ADIC1_DEUDOR)+"&ADIC2_DEUDOR="+encodeURIComponent(ADIC2_DEUDOR)+"&ADIC3_DEUDOR="+encodeURIComponent(ADIC3_DEUDOR)+"&RETIRO_SABADO="+encodeURIComponent(RETIRO_SABADO)+"&ACTIVO="+encodeURIComponent(ACTIVO)+"&USA_HONORARIOS="+encodeURIComponent(USA_HONORARIOS)+"&USA_INTERESES="+encodeURIComponent(USA_INTERESES)+"&FORMULA_INTERESES="+encodeURIComponent(FORMULA_INTERESES)+"&FORMULA_HONORARIOS="+encodeURIComponent(FORMULA_HONORARIOS)+"&FORMULA_HONORARIOS_FACT="+encodeURIComponent(FORMULA_HONORARIOS_FACT)

	if(sintNuevo!="0"){

			var criterios_ver ="alea="+Math.random()+"&accion_ajax=verifica_cliente&COD_CLIENTE="+COD_CLIENTE
			$('#verifica_cliente').load('FuncionesAjax/man_ClienteForm_ajax.asp', criterios_ver, function(){
				var valida_cliente =$('#valida_cliente').val()
				if(valida_cliente=="S")
				{
					alert("Codigo cliente ya existe")
					return
				}else{
					$('#guarda_cliente').load('FuncionesAjax/man_ClienteForm_ajax.asp', criterios, function(){})
				}					

			})

	}else{

		$('#guarda_cliente').load('FuncionesAjax/man_ClienteForm_ajax.asp', criterios, function(){})
		
	}	
}

function bt_eliminar(cod_cliente, nombre_archivo, pagina_origen, id_archivo)
{
	if(confirm("¿Esta seguro que desea eliminar el archivo, posterior a esta acción no podrá recuperarlo?"))
	{
		location.href="EliminarArchivo.asp?IntId="+cod_cliente+"&VarNombreFichero="+nombre_archivo+"&pagina_origen="+pagina_origen+"&id_archivo="+id_archivo
	}
	
	
}



function bt_descargar(ruta){

    window.open(ruta, "INFORMACION", "width=800, height=400, scrollbars=yes, menubar=no, location=no, resizable=yes");
     
    
}

function Refrescar(strTipo){

	mantenedorForm.action='man_ClienteForm.asp?strTipoPropiedad=' + strTipo;
	if (mantenedorForm.strFormMode.value == 'Nuevo') {
		location.href="man_ClienteForm.asp?sintNuevo=1&strTipoPropiedad=" + strTipo;
		}
	else {
	 	mantenedorForm.submit();
	}
}


function enviar(){

	var File1 = $('#File1').val()
	var IntId = $('#COD_CLIENTE').val() 

	if(File1=="")
	{
		alert("¡Debe seleccionar archivo!")
		return
	}
	
	var vec = File1.split("\\");
	var cont = 0
	for(i=0;i<(vec.length);i++)
		{
			cont = cont + 1
		}

	var archivo 	=vec[cont-1]
	var extension 	=archivo.split(".");
	var contEx 		=0			

	for(i=0;i<(extension.length);i++)
	{
		contEx = contEx + 1
	}
		
	var nombre_archivo = extension[contEx-2]
	var extension_archivo = extension[contEx-1]

	var archivo=archivo.replace(",","");
	var archivo=archivo.replace("Á","");
	var archivo=archivo.replace("É","");
	var archivo=archivo.replace("Í","");
	var archivo=archivo.replace("Ó","");
	var archivo=archivo.replace("Ú","");
	var archivo=archivo.replace("á","");
	var archivo=archivo.replace("é","");
	var archivo=archivo.replace("í","");
	var archivo=archivo.replace("ó","");
	var archivo=archivo.replace("ú","");
	var archivo=archivo.replace("ñ","");
	var archivo=archivo.replace("Ñ","");	

	var criterios ="alea="+Math.random()+"&accion_ajax=verifica_biblioteca_cliente&nombre_archivo="+archivo+"&IntId="+IntId
	$('#verifica_archivo').load('FuncionesAjax/verifica_archivo_ajax.asp', criterios, function () {

	    var archivo_validado = $('#archivo_validado').val()

	    if (archivo_validado == "no_existe") {
	        frmSend.action = "man_ClienteForm.asp?archivo=1&COD_CLIENTE=" + IntId; 
	        frmSend.submit();

	    } else {

	        alert("El archivo que intenta subir al sistema ya existe. Si desea subirlo igualmente, elimine el archivo anterior o cambie el nombre de éste.")
	        return
	    }

	})

}


</SCRIPT>



<%CerrarSCG()%>

</BODY>
</HTML>




