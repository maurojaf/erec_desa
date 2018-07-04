<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html>
<html lang="es">
<HEAD>
    <meta http-equiv="X-UA-Compatible" content="IE=edge,chrome=1">
	<META HTTP-EQUIV="Cache-Control" CONTENT ="no-cache">
    <meta charset="utf-8">
	
   	<!--#include file="arch_utils.asp"-->
	<!--#include file="sesion_inicio.asp"-->
	<!--#include file="../lib/comunes/rutinas/TraeCampo.inc" -->
	<!--#include file="../lib/asp/comunes/general/RutinasVarias.inc" -->
	
  
<%
	    Response.CodePage=65001
	    Response.charset ="utf-8"
	    abrirscg()

	    intCOD_CLIENTE 		=request("COD_CLIENTE") 
		strEstadoProceso    =request("strEstadoProceso")
		termino             =request("termino")
		strhorario          =request("strhorario")	
		strTipoGestion      =request("strTipoGestion")
		Busqueda            =request("Busqueda")

%>

<%
        if Busqueda = 1 then '' documentos
	        nombre = "Documentos_" & replace(Time(),":","")
        else
	        nombre = "Direcciones_" & replace(Time(),":","")
        end if 	
	        fileName ="Informe_Ruta_" & nombre  & ".xls"
            Response.AddHeader "content-disposition", "attachment; filename=" & fileName
            Response.ContentType = "application/octet-stream"
            Response.Flush()
%>
    <style type="text/css">
        .style1
        {
            height: 50px;
            width: 44px;
        }
        .style9
        {
            height: 41px;
            width: 46px;
        }
        .style27
        {
            height: 50px;
            width: 32px;
        }
        .style29
        {
            height: 67px;
            width: 44px;
        }
        .style41
        {
            height: 67px;
            width: 32px;
        }
        .style42
        {
            height: 50px;
            width: 51px;
        }
        .style43
        {
            height: 67px;
            width: 51px;
        }
        .style44
        {
            height: 41px;
            width: 54px;
        }
        .style45
        {
            height: 67px;
            width: 54px;
        }
        .style46
        {
            height: 50px;
            width: 41px;
        }
        .style47
        {
            height: 67px;
            width: 41px;
        }
        .style48
        {
            height: 41px;
            width: 52px;
        }
        .style49
        {
            height: 67px;
            width: 52px;
        }
        .style53
        {
            height: 67px;
            width: 71px;
        }
        .style54
        {
            height: 50px;
            width: 56px;
        }
        .style55
        {
            height: 67px;
            width: 56px;
        }
        .style56
        {
            height: 41px;
            width: 48px;
        }
        .style57
        {
            height: 67px;
            width: 48px;
        }
        .style58
        {
            height: 50px;
            width: 58px;
        }
        .style59
        {
            height: 67px;
            width: 58px;
        }
        #Table1
        {
            width: 78%;
        }
        .style60
        {
            height: 54px;
        }
        .style63
        {
            height: 54px;
            width: 54px;
        }
        .style68
        {
            height: 41px;
            width: 24px;
        }
        .style69
        {
            height: 54px;
            width: 24px;
        }
        .style77
        {
            height: 54px;
            width: 80px;
        }
        .style78
        {
            height: 41px;
            width: 80px;
        }
        .style79
        {
            height: 41px;
            width: 89px;
        }
        .style80
        {
            height: 54px;
            width: 89px;
        }
        .style81
        {
            height: 41px;
            width: 93px;
        }
        .style82
        {
            height: 54px;
            width: 93px;
        }
        .style84
        {
            height: 54px;
            width: 118px;
        }
        .style85
        {
            height: 41px;
            width: 118px;
        }
        .style90
        {
            height: 54px;
            width: 48px;
        }
        .style91
        {
            height: 41px;
            width: 65px;
        }
        .style92
        {
            height: 54px;
            width: 65px;
        }
        .style94
        {
            height: 54px;
            width: 68px;
        }
        .style95
        {
            height: 41px;
            width: 68px;
        }
        .style99
        {
            height: 50px;
            width: 71px;
        }
        .style101
        {
            height: 67px;
            width: 80px;
        }
        .style102
        {
            height: 50px;
            width: 72px;
        }
        .style103
        {
            height: 67px;
            width: 72px;
        }
        .style104
        {
            height: 67px;
            width: 70px;
        }
        .style105
        {
            height: 50px;
            width: 70px;
        }
        .style106
        {
            height: 41px;
            width: 105px;
        }
        .style107
        {
            height: 54px;
            width: 105px;
        }
        .style108
        {
            height: 50px;
            width: 50px;
        }
        .style109
        {
            height: 67px;
            width: 50px;
        }
        .style110
        {
            height: 50px;
            width: 59px;
        }
        .style111
        {
            height: 67px;
            width: 59px;
        }
        .style112
        {
            height: 50px;
            width: 54px;
        }
        .style113
        {
            height: 50px;
            width: 52px;
        }
        .style114
        {
            height: 50px;
            width: 48px;
        }
        .style115
        {
            height: 50px;
            width: 80px;
        }
        .style117
        {
            height: 28px;
            width: 72px;
        }
        .style119
        {
            height: 28px;
            width: 71px;
        }
        .style120
        {
            height: 28px;
            width: 56px;
        }
        .style121
        {
            height: 28px;
            width: 80px;
        }
        .style123
        {
            height: 26px;
            width: 70px;
        }
        .style124
        {
            height: 26px;
            width: 50px;
        }
        .style126
        {
            height: 26px;
            width: 41px;
        }
        .style127
        {
            height: 26px;
            width: 54px;
        }
        .style130
        {
            height: 28px;
            width: 54px;
        }
    </style>
</head>
<body>
<%if Busqueda =  1 then 
SQlFrom  = strSql & " FROM DEUDOR D	  INNER JOIN CUOTA C ON D.RUT_DEUDOR = C.RUT_DEUDOR AND D.COD_CLIENTE in (" & intCOD_CLIENTE & ")"
			SQlFrom = SQlFrom & " 				  INNER JOIN GESTIONES G ON G.RUT_DEUDOR = D.RUT_DEUDOR AND G.FECHA_COMPROMISO IS NOT NULL"
			SQlFrom = SQlFrom & " 				  INNER JOIN GESTIONES_CUOTA GC ON G.ID_GESTION = GC.ID_GESTION AND C.ID_CUOTA = GC.ID_CUOTA" 
			SQlFrom = SQlFrom & " 				  INNER JOIN GESTIONES_TIPO_GESTION GTG ON G.COD_CATEGORIA = GTG.COD_CATEGORIA "
			SQlFrom = SQlFrom & " 							 AND G.COD_SUB_CATEGORIA = GTG.COD_SUB_CATEGORIA "
			SQlFrom = SQlFrom & " 							 AND G.COD_GESTION = GTG.COD_GESTION"
			SQlFrom = SQlFrom & " 							 AND GTG.COD_CLIENTE = D.COD_CLIENTE"
			SQlFrom = SQlFrom & " 				  LEFT JOIN DEUDOR_DIRECCION DD ON G.ID_DIRECCION_COBRO_DEUDOR = DD.ID_DIRECCION AND C.RUT_DEUDOR = DD.RUT_DEUDOR "
			SQlFrom = SQlFrom & " 				  LEFT JOIN FORMA_RECAUDACION FR ON G.ID_FORMA_RECAUDACION = FR.ID_FORMA_RECAUDACION "
			SQlFrom = SQlFrom & " 				  LEFT JOIN EMPRESAS_RECAUDADORAS ER ON ISNULL(RTRIM(DD.COMUNA),FR.NOMBRE+' '+FR.UBICACION) = ER.NOMBRE_COMUNA AND ER.COD_CLIENTE = D.COD_CLIENTE"
			SQlFrom = SQlFrom & " 				  LEFT JOIN CAJA_FORMA_PAGO ON G.FORMA_PAGO = CAJA_FORMA_PAGO.ID_FORMA_PAGO "
			SQlFrom = SQlFrom & " 				  LEFT JOIN DEUDOR_TELEFONO DT ON DT.ID_TELEFONO = G.ID_MEDIO_GESTION "
			SQlFrom = SQlFrom & " 				  LEFT JOIN TELEFONO_CONTACTO      TC  ON G.ID_CONTACTO_FONO_COBRO = TC.ID_CONTACTO"
			SQlFrom = SQlFrom & " 				  LEFT JOIN EMAIL_CONTACTO         EC   ON G.ID_MEDIO_GESTION = EC.ID_CONTACTO"
			SQlFrom = SQlFrom & "                   LEFT JOIN USUARIO u ON              gc.USUARIO_ESTADO_RUTA = u.ID_USUARIO"
			SQlFrom = SQlFrom & " WHERE ((ISNULL(GTG.CONFIRMA_CP,0) = 1 AND ISNULL(GC.CONFIRMACION_CP,'N')='S' )"
			SQlFrom = SQlFrom & " OR   (ISNULL(GTG.CONFIRMA_CP,0) = 0))"
			SQlFrom = SQlFrom & " AND (G.ID_DIRECCION_COBRO_DEUDOR IS NOT NULL"
			SQlFrom = SQlFrom & " OR ISNULL(FR.TIPO,'') = 'RETIRO' )"
					
			If strEstadoProceso = "0" then
			    SQlFrom = SQlFrom & " AND DATEDIFF(D,G.FECHA_COMPROMISO,'"& termino & "') >= 0 "
			    SQlFrom = SQlFrom & " AND DATEDIFF(D,G.FECHA_COMPROMISO,'"& termino & "') <= 30 "
			Else
                SQlFrom = SQlFrom & " AND DATEDIFF(D,G.FECHA_COMPROMISO,'"& termino & "') = 0 "
			End If

            if Trim(strEstadoProceso) <> "3" Then
				SQlFrom = SQlFrom & " 	AND ISNULL(GC.ESTADO_RUTA,0) = '" & strEstadoProceso & "'"
			End If
					
		    if Trim(strTipoGestion) = "1" Then 'Recaudacion
												
			    SQlFrom = SQlFrom & " AND (GTG.GESTION_MODULOS IN (1,11))"
												
		    ElseIf Trim(strTipoGestion) = "2" Then 'Notificacion
			    SQlFrom = SQlFrom & " AND (GTG.GESTION_MODULOS = 13)"
		    Else 
			    SQlFrom = SQlFrom & " AND (GTG.GESTION_MODULOS IN (1,11,13))"
		    End If

            SQlFrom = SQlFrom & " AND  G.ID_GESTION = C.ID_ULT_GEST_GENERAL"
            
            if Trim(strHorario) = "1" Then
			    SQlFrom = SQlFrom & " 	AND (isnull(SUBSTRING(G.HORA_HASTA,1,2),9) <= 14)"
		    End If

			if Trim(strHorario) = "2" Then
				SQlFrom = SQlFrom & " 	AND (isnull(SUBSTRING(G.HORA_HASTA,1,2),9) > 14)"
			End If

 				SQlFrom = SQlFrom & " AND C.ESTADO_DEUDA IN (1,2,7,8) "
 				SQlFrom = SQlFrom & " AND C.COD_CLIENTE in (" & intCOD_CLIENTE & ")"
                dim Tam 
                Tam = 2

%>

			<table id='tbl_Procesa'  border='1'  align='center'>
			<thead>
	        <tr >
            <% if  intCOD_CLIENTE <> "1100"  then 
                Tam =  3    
            %>
                <td  style="with:5px;font-size:10px;background:#989898;" class="style112">ID</td>
             <%end if %>
				<td style="with:5px;font-size:10px;background:#989898;" class="style1">EMPRESA</td>
				<td style="with:5px;font-size:10px;background:#989898;" class="style105">FECHA RUTA</td>
				<td style="with:5px;font-size:10px;background:#989898;" class="style108">RUT CLIENTE</td>
				<td style="with:5px;font-size:10px;background:#989898;" class="style112">CLIENTE</td>
				<td style="with:5px;font-size:10px;background:#989898;" class="style46">EJECUTIVO</td>
				<td style="with:5px;font-size:10px;background:#989898;" class="style110">RUT DEUDOR</td>
				<td style="with:5px;font-size:10px;background:#989898;" class="style113">DEUDOR</td>
				<td style="with:5px;font-size:10px;background:#989898;" class="style42">FACTURA</td>
				<td style="with:5px;font-size:10px;background:#989898;" class="style114">MONTO</td>
				<td style="with:5px;font-size:10px;background:#989898;" class="style58">COMUNA</td>
				<td style="with:5px;font-size:10px;background:#989898;" class="style102">DIRECCION</td>
				<td style="with:5px;font-size:10px;background:#989898;" class="style115">DOCUMENTOS NECESARIOS</td>
				<td style="with:5px;font-size:10px;background:#989898;" class="style99">OBSERVACION</td>
				<td style="with:5px;font-size:10px;background:#989898;" class="style54">HORARIO ENTREGA PAGOS</td>
				<td style="with:5px;font-size:10px;background:#989898;" class="style27">H</td>
	        </tr>
            </thead>	
	    	<tbody>
	     <%

			strSql ="SELECT CONVERT(VARCHAR(10),g.FECHA_COMPROMISO,103)AS FECHA_PAGOS ,c.RUT_SUBCLIENTE		RUTCLIENTE " 
            strSql = strSql  & ",c.NOMBRE_SUBCLIENTE	NOMBRECLIENTE,d.RUT_DEUDOR ,d.NOMBRE_DEUDOR DEUDOR ,c.NRO_DOC FACTURA  " 
            strSql = strSql  & ",CAST(CAST(c.SALDO AS BIGINT) AS VARCHAR(10)) AS MONTO ,c.Id_Cuota,gc.Id_Gestion "
            strSql = strSql  & ",g.HORA_DESDE + ' - ' + g.HORA_HASTA AS HORARIO_ENTREGA_PAGOS " 
            strSql = strSql  & ",CASE WHEN isnull(substring(g.HORA_DESDE,1,2),9) <= 13 and isnull(substring(g.HORA_HASTA,1,2),9) <= 14	THEN 'AM'	 ELSE 'PM'	END AS AM_PM" 
            strSql = strSql  & ",Isnull(UPPER(ISNULL(UPPER(FR.NOMBRE+' '+FR.UBICACION), upper(DD.CALLE+' '+DD.NUMERO+' '+DD.RESTO+' '+DD.comuna))),'') AS DIRECCION" 
            strSql = strSql  & ",g.DOC_GESTION AS DOCUMENTOS_NECESARIOS" 
			strSql = strSql  & ",DT.TELEFONO"
			strSql = strSql  & ",ISNULL(TC.CONTACTO,EC.CONTACTO) AS CONTACTO "
			strSql = strSql & " ,REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(G.observaciones_CAMPO ,'Ñ','N'),'Á','A'),'É','E'),'Í','I'),'Ó','O'),'Ú','U'),CHAR(10),' '),CHAR(13),' ') AS 'OBSERVACION'"
			strSql = strSql & "	,(ISNULL(DD.COMUNA,'NO DEFINIDA')) AS COMUNA"
			strSql = strSql & "	,CALLE AS CALLE_GEO"
			strSql = strSql & "	,NUMERO AS NUMERO_GEO"
			strSql = strSql & "	,COMUNA AS COMUNA_GEO"
			strSql = strSql & " 	    	,CASE WHEN GC.NUEVA_EMPRESA_REC IS NOT NULL"
			strSql = strSql & " 	    	 THEN GC.NUEVA_EMPRESA_REC"
			strSql = strSql & " 	    	 WHEN ER.EMPRESA_REC IS NOT NULL"
			strSql = strSql & "	  			 THEN ER.EMPRESA_REC"
			strSql = strSql & " 	 		 WHEN DD.COMUNA IS NOT NULL"
			strSql = strSql & " 	  		 THEN UPPER(DD.COMUNA)"
			strSql = strSql & " 	 		 ELSE 'NO DEFINIDA'"
			strSql = strSql & " 		END AS EMPRESA_RECAUDADORA,"
			strSql = strSql & " 	   'FECHA INGRESO: '+ (CONVERT(VARCHAR(10),g.FECHA_INGRESO,103) + ' ' + CONVERT(VARCHAR(5),convert(datetime, g.HORA_INGRESO), 108)  +' '+"
			strSql = strSql & " 	   'USUARIO ESTADO RUTA: ' + ISNULL(u.LOGIN,'NO PROCE.') +"
			strSql = strSql & " 	   ' FECHA ESTADO: ' + ISNULL(CONVERT(VARCHAR(10),GC.FECHA_ESTADO_RUTA,103)+' '+CONVERT(VARCHAR(10),GC.FECHA_ESTADO_RUTA,108),'NO PROC.')) AS OBSERVACION_GENERAL_RUTA"
			
            strSql = strSql & SQlFrom

			               ' response.write(strSql)
							'response.end
							set rsDet=Conn.execute(strSql)
									
                                   dim i 
                                   i = 0
                                   dim rut_Evaluar 
                                   rut_Evaluar = "0"
                                   dim Total_Doc
                                   Total_Doc = 0 

							if not rsDet.eof then
							do until rsDet.eof
							
                   			If rsDet("TELEFONO") <> "" then
				                strFonoCobro =  rsDet("TELEFONO")
			                Else
				                strFonoCobro = ""
			                End If	

                       		If rsDet("CONTACTO") <> "" then
						                strContactoCobro = " " + rsDet("CONTACTO")
			                Else
						                strContactoCobro = ""
			                End If


                            if rut_Evaluar  <> rsDet("RUT_DEUDOR") then 
                               i = i + 1
                               rut_Evaluar  =rsDet("RUT_DEUDOR") 
                            end if 

							If Trim(rsDet("DIRECCION")) <> "" Then
									srtAnexoMsg1 = rsDet("DIRECCION")
							Else
									srtAnexoMsg1 = "SIN INFORMACIÓN"
							End If

						   If rsDet("DOCUMENTOS_NECESARIOS") <> "" Then
								srtAnexoMsg2 = rsDet("DOCUMENTOS_NECESARIOS")
							Else
								srtAnexoMsg2 = "SIN INFORMACIÓN"
							End If	

                    		If rsDet("OBSERVACION") <> "" then
					                strObservacionRuta = " OBS: " + rsDet("OBSERVACION")
			                Else
					                strObservacionRuta = " OBS: SIN INFORMACIÓN" 
			                End If

							
							srtAnexoMsg3 = strFonoCobro + strContactoCobro + strObservacionRuta							

                            ''' suma el total de documentos $$$$
                            Total_Doc =  Total_Doc  +  rsDet("MONTO")

                           LargoRut = len(rsDet("RUTCLIENTE"))
                           Rut_cliente = left(rsDet("RUTCLIENTE"),LargoRut-2)
                           Rut_cliente = formatnumber(Rut_cliente,0)
                           Rut_cliente = Rut_cliente & Right(rsDet("RUTCLIENTE"),2)


                           LargoRutD = len(rsDet("RUT_DEUDOR"))
                           Rut_Deudor = left(rsDet("RUT_DEUDOR"),LargoRutD-2)
                           Rut_Deudor= formatnumber(Rut_Deudor,0)
                           Rut_Deudor= Rut_Deudor & Right(rsDet("RUT_DEUDOR"),2)
                           

%>

				<tr>
                    <% if  intCOD_CLIENTE <> "1100"  then %>
                    <td style="with:5px;font-size:10px;" class="style45"><%=i%></td>
                    <%end if %>
					<td style="with:5px;font-size:10px;" class="style29"><%=rsDet("EMPRESA_RECAUDADORA") %></td>
					<td style="with:5px;font-size:10px;" class="style104"><%=rsDet("FECHA_PAGOS")%></td>
					<td style="with:5px;font-size:10px;" class="style109"><%=Rut_cliente%></td>
					<td style="with:5px;font-size:10px;" class="style45"><%=rsDet("NOMBRECLIENTE")%></td>
					<td style="with:5px;font-size:10px;" class="style47">LLACRUZ</td>
					<td style="with:5px;font-size:10px;" class="style111"><%=Rut_Deudor%></td>
					<td style="with:10px;font-size:10px;" class="style49"><%=rsDet("DEUDOR")%></td>
					<td style="with:10px;font-size:10px;" class="style43"><%=rsDet("FACTURA")%></td>
					<td style="with:10px;font-size:10px;" class="style57"><%=formatnumber(rsDet("MONTO"),0)%></td>
					<td style="with:10px;font-size:10px;" class="style59"><%=rsDet("COMUNA")%></td>
					<td style="with:10px;font-size:10px;" class="style103"><%=srtAnexoMsg1%></td>
					<td style="with:10px;font-size:10px;" class="style101"><%=srtAnexoMsg2%></td>
					<td style="with:10px;font-size:10px;" class="style53"><%=srtAnexoMsg3%></td>
					<td style="with:10px;font-size:10px;" class="style55"><%=rsDet("HORARIO_ENTREGA_PAGOS")%></td>
					<td style="with:10px;font-size:10px;" class="style41"><%=rsDet("AM_PM")%></td>
				</tr>
	      <%
	      		Response.Flush
           	    rsDet.movenext
		        loop
				end if 
	      %>
                <tr></tr>
                <tr></tr>
                <tr>
                <td colspan="<%=Tam%>" style="with:5px;font-size:10px;background:#989898;">EMPRESA</td>
                <td colspan="3" style="with:5px;font-size:10px;background:#989898;" class="style127">RUT DEUDOR</td>
         		<td colspan="3"style="with:5px;font-size:10px;background:#989898;">DEUDOR</td>
				<td colspan="2"style="with:5px;font-size:10px;background:#989898;">CANTIDAD DOCUMENTOS</td>
				<td colspan="2"style="with:5px;font-size:10px;background:#989898;">DOCUMENTOS NECESARIOS</td>
				<td colspan="3"style="with:5px;font-size:10px;background:#989898;">MONTO A RECAUDAR</td>
				

	            </tr>

           <% '''''''''''''' *********************************** calcula el total de lo a recaudar

                strSql = "" 
                strSql =  "SELECT D.RUT_DEUDOR ,D.NOMBRE_DEUDOR DEUDOR ,COUNT(C.NRO_DOC) DOCUMENTOS,SUM(CAST(C.SALDO AS BIGINT)) AS MONTO, g.DOC_GESTION AS DOCUMENTOS_NECESARIOS " 
                strSql = strSql & ",CASE WHEN GC.NUEVA_EMPRESA_REC IS NOT NULL THEN GC.NUEVA_EMPRESA_REC WHEN ER.EMPRESA_REC IS NOT NULL THEN ER.EMPRESA_REC WHEN DD.COMUNA IS NOT NULL THEN UPPER(DD.COMUNA) ELSE 'NO DEFINIDA' END AS EMPRESA_RECAUDADORA"
                strSql  = strSql & SQlFrom 
                strSql  = strSql & " group by d.RUT_DEUDOR  ,d.NOMBRE_DEUDOR , g.DOC_GESTION , " 
                strSql  = strSql & " (CASE WHEN GC.NUEVA_EMPRESA_REC IS NOT NULL THEN GC.NUEVA_EMPRESA_REC WHEN ER.EMPRESA_REC IS NOT NULL THEN ER.EMPRESA_REC WHEN DD.COMUNA IS NOT NULL THEN UPPER(DD.COMUNA) ELSE 'NO DEFINIDA' END)"   

                ''response.Write("strSql " & strSql)
                
		        set rsDet2=Conn.execute(strSql)
		        if not rsDet2.eof then
		        do until rsDet2.eof

                    LargoRutD = len(rsDet2("RUT_DEUDOR"))
                    Rut_Deudor = left(rsDet2("RUT_DEUDOR"),LargoRutD-2)
                    Rut_Deudor= formatnumber(Rut_Deudor,0)
                    Rut_Deudor= Rut_Deudor & Right(rsDet2("RUT_DEUDOR"),2)
                    DEUDOR = rsDet2("DEUDOR")
                    documentos = rsDet2("documentos")
                    Empresa =rsDet2("EMPRESA_RECAUDADORA")
                    If rsDet2("DOCUMENTOS_NECESARIOS") <> "" Then
						srtAnexoMsg2 = rsDet2("DOCUMENTOS_NECESARIOS")
					Else
						srtAnexoMsg2 = "SIN INFORMACIÓN"
					End If	

                    MONTO = rsDet2("MONTO")
    %>
				<tr>
                    <td  colspan="<%=Tam%>"   style="with:10px;font-size:10px;" ><%=Empresa%></td>
                    <td  colspan="3" style="with:10px;font-size:10px;" ><%=Rut_Deudor%></td>
					<td  colspan="3" style="with:10px;font-size:10px;" ><%=DEUDOR%></td>
					<td  colspan="2" style="with:10px;font-size:10px;" align="center"  ><%=documentos%></td>
					<td  colspan="2" style="with:10px;font-size:10px;" ><%=srtAnexoMsg2%></td>
					<td  colspan="3" style="with:10px;font-size:10px;" ><%=formatnumber(MONTO,0)%></td>
				</tr>
	      <%
	      		Response.Flush
           	    rsDet2.movenext
		        loop
				end if 
	      %>


	      	</tbody>
	        </table>


            

<%else%>			

		<table id='Table1' border='1'  align='center'>
			<thead>
	        <tr>
                <% if  intCOD_CLIENTE <> "1100"  then %>
                <td style="with:5px;font-size:10px;background:#989898;" class="style68">ID</td>
                <%end if%>
				<td style="with:5px;font-size:10px;background:#989898;" class="style44">EMPRESA</td>
				<td style="with:5px;font-size:10px;background:#989898;" class="style91">FECHA RUTA</td>
				<td style="with:5px;font-size:10px;background:#989898;" class="style56">EJECUTIVO</td>
				<td style="with:5px;font-size:10px;background:#989898;" class="style95">RUT DEUDOR</td>
				<td style="with:5px;font-size:10px;background:#989898;" class="style81">DEUDOR</td>
                <td style="with:5px;font-size:10px;background:#989898;" class="style48">CANT DOC</td>
                <td style="with:5px;font-size:10px;background:#989898;" class="style48">MONTO DOC</td>
				<td style="with:5px;font-size:10px;background:#989898;" class="style9">HORARIO</td>
				<td style="with:5px;font-size:10px;background:#989898;" class="style79">COMUNA</td>
				<td style="with:5px;font-size:10px;background:#989898;" class="style78">DIRECCION</td>
				<td style="with:5px;font-size:10px;background:#989898;" class="style85">OBSERVACION</td>
                <td style="with:5px;font-size:10px;background:#989898;" class="style106">TIPOGESTION</td>
			</tr>
            </thead>	
	    	<tbody>
            <%
			strSql ="SELECT " & _
					" CONVERT(VARCHAR(10),g.FECHA_COMPROMISO,103)AS FECHA_PAGOS  " & _
					" ,d.RUT_DEUDOR  " & _
					" ,d.NOMBRE_DEUDOR DEUDOR " & _
					" ,count(gc.Id_Cuota) CantidadDocumentos  "  & _
                    " ,SUM(C.SALDO)  MONTO"  & _
					" ,CASE WHEN isnull(substring(g.HORA_DESDE,1,2),9) <= 13 and isnull(substring(g.HORA_HASTA,1,2),9) <= 14 THEN 'AM' ELSE 'PM' END AS AM_PM " & _
					" ,Isnull(UPPER(ISNULL(UPPER(FR.NOMBRE+' '+FR.UBICACION),upper(DD.CALLE+' '+DD.NUMERO+' '+DD.RESTO+' '+DD.comuna))),'') AS DIRECCION " & _
					" ,CASE WHEN GC.NUEVA_EMPRESA_REC IS NOT NULL " & _
					"	 THEN GC.NUEVA_EMPRESA_REC " & _
					"	 WHEN ER.EMPRESA_REC IS NOT NULL " & _
					"		 THEN ER.EMPRESA_REC " & _
					"	 WHEN DD.COMUNA IS NOT NULL " & _
					"	 THEN UPPER(DD.COMUNA) " & _
					"	 ELSE 'NO DEFINIDA' " & _
					"    END AS EMPRESA_RECAUDADORA " & _
					"    ,gc.Id_Gestion , 'hola' OBSERVACION_GENERAL_RUTA " 
            strSql = strSql & ",DT.TELEFONO"
			strSql = strSql & ",ISNULL(TC.CONTACTO,EC.CONTACTO) AS CONTACTO "
			strSql = strSql & " ,REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(G.observaciones_CAMPO ,'Ñ','N'),'Á','A'),'É','E'),'Í','I'),'Ó','O'),'Ú','U'),CHAR(10),' '),CHAR(13),' ') AS 'OBSERVACION'"
			strSql = strSql & "	,DD.COMUNA,CALLE AS CALLE_GEO"
			strSql = strSql & "	,NUMERO AS NUMERO_GEO"
			strSql = strSql & "	,COMUNA AS COMUNA_GEO"
			strSql = strSql & " ,'FECHA INGRESO: '+ (CONVERT(VARCHAR(10),g.FECHA_INGRESO,103) + ' ' + CONVERT(VARCHAR(5),convert(datetime, g.HORA_INGRESO), 108)  +' '+"
			strSql = strSql & "  'USUARIO ESTADO RUTA: ' + ISNULL(u.LOGIN,'NO PROCE.') +"
			strSql = strSql & "' FECHA ESTADO: ' + ISNULL(CONVERT(VARCHAR(10),GC.FECHA_ESTADO_RUTA,103)+' '+CONVERT(VARCHAR(10),GC.FECHA_ESTADO_RUTA,108),'NO PROC.')) AS OBSERVACION_GENERAL_RUTA"
            strSql = strSql & ", CASE WHEN GTG.GESTION_MODULOS IN (1,11) THEN 'RECAUDACION' ELSE 'NOTIFICACION' END TIPOGESTION"  
			strSql = strSql & " FROM DEUDOR D	  INNER JOIN CUOTA C ON D.RUT_DEUDOR = C.RUT_DEUDOR AND D.COD_CLIENTE in (" & intCOD_CLIENTE & ")"
			strSql = strSql & " 				  INNER JOIN GESTIONES G ON G.RUT_DEUDOR = D.RUT_DEUDOR AND G.FECHA_COMPROMISO IS NOT NULL"
			strSql = strSql & " 				  INNER JOIN GESTIONES_CUOTA GC ON G.ID_GESTION = GC.ID_GESTION AND C.ID_CUOTA = GC.ID_CUOTA" 
			strSql = strSql & " 				  INNER JOIN GESTIONES_TIPO_GESTION GTG ON G.COD_CATEGORIA = GTG.COD_CATEGORIA "
			strSql = strSql & " 							 AND G.COD_SUB_CATEGORIA = GTG.COD_SUB_CATEGORIA "
			strSql = strSql & " 							 AND G.COD_GESTION = GTG.COD_GESTION"
			strSql = strSql & " 							 AND GTG.COD_CLIENTE = D.COD_CLIENTE"
			strSql = strSql & " 				  LEFT JOIN DEUDOR_DIRECCION DD ON G.ID_DIRECCION_COBRO_DEUDOR = DD.ID_DIRECCION AND C.RUT_DEUDOR = DD.RUT_DEUDOR "
			strSql = strSql & " 				  LEFT JOIN FORMA_RECAUDACION FR ON G.ID_FORMA_RECAUDACION = FR.ID_FORMA_RECAUDACION "
			strSql = strSql & " 				  LEFT JOIN EMPRESAS_RECAUDADORAS ER ON ISNULL(RTRIM(DD.COMUNA),FR.NOMBRE+' '+FR.UBICACION) = ER.NOMBRE_COMUNA AND ER.COD_CLIENTE = D.COD_CLIENTE"
			strSql = strSql & " 				  LEFT JOIN CAJA_FORMA_PAGO ON G.FORMA_PAGO = CAJA_FORMA_PAGO.ID_FORMA_PAGO "
			strSql = strSql & " 				  LEFT JOIN DEUDOR_TELEFONO DT ON DT.ID_TELEFONO = G.ID_MEDIO_GESTION "
			strSql = strSql & " 				  LEFT JOIN TELEFONO_CONTACTO      TC  ON G.ID_CONTACTO_FONO_COBRO = TC.ID_CONTACTO"
			strSql = strSql & " 				  LEFT JOIN EMAIL_CONTACTO         EC   ON G.ID_MEDIO_GESTION = EC.ID_CONTACTO"
			strSql = strSql & "                   LEFT JOIN USUARIO u ON              gc.USUARIO_ESTADO_RUTA = u.ID_USUARIO"
			strSql = strSql & " WHERE ((ISNULL(GTG.CONFIRMA_CP,0) = 1 AND ISNULL(GC.CONFIRMACION_CP,'N')='S' )"
			strSql = strSql & " OR   (ISNULL(GTG.CONFIRMA_CP,0) = 0))"
			strSql = strSql & " AND (G.ID_DIRECCION_COBRO_DEUDOR IS NOT NULL"
			strSql = strSql & " OR ISNULL(FR.TIPO,'') = 'RETIRO' )"
					
			If strEstadoProceso = "0" then

			strSql = strSql & " AND DATEDIFF(D,G.FECHA_COMPROMISO,'"& termino & "') >= 0 "
			strSql = strSql & " AND DATEDIFF(D,G.FECHA_COMPROMISO,'"& termino & "') <= 30 "

			Else

			strSql = strSql & " AND DATEDIFF(D,G.FECHA_COMPROMISO,'"& termino & "') = 0 "

			End If
			if Trim(strEstadoProceso) <> "3" Then
				strSql = strSql & " 	AND ISNULL(GC.ESTADO_RUTA,0) = '" & strEstadoProceso & "'"
			End If
			if Trim(strTipoGestion) = "1" Then 'Recaudacion
				strSql = strSql & " AND (GTG.GESTION_MODULOS IN (1,11))"
			ElseIf Trim(strTipoGestion) = "2" Then 'Notificacion
				strSql = strSql & " AND (GTG.GESTION_MODULOS = 13)"
			Else 
				strSql = strSql & " AND (GTG.GESTION_MODULOS IN (1,11,13))"
			End If
				strSql = strSql & " AND  G.ID_GESTION = C.ID_ULT_GEST_GENERAL"
			if Trim(strHorario) = "1" Then
				strSql = strSql & " 	AND (isnull(SUBSTRING(G.HORA_HASTA,1,2),9) <= 14)"
			End If
			if Trim(strHorario) = "2" Then
				strSql = strSql & " 	AND (isnull(SUBSTRING(G.HORA_HASTA,1,2),9) > 14)"
			End If
			strSql = strSql & " AND C.ESTADO_DEUDA IN (1,2,7,8) "
			strSql = strSql & " AND C.COD_CLIENTE in (" & intCOD_CLIENTE & ")"
			strSql = strSql & "  group  by	 d.RUT_DEUDOR ,d.NOMBRE_DEUDOR,g.HORA_DESDE,g.HORA_HASTA ,FR.NOMBRE,FR.UBICACION " 
			strSql = strSql & " ,DD.CALLE,DD.NUMERO,DD.RESTO,DD.comuna,g.FECHA_COMPROMISO " 
			strSql = strSql & " ,GC.NUEVA_EMPRESA_REC,ER.EMPRESA_REC,gc.Id_Gestion,DT.TELEFONO,TC.CONTACTO,EC.CONTACTO,G.observaciones_CAMPO,COMUNA,NUMERO,CALLE,GC.FECHA_ESTADO_RUTA" 
			strSql = strSql & " ,g.FECHA_INGRESO,g.HORA_INGRESO,u.LOGIN,GTG.GESTION_MODULOS"
                                   
	   On error Resume Next 
		 	set rsDet=Conn.execute(strSql)

		 '
         'response.write(",Error," & strSql)
		 ' response.end

	   If Err.number<>0 then
		   response.write(",Error," & Err.Description )
		   response.end
	   end if 

                                   dim j 
                                   j = 0
                                   dim  Total_Doc_Dir 
                                   Total_Doc_Dir = 0

			if not rsDet.eof then
			do until rsDet.eof
				j = j +1


			If rsDet("TELEFONO") <> "" then
				strFonoCobro =  rsDet("TELEFONO")
			Else
				strFonoCobro = ""
			End If	

			If rsDet("CONTACTO") <> "" then
						strContactoCobro = " " + rsDet("CONTACTO")
			Else
						strContactoCobro = ""
			End If	


			If rsDet("OBSERVACION") <> "" then
					strObservacionRuta = " OBS: " + rsDet("OBSERVACION")
			Else
					strObservacionRuta = " OBS: SIN INFORMACIÓN" 
			End If


			srtAnexoMsg3 = strFonoCobro + strContactoCobro + strObservacionRuta

			strDireccion2 = rsDet("CALLE_GEO") & " " & rsDet("NUMERO_GEO") & " " & rsDet("COMUNA_GEO")
			strDireccion2 = Trim(strDireccion2)

			strDireccion_geo = strDireccion_geo
			strComuna = rsDet("COMUNA")
			

            ''' suma el total de documentos $$$$
                            Total_Doc_Dir =  Total_Doc_Dir +  rsDet("MONTO")

                              LargoRutD = len(rsDet("RUT_DEUDOR"))
                           Rut_Deudor = left(rsDet("RUT_DEUDOR"),LargoRutD-2)
                           Rut_Deudor= formatnumber(Rut_Deudor,0)
                           Rut_Deudor= Rut_Deudor & Right(rsDet("RUT_DEUDOR"),2)

			%>

				<tr>
                    <% if  intCOD_CLIENTE <> "1100"  then %>
                        <td style="with:5px;font-size:10px;"  class="style69"><%=j%></td>
                    <%end if %>
					<td style="with:5px;font-size:10px;"  class="style63" ><%=rsDet("EMPRESA_RECAUDADORA") %></td>
					<td style="with:5px;font-size:10px;" class="style92" ><%=rsDet("FECHA_PAGOS")%></td>
					<td style="with:5px;font-size:10px;" class="style90" >LLACRUZ</td>
					<td style="with:5px;font-size:10px;" class="style94" ><%=Rut_Deudor%></td>
                    <td style="with:5px;font-size:10px;" class="style82" ><%=rsDet("DEUDOR")%></td>
                    <td style="with:5px;font-size:10px;" class="style60" ><%=rsDet("CantidadDocumentos")%></td>
                    <td style="with:5px;font-size:10px;" class="style60" ><%=formatnumber(rsDet("MONTO"),0)%></td>
					<td style="with:5px;font-size:10px;" class="style60" ><%=rsDet("AM_PM")%></td>
					<td style="with:5px;font-size:10px;" class="style80" ><%=rsDet("comuna")%></td>
					<td style="with:5px;font-size:10px;" class="style77" ><%=rsDet("DIRECCION")%></td>
					<td style="with:5px;font-size:10px;" class="style84"><%=srtAnexoMsg3%></td>
                    <td style="with:5px;font-size:10px;" class="style107"><%=rsDet("TIPOGESTION")%></td>
                    
				</tr>
	      <%
	      		Response.Flush
           	    rsDet.movenext
		        loop
				end if 
	      %>
            <% if  intCOD_CLIENTE <> "1100"  then %>
            <tr >
            <td bgcolor="#6e6e6e" align="right" colspan="7"><font color="#ffffff">TOTAL</font></td>
            <td align="left" bgcolor="#6E6E6E" colspan="6"><font color="#ffffff"><%=formatnumber(Total_Doc_Dir,0)%></font></td>
            </tr>
            <%else %>
             <tr >
            <td bgcolor="#6e6e6e" align="right" colspan="6"><font color="#ffffff">TOTAL</font></td>
            <td align="left" bgcolor="#6E6E6E" colspan="6"><font color="#ffffff"><%=formatnumber(Total_Doc_Dir,0)%></font></td>
            </tr>
            <%end if %>

	      	</tbody>
	        </table>


<%end if%>			
</body>
</html>
<%CerrarSCG1()%>

