<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001" LCID = 1034%>
<!--#include file="../arch_utils.asp"-->
<!--#include file="../../lib/comunes/rutinas/funciones.inc" -->
<!--#include file="../../lib/comunes/rutinas/TraeCampo.inc" -->
<!--#include file="../../lib/asp/comunes/general/rutinasBooleano.inc"-->
<!--#include file="../../lib/lib.asp"-->
<link href="../css/style_generales_sistema.css" rel="stylesheet">

<!--ingresa los datos de la campana-->
<%
Response.CodePage = 65001
Response.charset="utf-8"

accion_ajax 		=request("accion_ajax")

abrirscg()



if trim(accion_ajax)="Ingresa_Campana" then

ID_CAMPANA 		=request("ID_CAMPANA")
COD_CLIENTE 	=request("COD_CLIENTE")
NOMBRE			=request("Nom_Campana")
DESCRIPCION		=request("DESCRIPCION")
fecha_inicio	=request("fecha_inicio")	
fecha_termino	=request("fecha_termino")
id_usuario 		=request("id_usuario")
Observacaion 	=request("Observacaion")
Ruts 			=request("Ruts")


''' crea la campana
strSql =ID_CAMPANA & "," &  COD_CLIENTE & ",'" & NOMBRE  & "','" & DESCRIPCION  & "','" &  fecha_inicio   & "','" & fecha_termino & "','"& Observacaion & "',"  & id_usuario
 strSql="EXEC  proc_ins_campana " &  strSql
			
	    set rsCampana=Conn.execute(strSql)
	    If not rsCampana.eof Then
		    Estado = UCase(rsCampana("estado"))
		    mensaje =UCase(rsCampana("mensaje"))
	   End if


    if Estado <> "OK" then
		response.write("'',"&	Estado	&","& mensaje)
		response.end
	else 
		ValorNuevo_rut 		=split(Ruts,CHR(10))
		total_rut 			=ubound(ValorNuevo_rut)
		
		' strSql2 = total_rut
		' rut = ""
		 TotalExistoso =  0 
		 TotalFallidos = 0
		 Total = 0
		
		if(total_rut > 3000) 	then	
			response.write(",ERROR," & "Cantidad de Rut Ingresados Supera el Maximo de Datos 3000 Registros")
			response.end
		end if 
		
		ID_CAMPANAS = mensaje 
		 For indice = 0 to total_rut 

		 if TRIM(ValorNuevo_rut(indice))<>CHR(10) and TRIM(ValorNuevo_rut(indice))<>CHR(13) and TRIM(ValorNuevo_rut(indice))<>"" then
		 
			 Total 		=cint(Total) + 1
		 
			 strSql =	ID_CAMPANAS  & ",'" &  COD_CLIENTE & "','" & ValorNuevo_rut(indice)  & "'," & Total 
			 ''' marca los deudores con la campana creada 
			 strSql="EXEC  proc_ins_campana_deudor " &  strSql
	
			 set rsCampana2=Conn.execute(strSql)
			 
			 If not rsCampana2.eof Then
				Estado = UCase(rsCampana2("estado"))
				mensaje = UCase(rsCampana2("mensaje"))
			 End if
						if Estado <> "OK" then
							 TotalFallidos 		=cint(TotalFallidos) + 1
						 else 
							 TotalExistoso 		=cint(TotalExistoso) + 1
						 end if 
		end if
		next
		
		response.write(",OK," & "Total de Datos " & Total & "<br>Total Existosos:" &  TotalExistoso & "<br>Total Fallidos:" & TotalFallidos)
		response.end
	

end if 
end if	


%>

<!--Carga los datos de la campana -->
<%
if (trim(accion_ajax)="carga_Campana") then


 ID_CAMPANA 		=request("ID_CAMPANA")
 COD_CLIENTE 	=request("COD_CLIENTE")
 Mostrar 	=request("Mostrar")

 strSql =ID_CAMPANA & "," &  COD_CLIENTE 

	strSql="EXEC  Proc_Get_Campanas " &  strSql

	set rsCampana2=Conn.execute(strSql)
	
	'' si es 0 cargo los combos 
	if Mostrar =0 then
				DO WHILE NOT  rsCampana2.eof
			  
				valor = valor + "<option value='"  &  rsCampana2("id_Campana")  &  "'>" &  rsCampana2("nombre") &  "</option>" 
				
				rsCampana2.movenext
				loop
				Objeto  =valor 
				response.write (Objeto)
				response.end
	end if 
	'' carga los datos de la campana en las cajas de texto
	if Mostrar =1 then
		response.write (","& rsCampana2("id_Campana") & "," & rsCampana2("nombre")  & ","& rsCampana2("descripcion") & "," & rsCampana2("fecha_creacion") & "," & rsCampana2("fecha_inicio") & "," &  rsCampana2("fecha_termino") & "," &  rsCampana2("fecha_modificacion") & ","&  rsCampana2("observacion") & "," & rsCampana2("usuario") & "," & rsCampana2("usuario_Modifica")& "," & rsCampana2("rut_Campana")  )
			response.end
	end if 

end if
%>

<!-- marca a los deudores con la campana  en null para la nueva asignacion (MODIFICAR)--> 
<%
if (trim(accion_ajax)="refresca_Campana_deudores") then

 ID_CAMPANA 	=request("ID_CAMPANA")
 COD_CLIENTE 	=request("COD_CLIENTE")
 Mostrar 		=request("Mostrar")
 
	strSql =ID_CAMPANA & "," &  COD_CLIENTE 
 	strSql="EXEC  proc_upd_campana_Deudor " &  strSql
	
    set rsCampana2=Conn.execute(strSql)
    response.write(",ok,datos Actualizados") 
	
end if
%>

<!--Elimina los deudores/gestiones de esta campaña para luego eliminarla (ELIMINAR)-->
<%
if (trim(accion_ajax)="Elimina_Campana") then


	 ID_CAMPANA 	=request("ID_CAMPANA")
	 COD_CLIENTE 	=request("COD_CLIENTE")
	 Mostrar 	    =request("Mostrar")

	strSql =ID_CAMPANA & "," &  COD_CLIENTE 
 	
 	strSql="EXEC  proc_del_campana_deudor " &  strSql
    set rsCampana2=Conn.execute(strSql)
    response.write(",""ok,datos Eliminados") 
end if
%>

<!-- llenar el reporte " Resumen General" (SELECCIONAR CAMPANA)-->
<%
if (trim(accion_ajax)="refresca_Reporte_Asignacion") then
Response.charset="utf-8"
 ID_CAMPANA 	=request("ID_CAMPANA")
 COD_CLIENTE 	=request("COD_CLIENTE")
 
 if (ID_CAMPANA=0)then
 response.write("")
 response.end
 end if 
 
 
	strSql =ID_CAMPANA & "," &  COD_CLIENTE 
 	strSql="EXEC  Proc_get_Campana_Detalle  " &  strSql
	
	Cabecera = "<table width='80%' border='2' align='center' cellpadding='5' bordercolor='#E3EDF8' class='tabla_general'>" & _
				"<tr class='tabla_titulos' border='2' bordercolor='#E3EDF8'>" &	_
						"<td width='15%'>Total Casos</td>" & _
						"<td width='15%'>Total  Monto Casos</td>" & _
						"<td width='15%'>Promedio Monto Por Caso</td>" & _
						"<td width='15%'>M&iacutenimo Monto Por Caso</td>" & _
						"<td width='15%'>M&aacuteximo Monto Por Caso</td>" & _
						"<td width='15%'>Monto No Asignado</td>" & _
				"</tr>"	& _
				"<tr class='tabla_info' border='2' bordercolor='#E3EDF8'>"
	
    set rsCampana2=Conn.execute(strSql)
	'if  NOT  rsCampana2.eof
	    valor = valor + "<td>"   &  rsCampana2("total_casos")  &  "</td>"
		valor = valor + "<td>$"  &  formatNumber(rsCampana2("Total_monto"),0)   &  "</td>"
		valor = valor + "<td>$"  &  formatNumber(rsCampana2("PROMEDIO_SALDO_CASO"),0) &  "</td>"
		valor = valor + "<td>$"  &  formatNumber(rsCampana2("MIN_CASO"),0)  &  "</td>"
		valor = valor + "<td>$"  &  formatNumber(rsCampana2("MAX_CASO") ,0) &  "</td>"
		valor = valor + "<td>$"  &  formatNumber(rsCampana2("TOTAL_MONTO_NO_ASIG"),0) &  "</td>  		</tr>"	
		
		'rsCampana2("TOTAL_DOCUMENTOS") 
		
		 valor = valor + "<tr class='tabla_titulos' border='2' bordercolor='#E3EDF8'>" & _
						 "<td>Casos No Asignados</td>" &	_
						 "<td>Total Documentos</td>" &	_
						 "<td>Promedio Documentos</td>" &	_
						 "<td>M&iacutenimo Documentos Por Caso</td>" &	_
						 "<td>M&aacuteximo Documentos Por Caso</td>" &	_
						 "<td>Documentos No Asignados</td></tr>" 	
		
		valor = valor + "<tr class='tabla_info' border='2' bordercolor='#E3EDF8'><td>"  &  formatNumber(rsCampana2("TOTAL_CASOS_NO_ASIG"),0)    &  "</td>"
		valor = valor + "<td>"  &  formatNumber(rsCampana2("TOTAL_DOCUMENTOS"),0)  &  "</td>"
		valor = valor + "<td>"  &  formatNumber(rsCampana2("PROMEDIO_DOC_CASO"),0)  &  "</td>"
		valor = valor + "<td>" &  formatNumber(rsCampana2("MIN_DOCUMENTOS"),0) &  "</td>"
		valor = valor + "<td>" &  formatNumber(rsCampana2("MAX_DOCUMENTOS"),0)  &  "</td>"
		valor = valor + "<td>"  &  formatNumber(rsCampana2("TOTAL_DOC_NO_ASIG"),0)   &  "</td></tr></table>"	
	
	objeto = Cabecera  + valor
		
	
    response.write(objeto) 
	
end if
%>
<!-- llenar el reporte "Ejecutivos Asociados a Campaña" (SELECCIONAR CAMPANA)-->
<%
if (trim(accion_ajax)="refresca_Reporte_Ejecutivos") then
Response.charset="utf-8"
 ID_CAMPANA 	=request("ID_CAMPANA")
 COD_CLIENTE 	=request("COD_CLIENTE")
 
	if (ID_CAMPANA=0)then
		 response.write("")
		 response.end
	 end if 
 
	strSql =ID_CAMPANA & "," &  COD_CLIENTE 
 	strSql="EXEC  proc_get_Campana_Asignacion  " &  strSql
	
	
	Cabecera = "<table ID='table_tablesorter' border='0' ALIGN='CENTER' class='intercalado' style='width:100%;' cellpadding='6'>" & _
				  "<thead class='tabla_titulos' border='2' bordercolor='#E3EDF8'>" & _
						  "<tr class='tabla_titulos' border='2' bordercolor='#E3EDF8'>" &	_
						  "<td width='20%'>Codigo Ejecutivo</td>" & _
						  "<td width='20%'>Nombre Ejecutivo</td>" & _
						  "<td width='20%'>Rut Asignados</td>" & _
						 "<td width='20%'>Documentos Asignados</td>" & _
						  "<td width='20%'>Monto Asignado</td>" & _
				  "</tr></thead>"	
	 valor =  "<tbody>"

	set rsCampana2=Conn.execute(strSql)
	
	Cont_Ejecutivos = 0 
	Rut_Asignados = 0
	
	 DO WHILE NOT  rsCampana2.eof
	
		if (rsCampana2("codigo_usuario") <>0) then 
				Cont_Ejecutivos   =  Cont_Ejecutivos + 1
		end if 
		Rut_Asignados = Rut_Asignados + rsCampana2("Rut_Asignados") 
		Documentos_asignados = Documentos_asignados + rsCampana2("Documentos_asignados") 
		Monto_Asignado = Monto_Asignado + rsCampana2("Monto_Asignado") 
		
		
		 valor = valor + "<tr bordercolor='#999999'><td width='20%' align='right'>"  &  rsCampana2("codigo_usuario")  &  "</td>"
		 valor = valor + "<td width='20%' align='right'>"  &  rsCampana2("Nombre_usuario")  &  "</td>"
		 valor = valor + "<td  width='20%' align='right'>"  &  formatNumber(rsCampana2("Rut_Asignados"), 0)  &  "</td>"
		 valor = valor + "<td  width='20%' align='right'>"  &  formatNumber(rsCampana2("Documentos_asignados"), 0)  &  "</td>"
		 valor = valor + "<td  width='20%' align='right'> $"  & formatNumber(rsCampana2("Monto_Asignado"), 0)  &  "</td>  </tr>"	

		 rsCampana2.movenext
	 loop
	 
	 valor = valor + "</tbody><thead <tr>" & _
			  "	<td width='20%' align='left'><div ALIGN='LEFT'><h5>TOTALES</h5></div></td>" & _
 	          " <td width='20%' align='left'><div ALIGN='RIGHT'>" &  formatNumber(Cont_Ejecutivos, 0) &"</div></td>" & _
 	          " <td width='20%' align='right'>" &  formatNumber(Rut_Asignados, 0) &"</td>" & _
 	          " <td width='20%' align='right'><div ALIGN='RIGHT'>" & formatNumber(Documentos_asignados, 0) &"</div></td>" & _
 	          " <td width='20%' align='right'><div ALIGN='RIGHT'>$" & formatNumber( Monto_Asignado, 0) &"</div></td>" & _
 	          "</tr><thead>"
	
	objeto =  Cabecera + valor + "</tbody></table>"
		
	
    response.write(objeto) 
	
end if
%>  

	
			
			
<%
cerrarscg()
%>

