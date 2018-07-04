<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html>
<html lang="es">
<head>
    <meta http-equiv="X-UA-Compatible" content="IE=edge,chrome=1">
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8"/> 
    <link href="../css/normalize.css" rel="stylesheet">
    <link href="../css/style_generales_sistema.css" rel="stylesheet">
    <link href="../Componentes/prettyLoader/css/prettyLoader.css" rel="stylesheet">
	<!--#include file="sesion.asp"-->
	<!--#include file="arch_utils.asp"-->
	<!--#include file="../lib/conex_mysql.inc"-->
	<!--#include file="../lib/comunes/rutinas/funciones.inc" -->
	<!--#include file="../lib/comunes/rutinas/TraeCampo.inc" -->
	<!--#include file="../lib/asp/comunes/general/rutinasSCG.inc" -->
	<!--#include file="../lib/asp/comunes/general/rutinasBooleano.inc"-->
	<!--#include file="../lib/lib.asp"-->
	<!--#include file="../lib/comunes/rutinas/chkFecha.inc"-->
	<!--#include file="../lib/comunes/rutinas/sondigitos.inc"-->
	<!--#include file="../lib/comunes/rutinas/formatoFecha.inc"-->
	<!--#include file="../lib/comunes/rutinas/validarFecha.inc"-->
	<!--#include file="../lib/comunes/rutinas/diasEnMes.inc"-->
	<!--#include file="../lib/asp/comunes/general/rutinasBooleano.inc"-->
	<link href="../css/style.css" rel="stylesheet" type="text/css">
	<link href="../css/style_generales_sistema.css" rel="stylesheet">
    <link rel="stylesheet" type="text/css" href="../Componentes/jquery-1.9.2/css/start/jquery-ui-1.9.2.custom.css">
   

<%
	'cod_caja=110
	Response.CodePage=65001
	Response.charset ="utf-8"

	cod_caja=Session("intCodUsuario")
    strLetrasColorModPrev = "F3F3F3"
	AbrirSCG()

	strAnexo = request("cmb_usuario")
    strTipotelefono =   request("CB_TipoTelefono")
    strTelefono= Replace(request("TX_Telefono"),"-","")
    strfono= request("TX_Telefono")
	

    IF strTelefono <> "" THEN strCodArea = LEFT(strTelefono,2)


    pagina= request("pagina")
	PaginasTotales =  100	
    paginaDesde = pagina
    strLlamada = request("CB_LLAMADA")
    
    resp = request("resp")

    StrCodCliente= session("ses_codcli") 	
    

    If Trim(strLlamada) = "" Then strLlamada = "OUTBOUND" 'strLlamada = "TODAS"'
    if Trim(strCodArea)="" then strCodArea = ""  

	
	if sucursal="" then sucursal="0"
   
    if pagina="" OR pagina< 0  then pagina=0
    if paginaDesde ="" OR paginaDesde <= 0  then 
     paginaDesde=1 
    end if 
    

	usuario = request("cmb_usuario")

	if usuario = "" then usuario = "0"
	
    termino = request("termino")
	inicio = request("inicio")
	
	if Trim(inicio) = "" Then
		inicio = TraeFechaMesActual(Conn,-1)
		inicio = "01" & Mid(inicio,3,10)
		inicio = TraeFechaActual(Conn)
	End If
	if Trim(termino) = "" Then
		termino = TraeFechaActual(Conn)
	End If
	CLIENTE = REQUEST("CLIENTE")

   ' response.Write("inicio-->" & inicio)
   ' response.Write("<br/>termino-->" & termino)

%>
<title>CRM Cobros</title>
<style type="text/css">
<!--
.Estilo13 {color: #FFFFFF}
.Estilo27 {color: #FFFFFF}
-->
</style>

<script language="JavaScript" src="../javascripts/cal2.js"></script>
<script language="JavaScript" src="../javascripts/cal_conf2.js"></script>
<script language="JavaScript" src="../javascripts/validaciones.js"></script>
<script src="../javascripts/SelCombox.js"></script>
<script src="../javascripts/OpenWindow.js"></script>

<script src="../Componentes/jquery-1.9.2/js/jquery-1.8.3.js"></script>
<script src="../Componentes/jquery-1.9.2/js/jquery-ui-1.9.2.custom.js"></script>
<script src="../Componentes/jquery.tablesorter/jquery.tablesorter.js"></script>
<script src="../Componentes/prettyLoader/js/jquery.prettyLoader.js"></script> 

<script language="JavaScript " type="text/JavaScript">
$(document).ready(function(){

    $.prettyLoader();
   

    $('#inicio').datepicker( {changeMonth: true,changeYear: true, dateFormat: 'dd/mm/yy',
			    beforeShowDay: DisableDays })

    $('#termino').datepicker( {changeMonth: true,changeYear: true, dateFormat: 'dd/mm/yy',
			    beforeShowDay: DisableDays })


	
	$("#table_tablesorter").tablesorter({
	    headers: {
	        9: { sorter: 'decimal'} // column number, type
	    }
	})

})
function Refrescar()
{
	resp='no'
	datos.action = "listado_grabaciones.asp?resp="+ resp +"";
	datos.submit();
}

function ImpBoleta(intCompPago)
{
	window.open("imprime_boleta.asp?intNroComp=" + intCompPago,"INFORMACION","width=800, height=600, scrollbars=yes, menubar=no, location=no, resizable=yes");
}

function Limpiar() {

    $('#CB_LLAMADA').val('OUTBOUND');
    $('#cmb_usuario').val('0');
    $('#inicio').val('');
    $('#termino').val('');
    $('#CB_TipoTelefono').val('0');
    $('#TX_Telefono').val('');

        
    $.prettyLoader.show();

    resp = 'si'
    document.datos.action = "listado_grabaciones.asp?resp=" + resp + "";
    document.datos.submit();
    

}


function envia()
{
	
    Telefono = $('#TX_Telefono').val();

    //alert(Telefono.length)

    if ((Telefono.length > 1  & Telefono.length < 8))
    {
        alert("Largo Telefono no valido")
        return ;
    }

    var filter = /^[0-9--]*$/;

    if (filter.test(Telefono)) {
      
    }else{
        alert('Ingrese fortmato Telefono CodArea - Telefono');
        $('#TX_Telefono').val('');
        return;
    }

    $.prettyLoader.show();

	resp='si'
	document.datos.action = "listado_grabaciones.asp?resp="+ resp +"";
	document.datos.submit();
}

function imprimir()
{
	datos.action = "imprime_comprobantes.asp";
	datos.submit();
}

function TraerGrabacion (strIdGrab){
	URL='EscucharGrabacion.asp?strIdGrab=' + strIdGrab
	window.open(URL,"DATOS_GRABACION","width=470, height=230, scrollbars=no, menubar=no, location=no, resizable=yes")
}



function IrPagina(sintAccion) {
    resp = 'Pag'
    $.prettyLoader.show();
    if (sintAccion == 'Retroceder') {

        self.location.href = 'listado_grabaciones.asp?resp=' + resp + '&pagina=<%=pagina-PaginasTotales %>&cmb_usuario=<%=strAnexo%>&CB_TipoTelefono=<%=strTipotelefono%>&TX_CodArea=<%=strCodArea%>&TX_Telefono=<%=strTelefono%>&CB_LLAMADA=<%=strLlamada%>&inicio=<%=inicio%>&termino=<%=termino%>&ANEXO=<%=strAnexo%>'
    }
    if (sintAccion == 'Avanzar') {
        self.location.href = 'listado_grabaciones.asp?resp=' + resp + '&pagina=<%=PaginasTotales + pagina%>&cmb_usuario=<%=strAnexo%>&CB_TipoTelefono=<%=strTipotelefono%>&TX_CodArea=<%=strCodArea%>&TX_Telefono=<%=strTelefono%>&CB_LLAMADA=<%=strLlamada%>&inicio=<%=inicio%>&termino=<%=termino%>&ANEXO=<%=strAnexo%>'
    }

}





</script>

</head>

<body>

<input type="HIDDEN" name="strArrFeriados" 	id="strArrFeriados" 	value="<%=strArrFeriados%>">

<form name="datos" method="post">
<div class="titulo_informe">LISTADO DE GRABACIONES</div>	
<br>
<table width="90%" height="500" border="0" align="center">
  <tr>
    <td valign="top" colspan="3" height="70px">
	<table width="100%" border="0" bordercolor="#999999" class="estilo_columnas">
		<thead>
	      <tr   bordercolor="#999999"  bgcolor="#<%=session("COLTABBG")%>" class="Estilo13">
	      	<td>TIPO LLAMADA</td>
			<td>USUARIO</td>
			<td>DESDE</td>
			<td>HASTA</td>
			<td>TIPO TEL&EacuteFONO</td>
			<td>TEL&EacuteFONO</td>
			<td>&nbsp;</td>
	      </tr>
	     </thead>
		  <tr bordercolor="#999999" class="Estilo8">
		  	<td>
			        <SELECT NAME="CB_LLAMADA" id="CB_LLAMADA">
				        <option value="TODAS" <%if Trim(strLlamada)="TODAS" then response.Write("Selected") end if%>>TODAS</option>
                        <option value="INBOUND" <%if Trim(strLlamada)="INBOUND" then response.Write("Selected") end if%>>INBOUND</option>
				        <option value="OUTBOUND" <%if Trim(strLlamada)="OUTBOUND" then response.Write("Selected") end if%>>OUTBOUND</option>
			        </SELECT>
			</td>
			<td>
				    <SELECT NAME="cmb_usuario" id="cmb_usuario">
					    <option value="0-0">TODOS</option>
					    <%
					    stsSql="select u.ID_USUARIO,u.ANEXO,u.LOGIN from USUARIO u inner join USUARIO_CLIENTE UC  on uc.ID_USUARIO = u.ID_USUARIO and uc.COD_CLIENTE =  " & StrCodCliente & " where activo = 1 and  PuedenEscucharMisGrabaciones = 1 and anexo is not null "

                        set rsUsu=Conn.execute(stsSql)
					    if not rsUsu.eof then
						    do until rsUsu.eof
						    %>
						    <option value="<%=rsUsu("ANEXO") &"-"& rsUsu("ID_USUARIO") %>"  <%if Trim(strAnexo)=(Trim(rsUsu("ANEXO"))&"-"& Trim(rsUsu("ID_USUARIO")))   then response.Write("Selected") end if%>> <%=ucase(rsUsu("LOGIN"))%></option> 

						    <%rsUsu.movenext
						    loop
					    end if
					    rsUsu.close
					    set rsUsu=nothing
					    %>
				    </SELECT>
			</td>
			<td>        
                    <input name="inicio" readonly="true" type="text" id="inicio" value="<%=inicio%>" size="10" maxlength="10" >
			</td>
			<td>
                    <input name="termino" readonly="true" type="text" id="termino" value="<%=termino%>" size="10" maxlength="10">
			</td>
			<td>
			        <SELECT NAME="CB_TipoTelefono" id="CB_TipoTelefono">
				        <option value="TODOS" <%if Trim(strTipotelefono)="TODOS" then response.Write("Selected") end if%>>TODOS</option>
                        <option value="CELULAR" <%if Trim(strTipotelefono)="CELULAR" then response.Write("Selected") end if%>>CELULAR</option>
				        <option value="RED FIJA" <%if Trim(strTipotelefono)="RED FIJA" then response.Write("Selected") end if%>>RED FIJA</option>
			        </SELECT>
            </td>
			<td>
                    
                    <input title="(2-23464320)" name="TX_Telefono" type="text" id="TX_Telefono" value="<%=strfono%>" size="13" maxlength="12">
                    
                    </td>
			<td>
                    <input type="Button" name="Submit0" value="Limpiar" class="fondo_boton_100" 
                        onClick="Limpiar();">&nbsp;
                    <input type="Button" name="Submit" value="Ver" class="fondo_boton_100" onClick="envia();">
            </td>
	      </tr>
    </table>
    </td></tr>
    <tr>
    <td valign="top" colspan="3">
	<table  class="tablesorter"  id="table_tablesorter" style="width:100%;">
     
	<thead>
		<tr >
			<td>&nbsp;</th>
			<th>FECHA</th>
			<th>HORA</th>
			<th>EJECUTIVO</th>
			<th>TIPO LLAMADA</th>
			<th>TIPO TEL&EacuteFONO</th>
			<th>ANEXO</th>
			<th>TEL&EacuteFONO</th>
			<th>D&UacuteRACION (MIN)</th>
			<td>&nbsp;</td>
            <th>CARGADO</th>
		</tr>
	</thead>
	
    <tbody>
    
	<%


	AbrirSCG1()
		
       

    if strAnexo = "0-0" or  strAnexo = "" then 
    
        ssql="select ANEXO from USUARIO u inner join USUARIO_CLIENTE UC  on uc.ID_USUARIO = u.ID_USUARIO and uc.COD_CLIENTE =  " & StrCodCliente & " where activo = 1 and  PuedenEscucharMisGrabaciones = 1 and anexo is not null "
        set rsCLI=Conn1.execute(ssql)
		'Response.Write(ssql)
		Do While not rsCLI.eof
			'Response.Write(rsCLI("ANEXO") & "<br />***")
			strAnexos = strAnexos & Trim(rsCLI("ANEXO")) & ","
			rsCLI.movenext
		Loop
		rsCLI.close
		set rsCLI=nothing
		
		if Trim(strAnexos) <> "" then
		
			strAnexos=Mid(strAnexos,1,len(strAnexos)-1)
			
		end if
		
    else 
	
        My_Array=split(strAnexo,"-")
		
        strAnexos  =My_Array(0)

    end if 

	


	inicio = Mid(inicio,7,4) & "-" & Mid(inicio,4,2) & "-" & Mid(inicio,1,2)
	termino = Mid(termino,7,4) & "-" & Mid(termino,4,2) & "-" & Mid(termino,1,2)

        if strCodArea <> "09" AND strCodArea  <> "02"  then 
            strTelefono  ="0" & strTelefono 
        end if 

       strSql = "select uniqueid,replace(replace(replace(replace(clid,'"""
       strSql = strSql & "',''),'<',''),'>',''),src,'') clid,  ROUND((billsec/60),2) duration, src, dst,DATE_FORMAT(calldate, '%H:%i:%s') HoraLLamada,calldate,DATE_FORMAT(calldate, '%d/%m/%Y') FechaLLamada, lastapp " 
       strSql = strSql & "from asteriskcdrdb.cdr where billsec >= 10  and (calldate BETWEEN CAST('" & inicio & " 00:00:00.000' AS DATETIME) AND '" & termino & " 23:59:59.000')"

       

	If Trim(strLlamada)="OUTBOUND" Then ''' salientes

		if Trim(strAnexos) <> "" then
			strSql = strSql & " and src in (" & strAnexos & ")"
		end if

        if Trim(strTelefono) <> ""  and  Trim(strCodArea) <> ""  Then
			strSql = strSql & " and dst ='" & strTelefono & "'"
		End If

          ''' celular o fijo
        if Trim(strTipotelefono) = "CELULAR" Then
			strSql = strSql & " and   left(dst,2)  =09"
        else if Trim(strTipotelefono) = "RED FIJA" then
            strSql = strSql & " and   left(dst,2)  <> 09"
		End If
        End If

		strSql = strSql & " and disposition = 'ANSWERED' AND  dst REGEXP ('^[0-9]+$')  order by calldate desc LIMIT " & pagina & " ," & PaginasTotales 

  Else If Trim(strLlamada)="INBOUND" Then ''' entrantes buscara por el anexo ya uqe hay queda registrado el telefono
	
         strSql = strSql & " and length(src) > 3"
		
         if strAnexo <> "0-0" or  strAnexo <> "" then 
			if Trim(strAnexos) <> "" then
                strSql = strSql & " and dst in (" &  strAnexos & ") "
			end if
         end if 

		if Trim(strTelefono) <> ""  and  Trim(strCodArea) <> ""  Then
			strSql = strSql & " and src ='" & Mid(strTelefono, 2,  len(strTelefono)) & "'"
		End If

          ''' celular o fijo
        if Trim(strTipotelefono) = "CELULAR" Then
			strSql = strSql & " and   left(src,1)  =9"
        else if Trim(strTipotelefono) = "RED FIJA" then
            strSql = strSql & " and   left(src,1)  <> 9"
		End If
        End If

		strSql = strSql & " and disposition = 'ANSWERED'  AND  dst REGEXP ('^[0-9]+$')  order by calldate desc LIMIT " & pagina & " ," & PaginasTotales 
   
   else 

         
        if strAnexo = "0-0" or  strAnexo = "" then 
            if Trim(strAnexos) <> "" then
				strSql = strSql & "and (length(src) > 3 or src in (" &  strAnexos & ")) "
			else
				strSql = strSql & " and length(src) > 3 "
			end if
		else
            strSql = strSql & "and src in ((" &  Mid(strTelefono, 2,  len(strTelefono)) & ")) "
        end if   
		
        
        if Trim(strTelefono) <> ""  and  Trim(strCodArea) <> ""  Then
			strSql = strSql & " and ( dst ='" &   strTelefono & "'"
            strSql = strSql & " or  src ='" & Mid(strTelefono, 2,  len(strTelefono)) & "')"
		End If

        ''' celular o fijo
        if Trim(strTipotelefono) = "CELULAR" Then
			strSql = strSql & " and  ((left(dst,2)  =09 and length(src) <= 3 ) OR (left(src,1) = 9 and length(src) > 3 ))"
        else if Trim(strTipotelefono) = "RED FIJA" then
            strSql = strSql & " and   ((left(dst,2)  <> 09 and length(src) <= 3 )OR (left(src,1) <> 9 and length(src) > 3))"
		End If
        End If

		strSql = strSql & " and disposition = 'ANSWERED' AND  dst REGEXP ('^[0-9]+$')   order by calldate desc LIMIT " & pagina & " ," & PaginasTotales 

    End If
	End If

	
     'Response.write strSql
	 'Response.End
    Marca  = 0
	if strSql <> "" AND pagina >= 0 then
        
	    set rsGrab= ConnMySql.execute(strSql)
        

        if not rsGrab.eof then
            Marca = Marca + 1
            intReg = pagina
			intReg=intReg + 1
			'response.Write("aaaa" & Marca)
            if Marca = 1 then 
                paginaDesde = intReg
            end if 


            do while not rsGrab.eof

				strIdGrab = Trim(rsGrab("uniqueid"))
				strAnexo= Trim(rsGrab("src"))
				intDuracion = Trim(rsGrab("duration"))
				strTelefono = Trim(rsGrab("dst"))
				strTipo = Trim(rsGrab("lastapp"))
				'strFechaHora = Trim(rsGrab("calldate"))
				strFechaGrab = rsGrab("FechaLLamada") 'Mid(strFechaHora,1,10)
				strHoraGrab = Mid(rsGrab("HoraLLamada"),1,5) ' Mid(strFechaHora,12,5)
                strUsuario = UCase(Trim(rsGrab("clid")))
                strAnexoG = Trim(rsGrab("src"))
                strOrigen = strAnexo

                'strTipoLlamada =  Left(strTelefono,2)  
              

                if lEN(strOrigen) > 3 then
                    strOrigen = "INBOUND"
                    strTelefonoform = "0" & strAnexoG
                ELSE 
                    strOrigen = "OUTBOUND"
                    strTelefonoform = strTelefono
                end if 
                
                strTipoLlamada =  Left(strTelefonoform,2)  

                
                        if strTipoLlamada = "09" then
                            strTelefonoform = Left(strTelefonoform,2)  & "-"  & right(strTelefonoform,len(strTelefonoform)-2)
                            strTipoLlamada = "CELULAR"
                          
                        ELSE 
                        
                         if strTipoLlamada = "02" then 
                                strTelefonoform = Left(strTelefonoform,2) & "-" & right(strTelefonoform,len(strTelefonoform)-2)
                         else
                               strTelefonoform  = Left(strTelefonoform,3) & "-"  & right(strTelefonoform,len(strTelefonoform)-3)
                         end if 
                            strTipoLlamada = "RED FIJA"
                        end if 
             

            
		
        dim strExisteFono

        ssql="select top 1 ID_TELEFONO from DEUDOR_TELEFONO where TELEFONO_DAL = '" &  Replace(strTelefonoform,"-","") & "'"
        

		set rsTE=Conn1.execute(ssql)
        dim ExisteFono 
        ExisteFono = false
        if rsTE.eof then 
            ExisteFono = false
            strExisteFono  = "<img src='../imagenes/no.ico' border='0'>"
        else 
            strExisteFono  = "<img src='../imagenes/ok.ico' border='0'>" 
            ExisteFono = true
        end if 
	
    	set rsCLI=nothing
     


         If Trim(strOrigen)<>"INBOUND"   then 
             strAnexoG      =  strAnexoG 
           else 
             strAnexoG      = strTelefono
          END IF 

         strTelefono = strTelefonoform
     
         if ExisteFono  = "Verdadero"  then 
            strTelefono    = "<a href='Busqueda.asp?TX_TELASOCIADO=" & Replace(strTelefonoform,"-","") & "'>" & strTelefonoform & "</a>" 
        end if 

			
            %>
            
			<tr>
				<td><%=intReg%></td>
				<td><%=strFechaGrab%></td>
				<td><%=strHoraGrab%></td>
				<td><%=strUsuario%></td>
                <td><%=strOrigen%></td>
				<td><%=strTipoLlamada%></td>
                <td><%=strAnexoG%></td><!--strAnexoG-->
                <td><%=strTelefono %></td>
                
				<td><%=Replace(intDuracion,",",".")%></td>
				<td align="center">
					<A HREF="#" onClick="TraerGrabacion('<%=strIdGrab%>');return false;";>
						<img src="../imagenes/sound.png" border="0">
					</A> 
				</td>

                <td align="center"><%=strExisteFono%></td>
                
			</tr>
			<%
				intReg=intReg+1
				rsGrab.movenext
			loop %>
        <%else %>
        <tr><td colspan="11">BUSQUEDA SIN RESULTADOS</td> </tr>
        <%end if%>
    

	<%end if%>
	</tbody>
    
	</table>
   </td>
   </tr>
   <tr>
       <%if intReg > 1 then %>
        <TD ALIGN=center colspan="3" height="15px">
            <A HREF="#" onClick="IrPagina( 'Retroceder')";><img src="../imagenes/previous.gif" border="0"></A>
            &nbsp;
            <FONT FACE="verdana, Sans-Serif" Size=1 ><b>Resultado Desde <%= paginaDesde %> Hasta <%= intReg-1 %> </b></FONT>
            &nbsp;
            <A HREF="#" onClick="IrPagina( 'Avanzar')";><img src="../imagenes/next.gif" border="0"></A>
		</TD>
        <%end if %>
        </tr>  
  </table>
  <br>
  <br>
</form>
<script>
    $("#ventana_procesa").css('visibility', 'hidden');
</script>

</body>
</html>
<%
    
    strSql = "[dbo].[sp_genera_dias_inahabiles] "&trim(year(date()))&", " & trim(1)
    set rsFeriados=Conn1.execute(strSql)

	strArrFeriados=""
	Do While not rsFeriados.eof
		strArrFeriados = strArrFeriados & "'" & rsFeriados("FECHA") & "',"
		rsFeriados.movenext
	Loop
	strArrFeriados = Mid(strArrFeriados,1,len(strArrFeriados)-1)

	'CerrarScg1()
    %>




<% CerrarSCG1() %>


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



			


    </SCRIPT>