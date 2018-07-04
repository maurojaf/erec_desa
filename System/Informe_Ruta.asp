<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"  LCID = 1034%>
<!DOCTYPE html>
<html lang="es">
<HEAD>
   
    <meta http-equiv="X-UA-Compatible" content="IE=edge,chrome=1">
    <meta charset="utf-8">
   
	<!--#include file="sesion.asp"-->
	<!--#include file="arch_utils.asp"-->
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


    <%
		Id_Usuario = session("session_idusuario")
        Perfil_Usuario = session("perfil_emp")
       ' response.Write("---->" & Perfil_Usuario )
	%>



<%
	Response.CodePage=65001
	Response.charset ="utf-8"
	AbrirSCG()
	strEstadoProceso 	= request("cmb_estadoProceso")
	str_empresa_rec = request("cmb_empresa_rec")
	inicio 			= request("inicio")
	termino 		= request("termino")
	resp 			= request("resp")
	strhorario 		= REQUEST("cmb_horario")
	strTipoGestion	= request("cmb_tipogestion")
    strTipoBusqueda	= request("Busqueda")
    strRecarga	    = request("strRecarga")
    intCod_Cliente =  request("CB_CLIENTE")
    
    ''' valor utilzado para ver actualziar las empresas
    if strRecarga = "NO" or  str_empresa_rec = ""  then str_empresa_rec = "0"
    if intCod_Cliente ="" then  intCod_Cliente = session("ses_codcli")       

	if strEstadoProceso = "" then strEstadoProceso = "0"
	if strhorario = "" then strhorario = "1"
	if strTipoGestion = "" then strTipoGestion = "0" end if 
    if strTipoBusqueda = "" then strTipoBusqueda = "1"

	if Trim(inicio) = "" Then
		strMesActual = Month(TraeFechaActual(Conn))
		strAnoActual = Cdbl(Year(TraeFechaActual(Conn)))
		''strDiaActual = Cdbl(day(TraeFechaActual(Conn)))
		If strMesActual = 1 Then strAnoActual = strAnoActual - 1
		If strMesActual = 1 Then strMesActual = 12
		strMesActual = strMesActual - 1
		if Len(strMesActual) = 1 Then strMesActual = "0" & strMesActual
		If Trim(inicio) = "" Then inicio = "01/" & strMesActual & "/" & strAnoActual
	End If

	
	'termino = session("termino")

	if Trim(termino) = "" Then
		termino = TraeFechaActual(Conn)
	End If

	Hoy = TraeFechaActual(Conn)


	'response.write "<br>intCod_Cliente=" & intCod_Cliente
%>


    <link href="../css/style_multi_select.css" rel="stylesheet"> 
    <link href="../css/normalize.css" rel="stylesheet">	
	<link rel="stylesheet" type="text/css" href="../Componentes/jquery-1.9.2/css/start/jquery-ui-1.9.2.custom.css">
	<link href="../css/style_generales_sistema.css" rel="stylesheet"> 
	<link href="../Componentes/prettyLoader/css/prettyLoader.css" rel="stylesheet">
    <link href="../css/jquery.alerts.css" rel="stylesheet"> 
	<script src="../Componentes/jquery-1.9.2/js/jquery-1.8.3.js"></script>
	<script src="../Componentes/jquery-1.9.2/js/jquery-ui-1.9.2.custom.js"></script>
	<script src="../Componentes/prettyLoader/js/jquery.prettyLoader.js"></script> 
	<script src="../Componentes/jquery.alerts.mod.js"></script>
    
    <script src="../Componentes/jquery.multiselect.js"></script>
    

	<script language="JavaScript " type="text/JavaScript">

	    $(document).ready(function () {

	        $(document).tooltip();
	        $.prettyLoader();
	        $('#termino').datepicker({ changeMonth: true, changeYear: true, dateFormat: 'dd/mm/yy' });
	        $("#CB_CLIENTE").multiselect({ minWidth: 175 });


	        var cambiaestado = $('#TXT_CAMBIA').val()
	        if (cambiaestado == '') {
	            envia();
	        } else {
	            consulta_resumen2('','','');
	        }

	 

	    });

	    $(document).delegate('#CB_CLIENTE', 'change', function () {
	        var str = $('#CB_CLIENTE option:selected').val(); // Get Value of Option

	        if (str == 'close') {
	            /*$('#selectmenu').selectmenu('close');*/
	            alert('asdasd')
	        }
	    });



	    function MarcaChexbok() {
	        var Indice = $('#Desde').val()
	        
           var chkHeader = document.getElementById("chckHead");
           for (var i = 1; i < document.getElementById('tbl_Procesa').rows.length; i++)
            {
                chk = document.getElementById("ChckRow_" + Indice)
                chk.checked = chkHeader.checked
                Indice = parseInt(Indice)+ 1
                
           }
        }


        function MostrarVentana(valor) {
          
            var ancho = 550;
            var alto = 270;
            var titulo="Procesar Rutas"
            var cont = 0;
            var Indice = $('#Desde').val()

            for (var i = 1; i < document.getElementById('tbl_Procesa').rows.length; i++) {
                chk = document.getElementById("ChckRow_" + Indice)
                if (chk.checked == true) { 
                    cont=cont+1
                }
                Indice = parseInt(Indice) + 1;
            }

            if (cont == 0) {
                var Busqueda = $("#Busqueda").val();
                var msj = ""
                 if (Busqueda == 1) { msj = "Documento" } else { msj = "Direccion" }

                jAlert("Indique " + msj + " a Modificar", "Advertencia!")
                return;
                }
            //  alert(valor)
            //  procesar ruta
                
                // NOTIFICACIONES MOSTRAMOS OTROS CAMPOS FECHA Y HORA DE RETIRO 
                var tipo_Gestion = $("#cmb_tipogestion").val();
                if (tipo_Gestion == 2) 
                {
                    $("#cmb_horario0").css('visibility', 'visible');
                    $("#TerminoNotificacion").css('visibility', 'visible');
                    $("#Label3").css('visibility', 'visible');
                    $("#Label2").css('visibility', 'visible');
          
                } else {
               
                    $("#cmb_horario0").css('visibility', 'hidden');
                    $("#TerminoNotificacion").css('visibility', 'hidden');
                    $("#Label3").css('visibility', 'hidden');
                    $("#Label2").css('visibility', 'hidden');
                   
                }


            if (valor == 1) {
                    $("#span_Empresa").css('display', 'Block');
                    $("#CB_EMPRESA_REC").css('display', 'Block');
                    $("#Label1").css('display', 'Block');
                                     
                    $("#Estado").css('display', 'None');
                    $("#LblEstado").css('display', 'None');
                    $("#span_Estado").css('display', 'None');
                    $("#LblObservacion").css('display', 'None');
                    $("#Txt_Observacion").css('display', 'None');

                    $("#cmb_horario0").css('visibility', 'hidden');
                    $("#TerminoNotificacion").css('visibility', 'hidden');
                    $("#Label3").css('visibility', 'hidden');
                    $("#Label2").css('visibility', 'hidden');

                    alto = 250;
                    ancho = 580;
                    titulo = "Modificar Empresa Recaudadora"
                } // cambio empresa
                else 
                {

                    $("#span_Empresa").css('display', 'None');
                    $("#CB_EMPRESA_REC").css('display', 'None');
                    $("#Label1").css('display', 'none');
            
                    $("#Estado").css('display', 'Block');
                    $("#LblEstado").css('display', 'Block');
                    $("#span_Estado").css('display', 'Block');
                    $("#Txt_Observacion").css('display', 'Block');
                    $("#LblObservacion").css('display', 'Block');
                }

            
            
            $("#Txt_Observacion").val("");
            $("#CB_EMPRESA_REC").val(0);
            $("#Estado").val(1);
          
	        $('#ventana_procesa').dialog({
	            show: "blind",
	            hide: "explode",
	            width: 550,
	            height: alto,
	            modal: true,
               title: titulo,
	            buttons: {
	                Si: function () {
	                    Procesa(valor);
	                   
	                },
	                Cerrar: function () {
	                    $('#ventana_procesa').dialog("close");
	                }
	            }
	        });

	    }
	  

	function consulta_resumen2(Desde,Hasta,Estado) {
	    
	    
	    var Cod_Cliente = $('#CB_CLIENTE').val()
	    strEstadoProceso = $("#cmb_estadoProceso").val();  
	    termino = $("#termino").val(); 
	    strhorario = $("#cmb_horario").val();
	    strTipoGestion = $("#cmb_tipogestion").val();
	    Busqueda = $("#Busqueda").val();
	    str_empresa_rec = $("#cmb_empresa_rec").val();
        var Paginado  = 50

        /*avanza en el paginado*/
         if (Estado == 1) {
            
                if (Desde == ""){ Desde = 1 }
                else{ Desde = Hasta + 1}

                if (Hasta == "") { Hasta = Paginado }
                else { Hasta = Hasta + Paginado }
            }

            
        /*Retroceder en el paginado*/
            if (Estado == 0) {

                if (Desde == "") {
                Desde = 1 
            }
            else {
                Desde = Desde - Paginado 
            }

            if (Hasta == "") { Hasta = Paginado }
            else { Hasta = Hasta - Paginado }
        }



        if (Desde < 0) {
            Desde = 1
            Hasta = Paginado 
        }


        $("#Desde").val(Desde)



	    if (Cod_Cliente == null) {
	        return;
	    }
	    $('#refresca_resumen').text("")

	    $.ajax({ url: "FuncionesAjax/proceso_rutas_ajax.asp?accion_ajax=refresca_resumen&Cod_Cliente=" + Cod_Cliente , 
			type: "POST",
			data: { termino: termino, strEstadoProceso: strEstadoProceso, strhorario: strhorario, strTipoGestion: strTipoGestion, Busqueda: Busqueda, str_empresa_rec: str_empresa_rec, Desde: Desde, Hasta: Hasta, Estado: Estado },
			success: function(msg) {
				$('#refresca_resumen').append(msg)

			}
            ,error: function(XMLHttpRequest, textStatus, errorThrown){
								jAlert("Error al procesar los datos  ('"+errorThrown+"')")
							}
           });
	}
		

	function envia() {
	    $.prettyLoader.show();
	    /*datos.Submit.disabled = true;*/
	    document.datos.action = 'Informe_ruta.asp?resp=si';
	    document.datos.submit();
	  
	}



	function Procesa(valor) {

	    var Observacion = $("#Txt_Observacion").val();
	    var Empresa = $("#CB_EMPRESA_REC").val();
	    var Estado = $("#Estado").val();
	    var Busqueda = $("#Busqueda").val();
	    var strTipoGestion = $("#cmb_tipogestion").val();
	    var Cod_Cliente = $('#CB_CLIENTE').val();
	    var Perfil_Usuario = $('#Perfil_Usuario').val();
	    var horarioNotificaion = $("#cmb_horario0").val();
	    var FECHA_COMPROMISO = $('#TerminoNotificacion').val();
	    var HORA_HASTA
	    var HORA_DESDE


	    //alert(horarioNotificaion);
	    //return;
        
	    if (horarioNotificaion == "1") {
	        HORA_DESDE = "09:00"
	        HORA_HASTA = "14:00"

	    } else {
	        HORA_DESDE = "14:00"
	        HORA_HASTA = "19:00"
	    }
	   // alert(HORA_DESDE)
	   // alert(HORA_HASTA)

	    
	    if (valor == 1) {
	        if (Empresa == 0) {
	                    $('#span_Empresa').css('border-color', '#FE2E2E')
	                    $('#span_Empresa').text("*")
	                    jAlert("Seleccione Empresa Recaudadora", "Advertencia!")
	                    return;
	        } else {
	            $('#span_Empresa').text("")
	            $(this).css('border-color', '')
	        }

	    
	    } else {
                        // sin seleccionar
	                    if (Estado == 3) {
	                        $('#span_Estado').css('border-color', '#FE2E2E')
	                        $('#span_Estado').text("*")
	                        jAlert("Seleccione Estado Ruta", "Advertencia!")
	                        return;
	                    } else {
	                        $('#span_Estado').text("")
	                        $(this).css('border-color', '')
	                    }



	                    if (Estado == 2) {
	                        if (Observacion.length == 0) {
	                            jAlert("Debe Ingrese Observación", "Advertencia!")
	                            return;
	                        }
	                        else {
	                            $('#span_Observacion').text("")
	                            $(this).css('border-color', '')
	                        }
	                    }
	    }
	                /*alert(Perfil_Usuario)*/


	    jConfirm("¿Esta seguro De Modificar La Ruta?", "Advertencia!", function (r) {
	        if (r) {

	            var Indice = $('#Desde').val()

	            for (var i = 1; i < document.getElementById('tbl_Procesa').rows.length; i++) {

	                chk = document.getElementById("ChckRow_" + Indice).checked
	                var id = document.getElementById('tbl_Procesa').rows[i].cells[0].innerHTML;
	                var EmpresaR = document.getElementById('tbl_Procesa').rows[i].cells[2].innerHTML;
	                var id_gestion = 0
	                var id_cuota = 0




	                // documentos
	                if (Busqueda == 1) {

	                    if ((Perfil_Usuario.toUpperCase() != "TRUE")) {
	                        /*alert("acac")*/
	                        id_gestion = document.getElementById('tbl_Procesa').rows[i].cells[18].innerHTML;
	                        id_cuota = document.getElementById('tbl_Procesa').rows[i].cells[19].innerHTML;
	                    } else {
	                        id_gestion = document.getElementById('tbl_Procesa').rows[i].cells[17].innerHTML;
	                        id_cuota = document.getElementById('tbl_Procesa').rows[i].cells[18].innerHTML;
	                    }
	                }
	                else // direcciones
	                {
	                    if ((Perfil_Usuario.toUpperCase() != "TRUE")) {
	                        id_gestion = document.getElementById('tbl_Procesa').rows[i].cells[14].innerHTML;
	                    } else {

	                        id_gestion = document.getElementById('tbl_Procesa').rows[i].cells[13].innerHTML;
	                    }
	                }

	                /*alert(id_gestion);*/


	                if (chk == true) 
                    {

	                    $.ajax({ url: "FuncionesAjax/proceso_rutas_ajax.asp?accion_ajax=Procesa_Rutas",
	                        type: "POST",
	                        data: { Observacion: Observacion, Empresa: Empresa, Estado: Estado, id_gestion: id_gestion, id_cuota: id_cuota, EmpresaR: EmpresaR, id: id, valor: valor, strTipoGestion: strTipoGestion, HORA_DESDE: HORA_DESDE, HORA_HASTA: HORA_HASTA, FECHA_COMPROMISO: FECHA_COMPROMISO },
	                        success: function (msg) {
	                            var Mensaje = msg.split(',');
	                            var estado = Mensaje[1];
	                            var Msj = Mensaje[2];
	                            /*$('#refresca_resumen2').append(msg)*/
	                            if (estado.toUpperCase() != "OK") {
	                                jAlert(Msj, "Procesar Rutas");
	                                return;
	                            } else {
	                                /*$('#refresca_resumen2').append(Msj)  */
	                            }
	                        }, error: function (XMLHttpRequest, textStatus, errorThrown) {
	                            jAlert("Error al procesar los datos  ('" + errorThrown + "')")
	                        }
                          
	                    });


	                }
	                Indice = parseInt(Indice) + 1;

	            }

	            // window.location.href("Informe_Ruta.asp?resp=si");

	            document.datos.action = 'Informe_ruta.asp?strRecarga=si';
	            document.datos.submit();

	        } else {
	            jAlert("Proceso Cancelado", "Advertencia!");
	        }
	        
	    }); 
		
	}

	function exportar() {

	    var Cod_Cliente = $('#CB_CLIENTE').val();
	    var strEstadoProceso = $("#cmb_estadoProceso").val();
	    var termino = $("#termino").val();
	    var strhorario = $("#cmb_horario").val();
	    var strTipoGestion = $("#cmb_tipogestion").val();
	    var Busqueda = $("#Busqueda").val();


	    var pagina = 'exp_Informe_rutas.asp?Cod_Cliente=' + Cod_Cliente + '&strEstadoProceso=' + strEstadoProceso + '&termino=' + termino + '&strhorario=' + strhorario + '&strTipoGestion=' + strTipoGestion + '&Busqueda=' + Busqueda
  	    window.open(pagina, 'window', 'params');
	}

	
	function bt_geolocalizacion(direccion) {

	    direccion = direccion.replace("%", " ")
	    direccion = direccion.replace("%", " ")
	    direccion = direccion.replace("%", " ")
	    direccion = direccion.replace("%", " ")
	    direccion = direccion.replace("%", " ")
	    direccion = direccion.replace("%", " ")
	    direccion = direccion.replace("%", " ")
	    direccion = direccion.replace("%", " ")


	    direccion = direccion + ', CHILE'
        
        window.open('geolocalizacion.asp?direccion='+encodeURIComponent(direccion),"DATOS1","width=610, height=610, scrollbars=no, menubar=no, location=no, resizable=yes")

	}
	
    </script>

	<style type="text/css">
        .hiddencol
        {
            display:none;
        }
        .span_aviso_rojo{
		color:#FE2E2E;
		font-size:12px;
		width:1px;
		margin-right:2cm;
	}
	.abrir_cerrar {
	cursor: pointer;
	display: inline-block;
	float: right;
	position: relative;
	margin-right:2cm;
	</style>
</head>

<body leftmargin="0" rightmargin="0" marginwidth="0" topmargin="0" marginheight="0">
<form name="datos" method="post">
<div class="titulo_informe">Gestión de Rutas</div>
<br>
&nbsp;<table width="90%" class="estilo_columnas" align="center">
		<thead>
	      <tr height="20" >
			<td>CLIENTE</td>
			<td>ESTADO PROCESO RUTA</td>
			<td class="style1">FECHA RUTA</td>
			<td>HORARIO RUTA</td>
			<td>EMPRESA RECAUDADORA</td>
			<td>TIPO GESTIÓN</td>
			<td>TIPO INFORMACION</td>
			<td >&nbsp;</td>
		  </tr>
		</thead>
			
		  <tr >
         
        
			<td>
		
                <select name="CB_CLIENTE" id="CB_CLIENTE"  multiple >
				<%
				ssql="SELECT  COD_CLIENTE,RAZON_SOCIAL FROM CLIENTE WHERE ACTIVO = 1 AND COD_CLIENTE IN (SELECT COD_CLIENTE FROM USUARIO_CLIENTE WHERE ID_USUARIO = " & session("session_idusuario") & ") ORDER BY COD_CLIENTE,RAZON_SOCIAL ASC"
				set rsTemp= Conn.execute(ssql)
				if not rsTemp.eof then
			    
                do until rsTemp.eof
                ''' recorremos los valores para ver q clientes estan seleccioandos
                   cadena = intCOD_CLIENTE
                   y =  rsTemp("COD_CLIENTE") 
                    strChecked = ""
                    a=Split(cadena,",")
                    for each x in a
                        IF  trim(x) = trim(y)  THEN
                            strChecked = " selected "
                        END IF 
                    next
                    
                    %>
					<option  <%=strChecked%> value="<%=rsTemp("COD_CLIENTE")%>" ><%=rsTemp("RAZON_SOCIAL")%></option>
					<%

                    
					rsTemp.movenext
					loop
				end if
				rsTemp.close
				set rsTemp=nothing
				%>
				</select>
                
                
                </td>


			<td>
				<SELECT NAME="cmb_estadoProceso" id="cmb_estadoProceso" onChange="envia();">
					
					<option value="0" <%If Trim(strEstadoProceso)="0" Then Response.write "SELECTED"%>>NO PROCESADO</option>
					<option value="1" <%If Trim(strEstadoProceso)="1" Then Response.write "SELECTED"%>>PROCESADO</option>
					<option value="2" <%If Trim(strEstadoProceso)="2" Then Response.write "SELECTED"%>>RECHAZADO</option>
				</SELECT>
			</td>

			<td class="style1"><input name="termino" type="text" id="termino" value="<%=termino%>" readonly="true" size="10" maxlength="10" onChange="envia();">
			</td>

			<td>
				<SELECT NAME="cmb_horario" id="cmb_horario" onChange="envia();">
					<option value="0" <%If Trim(strhorario)="0" Then Response.write "SELECTED"%>>TODOS</option>
					<option value="1" <%If Trim(strhorario)="1" Then Response.write "SELECTED"%>>AM</option>
					<option value="2" <%If Trim(strhorario)="2" Then Response.write "SELECTED"%>>PM</option>
				</SELECT>
			</td>
<%

					strSql = "	SELECT "
					strSql = strSql & " (CASE WHEN GC.NUEVA_EMPRESA_REC IS NOT NULL "
					strSql = strSql & " 	  THEN GC.NUEVA_EMPRESA_REC "
					strSql = strSql & " 	  WHEN ER.EMPRESA_REC IS NOT NULL "
					strSql = strSql & " 	  THEN ER.EMPRESA_REC "
					strSql = strSql & " 	  WHEN DD.COMUNA IS NOT NULL "
					strSql = strSql & " 	  THEN UPPER(DD.COMUNA) "
					strSql = strSql & "  ELSE 'NO DEFINIDA' END) AS EMPRESA_RECAUDADORA"

					
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
					
 					strSql = strSql & " WHERE ((ISNULL(GTG.CONFIRMA_CP,0) = 1 AND ISNULL(GC.CONFIRMACION_CP,'N')='S' )"
 					strSql = strSql & " OR   (ISNULL(GTG.CONFIRMA_CP,0) = 0))"
					
 					strSql = strSql & " AND (G.ID_DIRECCION_COBRO_DEUDOR IS NOT NULL"
 					strSql = strSql & " OR ISNULL(FR.TIPO,'') = 'RETIRO' )"
					
					If strEstadoProceso = "0" then 'No procesados ''  todos
					        strSql = strSql & " AND DATEDIFF(D,G.FECHA_COMPROMISO,'"& termino & "') >= 0 "
					        strSql = strSql & " AND DATEDIFF(D,G.FECHA_COMPROMISO,'"& termino & "') <= 30 "
					Else
                        strSql = strSql & " AND DATEDIFF(D,G.FECHA_COMPROMISO,'"& termino & "') = 0 "
					End If

					if Trim(strEstadoProceso) <> "3" Then 'Todos
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

 					strSql = strSql & " GROUP BY (CASE WHEN GC.NUEVA_EMPRESA_REC IS NOT NULL THEN GC.NUEVA_EMPRESA_REC WHEN ER.EMPRESA_REC IS NOT NULL THEN ER.EMPRESA_REC WHEN DD.COMUNA IS NOT NULL THEN UPPER(DD.COMUNA) ELSE 'NO DEFINIDA' END) "

 					strSql = strSql & " ORDER BY EMPRESA_RECAUDADORA ASC "
					'response.write "<br>strSql=" & strSql
					'response.end()

%>
			<td>
				<SELECT NAME="cmb_empresa_rec" id="cmb_empresa_rec" onChange="envia();">
					<option value="0">TODOS</option>
					<%
						set rsDetEmpresa=Conn.execute(strSql)

							if not rsDetEmpresa.eof then
								do until rsDetEmpresa.eof
								%>
								<option value="<%=rsDetEmpresa("EMPRESA_RECAUDADORA")%>"
								<%if Trim(str_empresa_rec)=Trim(rsDetEmpresa("EMPRESA_RECAUDADORA")) then
									response.Write("Selected")
								end if%>
								><%=ucase(rsDetEmpresa("EMPRESA_RECAUDADORA"))%></option>

								<%rsDetEmpresa.movenext
								loop
                                'rsDetEmpresa.MoveFirst 
							end if
							
				
                           	rsDetEmpresa.close
					        set rsDetEmpresa=nothing
					%>
				</SELECT>

			</td>

			<td>
                <SELECT NAME="cmb_tipogestion" id="cmb_tipogestion" onChange="envia();">
					<option value="0" <%If Trim(strTipoGestion)="0" Then Response.write "SELECTED"%>>TODOS</option>
					<option value="1" <%If Trim(strTipoGestion)="1" Then Response.write "SELECTED"%>>RECAUDACIÓN</option>
					<option value="2" <%If Trim(strTipoGestion)="2" Then Response.write "SELECTED"%>>NOTIFICACIÓN</option>
				</SELECT></td>
			
			<td align="center">
                <!--<input type="radio"   name="Busqueda" value="1" > Documentos &nbsp; 
                <input type="radio"   name="Busqueda" value="0" > Direcciones-->
                <SELECT  id="Busqueda" name="Busqueda" onChange="envia();">
					<option value="1" <%If Trim(strTipoBusqueda)="1" Then Response.write "SELECTED"%>>DOCUMENTOS</option>
					<option value="2" <%If Trim(strTipoBusqueda)="2" Then Response.write "SELECTED"%>>DIRECCIONES</option>
				</SELECT></td>

			<td align="center">
			    
			
				 <input type="Button" class="fondo_boton_100" name="Submit" value="Ver" onclick= "envia();"></tr>
    </table>
	
	<input type="hidden" id="TXT_CAMBIA" value='<%=resp%>'/>
	<input type="hidden" id="Perfil_Usuario" value='<%=session("perfil_emp")%>'/>
    <input type="hidden" id="Desde" value='<%=session("perfil_emp")%>'/>
    
    <!--  **************************************   div procesar ************************** -->
	 <div id="ventana_procesa"  style="display:none;">
    <!-- <div id="ventana_procesa"  >		-->
    
	<table align="center" width="500" align="right" cellSpacing="0" cellPadding="0" border="0">
	<tr>		
		<td><label id="LblEstado">Estado Ruta:</label></td>		
		<td >
        <select id="Estado">
					<option value="3" >SELECIONE</option>
					<option value="1" >PROCESAR</option>
                    <option value="0" >NO PROCESAR</option>
					<option value="2" >RECHAZADO</option>
		</select>
        </td>
        <td align="left"> 
            <span id="span_Estado" class="span_aviso_rojo">*</span>
        </td>		
	</tr>	
	<tr>		
		<td><label id="LblObservacion">Observación:</label>
        </td>	
        <td>
        <textarea cols="45" rows="4"  ID="Txt_Observacion"  maxlength="80" ></textarea>
     </td>
	</tr>
	<tr>		
		<td><label id="Label3">Fecha Ruta:</label></td>	
        <td>
            <input type="text" id="TerminoNotificacion" value="<%=termino%>" readonly="true" ></td>
	</tr>
	<tr>		
		<td><label id="Label2" >Horario RuTa:</label></td>	
        <%  if strhorario ="0" then 
             strhorario2 = "1"
            end if 
        %>
        	<td>	
            <SELECT id="cmb_horario0" >
					<option value="1" <%If Trim(strhorario2)="1" Then Response.write "SELECTED"%>>AM</option>
					<option value="2" <%If Trim(strhorario2)="2" Then Response.write "SELECTED"%>>PM</option>
				</SELECT>
                
                </td>
                </tr>
    <tr>
	<td ><label id="Label1">Emp. Recaudadora:</label></td>		
		<td>
			                    <select id="CB_EMPRESA_REC" onchange="this.style.width=130">
                                <option value="0" >SELECCIONAR</option>
									<%strSql = " SELECT DISTINCT EMPRESA_REC FROM EMPRESAS_RECAUDADORAS WHERE COD_CLIENTE in (" & intCOD_CLIENTE & ") AND TIPO_EMPRESA = 1 "
									strSql=strSql & " ORDER BY EMPRESA_REC"
                                    
									set rsTemp= Conn.execute(strSql)
									if not rsTemp.eof then
										Do until rsTemp.eof
											If Trim(strEmpresaRec) = Trim(rsTemp("EMPRESA_REC")) Then strSelEmpRec = "SELECTED" Else strSelEmpRec = "" End If
										%>
											<option value="<%=rsTemp("EMPRESA_REC")%>" <%=strSelEmpRec%>><%=UCASE(rsTemp("EMPRESA_REC"))%></option>
											<%
											rsTemp.movenext
										Loop
									end if
									rsTemp.close
									set rsTemp=nothing

									%>
								</select>
        </td>
        <td align="left">
        <span id="span_Empresa" class="span_aviso_rojo" >*</span>
		</td>			</tr>	
	</table>
    <div id="refresca_resumen2"></div>
	</div>	
<%

%>
	
		

	<%
            ''''' reporte estados
   
					strSql = "SELECT	SUM(CASE WHEN ISNULL(GC.ESTADO_RUTA,0) = 0 AND (isnull(SUBSTRING(G.HORA_HASTA,1,2),9) <= 14) THEN 1 ELSE 0 END) AS RNP_AM,"
 					strSql = strSql & " SUM(CASE WHEN ISNULL(GC.ESTADO_RUTA,0) = 1 AND (isnull(SUBSTRING(G.HORA_HASTA,1,2),9) <= 14) AND DATEDIFF(D,G.FECHA_COMPROMISO,'"& termino & "') = 0 THEN 1 ELSE 0 END) AS RP_AM,"
 					strSql = strSql & " SUM(CASE WHEN ISNULL(GC.ESTADO_RUTA,0) = 2 AND (isnull(SUBSTRING(G.HORA_HASTA,1,2),9) <= 14) AND DATEDIFF(D,G.FECHA_COMPROMISO,'"& termino & "') = 0 THEN 1 ELSE 0 END) AS RR_AM,"
 					strSql = strSql & " SUM(CASE WHEN ISNULL(GC.ESTADO_RUTA,0) = 0 AND (isnull(SUBSTRING(G.HORA_HASTA,1,2),9) > 14) THEN 1 ELSE 0 END) AS RNP_PM,"
 					strSql = strSql & " SUM(CASE WHEN ISNULL(GC.ESTADO_RUTA,0) = 1 AND (isnull(SUBSTRING(G.HORA_HASTA,1,2),9) > 14) AND DATEDIFF(D,G.FECHA_COMPROMISO,'"& termino & "') = 0 THEN 1 ELSE 0 END) AS RP_PM,"
 					strSql = strSql & " SUM(CASE WHEN ISNULL(GC.ESTADO_RUTA,0) = 2 AND (isnull(SUBSTRING(G.HORA_HASTA,1,2),9) > 14) AND DATEDIFF(D,G.FECHA_COMPROMISO,'"& termino & "') = 0 THEN 1 ELSE 0 END) AS RR_PM,"

 					strSql = strSql & " COUNT(DISTINCT (CASE WHEN ISNULL(GC.ESTADO_RUTA,0) = 0 AND (isnull(SUBSTRING(G.HORA_HASTA,1,2),9) <= 14) THEN ISNULL(G.ID_DIRECCION_COBRO_DEUDOR,G.ID_FORMA_RECAUDACION) END)) AS DNP_AM,"
 					strSql = strSql & " COUNT(DISTINCT CASE WHEN ISNULL(GC.ESTADO_RUTA,0) = 1 AND (isnull(SUBSTRING(G.HORA_HASTA,1,2),9) <= 14) AND DATEDIFF(D,G.FECHA_COMPROMISO,'"& termino & "') = 0 THEN ISNULL(G.ID_DIRECCION_COBRO_DEUDOR,G.ID_FORMA_RECAUDACION) END) AS DP_AM,"
 					strSql = strSql & " COUNT(DISTINCT CASE WHEN ISNULL(GC.ESTADO_RUTA,0) = 2 AND (isnull(SUBSTRING(G.HORA_HASTA,1,2),9) <= 14) AND DATEDIFF(D,G.FECHA_COMPROMISO,'"& termino & "') = 0 THEN ISNULL(G.ID_DIRECCION_COBRO_DEUDOR,G.ID_FORMA_RECAUDACION) END) AS DR_AM,"

 					strSql = strSql & " COUNT(DISTINCT (CASE WHEN ISNULL(GC.ESTADO_RUTA,0) = 0 AND (isnull(SUBSTRING(G.HORA_HASTA,1,2),9) > 14) THEN G.ID_GESTION END)) AS DNP_PM,"
 					strSql = strSql & " COUNT(DISTINCT CASE WHEN ISNULL(GC.ESTADO_RUTA,0) = 1 AND (isnull(SUBSTRING(G.HORA_HASTA,1,2),9) > 14) AND DATEDIFF(D,G.FECHA_COMPROMISO,'"& termino & "') = 0 THEN ISNULL(G.ID_DIRECCION_COBRO_DEUDOR,G.ID_FORMA_RECAUDACION) END) AS DP_PM," 
 					strSql = strSql & " COUNT(DISTINCT CASE WHEN ISNULL(GC.ESTADO_RUTA,0) = 2 AND (isnull(SUBSTRING(G.HORA_HASTA,1,2),9) > 14) AND DATEDIFF(D,G.FECHA_COMPROMISO,'"& termino & "') = 0 THEN ISNULL(G.ID_DIRECCION_COBRO_DEUDOR,G.ID_FORMA_RECAUDACION) END) AS DR_PM"
					
 					strSql = strSql & " FROM DEUDOR D	 INNER JOIN CUOTA C ON D.RUT_DEUDOR = C.RUT_DEUDOR AND D.COD_CLIENTE in (" & intCOD_CLIENTE & ")"
					strSql = strSql & " 				 INNER JOIN GESTIONES G ON G.RUT_DEUDOR = D.RUT_DEUDOR AND G.FECHA_COMPROMISO IS NOT NULL"
					strSql = strSql & " 				 INNER JOIN GESTIONES_CUOTA GC ON G.ID_GESTION = GC.ID_GESTION AND C.ID_CUOTA = GC.ID_CUOTA" 
					strSql = strSql & " 				 INNER JOIN GESTIONES_TIPO_GESTION GTG ON G.COD_CATEGORIA = GTG.COD_CATEGORIA "
					strSql = strSql & " 							 AND G.COD_SUB_CATEGORIA = GTG.COD_SUB_CATEGORIA "
					strSql = strSql & " 							 AND G.COD_GESTION = GTG.COD_GESTION"
					strSql = strSql & " 							 AND GTG.COD_CLIENTE = D.COD_CLIENTE"
								
					strSql = strSql & " 				 LEFT JOIN DEUDOR_DIRECCION DD ON G.ID_DIRECCION_COBRO_DEUDOR = DD.ID_DIRECCION AND C.RUT_DEUDOR = DD.RUT_DEUDOR "
					strSql = strSql & " 				 LEFT JOIN FORMA_RECAUDACION FR ON G.ID_FORMA_RECAUDACION = FR.ID_FORMA_RECAUDACION "
					strSql = strSql & " 				 LEFT JOIN EMPRESAS_RECAUDADORAS ER ON ISNULL(RTRIM(DD.COMUNA),FR.NOMBRE+' '+FR.UBICACION) = ER.NOMBRE_COMUNA AND ER.COD_CLIENTE = D.COD_CLIENTE" 
					strSql = strSql & " 				 LEFT JOIN CAJA_FORMA_PAGO ON G.FORMA_PAGO = CAJA_FORMA_PAGO.ID_FORMA_PAGO "

					strSql = strSql & " WHERE ((ISNULL(GTG.CONFIRMA_CP,0) = 1 AND ISNULL(GC.CONFIRMACION_CP,'N')='S' )"
					strSql = strSql & " OR   (ISNULL(GTG.CONFIRMA_CP,0) = 0))"
					strSql = strSql & " AND (G.ID_DIRECCION_COBRO_DEUDOR IS NOT NULL"
					strSql = strSql & " OR ISNULL(FR.TIPO,'') = 'RETIRO' )"
					
	
                    strSql = strSql & " AND ((DATEDIFF(D,G.FECHA_COMPROMISO,'"& termino & "') >= 0 "
                    strSql = strSql & " AND   DATEDIFF(D,G.FECHA_COMPROMISO,'"& termino & "') <=  30 )" 
                    strSql = strSql & " or ( DATEDIFF(D,G.FECHA_COMPROMISO,'"& termino & "') = 0))"
					
          
					if Trim(strHorario) = "1" Then
						strSql = strSql & " 	AND (isnull(SUBSTRING(G.HORA_HASTA,1,2),9) <= 14)"
					End If

					if Trim(strHorario) = "2" Then
						strSql = strSql & " 	AND (isnull(SUBSTRING(G.HORA_HASTA,1,2),9) > 14)"
					End If
					
					if Trim(strTipoGestion) = "1" Then 'Recaudacion
					
						strSql = strSql & " AND (GTG.GESTION_MODULOS IN (1,11))"
					
					ElseIf Trim(strTipoGestion) = "2" Then 'Notificacion
					
						strSql = strSql & " AND (GTG.GESTION_MODULOS = 13)"
					
					Else 
					
						strSql = strSql & " AND (GTG.GESTION_MODULOS IN (1,11,13))"
						
					End If

					strSql = strSql & " AND  G.ID_GESTION = C.ID_ULT_GEST_GENERAL"

					strSql = strSql & " AND C.ESTADO_DEUDA IN (1,2,7,8) "
					strSql = strSql & " AND C.COD_CLIENTE in (" & intCOD_CLIENTE & ")"

                  

					if Trim(str_empresa_rec) <> "0" Then
						strSql = strSql & "		   AND (CASE WHEN GC.NUEVA_EMPRESA_REC IS NOT NULL"
						strSql = strSql & " 	    	 THEN GC.NUEVA_EMPRESA_REC "
						strSql = strSql & " 	    	 WHEN ER.EMPRESA_REC IS NOT NULL"
						strSql = strSql & "	  			 THEN ER.EMPRESA_REC"
						strSql = strSql & " 	 		 WHEN DD.COMUNA IS NOT NULL"
						strSql = strSql & " 	  		 THEN UPPER(DD.COMUNA)"
						strSql = strSql & " 	 		 ELSE 'NO DEFINIDA'"
						strSql = strSql & " 		END) = '" & str_empresa_rec & "'"
					End if
					
					''''''''''''''Response.WRITE strSql

					if strSql <> "" then
						set rsDet=Conn.execute(strSql)

						if not rsDet.eof then
							intReg = 0
							intTotalCasos = 0
							intRNP_am = 0
							intRR_am = 0
							intRP_am = 0
							intRNP_pm = 0
							intRR_pm = 0
							intRP_pm = 0

							do while not rsDet.eof
							
								intReg = intReg + 1

								intRNP_am = rsDet("RNP_AM")
								intRR_am = rsDet("RR_AM")
								intRP_am = rsDet("RP_AM")
								intRNP_pm = rsDet("RNP_PM")
								intRR_pm = rsDet("RR_PM")
								intRP_pm = rsDet("RP_PM")

								intDNP_am = rsDet("DNP_AM")
								intDR_am = rsDet("DR_AM")
								intDP_am = rsDet("DP_AM")
								intDNP_pm = rsDet("DNP_PM")
								intDR_pm = rsDet("DR_PM")
								intDP_pm = rsDet("DP_PM")
								
								rsDet.movenext
							loop
						end if


						intotalDir_AM = intDNP_am + intDR_am + intDP_am
						intotalDir_PM = intDNP_pm + intDR_pm + intDP_pm
						intTotalDir   = intotalDir_AM + intotalDir_PM

						intotalRutas_AM = intRNP_am + intRR_am + intRP_am
						intotalRutas_PM = intRNP_pm + intRR_pm + intRP_pm
						intTotalRutas   = intotalRutas_AM + intotalRutas_PM
						
						intRNP = intRNP_am + intRNP_pm
						intRR = intRR_am + intRR_pm
						intRP = intRP_am + intRP_pm
						
						intDNP = intDNP_am + intDNP_pm
						intDR = intDR_am + intDR_pm
						intDP = intDP_am + intDP_pm
          
						%>

                <table border="0" class="intercalado" align="center">

						<%if intReg > 0 then             %>
                        <thead>
						    <tr >
									<td>ESTADO</td>
									<td align="center" >TOTAL DOCUMENTOS</td>
									<td align="center">TOTAL DIRECCIONES</td>
								</tr>
							</thead>
							<tbody>
								<tr >
								<td>NO PROCESADO</td>
								<td align="center"><%=FN(intRNP,0)%></td>
								
								<td align="center"><%=FN(intDNP,0)%></td>
								</tr>

								<tr >
								<td>RECHAZADO</td>

								<td align="center"><%=FN(intRR,0)%></td>
								
								<td align="center"><%=FN(intDR,0)%></td>
								</tr>

								<tr >
								<td>PROCESADO</td>
								<td align="center"><%=FN(intRP,0)%></td>
								
								<td align="center"><%=FN(intDP,0)%></td>
								</tr>

								<tr class="totales">
									<td>TOTALES</td>
									<td align="center"><%=FN(intTotalRutas,0)%></td>
									
									<td align="center"><%=FN(intTotalDir,0)%></td>
								</tr>
							</tbody>
						<%Else%>
							<thead>

								<tr >
									<td ALIGN="CENTER" Colspan = "18" bgcolor="#ffffff">&nbsp;</td>
								</tr>

								<tr class="estilo_columna_individual">
									<td ALIGN="CENTER" Colspan = "4">NO EXISTEN DOCUMENTOS ENRUTADOS SEGÚN PARÁMETROS DE BÚSQUEDA</td>
								</tr>


							</thead>

						<%end if%>

					<%end if%>

	</table>
	<br>

	<INPUT id="hdnScrollPos" type=hidden NAME="hdnScrollPos">

	<div id="divScroll" style="overflow:auto; width:100%; height:100 px" onscroll="saveScroll()">

	<table width="100%" border="0" bordercolor="#000000" class="intercalado" align="center">
		<thead>
		<%if (intotalRutas_AM > 0 and Trim(strHorario) = "1") or (intotalRutas_PM > 0 and Trim(strHorario) = "2") or ((intotalRutas_AM > 0 or intotalRutas_PM > 0 ) and Trim(strHorario) = "0") then %>

		<tr bordercolor="#999999"  bgcolor="#<%=session("COLTABBG")%>" class="Estilo13">
			<td rowspan = "2" Width = "230">EMPRESA RECAUDADORA</td>
			<td align = "Middle" Colspan = "4">DOCUMENTOS</td>
			<td align = "Middle" Colspan = "4">DIRECCIONES</td>
		</tr>
		
		<tr bordercolor="#999999"  bgcolor="#<%=session("COLTABBG")%>" class="Estilo13">
			<td align="center" width = "10%" >NO PROCESADOS</td>
			<td align="center" width = "10%" >PROCESADOS</td>
			<td align="center" width = "10%" >RECHAZADOS</td>
			<td align="center" width = "10%" bgcolor='#6e6e6e' >TOTAL</td>
			<td align="center" width = "10%" >NO PROCESADOS</td>
			<td align="center" width = "10%" >PROCESADOS</td>
			<td align="center" width = "10%" >RECHAZADOS</td>
			<td align="center" width = "10%" bgcolor='#6e6e6e'>TOTAL</td>
		</tr>

		<%end if%>
		</thead>
		
		
		<tbody>
<%

                       '''''''''''''''''''''  reporte informe empresa 
                      
                       '''''''''''''''''''''  reporte informe empresa 

					strSql = "SELECT "
 					strSql = strSql & " (CASE WHEN GC.NUEVA_EMPRESA_REC IS NOT NULL"
 					strSql = strSql & " 	  THEN GC.NUEVA_EMPRESA_REC"
 					strSql = strSql & " 	  WHEN ER.EMPRESA_REC IS NOT NULL"
 					strSql = strSql & " 	  THEN ER.EMPRESA_REC"
 					strSql = strSql & " 	  WHEN DD.COMUNA IS NOT NULL"
 					strSql = strSql & " 	  THEN UPPER(DD.COMUNA)"
 					strSql = strSql & " 	  ELSE 'NO DEFINIDA'"
 					strSql = strSql & "  END) AS EMPRESA_RECAUDADORA,"
					
 					strSql = strSql & " SUM(CASE WHEN ISNULL(GC.ESTADO_RUTA,0) = 0 THEN 1 ELSE 0 END) AS RNP,"
 					strSql = strSql & " SUM(CASE WHEN ISNULL(GC.ESTADO_RUTA,0) = 1 AND DATEDIFF(D,G.FECHA_COMPROMISO,'"& termino & "') = 0 THEN 1 ELSE 0 END) AS RP,"
 					strSql = strSql & " SUM(CASE WHEN ISNULL(GC.ESTADO_RUTA,0) = 2 AND DATEDIFF(D,G.FECHA_COMPROMISO,'"& termino & "') = 0 THEN 1 ELSE 0 END) AS RR,"
					if Trim(strHorario) = "1" Then ''' AM
 					strSql = strSql & " COUNT(DISTINCT (CASE WHEN ISNULL(GC.ESTADO_RUTA,0) = 0 AND (isnull(SUBSTRING(G.HORA_HASTA,1,2),9) <= 14) THEN ISNULL(G.ID_DIRECCION_COBRO_DEUDOR,G.ID_FORMA_RECAUDACION) END)) AS DNP,"
                    elseif Trim(strHorario) = "2"  then ''' PM
                    strSql = strSql & " COUNT(DISTINCT (CASE WHEN ISNULL(GC.ESTADO_RUTA,0) = 0 AND (isnull(SUBSTRING(G.HORA_HASTA,1,2),9) > 14) THEN ISNULL(G.ID_DIRECCION_COBRO_DEUDOR,G.ID_FORMA_RECAUDACION) END)) AS DNP,"
                    else 
                    strSql = strSql & " COUNT(DISTINCT (CASE WHEN ISNULL(GC.ESTADO_RUTA,0) = 0  THEN ISNULL(G.ID_DIRECCION_COBRO_DEUDOR,G.ID_FORMA_RECAUDACION) END)) AS DNP,"
                    end if
 					strSql = strSql & " COUNT(DISTINCT CASE WHEN ISNULL(GC.ESTADO_RUTA,0) = 1 AND DATEDIFF(D,G.FECHA_COMPROMISO,'"& termino & "') = 0 THEN ISNULL(G.ID_DIRECCION_COBRO_DEUDOR,G.ID_FORMA_RECAUDACION) END) AS DP," 
 					strSql = strSql & " COUNT(DISTINCT CASE WHEN ISNULL(GC.ESTADO_RUTA,0) = 2 AND DATEDIFF(D,G.FECHA_COMPROMISO,'"& termino & "') = 0 THEN ISNULL(G.ID_DIRECCION_COBRO_DEUDOR,G.ID_FORMA_RECAUDACION) END) AS DR"
                    strSql = strSql & " ,CASE WHEN GTG.GESTION_MODULOS = 13 THEN  ( isnull(SUM(CASE WHEN ISNULL(GC.ESTADO_RUTA,0) = 1 AND DATEDIFF(D,G.FECHA_COMPROMISO,'"& termino & "') = 0 THEN 1 ELSE 0 END),0)) ELSE 0 END NOT_DOC_P"
                    strSql = strSql & " ,CASE WHEN GTG.GESTION_MODULOS = 13 THEN  ( COUNT(DISTINCT CASE WHEN ISNULL(GC.ESTADO_RUTA,0) = 1 AND DATEDIFF(D,G.FECHA_COMPROMISO,'"& termino & "') = 0 THEN ISNULL(G.ID_DIRECCION_COBRO_DEUDOR,G.ID_FORMA_RECAUDACION) END)) ELSE 0 END NOT_DIR_P "
                    strSql = strSql & " ,CASE WHEN GTG.GESTION_MODULOS in (1,11) Then (isnull(SUM(CASE WHEN ISNULL(GC.ESTADO_RUTA,0) = 1 AND DATEDIFF(D,G.FECHA_COMPROMISO,'"& termino & "') = 0 THEN 1 ELSE 0 END),0)) ELSE 0 END REC_DOC_P"
                    strSql = strSql & " ,CASE WHEN GTG.GESTION_MODULOS in (1,11) Then (COUNT(DISTINCT CASE WHEN ISNULL(GC.ESTADO_RUTA,0) = 1 AND DATEDIFF(D,G.FECHA_COMPROMISO,'"& termino & "') = 0 THEN ISNULL(G.ID_DIRECCION_COBRO_DEUDOR,G.ID_FORMA_RECAUDACION) END))ELSE 0 END REC_DIR_P "


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

					if Trim(strTipoGestion) = "1" Then 'Recaudacion
					
						strSql = strSql & " AND (GTG.GESTION_MODULOS IN (1,11))"
					
					ElseIf Trim(strTipoGestion) = "2" Then 'Notificacion
					
						strSql = strSql & " AND (GTG.GESTION_MODULOS = 13)"
					
					Else 
					
						strSql = strSql & " AND (GTG.GESTION_MODULOS IN (1,11,13))"
						
					End If
					
					if Trim(strHorario) = "1" Then
						strSql = strSql & " 	AND (isnull(SUBSTRING(G.HORA_HASTA,1,2),9) <= 14)"
					End If

					if Trim(strHorario) = "2" Then
						strSql = strSql & " 	AND (isnull(SUBSTRING(G.HORA_HASTA,1,2),9) > 14)"
					End If
					
					strSql = strSql & " AND  G.ID_GESTION = C.ID_ULT_GEST_GENERAL"
					
					strSql = strSql & " AND ((ISNULL(GC.ESTADO_RUTA,0) = 0 )"  
					strSql = strSql & " 	  OR (ISNULL(GC.ESTADO_RUTA,0) = 1 AND DATEDIFF(D,G.FECHA_COMPROMISO,'"& termino & "') = 0)"
					strSql = strSql & " 	  OR (ISNULL(GC.ESTADO_RUTA,0) = 2 AND DATEDIFF(D,G.FECHA_COMPROMISO,'"& termino & "') = 0 ))"
					
					strSql = strSql & " AND C.ESTADO_DEUDA IN (1,2,7,8) "
					strSql = strSql & " AND C.COD_CLIENTE in (" & intCOD_CLIENTE & ")"
					
					if Trim(str_empresa_rec) <> "0" Then
						strSql = strSql & "		   AND (CASE WHEN GC.NUEVA_EMPRESA_REC IS NOT NULL"
						strSql = strSql & " 	    	 THEN GC.NUEVA_EMPRESA_REC"
						strSql = strSql & " 	    	 WHEN ER.EMPRESA_REC IS NOT NULL"
						strSql = strSql & "	  			 THEN ER.EMPRESA_REC"
						strSql = strSql & " 	 		 WHEN DD.COMUNA IS NOT NULL"
						strSql = strSql & " 	  		 THEN UPPER(DD.COMUNA)"
						strSql = strSql & " 	 		 ELSE 'NO DEFINIDA'"
						strSql = strSql & " 		END) = '" & str_empresa_rec & "'"
					End if

                    
 					strSql = strSql & " GROUP BY GTG.GESTION_MODULOS,(CASE WHEN GC.NUEVA_EMPRESA_REC IS NOT NULL THEN GC.NUEVA_EMPRESA_REC WHEN ER.EMPRESA_REC IS NOT NULL THEN ER.EMPRESA_REC WHEN DD.COMUNA IS NOT NULL THEN UPPER(DD.COMUNA) ELSE 'NO DEFINIDA' END) "

 					strSql = strSql & " ORDER BY EMPRESA_RECAUDADORA ASC "
                        

						'if not rsDetEmpresa.eof then

                        intTotalEmpresas = 0


                       set rsDetEmpresa=Conn.execute(strSql)

                        do while not rsDetEmpresa.eof
                        

                        dim Notificaciones
                        dim Recaudacion
			
                        Notificaciones = "<br/>NOTIFICACIÓN:<b>" & rsDetEmpresa("NOT_DOC_P") & "</b>"
                        NotificacionesDir = "<br/>NOTIFICACIÓN:<b>" & rsDetEmpresa("NOT_DIR_P") & "</b>"
                
                        
				        Recaudacion = "RECAUDACIÓN:<b>" & rsDetEmpresa("REC_DOC_P") &"</b>"
                        RecaudacionDir = "RECAUDACIÓN:<b>" & rsDetEmpresa("REC_DIR_P") &"</b>"
                



							intCasosRNP = rsDetEmpresa("RNP")
							intCasosRP = rsDetEmpresa("RP")
							intCasosRR = rsDetEmpresa("RR")

							intCasosDNP = rsDetEmpresa("DNP")
							intCasosDP = rsDetEmpresa("DP")
							intCasosDR = rsDetEmpresa("DR")
							
							intTotalEmpresas = intTotalEmpresas + 1
							intTotalCasosEmpresa = intCasosRNP + intCasosRP + intCasosRR
							intTotalCasosEmpresaDir = intCasosDNP + intCasosDP + intCasosDR

							intTotalCasosRNP = intTotalCasosRNP + intCasosRNP
							intTotalCasosRP = intTotalCasosRP + intCasosRP
							intTotalCasosRR = intTotalCasosRR + intCasosRR

							intTotalCasosDNP = intTotalCasosDNP + intCasosDNP
							intTotalCasosDP = intTotalCasosDP + intCasosDP
							intTotalCasosDR = intTotalCasosDR + intCasosDR

							intTotalCasosEmpresas = intTotalCasosEmpresas + intTotalCasosEmpresa
							intTotalCasosEmpresasDir = intTotalCasosEmpresasDir + intTotalCasosEmpresaDir

				%>


							<tr >

									<td Width = "230" align="LEFT"><%=rsDetEmpresa("EMPRESA_RECAUDADORA")%></td>

									<td Width = "207" align="RIGHT">
                                    <% if intCasosRNP > 0 then %> 
															<A HREF="Informe_Ruta.asp?cmb_estadoProceso=<%=0%>&cmb_Horario=<%=strhorario%>&termino=<%=termino%>&cmb_empresa_rec=<%=rsDetEmpresa("EMPRESA_RECAUDADORA")%>&CB_CLIENTE=<%=intCod_Cliente%>&cmb_tipogestion=<%=strTipoGestion%>&Busqueda=<%=strTipoBusqueda%>">
															<acronym title="Llevar a pantalla de selección"><%=intCasosRNP%></acronym></A>
			                        <%else %>												
                                    <%=intCasosRNP%>
                                    <% end if %>
									</td>

									<td Width = "193" align="RIGHT" title="<%=Recaudacion & Notificaciones%>">
                                      <% if intCasosRP > 0 then %> 
															<A HREF="Informe_Ruta.asp?cmb_estadoProceso=<%=1%>&cmb_Horario=<%=strhorario%>&termino=<%=termino%>&cmb_empresa_rec=<%=rsDetEmpresa("EMPRESA_RECAUDADORA")%>&CB_CLIENTE=<%=intCod_Cliente%>&cmb_tipogestion=<%=strTipoGestion%>&Busqueda=<%=strTipoBusqueda%>">
															<acronym title="Llevar a pantalla de selección"><%=intCasosRP%></acronym>
															</A>
                                    <%else %>												
                                    <%=intCasosRP%>
                                    <% end if %>
									</td>

									<td Width = "190" align="RIGHT">
                                    <% if intCasosRR > 0 then %> 
															<A HREF="Informe_Ruta.asp?cmb_estadoProceso=<%=2%>&cmb_Horario=<%=strhorario%>&termino=<%=termino%>&cmb_empresa_rec=<%=rsDetEmpresa("EMPRESA_RECAUDADORA")%>&CB_CLIENTE=<%=intCod_Cliente%>&cmb_tipogestion=<%=strTipoGestion%>&Busqueda=<%=strTipoBusqueda%>">
															<acronym title="Llevar a pantalla de selección"><%=intCasosRR%></acronym>
															</A>
                                    <%else %>												
                                    <%=intCasosRR%>
                                    <% end if %>
									</td>

									<td align="RIGHT"><%=intTotalCasosEmpresa%></td>


									<td Width = "207" align="RIGHT">
                                     <% if intCasosRR > 0 then %> 
															<A HREF="Informe_Ruta.asp?cmb_estadoProceso=<%=0%>&cmb_Horario=<%=strhorario%>&termino=<%=termino%>&cmb_empresa_rec=<%=rsDetEmpresa("EMPRESA_RECAUDADORA")%>&CB_CLIENTE=<%=intCod_Cliente%>&cmb_tipogestion=<%=strTipoGestion%>&Busqueda=<%=strTipoBusqueda%>">
															<acronym title="Llevar a pantalla de selección"><%=intCasosDNP%></acronym>
															</A>
                                    <%else %>												
                                    <%=intCasosDNP%>
                                    <% end if %>
									</td>

									<td Width = "193" align="RIGHT" title="<%=RecaudacionDir & NotificacionesDir%>">
                                     <% if intCasosDP > 0 then %> 
															<A HREF="Informe_Ruta.asp?cmb_estadoProceso=<%=1%>&cmb_Horario=<%=strhorario%>&termino=<%=termino%>&cmb_empresa_rec=<%=rsDetEmpresa("EMPRESA_RECAUDADORA")%>&CB_CLIENTE=<%=intCod_Cliente%>&cmb_tipogestion=<%=strTipoGestion%>&Busqueda=<%=strTipoBusqueda%>">
															<acronym title="Llevar a pantalla de selección"><%=intCasosDP%></acronym>
															</A>
                                    <%else %>												
                                    <%=intCasosDP%>
                                    <% end if %>
									</td>

									<td Width = "190" align="RIGHT">
                                     <% if intCasosDR > 0 then %> 
															<A HREF="Informe_Ruta.asp?cmb_estadoProceso=<%=2%>&cmb_Horario=<%=strhorario%>&termino=<%=termino%>&cmb_empresa_rec=<%=rsDetEmpresa("EMPRESA_RECAUDADORA")%>&CB_CLIENTE=<%=intCod_Cliente%>&cmb_tipogestion=<%=strTipoGestion%>&Busqueda=<%=strTipoBusqueda%>">
															<acronym title="Llevar a pantalla de selección"><%=intCasosDR%></acronym>
															</A>
                                    <%else %>												
                                    <%=intCasosDR%>
                                    <% end if %>
									</td>

									<td align="RIGHT"><%=intTotalCasosEmpresaDir%></td>

							</tr>

							<%	rsDetEmpresa.movenext
							loop
						
'					end if

                      '  rsDetEmpresa.close
					'	set rsDetEmpresa=nothing
						   ' response.write "intotalRutas_AM = " & intotalRutas_AM
                           ' response.write "strHorario = " & strHorario
                           ' response.write "intotalRutas_PM = " & intotalRutas_PM
                           ' response.write "intTotalCasos = " & intTotalCasos


						 if (intotalRutas_AM > 0 and Trim(strHorario) = "1") or (intotalRutas_PM > 0 and Trim(strHorario) = "2") or (intReg > 0 and Trim(strHorario) = "0")then%>
						</tbody>

						<tr bordercolor="#999999"  bgcolor="#<%=session("COLTABBG")%>" class="totales">
								<td Width = "230" >TOTAL SEGMENTACIÓN <%=intTotalEmpresas%></td>
								<td Width = "207" align="RIGHT"><%=intTotalCasosRNP%></td>
								<td Width = "193" align="RIGHT"><%=intTotalCasosRP%></td>
								<td Width = "190" align="RIGHT"><%=intTotalCasosRR%></td>
								<td align="RIGHT"><%=intTotalCasosEmpresas%></td>
								<td Width = "207" align="RIGHT"><%=intTotalCasosDNP%></td>
								<td Width = "193" align="RIGHT"><%=intTotalCasosDP%></td>
								<td Width = "190" align="RIGHT"><%=intTotalCasosDR%></td>
								<td align="RIGHT"><%=intTotalCasosEmpresasDir%></td>
						</tr>

						<%end if%>

	</table>

	</div>
	
	<div id="refresca_resumen"></div>

</form>

</body>
</html>
<%cerrarscg()%>

