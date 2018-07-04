<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>

<!DOCTYPE html>
<html>
<head>
<%
    Response.CodePage=65001
    Response.charset ="utf-8"
    
    CodigoCliente 	= request("CodigoCliente")
    RutDeudor 		= request("RutDeudor")
    CodigoUsuario	= request("CodigoUsuario")
    IdBlockVisibles	= request("IdBlockVisibles")

    'DECLARAMOS CODIGOS DE ATRIBUTOS PARA EL BLOQUE FORMA DE PAGO
    CAMPO_CUANDO_PAGA = 1
    CAMPO_DIA_HORA_PAGO_ESPECIAL = 2
    CAMPO_COMO_PAGA = 3
    CAMPO_EXIGENCIAS_ESPECIALES = 4
    CAMPO_PORTALES_CONSULTA_PAGO = 5
    CAMPO_PRE_ENVIO_FACTURA_REGION = 6
    CAMPO_PAGO_FACTORING = 7
    CAMPO_PAGO_SOLO_CLIENTE = 8
    
    'DECLARAMOS CODIGOS DE DOMINIOS CAMPOS
    DOMINIO_CAMPO_SI = 1
    DOMINIO_CAMPO_NO = 2
    DOMINIO_CAMPO_SA = 3
    DOMINIO_CAMPO_VALE_VISTA = 4
    DOMINIO_CAMPO_TRANSFERENCIA = 5
    DOMINIO_CAMPO_CHEQUE = 6
    DOMINIO_CAMPO_TEXTO = 7
        
    ' DECLARAMOS POSICIONES DE COLUMNAS PARA EL ArrDatosCampo
    COLUMNA_FECHA = 0
    COLUMNA_ID_CAMPO = 1
    COLUMNA_ID_DOMINIO_CAMPO = 2
    COLUMNA_VALOR = 3
    COLUMNA_TEXTO = 4
%>
    <meta http-equiv="X-UA-Compatible" content="IE=edge,chrome=1"></meta>
    <meta charset="utf-8"></meta>
    <title>Ficha deudor</title>
    
    <script src="../Javascripts/FichaDeudor/jquery-1.9.0.min.js"></script>
    <script src="../Javascripts/FichaDeudor/jquery-ui.min.js"></script>
        
    <link rel="stylesheet" href="../Css/FichaDeudor/jquery-ui.min.css">
            
    <script src="../Javascripts/FichaDeudor/metodos.js"></script>
    
    <link rel="stylesheet" type="text/css" href="../Css/FichaDeudor/normalize.css" media="screen" />
        
    <style type="text/css">
        
        html, body {
            height:100%;
            
            margin:0;
            padding:0;
            
            font-size: 12px;
        }

        .Visible {
            display: block;
        }

        .NoVisible {
            display: none;
        }

        /* ESTILO DE LA PAGINA */

        #ContainerPrincipal {
            width: 100%;
            height: calc(100% - 40px);
            
            position: relative;
        }

        #Menu {
            width: 90%;
            height: 30px;
            
            background-color: #FAFAFA;
            
            position: fixed;
            z-index: 10;
            
            padding-left: 5%;
            padding-top: 10px;
            padding-bottom: 10px;
        }
        
        .ui-widget button {
            font-size:11px;
            font-family: "Trebuchet MS";
        }
        
        .BotonHover {
            background-image: url("../Imagenes/fondo_botones_hover.jpg");
            background-repeat:repeat-x;
            height: 23px;
            border:0px solid #ccc;
            color:#fff;
            font-size:11px;
            font-style:normal;
            font-family: "Trebuchet MS";
            border-radius: 4px;    
        }
        
        .Boton{
            background-image: url("../Imagenes/fondo_botones.jpg");
            background-repeat:repeat-x;
            height: 23px;
            border:0px solid #ccc;
            color:#fff;
            font-size:11px;
            font-style:normal;
            font-family: "Trebuchet MS";
            border-radius: 4px;
        }

        .Dialog {
            font-size: 12px;
            text-align: center;
        }
        
        .Dialog div.Buttons {
            width: 100%;
            text-align: right;
            vertical-align: top;
            float: left;
        }
        
        .Dialog div.ButtonsGuardarCambios {
        
            width: 50%;
        
        }
        
        .Dialog div.Message {
            width: 50%;
            text-align: left;
            vertical-align: top;
            float: left;
        }
        
            .Dialog textarea {
                width: 300px;
                height: 185px;
                
                margin-top: 10px;
                margin-bottom: 10px;
            }

        #ContainerBloques {
            width: 100%;
            
            position: absolute;
            top: 60px;
        }
        
        .Block {
            width: calc(100% - 2px);
            height: auto;
        }

            .Block table {					
                position: relative;
                left: 5%;
                right: 5%;
                
                border-style: solid;
                border-width: 1px;
                
                margin-bottom: 10px;
            }
            
                table .Title {		
                    background-color: #989898;
                    
                    font-size: 12px;
                    font-weight: bold;
                    font-family: Tahoma;
                    color: #FFFFFF;
                    
                    padding-top: 5px;
                    padding-bottom: 5px;
                    padding-left: 10px;
                    
                    border-color: #000000;
                }
                
                table .TitleAllHidden {
                
                    text-align: center;
                
                }
            
                table .TitleIcon {						
                    margin-right: 10px;
                }
            
                table tr {
                    width: 100%;
                }
            
                table td {
                    padding-left: 1%;
                    padding-top: 5px;
                    padding-bottom: 5px;
                    
                    border-style: solid;
                    border-width: 1px;
                }
                
                table .NombreAtributo {
                    width: 25%;
                    background-color: #c9def2;
                    
                    padding-left: 1%;
                    padding-top: 5px;
                    padding-bottom: 5px;
                    font-size: 11px;
                    font-family: Tahoma;
                }
                
                table .Atributo {
                    width: 50%;
                    text-align: left;
                    font-size: 11px;
                    font-family: Tahoma;
                }
                
                table td.InformacionDeudor {
                    width: 75%;
                }
                    
                    .Atributo input[type=text] {
                        width: 97%;
                    }
                    
                    .Atributo input[type=radio] {
                        margin-left: 10px;
                        margin-right: 10px;
                        cursor: pointer;
                    }
                    
                table .AtributoGris {
                    background-color: #F0F0F0;
                    font-size: 11px;
                    font-family: Tahoma;
                }
                                        
                table .FechaModificacion {
                    text-align: center;
                    font-size: 11px;
                    font-family: Tahoma;
                }
                
                table .BotonHistorial {
                
                    width: 10%;
                    
                    text-align: center;
                    
                    padding-left: 0px;
                }
                
                    table .BotonHistorial img {
                        width: 15px;
                    
                        cursor: pointer;
                    }
                    
        .noclose .ui-dialog-titlebar-close
        {
            display:none;
        }
        
        table.HistorialCambios td {
            padding: 5px;
            text-align: center;
            border-style: solid;
            border-width: 1px;
        }
        
        table.HistorialCambios td.Texto {
            text-align: left;
        }
        
        table.HistorialCambios tr.Atributo {
            font-size: 11px;
            font-family: Tahoma;
        }
        
        table.HistorialCambios tr.AtributoGris {
            background-color: #F0F0F0;
            font-size: 11px;
            font-family: Tahoma;
        }
        
        table.HistorialCambios div.Observaciones {
            word-wrap: break-word;
            width: 135px;
            text-align: left;
        }

        .accordion {
            width: 90%;
            margin: 0 auto;
            border-bottom: solid 1px #c4c4c4;
            font-size: 11px;
            font-family: Tahoma;
        }

        .accordion h3 {
            background: #e9e7e7 url(images/arrow-square.gif) no-repeat right -51px;
            padding: 7px 15px;
            margin: 0;
            font: bold 120%/100% Tahoma, sans-serif;
            border: solid 1px #c4c4c4;
            border-bottom: none;
            cursor: pointer;
        }
        .accordion h3:hover {
            background-color: #e3e2e2;
        }
        .accordion h3.active {
            background-position: right 5px;
        }

       .Block .tablaUbicabilidad {					
               margin: 0px;
        }

        .botonsLeft {
            float: left;
        }

         .botonsRight {
            float: right;
        }

       .ui-accordion .ui-accordion-content {    
        border-top: 0;
        overflow:hidden;
        padding: 0;

       }

        .Block .tablaInterior {
            margin-left: -5%;
        }

        .accordion div {
            text-align: left;
        }
    
    </style>
    
    <!--#include file="arch_utils.asp"-->
    <!--#include file="../lib/comunes/rutinas/funciones.inc" -->
    
    <!--[if lt IE 9]>
        <script src="../Javascripts/FichaDeudor/html5shiv.min.js"></script>
    <![endif]-->
</head>

<body>
<form action="FichaDeudor.asp" method="post">
<input type="hidden" id="CodigoCliente" name="CodigoCliente" value="<%=CodigoCliente%>"></input>
<input type="hidden" id="RutDeudor" name="RutDeudor" value="<%=RutDeudor%>"></input>
<input type="hidden" id="CodigoUsuario" name="CodigoUsuario" value="<%=CodigoUsuario%>"></input>
<input type="hidden" id="IdBlockVisibles" name="IdBlockVisibles" value="<%=IdBlockVisibles%>" />
    <article id="ContainerPrincipal">
    
        <nav id="Menu">
            <div class="botonsLeft">
                <button id="ButtonDeudor" class="Boton" value="Deudor">Deudor</button>
                <button id="ButtonPago" class="Boton" value="Pago">Pago</button>
                <button id="ButtonContabilidad" class="Boton" value="Contactabilidad">Contactabilidad</button>
                <button id="ButtonCliente" class="Boton" style="display: none;" value="Cliente">Cliente</button>
            </div>
            <div class="botonsRight">
                <button id="ButtonGuardar" class=" Boton" value="Guardar">Guardar</button>
            </div>
        </nav>	
        
        <div id="DialogHistorial" class="Dialog">
        </div>
        
        <div id="DialogObservacion" title="Observación" class="Dialog">
            <textarea id="TxaObservacion" maxlength="300"></textarea><br>
            <div class="Message">
                M&aacute;x. 300 caracteres.
            </div>
            <div class="Buttons ButtonsGuardarCambios">
                <button id="ButtonGuardarFicha" class="Boton">Guardar</button>
                <button id="ButtonCancelarGuardarFicha" class="Boton">Cancelar</button>
            </div>
        </div>
        
        <div id="DialogNoCambios" title="Atención" class="Dialog">
            <p>No se han realizado cambios en la ficha deudor.</p>
            <div class="Buttons">
                <button id="ButtonAceptarNoCambios" class="Boton">Aceptar</button>
            </div>
        </div>
        
        <div id="DialogCambiosGuardados" title="Atención" class="Dialog">
            <p>Los cambios en la ficha deudor se han guardado.</p>
            <div class="Buttons">
                <button id="ButtonAceptarCambiosGuardados" class="Boton">Aceptar</button>
            </div>
        </div>
            
        <section id="ContainerBloques">
        
            <%
                AbrirSCG1()
                
                    strSql = "SELECT MIN(FECHA_ESTADO) AS FECHA_ESTADO,  " 
                        strSql = strSql & "MIN(FECHA_ESTADO_CUSTODIO) AS FECHA_ESTADO_CUSTODIO, "
                        strSql = strSql & "MAX(FECHA_VENC) AS FECHA_VENCIMIENTO, "
                        strSql = strSql & "COUNT(ID_CUOTA) AS TOTAL_DOCUMENTOS, "
                        strSql = strSql & "SUM(SALDO) AS TOTAL_MONTOS "
                    strSql = strSql & "FROM CUOTA INNER JOIN ESTADO_DEUDA "
                        strSql = strSql & "ON CUOTA.ESTADO_DEUDA = ESTADO_DEUDA.CODIGO "
                    strSql = strSql & "WHERE "
                        strSql = strSql & "ESTADO_DEUDA.ACTIVO = '1' "
                        strSql = strSql & "AND CUOTA.RUT_DEUDOR = '"& RutDeudor &"' "
                        strSql = strSql & "AND CUOTA.COD_CLIENTE = '"& CodigoCliente &"' "
                        
                    set RsDeudor = Conn1.execute(strSql)
                    
                    if not RsDeudor.eof then
                    
                        if(RsDeudor("FECHA_ESTADO_CUSTODIO") <> "") then
                            strFechaEstado = FormatDateTime(RsDeudor("FECHA_ESTADO_CUSTODIO"), 2)
                        elseif(RsDeudor("FECHA_ESTADO") <> "") then
                            strFechaEstado = FormatDateTime(RsDeudor("FECHA_ESTADO"), 2)
                        end if
                        
                        if(strFechaEstado <> "")then
                            DiasDeCobranza = DateDiff("d", strFechaEstado, now)
                        else
                            DiasDeCobranza = ""
                        end if
                        
                        VencimientoMayor = RsDeudor("FECHA_VENCIMIENTO")
                        DiasVencimiento = DateDiff("d", VencimientoMayor, now) 
                        TotalDocumentos = RsDeudor("TOTAL_DOCUMENTOS")
                        TotalMontos = RsDeudor("TOTAL_MONTOS")
                        
                    end if
                        
                    strSql = "SELECT TOP 1 " 
                        strSql = strSql & "DONDE_PAGA = UPPER(ISNULL(UPPER(RE.NOMBRE+' '+RE.UBICACION), upper(REPLACE(isnull(DD.CALLE,'') + ' ' + isnull(DD.NUMERO,'') + ' ' + isnull(DD.RESTO,'') + ' ' + isnull(DD.COMUNA,''),'  ',' ') ))), "
                        strSql = strSql & "HORARIO_PAGO_DESDE = G.HORA_DESDE, "
                        strSql = strSql & "HORARIO_PAGO_HASTA = G.HORA_HASTA, "
                        strSql = strSql & "ULTIMA_FORMA_PAGO = FPC.DESC_FORMA_PAGO, "
                        strSql = strSql & "DOCUMENTOS_PARA_PAGO = ISNULL(DOC_GESTION,'') "

                    strSql = strSql & "FROM GESTIONES G INNER JOIN GESTIONES_TIPO_GESTION GTG "
                        strSql = strSql & "ON G.COD_CATEGORIA=GTG.COD_CATEGORIA "
                            strSql = strSql & "AND G.COD_SUB_CATEGORIA=GTG.COD_SUB_CATEGORIA "
                            strSql = strSql & "AND G.COD_GESTION=GTG.COD_GESTION "
                            strSql = strSql & "AND G.COD_CLIENTE=GTG.COD_CLIENTE "
                                                                                  
                        strSql = strSql & "LEFT JOIN FORMA_RECAUDACION RE "
                            strSql = strSql & "ON RE.ID_FORMA_RECAUDACION=G.ID_FORMA_RECAUDACION "
                        strSql = strSql & "LEFT JOIN FORMA_PAGO_CLIENTE FPC "
                            strSql = strSql & "ON G.FORMA_PAGO=FPC.ID_FORMA_PAGO AND G.COD_CLIENTE=FPC.COD_CLIENTE "
                        strSql = strSql & "LEFT JOIN DEUDOR_DIRECCION DD "
                            strSql = strSql & "ON DD.ID_DIRECCION=G.ID_DIRECCION_COBRO_DEUDOR " 

                        strSql = strSql & "WHERE G.COD_CLIENTE = '" & CodigoCliente & "' "
                            strSql = strSql & "AND G.RUT_DEUDOR = '" & RutDeudor & "' "
                            strSql = strSql & "AND GTG.GESTION_MODULOS IN (11,1) "

                        strSql = strSql & "ORDER BY G.ID_GESTION DESC"
                        
                    set RsDeudor = Conn1.execute(strSql)
                    
                    if not RsDeudor.eof then
                    
                        DondePaga = RsDeudor("DONDE_PAGA")
                        HorarioPagoDesde = RsDeudor("HORARIO_PAGO_DESDE")
                        HorarioPagoHasta = RsDeudor("HORARIO_PAGO_HASTA")
                        UltimaFormaPago = RsDeudor("ULTIMA_FORMA_PAGO")
                        DocumentosParaPago = RsDeudor("DOCUMENTOS_PARA_PAGO")
                    
                    else
                    
                        DondePaga = "SIN INFORMACION"
                        HorarioPagoDesde = "SIN INFORMACION"
                        HorarioPagoHasta = "SIN INFORMACION"
                        UltimaFormaPago = "SIN INFORMACION"
                        DocumentosParaPago = "SIN INFORMACION"
                    
                    end if
                    
                    strSql = "SELECT VC.IdCampo, DC.IdDominioCampo, DC.Valor, VC.IsDefault "
                    strSql = strSql & "FROM DominioCampo DC INNER JOIN ValorCampo VC "
                    strSql = strSql & "ON DC.IdDominioCampo = VC.IdDominioCampo "
                    strSql = strSql & "ORDER BY VC.IdCampo, VC.Orden"
                    
                    set rsDominioCampo = Conn1.execute(strSql)
                    arrDominioCampo = rsDominioCampo.getRows()
                    
                    maxIDominioCampo = UBound(ArrDominioCampo,2)
                        
                
                    strSql = "SELECT  FD.Fecha, "
                        strSql = strSql & "FD.IdCampo, "
                        strSql = strSql & "FD.IdDominioCampo, "
                        strSql = strSql & "DC.Valor, "
                        strSql = strSql & "VCT.Texto "
                    strSql = strSql & "FROM    dbo.FichaDeudor FD "
                    strSql = strSql & "INNER JOIN ( SELECT MAX(Fecha) AS Fecha, "
                    strSql = strSql & "					IdCampo, "
                    strSql = strSql & "					CodigoCliente, "
                    strSql = strSql & "					RutDeudor "
                    strSql = strSql & "				 FROM   dbo.FichaDeudor "
                    strSql = strSql & "				 GROUP BY IdCampo, "
                    strSql = strSql & "					CodigoCliente, "
                    strSql = strSql & "					RutDeudor "
                    strSql = strSql & "			   ) MFD ON FD.IdCampo = MFD.IdCampo "
                    strSql = strSql & "			   			AND FD.CodigoCliente = MFD.CodigoCliente "
                    strSql = strSql & "			   			AND FD.RutDeudor = MFD.RutDeudor "
                    strSql = strSql & "						AND FD.Fecha = MFD.Fecha "
                    strSql = strSql & "INNER JOIN dbo.DominioCampo DC ON FD.IdDominioCampo = DC.IdDominioCampo "
                    strSql = strSql & "LEFT JOIN ValorCampoTexto VCT ON FD.IdFichaDeudor = VCT.IdFichaDeudor "
                    strSql = strSql & "WHERE FD.CodigoCliente = '" & CodigoCliente & "' AND FD.RutDeudor = '" & RutDeudor & "' "
                    strSql = strSql & "ORDER BY FD.IdCampo ASC "
                     
                    set rsDatosCampo = Conn1.execute(strSql)
                    
                    if not rsDatosCampo.eof then
                        arrDatosCampo = rsDatosCampo.getRows()
                        
                        maxIDatosCampo = UBound(ArrDatosCampo,2)
                    else
                        Redim arrDatosCampo(0,0)
                        
                        maxIDatosCampo = -1
                    end if


                strSql = "SELECT  TC.Glosa, "
                    strSql = strSql & "MU.Nombre, "
                    strSql = strSql & "U.Dato "
                strSql = strSql & "FROM Ubicabilidad U "
                strSql = strSql & "INNER JOIN MedioUbicabilidad MU "
                strSql = strSql & "ON MU.IdMedioUbicabilidad = U.IdMedioUbicabilidad "
                strSql = strSql & "INNER JOIN TipoContacto TC "
                strSql = strSql & "ON U.IdTipoContacto = TC.IdTipoContacto "
                strSql = strSql & "WHERE U.CodigoCliente = '" & CodigoCliente & "' "
                strSql = strSql & "AND U.RutDeudor= '" & RutDeudor & "' "                 
                strSql = strSql & "ORDER BY TC.Glosa, MU.IdMedioUbicabilidad, U.IdUbicabilidad "

                set rsDatosUbicabilidad = Conn1.execute(strSql)

                if not rsDatosUbicabilidad.eof then
                    arrDatosUbicabilidad = rsDatosUbicabilidad.getRows()
                else
                    Redim arrDatosUbicabilidad(0,0)
                end if
                    
                CerrarSCG1()
            %>
            
            <%
            
                    Function ValidarCampoVacio(valor)
                    
                        if(Trim(valor) <> "") then
                        
                            ValidarCampoVacio = valor
                            
                        else
                        
                            ValidarCampoVacio = "SIN INFORMACION"
                            
                        end if
                        
                    end function
                    
                    
                    
                    Function ValidarIsDefault(valor)
                    
                        if(valor) then
                        
                            ValidarIsDefault = "checked='checked'"
                            
                        else
                        
                            ValidarIsDefault = ""
                            
                        end if
                        
                    end function
                    
                    
                    
                    Function GetValueForFieldFromArrayOfData(arrayOfData, fieldId, domainFieldId, columnToGetValueFrom)
                    
                        GetValueForFieldFromArrayOfData = ""
                        
                        for k = 0 to maxIDatosCampo
                        
                            fieldIdFromArray = arrayOfData(COLUMNA_ID_CAMPO, k)
                            
                            domainFieldIdFromArray = arrayOfData(COLUMNA_ID_DOMINIO_CAMPO, k)
                            
                            if fieldId = fieldIdFromArray and domainFieldId = domainFieldIdFromArray then
                                
                                GetValueForFieldFromArrayOfData = arrayOfData(columnToGetValueFrom, k)
                            
                                exit for
                            
                            end if
                            
                        next
                        
                    End Function
                    
                    
                    
                    Function GetDateTimeFieldFromArrayOfData(fieldId, domainFieldId)
                    
                        dateOfModification = GetValueForFieldFromArrayOfData(ArrDatosCampo, _
                                                                            fieldId, _
                                                                            domainFieldId, _
                                                                            COLUMNA_FECHA)
                                                                            
                        if dateOfModification <> "" then
                        
                            GetDateTimeFieldFromArrayOfData = FormatDateTime(dateOfModification, 2)
                            
                        else
                        
                            GetDateTimeFieldFromArrayOfData = ""
                            
                        end if
                    
                    End Function
                    
                    
                    
                    Function GenerateFieldAndGetDomainFieldIdToUse(fieldId)
                    
                        GenerateFieldAndGetDomainFieldIdToUse = ""
                        
                        hasBeenChecked = false
                        
                        for i = 0 to maxIDominioCampo
                        
                            checked = ""
                            
                            if(ArrDominioCampo(0, I) = fieldId) then
                            
                                if GetValueForFieldFromArrayOfData(ArrDatosCampo, _
                                                                fieldId, _
                                                                ArrDominioCampo(1, I), _
                                                                COLUMNA_ID_DOMINIO_CAMPO) = ArrDominioCampo(1, I) then
                                    
                                    checked = "checked='checked'"
                                    
                                    hasBeenChecked = true
                                    
                                    GenerateFieldAndGetDomainFieldIdToUse = ArrDominioCampo(1, I)
                                        
                                else
                                
                                    if not hasBeenChecked then
                                    
                                        checked = ValidarIsDefault(arrDominioCampo(3, I))
                                    
                                        if (arrDominioCampo(3, I)) then
                                        
                                            GenerateFieldAndGetDomainFieldIdToUse = ArrDominioCampo(1, I)
                                            
                                        end if
                                        
                                    end if
                                    
                                end if
                                
                                Response.Write("<input " & checked & " type='radio' id='Campo_" & fieldId & _
                                            "_" & ArrDominioCampo(1,I) & "' value='" & ArrDominioCampo(1,I) & "' name='Campo_" & _
                                            fieldId & "' />" & ArrDominioCampo(2,I))
                                
                            end if
                            
                        next
                    
                    End Function
            %>
            
            <div id="BlockDeudor" class="Block Visible">
                <table width="90%" height="100%" cellspacing="0" id="TablaDeudor">
                    <tr><td colspan="2" class="Title"><img width="16px" class="TitleIcon" src="../Imagenes/FichaDeudor/icon/deudor.png">INFORMACION DEUDOR</td></tr>
                    <tr>
                        <td class="NombreAtributo">Fecha inicio cobranza</td>
                        <td class="Atributo InformacionDeudor"><%=ValidarCampoVacio(strFechaEstado)%></td>
                    </tr>
                    <tr>
                        <td class="NombreAtributo">Dias de cobranza</td>
                        <td class="AtributoGris InformacionDeudor"><%=ValidarCampoVacio(DiasDeCobranza)%></td>
                    </tr>
                    <tr>
                        <td class="NombreAtributo">Vencimiento mayor</td>
                        <td class="Atributo InformacionDeudor"><%=ValidarCampoVacio(VencimientoMayor) %></td>
                    </tr>
                    <tr>
                        <td class="NombreAtributo">Dias vencimiento</td>
                        <td class="AtributoGris InformacionDeudor"><%=ValidarCampoVacio(DiasVencimiento)%></td>
                    </tr>
                    <tr>
                        <td class="NombreAtributo">Total documentos</td>
                        <td class="Atributo InformacionDeudor"><%=ValidarCampoVacio(TotalDocumentos) %></td>
                    </tr>
                    <tr>
                        <td class="NombreAtributo">Total montos</td>
                        <td class="AtributoGris InformacionDeudor"><%=ValidarCampoVacio(fn(TotalMontos,0)) %></td>
                    </tr>
                    <%	AbrirSCG1()
                    
                            StrSQL = "SELECT CLIENTE.ADIC_1 AS AdicionalUnoCliente, "
                                StrSQL = StrSQL & "CLIENTE.ADIC_2 AS AdicionalDosCliente, "
                                StrSQL = StrSQL & "CLIENTE.ADIC_3 AS AdicionalTresCliente, "
                                StrSQL = StrSQL & "DEUDOR.ADIC_1 AS AdicionalUnoDeudor, "
                                StrSQL = StrSQL & "DEUDOR.ADIC_2 AS AdicionalDosDeudor, "
                                StrSQL = StrSQL & "DEUDOR.ADIC_3 AS AdicionalTresDeudor "
                            StrSQL = StrSQL & "FROM CLIENTE INNER JOIN DEUDOR "
                                StrSQL = StrSQL & "ON CLIENTE.COD_CLIENTE = DEUDOR.COD_CLIENTE "
                            StrSQL = StrSQL & "WHERE CLIENTE.COD_CLIENTE='"& CodigoCliente &"' "
                            StrSQL = StrSQL & "AND DEUDOR.RUT_DEUDOR = '"& RutDeudor &"' "
                            
                            set RsCliente = Conn1.execute(StrSQL)
                            
                            if not RsCliente.eof then
                            
                                AdicionalUnoCliente = RsCliente("AdicionalUnoCliente")
                                AdicionalDosCliente = RsCliente("AdicionalDosCliente")
                                AdicionalTresCliente = RsCliente("AdicionalTresCliente")
                                
                                AdicionalUnoDeudor = RsCliente("AdicionalUnoDeudor")
                                AdicionalDosDeudor = RsCliente("AdicionalDosDeudor")
                                AdicionalTresDeudor = RsCliente("AdicionalTresDeudor")
                        
                            else
                            
                                AdicionalUnoCliente = "SIN INFORMACION"
                                AdicionalDosCliente = "SIN INFORMACION"
                                AdicionalTresCliente = "SIN INFORMACION"
                                
                                AdicionalUnoDeudor = "SIN INFORMACION"
                                AdicionalDosDeudor = "SIN INFORMACION"
                                AdicionalTresDeudor = "SIN INFORMACION"
                            
                            end if
                        
                        CerrarSCG1()
                        
                       if(AdicionalUnoCliente <> "") then
                    %>
                    <tr>
                        <td class="NombreAtributo"><%=ValidarCampoVacio(AdicionalUnoCliente) %></td>
                        <td class="Atributo InformacionDeudor">
                            <%=ValidarCampoVacio(AdicionalUnoDeudor) %>
                        </td>
                    </tr>
                    <% end if
                    
                       if(AdicionalDosCliente <> "") then
                    %>
                    <tr>
                        <td class="NombreAtributo"><%=ValidarCampoVacio(AdicionalDosCliente) %></td>
                        <td class="AtributoGris InformacionDeudor">
                            <%=ValidarCampoVacio(AdicionalDosDeudor) %>
                        </td>
                    </tr>
                    <% end if
                    
                       if(AdicionalTresCliente <> "") then
                    %>
                    <tr>
                        <td class="NombreAtributo"><%=ValidarCampoVacio(AdicionalTresCliente) %></td>
                        <td class="Atributo InformacionDeudor">
                            <%=ValidarCampoVacio(AdicionalTresDeudor) %>
                        </td>
                    </tr>
                    <% end if %>					
                </table>
            </div>
            
            <div id="BlockPago" class="Block NoVisible">
                <table width="90%" height="100%" cellspacing="0">
                    <tr><td colspan="4" class="Title"><img width="16px" class="TitleIcon" src="../Imagenes/FichaDeudor/icon/pago.png">FORMA DE PAGO</td></tr>
                    <tr>
                        <td class="NombreAtributo">Cuando paga</td>
                        <td class="Atributo"><input value="<% Response.Write(GetValueForFieldFromArrayOfData(ArrDatosCampo, _
                                                                                            CAMPO_CUANDO_PAGA, _
                                                                                            DOMINIO_CAMPO_TEXTO, _
                                                                                            COLUMNA_TEXTO)) %>" id="Campo_<%=CAMPO_CUANDO_PAGA%>_<%=DOMINIO_CAMPO_TEXTO%>" name="Campo_<%=CAMPO_CUANDO_PAGA%>" type="text" maxlength="50" /></td>
                        <td class="FechaModificacion">
                            <label title="Última fecha de modificación">
                            <%
                                fechaModificacionCampo = GetDateTimeFieldFromArrayOfData(CAMPO_CUANDO_PAGA, DOMINIO_CAMPO_TEXTO)
                                
                                Response.Write(fechaModificacionCampo)
                            %>
                            </label>
                        </td>
                        <td class="BotonHistorial">
                            <% if fechaModificacionCampo <> "" then %>
                            <img title="Ver histórico de cambios" src="../imagenes/icon_gestiones.jpg" class="ButtonHistorial" idCampo="<%=CAMPO_CUANDO_PAGA %>"></img>
                            <% end if %>
                        </td>
                    </tr>
                    <tr class="AtributoGris">
                        <td class="NombreAtributo">Dia / hora pago especial</td>
                        <td class="Atributo">
                            <%
                            domainFieldIdToUse = GenerateFieldAndGetDomainFieldIdToUse(CAMPO_DIA_HORA_PAGO_ESPECIAL)
                            %>
                        </td>
                        <td class="FechaModificacion">
                            <label title="Última fecha de modificación">
                            <%
                                fechaModificacionCampo = GetDateTimeFieldFromArrayOfData(CAMPO_DIA_HORA_PAGO_ESPECIAL, domainFieldIdToUse)
                                
                                Response.Write(fechaModificacionCampo)
                            %>
                            </label>
                        </td>
                        <td class="BotonHistorial">
                            <% if fechaModificacionCampo <> "" then %>
                            <img title="Ver histórico de cambios" src="../imagenes/icon_gestiones.jpg" class="ButtonHistorial" idCampo="<%=CAMPO_DIA_HORA_PAGO_ESPECIAL %>"></img>
                            <% end if %>
                        </td>
                    </tr>
                    <tr>
                        <td class="NombreAtributo">Donde paga</td>
                        <td class="Atributo"><%=ValidarCampoVacio(DondePaga)%></td>
                        <td></td>
                        <td class="BotonHistorial"></td>
                    </tr>
                    <tr class="AtributoGris">
                        <td class="NombreAtributo">Como paga</td>
                        <td class="Atributo">
                            <%
                            domainFieldIdToUse = GenerateFieldAndGetDomainFieldIdToUse(CAMPO_COMO_PAGA)
                            %>
                        </td>
                        <td class="FechaModificacion">
                            <label title="Última fecha de modificación">
                            <%
                                fechaModificacionCampo = GetDateTimeFieldFromArrayOfData(CAMPO_COMO_PAGA, domainFieldIdToUse)
                                
                                Response.Write(fechaModificacionCampo)
                            %>
                            </label>
                        </td>
                        <td class="BotonHistorial">
                            <% if fechaModificacionCampo <> "" then %>
                            <img title="Ver histórico de cambios" src="../imagenes/icon_gestiones.jpg" class="ButtonHistorial" idCampo="<%=CAMPO_COMO_PAGA %>"></img>
                            <% end if %>
                        </td>
                    </tr>
                    <tr>
                        <td class="NombreAtributo">Ultima forma pago</td>
                        <td class="Atributo"><%=ValidarCampoVacio(UltimaFormaPago)%></td>
                        <td></td>
                        <td class="BotonHistorial"></td>
                    </tr>
                    <tr class="AtributoGris">
                        <td class="NombreAtributo">Documentos para pago</td>
                        <td class="Atributo"><%=ValidarCampoVacio(DocumentosParaPago)%></td>
                        <td></td>
                        <td class="BotonHistorial"></td>
                    </tr>
                    <tr>
                        <td class="NombreAtributo">Exigencias especiales</td>
                        <td class="Atributo"><input value="<% Response.Write(GetValueForFieldFromArrayOfData(ArrDatosCampo, _
                                                                                            CAMPO_EXIGENCIAS_ESPECIALES, _
                                                                                            DOMINIO_CAMPO_TEXTO, _
                                                                                            COLUMNA_TEXTO)) %>" id="Campo_<%=CAMPO_EXIGENCIAS_ESPECIALES%>_<%=DOMINIO_CAMPO_TEXTO%>" name="Campo_<%=CAMPO_EXIGENCIAS_ESPECIALES%>" type="text" maxlength="50" /></td>
                        <td class="FechaModificacion">
                            <label title="Última fecha de modificación">
                            <%
                                fechaModificacionCampo = GetDateTimeFieldFromArrayOfData(CAMPO_EXIGENCIAS_ESPECIALES, DOMINIO_CAMPO_TEXTO)
                                
                                Response.Write(fechaModificacionCampo)
                            %>
                            </label>
                        </td>
                        <td class="BotonHistorial">
                            <% if fechaModificacionCampo <> "" then %>
                            <img title="Ver histórico de cambios" src="../imagenes/icon_gestiones.jpg" class="ButtonHistorial" idCampo="<%=CAMPO_EXIGENCIAS_ESPECIALES %>"></img>
                            <% end if %>
                        </td>
                    </tr>
                    <tr class="AtributoGris">
                        <td class="NombreAtributo">Portales consulta de pago</td>
                        <td class="Atributo"><input value="<% Response.Write(GetValueForFieldFromArrayOfData(ArrDatosCampo, _
                                                                                            CAMPO_PORTALES_CONSULTA_PAGO, _
                                                                                            DOMINIO_CAMPO_TEXTO, _
                                                                                            COLUMNA_TEXTO)) %>" id="Campo_<%=CAMPO_PORTALES_CONSULTA_PAGO%>_<%=DOMINIO_CAMPO_TEXTO%>" name="Campo_<%=CAMPO_PORTALES_CONSULTA_PAGO%>" type="text" maxlength="50" /></td>
                        <td class="FechaModificacion">
                            <label title="Última fecha de modificación">
                            <%
                                fechaModificacionCampo = GetDateTimeFieldFromArrayOfData(CAMPO_PORTALES_CONSULTA_PAGO, DOMINIO_CAMPO_TEXTO)
                                
                                Response.Write(fechaModificacionCampo)
                            %>
                            </label>
                        </td>
                        <td class="BotonHistorial">
                            <% if fechaModificacionCampo <> "" then %>
                            <img title="Ver histórico de cambios" src="../imagenes/icon_gestiones.jpg" class="ButtonHistorial" idCampo="<%=CAMPO_PORTALES_CONSULTA_PAGO %>"></img>
                            <% end if %>
                        </td>
                    </tr>
                    <tr>
                        <td class="NombreAtributo">Pre envio factura region</td>
                        <td class="Atributo">
                            <%
                            domainFieldIdToUse = GenerateFieldAndGetDomainFieldIdToUse(CAMPO_PRE_ENVIO_FACTURA_REGION)
                            %>
                        </td>
                        <td class="FechaModificacion">
                            <label title="Última fecha de modificación">
                            <%
                                fechaModificacionCampo = GetDateTimeFieldFromArrayOfData(CAMPO_PRE_ENVIO_FACTURA_REGION, domainFieldIdToUse)
                                
                                Response.Write(fechaModificacionCampo)
                            %>
                            </label>
                        </td>
                        <td class="BotonHistorial">
                            <% if fechaModificacionCampo <> "" then %>
                            <img title="Ver histórico de cambios" src="../imagenes/icon_gestiones.jpg" class="ButtonHistorial" idCampo="<%=CAMPO_PRE_ENVIO_FACTURA_REGION %>"></img>
                            <% end if %>
                        </td>
                    </tr>
                    <tr class="AtributoGris">
                        <td class="NombreAtributo">Paga a factoring</td>
                        <td class="Atributo">
                            <%
                            domainFieldIdToUse = GenerateFieldAndGetDomainFieldIdToUse(CAMPO_PAGO_FACTORING)
                            %>
                        </td>
                        <td class="FechaModificacion">
                            <label title="Última fecha de modificación">
                            <%
                                fechaModificacionCampo = GetDateTimeFieldFromArrayOfData(CAMPO_PAGO_FACTORING, domainFieldIdToUse)
                                
                                Response.Write(fechaModificacionCampo)
                            %>
                            </label>
                        </td>
                        <td class="BotonHistorial">
                            <% if fechaModificacionCampo <> "" then %>
                            <img title="Ver histórico de cambios" src="../imagenes/icon_gestiones.jpg" class="ButtonHistorial" idCampo="<%=CAMPO_PAGO_FACTORING %>"></img>
                            <% end if %>
                        </td>
                    </tr>
                    <tr>
                        <td class="NombreAtributo">Paga solo a Cliente</td>
                        <td class="Atributo">
                            <%
                            domainFieldIdToUse = GenerateFieldAndGetDomainFieldIdToUse(CAMPO_PAGO_SOLO_CLIENTE)
                            %>
                        </td>
                        <td class="FechaModificacion">
                            <label title="Última fecha de modificación">
                            <%
                                fechaModificacionCampo = GetDateTimeFieldFromArrayOfData(CAMPO_PAGO_SOLO_CLIENTE, domainFieldIdToUse)
                                
                                Response.Write(fechaModificacionCampo)
                            %>
                            </label>
                        </td>
                        <td class="BotonHistorial">
                            <% if fechaModificacionCampo <> "" then %>
                            <img title="Ver histórico de cambios" src="../imagenes/icon_gestiones.jpg" class="ButtonHistorial" idCampo="<%=CAMPO_PAGO_SOLO_CLIENTE %>"></img></img>
                            <% end if %>
                        </td>
                    </tr>
                </table>
            </div>
            
            <div id="BlockContactabilidad" class="Block Visible"> 
                <table width="90%" height="100%" cellspacing="0" class="tablaUbicabilidad">
                    <tr><td colspan="4" class="Title"><img width="16px" class="TitleIcon" src="../Imagenes/FichaDeudor/icon/contactabilidad.png">CONTACTABILIDAD</td></tr>
                </table>
                   <%

                       function ConcatenarDatoUbicabilidad(i, primerTipoContacto, arrDatosUbicabilidad, primerMedioUbicabilidad, primerDatoUbicabilidad)

                         ConcatenarDatoUbicabilidad = primerDatoUbicabilidad

                           if primerTipoContacto = arrDatosUbicabilidad(0, i + 1) then
                                
                                if primerMedioUbicabilidad = arrDatosUbicabilidad(1, i + 1) then

                                    ConcatenarDatoUbicabilidad = primerDatoUbicabilidad & " / " & arrDatosUbicabilidad(2, i + 1)

                                    i = i + 1

                                end if

                            end if

                       end function


                        if UBound(arrDatosUbicabilidad, 2) > 0 then
                        
                            largoUbicabilidad = UBound(arrDatosUbicabilidad, 2)

                            Response.Write("<div id=""accordion"" class=""accordion"">" )

                            primerTipoContacto = ""

                            indiceClaseCss = 0

                            for i = 0 to largoUbicabilidad

                                if primerTipoContacto <> arrDatosUbicabilidad(0, i) then

                                    if i > 0 then

                                        Response.Write("</table></div>")

                                    end if
                                
                                    Response.Write("<h3> " & arrDatosUbicabilidad(0, i) & "</h3> ")

                                    Response.Write("<div> <table width=""100%"" cellspacing=""0"" class=""tablaInterior"">")

                                end if
                       
                                primerTipoContacto = arrDatosUbicabilidad(0, i)

                                primerMedioUbicabilidad = arrDatosUbicabilidad(1, i)

                                primerDatoUbicabilidad = arrDatosUbicabilidad(2, i)

                                if i + 1 < largoUbicabilidad then
                       
                                   primerDatoUbicabilidad = ConcatenarDatoUbicabilidad(i, primerTipoContacto, arrDatosUbicabilidad, primerMedioUbicabilidad, primerDatoUbicabilidad)

                                else

                                   if  i < largoUbicabilidad then

                                    primerDatoUbicabilidad = ConcatenarDatoUbicabilidad(i, primerTipoContacto, arrDatosUbicabilidad, primerMedioUbicabilidad, primerDatoUbicabilidad)
                            
                                    end if

                                end if

                                if indiceClaseCss mod 2 = 0 then

                                    cssClass = "Atributo"

                                else

                                    cssClass = "AtributoGris"

                                end if

                                Response.Write("<tr><td class=""NombreAtributo"">" &  primerMedioUbicabilidad & "</td><td class=""" & cssClass & " InformacionDeudor"">" & primerDatoUbicabilidad & "</td></tr>")

                                indiceClaseCss = indiceClaseCss + 1

                            Next

                        else

                            Response.Write("<table width=""90%"" cellspacing=""0""><tr><td class=""Atributo"">No hay datos de ubicabilidad.</td></tr></table>")

                        end if

                       Response.Write("</table></div>")

                       Response.Write("</div>")
                   %> 
            </div>
            
            <div id="BlockCliente" class="Block NoVisible">
                <table width="90%" height="100%" cellspacing="0">
                    <tr><td colspan="4" class="Title"><img width="16px" class="TitleIcon" src="../Imagenes/FichaDeudor/icon/cliente.png">CLIENTE</td></tr>
                    <tr>
                        <td>UNO</td>
                        <td>DOS</td>
                        <td>TRES</td>
                        <td>CUATRO</td>
                    </tr>
                    <tr>
                        <td>UNO</td>
                        <td>DOS</td>
                        <td>TRES</td>
                        <td>CUATRO</td>
                    </tr>
                    <tr>
                        <td>UNO</td>
                        <td>DOS</td>
                        <td>TRES</td>
                        <td>CUATRO</td>
                    </tr>
                    <tr>
                        <td>UNO</td>
                        <td>DOS</td>
                        <td>TRES</td>
                        <td>CUATRO</td>
                    </tr>
                </table>	
            </div>
            
            <div id="BlockAllHidden" class="Block NoVisible">
                <table width="90%" height="100%" cellspacing="0" id="TablaDeudor">
                    <tr><td class="Title TitleAllHidden">SELECCIONE UNA DE LAS OPCIONES DISPONIBLES</td></tr>
                </table>
            </div>
            
        </section>
    </article>
</form>
</body>

<script type="text/javascript">
 
    var ID_DOMINIO_CAMPO_TEXTO = 7;

    var ArrayDatosFicha = [];
    
    var ArrayIdFicha = [];
    
    var ArrayDatosIngresados = [];
    
    var idBlockVisibles = "";
    
    var submitForm = false;
    
    function RestoreSections() {
    
        if ($("#IdBlockVisibles").val() != "") {
        
            var arrayBlockVisibles = $("#IdBlockVisibles").val().split(",");
            
            $("div[id^=Block]")
                .each(function(){

                    $(this)
                        .addClass("NoVisible")
                        .removeClass("Visible");
                
                });
                
            $("#Menu")
                .find("button")
                .each(function(){
                
                    $(this)
                        .addClass("Boton")
                        .removeClass("BotonHover");
                        
                });
            
            for (var i = 0; i < arrayBlockVisibles.length; i++) {
            
                $("#Button" + arrayBlockVisibles[i])
                    .addClass("BotonHover")
                    .removeClass("Boton");
                    
                $("div[id=Block" + arrayBlockVisibles[i] + "]")
                    .addClass("Visible")
                    .removeClass("NoVisible");
                    
            }
        
        }
    
    }
    
    function CargaArrayDatos() {
        var ArrayDatos = [];
        var i = 0;
        $("input[name^=Campo_]").each( function() {
            if(($(this).attr('type') == "radio" && $(this).is(":checked")) || $(this).attr('type') != "radio") {
                ArrayDatos[i] = $(this).val();
                i++;
            }	
        });
        return ArrayDatos;
    }
    
    function CargaArrayId() {
        var ArrayId = [];
        var i = 0;
        $("input[name^=Campo_]").each( function() {
            if(($(this).attr('type') == "radio" && $(this).is(":checked")) || $(this).attr('type') != "radio") {
                ArrayId[i] = $(this).attr('id');
                i++;
            }	
        });
        return ArrayId;
    }
    
    $(function(){
        ArrayDatosFicha = CargaArrayDatos();
        
        RestoreSections();
        
        $("form").submit(function( event ) {
        
            if (idBlockVisibles == "" && !submitForm) {
            
                event.preventDefault();
                
            }
            
        });

        $("#accordion").accordion({ collapsible: true });
    });
    
    $( "#ButtonGuardar" ).click(function() {
    
        ArrayDatosIngresados = CargaArrayDatos();
        
        var showObservacion = false;
        
        for(i = 0 ;i < ArrayDatosIngresados.length && !showObservacion; i++) {
        
            showObservacion = ArrayDatosIngresados[i] != ArrayDatosFicha[i];
        
        }
        
        if (showObservacion) {
        
            $( "#DialogObservacion" ).dialog("open");
            
        }
        else {
        
            $( "#DialogNoCambios" ).dialog("open");
            
        }
        
    });
    
    $("#ButtonAceptarNoCambios").click(function(){
    
        $( "#DialogNoCambios" ).dialog("close");
    
    });
    
    $("#ButtonAceptarCambiosGuardados").click(function(){
    
        $("button[id^=Button]")
            .each(function(){

                if ($(this).hasClass("BotonHover") && $.trim($(this).val()) != "") {
                
                    idBlockVisibles += $(this).val() + ",";
                
                }
            
            });
        
        if (idBlockVisibles != "") {
        
            idBlockVisibles = idBlockVisibles.substring(0, idBlockVisibles.length - 1);
        
        }
        
        $("#IdBlockVisibles").val(idBlockVisibles);
        
        $( "#DialogCambiosGuardados" ).dialog("close");
        
        submitForm = true;
        
        $("form").submit();
    
    });
    
    $("#ButtonCancelarGuardarFicha").click(function(){
    
        $( "#DialogObservacion" ).dialog("close");
    
    });
    
    $( "#ButtonGuardarFicha" ).click(function() {
    
        ArrayDatosIngresados = CargaArrayDatos();
        
        ArrayIdFicha = CargaArrayId();
        
        for(i = 0 ;i < ArrayDatosIngresados.length; i++) {
        
            if(ArrayDatosIngresados[i] != ArrayDatosFicha[i]) {
            
                var id = ArrayIdFicha[i].split("_");
                
                var texto = "";
                
                if(id[2] == ID_DOMINIO_CAMPO_TEXTO) {
                
                    texto = $("#Campo_"+id[1]+"_"+id[2]).val();
                    
                }
                
                var criterios = "Observacion=" + encodeURIComponent($("#TxaObservacion").val()) + 
                                "&IdCampo=" + id[1] + "&IdDominio=" + id[2] + "&RutDeudor=" + $("#RutDeudor").val() + 
                                "&CodigoCliente=" + $("#CodigoCliente").val() + "&CodigoUsuario=" + $("#CodigoUsuario").val() + 
                                "&Texto=" + encodeURIComponent(texto);
                
                $.post('FuncionesAjax/FichaDeudor/FichaDeudor.asp', criterios, function(data){});
                
            }
            
        }
        
        ArrayDatosFicha = CargaArrayDatos();
        
        $("#ButtonCancelarGuardarFicha").click();
        
        $( "#DialogCambiosGuardados" ).dialog("open");
        
    });
</script>
</html>