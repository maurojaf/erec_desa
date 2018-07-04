<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"  LCID = 1034%>
<!DOCTYPE html>
<html lang="es">
<HEAD>
    <meta http-equiv="X-UA-Compatible" content="IE=edge,chrome=1">
    <meta charset="utf-8">
    <link href="../css/normalize.css" rel="stylesheet">
	<link href="../css/style.css" rel="stylesheet" type="text/css">
	<link href="../css/style_generales_sistema.css" rel="stylesheet">

	<!--#include file="sesion.asp"-->

	<!--#include file="arch_utils.asp"-->

	<!--#include file="../lib/comunes/rutinas/funciones.inc" -->
	<!--#include file="../lib/comunes/rutinas/TraeCampo.inc" -->
	<!--#include file="../lib/comunes/rutinas/TraeTelefonoDeudor.inc" -->

	<!--#include file="../lib/asp/comunes/general/rutinasBooleano.inc"-->
	<!--#include file="../lib/asp/comunes/general/RutinasVarias.inc" -->
	<!--#include file="../lib/lib.asp"-->

<title>INGRESO DE GESTIONES</title>
<%
	Response.CodePage=65001
	Response.charset ="utf-8"

    fonoActual = Request("fono_actual")
	pagina_origen = Request("pagina_origen")

	strCobranza = ""
	
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
			
	'cerrarscg()
	
	if esPerfilEmpresa and Trim(strUsaCustodio) = "S" then
		strCobranza = "0"
	else
		strCobranza = ""
	end if
	
	strCodCliente = Request("cliente")
	
	
	
	'Response.write "<br>strCodCliente=" & strCodCliente
	
%>

<script src="../Componentes/jquery-1.9.2/js/jquery-1.8.3.js"></script>
<script src="../Componentes/jquery-1.9.2/js/jquery-ui-1.9.2.custom.js"></script>

<script src="../Componentes/jquery.tablesorter/jquery.tablesorter.js"></script>

<script src="../Componentes/PrettyNumber/jquery.prettynumber.js"></script>
<link href="../css/style_multi_select.css" rel="stylesheet"> 
<script src="../Componentes/jquery.multiselect.js"></script>
<link rel="stylesheet" type="text/css" href="../Componentes/jquery-1.9.2/css/start/jquery-ui-1.9.2.custom.css">
<link href="../Componentes/prettyLoader/css/prettyLoader.css" rel="stylesheet">
<script src="../Componentes/prettyLoader/js/jquery.prettyLoader.js"></script> 
<script type="text/javascript" charset="utf-8">

	var paramsCuotaChecked = '';

	function FormateaTodosNumeros() {
        $("#PP_TX_CAPITAL").val(FormatearNumero($("#PP_TX_CAPITAL").val()));
        $("#PP_TX_HONORARIOS").val(FormatearNumero($("#PP_TX_HONORARIOS").val()));
        $("#PP_TX_INTERES").val(FormatearNumero($("#PP_TX_INTERES").val()));
        $("#PP_TX_GASTOSPROTESTOS").val(FormatearNumero($("#PP_TX_GASTOSPROTESTOS").val()));
        $("#PP_TX_GASTOS").val(FormatearNumero($("#PP_TX_GASTOS").val()));
        $("#PP_TX_INDEM_COMP").val(FormatearNumero($("#PP_TX_INDEM_COMP").val()));
        $("#PP_TX_TOTALDEUDA").val(FormatearNumero($("#PP_TX_TOTALDEUDA").val()));
        $("#PP_TX_TOTALDEUDA_DESC").val(FormatearNumero($("#PP_TX_TOTALDEUDA_DESC").val()));
        $("#pie").val(FormatearNumero($("#pie").val()));
        $("#PP_TX_TOTALCONVENIO").val(FormatearNumero($("#PP_TX_TOTALCONVENIO").val()));
        $("#desc_interes").val(FormatearNumero($("#desc_interes").val()));
        $("#desc_gastos").val(FormatearNumero($("#desc_gastos").val()));
        $("#GASTOS_PROTESTOS").val(FormatearNumero($("#GASTOS_PROTESTOS").val()));
        $("#desc_indemComp").val(FormatearNumero($("#desc_indemComp").val()));
        $("#desc_honorarios").val(FormatearNumero($("#desc_honorarios").val()));
        $("#desc_capital").val(FormatearNumero($("#desc_capital").val()));
    }

	function LimpiaNumeros() {
        $("#PP_TX_CAPITAL").val($("#PP_TX_CAPITAL").val().replace(/\./g, ""))
		
        $("#PP_TX_HONORARIOS").val($("#PP_TX_HONORARIOS").val().replace(/\./g, ""))
		
        $("#PP_TX_INTERES").val($("#PP_TX_INTERES").val().replace(/\./g, ""))
		
        $("#PP_TX_GASTOSPROTESTOS").val($("#PP_TX_GASTOSPROTESTOS").val().replace(/\./g, ""))
		
        $("#PP_TX_GASTOS").val($("#PP_TX_GASTOS").val().replace(/\./g, ""))
		
        $("#PP_TX_INDEM_COMP").val($("#PP_TX_INDEM_COMP").val().replace(/\./g, ""))
		
        $("#PP_TX_TOTALDEUDA").val($("#PP_TX_TOTALDEUDA").val().replace(/\./g, ""))
		
        $("#PP_TX_TOTALDEUDA_DESC").val($("#PP_TX_TOTALDEUDA_DESC").val().replace(/\./g, ""))
		
        $("#pie").val($("#pie").val().replace(/\./g, ""))
		
        $("#PP_TX_TOTALCONVENIO").val($("#PP_TX_TOTALCONVENIO").val().replace(/\./g, ""))
		
        $("#desc_interes").val($("#desc_interes").val().replace(/\./g, ""))
		
        $("#desc_gastos").val($("#desc_gastos").val().replace(/\./g, ""))
		
        $("#GASTOS_PROTESTOS").val($("#GASTOS_PROTESTOS").val().replace(/\./g, ""))
		
        $("#desc_indemComp").val($("#desc_indemComp").val().replace(/\./g, ""))
		
        $("#desc_honorarios").val($("#desc_honorarios").val().replace(/\./g, ""))
		
        $("#desc_capital").val($("#desc_capital").val().replace(/\./g, ""))
	}

	function LimpiaNumero(numero) {

        var numero = new String(numero);
        var result = numero.replace(/\./g, "")
        return result;
    }

	function FormatearNumero(numero) {
        /*alert(numero);*/
        var number = new String(parseInt(numero.toString().replace(/\./g, "")));
        var result = '';
        while (number.length > 3) {
            result = '.' + number.substr(number.length - 3) + result;
            number = number.substring(0, number.length - 3);
        }
        result = number + result;
        /*alert(result);*/
        return result;

    }
	
	function func_porc_desc_capital(){
		

        if ($("#PP_TX_CAPITAL").val()=="")
        {
            alert("Indique Capital a calcular")
            $("#porc_desc_capital").val("");
            return false;
        }

        LimpiaNumeros()
        

        if (!/^([0-9])*$/.test($("#porc_desc_capital").val()))
        {
            alert("% Descuento Ingrese Solo Numeros");
            $("#desc_capital").val("");
             FormateaTodosNumeros();
            return;
        }
        
        
        if ($("#porc_desc_capital").val() == '' && $("#desc_capital").val() != '')
        {
            func_descuentos($("#desc_capital").val(),'DESCUENTO');
            return false;
        } 

        if ($("#porc_desc_capital").val() == 0) 
        {
            $("#porc_desc_capital").val(0);
            $("#desc_capital").val(0);
            $("#PP_TX_TOTALCONVENIO").val(parseInt($("#PP_TX_TOTALDEUDA").val()) - parseInt($("#pie").val()) - parseInt($("#desc_gastos").val()) - parseInt($("#desc_capital").val()) - parseInt($("#desc_interes").val()) - parseInt($("#desc_honorarios").val()));
            $("#PP_TX_TOTALDEUDA_DESC").val(parseInt($("#PP_TX_TOTALDEUDA").val()) - parseInt($("#desc_gastos").val()) - parseInt($("#desc_capital").val()) - parseInt($("#desc_interes").val()) - parseInt($("#desc_honorarios").val()));
            FormateaTodosNumeros();
		    MostrarPie();
            return ;
        } 


		var Total = Math.round(parseInt($("#PP_TX_CAPITAL").val()) * parseInt($("#porc_desc_capital").val()) / 100);
        

        if (parseInt(Total) <= 0) 
        {
            $("#porc_desc_capital").val(0);
            $("#desc_capital").val(0);
        }

		if (parseInt($("#PP_TX_CAPITAL").val()) < parseInt(Total))
        {
            alert("Monto Capital Descuento no debe ser mayor a Capital Monto de Deuda")
            func_descuentos($("#desc_capital").val(),'DESCUENTO');
            return false;
        }

        $("#desc_capital").val(Total)
		$("#PP_TX_TOTALCONVENIO").val(parseInt($("#PP_TX_TOTALDEUDA").val()) - parseInt($("#pie").val()) - parseInt($("#desc_gastos").val()) - parseInt($("#desc_capital").val()) - parseInt($("#desc_interes").val()) - parseInt($("#desc_honorarios").val()));
        $("#PP_TX_TOTALDEUDA_DESC").val(parseInt($("#PP_TX_TOTALDEUDA").val()) - parseInt($("#desc_gastos").val()) - parseInt($("#desc_capital").val()) - parseInt($("#desc_interes").val()) - parseInt($("#desc_honorarios").val()));

        FormateaTodosNumeros();

		func_porc_capital_pie();
        MostrarPie();
	}
	
	function func_porc_capital_pie()
    {

   /*
    if (datos.pie.value == '') datos.pie.value = 0;
    if (datos.porc_capital_pie.value == '') datos.porc_capital_pie.value = 0;
   */
    LimpiaNumeros();

    if ($("#porc_capital_pie").val() > 100 || $("#porc_capital_pie").val() < 0)
    {
        alert("% pie no Valido");
		CalculateCapitalPercentageAndRefreshAgreement();
        FormateaTodosNumeros();
        return false;
    }


        if ($("#porc_capital_pie").val() == '0' && $("#pie").val() != '0')
        {
            $("#porc_capital_pie").val('0');
            $("#pie").val('0');
            $("#PP_TX_TOTALCONVENIO").val((parseInt($("#PP_TX_TOTALDEUDA").val()) - parseInt($("#pie").val()) - parseInt($("#desc_gastos").val()) - parseInt($("#desc_capital").val()) - parseInt($("#desc_interes").val()) - parseInt($("#desc_honorarios").val())));
            FormateaTodosNumeros();
            return false;
        } 


        if ($("#porc_capital_pie").val() == '' && $("#pie").val() != '')
        {
			//alert();
            func_descuentos($("#desc_capital").val(),'PIE');
			return false;
        } 

		$("#pie").val((parseInt($("#porc_capital_pie").val()) * parseInt($("#PP_TX_TOTALDEUDA_DESC").val()))/ 100);
	    $("#pie").val(Math.round($("#pie").val()));
		
	    $("#PP_TX_TOTALCONVENIO").val((parseInt($("#PP_TX_TOTALDEUDA").val()) - parseInt($("#pie").val()) - parseInt($("#desc_gastos").val()) - parseInt($("#desc_capital").val()) - parseInt($("#desc_interes").val()) - parseInt($("#desc_honorarios").val())));
		
        FormateaTodosNumeros();
        /*MostrarPie();*/
	}
	
	function CalculateCapitalPercentageAndRefreshAgreement()
	{
		var pie = parseInt(LimpiaNumero($("input[name='pie']").val()));
		
		var totalDeudaConDescuento = parseInt(LimpiaNumero($("input[name='PP_TX_TOTALDEUDA_DESC']").val()));
	
		if (pie <= totalDeudaConDescuento) {
			$("input[name='porc_capital_pie']").val(Math.round(pie / totalDeudaConDescuento * 100, 0));
			
			$("input[name='pie']").val(FormatearNumero(pie));
			
			CalculateTotalForAgreement();
		}
		else {
			alert('El monto del pie no puede ser mayor al total deuda con descuento.');
			
			CalculateCapital();
		}
	}
	
	function CalculateCapital()
	{
		var porcentajePie = parseInt(LimpiaNumero($("input[name='porc_capital_pie']").val()));
		
		var totalDeudaConDescuento = parseInt(LimpiaNumero($("input[name='PP_TX_TOTALDEUDA_DESC']").val()));
		
		$("input[name='pie']").val(FormatearNumero(Math.round(porcentajePie / 100 * totalDeudaConDescuento, 0)));
	}
	
	function CalculateTotalForAgreement()
	{
		var totalDeuda = parseInt(LimpiaNumero($("input[name='PP_TX_TOTALDEUDA']").val()));
		
		var pie = parseInt(LimpiaNumero($("input[name='pie']").val()));
		
		var descuentosGastosJudiciales = parseInt(LimpiaNumero($("input[name='desc_gastos']").val()));
		
		var descuentosCapital = parseInt(LimpiaNumero($("input[name='desc_capital']").val()));
		
		var descuentosIntereses = parseInt(LimpiaNumero($("input[name='desc_interes']").val()));
		
		var descuentosHonorarios = parseInt(LimpiaNumero($("input[name='desc_honorarios']").val()));
		
		var descuentosIndemComp = parseInt(LimpiaNumero($("input[name='desc_indemComp']").val()));
		
		$("input[name='PP_TX_TOTALCONVENIO']").val(FormatearNumero(totalDeuda - pie - descuentosGastosJudiciales - descuentosCapital - descuentosIntereses - descuentosHonorarios - descuentosIndemComp));
	}
	
	<%
	
	strSql="SELECT USA_SUBCLIENTE, USA_INTERESES, USA_HONORARIOS, USA_PROTESTOS, FORMULA_HONORARIOS,FORMULA_INTERESES,PIE_PORC_CAPITAL, HON_PORC_CAPITAL, IC_PORC_CAPITAL, TASA_MAX_CONV, DESCRIPCION, RAZON_SOCIAL,INTERES_MORA, USA_CUSTODIO, COD_TIPODOCUMENTO_HON, MESES_TD_HON FROM CLIENTE WHERE COD_CLIENTE ='" & strCodCliente & "'"
	set rsTasa=Conn.execute(strSql)
	if not rsTasa.eof then
		intPorcPie = ValNulo(rsTasa("PIE_PORC_CAPITAL"),"N")/100
	else
		intPorcPie = 0
	end if
	
	%>
		
	function MostrarPie()
	{
		if (parseInt($("#PP_TX_TOTALDEUDA").val()) > 0 && $("#porc_capital_pie").val() == 0)
		{
			$("#porc_capital_pie").val(<%=intPorcPie*100%>);

			CalculateCapital();
			
			CalculateTotalForAgreement();
		}
	}
	
	function func_descuentos(objeto,origen){

		if ($("#PP_TX_CAPITAL").val()=="")
        {
            alert("Indique Capital a calcular")
            $("#porc_desc_capital").val("");
            return false;
        }

        LimpiaNumero(objeto);
		
		LimpiaNumeros();
		
		objeto = objeto.replace(/\./g, "")
		
		if (!/^([0-9])*$/.test(objeto))
        {
            alert("Ingrese Solo Numeros");
            return;
        }

        var por = 0 ;

        if (origen=="DESCUENTO")
        {
            if (parseInt($("#PP_TX_CAPITAL").val())< parseInt(objeto))
            {
                alert("Monto Capital Descuento no debe ser superior a Capital Monto de Deuda")
                func_porc_desc_capital();
                return false;
            }
             if (parseInt($("#PP_TX_CAPITAL").val()) > 0)
            {

                if ($("#porc_desc_capital").val() != '' && $("#desc_capital").val() == '')
                {
                    func_porc_desc_capital($("#porc_desc_capital").val());
                    return false;
                }
                por =  Math.round((objeto/$("#PP_TX_CAPITAL").val())*100);
                $("#porc_desc_capital").val(por);
            }
        }

        if (origen=="INTERES")
        {
        
        if (parseInt($("#PP_TX_INTERES").val())< parseInt(objeto))
            {
                alert("Monto Interes  Descuento no debe ser superior a Interes Monto de Deuda")
                func_porc_desc_interes();
                return false;
            }
               if (parseInt($("#PP_TX_INTERES").val()) > 0)
            {

              if ($("#porc_desc_interes").val() != '' && $("#desc_interes").val() == '')
                {
                    func_porc_desc_interes($("#porc_desc_interes").val());
                    return false;
                } 

            por =  Math.round((objeto/$("#PP_TX_INTERES").val())*100);
            $("#porc_desc_interes").val(por);
            }
        }

           if (origen=="JUDICIAL")
        {
            if (parseInt($("#PP_TX_GASTOS").val())< parseInt(objeto))
            {
                alert("Monto Gastos Judiciales Descuento no debe ser superior a Gastos Judiciales Monto de Deuda")
                func_porc_desc_gastos();
                return false;
            }
            if (parseInt($("#PP_TX_GASTOS").val()) > 0)
            {

              if ($("#porc_desc_gastos").val() != '' && $("#desc_gastos").val() == '')
                {
                    func_porc_desc_gastos($("#porc_desc_gastos").val());
                    return false;
                } 


                por =  Math.round((objeto/$("#PP_TX_GASTOS").val())*100);
                $("#porc_desc_gastos").val(por);
            }
        }

           if (origen=="PROTESTO")
        {
            if (parseInt($("#PP_TX_GASTOSPROTESTOS").val())< parseInt(objeto))
            {
                alert("Monto Gastos Protestos Descuento no debe ser superior a Gastos Protestos Monto de Deuda")
                func_porc_gastosprotestos();
                return false;
            }

             if ($("#porc_desc_gastosprotestos").val() != '' && $("#GASTOS_PROTESTOS").val() == '')
                {
                    func_porc_gastosprotestos($("#porc_desc_gastosprotestos").val());
                    return false;
                } 

                if (parseInt($("#PP_TX_GASTOSPROTESTOS").val()) > 0)
                {
                    por =  Math.round((objeto/$("#PP_TX_GASTOSPROTESTOS").val())*100);
                    $("#porc_desc_gastosprotestos").val(por);
                }
        }


        if (origen=="INDEMCOMP")
        {
            if (parseInt($("#PP_TX_INDEM_COMP").val())< parseInt(objeto))
            {
                alert("Monto IndemComp Descuento no debe ser superior a IndemComp Monto de Deuda")
                func_porc_indemComp();
                return false;
            }
            
            if (parseInt($("#PP_TX_INDEM_COMP").val()) > 0)
            {
               if ($("#porc_desc_indemComp").val() != '' && $("#desc_indemComp").val() == '')
                {
                    func_porc_indemComp($("#porc_desc_indemComp").val());
                    return false;
                } 

                 por =  Math.round((objeto/$("#PP_TX_INDEM_COMP").val())*100);
                 $("#porc_desc_indemComp").val(por);
            }
        }

          if (origen=="HONORARIOS")
        {
            if (parseInt($("#PP_TX_HONORARIOS").val())< parseInt(objeto))
            {
                alert("Monto Honorarios Descuento no debe ser superior a Honorarios Monto de Deuda")
                func_porc_desc_honorarios();
                return false;
            }

            if (parseInt($("#PP_TX_HONORARIOS").val()) > 0)
            {
                
                if ($("#porc_desc_honorarios").val() != '' && $("#desc_honorarios").val() == '')
                {
                    func_porc_desc_honorarios($("#porc_desc_honorarios").val());
                     MostrarPie();
                    return false;
                } 

                por =  Math.round((parseInt(objeto)/parseInt($("#PP_TX_HONORARIOS").val()))*100);
                $("#porc_desc_honorarios").val(parseInt(por));

                MostrarPie();
            }
        }

        if (origen=="PIE") 
        {

   
            if ($("#porc_capital_pie").val() != '0' && $("#porc_capital_pie").val() !== '')// && datos.pie.value == '')
            {
			
                func_porc_capital_pie($("#porc_capital_pie").val());
                return false;
            }else if ($("#pie").val() == '0')
            {
             $("#porc_capital_pie").val(0);
             $("#pie").val(0);
             $("#PP_TX_TOTALCONVENIO").val(parseInt($("#PP_TX_TOTALDEUDA").val()) - parseInt($("#pie").val()) - parseInt($("#desc_gastos").val()) - parseInt($("#desc_capital").val()) - parseInt($("#desc_interes").val()) - parseInt($("#desc_honorarios").val()));
             FormateaTodosNumeros();
             return false;
            }

        if (parseInt($("#PP_TX_TOTALDEUDA_DESC").val()) < parseInt(objeto))
        {
        
            

            alert("Pie a cancelar no debe ser superior a Total Deuda con Descuento")
            $("#PP_TX_TOTALCONVENIO").val(FormatearNumero($("#PP_TX_TOTALDEUDA").val()));
            func_porc_capital_pie();
            return false;
        }
        
        }
        

		if ($("#PP_TX_CAPITAL").val() == '') $("#PP_TX_CAPITAL").val(0);
		if ($("#desc_interes").val() == '') $("#desc_interes").val(0);
        if ($("#desc_gastos").val() == '') $("#desc_gastos").val(0);
        if ($("#GASTOS_PROTESTOS").val() == '') $("#GASTOS_PROTESTOS").val(0);
        if ($("#desc_indemComp").val() == '') $("#desc_indemComp").val(0);
        if ($("#desc_honorarios").val() == '') $("#desc_honorarios").val(0);
        if ($("#desc_capital").val() == '') $("#desc_capital").val(0);
		$("#PP_TX_TOTALCONVENIO").val(parseInt($("#PP_TX_TOTALDEUDA").val()) - parseInt($("#pie").val()) - parseInt($("#desc_gastos").val()) - parseInt($("#desc_capital").val()) - parseInt($("#desc_interes").val()) - parseInt($("#desc_honorarios").val()));

		$("#PP_TX_TOTALDEUDA_DESC").val(parseInt($("#PP_TX_TOTALDEUDA").val()) - parseInt($("#desc_gastos").val()) - parseInt($("#desc_capital").val()) - parseInt($("#desc_interes").val()) - parseInt($("#desc_honorarios").val()));
		

        var Por = 0
        Por  = ($("#pie").val()*100)/($("#PP_TX_TOTALDEUDA_DESC").val())
        $("#porc_capital_pie").val(Math.round(Por))
        
        FormateaTodosNumeros();

        MostrarPie();
	}

    $(document).ready(function(){


        $.prettyLoader({bind_to_ajax: false});

        if($("#fonoActual").val()!= "" ){
		
            $("#cmbcat").val(2).change();
			
				if($("#pagina_origen").val()== ""){
					$("#cmbcat").prop('disabled', true);
				}
        }

        filtro_historial($("#CB_FILTRO").val());
		
		$('#ventana_procesa').dialog({
			show:"blind", 
			hide:"explode",   		       	 
			width:550,
			height:370 ,
			modal:true,	
			autoOpen:false,	
			buttons: {
				Agendar: function() {
					agendamiento_gestion_sin_contacto()
				},
				Cancelar: function() {
					$(this).dialog( "close" );
				}
			}  	
		});	
		
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
		
		});
		
    });    

function ventana_procesa(){

		$('#ventana_procesa').dialog( "open" );
		$('#cmbcat').change()

}
	
</script>

<style type=text/css>

.HeaderWithoutSort {

	font-weight: bold; color: white; font-size: 12px; border: 1px solid #858585; text-align: center;

}

 body {
	 scrollbar-arrow-color: white;
	 scrollbar-dark-shadow-color: #000080;
	 scrollbar-track-color: #0080C0;
	 scrollbar-face-color: #0080C0;
	 scrollbar-shadow-color: white;
	 scrollbar-highlight-color: white;
	 scrollbar-3d-light-color: a;
	 overflow: auto;
	 overflow-x: hidden;
	 scrollbar-base-color:#ffeaff:

 }


.black_overlay{
	display: none;
	position: absolute;
	top: 0%;
	left: 0%;
	width: 100%;
	height: 100%;
	background-color: black;
	z-index:1001;
	-moz-opacity: 0.8;
	opacity:.80;
	filter: alpha(opacity=80);
}
.white_content {
	display: none;
	position: absolute;
	top: 10%;
	left: 10%;
	width: 80%;
	height: 60%;
	padding: 8px;
	border: 8px solid blue;
	background-color: white;
	z-index:1200;
	overflow: auto;
}


#ventana_envio_correo{
	text-align: center;

}
.opcion_envio_correo{
	font-family: "verdana";
	font-size: 12px;
	color:#2E2E2E;
	width: 150px;
	text-align: center;
	cursor: pointer;
	margin-top:20px;
	float:left;
}

.textarea_email{
	width: 303px;	
	height: 40px;
}
.enviado_a{
	width: 100%;
	text-align: left;
	font-family: "verdana";
	font-size: 12px;
	color:#2E2E2E;		
}

.bt_enviar_correo{
	float:left;
	width: 65px;
	font-size: 12px;
	color:#585858;
	border-right:1px solid #D8D8D8;
	border-bottom: 2px solid #BDBDBD;
	height:75px;
	margin-bottom: 10px;
	margin-right: 10px;
	background-color: #E6E6E6;
	border-radius: 5px; 
	cursor: pointer;
}

.texto_enviar{
	display: block;
	margin-top: 30px;

}


.contenido_correo{
	float:left;
	width: 550px;
}

.tabla_contenido{
	border-collapse: separate;
		border-spacing:  0px;
		width: 100%;
		
}

.titulo_envio{
	text-align: left;
	width: 100px;
	font-family: "Verdana";
	font-size: 11px;
	height: 16px;
}

.contenido_envio{
	text-align: left;
	font-family: "Verdana";
	font-size: 11px;
	height: 16px;
	background-color: #F2F2F2;
	border: 2px solid #FFF;
	width:200px;
}

#con_copia{
	width: 300px;
}

.encabezado_email{
	font-family: "Verdana";
	font-size: 14px;
	font-weight: bold;
	color:#585858;
	text-align: left;
}

.cuerpo_email{
	font-family: "Verdana";
	font-size: 12px;
	color:#585858;
	text-align: left;	
	margin-bottom: 10px;	
	margin-top: 10px;
}

.imagens_email{
	text-align: right;
	margin-bottom: 10px;
}

.importante_email{
	font-size: 12px;
	font-weight: bold;
	color:#000;
	text-align: left;
}

.comunicacion_email{
	font-family: "Verdana";
	color:#585858;
	text-align: left;
}
.firma_email{
	text-align: left;
}

.resumen_email{
	font-family: "Verdana";
	font-size: 12px;
	color:#585858;
	text-align: left;

}

.flecha_ordenamiento{
	float: right;
	width: 10px;
	height: 8px;
	margin-top: 3px;
	cursor: pointer;
}

.mas_registros{
	padding: 5px;
	width: 30%;
	font-size: 12px;
	font-weight: bold;
	font-family: "verdana";	
	cursor: pointer;
	margin: 10px;
}

.td_hover{
	height: 22px;
}
.td_hover:nth-child(even) {
    background: #F0F0F0; 
    height: 22px;
}
.td_hover:nth-child(odd) {
    background: #FFF;
    height: 22px;
}

.oculta
{
	background-image: url("../Imagenes/fondo_botones.jpg");
	background-repeat:repeat-x;
}
</style>

<script type="text/javascript">
function IsValidCamposContacto() {

	var validaContactoTelefono = $('#validaContactoTelefono').val()
	var validaContactoEmail = $('#validaContactoEmail').val()
	var validaContactoDireccion = $('#validaContactoDireccion').val()
	
	var TX_CONTACTO 	= $('#TX_CONTACTO').val()
	var TX_CARGO 		= $('#TX_CARGO').val()
	var TX_DPTO 		= $('#TX_DPTO').val()
	var TX_APELLIDO		= $('#TX_APELLIDO').val()
	
	if(validaContactoTelefono == 'True' && validaContactoEmail == 'True' && validaContactoDireccion == 'True') {
		if(TX_CONTACTO==''){
			alert('DEBE INGRESAR NOMBRE');
			return false
		}
		
		if(TX_APELLIDO==''){
			alert('DEBE INGRESAR APELLIDO');
			return false
		}
		
		if(TX_CARGO=='' && TX_DPTO==''){
			alert('DEBE INGRESAR AL MENOS UN CARGO O UN DEPARTAMENTO');
			return false
		}
	}	
	return true
}

function IsValidTipoContacto() {
	
	var validaTipoContacto = $('#validaTipoContacto').val()
	
	if(validaTipoContacto == 'True') {
		return true
	}
	return false
}

function ValidaHora( ObjIng, strHora )
{
    var er_fh = /^(00|01|02|03|04|05|06|07|08|09|10|11|12|13|14|15|16|17|18|19|20|21|22|23)\:([0-5]0|[0-5][1-9])$/
    if( strHora == "" )
    {
            alert("Introduzca la hora.")
            return false
    }
    if ( !(er_fh.test( strHora )) )
    {
            alert("El dato en el campo hora no es válido.");
            ObjIng.value = '';
            ObjIng.focus();
            return false
    }
    return true
}

function set_CB_ID_CONTACTO_GESTION(MEDIO_ASOCIADO,id){

	var criterios ="alea="+Math.random()+"&accion_ajax=actualiza_td_CB_ID_CONTACTO_GESTION&ID_CONTACTO_GESTION="+id+"&MEDIO_ASOCIADO="+MEDIO_ASOCIADO
	$('#td_CB_ID_CONTACTO_GESTION').load("FuncionesAjax/actualiza_medio_contacto.asp", criterios, function(){})
 
}

function set_CB_ID_CONTACTO_FONO_COBRO(id){
	var criterios ="alea="+Math.random()+"&accion_ajax=actualiza_td_CB_ID_CONTACTO_FONO_COBRO&ID_FONO_COBRO="+id
	$('#td_CB_ID_CONTACTO_FONO_COBRO').load('FuncionesAjax/actualiza_medio_contacto.asp', criterios, function(){})
}


function envia(strTipo) {

	if (strTipo == 'AF') {
		var strFonoAgestionar 	=$('#strFonoAgestionar').val()
		var rut 			 	=$('#rut_').val()
		var strCodCliente   	= $('#strCodCliente').val()
		var pagina_origen 		=$('#pagina_origen').val()
		var marco_nv 			="N"
		var marco_VA_SA 		="N"
		var contador 			=0		
		var sinTipoContacto = 0;
				
		if(IsValidTipoContacto()) {				
			$('input[name="correlativo_deudor"]').each(function(){
				var concat_tipoContacto ="#cbxTipoContacto_"+$(this).val()+" option:selected"
				var strTipoContacto		=$(concat_tipoContacto).val()
				
				if(strTipoContacto == "") {
					sinTipoContacto = sinTipoContacto+1;
				}
			})
		}
			
		if(sinTipoContacto == 0) {

			$('input[name="correlativo_deudor"]').each(function(){

				contador = contador + 1
				var concat_anexo 		="#TX_ANEXO_"+$(this).val()
				var concat_tipoContacto ="#cbxTipoContacto_"+$(this).val()+" option:selected"
				var concat_radiomail	="input[id='radiofon"+$(this).val()+"']:checked"
				var concat_TX_DESDE 	="#TX_DESDE_"+$(this).val()
				var concat_TX_HASTA 	="#TX_HASTA_"+$(this).val()
				var concat_CH_DIAS 		="input[id='CH_DIAS_"+$(this).val()+"']:checked"
				var strDiasAtencion     =""

				$(concat_CH_DIAS).each(function () {
					strDiasAtencion =$(this).val()+","+strDiasAtencion
				})

				strDiasAtencion =strDiasAtencion.substring(0, strDiasAtencion.length-1)
				
				var strAnexo  			=$(concat_anexo).val()
				var strTipoContacto		=$(concat_tipoContacto).val()
				var estado_correlativo 	=$(concat_radiomail).val()
				var CORRELATIVO 		=$(this).val()
				var TX_DESDE 			=$(concat_TX_DESDE).val()
				var TX_HASTA 			=$(concat_TX_HASTA).val()

				if(estado_correlativo==2){
					marco_nv ="S"
				}

				if(estado_correlativo==0){
					marco_VA_SA ="S"
				}

				if(estado_correlativo==1){
					marco_VA_SA ="S"
				}
				
				var criterios ="alea="+Math.random()+"&rut="+rut+"&strCodCliente="+strCodCliente+"&estado_correlativo="+estado_correlativo+"&strCodCliente="+strCodCliente+"&strAnexo="+encodeURIComponent(strAnexo)+"&CORRELATIVO="+CORRELATIVO+"&TX_DESDE="+TX_DESDE+"&TX_HASTA="+TX_HASTA+"&strDiasAtencion="+strDiasAtencion+"&strTipoContacto="+strTipoContacto

				$('#carga_funcion_ajax').load('FuncionesAjax/audita_fon_ajax.asp', criterios, function(data){

				    var criterios ="alea="+Math.random()+"&accion_ajax=refresca_ubicabilidad&rut="+rut+"&strCodCliente="+strCodCliente+"&descripcion_medio=telefono"+"&fono_actual="+$("#fonoActual").val()
					$('#opcion_telefono').load('FuncionesAjax/detalle_gestiones_ajax.asp', criterios , function(){


					    $('#imagen_contacto').click(function(){
							
								$('#muestra_email').val(0)
								$('#muestra_direccion').val(0)

								var muestra_contacto =$('#muestra_contacto').val()

								suma_muestra_contacto = parseInt(muestra_contacto) +1 

								$('#muestra_contacto').val(suma_muestra_contacto)

								muestra_contacto =$('#muestra_contacto').val()

								var muestra =$('#muestra').val()



								if(muestra_contacto==1)
								{
									$('#muestra').val("N")
									carga_funcion_telefono()

								}else{

									if(muestra=="S")
									{
										$('#muestra').val("N")
										carga_funcion_telefono()	
									}

									if(muestra=="N")
									{
										$('#muestra').val("S")
										$('#carga_funcion').html("")
									}	 	
								 }
						})
					})
				})				

				if($('input[name="correlativo_deudor"]').length == contador){
					alert("¡Datos actualizados!")
					carga_funcion_telefono()	
					actualiza_medio()	
				}
			});		
		} else {
			alert("TODOS LOS TELEFONOS DEBEN TENER UN TIPO DE CONTACTO SELECCIONADO.")
		}	
	}

	if (strTipo == 'NF') {

		var rut = $('#rut_').val()
		var strCodCliente   	= $('#strCodCliente').val()
		
		var criterios ="alea="+Math.random()+"&strOrigen=deudor_telefonos&strCodCliente="+strCodCliente+"&rut="+rut

		$('#carga_funcion').load('nuevo_tel.asp', criterios, function(data){

		})
	}

	if (strTipo == 'NV') {

		var rut =$('#rut_').val()
		var strCodCliente   	= $('#strCodCliente').val()
		var criterios ="alea="+Math.random()+"&strOrigen=deudor_telefonos&strCodCliente="+strCodCliente+"&strRUT_DEUDOR="+rut

		$('#carga_funcion').load('deudor_telefonos_nv.asp', criterios, function(data){})
	}

}

function actualiza_medio(){
	var rut 			=$('#rut_').val()
	var TIPO_GESTION 	=$('#TIPO_GESTION').val()
	var MEDIO_ASOCIADO 	=$('#MEDIO_ASOCIADO').val()
    var Forma_Pago 	    =$('#CB_FORMA_PAGO').val()
    var strCodCliente   = $('#strCodCliente').val()

	var criterios ="alea="+Math.random()+"&accion_ajax=actualiza_ID_MEDIO_AGENDAMIENTO&rut="+rut+"&MEDIO_ASOCIADO="+MEDIO_ASOCIADO
	$('#td_ID_MEDIO_AGENDAMIENTO').load('FuncionesAjax/actualiza_medio_contacto.asp',criterios, function(){})

	if (!$("#CB_ID_MEDIO_GESTION").is(":disabled")) {
		var criterios ="alea="+Math.random()+"&accion_ajax=actualiza_ID_MEDIO_GESTION&rut="+rut+"&MEDIO_ASOCIADO="+MEDIO_ASOCIADO
		$('#td_ID_MEDIO_GESTION').load('FuncionesAjax/actualiza_medio_contacto.asp',criterios, function(){})
	}
	
	var criterios ="alea="+Math.random()+"&accion_ajax=actualiza_ID_DIRECCION_COBRO_DEUDOR&strCodCliente="+strCodCliente+"&Forma_Pago="+Forma_Pago+"&rut="+rut
	$('#td_CB_ID_DIRECCION_COBRO_DEUDOR').load('FuncionesAjax/actualiza_medio_contacto.asp',criterios, function(){})

	var criterios ="alea="+Math.random()+"&accion_ajax=actualiza_ID_FONO_COBRO&rut="+rut
	$('#td_CB_ID_FONO_COBRO').load('FuncionesAjax/actualiza_medio_contacto.asp',criterios, function(){})

	set_CB_ID_CONTACTO_GESTION(MEDIO_ASOCIADO,0)
	set_CB_ID_CONTACTO_FONO_COBRO(0)
}

function nuevo_telefono(){
		var rut 			 =$('#rut_').val()
		var COD_AREA 		= $('#COD_AREA').val()
		var numero 			= $('#numero').val()
		var TX_CONTACTO 	= $('#TX_CONTACTO').val()
		var TX_CARGO 		= $('#TX_CARGO').val()
		var TX_DPTO 		= $('#TX_DPTO').val()
		var TX_APELLIDO		= $('#TX_APELLIDO').val()
		var CB_FUENTE 		= $('#CB_FUENTE').val()
		var TX_ANEXO		= $('#TX_ANEXO').val()
		var TX_DESDE 		= $('#TX_DESDE').val()
		var TX_HASTA 		= $('#TX_HASTA').val()
		var dias_atencion 	= "" 	
		var num_min 		= $('#num_min').val()		
		var cbxTipoContacto = $('#cbxTipoContacto').val()
		
		if(!IsValidCamposContacto()) {
			return
		}
		
		if(IsValidTipoContacto()) {
			if(cbxTipoContacto=="")
			{
				alert("DEBE SELECCIONAR UN TIPO DE CONTACTO.")
				return
			}
		}
		
		if(numero==''){
			alert('Debe ingresar un numero');

		}else if (valida_largo_nuevo(numero, num_min)){
		}else{

			$('input[name="CH_DIAS"]:checked').each(function () {

				dias_atencion =$(this).val()+","+dias_atencion
			})

			strDiasAtencion =dias_atencion.substring(0, dias_atencion.length-1)

			var criterios ="alea="+Math.random()+"&strOrigen=deudor_telefonos&COD_AREA="+COD_AREA+"&numero="+numero+"&rut="+rut+"&TX_CONTACTO="+encodeURIComponent(TX_CONTACTO)+"&TX_APELLIDO="+encodeURIComponent(TX_APELLIDO)+"&TX_CARGO="+encodeURIComponent(TX_CARGO)+"&TX_DPTO="+encodeURIComponent(TX_DPTO)+"&CB_FUENTE="+encodeURIComponent(CB_FUENTE)+"&CH_DIAS="+strDiasAtencion+"&TX_HASTA="+TX_HASTA+"&TX_DESDE="+TX_DESDE+"&TX_ANEXO="+encodeURIComponent(TX_ANEXO)+"&strTipoContacto="+cbxTipoContacto

			$('#carga_funcion').load('scg_tel.asp', criterios, function(data){

				actualiza_medio()
				carga_funcion_telefono()

				var criterios ="alea="+Math.random()+"&accion_ajax=refresca_ubicabilidad&rut="+rut+"&descripcion_medio=telefono"+"&fono_actual="+$("#fonoActual").val()
				$('#opcion_telefono').load('FuncionesAjax/detalle_gestiones_ajax.asp', criterios , function(){
				 	$('#imagen_contacto').click(function(){
					 	$('#muestra_email').val(0)
					 	$('#muestra_direccion').val(0)

					 	var muestra_contacto =$('#muestra_contacto').val()

					 	suma_muestra_contacto = parseInt(muestra_contacto) +1 

					 	$('#muestra_contacto').val(suma_muestra_contacto)

					 	muestra_contacto =$('#muestra_contacto').val()

					 	var muestra =$('#muestra').val()



					 	if(muestra_contacto==1)
					 	{
					 		$('#muestra').val("N")
					 		carga_funcion_telefono()

					 	}else{

						 	if(muestra=="S")
						 	{
						 		$('#muestra').val("N")
						 		carga_funcion_telefono()	
						 	}

						 	if(muestra=="N")
						 	{
						 		$('#muestra').val("S")
						 		$('#carga_funcion').html("")
						 	}	 	
						 }
				 	})
				})
			})
		}
}

 function modifica_contacto(rut,ID_TELEFONO){

	var criterios ="alea="+Math.random()+"&strOrigen=deudor_telefonos&strRut="+rut+"&intIdTelefono="+encodeURIComponent(ID_TELEFONO)

	$('#carga_funcion').load('modificar_contacto.asp', criterios, function(data){

	})
 
 }

function modifica_contacto_guarda(intIdTelefono){

	var rut 			= $('#rut_').val()
	var TX_CONTACTO 	= $('#TX_CONTACTO').val()
	var TX_APELLIDO		= $('#TX_APELLIDO').val()
	var TX_CARGO 		= $('#TX_CARGO').val()
	var TX_DPTO 		= $('#TX_DPTO').val()
	
	if(TX_CONTACTO==''){
		alert('DEBE INGRESAR NOMBRE');
		return
	}
	
	if(TX_APELLIDO==''){
		alert('DEBE INGRESAR APELLIDO');
		return
	}
	
	if(TX_CARGO=='' && TX_DPTO==''){
		alert('DEBE INGRESAR AL MENOS UN CARGO O UN DEPARTAMENTO');
		return
	}

	var criterios ="alea="+Math.random()+"&strOrigen=deudor_telefonos&strRut="+rut+"&TX_CARGO="+encodeURIComponent(TX_CARGO)+"&TX_CONTACTO="+encodeURIComponent(TX_CONTACTO)+"&TX_APELLIDO="+encodeURIComponent(TX_APELLIDO)+"&TX_DPTO="+encodeURIComponent(TX_DPTO)+"&accion_ajax=guardar_contacto&intIdTelefono="+intIdTelefono

	$('#carga_funcion').load('FuncionesAjax/modificar_contacto_ajax.asp', criterios, function(data){

		actualiza_medio()
	})
}

function modifica_contacto_elimina(strRut,strOrigen,intIdContacto,intIdTelefono)
{
	var criterios ="alea="+Math.random()+"&strOrigen="+strOrigen+"&strRut="+strRut+"&intIdContacto="+intIdContacto+"&intIdTelefono="+intIdTelefono+"&accion_ajax=eliminar_contacto"

	$('#carga_funcion').load('FuncionesAjax/modificar_contacto_ajax.asp', criterios, function(data){
		actualiza_medio()
	})	
}

function carga_funcion_email()
{
	var strFonoAgestionar 	=$('#strFonoAgestionar').val()
	var rut 			 	=$('#rut_').val()
	var strCodCliente   	=$('#strCodCliente').val()

	var criterios ="alea="+Math.random()+"&strFonoAgestionar="+strFonoAgestionar+"&rut="+rut+"&strCodCliente="+strCodCliente+"&muestra_envio_correo=S"

	$('#carga_funcion').load('deudor_email.asp', criterios, function(data){})
}

function agrega_contacto_mail(strOrigen,strRut,intIdEmail)
{	
	var criterios ="alea="+Math.random()+"&strOrigen="+strOrigen+"&strRut="+strRut+"&intIdEmail="+intIdEmail

	$('#carga_funcion').load('modificar_contacto_email.asp', criterios, function(data){})
}

function modifica_email_guarda(strOrigen, intIdEmail){
	var rut 			= $('#rut_').val()
	var TX_CONTACTO 	= $('#TX_CONTACTO').val()
	var TX_APELLIDO		= $('#TX_APELLIDO').val()
	var TX_CARGO 		= $('#TX_CARGO').val()
	var TX_DPTO 		= $('#TX_DPTO').val()

	if(TX_CONTACTO==''){
		alert('DEBE INGRESAR NOMBRE');
		return
	}
	
	if(TX_APELLIDO==''){
		alert('DEBE INGRESAR APELLIDO');
		return
	}
	
	if(TX_CARGO=='' && TX_DPTO==''){
		alert('DEBE INGRESAR AL MENOS UN CARGO O UN DEPARTAMENTO');
		return
	}

	var criterios ="alea="+Math.random()+"&strOrigen="+strOrigen+"&strRut="+rut+"&TX_CARGO="+encodeURIComponent(TX_CARGO)+"&TX_CONTACTO="+encodeURIComponent(TX_CONTACTO)+"&TX_APELLIDO="+encodeURIComponent(TX_APELLIDO)+"&TX_DPTO="+encodeURIComponent(TX_DPTO)+"&accion_ajax=guardar_mail&intIdEmail="+intIdEmail

	$('#carga_funcion').load('FuncionesAjax/modificar_email_ajax.asp', criterios, function(data){
		actualiza_medio()
	}) 	
}

function modifica_email_elimina(strRut,strOrigen,intIdContacto,intIdEmail){
	var rut 			= $('#rut_').val()
	var criterios ="alea="+Math.random()+"&strOrigen="+strOrigen+"&strRut="+strRut+"&intIdContacto="+encodeURIComponent(intIdContacto)+"&intIdEmail="+encodeURIComponent(intIdEmail)+"&accion_ajax=elimina_mail"

	$('#carga_funcion').load('FuncionesAjax/modificar_email_ajax.asp', criterios, function(data){
 		actualiza_medio()		
	})
}

function envia_email(strTipo) {

	if (strTipo == 'AE') {

		var strFonoAgestionar 	=$('#strFonoAgestionar').val()
		var strCodCliente   	= $('#strCodCliente').val()
		var rut 			 	=$('#rut_').val()
		var pagina_origen 		=$('#pagina_origen').val()
		var marco_nv 			="N"
		var marco_VA_SA 		="N"
		var contador 			=0
		var sinTipoContacto     = 0;
		
		if(IsValidTipoContacto()) {
			$('input[name="correlativo_deudor"]').each(function(){
				var concat_tipoContacto ="#cbxTipoContacto_"+$(this).val()+" option:selected"
				var strTipoContacto		=$(concat_tipoContacto).val()
				
				if(strTipoContacto == "") {
					sinTipoContacto = sinTipoContacto+1;
				}
			})
		}
		
		if(sinTipoContacto == 0) {
			$('input[name="correlativo_deudor"]').each(function(){
			
				var concat_tipoContacto ="#cbxTipoContacto_"+$(this).val()+" option:selected"
				var strTipoContacto		=$(concat_tipoContacto).val()
				var concat_anexo 		="#TX_ANEXO_"+$(this).val()
				var concat_radiomail	="input[id='radiomail"+$(this).val()+"']:checked"				
				var strAnexo  			=$(concat_anexo).val()
				var estado_correlativo 	=$(concat_radiomail).val()
				var CORRELATIVO 		=$(this).val()
				contador = contador + 1

				if(estado_correlativo==2){
					marco_nv ="S"
				}

				if(estado_correlativo==0){
					marco_VA_SA ="S"
				}

				if(estado_correlativo==1){
					marco_VA_SA ="S"
				}	

				var criterios ="alea="+Math.random()+"&strOrigen=deudor_email&rut="+rut+"&strCodCliente="+strCodCliente+"&strFonoAgestionar="+strFonoAgestionar+"&estado_correlativo="+estado_correlativo+"&strAnexo="+encodeURIComponent(strAnexo)+"&CORRELATIVO="+CORRELATIVO+"&accion_ajax=auditar_email"+"&strTipoContacto="+strTipoContacto
				$('#carga_funcion_ajax').load("FuncionesAjax/audita_cor_ajax.asp", criterios, function(data){

				    var criterios ="alea="+Math.random()+"&accion_ajax=refresca_ubicabilidad&rut="+rut+"&strCodCliente="+strCodCliente+"&descripcion_medio=email"+"&fono_actual="+$("#fonoActual").val()
					$('#opcion_email').load('FuncionesAjax/detalle_gestiones_ajax.asp', criterios , function(){	

						 $('#imagen_email').click(function(){

							$('#muestra_contacto').val(0)
							$('#muestra_direccion').val(0)
							var muestra_email =$('#muestra_email').val()
							suma_muestra_email = parseInt(muestra_email) +1
							$('#muestra_email').val(suma_muestra_email)
							muestra_email =$('#muestra_email').val()
							var muestra =$('#muestra').val()

							if(muestra_email==1)
							{
								$('#muestra').val("N")
								carga_funcion_email()

							}else{

								if(muestra=="S")
								{
									$('#muestra').val("N")
									carga_funcion_email()	
								}

								if(muestra=="N")
								{
									$('#muestra').val("S")
									$('#carga_funcion').html("")
								}	 	
							 }
						 })
					})
				})
				if($('input[name="correlativo_deudor"]').length == contador){

					alert("¡Datos actualizados!")
					carga_funcion_email()	
					actualiza_medio()
				}
			});
		} else {
			alert("TODOS LOS EMAIL DEBEN TENER UN TIPO DE CONTACTO SELECCIONADO.")
		}
	}
	
	if (strTipo == 'NE') {

		var rut = $('#rut_').val()
		var strCodCliente   	= $('#strCodCliente').val()

		var criterios ="alea="+Math.random()+"&strOrigen=deudor_email&strCodCliente="+strCodCliente+"&rut="+rut
		$('#carga_funcion').load('nuevo_cor.asp', criterios, function(data){})

	}

	if (strTipo == 'NV') {

		var rut = $('#rut_').val()
		var strCodCliente   	= $('#strCodCliente').val()
		
		var criterios ="alea="+Math.random()+"&strOrigen=deudor_email&strCodCliente="+strCodCliente+"&rut="+rut
		$('#carga_funcion').load('deudor_email_nv.asp', criterios, function(data){

		})		
	}	
}

function ValidarCorreo(strmail)
{
    var Formato = /^(([^<>()[\]\.,;:\s@\"]+(\.[^<>()[\]\.,;:\s@\"]+)*)|(\".+\"))@(([^<>()[\]\.,;:\s@\"]+\.)+[^<>()[\]\.,;:\s@\"]{2,})$/i;
    var Comparacion = Formato.test(strmail);    
    
    if(Comparacion == false){
        alert("El e-mail ingresado es invalido!");
        return false;
    }
    return true;
}


function ingresa_nuevo_mail(strOrigen)
{
	var EMAIL 			=$('#EMAIL').val()
	var TX_ANEXO 		=$('#TX_ANEXO').val()
	var TX_CONTACTO 	= $('#TX_CONTACTO').val()
	var TX_APELLIDO		= $('#TX_APELLIDO').val()
	var TX_CARGO 		= $('#TX_CARGO').val()
	var TX_DPTO 		= $('#TX_DPTO').val()
	var CB_FUENTE 		=$('#CB_FUENTE').val()
	var rut 			=$('#rut_').val()
	var strCodCliente   	= $('#strCodCliente').val()
	var cbxTipoContacto = $('#cbxTipoContacto').val()
	
	if(!IsValidCamposContacto()) {
		return
	}
	
	if(IsValidTipoContacto()) {
		if(cbxTipoContacto=="")
		{
			alert("DEBE SELECCIONAR UN TIPO DE CONTACTO.")
			return
		}
	}
	
	if(EMAIL=='')
	{
		alert('DEBE INGRESAR UN CORREO');
		return
	}

	if( ValidarCorreo(EMAIL) )
		{			
			var criterios ="alea="+Math.random()+"&EMAIL="+encodeURIComponent(EMAIL)+"&TX_ANEXO="+encodeURIComponent(TX_ANEXO)+"&TX_CONTACTO="+encodeURIComponent(TX_CONTACTO)+"&TX_CARGO="+encodeURIComponent(TX_CARGO)+"&TX_APELLIDO="+encodeURIComponent(TX_APELLIDO)+"&TX_DPTO="+encodeURIComponent(TX_DPTO)+"&CB_FUENTE="+encodeURIComponent(CB_FUENTE)+"&rut="+rut+"&strOrigen="+strOrigen+"&strTipoContacto="+cbxTipoContacto

			$('#carga_funcion').load('scg_cor.asp', criterios, function(){
				

				carga_funcion_email()

 				actualiza_medio()

 				var criterios ="alea="+Math.random()+"&accion_ajax=refresca_ubicabilidad&strCodCliente="+strCodCliente+"&rut="+rut+"&descripcion_medio=email"+"&fono_actual="+$("#fonoActual").val()
				$('#opcion_email').load('FuncionesAjax/detalle_gestiones_ajax.asp', criterios , function(){

					 $('#imagen_email').click(function(){

					 	$('#muestra_contacto').val(0)
					 	$('#muestra_direccion').val(0)

					 	var muestra_email =$('#muestra_email').val()

					 	suma_muestra_email = parseInt(muestra_email) +1 

					 	$('#muestra_email').val(suma_muestra_email)

					 	muestra_email =$('#muestra_email').val()

					 	var muestra =$('#muestra').val()


					 	if(muestra_email==1)
					 	{
					 		$('#muestra').val("N")
					 		carga_funcion_email()

					 	}else{

						 	if(muestra=="S")
						 	{
						 		$('#muestra').val("N")
						 		carga_funcion_email()	
						 	}

						 	if(muestra=="N")
						 	{
						 		$('#muestra').val("S")
						 		$('#carga_funcion').html("")
						 	}	 	
						 }
					 })
				})
			})
		}	
}


function carga_funcion_direccion()
{
	var strFonoAgestionar 				=$('#strFonoAgestionar').val()
	var rut 			 				=$('#rut_').val()
	var strCodCliente   				=$('#strCodCliente').val()
	var muestra_carga_funcion_direccion =$('#muestra_carga_funcion_direccion').val()
	
	var criterios ="alea="+Math.random()+"&strFonoAgestionar="+encodeURIComponent(strFonoAgestionar)+"&strCodCliente="+strCodCliente+"&rut="+rut+"&strCOD_CLIENTE="+strCodCliente+"&muestra_carga_funcion_direccion="+muestra_carga_funcion_direccion

	$('#carga_funcion').load('deudor_direcciones.asp', criterios, function(data){})	  
}

function agrega_direccion(strOrigen,strRut,intIdDireccion){

	var criterios ="alea="+Math.random()+"&strOrigen="+strOrigen+"&strRut="+strRut+"&intIdDireccion="+intIdDireccion
	$('#carga_funcion').load('modificar_contacto_dir.asp', criterios, function(data){
		actualiza_medio()
	})	
}

function agrega_contacto_direccion(intIdDireccion,strRut){
	var TX_CONTACTO 	= $('#TX_CONTACTO').val()
	var TX_APELLIDO		= $('#TX_APELLIDO').val()
	var TX_CARGO 		= $('#TX_CARGO').val()
	var TX_DPTO 		= $('#TX_DPTO').val()

	if(TX_CONTACTO==''){
		alert('DEBE INGRESAR NOMBRE');
		return
	}
	
	if(TX_APELLIDO==''){
		alert('DEBE INGRESAR APELLIDO');
		return
	}
	
	if(TX_CARGO=='' && TX_DPTO==''){
		alert('DEBE INGRESAR AL MENOS UN CARGO O UN DEPARTAMENTO');
		return
	}

	var criterios ="alea="+Math.random()+"&strRut="+strRut+"&intIdDireccion="+encodeURIComponent(intIdDireccion)+"&accion_ajax=agrega_direccion&TX_CONTACTO="+encodeURIComponent(TX_CONTACTO)+"&TX_APELLIDO="+encodeURIComponent(TX_APELLIDO)+"&strOrigen=deudor_direcciones&TX_CARGO="+encodeURIComponent(TX_CARGO)+"&TX_DPTO="+encodeURIComponent(TX_DPTO)

	$('#carga_funcion').load('FuncionesAjax/modificar_contacto_dir_ajax.asp', criterios, function(data){
		actualiza_medio()
	})	
}

function elimina_direccion(intIdContacto,intIdDireccion,strRut){

	var criterios ="alea="+Math.random()+"&intIdContacto="+intIdContacto+"&accion_ajax=elimina_direccion&intIdDireccion="+intIdDireccion+"&strRut="+strRut+"&strOrigen=deudor_direcciones"

	$('#carga_funcion').load('FuncionesAjax/modificar_contacto_dir_ajax.asp', criterios, function(data){
		actualiza_medio()
	})	
}

function envia_direccion(strTipo) {

	if (strTipo == 'AD') {		
		var strFonoAgestionar 	=$('#strFonoAgestionar').val()
		var rut 			 	=$('#rut_').val()
		var strCodCliente   	= $('#strCodCliente').val()
		var pagina_origen 		=$('#pagina_origen').val()
		var marco_nv 			="N"
		var marco_VA_SA 		="N"
		var cliente_ 			=$('#cliente_').val()
		var contador 			=0		
		var sinTipoContacto     = 0;
		
		if(IsValidTipoContacto()) {
			$('input[name="correlativo_deudor"]').each(function(){
				var concat_tipoContacto ="#cbxTipoContacto_"+$(this).val()+" option:selected"
				var strTipoContacto		=$(concat_tipoContacto).val()
				
				if(strTipoContacto == "") {
					sinTipoContacto = sinTipoContacto+1;
				}
			})
		}
		
		if(sinTipoContacto == 0) {
		
			$('input[name="correlativo_deudor"]').each(function(){

				var concat_tipoContacto ="#cbxTipoContacto_"+$(this).val()+" option:selected"
				var strTipoContacto		=$(concat_tipoContacto).val()
			
				var concat_anexo 		="#TX_ANEXO_"+$(this).val()
				var concat_radiomail	="input[id='radiodir"+$(this).val()+"']:checked"
				var concat_TX_DESDE 	="#TX_DESDE_"+$(this).val()
				var concat_TX_HASTA 	="#TX_HASTA_"+$(this).val()
				var concat_CH_DIAS 		="input[id='CH_DIAS_"+$(this).val()+"']:checked"
				var strDiasAtencion     =""
				contador = contador + 1 

				$(concat_CH_DIAS).each(function () {
					strDiasAtencion =$(this).val()+","+strDiasAtencion
				})

				strDiasAtencion =strDiasAtencion.substring(0, strDiasAtencion.length-1)

				var strAnexo  			=$(concat_anexo).val()
				var estado_correlativo 	=$(concat_radiomail).val()
				var CORRELATIVO 		=$(this).val()
				var TX_HASTA 			=$(concat_TX_HASTA).val()
				var TX_DESDE 			=$(concat_TX_DESDE).val()
				
				if(estado_correlativo==2){
					marco_nv ="S"
				}

				if(estado_correlativo==0){
					marco_VA_SA ="S"
				}

				if(estado_correlativo==1){
					marco_VA_SA ="S"
				}

				var criterios ="alea="+Math.random()+"&strOrigen=deudor_direcciones&rut="+rut+"&strCodCliente="+strCodCliente+"&strFonoAgestionar="+encodeURIComponent(strFonoAgestionar)+"&estado_correlativo="+encodeURIComponent(estado_correlativo)+"&strAnexo="+encodeURIComponent(strAnexo)+"&CORRELATIVO="+CORRELATIVO+"&accion_ajax=auditar_direccion&strDiasAtencion="+strDiasAtencion+"&TX_DESDE="+TX_DESDE+"&TX_HASTA="+TX_HASTA+"&strTipoContacto="+strTipoContacto

				$('#carga_funcion_ajax').load('FuncionesAjax/audita_dir_ajax.asp', criterios, function(data){

				    var criterios ="alea="+Math.random()+"&accion_ajax=refresca_ubicabilidad&rut="+rut+"&strCodCliente="+strCodCliente+"&descripcion_medio=direccion"+"&fono_actual="+$("#fonoActual").val()
					$('#opcion_direccion').load('FuncionesAjax/detalle_gestiones_ajax.asp', criterios , function(){

						$('#imagen_direccion').click(function(){
							$('#muestra_contacto').val(0)
							$('#muestra_email').val(0)

							var muestra_direccion =$('#muestra_direccion').val()

							suma_muestra_direccion = parseInt(muestra_direccion) +1 

							$('#muestra_direccion').val(suma_muestra_direccion)

							muestra_direccion =$('#muestra_direccion').val()

							var muestra =$('#muestra').val()

								if(muestra_direccion==1)
								{
									$('#muestra').val("N")
									carga_funcion_direccion()

								}else{

									if(muestra=="S")
									{
										$('#muestra').val("N")
										carga_funcion_direccion()	
									}

									if(muestra=="N")
									{
										$('#muestra').val("S")
										$('#carga_funcion').html("")
									}	 	
								 }
						 })
					})		 		
				})

				if($('input[name="correlativo_deudor"]').length == contador){
					alert("¡Datos actualizados!")
					carga_funcion_direccion()	
					actualiza_medio()	

				}
			});
		} else {
			alert("TODOS LAS DIRECCIONES DEBEN TENER UN TIPO DE CONTACTO SELECCIONADO.")
		}				
	}
	if (strTipo == 'ND') {

		var rut =$('#rut_').val()
		var strCodCliente   	= $('#strCodCliente').val()		

		var criterios ="alea="+Math.random()+"&strOrigen=deudor_direcciones&strCodCliente="+strCodCliente+"&rut="+rut

	 	$('#carga_funcion').load('nueva_dir.asp', criterios, function(data){
	 		
	 	})

	}

	if (strTipo == 'NV') {

		var rut =$('#rut_').val()
		var strCodCliente   	= $('#strCodCliente').val()		

		var criterios ="alea="+Math.random()+"&strOrigen=deudor_direcciones&strCodCliente="+strCodCliente+"&rut="+rut

	 	$('#carga_funcion').load('deudor_direcciones_nv.asp', criterios, function(data){})		
	}	
}
 
function guarda_nueva_direccion(){

	var comuna  		=$('#comuna').val()
	var calle  			=$('#calle').val()
	var numero  		=$('#numero').val()
	var resto  			=$('#resto').val()
	var strDiasAtencion =""
	var rut 			=$('#rut_').val()
	var TX_DESDE  		=$('#TX_DESDE').val()
	var TX_HASTA  		=$('#TX_HASTA').val()
	var TX_CONTACTO 	=$('#TX_CONTACTO').val()
	var TX_CARGO 		=$('#TX_CARGO').val()
	var TX_DPTO 		=$('#TX_DPTO').val()
	var TX_APELLIDO		=$('#TX_APELLIDO').val()
	var cliente_ 		=$('#cliente_').val()
	var cbxTipoContacto =$('#cbxTipoContacto').val()

	$("input[id='CH_DIAS']:checked").each(function () {
		strDiasAtencion =$(this).val()+","+strDiasAtencion
	})

	strDiasAtencion 	=strDiasAtencion.substring(0, strDiasAtencion.length-1)


	if (strDiasAtencion=="")
	{
		alert("Debe seleccionar al menos 1 día de pago");
		return
	}

	if (calle=="")
	{
		alert('Debe ingresar una calle');
		return
	}

	if (numero=="")
	{
		alert('Debe ingresar un numero');
		return
	}

	if (comuna=="")
	{
		alert('Debe seleccionar una comuna');
		return
	}
	
	if(!IsValidCamposContacto()) {
		return
	}
	
	if(IsValidTipoContacto()) {
		if(cbxTipoContacto=="")
		{
			alert("DEBE SELECCIONAR UN TIPO DE CONTACTO.")
			return
		}
	}

	var criterios ="alea="+Math.random()+"&strOrigen=deudor_direcciones&rut="+rut+"&comuna="+encodeURIComponent(comuna)+"&numero="+encodeURIComponent(numero)+"&calle="+encodeURIComponent(calle)+"&strDiasAtencion="+strDiasAtencion+"&TX_CONTACTO="+encodeURIComponent(TX_CONTACTO)+"&TX_APELLIDO="+encodeURIComponent(TX_APELLIDO)+"&TX_CARGO="+encodeURIComponent(TX_CARGO)+"&TX_DPTO="+encodeURIComponent(TX_DPTO)+"&TX_HASTA="+encodeURIComponent(TX_HASTA)+"&TX_DESDE="+TX_DESDE+"&resto="+encodeURIComponent(resto)+"&strTipoContacto="+cbxTipoContacto

 	$('#carga_funcion').load('scg_dir.asp', criterios, function(data){

 		carga_funcion_direccion()

		actualiza_medio()

		var criterios ="alea="+Math.random()+"&accion_ajax=refresca_ubicabilidad&rut="+rut+"&descripcion_medio=direccion"+"&fono_actual="+$("#fonoActual").val()
		$('#opcion_direccion').load('FuncionesAjax/detalle_gestiones_ajax.asp', criterios , function(){
			$('#imagen_direccion').click(function(){
			 	$('#muestra_contacto').val(0)
				$('#muestra_email').val(0)

			 	var muestra_direccion =$('#muestra_direccion').val()

			 	suma_muestra_direccion = parseInt(muestra_direccion) +1 

			 	$('#muestra_direccion').val(suma_muestra_direccion)

			 	muestra_direccion =$('#muestra_direccion').val()


			 	var muestra =$('#muestra').val()



				 	if(muestra_direccion==1)
				 	{
				 		$('#muestra').val("N")
				 		carga_funcion_direccion()

				 	}else{

					 	if(muestra=="S")
					 	{
					 		$('#muestra').val("N")
					 		carga_funcion_direccion()	
					 	}

					 	if(muestra=="N")
					 	{
					 		$('#muestra').val("S")
					 		$('#carga_funcion').html("")
					 	}	 	
					 }
			})
		})
 	})
}

function  asigna_minimo_a_variable(COD_AREA, num_min)
{
	$('#num_min').val(asigna_minimo_nuevo(COD_AREA,num_min))

} 

function bt_descargar(ruta)
{
	window.open(ruta, '_blank', 'width=600,height=600,resizable=yes') 

}
$(document).ready(function(){

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

    $("#table_tablesorter").tablesorter({dateFormat: "uk"});

	$(document).tooltip({
			
		open: function(event, ui)
		{
		    var $id = $(ui.tooltip).attr('id');

			$('#' + $id).show();
		},
		close: function(event, ui)
		{
			ui.tooltip.hover(function()
			{
				
				$(this).stop(true).fadeTo(400, 1); 
			},
			function()
			{
				$(this).fadeOut('400', function()
				{
					$(this).hide();
				});
			});
		}
	}).click(function(event){
		var target = $( event.target );
		if ( target.is( "a" ) ) {
			$("div.ui-tooltip").hide();
			$(".estilo_columna_individual").focus();
		}
	});

	$("#con_copia").multiselect(); 
	$('#TX_FEC_COMPROMISO').datepicker( {changeMonth: true,changeYear: true, dateFormat: 'dd/mm/yy'})
	$('#TX_FEC_COMPROMISO_RUTA').datepicker( {changeMonth: true,changeYear: true, dateFormat: 'dd/mm/yy'})
	$('#TX_FEC_GESTION_TERRENO').datepicker( {changeMonth: true,changeYear: true, dateFormat: 'dd/mm/yy'})
	$('#TX_FEC_NORM').datepicker( {changeMonth: true,changeYear: true, dateFormat: 'dd/mm/yy'})
	$('#TX_FEC_NORM2').datepicker( {changeMonth: true,changeYear: true, dateFormat: 'dd/mm/yy'})
	$('#TX_FEC_AGEND').datepicker( {changeMonth: true,changeYear: true, dateFormat: 'dd/mm/yy'})
	$('#TX_FEC_AGEND_SC').datepicker( {changeMonth: true,changeYear: true, dateFormat: 'dd/mm/yy'})
	$('#TX_FEC_AGEND_TEL').datepicker( {changeMonth: true,changeYear: true, dateFormat: 'dd/mm/yy'})
	$('#TX_FEC_AGEND_EMAIL').datepicker( {changeMonth: true,changeYear: true, dateFormat: 'dd/mm/yy'})
	$('#TX_FEC_AGEND_DIRECCION').datepicker( {changeMonth: true,changeYear: true, dateFormat: 'dd/mm/yy'})
    
    $('#FECHA_VENC').datepicker( {changeMonth: true,changeYear: true, dateFormat: 'dd/mm/yy'})

    $('#imagen_contacto').click(function(){

	 	$('#muestra_email').val(0)
	 	$('#muestra_direccion').val(0)

	 	var muestra_contacto =$('#muestra_contacto').val()

	 	suma_muestra_contacto = parseInt(muestra_contacto) +1 

	 	$('#muestra_contacto').val(suma_muestra_contacto)

	 	muestra_contacto =$('#muestra_contacto').val()

	 	var muestra =$('#muestra').val()

	 	if(muestra_contacto==1)
	 	{
	 		$('#muestra').val("N")
	 		carga_funcion_telefono()

	 	}else{

		 	if(muestra=="S")
		 	{
		 		$('#muestra').val("N")
		 		carga_funcion_telefono()	
		 	}

		 	if(muestra=="N")
		 	{
		 		$('#muestra').val("S")
		 		$('#carga_funcion').html("")
		 	}	 	
		 }
    })
   
 
	
	 $('#imagen_email').click(function(){

	 	$('#muestra_contacto').val(0)
	 	$('#muestra_direccion').val(0)

	 	var muestra_email =$('#muestra_email').val()

	 	suma_muestra_email = parseInt(muestra_email) +1 

	 	$('#muestra_email').val(suma_muestra_email)

	 	muestra_email =$('#muestra_email').val()

	 	var muestra =$('#muestra').val()


	 	if(muestra_email==1)
	 	{
	 		$('#muestra').val("N")
	 		carga_funcion_email()

	 	}else{

		 	if(muestra=="S")
		 	{
		 		$('#muestra').val("N")
		 		carga_funcion_email()	
		 	}

		 	if(muestra=="N")
		 	{
		 		$('#muestra').val("S")
		 		$('#carga_funcion').html("")
		 	}	 	
		 }
	 })

 $('#imagen_direccion').click(function(){
 	$('#muestra_contacto').val(0)
	$('#muestra_email').val(0)

 	var muestra_direccion =$('#muestra_direccion').val()

 	suma_muestra_direccion = parseInt(muestra_direccion) +1 

 	$('#muestra_direccion').val(suma_muestra_direccion)

 	muestra_direccion =$('#muestra_direccion').val()

 	var muestra =$('#muestra').val()

	 	if(muestra_direccion==1)
	 	{
	 		$('#muestra').val("N")
	 		carga_funcion_direccion()

	 	}else{

		 	if(muestra=="S")
		 	{
		 		$('#muestra').val("N")
		 		carga_funcion_direccion()	
		 	}

		 	if(muestra=="N")
		 	{
		 		$('#muestra').val("S")
		 		$('#carga_funcion').html("")
		 	}	 	
		 }
 })

 if ("<%= TraeSiNo(session("perfil_sup")) %>" == "No")
 {
 $("#imagen_contacto").click();
 }

		var SiNo = ("<%= TraeSiNo(session("perfil_emp")) %>" == "Si");
		if(SiNo){
			filtro_historial('EFECTIVAS ACTIVAS');
			$("#CB_FILTRO").val("EFECTIVAS ACTIVAS");
		}
		else{
			filtro_historial('EFECTIVAS ACTIVAS');
			$("#CB_FILTRO").val("EFECTIVAS ACTIVAS");
		}

		$('.cambio_flecha_ordenamiento').toggle(function(){
			$('.flecha_ordenamiento').attr('src', '../Imagenes/flecha_arriba_ordenamiento.png')
		}, function(){
			$('.flecha_ordenamiento').attr('src', '../Imagenes/flecha_abajo_ordenamiento.png')
		})
})

function bt_ver_historial(ID_CUOTA)
{

	window.open('historial_documentos_biblioteca_deudor.asp?ID_CUOTA='+ID_CUOTA,"_new","width=900, height=500, scrollbars=yes, menubar=no, location=no, resizable=yes")
}

function ingresa_cuotas(){
	var rut_ 				=$('#rut_').val()
	var strUsaSubCliente  	=$('#strUsaSubCliente').val()

	var criterios 	="alea="+Math.random()+"&accion_ajax=ingresa_cuotas_deudor&RUT_DEUDOR="+rut_+"&strUsaSubCliente="+strUsaSubCliente
	$('#refreca_acccion_archivo').load('FuncionesAjax/ingresa_deudor_deuda_ajax.asp', criterios, function(){
		$('#FECHA_VENC').datepicker( {changeMonth: true,changeYear: true, dateFormat: 'dd/mm/yy'})

	})
}

function bt_ingresa_cuotas_deudor(){
	
	var NRO_DOC 			=$('#NRO_DOC').val()
	var NRO_CUOTA 			=$('#NRO_CUOTA').val()
	var TIPO_DOCUMENTO 		=$('#TIPO_DOCUMENTO').val()
	var NRO_CLIENTE_DEUDOR 	=$('#NRO_CLIENTE_DEUDOR').val()
	var NRO_CLIENTE_DOC 	=$('#NRO_CLIENTE_DOC').val()
	var SUCURSAL 			=$('#SUCURSAL').val()
	var FECHA_VENC 			=$('#FECHA_VENC').val()
	var VALOR_CUOTA 		=$('#VALOR_CUOTA').val()
	var OBSERVACION 		=$('#OBSERVACION').val()
	var ADIC_1 				=$('#ADIC_1').val()
	var ADIC_2 				=$('#ADIC_2').val()
	var ADIC_3 				=$('#ADIC_3').val()
	var GASTOS_PROTESTOS 	=$('#GASTOS_PROTESTOS').val()
	var RUT_DEUDOR 	 		=$('#rut_').val()
	var ID_BANCO 			=$('#ID_BANCO').val()
	var strUsaSubCliente  	=$('#strUsaSubCliente').val()
	var RUT_SUBCLIENTE 		=$('#RUT_SUBCLIENTE').val()
	var NOMBRE_SUBCLIENTE 	=$('#NOMBRE_SUBCLIENTE').val()	

	if(isNaN(NRO_CUOTA)){
		alert("Formato invalido, solo ingrese numeros")
		return
	}

	if(isNaN(VALOR_CUOTA)){
		alert("Formato invalido, solo ingrese numeros")
		return
	}

	if(ID_BANCO==null){
		ID_BANCO =""
	}	

	if(RUT_SUBCLIENTE==null){
		RUT_SUBCLIENTE =""
	}

	if(NOMBRE_SUBCLIENTE==null){
		NOMBRE_SUBCLIENTE =""
	}

	if(confirm("¿Esta seguro de ingresar cuota al deudor?")){
		var criterios	="alea="+Math.random()+"&accion_ajax=inserta_cuotas_deudor&NRO_DOC="+NRO_DOC+" &NRO_CUOTA="+NRO_CUOTA+"&TIPO_DOCUMENTO="+TIPO_DOCUMENTO+"&NRO_CLIENTE_DEUDOR="+NRO_CLIENTE_DEUDOR+"&NRO_CLIENTE_DOC="+NRO_CLIENTE_DOC+"&SUCURSAL="+SUCURSAL+"&FECHA_VENC="+FECHA_VENC+"&VALOR_CUOTA="+VALOR_CUOTA+"&OBSERVACION="+OBSERVACION+"&ADIC_1="+ADIC_1+"&ADIC_2="+ADIC_2+"&ADIC_3="+ADIC_3+"&GASTOS_PROTESTOS="+GASTOS_PROTESTOS+"&RUT_DEUDOR="+RUT_DEUDOR+"&ID_BANCO="+ID_BANCO+"&strUsaSubCliente="+strUsaSubCliente+"&RUT_SUBCLIENTE="+RUT_SUBCLIENTE+"&NOMBRE_SUBCLIENTE="+NOMBRE_SUBCLIENTE

		$('#div_ingresa_cuotas_deudor_tabla').load('FuncionesAjax/ingresa_deudor_deuda_ajax.asp', criterios, function(){
			var ULT_ID_CUOTA = $('#ULT_ID_CUOTA').val()
			if(isNaN(ULT_ID_CUOTA)){

			}else{
				Refrescar()
				bt_limpia_campos()
			}
		})		
	}
}

function bt_muestra_banco(valor){
	if(valor=="2" || valor=="18"){
		$('#tr_muestra_banco').css('display', 'block')
	}else{		
		$('#tr_muestra_banco').css('display', 'none')
	}
	alert(valor)
	if(valor=="2" || valor=="3"  || valor=="3" || valor=="3"){
		$('#tr_muestra_cuotas').css('display', 'block')
	}else{		
		$('#tr_muestra_cuotas').css('display', 'none')
	}	
}

function bt_limpia_campos(){

	$('#NRO_DOC').val("")
	$('#NRO_CUOTA').val("")
	$('#TIPO_DOCUMENTO').val("")
	$('#NRO_CLIENTE_DEUDOR').val("")
	$('#NRO_CLIENTE_DOC').val("")
	$('#SUCURSAL').val("")
	$('#FECHA_VENC').val("")
	$('#VALOR_CUOTA').val("")
	$('#OBSERVACION').val("")
	$('#ADIC_1').val("")
	$('#ADIC_2').val("")
	$('#ADIC_3').val("")
	$('#GASTOS_PROTESTOS').val("")
	$('#rut_').val("")
	$('#ID_BANCO').val("")
	$('#strUsaSubCliente').val("")
	$('#RUT_SUBCLIENTE').val("")
	$('#NOMBRE_SUBCLIENTE').val("")
}


function bt_asocia_cuotas_archivo(){

	var contador =0
	$('input[id="CH_ID_CUOTA"]:checked').each(function(){
		contador = contador + 1
	})

	if(contador==0){
		alert("Debe seleccionar cuota")
		return
	}

	$('input[id="CH_ID_CUOTA"]:checked').each(function(){
		var strRut 				=$('#rut_').val()
		var strCodCliente   	=$('#strCodCliente').val()
		var TX_OBSERVACIONES	=""
		var ID_CUOTA 			=$(this).val()
		contador = contador +1

		var criterios ="alea="+Math.random()+"&accion_ajax=CARGA_ARCHIVOS_CUOTA&strRut="+strRut+"&strCodCliente="+strCodCliente+"&ID_CUOTA="+ID_CUOTA+"&TX_OBSERVACIONES="+encodeURIComponent(TX_OBSERVACIONES)

		$('#refreca_acccion_archivo').load('FuncionesAjax/ingresa_deudor_deuda_ajax.asp', criterios, function(){})

		if($('input[id="CH_ID_CUOTA"]:checked').size()==contador)
		{
			Refrescar()
		}
	})
}

function envia_agendamiento(){
datos.BT_CARTERA.disabled = true;
datos.action='modulo_agendamiento_tactico.asp?';
datos.submit();
}
</script>


<%
	strIDCuotas 		= request.querystring("strIDCuotas")
	pagina_origen 		= request.querystring("pagina_origen")
	strNuevaGestion 	= Request("strNuevaGestion")

	strFonoAgestionar 	= Request("strFonoAgestionar")
	strCategoria 		= Request("strCategoria")
	strRutSubCliente 	= Request("strRutSubCliente")
	strContactoSel 		= Request("strContactoSel")
	strRutDeudor		= request("rut")

	if instr(strRutDeudor,"-") = 0 then		

		Cadena1=strRutDeudor
		Cadena2="-"
		If InStr(Cadena1,Cadena2)<0 then
			strRutDeudor = mid(TRIM(strRutDeudor), 1 ,len(TRIM(strRutDeudor))-1) &"-"& mid(TRIM(strRutDeudor), len(TRIM(strRutDeudor)) , 1)
		Else
			strRutDeudor = replace(strRutDeudor,"-","")

			strRutDeudor = mid(TRIM(strRutDeudor), 1 ,len(TRIM(strRutDeudor))-1) &"-"& mid(TRIM(strRutDeudor), len(TRIM(strRutDeudor)) , 1)
		End if

	end if
	
	fono_con 			= request("fono_con")
	area_con 			= request("area_con")
	strGrabar 			= request("strGrabar")
	strPrioriza 		= request("strPrioriza")
	strChTodos 			= Request("CH_TODOS")
	cuotas_deudor 		= request("cuotas_deudor")
	strUsuario			= session("session_idusuario")
	
	'Response.write "<br>pagina_origen=" & pagina_origen 
	
	session("session_RUT_DEUDOR") = strRutDeudor

	If UCASE(Request("CH_TODOS")) = "ON" Then
		strChTodos="CHECKED"
	End if

	AbrirScg1()

		strSql="SELECT ISNULL(ID_CAMPANA,0) as ID_CAMPANA,ISNULL(DEUDOR.RESP_EMAIL,0) AS PRIORIZACION, REPLACE(REPLACE(OBSERVACIONES_CONF,char(13),' '),char(10),' ') as OBSERVACIONES_CONF, FECHA_CONF, USUARIO_CONF , IsNull(datediff(minute,FECHA_CONF,IsNull(FECHA_UG_TITULAR,'01/01/1900')),0) as DIFMINUTOS FROM DEUDOR WHERE RUT_DEUDOR='" & strRutDeudor & "' AND COD_CLIENTE='" & strCodCliente & "'"

		set rsDeudor = Conn1.execute(strSql)
		if not rsDeudor.eof then
			intIdCampana 	=rsDeudor("ID_CAMPANA")
			If (Trim(strObsConf) = "" or IsNull(strObsConf)) Then
				strObsConf 	= Trim(rsDeudor("OBSERVACIONES_CONF"))
			Else
				strObsConf 	= Trim(Replace(Replace(rsDeudor("OBSERVACIONES_CONF"), chr(13)," "), chr(10)," "))
			End if
			
			intMinDif 		= rsDeudor("DIFMINUTOS")

			strFechaConf 	= rsDeudor("FECHA_CONF")
			strUsuarioConf 	= rsDeudor("USUARIO_CONF")
			strPriorizacion = rsDeudor("PRIORIZACION")

			If Trim(strFechaConf) <> "" and Trim(strUsuarioConf) <> "" then
				strTextoConf = "Fecha : " & strFechaConf & " , Usuario : " & strUsuarioConf & ", Obs : "
				strTextoConf = Trim(Replace(Replace(strTextoConf, chr(13)," "), chr(10)," "))
			End If

		else
			intIdCampana=0

		end if
		rsDeudor.close
		set rsDeudor=nothing

	CerrarScg1()

	AbrirSCG1()
	ssql="EXEC proc_Parametros_Tabla_Cliente '"&TRIM(strRutDeudor)&"','"&TRIM(strCodCliente)&"'"

	set rsCLI=Conn1.execute(ssql)
	if not rsCLI.eof then
		strNomFormHon 		= ValNulo(rsCLI("FORMULA_HONORARIOS"),"C")
		strNomFormInt 		= ValNulo(rsCLI("FORMULA_INTERESES"),"C")

		strUsaSubCliente 	= rsCLI("USA_SUBCLIENTE")
		strUsaInteres 		= rsCLI("USA_INTERESES")
		strUsaHonorarios 	= rsCLI("USA_HONORARIOS")
		strUsaProtestos 	= rsCLI("USA_PROTESTOS")
		intUsaDiscador 	= rsCLI("USA_DISCADOR")
		bitAgendarSinContacto = rsCLI("AGENDAR_SIN_CONTACTO")
		intTipoAgendamiento = rsCLI("TIPO_AGENDAMIENTO")


		nombre_cliente 		=rsCLI("RAZON_SOCIAL")
		intRetiroSabado 	=Cint(rsCLI("RETIRO_SABADO"))
		strMsjRetiroSabado 	= ""
		If Trim(intRetiroSabado) = "1" Then
			strMsjRetiroSabado = "sabados,"
		End if

		strUbicFono 		=rsCLI("UBIC_FONO")
		strUbicEmail 		=rsCLI("UBIC_EMAIL")
		strUbicDireccion 	=rsCLI("UBIC_DIRECCION")
	end if
	rsCLI.close
	set rsCLI=nothing
	CerrarSCG1()
	
	'Response.write "<br>pagina_origen=" & pagina_origen 
	
	if bitAgendarSinContacto=true or pagina_origen="agendamiento_tactico" then
	muestraocultadivdivAgendSC = "display:inline"
	else 
	muestraocultadivdivAgendSC = "display:none"
	End If

	AbrirScg1()

		If Trim(strNuevaGestion) = "" Then

			If Trim(strMasTelefonos) = "S" Then
				 strFonoAsociado = strFonoAgestionar
				 strFonoAgend = strFonoAgestionar
				 intIdContacto = strContactoSel
			Else
				strSql = "SELECT TOP 1 FONO_AGEND, ID_CONTACTO, TELEFONO_ASOCIADO FROM GESTIONES WHERE COD_CLIENTE = '" & strCodCliente & "'"
				strSql = strSql & " AND RUT_DEUDOR = '" & strRutDeudor & "' ORDER BY FECHA_INGRESO DESC, CORRELATIVO DESC"

					strFonoAgend = ""
					intIdContacto = ""
					strFonoAsociado = ""

			End If

		End If


		strSql = "SELECT TOP 1 GESTIONES.HORA_DESDE,GESTIONES.HORA_HASTA,FORMA_PAGO,UPPER(ISNULL(UPPER(RE.NOMBRE+' '+RE.UBICACION), upper(DD.CALLE+' '+DD.NUMERO+' '+DD.RESTO+' '+DD.comuna))) LUGAR_PAGO, ISNULL(DOC_GESTION,'') AS DOC_GESTION FROM GESTIONES "
		
		strSql = strSql & " LEFT JOIN FORMA_RECAUDACION RE ON RE.ID_FORMA_RECAUDACION= GESTIONES.ID_FORMA_RECAUDACION "
		strSql = strSql & " LEFT JOIN DEUDOR_DIRECCION DD ON DD.ID_DIRECCION= GESTIONES.ID_DIRECCION_COBRO_DEUDOR "

		strSql = strSql & " WHERE GESTIONES.COD_CLIENTE = '" & strCodCliente & "'"
		strSql = strSql & " AND GESTIONES.RUT_DEUDOR = '" & strRutDeudor & "'"
		strSql = strSql & " AND CAST(COD_CATEGORIA AS VARCHAR(2)) + '-' + CAST(COD_SUB_CATEGORIA AS VARCHAR(2)) + '-' + CAST(COD_GESTION AS VARCHAR(2)) IN (SELECT CAST(COD_CATEGORIA AS VARCHAR(2)) + '-' + CAST(COD_SUB_CATEGORIA AS VARCHAR(2)) + '-' + CAST(COD_GESTION AS VARCHAR(2)) FROM GESTIONES_TIPO_GESTION "

		strSql = strSql & " WHERE GESTION_MODULOS= 11 AND COD_CLIENTE = '" & strCodCliente & "')"
		strSql = strSql & " ORDER BY GESTIONES.FECHA_INGRESO DESC, GESTIONES.CORRELATIVO DESC"

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

	CerrarScg1()

	AbrirScg1()
		strSql = "SELECT ID_TELEFONO FROM DEUDOR_TELEFONO WHERE RUT_DEUDOR='" & strRutDeudor & "' AND ESTADO <> 2 AND CAST(COD_AREA as VARCHAR(2)) + '-' + TELEFONO = '" & strFonoAsociado & "'"

		set rsFono=Conn1.execute(strSql)
		If not rsFono.eof Then
			intIdTelefonoAgestionar = rsFono("ID_TELEFONO")
		Else
			intIdTelefonoAgestionar = 0
		End If
	CerrarScg1()

	AbrirSCG1()
		strSql = "SELECT PR.ID_PRIORIZACION,PR.OBSERVACION_PRIORIZACION, USUARIO.LOGIN, PR.FECHA_PRIORIZACION,TSP.NOM_TIPO_SOLICITUD"
		strSql= strSql & " FROM PRIORIZACION PR INNER JOIN PRIORIZACIONES_CUOTA PRC ON PR.ID_PRIORIZACION = PRC.ID_PRIORIZACION"
		strSql= strSql & " 					 INNER JOIN CUOTA ON CUOTA.ID_CUOTA = PRC.ID_CUOTA"
		strSql= strSql & " 					 INNER JOIN ESTADO_DEUDA ON CUOTA.ESTADO_DEUDA = ESTADO_DEUDA.CODIGO"
		strSql= strSql & " 					 INNER JOIN USUARIO ON PR.ID_USUARIO_PRIORIZACION = USUARIO.ID_USUARIO"
		strSql= strSql & " 					 INNER JOIN TIPO_SOLICITUD_PRIORIZACION TSP ON TSP.ID_TIPO_SOLICITUD = PR.ID_TIPO_SOLICITUD"

		strSql= strSql & " WHERE PR.RUT_DEUDOR = '" & strRutDeudor & "' AND PRC.ESTADO_PRIORIZACION = 0 AND PR.COD_CLIENTE = '" & strCodCliente & "' AND ESTADO_DEUDA.ACTIVO = 1"

		strSql= strSql & " 					 GROUP BY PR.ID_PRIORIZACION,PR.OBSERVACION_PRIORIZACION,USUARIO.LOGIN, PR.FECHA_PRIORIZACION,TSP.NOM_TIPO_SOLICITUD"

		set RsPrio=Conn1.execute(strSql)

		If not RsPrio.eof then

			intEstadoPrior = 0

			Do While Not RsPrio.Eof

			intIdPriorizacion = RsPrio("ID_PRIORIZACION")

			strTotalDoc = ""

			AbrirSCG2()
						strSql = "SELECT CUOTA.NRO_DOC, (CASE WHEN (DATEDIFF(MINUTE,(ISNULL(FECHA_AGEND_ULT_GES,GETDATE()-200)+ convert(varchar(10),ISNULL(HORA_AGEND_ULT_GES,'00:00'),108)),GETDATE()) >= 0) OR [dbo].[fun_trae_fecha_ult_gestion] (CUOTA.ID_CUOTA) < PR.FECHA_PRIORIZACION THEN 1 ELSE 0 END) AS AGEND_PRIO"
						strSql= strSql & " FROM PRIORIZACION PR INNER JOIN PRIORIZACIONES_CUOTA PRC ON PR.ID_PRIORIZACION = PRC.ID_PRIORIZACION"
						strSql= strSql & " 					 INNER JOIN CUOTA ON CUOTA.ID_CUOTA = PRC.ID_CUOTA"
						strSql= strSql & " 					 INNER JOIN USUARIO ON PR.ID_USUARIO_PRIORIZACION = USUARIO.ID_USUARIO"

						strSql= strSql & " WHERE PRC.ID_PRIORIZACION = '" & intIdPriorizacion & "' AND PRC.ESTADO_PRIORIZACION = 0"

						set RsPrioDoc=Conn2.execute(strSql)

						If not RsPrioDoc.eof then

						intAgendPrio = 0

							Do While Not RsPrioDoc.Eof

								strDoc = RsPrioDoc("NRO_DOC")
								strTotalDoc = strTotalDoc & "-" & strDoc
								intAgendPrio = intAgendPrio + RsPrioDoc("AGEND_PRIO")

								RsPrioDoc.movenext
							Loop
						End If

			CerrarSCG2()

				strObsPrio = RsPrio("OBSERVACION_PRIORIZACION")
				strUsuarioPrio = RsPrio("LOGIN")
				strFechaPrio = RsPrio("FECHA_PRIORIZACION")
				strTipoSol = RsPrio("NOM_TIPO_SOLICITUD")

				If Trim(strTotalDoc) <> "" Then
					strTotalDoc = "Doc: " & Mid(strTotalDoc,2,Len(strTotalDoc))
				End If

				If Trim(strFechaPrio) <> "" and Trim(strUsuarioPrio) <> "" then
					strTextoPrio = "Fecha: " & strFechaPrio & " , Usuario : " & strUsuarioPrio & "<br>" & "Tipo Sol: " & strTipoSol & "<br>" & "Obs : " & strObsPrio & "<br>" & strTotalDoc & "<br>" & "<br>"

					strTextoPrioF = strTextoPrioF & strTextoPrio
					
				Else strTextoPrioF = "Sin priorizaciones Pendientes"
				
				End If

				RsPrio.movenext
			Loop
		End If
		
		If strTextoPrioF="" Then
		
		strTextoPrioF = "Sin priorizaciones Pendientes"
		
		End If
		
	CerrarSCG1()
	
	'CARGA DE DATOS DE VALIDACION DE CONTACTO
	AbrirScg2()
		strValidaContacto = "SELECT ValidaContactoTelefono, ValidaContactoEmail, ValidaContactoDireccion, ValidaTipoContacto FROM AtributosCliente WHERE CodigoCliente = '"& strCodCliente &"'"
		set rsValidaContacto = conn2.execute(strValidaContacto)
		
		If not rsValidaContacto.eof then
			strValidaContactoTelefono = rsValidaContacto("ValidaContactoTelefono")
			strValidaContactoEmail = rsValidaContacto("ValidaContactoEmail")
			strValidaContactoDireccion = rsValidaContacto("ValidaContactoDireccion")
			strValidaTipocontacto = rsValidaContacto("ValidaTipoContacto")
		else
			strValidaContactoTelefono = "False"
			strValidaContactoEmail = "False"
			strValidaContactoDireccion = "False"
			strValidaTipocontacto = "True"
		end if
	CerrarSCG2()

%>
</head>

<body leftmargin="0" rightmargin="0" marginwidth="0" topmargin="0" marginheight="0">

<input type="hidden" id="fonoActual" value="<%=fonoActual%>">
<input type="hidden" id="muestra" value="S">
<input type="hidden" id="muestra_contacto" value="0">
<input type="hidden" id="muestra_direccion" value="0">
<input type="hidden" id="muestra_email" value="0">
<input type="hidden" id="pagina_origen" 	name="pagina_origen" 	value="<%=pagina_origen%>">
<input type="hidden" id="strUsaSubCliente" 	name="strUsaSubCliente" value="<%=strUsaSubCliente%>">

<input type="hidden" id="validaContactoTelefono" value="<%=strValidaContactoTelefono%>">
<input type="hidden" id="validaContactoEmail" value="<%=strValidaContactoEmail%>">
<input type="hidden" id="validaContactoDireccion" value="<%=strValidaContactoDireccion%>">
<input type="hidden" id="validaTipoContacto" value="<%=strValidaTipoContacto %>">

<input type="hidden" name="visualiza_formato_correo" id="visualiza_formato_correo" value="N">
<div id="ventana_envio_correo" title="Envio correo" name="ventana_envio_correo" style="display:none;"></div>
<div id="PlanPagoIngresoGestion" name="PlanPagoIngresoGestion" style="display: none;"></div>
<div id="crea_archivos"></div>
<div id="crea_archivos_excel"></div>
<div id="descarga_archivo"></div>
<div id="refresca_fecha_hora">				
	<input type="hidden" name="fecha_generar_documentos" id="fecha_generar_documentos" value="<%=trim(replace(replace(now(),":","-")," ","_"))%>">	
</div>

<form name="datos" method="post">
<input name="num_min" id="num_min" type="hidden" value="10">
<div class="titulo_informe">MÓDULO INGRESO DE GESTIONES <%=nombre_cliente%>&nbsp;&nbsp;
</div>

 <table width="90%" border="0" align="center">
 <% 

    AbrirSCG1()
	ssql="proc_Parametros_Tabla_Deudor '"&TRIM(strCodCliente)&"','"&TRIM(strRutDeudor)&"'"
	set rsDEU=Conn1.execute(ssql)
	if not rsDEU.eof then
		strNombreDeudor 	= rsDEU("NOMBRE_DEUDOR")
		strNombreRpleg 	= rsDEU("REPLEG_NOMBRE")
		strRutDeudor 		= rsDEU("RUT_DEUDOR")
		strRutRpleg 		= rsDEU("REPLEG_RUT")
		strFechaProrroga 	= rsDEU("FECHA_PRORROGA")
		strTrmoVencimiento 	= rsDEU("TRAMO_VENC")
		strTramoMonto 	= rsDEU("TRAMO_MONTO")
		strNombreFoco = rsDEU("NOMBRE_FOCO")
		strTramoAsignacion 	= rsDEU("TRAMO_ASIG")
		strCampana = rsDEU("NOMBRE_CAMPANA")
		strCampanaCliente = rsDEU("NOMBRE_CAMPANA_CLI")
		strSucursal = rsDEU("SUCURSAL")
		strEjecutivo = rsDEU("USUARIO_ASIG")
		strEtapaCobranza = rsDEU("ETAPA_COBRANZA")
		intIdSegmentoVenc = rsDEU("ID_SEGMENTO_VENC")
		intIdSegmentoMonto = rsDEU("ID_SEGMENTO_MONTO")
		intIdSegmentoAsig = rsDEU("ID_SEGMENTO_ASIG")

	else
		strNombreDeudor = "SIN NOMBRE"
	end if
	rsDEU.close
	set rsDEU=nothing
	CerrarSCG1()

	%>

 <tr>
    <td height="242" valign="top">
	
		<table width="100%" border="0" bordercolor="#FFFFFF">
		  <tr>

			<td width="120" class="estilo_columna_individual">&nbsp;&nbsp;RUT DEUDOR&nbsp;&nbsp;&nbsp;&nbsp;
				<a href="javascript:ventanaBusqueda('Busqueda.asp?strOrigen=1&TX_RUT_DEUDOR=<%=strRutDeudor%>&TX_NOMBRE=<%=strNombreDeudor%>')"><img src="../imagenes/buscar.png" border="0"></a>
			
			<td width="80" class="Estilo10" bgcolor="#C9DEF2">
				<A HREF="principal.asp?TX_RUT=<%=strRutDeudor%>&cliente=<%=strCodCliente%>">
					<acronym title="Llevar a pantalla principal">&nbsp;<%=strRutDeudor%></acronym>
				</A>
			</td>	
				
			<td width="120" class="estilo_columna_individual">&nbsp;&nbsp;NOMBRE DEUDOR</td>
			<td width="300" class="Estilo10" bgcolor="#C9DEF2">&nbsp;<%=strNombreDeudor%></td>

			<td width="120" class="estilo_columna_individual">&nbsp;&nbsp;RUT REP. LEG.&nbsp;&nbsp;&nbsp;&nbsp;</td>
			<td width="80" class="Estilo10" bgcolor="#C9DEF2"><%=strRutRpleg%></td>
			
			<td width="120" class="estilo_columna_individual">&nbsp;&nbsp;NOMBRE REP. LEG.</td>
			<td width="300" class="Estilo10" bgcolor="#C9DEF2">&nbsp;<%=strNombreRpleg%></td>
			
			<td bgcolor="#C9DEF2" align="right" style="padding-right: 10px;">

				<abbr title="Agendamiento">
					<img style="cursor: pointer;" width="25px" name="BT_CARTERA" src="../Imagenes/48px-Crystal_Clear_app_kword.png" onClick="envia_agendamiento();"></img>
				<abbr>
			</td>
			
			<td bgcolor="#C9DEF2" align="right" style="padding-right: 10px;">

				<abbr title="Biblioteca Deudor">
					<img src="../imagenes/Icono_Biblioteca.png" onClick="javascript:ventanaBiblioteca('biblioteca_deudores.asp?strRut=<%=strRutDeudor%>');" value="Biblioteca"></img>
				<abbr>
			
			</td>
			
			<td width="50"  bgcolor="#C9DEF2" align="right" style="padding-right: 10px;">
			
			<%If strTextoPrioF = "Sin priorizaciones Pendientes" then%>

				<abbr title="<%=strTextoPrioF%>">
					<img src="../imagenes/priorizar_normal.png" border="0" onClick="javascript:VentanaPriorizarCaso('priorizar_caso.asp?strCodCliente=<%=strCodCliente%>&strRut=<%=strRutDeudor%>&strOrigen=1')">
				<abbr>
				
			<%else%>
			
				<abbr title="<%=strTextoPrioF%>">
					<img src="../imagenes/priorizar_urgente.png" border="0" onClick="javascript:VentanaPriorizarCaso('priorizar_caso.asp?strCodCliente=<%=strCodCliente%>&strRut=<%=strRutDeudor%>&strOrigen=1')">
				<abbr>				

			<%end if%>
			
			</td>	

				<script language="JavaScript" type="text/JavaScript">
					function VentanaPriorizarCaso(URL){
						window.open(URL,"DATOS","width=2000, height=1000, scrollbars=yes, menubar=no, location=no, resizable=yes")
					}
				</script>
				
			<td width="50" bgcolor="#C9DEF2" align="right" style="padding-right: 10px;">
				<abbr title="Ficha Deudor">
					<img style="cursor: pointer;" width="25px" src="../Imagenes/FichaDeudor/icon/botonFichaDeudor.png" onClick="javascript:VentanaFichaDeudor('FichaDeudor.asp?CodigoCliente=<%=strCodCliente%>&RutDeudor=<%=strRutDeudor%>&CodigoUsuario=<%=strUsuario%>')"></img>
				<abbr>
			</td>
				
				<script language="JavaScript" type="text/JavaScript">
					function VentanaFichaDeudor (URL){
						window.open(URL,"DATOS","width=818, height=500, scrollbars=yes, menubar=no, location=no, resizable=no")
					}
				</script>
			
			<td width="50" bgcolor="#C9DEF2" align="CENTER" id="opcion_telefono">

			<% If strUbicFono = "CONTACTADO" then %>


				<img src="../imagenes/mod_telefono_va.png" id="imagen_contacto" onclick="" style="cursor:pointer;" border="0">
	  
			<% ElseIf strUbicFono = "NO CONTACTADO" then %>


				 <img src="../imagenes/mod_telefono_sa.png" id="imagen_contacto" onclick="" style="cursor:pointer;" border="0">

			<% Else %>


				 <img src="../imagenes/mod_telefono_nv.png" id="imagen_contacto" onclick="" style="cursor:pointer;" border="0" >

			<% End If %>

			</td>

		   <td width="50" bgcolor="#C9DEF2" align="CENTER" id="opcion_email">

			<% If strUbicEmail = "CONTACTADO" then %>

				 <img src="../imagenes/mod_mail_va.png" border="0" id="imagen_email" onclick="" style="cursor:pointer;" >

			<% ElseIf strUbicEmail = "NO CONTACTADO" then %>

				 <img src="../imagenes/mod_mail_sa.png" border="0" id="imagen_email" onclick="" style="cursor:pointer;">

			<% Else %>

				 <img src="../imagenes/mod_mail_nv.png" border="0" id="imagen_email" onclick="" style="cursor:pointer;" >

			<% End If %>

			</td>

		   <td width="50" bgcolor="#C9DEF2" align="CENTER" id="opcion_direccion">

			<% If strUbicDireccion = "CONTACTADO" then %>

				 <img src="../imagenes/mod_direccion_va.png" border="0"  style="cursor:pointer;" onclick="" id="imagen_direccion">

			<% ElseIf strUbicDireccion = "NO CONTACTADO" then %>

				 <img src="../imagenes/mod_direccion_sa.png" border="0"  style="cursor:pointer;" onclick="" id="imagen_direccion">

			<% Else %>

				 <img src="../imagenes/mod_direccion_nv.png" border="0"  onclick="" style="cursor:pointer;" id="imagen_direccion">

			<% End If %>

			</td>
			
		  </tr>
		</table>
		
		&nbsp;
		  
		<table  class="estilo_columnas" style="width:100%;" border="1" cellSpacing="0" cellPadding="0">
		<thead>
		<tr class="Estilo34">
			<td width="80" align="center" >ETAPA COBRANZA</td>
			<td width="80" align="center" >FOCO</td>
			<td width="80" align="center" >TRAMO VENCIMIENTO</td>
			<td width="80" align="center" >TRAMO ASIGNACIÓN</td>
			<td width="80" align="center" >TRAMO MONTO</td>
			<td width="80" align="center" >CAMPAÑA</td>
			<td width="80" align="center" >C. CLIENTE</td>
			<td width="80" align="center" >SUCURSAL</td>
			<td width="80" align="center" >EJECUTIVO</td>
		</tr>
		</thead>
		
		<tbody>
		  
		  <tr class="Estilo10" bgcolor="#C9DEF2" border="0">
			<td align="center" height="20"><%=strEtapaCobranza%></td>
		    <td align="center" height="20"><%=strNombreFoco%></td>
			<td align="center" height="20"><%=strTrmoVencimiento%></td>
			<td align="center" ><%=strTramoAsignacion%></td>
			<td align="center" ><%=strTramoMonto%></td>
			<td align="center" ><%=strCampana%></td>
			<td align="center" ><%=strCampanaCliente%></td>
			<td align="center" ><%=strSucursal%></td>
			<td align="center" ><%=strEjecutivo%></td>
		  </tr>
		  
		</tbody>
		  
	  </table>
  
  
	<div id="carga_funcion_ajax" class="class_carga_funcion"></div>
	<div id="carga_funcion" class="class_carga_funcion"></div>	
			
			<div id="ingreso_gestion"></div>
			<div id="agendamiento_gestion_sin_contacto" style=display:none></div>

	<table width="100%">
		<tr>
			<td height="20" ALIGN=LEFT class="subtitulo_informe">
				> DETALLE DE DOCUMENTOS
			</td>
		</tr>
	</table>

	<table  border="1" class="intercalado" style="width:100%;" bordercolor="#000000" cellSpacing="0" cellPadding="1">
        		
	<tr class="Estilo34">
		<td colspan="5" align="LEFT">
		<a href="#" onClick= "marcar_boxes()">Marcar todos</a>&nbsp;&nbsp;&nbsp;
		<a href="#" onClick="desmarcar_boxes()">Desmarcar todos</a>
		</td>
		<td colspan="12" align="RIGHT">Mostrar Todos 
		<%
			if trim(strIDCuotas)="" then
				mostrar_todos =0
			else
				mostrar_todos =1
			end if		

			if trim(pagina_origen)="carga_masiva_archivos" then
				mostrar_todos =1
			else
				mostrar_todos =0
			end if		

		%>	
			<INPUT TYPE="checkbox" NAME="CH_TODOS_CUOTA" ID="CH_TODOS_CUOTA" <%if trim(mostrar_todos)=1 then response.write " checked " end if%> onClick="Refrescar();">
			<input type="hidden" 	id="cuotas_deudor" name="cuotas_deudor" value="">
			</td>
	</tr>
	</table>	
	<input type="hidden" ID="strUsaSubCliente" VALUE='<%=strUsaSubCliente%>'>
	<input type="hidden" ID="strUsaInteres" VALUE='<%=strUsaInteres%>'>
	<input type="hidden" ID="strUsaProtestos" VALUE='<%=strUsaProtestos%>'>
	<input type="hidden" ID="strUsaHonorarios" VALUE='<%=strUsaHonorarios%>'>
	
	<%
	
	AbrirSCG()
	
	strSqlCliente = "SELECT ISNULL(USA_CUSTODIO, 'N') AS USA_CUSTODIO FROM CLIENTE WHERE COD_CLIENTE = '" & strCodCliente & "'"
	
	set rsTemp= Conn.execute(strSqlCliente)
	
	strUsaCustodio = rsTemp("USA_CUSTODIO")
	
	%>	
	
	<div id="div_mostrar_todo">	

	  			<table  border="1" id="table_tablesorter"  class="tablesorter"  style="width:100%;" bordercolor="#000000" cellSpacing="0" cellPadding="1">
				<thead>
  				<tr class="Estilo34">
  					<td>&nbsp;</td>

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
  					<th align="center" width="70">INTERES</th>
  					<%End If%>
  					<%If Trim(strUsaProtestos)="1" Then%>
  					<th align="center" width="80">PROTESTOS</th>
  					<%End If%>
  					<%If Trim(strUsaHonorarios)="1" Then%>
  					<th align="center" width="90">HONORARIOS</th>
  					<%End If%>
  					<th align="center" width="70">ABONO</th>
  					<th align="center" width="70">SALDO</th>
  					<th >FECHA AGEND.</th>
					<% If Trim(strUsaCustodio) = "S" Then %>
					<td class="HeaderWithoutSort">CUSTODIO</td>
					<% End If %>
					<td>&nbsp;</td>
  					<td>&nbsp;</td>
  					<td>&nbsp;</td>
  					<td>&nbsp;</td>
  					<td>&nbsp;</td>
  				</tr>
  				</thead>
  				<tbody>
  				<%
				
				if trim(strIDCuotas)="" then
					mostrar_todos =0
				else
					mostrar_todos =1
				end if
				
				strSql ="exec proc_Trae_Cuotas_Deudor '"&trim(strCodCliente)&"','"&trim(strRutDeudor)&"','"&strIDCuotas&"','','"&trim(strNomFormInt)&"', '"&trim(strNomFormHon)&"', '1', '"&mostrar_todos&"', '" & strCobranza & "' "
				
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

  				Do until rsTemp.eof

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

  						intAntiguedad = ValNulo(rsTemp("ANTIGUEDAD"),"N")

  						intIntereses = rsTemp("INTERESES")
						intHonorarios = rsTemp("HONORARIOS")

  						intProtestos = ValNulo(rsTemp("GASTOS_PROTESTOS"),"N")

  						intTotDoc= intSaldo+intIntereses+intProtestos+intHonorarios

  						intTotSelSaldo = intTotSelSaldo+intSaldo
  						intTotSelAbono = intTotSelAbono+intAbono
  						intTotSelValorCuota = intTotSelValorCuota+intValorCuota

  						intTotSelIntereses= intTotSelIntereses+intIntereses
  						intTotSelProtestos= intTotSelProtestos+intProtestos
  						intTotSelHonorarios= intTotSelHonorarios+intHonorarios
  						intTotSelDoc = intTotSelDoc+intTotDoc

  						strArrConcepto = strArrConcepto & ";" & "CH_" & rsTemp("ID_CUOTA")
						strArrID_CUOTA = strArrID_CUOTA & ";" & rsTemp("ID_CUOTA")

						%>
  						<tr class="Estilo34">

  						<input name="TX_SALDO_<%=Replace(rsTemp("ID_CUOTA"),"-","_")%>" id="TX_SALDO_<%=Replace(rsTemp("ID_CUOTA"),"-","_")%>" type="hidden" value="<%=intTotDoc%>">
  						<input name="TX_CAPITAL_<%=Replace(rsTemp("ID_CUOTA"),"-","_")%>" id="TX_CAPITAL_<%=Replace(rsTemp("ID_CUOTA"),"-","_")%>" type="hidden" value="<%=intValorCuota%>">
  						<input name="TX_HONORARIOS_<%=Replace(rsTemp("ID_CUOTA"),"-","_")%>" id="TX_HONORARIOS_<%=Replace(rsTemp("ID_CUOTA"),"-","_")%>" type="hidden" value="<%=intHonorarios%>">
  						<input name="TX_INTERESES_<%=Replace(rsTemp("ID_CUOTA"),"-","_")%>" id="TX_INTERESES_<%=Replace(rsTemp("ID_CUOTA"),"-","_")%>" type="hidden" value="<%=intIntereses%>">
  						<input name="TX_PROTESTOS_<%=Replace(rsTemp("ID_CUOTA"),"-","_")%>" id="TX_PROTESTOS_<%=Replace(rsTemp("ID_CUOTA"),"-","_")%>" type="hidden" value="<%=intProtestos%>">


  						<TD width="12">
							<INPUT TYPE="checkbox" <% if esPerfilEmpresa and Trim(strCustodio) <> "LLACRUZ" or not esPerfilEmpresa or Trim(strUsaCustodio) = "N" then %> checked="checked" <% else %> disabled="disabled" <% end if %> NAME="CH_ID_CUOTA" id="CH_ID_CUOTA" value="<%=rsTemp("ID_CUOTA")%>">
  						</TD>

          		  		<%If Trim(strUsaSubCliente)="1" Then%>

          		  		<td width="69"><%=rsTemp("RUT_SUBCLIENTE")%></td>

          		  		<td><a href="javascript:ventanaBusqueda('Busqueda.asp?strOrigen=1&TX_RUT_DEUDOR=<%=rsTemp("RUT_DEUDOR")%>&TX_NOMBRE=<%=strNombreDeudor%>&TX_RUTSUBCLIENTE=<%=rsTemp("RUT_SUBCLIENTE")%>&TX_NOMBRE_SUBCLIENTE=<%=rsTemp("NOMBRE_SUBCLIENTE")%>')"><img src="../imagenes/buscar.png" border="0"></a>

          		 		<td title="<%=rsTemp("NOMBRE_SUBCLIENTE")%>">
          		  			<%=Mid(rsTemp("NOMBRE_SUBCLIENTE"),1,35)%>
          		  		<%End If%>

          		  		<td><%=strNroDoc%></td>
          		  		<td><%=strNroCuota%></td>

          		  		<%If intProrroga = "0" Then%>
  							<td>
  								<%=strFechaVenc%>
  							</td>
  						<%Else%>
  							<td bgcolor="#ff6666" title="Vencimiento Original: <%=strFechaVencOriginal%>">
          		  			<%=strFechaVenc%>
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
								<img src="../imagenes/<%=strImagenGest%>" border="0">
								</A>
						</td>

  						<td ALIGN="CENTER">

							<a href="javascript:ventanaGestionesPorDoc('gestiones_por_documento.asp?intID_CUOTA=<%=trim(rsTemp("ID_CUOTA"))%>&strCodCliente=<%=strCodCliente%>&strNroDoc=<%=trim(rsTemp("NRO_DOC"))%>strNroCuota=<%=rsTemp("NRO_CUOTA")%>')">
							<img src="../imagenes/icon_gestiones.jpg" border="0">
						</a>
						</td>

  						<td>
  							<a href="javascript:ventanaMas('<%=strDetCuota%>?ID_CUOTA=<%=trim(rsTemp("ID_CUOTA"))%>&cliente=<%=strCodCliente%>&strRUT_DEUDOR=<%=trim(rsTemp("RUT_DEUDOR"))%>&strNroDoc=<%=trim(rsTemp("NRO_DOC"))%>&strNroCuota=<%=rsTemp("NRO_CUOTA")%>')">
  							<img src="../imagenes/Carpeta3.png" border="0"></a>

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

%>
							<%
							strImagenGest1=""

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

							%>

								<% If strImagenGest1 <> "" Then %>
								<img src="../Imagenes/<%=strImagenGest1%>" border="0">
								<% Else %>
								&nbsp;

								<% End If %>
						</td>

  						</tr>

  						<%

  					rsTemp.movenext
  				intCorrelativo = intCorrelativo + 1
  				loop

				vArrConcepto = split(strArrConcepto,";")
				vArrID_CUOTA = split(strArrID_CUOTA,";")

				intTamvConcepto = ubound(vArrConcepto)

  				rsTemp.close
  				set rsTemp=nothing
	

  				strArrID_CUOTA = Mid(strArrID_CUOTA,2,len(strArrID_CUOTA))
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
  				<td>&nbsp;</td>
  				<td>&nbsp;</td>
				<td Colspan="3" Rowspan="2" align="center"><input type="button" id="ButtonPlanPago" name="ButtonPlanPago" value="Plan de Pago" class="fondo_boton_100" /></td>
  			</tr>


  			<tr class="Estilo34" height="25">

				<td <%=strColspan%>>&nbsp;&nbsp;&nbsp;&nbsp;Totales Seleccionados:</td>
				<td ALIGN="RIGHT"><span id="span_TX_CAPITAL" style="font-weight:bold;">0</span>
					<INPUT TYPE="hidden" NAME="TX_CAPITAL" ID="TX_CAPITAL" DISABLED style="text-align:right;width:90" size="10" size="10" onkeyup="format(this)" onchange="format(this)">
				</td>



				<% If Trim(strUsaInteres)="1" Then%>
					<td ALIGN="RIGHT"><INPUT TYPE="TEXT" NAME="TX_INTERESES"  ID="TX_INTERESES" DISABLED style="text-align:right;width:90" size="10" size="10" onkeyup="format(this)" onchange="format(this)"></td>
				<% Else%>
					<INPUT TYPE="hidden" NAME="TX_INTERESES" ID="TX_INTERESES">
				<% End If%>

				<% If Trim(strUsaProtestos)="1" Then%>
					<td ALIGN="RIGHT"><INPUT TYPE="TEXT" NAME="TX_PROTESTOS" ID="TX_PROTESTOS" DISABLED style="text-align:right;width:90" size="10" size="10" onkeyup="format(this)" onchange="format(this)"></td>
				<% Else%>
					<INPUT TYPE="hidden" NAME="TX_PROTESTOS" ID="TX_PROTESTOS">
				<% End If%>

				<% If Trim(strUsaHonorarios)="1" Then%>
					<td ALIGN="RIGHT"><INPUT TYPE="TEXT" NAME="TX_HONORARIOS" ID="TX_HONORARIOS" DISABLED style="text-align:right;width:90" size="10" onkeyup="format(this)" onchange="format(this)"></td>
				<% Else%>
					<INPUT TYPE="hidden" NAME="TX_HONORARIOS" ID="TX_HONORARIOS">
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
	  		</thead>


	  		<INPUT TYPE="hidden" NAME="strArrID_CUOTA" VALUE="<%=strArrID_CUOTA%>">
	  			
			</table>

		</div>


		<%if trim(pagina_origen)="carga_masiva_archivos" then%>
			<br>
			<div style="width:100%; text-align:right;" >
				<input type="button" class="fondo_boton_100" name="" id="" value="Asocia archivo" onclick="bt_asocia_cuotas_archivo()">
				&nbsp;
				<input type="button" class="fondo_boton_100" name="" id="" value="Ingresa cuotas" onclick="ingresa_cuotas()">
				&nbsp;
				<input type="button" class="fondo_boton_100" name="" id="" onclick="location.href='carga_masiva_archivos.asp'" value="Volver">
			</div>

			<br>
			<table  border="0" style="width:100%;" bordercolor="#000000" cellSpacing="0" cellPadding="1">
				<% 	
					ID_ARCHIVO_VER =Request.querystring("ID_ARCHIVO_VER")

					sql_ver_archivo = " SELECT NOMBRE_ARCHIVO, "
					sql_ver_archivo = sql_ver_archivo & " SUBSTRING(NOMBRE_ARCHIVO, CHARINDEX('.',NOMBRE_ARCHIVO)+1, LEN(NOMBRE_ARCHIVO)) EXTENCION "
					sql_ver_archivo = sql_ver_archivo & " FROM CARGA_ARCHIVOS CAR "
					sql_ver_archivo = sql_ver_archivo & " WHERE ID_ARCHIVO = " & ID_ARCHIVO_VER
					set rs_ver_archivo = conn.execute(sql_ver_archivo)
					if not rs_ver_archivo.eof then
						nombre_archivo 	=rs_ver_archivo("NOMBRE_ARCHIVO")
						extencion 		=rs_ver_archivo("EXTENCION")
					end if
				%>
				<tr><td align="center"><b>Nombre archivo :</b><%=nombre_archivo%></td><td></td></tr>
				<tr>
					<td width="400" valign="top">
						<iframe height="300" width="500" src="previsualizar_archivos.asp?archivo=<%=nombre_archivo%>&extencion=<%=extencion%>"></iframe>
					</td>
					<td valign="top" id="refreca_acccion_archivo"></td>
				</tr>

			</table>

		<%end if%>

		<%if trim(pagina_origen)<>"carga_masiva_archivos" then%>
			<div class="subtitulo_informe">> TIPO GESTIÓN</div>


	  	    <table width="100%" border="0" bordercolor="#FFFFFF" class="estilo_columnas">
	  	    <thead>
	        <tr bordercolor="#999999" bgcolor="#<%=session("COLTABBG")%>" class="Estilo13">
	          <td width="33%">CATEGORIA</td>
	          <td width="34%">SUBCATEGORIA</td>
	          <td width="33%">GESTION</td>
			  <td width="33%">&nbsp;</td>
	        </tr>
	    	</thead>
	        <tr bordercolor="#999999">
	          <td >
	          	<select name="cmbcat" id="cmbcat" onChange="cargasubcat(this.value);" >
	          <option value="">SELECCIONE</option>
	    <%

	          AbrirSCG1()
			  
				strSql = "SELECT DISTINCT A.COD_CATEGORIA, A.DESCRIPCION FROM GESTIONES_TIPO_CATEGORIA A, GESTIONES_TIPO_SUBCATEGORIA B, GESTIONES_TIPO_GESTION C "
				strSql = strSql & " WHERE A.COD_CATEGORIA = B.COD_CATEGORIA "
				strSql = strSql & " AND B.COD_CATEGORIA = C.COD_CATEGORIA "
				strSql = strSql & " AND B.COD_SUB_CATEGORIA = C.COD_SUB_CATEGORIA "
				strSql = strSql & " AND	C.COD_CLIENTE = '" & strCodCliente & "'"

				if TraeSiNo(session("perfil_adm")) <> "Si" and TraeSiNo(session("perfil_full")) <> "Si" Then

					if TraeSiNo(session("perfil_sup")) = "Si" and TraeSiNo(session("perfil_emp")) = "No" Then
						strSql=strSql & " AND ISNULL(VER_SUPERVISOR,0) = 1"
					Elseif TraeSiNo(session("perfil_cob")) = "Si" and TraeSiNo(session("perfil_emp")) = "No" Then
						strSql=strSql & " AND ISNULL(VER_COBRADOR,0) = 1"
					Elseif TraeSiNo(session("perfil_sup")) = "Si" and TraeSiNo(session("perfil_emp")) = "Si" Then
						strSql=strSql &  " AND ISNULL(VER_CLIENTE,0) = 1"
					Elseif TraeSiNo(session("perfil_cob")) = "Si" and TraeSiNo(session("perfil_emp")) = "Si" Then
						strSql=strSql &  " AND ISNULL(VER_CLIENTE_COBRADOR,0) = 1"
					End If

				End If

				if trim(pagina_origen)="casos_objetados" then
					strCategoria = 1
	          	end if					

			  	set rsGestCat=Conn1.execute(strSql)
			  	Do While not rsGestCat.eof

			  	strCategoriaCB = rsGestCat("COD_CATEGORIA")
				strSel=""
				if Trim(strCategoriaCB) = Trim(strCategoria) Then strSel = "SELECTED" Else strSel = ""
		%>
			  	<option value="<%=rsGestCat("COD_CATEGORIA")%>" <%=strSel%>><%=rsGestCat("DESCRIPCION")%></option>
		<%
			  	rsGestCat.movenext
			  	Loop
			  	rsGestCat.close
			  	set rsGestCat=nothing
			  CerrarSCG1()
		%>
	          </select>  
	          </td>
	          <td id="refresca_subcategoria">
	          <%
	          	if trim(pagina_origen)="casos_objetados" then
					AbrirSCG2()

					strSql = "SELECT DISTINCT A.COD_CATEGORIA, A.COD_SUB_CATEGORIA , A.DESCRIPCION "
					strSql = strSql & " FROM GESTIONES_TIPO_SUBCATEGORIA A, GESTIONES_TIPO_GESTION B "
					strSql = strSql & " WHERE A.COD_CATEGORIA = 1 "
					strSql = strSql & " AND A.COD_CATEGORIA = B.COD_CATEGORIA "
					strSql = strSql & " AND A.COD_SUB_CATEGORIA = B.COD_SUB_CATEGORIA "
					strSql = strSql & " AND B.COD_CLIENTE='" & strCodCliente & "'"

					if TraeSiNo(session("perfil_adm")) <> "Si" and TraeSiNo(session("perfil_full")) <> "Si" Then

							if TraeSiNo(session("perfil_sup")) = "Si" and TraeSiNo(session("perfil_emp")) = "No" Then
								strSql=strSql & " AND ISNULL(VER_SUPERVISOR,0) = 1"
							Elseif TraeSiNo(session("perfil_cob")) = "Si" and TraeSiNo(session("perfil_emp")) = "No" Then
								strSql=strSql & " AND ISNULL(VER_COBRADOR,0) = 1"
							Elseif TraeSiNo(session("perfil_sup")) = "Si" and TraeSiNo(session("perfil_emp")) = "Si" Then
								strSql=strSql &  " AND ISNULL(VER_CLIENTE,0) = 1"
							Elseif TraeSiNo(session("perfil_cob")) = "Si" and TraeSiNo(session("perfil_emp")) = "Si" Then
								strSql=strSql &  " AND ISNULL(VER_CLIENTE_COBRADOR,0) = 1"
							End If

					End If

					set rsGestSubCat=Conn2.execute(strSql)
					%>
					<select name="cmbsubcat" id="cmbsubcat" onChange="cargagest(this.value,cmbcat.value);"  >
					<%
					Do While not rsGestSubCat.eof
						%>
						<option value="<%=rsGestSubCat("COD_SUB_CATEGORIA")%>" <%if trim(rsGestSubCat("COD_SUB_CATEGORIA"))="1" then response.write " selected " end if%>><%=rsGestSubCat("DESCRIPCION")%></option>
						<%
					rsGestSubCat.movenext
					Loop
					rsGestSubCat.close
					set rsGestSubCat=nothing
					%></select><%
					CerrarSCG2()

				else%>

		          	<select name="cmbsubcat" id="cmbsubcat" onChange="cargagest(this.value,cmbcat.value);"  >
				  	  <option value="">SELECCIONE</option>
		          	</select>

				<%end if%>
	      </td>
	          <td id="refresca_gestion">
	          	<%if trim(pagina_origen)="casos_objetados" then
	          			AbrirSCG2()
						strSql="SELECT * FROM GESTIONES_TIPO_GESTION WHERE COD_CATEGORIA = 1 AND COD_SUB_CATEGORIA = 1 " 
						strSql = strSql & " AND COD_CLIENTE='" & strCodCliente & "'"

						if TraeSiNo(session("perfil_adm")) <> "Si" and TraeSiNo(session("perfil_full")) <> "Si" Then

							if TraeSiNo(session("perfil_sup")) = "Si" and TraeSiNo(session("perfil_emp")) = "No" Then
								strSql=strSql & " AND ISNULL(VER_SUPERVISOR,0) = 1"
							Elseif TraeSiNo(session("perfil_cob")) = "Si" and TraeSiNo(session("perfil_emp")) = "No" Then
								strSql=strSql & " AND ISNULL(VER_COBRADOR,0) = 1"
							Elseif TraeSiNo(session("perfil_sup")) = "Si" and TraeSiNo(session("perfil_emp")) = "Si" Then
								strSql=strSql &  " AND ISNULL(VER_CLIENTE,0) = 1"
							Elseif TraeSiNo(session("perfil_cob")) = "Si" and TraeSiNo(session("perfil_emp")) = "Si" Then
								strSql=strSql &  " AND ISNULL(VER_CLIENTE_COBRADOR,0) = 1"
							End If

						End If
						set rsGestion=Conn2.execute(strSql)

						If Not rsGestion.Eof Then
						%>
						 <select name="cmbgest" id="cmbgest" onChange="cajas_tipo_gestion();">
						 <option value="">SELECCIONE</option>	
						<%
							Do While Not rsGestion.Eof
							%>
								<option value="<%=rsGestion("COD_GESTION")%>"><%=rsGestion("DESCRIPCION")%></option>
							<%
							rsGestion.movenext
							Loop
						
						End if
						%>
						</select>
						<%
						CerrarSCG2()

	          	else%>

		          <select name="cmbgest" id="cmbgest" onChange="cajas_tipo_gestion();">
				   <option value="">SELECCIONE</option>
		          </select>

		        <%end if%>  
	      </td>
		  <td>
			<input  id="divAgendSC" name="divAgendSC" type="button" 	class="fondo_boton_100" STYLE="MARGIN-TOP:2PX; <%=muestraocultadivdivAgendSC%>" Value="Agendar Sin Contacto" onClick="ventana_procesa();">	
		  </td>
		</tr>
		</table>

		<div id="cajas_tipo_gestion"></div>

		<%end if%>
<input type="hidden" name="rut_" 				id="rut_" 				value="<%=strRutDeudor%>">
<input type="hidden" name="cliente_" 			id="cliente_" 			value="<%=strCodCliente%>">
<input type="hidden" name="intIdCampana" 		id="intIdCampana" 		value="<%=intIdCampana%>">

<input type="hidden" name="hd_gest_cp" 			id="hd_gest_cp"  		value="">
<input type="hidden" name="hd_gest_norm" 		id="hd_gest_norm"  		value="">

<input type="hidden" name="hd_tipo_agend" 		id="hd_tipo_agend"  	value="">
<input type="hidden" name="hd_tipo_gestion" 	id="hd_tipo_gestion"  	value="">
<input type="hidden" name="hd_obs_cliente" 		id="hd_obs_cliente"  	value="">

<input type="hidden" name="strContactoSel" 		id="strContactoSel"  	value="<%=strContactoSel%>">
<input type="hidden" name="strFonoAgestionar" 	id="strFonoAgestionar"  value="<%=strFonoAgestionar%>">
<input type="hidden" name="strCodCliente" 		id="strCodCliente"  	value="<%=strCodCliente%>">
<input type="hidden" name="pagina_origen" 		id="pagina_origen"  	value="<%=pagina_origen%>">

<input type="hidden" name="intIdSegmentoVenc" 	id="intIdSegmentoVenc"  value="<%=intIdSegmentoVenc%>">
<input type="hidden" name="intIdSegmentoMonto" 	id="intIdSegmentoMonto" value="<%=intIdSegmentoMonto%>">
<input type="hidden" name="intIdSegmentoAsig" 	id="intIdSegmentoAsig"  value="<%=intIdSegmentoAsig%>">


<div id="ventana_procesa" title="Agendar Sin Contacto" style="display:none;">	
	<table align="center" width="500" align="right" cellSpacing="0" cellPadding="0" border="0">
		<tr>		
			<td align="left" colspan="2" class="titulo_informe" width="200">> AGENDAMIENTO GESTION SIN CONTACTO</td>		
		</tr>
		<tr>		
			<td align="left" class="estilo_columna_individual" width="200">FECHA</td>		
			<td align="left" class="estilo_columna_individual" width="200">HORA</td>
		</tr>	
		<tr>
			<td align="left" class="" width="">
				<input name="TX_FEC_AGEND_SC" readonly type="text" id="TX_FEC_AGEND_SC" size="10" maxlength="10" onBlur="ValidaDifFechas();" value="<%=date()%>">
			</td>		
			<td align="left" class="" width="">
				<input name="TX_HORAAGEND_SC" type="text" id="TX_HORAAGEND_SC" size="5" maxlength="5" onChange="return ValidaHora(this,this.value)">
			</td>		
		</tr>	
		<tr><td colspan="2">&nbsp;</td></tr>	
		<tr>		
			<td align="left" colspan="2" class="titulo_informe">> OBSERVACIÓN</td>	
		</TR>
		<TR>
			<td align="left" colspan="2" >
				<TEXTAREA id="TX_OBSERVACION_CONSULTA" maxlength="199" style="width:350px; height:50px;" name="TX_OBSERVACION_CONSULTA"></TEXTAREA>
			</td>				
		</tr>
	</table>
</div>

</form>

<br>
<%if trim(pagina_origen)<>"carga_masiva_archivos" then%>
	<div class="subtitulo_informe" style="float:left;">
		> HISTORIAL DE GESTIONES
	</div>
	<div style="float:right;">
		<span class="subtitulo_informe">> Filtro Gestiones</span>
		<select name="CB_FILTRO" id="CB_FILTRO" onchange="filtro_historial(this.value)" >
			<option value="TODAS" >TODAS</option>
			<option value="EFECTIVAS ACTIVAS" selected="selected">ACTIVAS EFECTIVAS</option>
			<option value="ACTIVAS CALL">ACTIVAS CALL</option>
			<option value="ACTIVAS MASIVO">ACTIVAS MASIVO</option>
			<option value="EFECTIVAS">EFECTIVAS</option>
		</select>		
	</div>
	<table class="estilo_columnas" style="width:100%;" border="1" cellSpacing="0" cellPadding="0">
	<thead>
	<tr>
		<td width="20" >&nbsp;</td>
		<td width="20" >&nbsp;</td>
		<td width="70">FECHA</td>
		<td width="50">HORA</td>
		<td width="350">GESTION</td>
		<td width="60">F.COMP.</td>
		<td width="60">F.AGEND</td>
		<td width="50">H.AGEND</td>
		<td width="350">OBSERVACIONES</td>
		<td width="80">MEDIO G.</td>
		<td width="20">&nbsp;</td>
		<td width="20">F.AG.</td>
		<td width="80">EJECUTIVO</td>
		<td width="20">&nbsp;</td>
		<% if ((TraeSiNo(session("perfil_sup"))="Si" or  TraeSiNo(session("perfil_adm"))="Si") and TraeSiNo(session("perfil_emp"))<>"Si") Then %>
			<td width="20" class="Estilo4">&nbsp;</td>
		<% End If %>
	</tr>
	</thead>
	</table>

	<div id="frame2" style="width:100%;"></div>
	<br>
	<br>
<%end if%>
</body>
</html>

<style type="text/css" media="screen">
.festivos span {
 color: red !important; //muestra rojos los festivos
}
.ui-datepicker-week-end span {
 color: #333 !important; //muestra grises los fines de semana
}	
</style>

<script type="text/javascript">

	function carga_funcion_telefono(){

		var strFonoAgestionar 		=$('#strFonoAgestionar').val()
		var rut 			 		=$('#rut_').val()
		var strCodCliente   		=$('#strCodCliente').val()
		var intUsaDiscador			='<%=intUsaDiscador%>'

		var criterios ="alea="+Math.random()+"&strFonoAgestionar="+strFonoAgestionar+"&strRUT_DEUDOR="+rut+"&strCOD_CLIENTE="+strCodCliente+"&intUsaDiscador="+intUsaDiscador

		$('#carga_funcion').load('deudor_telefonos.asp', criterios, function(data){})
	}

	function bt_script_observacion(script_gestion){
		if (script_gestion=="borrar"){
			$('#TX_OBSERVACIONES').val("")
		}else{
			$('#TX_OBSERVACIONES').val(script_gestion)
		}			
	} 

	function filtro_historial(valor){
		var rut	 					=$('#rut_').val() 
		var strCodCliente   		=$('#strCodCliente').val()
		
		var criterios ="alea="+Math.random()+"&accion_ajax=refresca_historial&rut="+rut+"&strCodCliente="+strCodCliente+"&inicio=1&finales=25&CB_FILTRO="+valor+"&fono_actual="+$("#fonoActual").val()
		$('#frame2').load('FuncionesAjax/detalle_gestiones_ajax.asp', criterios, function(data){
			$('.td_hover').hover(function(){
				$(this).css('background-color','#CEE3F6')
			}, function(){
				$(this).css('background-color','')
			})
		})
	}

	function trae_cuotas_por_gestion(ID_GESTION){
		parent.trae_cuotas_por_gestion(ID_GESTION)

	}
	function Refrescar(){
		FrmHistorial.action='HistorialGestiones.asp?strRUT_DEUDOR=<%=strRutDeudor%>&strCOD_CLIENTE=<%=strCodCliente%>';
		FrmHistorial.submit();
	}

	function TraerGrabacion (strTelefono,strFecIngreso,strHoraIngreso,intIdusuario,strAnexo){
    	URL='EscucharGrabacion.asp?strTelefono=' + strTelefono + '&strFecIngreso=' + strFecIngreso + '&strHoraIngreso=' + strHoraIngreso + '&intIdusuario=' + intIdusuario + '&strAnexo=' + strAnexo
		window.open(URL,"DATOS_GRABACION","width=470, height=230, scrollbars=no, menubar=no, location=no, resizable=yes")
	}

	function ConfirmarCP(id_gestion, dtmFecCompGest, intCodGestConcat,mostrar)
	{
		datos.action = "confirmar_cp.asp?id_gestion=" + id_gestion + "&rut=<%=rut%>&cliente=<%=cliente%>&dtmFecCompGest=" + dtmFecCompGest + "&intCodGestConcat=" + intCodGestConcat+ "&mostrar=" + mostrar ;
		datos.submit();
	}

	function bt_mostrar_mas_registros(inicio, finales){

		var rut						=$('#rut_').val()
		var strCodCliente   		=$('#strCodCliente').val()
		var CB_FILTRO 	=$('#CB_FILTRO').val()

		inicio = parseInt(inicio) + 1
		var concat ="#refreso_mas_registros_"+finales
		var criterios ="alea="+Math.random()+"&accion_ajax=refresca_historial&inicio="+inicio+"&finales="+finales+"&rut="+rut+"&strCodCliente="+strCodCliente+"&CB_FILTRO="+CB_FILTRO+"&fono_actual="+$("#fonoActual").val()
		$(concat).load('FuncionesAjax/detalle_gestiones_ajax.asp', criterios, function(){
			$('.td_hover').hover(function(){
				$(this).css('background-color','#CEE3F6')
			}, function(){
				$(this).css('background-color','')
			})			
		})
	}

	function marcar_boxes(){
		var TX_MONTO_CANCELADO =	$('#TX_MONTO_CANCELADO').val()
		
		$('#TX_CAPITAL').val(0)
		$('#TX_INTERESES').val(0)
		$('#TX_PROTESTOS').val(0)
		$('#TX_HONORARIOS').val(0)
		$('#TX_SALDO').val(0)

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

		datos.TX_CAPITAL.value = 0;
		datos.TX_INTERESES.value = 0;
		datos.TX_PROTESTOS.value = 0;
		datos.TX_HONORARIOS.value = 0;
		datos.TX_SALDO.value = 0;

		if(TX_MONTO_CANCELADO!=null){
			$('#TX_MONTO_CANCELADO').val(0)
		}

		$('#span_TX_SALDO').text(0)
		$('#span_TX_CAPITAL').text(0)	

		$('input[id="CH_ID_CUOTA"]').each(function(){	
			$(this).removeAttr('checked');	
		})
	}

	function cajas_tipo_gestion(){
		var cmbcat  	=$('#cmbcat').val()
		var cmbsubcat  	=$('#cmbsubcat').val()
		var cmbgest  	=$('#cmbgest').val()
		var rut  		=$('#rut_').val()
		var strCodCliente   	=$('#strCodCliente').val()
		
		var TX_SALDO 	=$('#TX_SALDO').val()

		var criterios="alea="+Math.random()+"&accion_ajax=muestra_cajas_tipo_gestion&cmbcat="+cmbcat+"&cmbsubcat="+cmbsubcat+"&cmbgest="+cmbgest+"&rut="+rut+"&strCodCliente="+strCodCliente+"&TX_SALDO="+TX_SALDO+"&fono_actual="+$("#fonoActual").val()
		
       $('#cajas_tipo_gestion').load('FuncionesAjax/detalle_gestiones_ajax.asp', criterios, function(){

			$('#TX_FECHA_COMPROMISO').datepicker( {changeMonth: true,changeYear: true, dateFormat: 'dd/mm/yy',
			    beforeShowDay: DisableDays })

			$('#TX_FECHA_PAGO').datepicker( {changeMonth: true,changeYear: true, dateFormat: 'dd/mm/yy',
			    beforeShowDay: DisableDays })

			$('#TX_FEC_AGEND').datepicker( {changeMonth: true,changeYear: true, dateFormat: 'dd/mm/yy',
			    beforeShowDay: DisableDays })
		})
	}


function LugarPago()
{
			var forma_Pago 		= $('#CB_FORMA_PAGO').val();	
			var strCodCliente   =$('#strCodCliente').val()
            var rut  		=$('#rut_').val()
            $.ajax({url: "FuncionesAjax/detalle_gestiones_ajax.asp?accion_ajax=refresca_lugar_pago"+"&fono_actual="+$("#fonoActual").val(), 
			method: "POST", 
			data: {forma_Pago: forma_Pago,strCodCliente:strCodCliente,rut:rut},
			success: function(data) {
			             
				$("#td_CB_ID_DIRECCION_COBRO_DEUDOR").html(data);
			},
			error: function(XMLHttpRequest, textStatus, errorThrown){
			     alert("Error al procesar los datos ('"+errorThrown+"')");
			}
           });
}
	function asigna_minimo(campo, minimo1){
		if(campo.value!=0){
			if(campo.value==41 || campo.value==32 || campo.value==2){
				minimo1=7;
			}else if(campo.value.length==1){
				minimo1=8;
			}else {
				minimo1=7;
			}
		}else{minimo1=10}
		return(minimo1)
	}
		function asigna_minimo_nuevo(campo, minimo1){
			if (campo!=0)	{
				if(campo==41 || campo==32 || campo==45 || campo==57 || campo==55 || campo==72 || campo==71 || campo==73 || campo==75){
					minimo1=7;
				}else if(campo.length==1 || campo==2){
					minimo1=8;
				}else {
					minimo1=7;
				}
			}else{minimo1=10}
			return(minimo1)
		}

	function valida_largo(campo, minimo){
		var numero = $('#numero').val();
		var codigoArea = $('#COD_AREA').val();
		var inicioNumero = numero.substring(0,3);
		
		if(codigoArea != 0) {
			if(campo.length != minimo) {
				alert("Fono debe tener " + minimo + " digitos");
				$('#numero').select();
				$('#numero').focus();
				return(true);
			}	
		} else {
		
			if (numero.length < 9) {
				alert("El numero ingresado no es valido, debe ser un numero con prefijo 600 y de largo 10 o prefijo 800 o 197 y de largo 9.");
				return(true);
			}
			
			if (numero.length == 9 && inicioNumero != 800 && inicioNumero != 197) {
				alert("El numero ingresado no es valido, debe ser un numero con prefijo 800 o 197 y de largo 9.");
				return(true);
			}
			
			if (numero.length == 10 && inicioNumero != 600) {
				alert("El numero ingresado no es valido, debe ser un numero con prefijo 600 y de largo 10.");
				return(true);
			}
		}		
		return(false);
	}

	function valida_largo_nuevo(campo, minimo){
		var numero = $('#numero').val();
		var codigoArea = $('#COD_AREA').val();
		var inicioNumero = numero.substring(0,3);
		
		if(codigoArea != 0) {
			if(campo.length != minimo) {
				alert("Fono debe tener " + minimo + " digitos");
				$('#numero').select();
				$('#numero').focus();
				return(true);
			}	
		} else {
		
			if (numero.length < 9) {
				alert("El numero ingresado no es valido, debe ser un numero con prefijo 600 y de largo 10 o prefijo 800 o 197 y de largo 9.");
				return(true);
			}
		
			if (numero.length == 9 && inicioNumero != 800 && inicioNumero != 197) {
				alert("El numero ingresado no es valido, debe ser un numero con prefijo 800 o 197 y de largo 9.");
				return(true);
			}
			
			if (numero.length == 10 && inicioNumero != 600) {
				alert("El numero ingresado no es valido, debe ser un numero con prefijo 600 y de largo 10.");
				return(true);
			}
		}		
		return(false);
	}

	function solonumero(valor){
      if (isNaN(valor.value)) {
            valor.value=""
			return ""
      }else{
			valor.value
			return valor.value
      }
	}

    

	function ingreso_nueva_gestion(){
		var IntTipoGestion 				=$('#TIPO_GESTION').val()
		var intMedioAsociado			=$('#MEDIO_ASOCIADO').val()
		var strSeleccionado				="N"
		var strSeleccionadoDocGestion 	="N"
		var concat_cuotas_deudor 		=""
		var concat_doc_gestion	 		=""
		var rut  						=$('#rut_').val()
		var intIdCampana 				=$('#intIdCampana').val()
		var cmbcat 						=$('#cmbcat').val()
		var cmbsubcat 					=$('#cmbsubcat').val()
		var cmbgest 					=$('#cmbgest').val()
		var strGestionTipoGestion		=cmbcat+"*"+cmbsubcat+"*"+cmbgest
		var strPaginaOrigen 			=$('#pagina_origen').val()
		var strCodCliente   			= $('#strCodCliente').val()
		var dtmFechaCompromiso 			=$('#TX_FECHA_COMPROMISO').val()
		var strNroDocPago 				=$('#TX_NRO_DOC_PAGO').val()
		var dtmFechaPago				=$('#TX_FECHA_PAGO').val()
		var strHoraDesde				=$('#TX_HORA_DESDE').val()
		var strHoraHasta				=$('#TX_HORA_HASTA').val()
		var strFormaPago				=$('#CB_FORMA_PAGO').val()
		var strDocGestionNecesarios		=$('#TX_DOC_GESTION_NECESARIOS').val()
		var IntMontoCancelado 			=$('#TX_MONTO_CANCELADO').val()
		var strEnvioHdr					=$('#CB_ENVIO_HRD').val()
		var intIdFonoCobro				=$('#CB_ID_FONO_COBRO').val()
		var intIdContactoFonoCobro 		=$('#CB_ID_CONTACTO_FONO_COBRO').val()
		var intIdDireccionCobroDeudor 	=$('#CB_ID_DIRECCION_COBRO_DEUDOR').val()
		var strObservaciones 	 		=$('#TX_OBSERVACIONES').val()
		var intIdSegmentoVenc 	 		=$('#intIdSegmentoVenc').val()		
		var intIdSegmentoMonto 	 		=$('#intIdSegmentoMonto').val()		
		var intIdSegmentoAsig 	 		=$('#intIdSegmentoAsig').val()	

		
		if ("<%=bitAgendarSinContacto%>"=="True" || "<%=pagina_origen%>"=="agendamiento_tactico")
			MostrarFilas('divAgendSC');

		if(cmbcat==""){
			alert("¡Debe ingresar categoria!")
			return
		}
		if(cmbsubcat==""){
			alert("¡Debe ingresar subcategoria!")
			return
		}
		if(cmbgest==""){
			alert("¡Debe ingresar tipo gestión!")
			return
		}

		$('input[id="CH_ID_CUOTA"]:checked').each(function(){
			strSeleccionado ="S"
         	concat_cuotas_deudor = concat_cuotas_deudor +","+$(this).val()
		})

		$('input[id="CK_DOC_GESTION"]:checked').each(function(){
			strSeleccionadoDocGestion ="S"
			concat_doc_gestion = concat_doc_gestion +","+$(this).val()
		})	
		
		if(strSeleccionado=="N"){
			alert('Debe asociar al menos 1 documento a la gestión');
			return
		}

		if(IntTipoGestion==1){ //	COMPROMISO PAGO

			if(dtmFechaCompromiso==""){
				alert("Debe ingresar fecha compromiso")
				return
			}
			if(strFormaPago==""){
				alert("Debe ingresar forma pago")
				return
			}
			if(intIdDireccionCobroDeudor==""){
				alert("Debe ingresar lugar pago")
				return
			}

		}

		if(IntTipoGestion==2){ // COMPROMISO PAGO RUTA 

			if(dtmFechaCompromiso==""){
				alert("Debe ingresar fecha compromiso")
				return
			}

			if(strHoraDesde==""){
				alert("Debe ingresar hora desde")
				return
			}

			if(strHoraHasta==""){
				alert("Debe ingresar hora hasta")
				return
			}

			if(strFormaPago==""){
				alert("Debe ingresar forma pago")
				return
			}

			if(intIdDireccionCobroDeudor==""){
				alert("Debe ingresar lugar pago")
				return
			}


		}

		if(IntTipoGestion==3){ // GESTION 3 NORMALIZACION ( INDICA QUE PAGÓ) NO TIENE VALIDACIONES

		}
		
		if(IntTipoGestion==4){ // GESTION 4 NORMALIZACION ( INDICA QUE PAGÓ) NO TIENE VALIDACIONES

		}	
		
		
		if(IntTipoGestion==5){  // GESTION 5 VERIFICACION EN TERRENO 

			if(dtmFechaCompromiso==""){
				alert("Debe ingresar fecha gestión")
				return
			}

			if(strHoraDesde==""){
				alert("Debe ingresar hora desde")
				return
			}

			if(strHoraHasta==""){
				alert("Debe ingresar hora hasta")
				return
			}		

			if(intIdDireccionCobroDeudor==""){
				alert("Debe ingresar lugar gestión")
				return
			}

		}

		var dtmFecAgend 			=$('#TX_FEC_AGEND').val()
		var intIdMedioAgendamiento	=$('#CB_ID_MEDIO_AGENDAMIENTO').val()
		var intIdMedioGestion		=$('#CB_ID_MEDIO_GESTION').val()
		var strHoraAgend			=$('#TX_HORAAGEND').val()
		var intIdContactoGestion	=$('#CB_ID_CONTACTO_GESTION').val()
		var Obligatoriedad 			=$('#Obligatoriedad').val()

		if (Obligatoriedad==1) 
		{
			if(dtmFecAgend==""){
				alert("Debe ingresar fecha agendamiento")
				return
			}


			if(intIdMedioAgendamiento=="" && intIdMedioAgendamiento!=null){
				alert("Debe ingresar medio agendamiento")
				return
			}

			if(intIdMedioGestion=="" && intIdMedioGestion!=null){
				alert("Debe ingresar medio gestión")
				return
			}
		}
        
		$('#ingresar').prop('disabled', true);
        
		strDocGestion 	 		= concat_doc_gestion.substring(1, concat_doc_gestion.length)		
		concat_cuotas_deudor 	= concat_cuotas_deudor.substring(1, concat_cuotas_deudor.length)
		$('#cuotas_deudor').val(concat_cuotas_deudor)

		var criterios= "alea="+Math.random()+"&accion_ajax=ingreso_gestion&IntTipoGestion="+IntTipoGestion+"&cuotas_deudor="+concat_cuotas_deudor+"&strCodCliente="+strCodCliente+"&dtmFechaCompromiso="+dtmFechaCompromiso+"&strNroDocPago="+strNroDocPago+"&dtmFechaPago="+dtmFechaPago+"&strHoraDesde="+strHoraDesde+"&strHoraHasta="+strHoraHasta+"&strFormaPago="+strFormaPago+"&strDocGestion="+encodeURIComponent(strDocGestion)+"&strDocGestionNecesarios="+strDocGestionNecesarios+"&IntMontoCancelado="+IntMontoCancelado+"&strEnvioHdr="+strEnvioHdr+"&intIdFonoCobro="+intIdFonoCobro+"&intIdContactoFonoCobro="+intIdContactoFonoCobro+"&intIdDireccionCobroDeudor="+intIdDireccionCobroDeudor+"&dtmFecAgend="+dtmFecAgend+"&intIdMedioAgendamiento="+intIdMedioAgendamiento+"&intIdMedioGestion="+intIdMedioGestion+"&strHoraAgend="+strHoraAgend+"&intIdContactoGestion="+intIdContactoGestion+"&strObservaciones="+encodeURIComponent(strObservaciones)+"&rut="+rut+"&intIdCampana="+intIdCampana+"&strGestionTipoGestion="+strGestionTipoGestion+"&intMedioAsociado="+intMedioAsociado+"&strPaginaOrigen="+strPaginaOrigen+"&intIdSegmentoVenc="+intIdSegmentoVenc+"&intIdSegmentoMonto="+intIdSegmentoMonto+"&intIdSegmentoAsig="+intIdSegmentoAsig+"&fono_actual="+$("#fonoActual").val()         

		$('#ingreso_gestion').load('FuncionesAjax/detalle_gestiones_ajax.asp', criterios, function(){
		    $.prettyLoader.show(2000);
			var intIdGestion 				=$('#intIdGestion').val()
			var pagina_redireccionamiento 	=$('#pagina_redireccionamiento').val()
            
			//alert(pagina_redireccionamiento);

			if (strPaginaOrigen=="casos_objetados"){ 
				location.href="listado_expone_requerimientos.asp"
			}else{  
			    if (pagina_redireccionamiento=="confirmar_cp"){
					location.href="confirmar_cp.asp?id_gestion="+intIdGestion+"&dtmFecCompGest="+dtmFechaCompromiso	
				}		
			    else if (pagina_redireccionamiento=="principal"){ 

			        if("<%=intTipoAgendamiento%>"=="0" || "<%=pagina_origen%>"=="agendamiento_tactico"){

			            location.href="modulo_agendamiento_tactico.asp";
                        
			        }else if("<%=intTipoAgendamiento%>"=="1"){
                        
			            location.href="modulo_gestion_campanas.asp";

			        }else if("<%=intTipoAgendamiento%>"=="2"){
                        
			            location.href="modulo_gestion_campanas.asp";

			        }else{
			            location.href="principal.asp?a=1";
			        }
			    }
                
			}
			MostrarFilas("divAgendSC");

			var CB_FILTRO 	=$('#CB_FILTRO').val()

			var criterios ="alea="+Math.random()+"&accion_ajax=refresca_historial&rut="+rut+"&strCodCliente="+strCodCliente+"&inicio=1&finales=25&CB_FILTRO="+CB_FILTRO+"&fono_actual="+$("#fonoActual").val()
			$('#frame2').load('FuncionesAjax/detalle_gestiones_ajax.asp', criterios, function(){
				$('.td_hover').hover(function(){
					$(this).css('background-color','#CEE3F6')
				}, function(){
					$(this).css('background-color','')
				})

			})

			$('input[id="CH_TODOS_CUOTA"]').removeAttr('checked')

			Refrescar()   

			$('#ingresar').prop('disabled', false);

			if($("#fonoActual").val()!= ""){
                $("#cmbcat").val(2);
                var valorPrimeraOpcion = $('#cmbgest option:first').val();
                $('#cmbgest').val(valorPrimeraOpcion);

			}else{
			    $('#cmbcat').val("");
			    $('#cmbsubcat').val("");
			    $('#cmbgest').val("");

			    var criterios ="alea="+Math.random()+"&accion_ajax=refresca_subcategoria"+"&fono_actual="+$("#fonoActual").val()
			    $('#refresca_subcategoria').load('FuncionesAjax/detalle_gestiones_ajax.asp', criterios, function(){})			

			    var criterios ="alea="+Math.random()+"&accion_ajax=refresca_gestion"+"&fono_actual="+$("#fonoActual").val()
			    $('#refresca_gestion').load('FuncionesAjax/detalle_gestiones_ajax.asp', criterios, function(){})
			}

			$('#cajas_tipo_gestion').text("");

			OcultarFilas($('#cajas_tipo_gestion'));
		})
	}

	function trae_cuotas_por_gestion(ID_GESTION){
		var rut 			=$('#rut_').val()
		var strCodCliente   =$('#strCodCliente').val()
		
		$("#CB_COBRANZA option[value='']").attr('selected', 'selected');

		var criterios ="alea="+Math.random()+"&accion_ajax=mostrar_todos_cuotas&rut="+rut+"&strCodCliente="+strCodCliente+"&ID_GESTION="+ID_GESTION+"&fono_actual="+$("#fonoActual").val()

		$('#div_mostrar_todo').load('FuncionesAjax/detalle_gestiones_ajax.asp', criterios, function(){		
		
			$("#table_tablesorter").tablesorter({dateFormat: "uk"});
		})		
	}

	function cargaContacto(intIdTelefono, strCombo)
	{
		var comboBox = document.getElementById(strCombo);
		switch (intIdTelefono)
		{
			<%
			  AbrirSCG1()
				strSql="SELECT * FROM DEUDOR_TELEFONO WHERE RUT_DEUDOR = '" & strRutDeudor & "'"
				set rsDeudorTelefono=Conn1.execute(strSql)
				Do While not rsDeudorTelefono.eof
			%>

			case '<%=rsDeudorTelefono("ID_TELEFONO")%>':
				comboBox.options.length = 0;
				var newOption = new Option('SELECCIONE', '0');comboBox.options[comboBox.options.length] = newOption;
				<%
				AbrirSCG2()

				strSql=" SELECT 0 as ORDEN, ID_CONTACTO, CONTACTO FROM TELEFONO_CONTACTO WHERE ID_TELEFONO = " & rsDeudorTelefono("ID_TELEFONO")
				strSql = strSql & " UNION"
				strSql = strSql & " SELECT ORDEN, ID_CONTACTO, UPPER(CONTACTO_BASE) AS CONTACTO FROM CONTACTO_BASE WHERE COD_CLIENTE = '" & strCodCliente & "' ORDER BY ORDEN, ID_CONTACTO DESC "



				set rsContacto=Conn2.execute(strSql)
				Do While not rsContacto.eof
					%>
					var newOption = new Option('<%=UCASE(rsContacto("CONTACTO"))%>', '<%=rsContacto("ID_CONTACTO")%>');comboBox.options[comboBox.options.length] = newOption;
					<%
					rsContacto.movenext
				Loop
				rsContacto.close
				set rsContacto=nothing
				CerrarSCG2()
				%>
				break;


			<%
			  	rsDeudorTelefono.movenext
			  	Loop
			  	rsDeudorTelefono.close
			  	set rsDeudorTelefono=nothing
				CerrarSCG1()
			%>

		}
	}

	function Priorizar() {

		if (confirm("¿ Está seguro de priorizar ? De aceptar la priorizacion todos los documentos activos del deudor pasaran a PRIORIDAD 2,1 siempre y cuando esten gestionados y agendados."))
			{
				datos.BT_PRIORIZAR_EJE.disabled=true
				datos.action='detalle_gestiones_action.asp?strPrioriza=S';
				datos.submit();
			}
		else
			alert("proceso cancelado");

	}

	function agendamiento_gestion_sin_contacto(){
	    $('#Agendar_SC').prop('disabled', true); 
		var rut  					=$('#rut_').val()
		var strCodCliente   		=$('#strCodCliente').val()
		var dtmFecAgend 			=$('#TX_FEC_AGEND_SC').val()
		var strHoraAgend			=$('#TX_HORAAGEND_SC').val()
		var criterios= "alea="+Math.random()+"&accion_ajax=agendamiento_gestion_sin_contacto&rut="+rut+"&strCodCliente="+strCodCliente+"&dtmFecAgend="+dtmFecAgend+"&strHoraAgend="+strHoraAgend+"&fono_actual="+$("#fonoActual").val()       

		$('#agendamiento_gestion_sin_contacto').load('FuncionesAjax/detalle_gestiones_ajax.asp', criterios, function(){	
		    $.prettyLoader.show();
			
				if("<%=intTipoAgendamiento%>"=="0" || "<%=pagina_origen%>"=="agendamiento_tactico" ){

					location.href="modulo_agendamiento_tactico.asp";
					
				}else if("<%=intTipoAgendamiento%>"=="1"){
					
					location.href="modulo_gestion_campanas.asp";

				}else if("<%=intTipoAgendamiento%>"=="2"){
					
					location.href="modulo_gestion_campanas.asp";

				}else{
					location.href="principal.asp?a=1";
				}		    
		})
	}
	
	function valida_entero(ObjIng, valor){
		if(isNaN(valor)){
			alert("Formato monto cancelado invalido")
			ObjIng.value=''
			return
		} 

	}


	function Refrescar(){

		var rut 			 		=$('#rut_').val()
		var strCodCliente   		=$('#strCodCliente').val()
		var CH_TODOS_CUOTA 			=""

		if($('input[id="CH_TODOS_CUOTA"]').is(':checked')){
			CH_TODOS_CUOTA =1
		}else{
			CH_TODOS_CUOTA =0
		}
		
		var CB_COBRANZA = $("#CB_COBRANZA :selected").val();
		
		var criterios ="alea="+Math.random()+"&accion_ajax=mostrar_todos_cuotas&rut="+rut+"&strCodCliente="+strCodCliente+"&CH_TODOS_CUOTA="+CH_TODOS_CUOTA+"&CB_COBRANZA="+CB_COBRANZA+"&fono_actual="+$("#fonoActual").val()
		$('#div_mostrar_todo').load('FuncionesAjax/detalle_gestiones_ajax.asp', criterios, function()
		{
		    $("#table_tablesorter").tablesorter({dateFormat: "uk"}); 

		})
	}

	function cargaContactoEmail(intIdEmail, strCombo){
	{

		var comboBox = document.getElementById(strCombo);

		switch (intIdEmail)
		{
			<%
			  AbrirSCG1()
				strSql="SELECT * FROM DEUDOR_EMAIL WHERE RUT_DEUDOR = '" & rut & "'"
				set rsDeudorEmail=Conn1.execute(strSql)
				Do While not rsDeudorEmail.eof
			%>

			case '<%=rsDeudorEmail("ID_EMAIL")%>':
				comboBox.options.length = 0;
				var newOption = new Option('SELECCIONE', '0');comboBox.options[comboBox.options.length] = newOption;
				<%
				AbrirSCG2()
				strSql="SELECT * FROM EMAIL_CONTACTO WHERE ID_EMAIL = " & rsDeudorEmail("ID_EMAIL")
				set rsContacto=Conn2.execute(strSql)
				Do While not rsContacto.eof
					%>
					var newOption = new Option('<%=UCASE(rsContacto("CONTACTO"))%>', '<%=rsContacto("ID_CONTACTO")%>');comboBox.options[comboBox.options.length] = newOption;
					<%
					rsContacto.movenext
				Loop
				rsContacto.close
				set rsContacto=nothing
				CerrarSCG2()
				%>
				break;
			<%
			  	rsDeudorEmail.movenext
			  	Loop
			  	rsDeudorEmail.close
			  	set rsDeudorEmail=nothing
				CerrarSCG1()
			%>
		}
	}}
	function MostrarFilas(Fila) {
		var elementos = document.getElementsByName(Fila);
		for (k = 0; k< elementos.length; k++) {
				   elementos[k].style.display = "inline";
		}
	}
	function OcultarFilas(Fila) {
		var elementos = document.getElementsByName(Fila);
		for (k = 0; k< elementos.length; k++) {
				   elementos[k].style.display = "none";
		}
	}
	function cargasubcat(subCat)
	{
		
		if (($('#cmbcat').val()=="" && "<%=bitAgendarSinContacto%>"=="True") || "<%=pagina_origen%>"=="agendamiento_tactico")
		{
			MostrarFilas('divAgendSC');
		}
		else
		{
			OcultarFilas('divAgendSC');
		}

		var comboBox = document.getElementById('cmbsubcat');
		comboBox.options.length = 0;
		var newOption = new Option('SELECCIONE', '0');comboBox.options[comboBox.options.length] = newOption;
		
		var comboBox2 = document.getElementById('cmbgest');
		comboBox2.options.length = 0;
		var newOption = new Option('SELECCIONE', '0');comboBox2.options[comboBox2.options.length] = newOption;
		
		OcultarFilas('divCompPago');
		OcultarFilas('divCompPagoRuta');
		OcultarFilas('divNormalizacion');
		OcultarFilas('divNormalizacion1');
		OcultarFilas('divGestionTerreno');
		OcultarFilas('divObsGestion');
		OcultarFilas('divAgend');
		
		switch (subCat)

		{
			<%
			  AbrirSCG1()
				strSql="SELECT * FROM GESTIONES_TIPO_CATEGORIA"


				strSql = "SELECT DISTINCT A.COD_CATEGORIA, A.DESCRIPCION FROM GESTIONES_TIPO_CATEGORIA A, GESTIONES_TIPO_SUBCATEGORIA B, GESTIONES_TIPO_GESTION C "
				strSql = strSql & " WHERE A.COD_CATEGORIA = B.COD_CATEGORIA "
				strSql = strSql & " AND B.COD_CATEGORIA = C.COD_CATEGORIA "
				strSql = strSql & " AND B.COD_SUB_CATEGORIA = C.COD_SUB_CATEGORIA "
				strSql = strSql & " AND C.COD_CLIENTE='" & strCodCliente & "'"

				''Response.write "<br>strSql=" & strSql

				if TraeSiNo(session("perfil_adm")) <> "Si" and TraeSiNo(session("perfil_full")) <> "Si" Then

						if TraeSiNo(session("perfil_sup")) = "Si" and TraeSiNo(session("perfil_emp")) = "No" Then
							strSql=strSql & " AND ISNULL(VER_SUPERVISOR,0) = 1"
						Elseif TraeSiNo(session("perfil_cob")) = "Si" and TraeSiNo(session("perfil_emp")) = "No" Then
							strSql=strSql & " AND ISNULL(VER_COBRADOR,0) = 1"
						Elseif TraeSiNo(session("perfil_sup")) = "Si" and TraeSiNo(session("perfil_emp")) = "Si" Then
							strSql=strSql &  " AND ISNULL(VER_CLIENTE,0) = 1"
						Elseif TraeSiNo(session("perfil_cob")) = "Si" and TraeSiNo(session("perfil_emp")) = "Si" Then
							strSql=strSql &  " AND ISNULL(VER_CLIENTE_COBRADOR,0) = 1"
						End If

				End If

				''Response.write "strSql==================" & strSql

				set rsGestCat=Conn1.execute(strSql)
				Do While not rsGestCat.eof
			%>

		case '<%=rsGestCat("COD_CATEGORIA")%>':
				comboBox.options.length = 0;
				var newOption = new Option('SELECCIONE', '0');comboBox.options[comboBox.options.length] = newOption;
				<%
				AbrirSCG2()
				''strSql="SELECT * FROM GESTIONES_TIPO_SUBCATEGORIA WHERE COD_CATEGORIA = " & rsGestCat("COD_CATEGORIA")
				strSql = "SELECT DISTINCT A.COD_CATEGORIA, A.COD_SUB_CATEGORIA , A.DESCRIPCION "
				strSql = strSql & " FROM GESTIONES_TIPO_SUBCATEGORIA A, GESTIONES_TIPO_GESTION B "
				strSql = strSql & " WHERE A.COD_CATEGORIA = " & rsGestCat("COD_CATEGORIA")
				strSql = strSql & " AND A.COD_CATEGORIA = B.COD_CATEGORIA "
				strSql = strSql & " AND A.COD_SUB_CATEGORIA = B.COD_SUB_CATEGORIA "
				strSql = strSql & " AND B.COD_CLIENTE='" & strCodCliente & "'"

				if TraeSiNo(session("perfil_adm")) <> "Si" and TraeSiNo(session("perfil_full")) <> "Si" Then

						if TraeSiNo(session("perfil_sup")) = "Si" and TraeSiNo(session("perfil_emp")) = "No" Then
							strSql=strSql & " AND ISNULL(VER_SUPERVISOR,0) = 1"
						Elseif TraeSiNo(session("perfil_cob")) = "Si" and TraeSiNo(session("perfil_emp")) = "No" Then
							strSql=strSql & " AND ISNULL(VER_COBRADOR,0) = 1"
						Elseif TraeSiNo(session("perfil_sup")) = "Si" and TraeSiNo(session("perfil_emp")) = "Si" Then
							strSql=strSql &  " AND ISNULL(VER_CLIENTE,0) = 1"
						Elseif TraeSiNo(session("perfil_cob")) = "Si" and TraeSiNo(session("perfil_emp")) = "Si" Then
							strSql=strSql &  " AND ISNULL(VER_CLIENTE_COBRADOR,0) = 1"
						End If
				End If

				''Response.write "<br>strSql ========= " & strSql

				set rsGestSubCat=Conn2.execute(strSql)
				Do While not rsGestSubCat.eof
					%>
					var newOption = new Option('<%=rsGestSubCat("DESCRIPCION")%>', '<%=rsGestSubCat("COD_SUB_CATEGORIA")%>');comboBox.options[comboBox.options.length] = newOption;
					<%
					rsGestSubCat.movenext
				Loop
				rsGestSubCat.close
				set rsGestSubCat=nothing
				CerrarSCG2()
				%>
				break;

			<%
			  	rsGestCat.movenext
			  	Loop
			  	rsGestCat.close
			  	set rsGestCat=nothing
				CerrarSCG1()
			%>
			
		}
	}

	function cargagest(subCat,cat){
	{
		var comboBox = document.getElementById('cmbgest');
		comboBox.options.length = 0;

		OcultarFilas('divCompPago');
		OcultarFilas('divCompPagoRuta');
		OcultarFilas('divNormalizacion');
		OcultarFilas('divNormalizacion1');
		OcultarFilas('divGestionTerreno');
		OcultarFilas('divObsGestion');
		OcultarFilas('divAgend');
		
		switch (cat)
		{
			<%
			  AbrirSCG()
				strSql="SELECT * FROM GESTIONES_TIPO_CATEGORIA"

				strSql = "SELECT DISTINCT A.COD_CATEGORIA, A.DESCRIPCION FROM GESTIONES_TIPO_CATEGORIA A, GESTIONES_TIPO_SUBCATEGORIA B, GESTIONES_TIPO_GESTION C "
				strSql = strSql & " WHERE A.COD_CATEGORIA = B.COD_CATEGORIA "
				strSql = strSql & " AND B.COD_CATEGORIA = C.COD_CATEGORIA "
				strSql = strSql & " AND B.COD_SUB_CATEGORIA = C.COD_SUB_CATEGORIA "
				strSql = strSql & " AND C.COD_CLIENTE='" & strCodCliente & "'"
				if TraeSiNo(session("perfil_adm")) <> "Si" and TraeSiNo(session("perfil_full")) <> "Si" Then

						if TraeSiNo(session("perfil_sup")) = "Si" and TraeSiNo(session("perfil_emp")) = "No" Then
							strSql=strSql & " AND ISNULL(VER_SUPERVISOR,0) = 1"
						Elseif TraeSiNo(session("perfil_cob")) = "Si" and TraeSiNo(session("perfil_emp")) = "No" Then
							strSql=strSql & " AND ISNULL(VER_COBRADOR,0) = 1"
						Elseif TraeSiNo(session("perfil_sup")) = "Si" and TraeSiNo(session("perfil_emp")) = "Si" Then
							strSql=strSql &  " AND ISNULL(VER_CLIENTE,0) = 1"
						Elseif TraeSiNo(session("perfil_cob")) = "Si" and TraeSiNo(session("perfil_emp")) = "Si" Then
							strSql=strSql &  " AND ISNULL(VER_CLIENTE_COBRADOR,0) = 1"
						End If

				End If

				set rsGestCat=Conn.execute(strSql)
				Do While not rsGestCat.eof
			%>
			case '<%=rsGestCat("COD_CATEGORIA")%>':
				comboBox.options.length = 0;
				var newOption = new Option('SELECCIONE', 'X');comboBox.options[comboBox.options.length] = newOption;
				
				<%
				AbrirSCG1()
				''strSql="SELECT * FROM GESTIONES_TIPO_SUBCATEGORIA WHERE COD_CATEGORIA = " & rsGestCat("COD_CATEGORIA")

				strSql = "SELECT DISTINCT A.COD_CATEGORIA, A.COD_SUB_CATEGORIA , A.DESCRIPCION "
				strSql = strSql & " FROM GESTIONES_TIPO_SUBCATEGORIA A, GESTIONES_TIPO_GESTION B "
				strSql = strSql & " WHERE A.COD_CATEGORIA = " & rsGestCat("COD_CATEGORIA")
				strSql = strSql & " AND A.COD_CATEGORIA = B.COD_CATEGORIA "
				strSql = strSql & " AND A.COD_SUB_CATEGORIA = B.COD_SUB_CATEGORIA "
				strSql = strSql & " AND B.COD_CLIENTE='" & strCodCliente & "'"

				if TraeSiNo(session("perfil_adm")) <> "Si" and TraeSiNo(session("perfil_full")) <> "Si" Then

						if TraeSiNo(session("perfil_sup")) = "Si" and TraeSiNo(session("perfil_emp")) = "No" Then
							strSql=strSql & " AND ISNULL(VER_SUPERVISOR,0) = 1"
						Elseif TraeSiNo(session("perfil_cob")) = "Si" and TraeSiNo(session("perfil_emp")) = "No" Then
							strSql=strSql & " AND ISNULL(VER_COBRADOR,0) = 1"
						Elseif TraeSiNo(session("perfil_sup")) = "Si" and TraeSiNo(session("perfil_emp")) = "Si" Then
							strSql=strSql &  " AND ISNULL(VER_CLIENTE,0) = 1"
						Elseif TraeSiNo(session("perfil_cob")) = "Si" and TraeSiNo(session("perfil_emp")) = "Si" Then
							strSql=strSql &  " AND ISNULL(VER_CLIENTE_COBRADOR,0) = 1"
						End If
				End If

				''Response.write "<br>strSql ========= " & strSql

				set rsGestSubCat=Conn1.execute(strSql)
				If Not rsGestSubCat.eof Then
					Do While not rsGestSubCat.eof
						%>
						if (subCat=='<%=rsGestSubCat("COD_SUB_CATEGORIA")%>') {
							<%

							AbrirSCG2()
							strSql="SELECT * FROM GESTIONES_TIPO_GESTION WHERE COD_CATEGORIA = " & rsGestCat("COD_CATEGORIA") & " AND COD_SUB_CATEGORIA = " & rsGestSubCat("COD_SUB_CATEGORIA")
							strSql = strSql & " AND COD_CLIENTE='" & strCodCliente & "'"

							if TraeSiNo(session("perfil_adm")) <> "Si" and TraeSiNo(session("perfil_full")) <> "Si" Then

								if TraeSiNo(session("perfil_sup")) = "Si" and TraeSiNo(session("perfil_emp")) = "No" Then
									strSql=strSql & " AND ISNULL(VER_SUPERVISOR,0) = 1"
								Elseif TraeSiNo(session("perfil_cob")) = "Si" and TraeSiNo(session("perfil_emp")) = "No" Then
									strSql=strSql & " AND ISNULL(VER_COBRADOR,0) = 1"
								Elseif TraeSiNo(session("perfil_sup")) = "Si" and TraeSiNo(session("perfil_emp")) = "Si" Then
									strSql=strSql &  " AND ISNULL(VER_CLIENTE,0) = 1"
								Elseif TraeSiNo(session("perfil_cob")) = "Si" and TraeSiNo(session("perfil_emp")) = "Si" Then
									strSql=strSql &  " AND ISNULL(VER_CLIENTE_COBRADOR,0) = 1"
								End If
							End If

							''Response.write "sql=" & strSql
							set rsGestion=Conn2.execute(strSql)
							If Not rsGestion.Eof Then
								Do While Not rsGestion.Eof

									if (Trim(rsGestCat("COD_CATEGORIA")) = "1" and Trim(rsGestSubCat("COD_SUB_CATEGORIA")) = "1" and Trim(rsGestion("COD_GESTION")) = "6") OR (Trim(rsGestCat("COD_CATEGORIA")) = "1" and Trim(rsGestSubCat("COD_SUB_CATEGORIA")) = "2" and Trim(rsGestion("COD_GESTION")) = "6") OR 	(Trim(rsGestCat("COD_CATEGORIA")) = "4" and Trim(rsGestSubCat("COD_SUB_CATEGORIA")) = "2" and Trim(rsGestion("COD_GESTION")) = "4") Then
										if TraeSiNo(session("perfil_sup")) = "Si" and TraeSiNo(session("perfil_emp")) = "No" Then
											strMuestra="SI"
										Else
											strMuestra="NO"
										End if
									Else
										strMuestra="SI"
									End if

									If strMuestra="SI" Then
									%>
										var newOption = new Option('<%=rsGestion("DESCRIPCION")%>', '<%=rsGestion("COD_GESTION")%>');comboBox.options[comboBox.options.length] = newOption;
									<%
									End If
									rsGestion.movenext
								Loop
							Else
							%>
								var newOption = new Option('SIN GESTION LA CATEGORIA', '0');comboBox.options[comboBox.options.length] = newOption;
							<%
							End if
							CerrarSCG2()
							%>
							break;
						}
						<%
						rsGestSubCat.movenext
					Loop
					rsGestSubCat.close
					set rsGestSubCat=nothing

				Else
					%>
					{
						var newOption = new Option('SELECCIONE', 'X');comboBox.options[comboBox.options.length] = newOption;
						var newOption = new Option('SIN GESTION LA CATEGORIA', '0');comboBox.options[comboBox.options.length] = newOption;
					}
					break;
				<%
				End If
				CerrarSCG1()
				%>
			<%
			  	rsGestCat.movenext
			  	Loop
			  	rsGestCat.close
			  	set rsGestCat=nothing
				CerrarSCG()
			%>
		}
		}

	}

	<% If Trim(strCategoria) = "2" Then %>
		datos.cmbcat.value = 2;
		cargasubcat(datos.cmbcat.value);
	<% End If %>

	<% If Trim(strCategoria) = "3" Then %>
		datos.cmbcat.value = 3;
		cargasubcat(datos.cmbcat.value);

	<% End If %>
	marcar_boxes();

	function AuditarDoc(id_cuota)
	{
		with( document.datos )
		{
			action = "auditar_doc.asp?id_cuota=" + id_cuota ;
			submit();
		}
	}


	function bt_geolocalizacion(direccion){
		window.open('geolocalizacion.asp?direccion='+encodeURIComponent(direccion),"DATOS1","width=610, height=610, scrollbars=no, menubar=no, location=no, resizable=yes")

	}


	function ventanaMas (URL){
		window.open(URL,"DATOS1","width=840, height=400, scrollbars=no, menubar=no, location=no, resizable=yes")
	}

	function TraerGrabacion (strTelefono,strFecIngreso,strHoraIngreso,intIdusuario,strAnexo){
		URL='EscucharGrabacion.asp?strTelefono=' + strTelefono + '&strFecIngreso=' + strFecIngreso + '&strHoraIngreso=' + strHoraIngreso + '&intIdusuario=' + intIdusuario + '&strAnexo=' + strAnexo
		window.open(URL,"DATOS_GRABACION","width=470, height=230, scrollbars=no, menubar=no, location=no, resizable=yes")
	}

	function ventanaGestionesPorDoc (URL){
		window.open(URL,"DATOS2","width=1500, height=600, scrollbars=yes, menubar=no, location=no, resizable=yes")
	}

	function ventanaGestionesFonos (URL){
		window.open(URL,"DATOS2","width=1300, height=600, scrollbars=yes, menubar=no, location=no, resizable=yes")
	}
	
	function ventanaBusqueda (URL){
		window.open(URL,"DATOS3","width=1050, height=700, scrollbars=yes, menubar=no, location=no, resizable=yes")
	}


	function ventana_simulacion_convenio(email, NOMBRE_DEUDOR){
		var concat 						=""
		var rut 						=$('#rut_').val()
		var strCodCliente   		=$('#strCodCliente').val()
		var fecha_generar_documentos 	=$('#fecha_generar_documentos').val()


		$('input[id="CH_ID_CUOTA"]:checked').each(function(){	
			concat 				= concat + ","+$(this).val()
		})

		var cuotas_rut =concat.substring(1, concat.length)
		

		if(concat==""){
			alert("Debe seleccionar al menos un documento ")
			return

		}else{
			NOMBRE_DEUDOR =NOMBRE_DEUDOR.replace("Á","A");
			NOMBRE_DEUDOR =NOMBRE_DEUDOR.replace("É","E");
			NOMBRE_DEUDOR =NOMBRE_DEUDOR.replace("Í","I");
			NOMBRE_DEUDOR =NOMBRE_DEUDOR.replace("Ó","O");
			NOMBRE_DEUDOR =NOMBRE_DEUDOR.replace("Ú","U");
			NOMBRE_DEUDOR =NOMBRE_DEUDOR.replace("Ñ","N");

			NOMBRE_DEUDOR =NOMBRE_DEUDOR.replace("á","a");
			NOMBRE_DEUDOR =NOMBRE_DEUDOR.replace("é","e");
			NOMBRE_DEUDOR =NOMBRE_DEUDOR.replace("í","i");
			NOMBRE_DEUDOR =NOMBRE_DEUDOR.replace("ó","o");
			NOMBRE_DEUDOR =NOMBRE_DEUDOR.replace("ú","u");
			NOMBRE_DEUDOR =NOMBRE_DEUDOR.replace("ñ","n");

			var criterios ="alea="+Math.random()+"&email="+email+"&NOMBRE_DEUDOR="+NOMBRE_DEUDOR+"&cuotas_rut="+cuotas_rut+"&accion_ajax=crea_correo_electronico&rut="+rut+"&strCodCliente="+strCodCliente+"&fecha_generar_documentos="+encodeURIComponent(fecha_generar_documentos)
			$('#ventana_envio_correo').load('FuncionesAjax/envio_correo_ajax.asp', criterios, function(){			
				$('.bt_enviar_correo').hover(function(){
					$(this).css('background-color', '#D2D2D2')
				}, function(){
					$(this).css('background-color', '#E6E6E6')
				})			
				var criterios ="alea="+Math.random()+"&rut="+rut+"&strCodCliente="+strCodCliente+"&cuotas_rut="+cuotas_rut+"&NOMBRE_DEUDOR="+NOMBRE_DEUDOR+"&fecha_generar_documentos="+encodeURIComponent(fecha_generar_documentos)
				$('#crea_archivos').load('pdf.asp', criterios, function(){


					var criterios ="alea="+Math.random()+"&rut="+rut+"&strCodCliente="+strCodCliente+"&cuotas_rut="+cuotas_rut+"&NOMBRE_DEUDOR="+NOMBRE_DEUDOR+"&fecha_generar_documentos="+encodeURIComponent(fecha_generar_documentos)
					$('#crea_archivos_excel').load('crea_excel.asp', criterios, function(){

						$('#txt_observacion_email').blur(function(){
							$('#visual_txt_observacion_email').text($(this).val())
							$('#visual_txt_observacion_email_titulo').text("Observacion:")
							if($(this).val()==""){
								$('#visual_txt_observacion_email_titulo').text("")
							}else{
								$('#visual_txt_observacion_email_titulo').text("Observacion:")
							}
						})

						$("#con_copia").multiselect(); 
						$('#ventana_envio_correo').dialog({
					   		show:"blind", 
					   		hide:"explode",   		       	 
					    	width:700,
					    	height:780 ,
					    	modal:true	
						    	
						});	

					})

				})

			})
				
		}

	}

	function bt_tipo_correo(CORREO_SALIENTE, NOMBRE_DEUDOR){
		var rut 						=$('#rut_').val()
		var strCodCliente   		=$('#strCodCliente').val()
		var concat 						=""
		var txt_observacion_email 		=$('#txt_observacion_email').val()
		var email_enviar 				=$('#email_para').text()
		var adj_pdf 					=$('input[id="adj_pdf"]:checked').val()
		var adj_excel 					=$('input[id="adj_excel"]:checked').val()
		var concat_con_copia 			=""
		var visualiza_formato_correo 	=$('#visualiza_formato_correo').val()
		var COD_CORREO 					=$('#COD_CORREO').val()
		var NOMBRE_DEUDOR 				=$('#NOMBRE_DEUDOR').val()
		var fecha_generar_documentos 	=$('#fecha_generar_documentos').val()
		if(visualiza_formato_correo=="S")
		{
			$('select[id="con_copia"] option:checked').each(function(){
				concat_con_copia =concat_con_copia + '***'+$(this).val()
			})

			var concat_con_copia = concat_con_copia.substring(3, concat_con_copia.length)

			if(adj_pdf==null){
				adj_pdf =""
			}
			if(adj_excel==null){
				adj_excel =""
			}

			$('input[id="CH_ID_CUOTA"]:checked').each(function(){
			
				concat 				= concat + ","+$(this).val()
			})

			if(concat==""){
				alert("Debe seleccionar al menos un documento ")
				return
			}

			txt_observacion_email = txt_observacion_email.replace("'","")
			txt_observacion_email = txt_observacion_email.replace("'","")
			txt_observacion_email = txt_observacion_email.replace("'","")
			txt_observacion_email = txt_observacion_email.replace("'","")
			txt_observacion_email = txt_observacion_email.replace("'","")
			txt_observacion_email = txt_observacion_email.replace("'","")
			
			var cuotas_rut =concat.substring(1, concat.length)

			var criterios ="alea="+Math.random()+"&accion_ajax=envio_plan_pago&email="+email_enviar+"&rut="+rut+"&strCodCliente="+strCodCliente+"&cuotas_rut="+cuotas_rut+"&NOMBRE_DEUDOR="+NOMBRE_DEUDOR+"&txt_observacion_email="+encodeURIComponent(txt_observacion_email)+"&CORREO_SALIENTE="+encodeURIComponent(CORREO_SALIENTE)+"&adj_pdf="+adj_pdf+"&adj_excel="+adj_excel+"&concat_con_copia="+concat_con_copia+"&COD_CORREO="+COD_CORREO+"&fecha_generar_documentos="+fecha_generar_documentos

			$('#envio_correo_plan_pago').load('FuncionesAjax/envio_correo_ajax.asp', criterios, function(){

				var criterios ="alea="+Math.random()+"&accion_ajax=actualiza_fecha_hora"
				$('#refresca_fecha_hora').load('FuncionesAjax/envio_correo_ajax.asp', criterios, function(){})

				setTimeout("$('#ventana_envio_correo').dialog('close'); $('#envio_correo_plan_pago').text('') ",1000)

			})

		}else{
			alert("Debe visualizar email antes de enviar")
			return
		}
	}
	

	function bt_visualiza_correo(cod_correo,rut,NOMBRE_DEUDOR,nombres_usuario,apellido_paterno,correo_electronico,cuotas_rut)
	{
		var txt_observacion_email 	=$('#txt_observacion_email').val() 
		var strCodCliente   		=$('#strCodCliente').val()
		
		$('#visualiza_formato_correo').val("S")
		var criterios ="alea="+Math.random()+"&accion_ajax=visualiza_correo&cod_correo="+cod_correo+"&rut="+rut+"&strCodCliente="+strCodCliente+"&NOMBRE_DEUDOR="+NOMBRE_DEUDOR+"&txt_observacion_email="+txt_observacion_email+"&nombres_usuario="+nombres_usuario+"&apellido_paterno="+apellido_paterno+"&correo_electronico="+correo_electronico+"&cuotas_rut="+cuotas_rut

		$('#visualiza_correo').load('FuncionesAjax/envio_correo_ajax.asp', criterios, function(){
			var NOM_CORREO 			=$("#NOM_CORREO").val()
			var ASUNTO_CORREO 		=$("#ASUNTO_CORREO").val()
			var CORREO_SALIENTE 	=$("#CORREO_SALIENTE").val()
			var CUERPO_CORREO 		=$("#CUERPO_CORREO").val()
			var FIRMA 				=$("#FIRMA").val()
			var COD_CORREO 			=$("#COD_CORREO").val()

			$("#email_de").text(CORREO_SALIENTE)
			$("#email_asunto").text(ASUNTO_CORREO+' Rut: '+rut+' Nombre: '+NOMBRE_DEUDOR)
		})
	}
	
	function ventanaBiblioteca (URL){
	window.open(URL,"INFORMACION","width=1000, height=500, scrollbars=yes, menubar=no, location=no, resizable=yes")
	}
		
</script>