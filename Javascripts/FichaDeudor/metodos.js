$(function() {

	$( "#DialogHistorial" ).dialog({
		autoOpen: false,
		width: 640,
		height: 480,
		modal: true,
		resizable: false
	});
	
	$( ".ButtonHistorial" ).click(function() {
	
		var args = "IdCampo=" + $(this).attr("idCampo") + 
				"&RutDeudor=" + $("#RutDeudor").val() + 
				"&CodigoCliente=" + $("#CodigoCliente").val();
		
		$.post("FuncionesAjax/FichaDeudor/HistorialCambiosCampo.asp", 
			args, 
			function(data) {
				$("#DialogHistorial").html(data);
			});
	
		$.post("FuncionesAjax/FichaDeudor/GetCampo.asp", 
			"IdCampo=" + $(this).attr("idCampo"), 
			function(campo) {
				$("#DialogHistorial").dialog("option", "title", "Historial de cambios para \"" + campo.Nombre + "\"");
			}, "json");
	
		$( "#DialogHistorial" ).dialog( "open" );
	});
	
	$( "#DialogObservacion" ).dialog({
		autoOpen: false,
		width: 340,
		height: 290,
		modal: true,
		resizable: false
	});
	
	$( "#DialogNoCambios" ).dialog({
		autoOpen: false,
		width: 360,
		height: 120,
		modal: true,
		resizable: false
	});
	
	$( "#DialogCambiosGuardados" ).dialog({
		autoOpen: false,
		width: 360,
		height: 120,
		modal: true,
		resizable: false,
		closeOnEscape: false,
		beforeClose: function (event, ui) { return false; },
		dialogClass: "noclose"
	});

	$("#Menu")
		.find("button")
		.each(function(){
		
			$(this).click(function() {
			
				var elementId = "#Block" + $(this).val();
			
				if( $(elementId).hasClass( "Visible" ) ) {
					
					$(this)
						.addClass("Boton")
						.removeClass("BotonHover");
					
					HideElement(elementId);
					
				} else {
				
					$(this)
						.addClass("BotonHover")
						.removeClass("Boton");
				
					ShowElement(elementId);
					
				}
				
				if ($(elementId).attr("id") != undefined) {
				
					var allHidden = true;
					
					$("div[id^=Block]")
						.not("BlockAllHidden")
						.each(function() {
						
						allHidden = allHidden && $(this).hasClass( "NoVisible" );
						
					});
					
					if (allHidden) {
					
						ShowElement("#BlockAllHidden");
					
					}
					else {
					
						HideElement("#BlockAllHidden");
					
					}
				
				}
			});	
		
		});
	
	$("#ButtonDeudor")
		.addClass("BotonHover")
		.removeClass("Boton");
		
	$("#ButtonContabilidad")
		.addClass("BotonHover")
		.removeClass("Boton");	
		
	$(document)
		.find("button")
		.hover(function() {
		
			$(this)
				.addClass("BotonHover")
				.removeClass("Boton");
		
		},
		function() {
		
			var elementId = "#Block" + $(this).val();
			
			if( !$(elementId)
					.hasClass( "Visible" ) ) {
		
				$(this)
					.addClass("Boton")
					.removeClass("BotonHover");
				
			}
		
		});

	$('#TxaObservacion').keyup(function (e) {
	
		var maxLength = 300;
	
		var text = $(this).val();

		var textLength = text.length;

		if (text.length > maxLength) {

		   $(this).val(text.substring(0, (maxLength)));
		   
		   e.preventDefault();
		   
		   return;
		}

	});
	
	$("img").tooltip();
	
	$("label").tooltip();
		
});

function ShowElement(elementId) {
	$( elementId )
		.addClass( "Visible" )
		.removeClass( "NoVisible" );
}

function HideElement(elementId) {
	$( elementId )
		.removeClass( "Visible" )
		.addClass( "NoVisible" );
}