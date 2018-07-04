<SCRIPT LANGUAGE=JavaScript>

<% If sstrMensajeBorrar <> "" Then Response.Write "alert( """ & sstrMensajeBorrar & """ ) " %>
<% If MensajeNuevo <> "" Then Response.Write "alert( """ & MensajeNuevo & """ ) " %>

document.body.topMargin  = 0;
document.body.leftMargin = 0;

function BAceptar( Form ) {
  Form.submit()
}


function IrPagina( Accion, Formulario ) {
    cambiarsintPagina = false;
    
    
    if( Accion == "Avanzar" ) {
        if ( <%= sintPagina %> < <%= sintTotalPaginas %> ) {
            Formulario.sintPagina.value = parseInt( Formulario.sintPagina.value ) + 1
            cambiarsintPagina = true
        }
    }
    else if( Accion == "Retroceder" ) {
        if ( <%= sintPagina %> > 1 ) {
            Formulario.sintPagina.value = parseInt( Formulario.sintPagina.value ) - 1
            cambiarsintPagina = true
        }
    }
    else if( Accion == "Primero" ) {
        if ( <%= sintPagina %> > 1 ) {
            Formulario.sintPagina.value = 1
            cambiarsintPagina = true
        }
    }
    else if( Accion == "Ultimo" ) {
        if ( <%= sintPagina %> > 0 ) {
            Formulario.sintPagina.value = <%= sintTotalPaginas %>
            cambiarsintPagina = true
            BAceptar( Formulario )
        }
    }
    else  {
    	//alert(Accion)
        Formulario.sintPagina.value = Accion
        cambiarsintPagina = true
    }

    if ( cambiarsintPagina && ( ChequeaCambio() == true ) ) {
        if (confirm("¿Está seguro de ir cambiar la página sin Actualizar los cambios?")) {
            DeshacerCambios()
            BAceptar( Formulario )
        }
    }
    else if ( cambiarsintPagina )
        BAceptar( Formulario )
}

//function IrTerminar
//    Acción del botón terminar
//
function IrTerminar( sintPaginaTerminar ) {
    if (ChequeaCambio() == true) {
      if (confirm("¿Está seguro de salir sin Actualizar los cambios?")) {
        DeshacerCambios()
        self.location.href = sintPaginaTerminar
      }
    }
    else {
	self.location.href = sintPaginaTerminar
    }
}

//function IrNuevo
//    Acción del botón nuevo
//
function IrNuevo(sintPaginaForm) {
  self.location.href = sintPaginaForm + '?sintNuevo=1&FormMode=Nuevo'
}

function IrNuevoId(sintPaginaForm) {
  self.location.href = sintPaginaForm + '&sintNuevo=1&FormMode=Nuevo'
}

// Función ChequeaCambio()
//    Chequea si cambió algún registro en toda la página (a modificar o nuevos)
// Retorna:
//    true o false
function ChequeaCambio(){
  cambio = false
  <% For i = intIndiceFilasIni to intIndiceFilasFin %>
      if (document.mantenedorForm.borrar<%= i %>.checked == true)
        cambio = true
  <% Next %>

  return cambio
}

function DeshacerCambios(){
  // Deshace los cambios
  <% For i = intIndiceFilasIni to intIndiceFilasFin %>
    document.mantenedorForm.borrar<%= i %>.checked = false
  <% Next %>
}


function Buscar(sintPaginaForm) {
    self.location.href = sintPaginaForm + '?sintNuevo=1&FormMode=Nuevo'
}

</SCRIPT>