<SCRIPT LANGUAGE="JavaScript">

function chkFecha(f) {
  str = f.value
  
  if (str.length<10){
  	alert("Error - Ingresó una fecha no válida");
  	f.focus();
  	f.select();
  }
  
  if ( !formatoFecha(str) ) {
    alert("Debe indicar la Fecha en formato DD/MM/AAAA. Ejemplo: 'Para 20 de Diciembre de 1982 se debe ingresar 20/12/1982'");
    f.select()
    f.focus()
    return false
  }

  // validacion de la fecha
  if ( !validarFecha(str) ) {
    // Los mensajes de error están dentro de validarFecha.
    f.select()
    f.focus()
    return false
  }

  return true
}

</SCRIPT>