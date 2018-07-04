<SCRIPT LANGUAGE="JavaScript">
//-------------------------------------------------------------------------------
// Propósito: Cambiar el campo actual y validar el RUT cuando el usuario presiona ENTER
// Entradas: 	ObjetoOrigen: El objeto a validar
//		ObjetoDestino: El ObjetoDestino destino
//		Tipo: El Tipo de Seleccion 1: Completa 2: Cursor
// Retorna: false
//-------------------------------------------------------------------------------
function RUTEnter (ObjetoOrigen , ObjetoDestino , Tipo) {
	if( window.event.keyCode == 13 ) {
		if( checkRutField( ObjetoOrigen.value ) ) {
			if (Tipo==1) {
				ObjetoDestino.focus()
				ObjetoDestino.select()
			}
			else	ObjetoDestino.focus()
		}
		else	return false;
	}
}

//-------------------------------------------------------------------------------
// Propósito: Cambiar el campo actual y validar el RUT cuando el usuario presiona 
//		ENTER e irse a una URL si esta OK
// Entradas: 	ObjetoOrigen: El objeto a validar
//		ObjetoDestino: El ObjetoDestino destino
//		Tipo: El Tipo de Seleccion 1: Completa 2: Cursor
// Retorna: false
//-------------------------------------------------------------------------------
function RUTEnterURL (ObjetoOrigen , URL) {
	if( window.event.keyCode == 13 ) {
		if( checkRutField( ObjetoOrigen.value ) ) {
			location.href= URL
		}
		else	return false;
	}
}

//-------------------------------------------------------------------------------
// Propósito: Cambiar el campo actual y validar si es numérico
// Entradas: 	ObjetoOrigen: El objeto a validar
//		ObjetoDestino: El ObjetoDestino destino
//		Tipo: El Tipo de Seleccion 1: Completa 2: Cursor
// Retorna: false
//-------------------------------------------------------------------------------
function NumeroEnteroEnter (ObjetoOrigen , ObjetoDestino , Tipo) {
	if( window.event.keyCode == 13 ) {
		if( esNumero( 'int', ObjetoOrigen.value ) == 'true' ) {
			if (Tipo==1) {
				ObjetoDestino.focus()
				ObjetoDestino.select()
			}
			else	ObjetoDestino.focus()
		}
		else	alert('El Numero ha sido mal ingresado')
	}
}

</SCRIPT>
