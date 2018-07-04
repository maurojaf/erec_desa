/*
Objetivos    : Validar cualquier tipo de cadena
Parametros   :	Item : Se entrega el objeto 
                CadenaOk : Cadena de caracteres validos para la funcion 
Retorna      : Retorna true o false
*/
function ValString()
//item, checkOK
{
var oCampo = ValString.arguments[0];
var sCadenaOk = ValString.arguments[1];
var iCam, iCad ;	
var lSw = false; 
var lRetorno = true;

	if( oCampo.value.length == 0 )
	{ lRetorno = false; }
	for( iCam = 0; iCam < oCampo.value.length; iCam++ )
	{	sSw = false
		for( iCad = 0; iCad < sCadenaOk.length; iCad++ )
		{	if( oCampo.value.charAt(iCam) == sCadenaOk.charAt(iCad) )
			{ sSw = true; break; }
		}
		if( sSw == false )
		{	lRetorno = false; iCam = oCampo.value.length; }
	}
	return lRetorno
}
