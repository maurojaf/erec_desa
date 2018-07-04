

<!--#include file="esNumero.inc"-->
<!--#include file="chkFecha.inc"-->
<!--#include file="formatoFecha.inc"-->
<!--#include file="validarFecha.inc"-->
<!--#include file="diasEnMes.inc"-->
<!--#include file="sonDigitos.inc"-->
<!--#include file="RutinasEnter.inc"-->

<SCRIPT LANGUAGE="JavaScript">


// Función: chkHora(f)
// Descripción: Chequea el formato de la hora. Comprueba que el formato
//             sea hh:mm donde y que la hora sea válida (ora 1 a 24 min 0 a 59)
// Argumentos:  f:  objeto del formulario de introducción de la hora
// Retorna: true o false
//
function chkHora(f) {
  str = f.value
  if ( !formatoHora(str) ) {
    alert("Debe indicar la Hora en formato HH:MM. Entre 00:00 y 23:59");
    f.select()
    f.focus()
    return false
  }

  // validacion de la fecha
  if ( !validarHora(str) ) {
    // Los mensajes de error están dentro de validarHora
    f.select()
    f.focus()
    return false
  }

  return true
}


// Función: formatoHora()
// Descripción: Comprueba que el formato del string sea hh:mm, donde
//               x es dígito decimal.
// Argumentos:  str:  string
// Retorna: true o false
//

function formatoHora(str) {
  var sl1, ui, hhstr, mmstr;

  // El formato debe ser h:mm, hh:mm,
  // Las posiciones son a partir de 0
  if (str.length < 4 &&  str.length > 5)    // el tamagno es fijo de 4 o 5
    return false

  sl1=str.indexOf(":")
  if (sl1 < 1 && sl1 > 2 )    // el primer slash debe estar en la 1 o 2
    return false


  hhstr = str.substring(0,sl1)
  mmstr = str.substring(sl1+1,str.length)
  if ( !sonDigitos(hhstr) || !sonDigitos(mmstr) )
    return false
  return true
}


// Función: validarHora(f)
// Descripción: Valida que el campo sea una hora correcta.
// Argumentos:  str_hora: string con la hora, con el formato hh:mm
// Retorna: true o false
//
function validarHora(str_hora){

  // Inicialmente habia puesto parseInt para transformar a numeros
  // enteros, pero con 08 y 09 en el dia transformaba a 0 en
  // explorer 4.0. Dionel.


  var sl1=str_hora.indexOf(":")
  var inhora = parseFloat(str_hora.substring(0,sl1))
  var inmin= parseFloat(str_hora.substring(sl1+1,str_hora.length))


  if (inhora< 0 || inhora> 23  || inmin <0 || inmin > 59 ) {
    alert("Debe indicar la Hora en formato HH:MM. Entre 00:00 y 23:59");
    return false;
  }

  return true
}


function checkCDV( dvr )
{
    dv = dvr + ""
    if ( dv != '0' && dv != '1' && dv != '2' && dv != '3' && dv != '4' && dv != '5' && dv != '6' && dv != '7' && dv != '8' && dv != '9' && dv != 'k'  && dv != 'K')
    {
        alert("Debe ingresar un digito verificador valido");
        document.forms[0].RUT.focus();
        document.forms[0].RUT.select();
        return false;
    }
    return true;
}


function checkDV( crut )
{
    var tmpstr = "";
    largo = crut.length;
    if ( largo < 2 )
    {
        alert("Debe ingresar el rut completo");
        document.forms[0].RUT.focus();
        document.forms[0].RUT.select();
        return false;
    }



    for ( i=0; i < crut.length ; i++ )
                if ( crut.charAt(i) != ' ' && crut.charAt(i) != '.' && crut.charAt(i) != '-' )
                {
                tmpstr = tmpstr + crut.charAt(i);
                }
            rut = tmpstr;
    crut=tmpstr;
    largo = crut.length;



    if ( largo > 2 )
        rut = crut.substring(0, largo - 1);
    else
        rut = crut.charAt(0);

    dv = crut.charAt(largo-1);




    if ( rut == null || dv == null )
            return 0;


    var dvr = '0';
    suma = 0;
    mul  = 2;

    for (i= rut.length-1 ; i >= 0; i--)
    {
        suma = suma + rut.charAt(i) * mul;
        if (mul == 7)
            mul = 2;
        else
            mul++;
    }


    res = suma % 11;
    if (res==1)
        dvr = 'k';
    else if (res==0)
        dvr = '0';
    else
    {
        dvi = 11-res;
        dvr = dvi + "";
    }


    if ( dvr != dv.toLowerCase() )
    {
        alert("El rut es incorrecto "+ dvr );
        document.forms[0].RUT.focus();
            document.forms[0].RUT.select();
        return false;
    }

    return true;
}

function checkRutField(texto)

{

    var tmpstr = "";

    for ( i=0; i < texto.length ; i++ )
        if ( texto.charAt(i) != ' ' && texto.charAt(i) != '.' && texto.charAt(i) != '-' )
            tmpstr = tmpstr + texto.charAt(i);
    texto = tmpstr;
    largo = texto.length;


    if ( largo < 2 )
    {
        alert("Debe ingresar el rut completo");
        document.forms[0].RUT.focus();
        document.forms[0].RUT.select();
        return false;
    }


    for (i=0; i < largo ; i++ )

    {
        if ( texto.charAt(i) !="0" && texto.charAt(i) != "1" && texto.charAt(i) !="2" && texto.charAt(i) != "3" && texto.charAt(i) != "4" && texto.charAt(i) !="5" && texto.charAt(i) != "6" && texto.charAt(i) != "7" && texto.charAt(i) !="8" && texto.charAt(i) != "9" && texto.charAt(i) !="k" && texto.charAt(i) != "K" )
        {
            alert("El valor ingresado no corresponde a un R.U.T valido");
            document.forms[0].RUT.focus();
            document.forms[0].RUT.select();
            return false;
        }
    }



    var invertido = "";


    for ( i=(largo-1),j=0; i>=0; i--,j++ )
        invertido = invertido + texto.charAt(i);


    var dtexto = "";
    dtexto = dtexto + invertido.charAt(0);
    dtexto = dtexto + '-';
    cnt = 0;

    for ( i=1,j=2; i<largo; i++,j++ )
    {
        if ( cnt == 3 )
        {
            dtexto = dtexto + '.';
            j++;
            dtexto = dtexto + invertido.charAt(i);
            cnt = 1;
        }
        else
        {
            dtexto = dtexto + invertido.charAt(i);
            cnt++;
        }
    }

    invertido = "";

    for ( i=(dtexto.length-1),j=0; i>=0; i--,j++ )
        invertido = invertido + dtexto.charAt(i);

    document.forms[0].RUT.value = invertido;

    if ( checkDV(texto) )
        return true;

    return false;

}

function limpiaRUT( crut )
{
    var tmpstr = "";
    largo = crut.length;

    for ( i=0; i < crut.length ; i++ )
                if ( crut.charAt(i) != ' ' && crut.charAt(i) != '.' && crut.charAt(i) != '-' )
                {
                tmpstr = tmpstr + crut.charAt(i);
                }
            rut = tmpstr;
    return tmpstr;
}

function cambiaLiquido( bruto ){
	document.forms[0].RETENCION.value =  (bruto * (document.forms[0].PORC_RET.value / 100 ));
	document.forms[0].LIQUIDO.value = bruto - document.forms[0].RETENCION.value;
	document.forms[0].NUMCERTIF.focus();
	return false
}

function PoneValor(intValor){
	if (intValor==1){
		document.forms[0].PORC_DIETA.value='10';
		return false
	}
	if (intValor==2){
		document.forms[0].PORC_DIETA.value='0';
		return false
	}
}

//Funcion que oculta  campos



function ValidaTipoRut (rut, tipo)
{
if (tipo==1){
	checkRutField(rut);
	}
else 
	{
	return true;	
	
	}


}


</SCRIPT>
