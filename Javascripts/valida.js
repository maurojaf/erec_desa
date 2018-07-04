function validaDV( crut,cajarut ){
  largo = crut.length;
  if ( largo >= 8 )
     rut = crut.substring(0, largo - 1);
  else
     rut = crut.charAt(0);
	 
  dv = crut.charAt(largo-1);
  if ( rut == null || dv == null )
      return 0;

  var dvr = '0';
  suma = 0;
  mul  = 2;
  for (i= rut.length -1 ; i >= 0; i--){
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
  else{
    dvi = 11-res;
    dvr = dvi + "";
  }

  if ( dvr != dv.toLowerCase() ){
    alert("RUT NO VALIDO"); // El rut es invalido
    switch(cajarut) {
    case '1' : 
      window.document.datos.rut.value="";
      window.document.datos.rut.focus();
      break;
    }
    return false;
  }
  return true;
}


function valida_rut(texto,cajarut){
if (texto!=''){
  var tmpstr = "";
  for ( i=0; i < texto.length ; i++ )
    if ( texto.charAt(i) != ' ' && texto.charAt(i) != '.' && texto.charAt(i) != '-' )
      tmpstr = tmpstr + texto.charAt(i);
      
  texto = tmpstr;
  largo = texto.length;

  if ( largo < 8 ){
    alert("RUT NO VALIDO");
	switch(cajarut) {
    case '1' : 
      window.document.datos.rut.value="";
      window.document.datos.rut.focus();
      break;
    }
	 
    return false;
  }

  var invertido = "";
  for ( i=(largo-1),j=0; i>=0; i--,j++ )
    invertido = invertido + texto.charAt(i);

  var dtexto = "";
  dtexto = dtexto + invertido.charAt(0);
  dtexto = dtexto + '-';
  cnt = 0;
  for ( i=1,j=2; i<largo; i++,j++ ){    
    if ( cnt == 3 ){
      //dtexto = dtexto + '.';
      j++;
      dtexto = dtexto + invertido.charAt(i);
      cnt = 1;
    }else{ 
      dtexto = dtexto + invertido.charAt(i);
      cnt++;
    }
  }

  invertido = "";
  for ( i=(dtexto.length-1),j=0; i>=0; i--,j++ )
    invertido = invertido + dtexto.charAt(i);

  switch(cajarut) {
    case '1' : 
	  window.document.datos.rut.value = invertido;
      break;
    }
  
  if ( validaDV(texto,cajarut) )
    return true;
    
  // El rut es incorrecto
  return false;
}
}



function valida_fecha(campo,caja){

if(campo.length==10){
	largo=campo.length;
	fecha=campo;
	
	dia=fecha.substr(0,2);
	mes=fecha.substr(3,2);
	ano=fecha.substr(6,4);
	
	if((!isNaN(dia))&&(!isNaN(mes))&&(!isNaN(ano))){
		bisiesto=false;
		if ((ano % 4)==0){
			bisiesto=true;
		} 
		if(mes>=1&&mes<=12){
			if(mes==1||mes==3||mes==5||mes==7||mes==8||mes==10||mes==12){
				if(dia>0&&dia<32){
					return true;
				}
				else{
					alert('INGRESE FECHA CORRECTA');
					caja.value='';
					caja.focus();
					return false;
				}	
			}
			if(mes==4||mes==6||mes==9||mes==11){
				if(dia>0&&dia<31){
					return true;
				}
				else{
					alert('INGRESE FECHA CORRECTA');
					caja.value='';
					caja.focus();
					return false;
				}
			}
			if(mes==2&&bisiesto){
				if(dia>0&&dia<30){
					return true;
				}
				else{
					alert('INGRESE FECHA CORRECTA');
					caja.value='';
					caja.focus();
					return false;
				}
			}else{
				if(dia>0&&dia<29){
					return true;
				}
				else{
					alert('INGRESE FECHA CORRECTA');
					caja.value='';
					caja.focus();
					return false;
				}
			}
		}else{
			alert('INGRESE FECHA CORRECTA');
			caja.value='';
			caja.focus();
			return false;
		}
	}else{
		alert('INGRESE FECHA CORRECTA');
		caja.value='';
		caja.focus();
		return false;
	}
	
}else{
	alert('INGRESE FECHA CORRECTA');
	caja.value='';
	caja.focus();
	return false;
}
}

function no_numeros()
{
if ((event.keyCode > 32 && event.keyCode < 48) || (event.keyCode > 45 && event.keyCode < 65) || (event.keyCode > 90 && event.keyCode < 97)) event.returnValue = false;
}

function no_letras()
{if (event.keyCode < 45 || event.keyCode > 57) event.returnValue = false;}


function imprime(){
window.print();
}