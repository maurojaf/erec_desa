<SCRIPT LANGUAGE="JavaScript">


function formatoFecha(str) {
  var sl1, sl2, ui, ddstr, mmstr, aaaastr;

  // El formato debe ser d/m/aaaa, d/mm/aaaa, dd/m/aaaa, dd/mm/aaaa,
  // Las posiciones son a partir de 0
  if (str.length < 8 &&  str.length > 10)    // el tamagno es fijo de 8, 9 o 10
    return false


  sl1=str.indexOf("/")
  if (sl1 < 1 && sl1 > 2 )    // el primer slash debe estar en la 1 o 2
    return false

  sl2=str.lastIndexOf("/")
  if (sl2 < 3 &&  sl2 > 5)    // el último slash debe estar en la 3, 4 o 5
    return false

  ddstr = str.substring(0,sl1)
  mmstr = str.substring(sl1+1,sl2)
  aaaastr = str.substring(sl2+1,str.length)

  if ( !sonDigitos(ddstr) || !sonDigitos(mmstr) || !sonDigitos(aaaastr) )
    return false

  return true
}

</SCRIPT>