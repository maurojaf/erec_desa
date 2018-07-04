
<SCRIPT LANGUAGE="JavaScript">

function sonDigitos(str) {
  var l, car

  l = str.length
  if ( l<1 )
    return false

  for ( i=0; i<l; i++) {
    car = str.substring(i,i+1)
    if ( "0" <= car &&  car <= "9" )
      continue
    else
      return false
  }
  return true
}

</SCRIPT>