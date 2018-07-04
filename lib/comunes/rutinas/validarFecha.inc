<SCRIPT LANGUAGE="JavaScript">


function validarFecha(str_fecha){

   var sl1=str_fecha.indexOf("/")
  var sl2=str_fecha.lastIndexOf("/")
  var inday = parseFloat(str_fecha.substring(0,sl1))
  var inmonth = parseFloat(str_fecha.substring(sl1+1,sl2))
  var inyear = parseFloat(str_fecha.substring(sl2+1,str_fecha.length))

  //alert("day:" + inday + ", mes:" + inmonth + ", agno: " + inyear)

  if (inmonth < 1 || inmonth > 12) {
    alert("Mes inválido en la fecha");
    return false;
  }
  if (inday < 1 || inday > diasEnMes(inmonth, inyear)) {
    alert("Día inválido en la fecha");
    return false;
  }

  return true
}

</SCRIPT>