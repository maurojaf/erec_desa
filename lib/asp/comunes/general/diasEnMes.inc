
<SCRIPT LANGUAGE="JavaScript">

function diasEnMes (month, year)
{
  if (month == 1 || month == 3 || month == 5 || month == 7 || month == 8 || month == 10 || month == 12)
    return 31;
  else if (month == 2)
    // February has 29 days in any year evenly divisible by four,
      // EXCEPT for centurial years which are not also divisible by 400.
      return (  ((year % 4 == 0) && ( (!(year % 100 == 0)) || (year % 400 == 0) ) ) ? 29 : 28 );
  else if (month == 4 || month == 6 || month == 9 || month == 11)
    return 30;
  // En caso contrario:
  alert("diasEnMes: Mes inválido");
  return -1;
}

</SCRIPT>