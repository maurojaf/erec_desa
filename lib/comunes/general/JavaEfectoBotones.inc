<STYLE>
  .bigChange {color:<%=strBgColTitG%>; font-weight:bolder; font-size:175%; letter-spacing:4px;  text-transform: uppercase; background:#000080}
  .start {color:#FFFFFF; background:<%=strBgColTitG%>} 
</STYLE>

<SCRIPT LANGUAGE=JAVASCRIPT>
  function highlightButton(s) {
    if ("INPUT"==event.srcElement.tagName)
      event.srcElement.className=s
  }

function entrar_correo()
{
 var accion="Procesa_Login.asp";
 window.document.Cuerpo.action=accion;
 window.document.Cuerpo.submit();
 }

</SCRIPT>
