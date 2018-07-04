  <!--Tabla 2 filas y 3 columnas: botones de navegación y contador de páginas-->
  <!-- NO MODIFICAR -->
  <TABLE BORDER=0 CELLSPACING=0 CELLPADDING=0 WIDTH="630" HEIGHT="36" BGCOLOR=<%=strBgColDet%> ALIGN="CENTER" >
  <TR bgColor=<%=strBgColDet%>>
  <TD WIDTH="5%" ALIGN=LEFT><INPUT TYPE=BUTTON NAME="Primero" VALUE="  &lt;&lt;  " onClick="IrPagina( 'Primero', this.form )"></TD>
  <TD WIDTH="5%" ALIGN=LEFT><INPUT TYPE=BUTTON NAME="Retroceder" VALUE="  &lt;  " onClick="IrPagina( 'Retroceder', this.form )"></TD>
  <TD WIDTH="100%" ALIGN=center>
  <FONT FACE="verdana, Sans-Serif" Size=2><b>Página <%= Pagina %> de <%= TotalPaginas %></b></FONT>
  </TD>
  <TD WIDTH="25%" ALIGN=right>&nbsp</TD>
  <TD WIDTH="25%" ALIGN=right><INPUT TYPE=BUTTON NAME="Avanzar" VALUE="  &gt;  " onClick="IrPagina( 'Avanzar', this.form )"></TD>
  <TD WIDTH="25%" ALIGN=right><INPUT TYPE=BUTTON NAME="Ultimo" VALUE="  &gt;&gt;  " onClick="IrPagina( 'Ultimo', this.form )"></TD>
  </TR>
  </TABLE>
