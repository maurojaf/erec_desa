<%
  '------------  Cambio: Guarda en una sesion la ultima pagina accedida ---------------------'
  If sintPagina = "" Then _
  	sintPagina = Session("sintPagina" & Request.ServerVariables("PATH_TRANSLATED"))
  '------------------------------------------------------------------------------------------
  If Trim(Request( "sintNumFilasPagina" )) = "" Then
  	sintNumFilasPagina = 25
  Else
  	sintNumFilasPagina = CInt(Request( "sintNumFilasPagina" ))
  End If
  'response.write sintNumFilasPagina
  ' Calcula el total de páginas
  sintTotalPaginas = Int( CDbl(sintTotalRegistros)/CDbl(sintNumFilasPagina) + 1 )
  ' Chequea condiciones de borde para sintPaginamiento'
  If CInt( sintPagina ) > sintTotalPaginas Then sintPagina = sintTotalPaginas
  If sintPagina < 1 Then sintPagina = 1
  intIndiceFilasIni = sintNumFilasPagina * (sintPagina - 1)
  intIndiceFilasFin = general_Minimo(Cdbl(sintTotalRegistros), Cdbl(sintNumFilasPagina*sintPagina-1))

 '------------  Cambio: Guarda en una sesion la ultima pagina accedida ---------------------'
  Session("sintPagina" & Request.ServerVariables("PATH_TRANSLATED")) = sintPagina
 '------------------------------------------------------------------------------------------
%>