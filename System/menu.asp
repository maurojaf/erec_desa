<!--#include file="arch_utils.asp"-->
<link href="../css/style_menu.css" rel="stylesheet" type="text/css" />
<%
   Response.CodePage =65001
   Response.charset  ="utf-8"

   IdUsuario = session("session_idusuario")
   strCodCliente = session("ses_codcli")
   
   AbrirSCG1()
   
	strSql="EXEC proc_Parametros_Tabla_Cliente '""','"&TRIM(strCodCliente)&"'"

	set rsCLI=Conn1.execute(strSql)
	if not rsCLI.eof then
		intUsaDiscador 	= rsCLI("USA_DISCADOR")
		intTipoAgendamiento = rsCLI("TIPO_AGENDAMIENTO")
	end if
	rsCLI.close
	set rsCLI=nothing
	
   CerrarSCG1()
   
   'Response.write "<br>intUsaDiscador=" & intUsaDiscador

   AbrirSCG()

      strSql =" SELECT PERFIL_ADM, PERFIL_EMP, PERFIL_SUP, PERFIL_SUP_CAJA, PERFIL_COB, PERFIL_CAJA, PERFIL_BACK,PuedoEscucharGrabaciones "
      strSql = strSql & " FROM USUARIO "
      strSql = strSql & " WHERE ID_USUARIO = " & IdUsuario
      set rsPerfilUsu = Conn.execute(strSql)

      strPerfilAdm      = rsPerfilUsu("PERFIL_ADM")
      strPerfilCliente  = rsPerfilUsu("PERFIL_EMP")
      strPerfilSup      = rsPerfilUsu("PERFIL_SUP")
      strPerfilCob      = rsPerfilUsu("PERFIL_COB")
	  strPerfilSupCaja      = rsPerfilUsu("PERFIL_SUP_CAJA")
      strPerfilCaja     = rsPerfilUsu("PERFIL_CAJA")
      strPerfilBack     = rsPerfilUsu("PERFIL_BACK")
      strPuedoEscucharGrabaciones = rsPerfilUsu("PuedoEscucharGrabaciones")
      strModulosVisibles      = "(1,2,3,8)"
      strSubModulosVisibles   = "(0)"

		if strPerfilCliente = "Verdadero" then
			intmodulosagendamiento = "0"
		elseif intTipoAgendamiento=0 then
			intmodulosagendamiento = "53"
		elseif intTipoAgendamiento=1 then
			intmodulosagendamiento = "52,53"
		elseif intTipoAgendamiento=2 and strPerfilSup = "Verdadero" then
			intmodulosagendamiento = "52,53"
		elseif intTipoAgendamiento=2 then
			intmodulosagendamiento = "52"
		else
			intmodulosagendamiento = "0"
		End If

      '--Perfil Cliente Supervisor--'

      If strPerfilCliente="Verdadero" and strPerfilSup = "Verdadero" Then

         strModulosVisibles      = "(1,2,3,7,8)"
         strSubModulosVisibles   = "(1,"&intmodulosagendamiento&",4,5,6,7,9,10,11,13,15,16,17,18,19,30,37,38,39,44,49)"

      '--Perfil Cliente Cobrador--'

      ElseIf strPerfilCliente="Verdadero" and strPerfilCob = "Verdadero" Then

         strModulosVisibles      = "(1,2,3,8)"
         strSubModulosVisibles   = "(1,"&intmodulosagendamiento&",4,5,7,9,10,11,13,15,16,17,18,37,38)"

      '--Perfil Supervisor - Caja--'
      ElseIf strPerfilSupCaja = "Verdadero"  Then

         strModulosVisibles      = "(1,2,3,4,5,8)"
         strSubModulosVisibles   = "(1,"&intmodulosagendamiento&",4,5,7,8,9,10,11,12,13,15,16,17,18,19,21,22,23,24,37,38,39,40,41,42)"
		 
      '--Perfil Supervisor--'
      ElseIf strPerfilSup = "Verdadero" Then

         strModulosVisibles      = "(1,2,3,7,8,10)"
         '' se restringue el acceso al form 
         if strPuedoEscucharGrabaciones = "Verdadero" then
                strSubModulosVisibles   = "(1,"&intmodulosagendamiento&",4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20,30,31,32,33,34,35,36,37,38,39,44,45,46,47,48,49,54,55)"
         else
                strSubModulosVisibles   = "(1,"&intmodulosagendamiento&",4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,30,31,32,36,37,38,39,44,45,46,48,49,54,55)"
         end if 

      '--Perfil Cobrador--'

      ElseIf strPerfilCob = "Verdadero" Then

         strModulosVisibles      = "(1,2,3,8)"
         strSubModulosVisibles   = "(1,"&intmodulosagendamiento&",4,5,7,8,9,10,11,12,13,15,16,17,18,19,37,38)"

      '--Perfil Backoffice--'

      ElseIf strPerfilBack = "Verdadero" Then

         strModulosVisibles      = "(1,2,3,8)"

         if strPuedoEscucharGrabaciones   = "Verdadero" then
                strSubModulosVisibles   = "(1,6,7,8,9,10,11,12,13,18,19,20,33,34,37,38,39)"
         else
                strSubModulosVisibles   = "(1,6,7,8,9,10,11,12,13,18,19,37,38)"
         end if

	  	'--Perfil CLIENTE--'
		ElseIf strPerfilCliente = "Verdadero" Then

          strModulosVisibles      = "(1,2,3,8)"
         strSubModulosVisibles   = "(1,"&intmodulosagendamiento&",4,5,7,9,10,11,13,15,16,17,18,37,38)"
      End If
		

	  strSql=" SELECT MP.COD_MODULO,MP.COD_SUB_MODULO"
	  strSql = strSql & " FROM MODULO_PERFIL MP INNER JOIN MODULO_PERFIL_USUARIO MPU ON MP.PER_CODIGO=MPU.PER_CODIGO"
	  strSql = strSql & " 	WHERE ID_USUARIO = " & IdUsuario
	  set rsModUsu= Conn.execute(strSql)
	  
	  'Response.write "strSql=" & strSql
	  'Response.end
	  
		  if not rsModUsu.eof then		

			strModulosVisibles=""
			strSubModulosVisibles=""
		  
			do until rsModUsu.eof
			
			intCodModulo = rsModUsu("COD_MODULO")
			intCodSubModulo = rsModUsu("COD_SUB_MODULO")
			
			strModulosVisibles = strModulosVisibles & "," & intCodModulo
			
			strSubModulosVisibles = strSubModulosVisibles & "," & intCodSubModulo
			
			'Response.write "strSql=" & strModulosVisibles
			
			rsModUsu.movenext
			loop
			
			strModulosVisibles = "(" & Mid(strModulosVisibles,2,Len(strModulosVisibles)) & ")"
			
			strSubModulosVisibles = "(" & Mid(strSubModulosVisibles,2,Len(strSubModulosVisibles)) & ")"
			
			'Response.write "strSql=" & strModulosVisibles
			'Response.write "strSql=" & strSubModulosVisibles
			
		  end if
        
%>
<div id='cssmenu'>
<ul>
<%

  strSql=" SELECT NOM_MODULO,COD_MODULO,ID_MODULO"
  strSql = strSql & " FROM MODULOS_TIPO_CATEGORIA"
  strSql = strSql & " WHERE '" & strPerfilAdm & "' = 'Verdadero' OR COD_MODULO IN " & strModulosVisibles
  strSql = strSql & " ORDER BY ORDEN"
  set rsMod= Conn.execute(strSql)
  
   if not rsMod.eof then
      do until rsMod.eof

         intCodModulo      = rsMod("COD_MODULO")
         strNombreModulo   = rsMod("NOM_MODULO")
         strIdModulo       = rsMod("ID_MODULO")

         %>

         <li class='has-sub'><a href='#'><span><%=strNombreModulo%></span></a>
         <ul>
         <%
            strSql=" SELECT COD_MODULO,NOM_SUB_MODULO,ASP_SUB_MODULO,PARAMETROS"
            strSql = strSql & " FROM MODULOS_TIPO_SUBCATEGORIA"
            strSql = strSql & " WHERE COD_MODULO = " & intCodModulo & " AND ('" & strPerfilAdm & "' = 'Verdadero' or  COD_SUB_MODULO IN " & strSubModulosVisibles & ")"
            strSql = strSql & " ORDER BY ORDEN"
            set rsSubMod= Conn.execute(strSql)
      
            if not rsSubMod.eof then
               
               do until rsSubMod.eof

                  strNombreSubModulo   = Trim(rsSubMod("NOM_SUB_MODULO"))
                  strNombreAspModulo   = Trim(rsSubMod("ASP_SUB_MODULO"))
                  strParametros        = Trim(rsSubMod("PARAMETROS"))
                  %>            
                  <li class='has-sub'><a href='<%=strNombreAspModulo%>'><span><%=strNombreSubModulo%></span></a></li>
                  <%
               rsSubMod.movenext
               loop
            end if %>               
         </ul>
         </li> 

<%
      rsMod.movenext
      loop

   end if

%>
</ul>
</div>


<br>
<br>
<br>