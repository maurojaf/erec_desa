<html lang="es">
<head>
    <meta http-equiv="X-UA-Compatible" content="IE=edge,chrome=1">
    <meta charset="utf-8">
    <title>Acceso e-Rec de Llacruz</title>
    <meta name="description" content="Acceso e-Rec de Llacruz">
    <meta name="author" content="Departamento desarrollo Llacruz">
    <link href="../css/normalize.css" rel="stylesheet"> 
    <link href="../css/style_generales_sistema.css" rel="stylesheet">
    <link href="../css/style_Carga_Gestiones_Masiva.css" rel="stylesheet">
   

	<script type='text/javascript' src='../Componentes/jquery-1.9.2/js/jquery-1.8.3.js' ></script>

</head>
<body>
<div class="titulo_informe">CARGA Contactabilidad</div>   
<table width="90%" border="0" align="center" cellpadding="0" cellspacing="0">
	<tr>
        <td colspan="2" align="right">
          <table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
         
            <tr valign="TOP">
              <td colspan="3">
                  <table width="100%" height="335" border="0" cellpadding="0" cellspacing="0" >
                    <tr>
                      <td height="331" valign="top"> 
                      <table width="100%" border="0">
                          <tr>
                            <td>
							
                            </td>
                          </tr>
                        </table>

						<input class="parametro" type="hidden" id="cliente-codigo" value="<%=trim(session("ses_codcli"))%>">
	<input class="parametro" type="hidden" id="usuario" value="<%=trim(session("session_idusuario"))%>">

</body>
</html>

<%
if session("session_idusuario") <> "" then 

dim archivo 
dim ruta 
ruta = "d:\app\EREC\Archivo\Integracion\"
archivo  = ruta  & "Intg_"& replace(Request.ServerVariables("REMOTE_ADDR"),".","") &".txt"
'response.write (archivo)
Dim objFSO, objTextFile
Dim sRead, sReadLine, sReadAll
Const ForReading = 1, ForWriting = 2, ForAppending = 8

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objTextFile = objFSO.CreateTextFile(archivo, True)

Nombre_usuario	=session("nombre_user")
Perfil			=session("session_tipo")
Logins			=session("session_login")
Cod_Cliente		=session("ses_codcli")'''
Id_Usuario		=session("session_idusuario")
Rut_Usuario		=session("session_user") 	 
iniciosesion 	= Date() &" - "& time()
pass			=session("ses_clave") '''
perfil_adm		=session("perfil_adm") 
perfil_caja		=session("perfil_caja") 
perfil_emp		=session("perfil_emp") 
perfil_sup		=session("perfil_sup") 
perfil_full		=session("perfil_full") 
perfil_cob		=session("perfil_cob") 
valor_uf		=session("valor_uf")
valor_moneda	=session("valor_moneda") 

 '' nombre/perfil/login/cod_cliente/id_usuario/rut_usurio/fecha+hora/ses_clave/perfil_adm/perfil_caja/perfil_emp/perfil_sup/perfil_full/Pagina/valormoneda/valorUF
objTextFile.Write(Nombre_usuario)
objTextFile.Write(";"&Perfil)
objTextFile.Write(";"&Logins)
objTextFile.Write(";"&Cod_Cliente)
objTextFile.Write(";"&Id_Usuario)
objTextFile.Write(";"&Rut_Usuario)
objTextFile.Write(";"&iniciosesion)
objTextFile.Write(";"&pass)
objTextFile.Write(";"&perfil_adm)
objTextFile.Write(";"&perfil_caja)
objTextFile.Write(";"&perfil_emp)
objTextFile.Write(";"&perfil_sup)
objTextFile.Write(";"&perfil_full)
objTextFile.Write(";"&perfil_cob)
objTextFile.Write(";"&"CargaTelefonos.aspx")
objTextFile.Write(";"&valor_uf)
objTextFile.Write(";"&valor_moneda)

objTextFile.WriteBlankLines(17)
objTextFile.Close
%>
<script>
{
top.location.href = "http://sistemas.llacruz.cl/CargaContactabilidad/CargaTelefonos.aspx";
}
</script>
<%else
response.write("session expirada")
end if%>

