<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001" LCID = 1034%>
<!--#include file="System/arch_utils_login.asp"-->
<%
Response.CodePage=65001
Response.charset ="utf-8"

AbrirSCG()
sql_sel_empresa ="exec proc_informacion_empresa"
set rs_sel_empresa = conn.execute(sql_sel_empresa)
if err then
    Response.write "Error información de empresa. Favor contactar a administrador"
end if

if not rs_sel_empresa.eof then
    DIRECCION_EMPRESA           =rs_sel_empresa("DIRECCION_EMPRESA")
    TELEFONOS_EMPRESA           =rs_sel_empresa("TELEFONOS_EMPRESA")
    CORREO_CONTACTO             =rs_sel_empresa("CORREO_CONTACTO")
    CORREO_SOPORTE              =rs_sel_empresa("CORREO_SOPORTE")
    TELELFONO_MESA_AYUDA        =rs_sel_empresa("TELELFONO_MESA_AYUDA")
    HORARIO_ATENCION            =rs_sel_empresa("HORARIO_ATENCION")
end if
CerrarSCG()

usuario_nombre  =request.cookies("usuario_nombre")
contrasena      =request.cookies("contrasena")

if trim(usuario_nombre)<>"" AND trim(contrasena)<>"" then
    muestra_redordarme =" checked "
end if
%> 
<!DOCTYPE html>
<html lang="es">
<head>
    <meta http-equiv="X-UA-Compatible" content="IE=edge,chrome=1">
    <meta charset="utf-8">

    <title>Acceso e-Rec de Llacruz</title>
    <meta name="description" content="Acceso e-Rec de Llacruz">
    <meta name="author" content="Departamento desarrollo Llacruz">
    <!--[if lt IE 9]>
        <script src="http://html5shim.googlecode.com/svn/trunk/html5.js"></script>
    <![endif]-->

    <link href="css/normalize.css" rel="stylesheet"> 
    <link href="css/bootstrap.css" rel="stylesheet"> 
    <link href="css/style_login.css" rel="stylesheet"> 
    <link href="Componentes/jquery-1.9.2/css/start/jquery-ui-1.9.2.custom.css" rel="stylesheet"> 
    <!--#include file="system/arch_utils.asp"-->
 
</head>
<body>

	<header>
        <figure>
            <div>
                <img class="img_login" src="Imagenes/logo.png">         
            </div>
        </figure>
	</header>	

	<section>
		<form action="" method="post" accept-charset="UTF-8">
				
			<div class="container2">                
                <h1>Bienvenido</h1>
                <div class="titulo2">al sistema de cobranzas e-Rec de Llacruz</div>

                <div class="contenido_inicio">

                    <h3>Iniciar sesión e-Rec</h3>
                    <div  class="contenedores_inicio">
                        <div class="titulo_inicio">USUARIO</div>
                        <i class="icon-user"></i>
                        <input class="input_inicio" type="text" name="usuario_nombre" id="usuario_nombre" value="<%=trim(usuario_nombre)%>">
                        <span class="muted">Ej: Pedro Millalaf > pmillalaf</span>
                        <span class="text-error alert_usuario">¡Ingresa nombre usuario!</span>
                    </div>
                    
                    <div  class="contenedores_inicio">
                        <div class="titulo_inicio">CLAVE</div>
                        <i class="icon-lock"></i>
                        <input class="input_inicio" type="password" name="contrasena" id="contrasena" value="<%=trim(contrasena)%>">
                        <span class="btn-link" onclick="bt_olvida_contrasena()">¿Olvidó su contraseña?</span>
                        <span class="text-error alert_contrasena">¡Ingresa contraseña!</span>
                    </div>
                    <div class="contenedores_inicio">
                        <label class="checkbox">
                              <input type="checkbox" <%=muestra_redordarme%> id="login_recordarme" name="login_recordarme"> Recordarme
                        </label>
                    </div>

                    <div class="contenedores_inicio">
                        <input type="button" class="btn btn-small" name="" onclick="bt_validar_usuario()" value="Ingresar">
                    </div>
                    <div id="span_mensaje_error"></div>
                </div>

			</div>
            <div id="action_section"></div>
            <div id="verifica_intentos_fallidos"></div>
		</form>

	</section>
        <div class="footer_ayuda">
           
            <h6> 
                <i class="icon-envelope"></i>
                <span class="btn-link"> ¿Necesita ayuda?</span> Comuníquese con la mesa de ayuda al <%=TRIM(TELELFONO_MESA_AYUDA)%> o escríbanos a <span class="btn-link"><%=TRIM(CORREO_SOPORTE)%></span>
            </h6>

        </div> 
	<footer>

        <div>
            <figure>
                <img class="footer_img" src="Imagenes/logollacruzgris.png"> 
                <span class="footer_label">
                    <span class="text-info">Asesorías e Inversiones Ltda.</span> | Teléfono: <%=TRIM(TELEFONOS_EMPRESA)%> Dirección: <%=TRIM(DIRECCION_EMPRESA)%> Mail: <span class="btn-link"><%=TRIM(CORREO_CONTACTO)%>. </span>Horario de atención: <%=TRIM(HORARIO_ATENCION)%>
                </span>    
            </figure>
        </div>
        <div id="contrasena_olvidada" title="Olvidaste tu contraseña" style="display:none;">
            <div id="contrasena_olvidada"></div>
            <div class="titulo_olvido_contrasena">Ingresa Usuario</div>
            <div><input type="text" name="login_usuario" id="login_usuario" value=""></div><span id="span_CORREO_ELECTRONICO" class="span_aviso_rojo"></span><span id="span_CORREO_ELECTRONICO_enviado" class="span_aviso_azul"></span>
        </div>

        <div id="bloqueo_3_fallidos" style="display:none;">
            <p class="bloqueo_3_fallidos_alert">Computador bloqueado por tres intentos fallidos de logeo</p><p>Espera 5 min y se desbloqueará automáticamente</p>
        </div>

	</footer>

    <script src="Componentes/jquery-1.9.2/js/jquery-1.8.3.js"></script>  
    <script src="Componentes/jquery-1.9.2/js/jquery-ui-1.9.2.custom.min.js"></script>  
    <script src="Javascripts/bootstrap.js"></script>
    <script src="Javascripts/js_login.js"></script>

</body>
</html>
