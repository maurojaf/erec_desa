<%


            session("ses_clave") 		 =request.QueryString("clave")
			session("ses_codcli") 		 =request.QueryString("CB_CLIENTE")
			session("session_idusuario") =request.QueryString("ID_USUARIO")
			session("session_user") 	 =request.QueryString("RUT_USUARIO")
			session("session_login") 	 =request.QueryString("LOGIN")
            session("nombre_user") 		 =request.QueryString("nombre_user")
            session("session_tipo") 	 =request.QueryString("PERFIL")
            session("iniciosesion") 	 = Date() &" - " & Time()   'rsUSU("FH") & " - " & rsUSU("HH")
            
            session("perfil_adm") 		 =request.QueryString("PERFIL_ADM")
			session("perfil_caja") 		 =request.QueryString("PERFIL_CAJA")
			session("perfil_emp") 		 =request.QueryString("PERFIL_EMP")
			session("perfil_sup") 		 =request.QueryString("PERFIL_SUP")
			session("perfil_full") 		 =request.QueryString("PERFIL_FULL")
			session("perfil_cob") 		 =request.QueryString("PERFIL_COB")
			session("Pagina")				 		 =request.QueryString("Pagina")


            response.Write("ses_clave:" & session("ses_clave"))
            response.Write("<br/>ses_codcli:" & session("ses_codcli"))
            response.Write("<br/>session_idusuario:" & session("session_idusuario"))
            response.Write("<br/>session_user:" & session("session_user"))
            response.Write("<br/>session_login:" & session("session_login"))
            response.Write("<br/>nombre_user:" & session("nombre_user"))
            response.Write("<br/>session_tipo:" & session("session_tipo"))
            response.Write("<br/>iniciosesion:" & session("iniciosesion"))


            response.Write("<br/>perfil_adm:" & session("perfil_adm"))
            response.Write("<br/>perfil_caja:" & session("perfil_caja"))
            response.Write("<br/>perfil_emp:" & session("perfil_emp"))
            response.Write("<br/>perfil_sup:" & session("perfil_sup"))
            response.Write("<br/>perfil_full:" & session("perfil_full"))
            response.Write("<br/>perfil_cob:" & session("perfil_cob"))
			response.Write("<br/>Pagiina:" & Pagina)

           ' response.end
         'response.Redirect(pagina)  
		  response.Redirect("default.asp")

		
		%>