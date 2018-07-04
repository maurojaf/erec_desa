<% 
'------------------------------------------------------------------------------------
' by gsus internet art 
' http://www.cedecero.com/gsus
' codigo para su libre utilización
'------------------------------------------------------------------------------------
   
    'Previamente el fichero  Anterior.txt 
    'ha de existir en nuestra carpeta.
	
    'Declaracion de variables
    Dim FSO, Fich , NombreAnterior, NombreNuevo 
    'Inicialización
  
    DestinationPath = Server.mapPath("../Archivo/BibliotecaDeudores") 

  ' Instanciamos el objeto
   Set Obj_FSO = CreateObject("Scripting.FileSystemObject") 
   ' Asignamos el fichero a renombrar a la variable fich
  Set objfolder = Obj_FSO.GetFolder (DestinationPath)
   ' llamamos a la funcion copiar, 
   'y duplicamos el archivo pero con otro nombre

    cont=0
    for each objfile22 in objfolder.files  'Recorre los ficheros del directorio raiz
      cont=cont+1
      if trim(objfile22.name)<>"Thumbs.db" then
       
        'response.write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"&objfile22.name&"<br>" 
        Set Fich = Obj_FSO.GetFile((DestinationPath&"/"&objfile22.name)) 
      
        'Fich.Copy((DestinationPath&"/"&objfile22.name)) 
        response.write replace(objfile22.name,"á","a")&"<br>" 
        'Fich.Copy((DestinationPath&"/"&replace(objfile22.name,"á","a"))) 
        
        'Fich.Delete() 

      end if
      
    next



 %>