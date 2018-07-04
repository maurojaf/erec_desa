<SCRIPT LANGUAGE="JavaScript">
// Este SCRIPT se incluye en la ventana de quienes lamana a los buscadores

function Selector_Actualizar( CampoHTML, Resultado )  {

    if( CampoHTML != "" ) {
	//Recupera objeto campo
	var objCampo = document.forms[0].elements[ CampoHTML ];
	objCampo.value = Resultado;
    }
}

//-------------------------------------------------------------------------------
// Propósito: Abrir una ventana de tipo Buscador, que permita seleccionar un registro
// Entradas: 	URL = direccion del buscador
//		Titulo = titulo de la ventana del buscador
//		Codigo = Codigo de un formulario a actualizar
//		ancho = ancho de la ventana del buscador
//		alto = alto de la ventana del buscador
// Retorna: false
//-------------------------------------------------------------------------------
function Selector_Item(URL, Titulo, Codigo, ancho, alto)  {
     //Deja en variable de ventana el campo a modificar
     window.Selector_CodigoAModificar = Codigo;
     window.Selector_DescripcionAModificar = "";
      //Abre ventana

     settings=
     	"toolbar=no,location=no,directories=no,"+
     	"status=no,menubar=no,scrollbars=yes,"+
     	"resizable=no,width=" + ancho + ",height=" + alto +
     	",top=20,left=0";
     window.open( URL, Titulo, settings);
	 //Retorna
    return( false );
}
function Selector_Item_Descripcion(URL, Titulo, Codigo, Descripcion, ancho, alto)  {
     //Deja en variable de ventana el campo a modificar
     window.Selector_DescripcionAModificar = Descripcion;
     window.Selector_CodigoAModificar = Codigo;
      //Abre ventana
     settings="toolbar=no,location=no,directories=no,"+
     	"status=no,menubar=no,scrollbars=yes,"+
     	"resizable=no,width=" + ancho + ",height=" + alto +
     	",top=20,left=0";
     window.open( URL, Titulo, settings);
	 //Retorna
    return( false );
}

</SCRIPT>

