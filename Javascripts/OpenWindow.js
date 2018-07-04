function OpenWindow()
{
var NameWindow    = OpenWindow.arguments[0];
var window_width  = OpenWindow.arguments[1];
var window_height = OpenWindow.arguments[2];
var type_window   = OpenWindow.arguments[3];
var strResizable  = OpenWindow.arguments[4];
var strTitulo     = OpenWindow.arguments[5];
var strScroll     = OpenWindow.arguments[6];
var window_top    = OpenWindow.arguments[7];
var window_left   = OpenWindow.arguments[8];
var window_prop   = "";

	if(typeof(type_window) == 'undefined')
	{	type_window = document.all;	}
	
  if(typeof(window_height) == 'undefined')
	{	if(type_window) 
		{	window_height = 514; }
		else
		{	window_height = 390; }
  }
	if(typeof(window_width) == 'undefined')
	{	window_width  = 580; }
	
	if (typeof(strResizable) == 'undefined')
	{	strResizable  = 'no'; }

	if (typeof(strScroll) == 'undefined')
	{	strScroll  = 'no'; }

	if(typeof(strTitulo) == 'undefined')
	{	strTitulo  = '';	}

  //Calcula posicion de ventana
	if(typeof(window_top) == 'undefined') {
		window_top	= (screen.availHeight - window_height) / 2;
	}
	if(typeof(window_left) == 'undefined') {
  	window_left	= (screen.availWidth  - window_width ) / 2;
	}

  //Construye string con propiedades de ventana a abrir
  if (type_window == 'win_nor')
	{	
	  window_prop += "status=no,toolbar=no,menubar=no,location=no,";
		window_prop += "resizable=" + strResizable + ",";
    window_prop += "height=" + window_height + ",";
    window_prop += "width="  + window_width  + ",";
    window_prop += "top="    + ( ( window_top  > 0 ) ? window_top  : 0 ) + ",";
    window_prop += "left="   + ( ( window_left > 0 ) ? window_left : 0 ) + ",";
    window_prop += "dependent=yes" + ",";
    window_prop += "scrollbars=" + strScroll ;
  }
  else
	{	window_prop += "status=no" + ";";
    window_prop += "center=yes" + ";";
    window_prop += "dialogHeight=" + window_height + "px;";
    window_prop += "dialogWidth="  + window_width  + "px;";
		window_prop += "dialogTop=" + window_top + "px;";
    window_prop += "dialogLeft="  + window_left + "px;";
  }
  //Abre ventana
  if (type_window == 'win_nor')
	{	window.open(NameWindow, strTitulo, window_prop); }
  else
	{	window.showModalDialog(NameWindow, this, window_prop ); }
}
