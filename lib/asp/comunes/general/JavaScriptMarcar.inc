<SCRIPT LANGUAGE=JavaScript>

function MarcaTodos(objeto) {
	for (i = 0; i < objeto.length; i++)
	{
	  objeto.options[i].selected=true;
	}
	objeto.disabled = true
}

function DesMarcaTodos(objeto) {
	for (i = 0; i < objeto.length; i++)
	{
	  objeto.options[i].selected=false;
	}
	objeto.disabled = false
}


</SCRIPT>