<SCRIPT LANGUAGE="JavaScript">

//-------------------------------------------------------------------------------
// Propósito:   Devuelve verdadero si la entrada a evaluar corresponde al tipo de dato requerido
// Entradas:    tipo: el tipo de dato a evaluar (tyniint,smallint,int,float) - inputText: el texto a evaluar
// Returns: True - si es el tipo de dato correspondiente , False no lo es
//-------------------------------------------------------------------------------

function esNumero(tipo,inputText)
{
    var tipo,int,smallint,tyniint;
    var inputLength = inputText.length;
    var isNumber = "true";

    for (i = 0; i < inputLength; i++) {
        var aChar = inputText.substring(i,i+1);
	if (aChar < "0" || aChar > "9" )  {
		if (aChar == "ƒ") {
		        isNumber = "true";
		    }
		else
		    {
			if ((aChar == ".") && ((tipo == 'money')||(tipo == 'real')))
				{
				isNumber = "true";
				}
			else
				{
				isNumber = "false";
				}
		    }
   		}	
	}



        if (inputLength == 0)
          isNumber = "false";

        if( isNumber == "true" ) {

        intVal = parseFloat(inputText);
        switch (tipo){
            case 'float':
                if ((intVal < -9999999999999999999999999999999999999) || (intVal > 9999999999999999999999999999999999999))
                isNumber = "false";
                break;
            case 'int':
                if ((intVal < -2147483648) || (intVal > 2147483648))
                isNumber = "false";
                break;
            case 'smallint':
                if ((intVal < -32768) || (intVal > 32767 ))
                isNumber = "false";
                break;
            case 'tinyint':
                if ((intVal < 0) || (intVal > 255 ))
                isNumber = "false";
                break;
            case 'real':
				if ((intVal < -922337203685477.5808) || (intVal > 922337203685477.5807 ))
				isNumber = "false";
				break;
            case 'money':
	        if ((intVal < -922337203685477.5808) || (intVal > 922337203685477.5807 ))
	        isNumber = "false";
                break;
            default:


        }

    }
        return isNumber;

}


function esNumeroNegativo(tipo,inputText)
{
    var tipo,int,smallint,tyniint;
    var inputLength = inputText.length;
    var isNumber = "true";

    for (i = 0; i < inputLength; i++) {
        var aChar = inputText.substring(i,i+1);
	if ((aChar < "0" || aChar > "9" ) && (aChar != "-" )) {
		if (aChar == "ƒ") {
		        isNumber = "true";
		    }
		else
		    {
			if ((aChar == ".") && ((tipo == 'money')||(tipo == 'real')))
				{
				isNumber = "true";
				}
			else
				{
				isNumber = "false";
				}
		    }
   		}	
	}



        if (inputLength == 0)
          isNumber = "false";

        if( isNumber == "true" ) {

        intVal = parseFloat(inputText);
        switch (tipo){
            case 'float':
                if ((intVal < -9999999999999999999999999999999999999) || (intVal > 9999999999999999999999999999999999999))
                isNumber = "false";
                break;
            case 'int':
                if ((intVal < -2147483648) || (intVal > 2147483648))
                isNumber = "false";
                break;
            case 'smallint':
                if ((intVal < -32768) || (intVal > 32767 ))
                isNumber = "false";
                break;
            case 'tinyint':
                if ((intVal < 0) || (intVal > 255 ))
                isNumber = "false";
                break;
            case 'real':
				if ((intVal < -922337203685477.5808) || (intVal > 922337203685477.5807 ))
				isNumber = "false";
				break;
            case 'money':
	        if ((intVal < -922337203685477.5808) || (intVal > 922337203685477.5807 ))
	        isNumber = "false";
                break;
            default:


        }

    }
        return isNumber;

}



</SCRIPT>

