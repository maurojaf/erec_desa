function Trim(strTexto)
{
var intLargo = strTexto.length;
var intInd = -1;
var intPosI = -1;
var intPosD = 0;

	if( intLargo == 0 ) { return '';}
	for( intInd = 0; intInd < intLargo; intInd++)
	{ if( strTexto.charAt(intInd) != ' ' )
		{ intPosI = intInd, intInd = intLargo }
	}
	if( intPosI == -1 ) { return '' }
	for( intInd = intLargo-1; intInd > -1; intInd-- )
	{ if( strTexto.charAt(intInd) != ' ' )
		{ intPosD = intInd; intInd = -1 }
	}
	return strTexto.substring(intPosI, intPosD+1)
}
