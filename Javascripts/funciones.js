function Valida_Rut(Vrut)
{
	var dig	
	Vrut = Vrut.split("-");

	if (!isNaN(Vrut[0]))
	{					
		largo_rut = Vrut[0].length;
		if ((largo_rut >= 7 ) && (largo_rut <= 8))
		{
			if (largo_rut > 7)
			{
				multiplicador = 3;
			}
			else
			{
				multiplicador = 2;
			}
			suma = 0;
			contador = 0;
				do
				{
					digito = Vrut[0].charAt(contador);
					digito = Number(digito);
						if (multiplicador == 1)
						{
							multiplicador = 7;
						}

					suma = suma + (digito * multiplicador);
					multiplicador --;
					contador ++;
				}
				while (contador < largo_rut);
			resto = suma % 11	
			dig_verificador = 11 - resto;

				if (dig_verificador == 10)
				{
					dig = "k";
				}
				else if (dig_verificador == 11)
				{
					dig = 0
				}
				else
				{
					dig = dig_verificador;
				}

				if (dig != Vrut[1])
				{
					//window.alert ("El Rut 1 es invalido !");
					return 0;
				}
		}
		else
		{
			//window.alert("El Rut  2 es invalido ! ");
			return 0;
		}
	}
	else
	{
		//window.alert("El Rut 3 es invalido ! ");
		return 0;
	}
		return 1;
}

