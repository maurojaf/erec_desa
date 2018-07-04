function selCombox(objCombo, valor) {

	for (i=0;i < objCombo.options.length; i ++) {
		if (objCombo.options[i].value == valor) {
				objCombo.options[i].selected = true; 
		}
	}
}