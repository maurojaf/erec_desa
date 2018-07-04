<TABLE ALIGN="CENTER" id="" BORDER="1" width="100%">
			<TR>
					<TD valign="TOP" WIDTH="35%" >
						<TABLE ALIGN="CENTER" WIDTH="100%" BORDER="0">
							<TR HEIGHT="30">
								<td COLSPAN="2" ALIGN="CENTER" class="estilo_columna_individual">
									MONTO DE DEUDA
								</td>
							</TR>

							<TR HEIGHT="30">
								<TD WIDTH="150" ALIGN="LEFT" class="Estilo22">Capital: </TD>
								<TD><input name="PP_TX_CAPITAL" id="PP_TX_CAPITAL" value="" type="text" size="10"  readonly="readonly"></TD>
							</TR>
							<TR HEIGHT="30">
								<TD ALIGN="LEFT" class="Estilo22">Interes: </TD>
								<TD><input name="PP_TX_INTERES" id="PP_TX_INTERES" type="text" size="10" value="0"  readonly="readonly"></TD>
							</TR>
							<TR HEIGHT="30">
								<TD align="LEFT" class="Estilo22">Gastos Judiciales: </TD>
								<TD><input name="PP_TX_GASTOS" id="PP_TX_GASTOS" type="text" size="10" value="<%=0%>"    readonly="readonly"></TD>
							</TR>
							<TR HEIGHT="30">
								<TD align="LEFT" class="Estilo22">Gastos Protestos: </TD>
								<TD><input name="PP_TX_GASTOSPROTESTOS" id="PP_TX_GASTOSPROTESTOS" type="text" size="10" value="<%=0%>" readonly="readonly"></TD>
							</TR>
							<TR HEIGHT="30">
								<TD align="LEFT" class="Estilo22">Indem.Comp.: </TD>
								<TD><input name="PP_TX_INDEM_COMP" id="PP_TX_INDEM_COMP" type="text" size="10" value="<%=0%>" readonly="readonly"></TD>
							</TR>

							<TR HEIGHT="30">
								<TD align="LEFT" class="Estilo22">Honorarios : </TD>
								<TD><input name="PP_TX_HONORARIOS" id="PP_TX_HONORARIOS" type="text" size="10" value="" readonly="readonly"></TD>
							</TR>
							<TR HEIGHT="30">
								<TD>&nbsp;</TD>
								<TD>______________</TD>
							</TR>
								<TR HEIGHT="30">
									<TD align="right" class="Estilo22">Total Deuda: </TD>
									<TD><input disabled value ="" name="PP_TX_TOTALDEUDA" id="PP_TX_TOTALDEUDA" type="text" size="10" readonly="readonly"></TD>
								</TR>
						</TABLE>


					</TD>

					<TD valign="TOP" WIDTH="35%" >

						<TABLE ALIGN="CENTER" WIDTH="100%" BORDER="0">
						  <TR HEIGHT="30">
							<TH COLSPAN="2" align="CENTER" class="estilo_columna_individual">
								DESCUENTOS
							</TH>

						  </TR>
						  <TR HEIGHT="30">
							  <TD width="150" ALIGN="LEFT" class="Estilo23">Capital:</TD>
							  <TD>
									%<input class="porc_desc_capital" id="porc_desc_capital" name="porc_desc_capital" type="text" size="3"   onblur="func_porc_desc_capital();"  value="0" maxlength="3">
									$<input class="desc_capital" id="desc_capital" name="desc_capital" type="text" size="8"  onblur="func_descuentos(this.value,'DESCUENTO');" value="0">
							  </TD>
						  </TR>
						  <TR HEIGHT="30">
							  <TD ALIGN="LEFT" class="Estilo23">Interes:</TD>
							  <TD>
									%<input class="porc_desc_interes" id="porc_desc_interes" name="porc_desc_interes" type="text" size="3"      onblur="func_porc_desc_interes();"  value="0" maxlength="3">
									$<input class="desc_interes" id="desc_interes" name="desc_interes" type="text" size="8"     onblur="func_descuentos(this.value,'INTERES');" value="0">
							  </TD>
						  </TR>
						  <TR HEIGHT="30">
								<TD ALIGN="LEFT" class="Estilo23">Gastos Judiciales:</TD>
								<TD>
									%<input DISABLED name="porc_desc_gastos" id="porc_desc_gastos" type="text" size="3"   onblur="func_porc_desc_gastos();" value="0" maxlength="3">
									$<input DISABLED name="desc_gastos" id="desc_gastos" type="text" size="8"  value="0" onblur="func_descuentos(this.value,'JUDICIAL');">
								</TD>
						  </TR>
						   <TR HEIGHT="30">
								<TD ALIGN="LEFT" class="Estilo23">Gastos Protestos:</TD>
								<TD>
									%<input DISABLED name="porc_desc_gastosprotestos" id="porc_desc_gastosprotestos" type="text" size="3"  onblur="func_porc_gastosprotestos();" value="0" maxlength="3">
									$<input DISABLED name="GASTOS_PROTESTOS" id="GASTOS_PROTESTOS" type="text" size="8" value="0"    onblur="func_descuentos(this.value,'PROTESTO');"></TD>
						  </TR>

						 <TR HEIGHT="30">
								<TD ALIGN="LEFT" class="Estilo23">IndemComp:</TD>
								<TD>
									%<input DISABLED name="porc_desc_indemComp" id="porc_desc_indemComp" type="text" size="3"   onblur="func_porc_indemComp();"  value="0" maxlength="3">
									$<input DISABLED name="desc_indemComp" id="desc_indemComp" type="text" size="8"   value="0" onblur="func_descuentos(this.value,'INDEMCOMP');">
								</TD>
						  </TR>

						  <TR HEIGHT="30">
							<td  align="LEFT" class="Estilo23"> Honorarios:</TD>
							<TD>
								%<input class="porc_desc_honorarios" id="porc_desc_honorarios"  name="porc_desc_honorarios" type="text" size="3"    onblur="func_porc_desc_honorarios();"value="0"  maxlength="3">
								$<input class="desc_honorarios" id="desc_honorarios"  name="desc_honorarios" type="text" size="8"  value="0"    onblur="func_descuentos(this.value,'HONORARIOS');">
							</TD>
						   </TR>
							<TR HEIGHT="30">
								<TD>&nbsp;</TD>
								<TD>______________</TD>
							</TR>

						   <TR HEIGHT="30">
								<TD>Total Deuda con Descuento</TD>
								<TD>$<input name="PP_TX_TOTALDEUDA_DESC" id="PP_TX_TOTALDEUDA_DESC" type="text" size="10" disabled readonly="readonly" value ="<%=FormatNumber(intTotalDeuda,0) %>" ></TD>
							</TR>

						  </TABLE>

					</TD>
					<TD valign="TOP" WIDTH="32%" >

					  <TABLE ALIGN="CENTER" BORDER="0" WIDTH="100%" valign="TOP">
					  <TR HEIGHT="30">
							<TH COLSPAN=2 ALIGN="CENTER" class="estilo_columna_individual">
								MODALIDAD DEL PAGO
							</TH>
					  </TR>
						<TR valign="TOP">
							<TD HEIGHT="30" ALIGN="LEFT" class="Estilo23">
							Pie a cancelar:$
							</TD>
						</TR>
						<TR HEIGHT="30" valign="TOP">
							<TD ALIGN="LEFT" class="Estilo23">
							Abono Deuda&nbsp;
                            %<input class="porc_capital_pie" id="porc_capital_pie"  name="porc_capital_pie" type="text" size="2" value="<%=intPorcPie*100%>" onblur="func_porc_capital_pie();" maxlength="3"  >
							$<input class="pie" id="pie" name="pie" type="text" size="10"   
                                    
									onblur="CalculateCapitalPercentageAndRefreshAgreement();"
                                    value="0"  
                                     maxlength="10" >
                        </TD>
						</TR>
					  <TR HEIGHT="30">
						  <TD width="200" ALIGN="LEFT" class="Estilo23" >Cantidad de cuotas: 
							<select name="cuotas" size="1" style="width:50px;" id="cuotas">
									<option value="-">-</option>
									<option value="1">1</option>
									<option value="2">2</option>
									<option value="3">3</option>
									<option value="4">4</option>
									<option value="5">5</option>
									<option value="6">6</option>
									<option value="7">7</option>
									<option value="8">8</option>
									<option value="9">9</option>
									<option value="10">10</option>
									<option value="11">11</option>
									<option value="12">12</option>
									<option value="13">13</option>
									<option value="14">14</option>
									<option value="15">15</option>
									<option value="16">16</option>
									<option value="17">17</option>
									<option value="18">18</option>
									<option value="19">19</option>
									<option value="20">20</option>
									<option value="21">21</option>
									<option value="22">22</option>
									<option value="23">23</option>
									<option value="24">24</option>
									<option value="25">25</option>
									<option value="26">26</option>
									<option value="27">27</option>
									<option value="28">28</option>
									<option value="29">29</option>
									<option value="30">30</option>
									<option value="31">31</option>
									<option value="32">32</option>
									<option value="33">33</option>
									<option value="34">34</option>
									<option value="35">35</option>
									<option value="36">36</option>
							  </select>
							  &nbsp;<br />
							  Dia de Pago: 
							  <input name="PP_TX_DIAPAGO" id="PP_TX_DIAPAGO" type="text" value="5" size="3" maxlength="5"  onkeyUp="return ValNumero(this);" >
						  </TD>
					  </TR>
						<TR>
							<TD align="LEFT" class="Estilo22">&nbsp;</TD>
						</TR>
						<TR>
							<TD align="LEFT" class="Estilo22">&nbsp;</TD>
						</TR>
												
						<TR>
							<TD align="LEFT" class="Estilo22">Total A Convenir: </TD>
						</TR>
						<TR>
							<TD><input name="PP_TX_TOTALCONVENIO" id="PP_TX_TOTALCONVENIO" type="text" size="10" readonly="readonly" 
                                    value ="" ></TD>
						</TR>

					  </TABLE>
				 </TD>
			</TR>
</TABLE>

<br />

<table width="100%">
	<tr>
		<td style="text-align: right;">
			<input class="fondo_boton_100" type="button" id="ButtonGenerarPlanPago" name="ButtonGenerarPlanPago" value="Generar" />
		</td>
	</tr>
</div>