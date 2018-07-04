<%
Set ConexionSDR = Server.CreateObject("ADOdb.Connection")
ConexionSDR.Open "uid=usr_sdr;pwd=usr_sdr;driver={SQL Server};server=ICARUS;database=SDR"
%>