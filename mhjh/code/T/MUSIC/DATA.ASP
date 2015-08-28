<%
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Driver={Microsoft Access Driver (*.mdb)}; DBQ=" & Server.MapPath("ysqx.asp")
Set rs = Server.CreateObject("ADODB.Recordset")
%>

