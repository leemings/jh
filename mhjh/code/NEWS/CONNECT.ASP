<%
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Driver={Microsoft Access Driver (*.mdb)}; DBQ=" & Server.MapPath("mlimlihong.mdb")
Set rs = Server.CreateObject("ADODB.Recordset")
%>
