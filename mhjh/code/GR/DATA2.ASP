<%
Set connb = Server.CreateObject("ADODB.Connection")
connb.Open "Driver={Microsoft Access Driver (*.mdb)}; DBQ=" & Server.MapPath("../news/chaoliu.asp")
Set rs = Server.CreateObject("ADODB.Recordset")
%>

