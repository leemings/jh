<%
Set connb = Server.CreateObject("ADODB.Connection")
connb.Open "Driver={Microsoft Access Driver (*.mdb)}; DBQ=" & Server.MapPath("../../21ex/gd.asp")
Set rs = Server.CreateObject("ADODB.Recordset")
%>

