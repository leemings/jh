<%
Set connb = Server.CreateObject("ADODB.Connection")
connb.Open "Driver={Microsoft Access Driver (*.mdb)}; DBQ=" & Server.MapPath("../../21ex/xw.asp")
Set rs = Server.CreateObject("ADODB.Recordset")
%>

