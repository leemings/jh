<%
Set cn = Server.CreateObject("ADODB.Connection")
cn.Open "Driver={Microsoft Access Driver (*.mdb)}; DBQ=" & Server.MapPath("../../21ex/mj.asp")
Set rs = Server.CreateObject("ADODB.Recordset")
%>

