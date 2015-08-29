<%
dim dbplus,connplus,plusstr
dbplus="data/e_plus.asp"   '请改成您自己的数据库文件名
Set connplus = Server.CreateObject("ADODB.Connection")
plusstr="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(dbplus)
connplus.Open plusstr
%>