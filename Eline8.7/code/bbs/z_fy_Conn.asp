<%
dim dbfy,connfy,fystr
dbfy="data/e_fyjy.asp"   '请改成您自己的数据库文件名
Set connfy = Server.CreateObject("ADODB.Connection")
fystr="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(dbfy)
connfy.Open fystr
%>
