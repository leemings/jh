<%
dim dbfastpost,connfastpost,fastpoststr
dbfastpost="data/e_fastpost.asp"   '请改成您自己的数据库文件名
Set connfastpost = Server.CreateObject("ADODB.Connection")
fastpoststr="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(dbfastpost)
connfastpost.Open fastpoststr
%>