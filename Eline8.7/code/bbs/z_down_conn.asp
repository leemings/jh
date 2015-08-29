<%
dim dbdown,conndown,downstr 
dbdown="data/e_download.asp"    '下载数据库名
Set conndown = Server.CreateObject("ADODB.Connection")
downstr="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(dbdown)
conndown.Open downstr
%>