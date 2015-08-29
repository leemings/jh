<%
dim connurl,strurl,dburl
set connurl=server.createobject("adodb.connection")
dburl="e_url.asp"
strurl="provider=Microsoft.Jet.OLEDB.4.0;Data Source="&Server.MapPath(dburl)
'如果您的服务器使用较老版本的ACCESS数据库，请使用下面的连接方法
'strurl="driver={Microsoft Access Driver(*.mdb)};dbq=&Server.MapPath(dburl)
connurl.open strurl
%>