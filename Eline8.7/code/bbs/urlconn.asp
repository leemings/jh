<%
dim connurl,strurl,dburl
set connurl=server.createobject("adodb.connection")
dburl="e_url.asp"
strurl="provider=Microsoft.Jet.OLEDB.4.0;Data Source="&Server.MapPath(dburl)
'������ķ�����ʹ�ý��ϰ汾��ACCESS���ݿ⣬��ʹ����������ӷ���
'strurl="driver={Microsoft Access Driver(*.mdb)};dbq=&Server.MapPath(dburl)
connurl.open strurl
%>