<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
username=Session("Ba_jxqy_username")
if username="" then Response.Redirect "../error.asp?id=016"
set conn=server.CreateObject("adodb.connection")
conn.Open Application("Ba_jxqy_connstr")
set rst=server.CreateObject("adodb.recordset")
rst.Open "select * from pc21 where username='"&username&"'",conn
if rst.EOF or rst.BOF then Response.Redirect "../error.asp?id=100"
zpoint=rst("zpoint")
upoint=rst("upoint")
bet=rst("bet")
msg="ׯ��"&zpoint&"�㣬��"&upoint&"�㣬"
if not(upoint>21 or zpoint>=upoint) then 
	conn.Execute "update �û� set ����=����+"&bet*2&" where ����='"&username&"'"
	conn.Execute "update pc21 set win=win+1,zpoint=21 where username='"&username&"'"
	msg=msg&"��Ӯ��"&bet&"�����ӡ�"
else
	conn.Execute "update pc21 set zpoint=21,bet=0 where username='"&username&"'"
	msg=msg&"������˼��������"&bet&"�����ӡ�"
end if
msg=msg&"<script language=javascript>parent.imgfrm.showcard("&rst("z0")&","&rst("z1")&","&rst("z2")&","&rst("z3")&","&rst("z4")&","&rst("z5")&","&rst("u0")&","&rst("u1")&","&rst("u2")&","&rst("u3")&","&rst("u4")&","&rst("u5")&");</script>"
%>
<html>
<head>
<link rel=stylesheet href='../style.css'>
</head>
<body oncontextmenu=self.event.returnValue=false bgcolor='<%=Application("Ba_jxqy_backgroundcolor")%>' background='<%=Application("Ba_jxqy_backgroundimage")%>'>
<p align=center>
<%=msg%>
<input type=button value='����һ��' onclick=javascript:location.replace('pcbet.asp');>
</p>
</body>
</html>