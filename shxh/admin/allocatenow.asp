<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
bgcolor=Application("Ba_jxqy_backgroundcolor")
bgimage=Application("Ba_jxqy_backgroundimage")
username=session("Ba_jxqy_username")
usercorp=session("Ba_jxqy_usercorp")
usergrade=session("Ba_jxqy_usergrade")
stock=Trim(Request.Form("stock"))
money=Request.Form("money")
if username="" then Response.Redirect "../error.asp?id=016"
if not(usergrade=Application("Ba_jxqy_allright") and usercorp="�ٸ�") then Response.Redirect "../error.asp?id=046"
set conn=server.CreateObject("adodb.connection")
conn.Open Application("Ba_jxqy_connstr")
set rst=server.CreateObject("adodb.recordset")
rst.Open "select * from �ֹ� where ����='"&stock&"'",conn
do while not (rst.EOF or rst.BOF)
	conn.Execute "update �û� set ���=���+"&rst("��Ȩ")*money&" where ����='"&rst("�ֹ���")&"'"
	rst.MoveNext
loop
rst.Close
set rst=nothing
conn.Execute "insert into ֤������(ʱ��,����,����) values(now(),'"&stock&"���ֽ�ֺ�','"&stock&"��Ӫ���������ʾӮ��,���»�����ֽ�ֺ�,ÿ����Ϣ"&money&"')"
conn.Close
set conn=nothing
Response.Write "<head><title>"&Application("Ba_jxqy_systemname")&"</title><script language=javascript>setTimeout('history.go(-1);',3000);</script><link rel='stylesheet' href='../chatroom/css.css'></head><body oncontextmenu=self.event.returnValue=false  background='"&bgimage&"' bgcolor='"&bgcolor&"'><p align=center><font color=0000ff size=4>��Ϣ�ֺ�</font><br>��Ʊ�ֺ����.�����Ӻ��Զ�����<br><a href='javascript:history.go(-1);'>����</a></p>"
%>