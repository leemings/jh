<%
username=session("Ba_jxqy_username")
usersex=session("Ba_jxqy_usersex")
if username="" then Response.Redirect "../error.asp?id=016"
if usersex="��" then 
	matesex="Ů"
else
	matesex="��"
end if
quest=server.HTMLEncode(Trim(Request.Form("quest")))
content=server.HTMLEncode(Trim(Request.Form("content")))
opt=server.HTMLEncode(Trim(Request.Form("opt")))
if content="" then Response.Redirect "../error.asp?id=043"
if instr(quest,"'")<>0  then Response.Redirect "../error.asp?id=900"
nowtime=now()
nowtimetype="#"&month(nowtime)&"/"&day(nowtime)&"/"&year(nowtime)&" "&hour(nowtime)&":"&minute(nowtime)&":"&second(nowtime)&"#"
bgcolor=Application("Ba_jxqy_backgroundcolor")
bgimage=Application("Ba_jxqy_backgroundimage")
set conn=server.createobject("adodb.connection") 
conn.open Application("Ba_jxqy_connstr")
set rst=server.CreateObject ("adodb.recordset")
rst.open "select ��ż from �û� where ����='"&username&"'",conn
mate=rst("��ż")
if opt="���" then
	if mate<>"��" then Response.Redirect "../error.asp?id=038"
	rst.Close
	rst.Open "select ���� from ý�� where ������='"&username&"'",conn
	if not (rst.EOF or rst.BOF)then
		if rst("����")<>quest then 
			Response.Redirect "../error.asp?ID=039"
		else
			Response.Redirect "../error.asp?id=040"
		end if	
	end if
	rst.Close 
	rst.Open "select * from �û� where ����='"&quest&"' and �Ա�='"&matesex&"'",conn
	if rst.EOF or rst.BOF then Response.Redirect "../error.asp?id=041"
	sqlstr="insert into ý��(������,����,˵��,ʱ��,����) values('"&username&"','"&quest&"','"&content&"',"&nowtimetype&",False)"
else
	if mate="��" then Response.Redirect "../error.asp?ID=042"
	sqlstr="select * from ý�� where ������='"&username&"'"
	rst.Close
	rst.Open sqlstr,conn
	if not(rst.EOF or rst.BOF) then Response.Redirect "../error.asp?id=042"
	sqlstr="insert into ý��(������,����,˵��,ʱ��,����) values('"&username&"','"&mate&"','"&content&"',"&nowtimetype&",True)"
end if
rst.Close
set rst=nothing
conn.Execute sqlstr
conn.Close 
set conn=nothing
%>
<head><title><%=Application("Ba_jxqy_systemname")%></title><LINK href="../chatroom/css.css" rel=stylesheet><script language=javascript>setTimeout('history.back();',3000);</script></head>
<body bgcolor='<%=bgcolor%>' background='<%=bgimage%>' oncontextmenu=self.event.returnValue=false topMargin=200>
<div align=center>3���Ӻ��Զ�����<br>
  <font color=FF0000><img src="../images/error.gif" width="32" height="32"> ����������ɣ�лл���Ĺ��</font></div>
<p align=right><a href="#" onclick=javascript:history.back(); onmouseover="window.status='����';return true;" onmouseout="window.status='';return true;">����</a></p>
</body>