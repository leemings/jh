<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
usercorp=Session("Ba_jxqy_usercorp")
usergrade=Session("Ba_jxqy_usergrade")
bgcolor=Application("Ba_jxqy_backgroundcolor")
bgimage=Application("Ba_jxqy_backgroundimage")
uname=Trim(Request.Form("uname"))
submit1=Trim(Request.Form("submit1"))
if instr(uname,"'")<>0 then Response.End
Set conn=Server.CreateObject("ADODB.CONNECTION")
conn.open application("Ba_jxqy_connstr")
set rst=Server.CreateObject("adodb.recordset")
rst.Open "select �ȼ� from �û� where ����='"&uname&"'",conn
if rst.EOF or rst.BOF then Response.Redirect "../error.asp?id=101"
ugrade=rst("�ȼ�")
rst.Close
set rst=nothing
if submit1="�� ��" and (usergrade>=Application("Ba_jxqy_exaltgraderight") and ugrade<5 or usergrade=Application("Ba_jxqy_allright")) and usercorp="�ٸ�" and usergrade>ugrade Then
	conn.Execute "update �û� set �ȼ�=�ȼ�+1 where ����='"&uname&"'"
	msg=uname&"�ĵȼ���Ϊ"&ugrade+1&"����"
elseif submit1="�� ��" and ((usergrade>=Application("Ba_jxqy_declinegraderight") and ugrade>0) or usergrade=Application("Ba_jxqy_allright")) and usercorp="�ٸ�"  and usergrade>ugrade then
	conn.Execute "update �û� set �ȼ�=�ȼ�-1 where ����='"&uname&"'"
	msg=uname&"�ĵȼ���Ϊ"&ugrade-1&"����"
elseif submit1="�� ��" and usercorp="�ٸ�" and usergrade=Application("Ba_jxqy_allright") and usergrade>ugrade then
	conn.Execute "update �û� set �ȼ�=1,����='��' where ����='"&uname&"'"
	msg="���"&uname&"����Ȩ����ɣ�"
elseif submit1="Ƹ ��" and usercorp="�ٸ�" and usergrade=Application("Ba_jxqy_allright") and usergrade>ugrade then
	conn.Execute "update �û� set �ȼ�=10,����='�ٸ�' where ����='"&uname&"'"
	msg="Ƹ��"&uname&"Ϊ�ٸ����ˣ���ʼ�ȼ�Ϊ10��"
else 
	msg="�޷������Ҫ��Ĳ���������������Ȩ�޲�����"
end if
conn.Close 
set conn=nothing
%>
<head>
<title><%=Application("Ba_jxqy_systemname")%></title>
</head>
<body bgcolor='<%=bgcolor%>' background='<%=bgimage%>' oncontextmenu=self.event.returnValue=false >
<p align=center><%=msg%><a href="javascript:location.replace('showgov.asp')" onmouseover="window.status='';return true;" onmouseout="window.status='';return true;">����</a></p>
</body>