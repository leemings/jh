<%
Response.Expires=-1
utime=Request.Form("utime")
if not isnumeric(utime) then Response.Redirect "../error.asp?id=034"
if utime<1 then Response.Redirect "../error.asp?id=035"
bgcolor=Application("Ba_jxqy_backgroundcolor")
bgimage=Application("Ba_jxqy_backgroundimage")
username=session("Ba_jxqy_username")
if username="" then Response.Redirect "../error.asp?id=016"
set conn=server.CreateObject("adodb.connection")
conn.Open Application("Ba_jxqy_connstr")
set rst=server.CreateObject("adodb.recordset")
rst.Open "select * from �û� where ����='"&username&"'",conn
if rst.EOF or rst.BOF then session.Abandon
umoney=rst("����")
rst.Close
set rst=nothing
bill=utime*60
etime=dateadd("h",utime,now())
if umoney<bill then Response.Redirect "../error.asp?id=030"
conn.Execute "update �û� set ����¼ʱ��='"&etime&"',״̬='��',����=����+"&bill&",����=����-"&bill&" where ����='"&username&"'"
conn.Close
set conn=nothing
session.Abandon
%>
<head><title>������ջ</title><link rel="stylesheet" href="../chatroom/css.css"></head>
<body bgcolor='<%=bgcolor%>' background='<%=bgimage%>' oncontextmenu=self.event.returnValue=false topmargin=100>
<p align=center><font color=0000FF size=6 face='��Բ'><%=Application("Ba_jxqy_systemname")%>������ջ</font></p>
<p><font color=4>�л�����<%=bill%>�����ǳ���л�����ǽ�����<%=etime%>֮ǰ��ֹ����ʹ������ʺţ���ӭ������</font></p>
<p align=right><a href="#" onclick=javascript:window.close(); onmouseover="window.status='�ر�';return true;" onmouseout="window.status='';return true;">�ر�</a></p>
</body>
