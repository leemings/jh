<%
bgcolor=Application("Ba_jxqy_backgroundcolor")
bgimage=Application("Ba_jxqy_backgroundimage")
username=session("Ba_jxqy_username")
usercorp=session("Ba_jxqy_usercorp")
usergrade=session("Ba_jxqy_usergrade")
if username="" then Response.Redirect "../error.asp?id=016"
if not(usergrade=Application("Ba_jxqy_allright") and usercorp="�ٸ�") then Response.Redirect "../error.asp?id=046"
howday=Request.Form("howday")
if not isnumeric(howday) then Response.Redirect "../error.asp?id=024"
howday=clng(howday)
nowtime=now()
abatetime=dateadd("d",-howday,nowtime)
abatetimetype="#"&month(abatetime)&"/"&day(abatetime)&"/"&year(abatetime)&" "&hour(abatetime)&":"&minute(abatetime)&":"&second(abatetime)&"#"
set conn=server.CreateObject("adodb.connection")
conn.Open Application("Ba_jxqy_connstr")
conn.Execute "delete  from �û� where ����¼ʱ��<"&abatetimetype
conn.Close
set conn=nothing
%>
<html>
<head>
<link rel=stylesheet href='../chatroom/css.css'>
</head>
<body oncontextmenu=self.event.returnValue=false background='<%=Application("Ba_jxqy_backgroundimage")%>' bgcolor='<%=Application("Ba_jxqy_backgroundcolor")%>'>
ҵ��ɾ��<%=formatdatetime(abatetime,1)%>֮��ĩʹ��֮�ʺţ��������ݿռ䲻��վ���ͷţ��������������ݿ�(Dafault:system/mxcz.asp)��ʹ��access2000ѹ���������Ч��<br>�밴���ؼ�������һҳ<p align=center><input type=button onclick=javascript:location.href='selectuser.asp' value='����'></p>
</body>
</html>