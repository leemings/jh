<%
username=trim(Request.Form("username"))
password=trim(Request.Form("password"))
if Application("Ba_jxqy_systemname")="" then Response.Redirect "error.asp?id=001"
if username="" or password="" then Response.Redirect "error.asp?id=002"
for i=1 to len(username)
	usernamechr=mid(username,i,1)
	if asc(usernamechr)>0 then Response.Redirect "error.asp?id=003"
next
if len(password)<6 then Response.Redirect "error.asp?id=004"
for i=1 to len(password)
	passwordchr=asc(mid(password,i,1))
	if passwordchr<48 or (passwordchr>57 and passwordchr<65) or (passwordchr>90 and passwordchr<97) or passwordchr>122 then Response.Redirect "error.asp?id=005"
next
lastip=Request.ServerVariables("REMOTE_ADDR")
lastlogintime=now()
lastlogintimetype="#"&month(lastlogintime)&"/"&day(lastlogintime)&"/"&year(lastlogintime)&" "&hour(lastlogintime)&":"&minute(lastlogintime)&":"&second(lastlogintime)&"#"
set conn=server.CreateObject("adodb.connection")
conn.Open Application("Ba_jxqy_connstr")
set rst=server.CreateObject("adodb.recordset")
rst.Open "select ����,���,�Ա�,�ȼ�,����,״̬,����¼ʱ��,��Ա,��Աʱ�� from �û� where ����='"&username&"' and ����='"&password&"'",conn
if rst.EOF or rst.BOF then Response.Redirect "error.asp?id=010"
mycorp=rst("����")
myidentity=rst("���")
mysex=rst("�Ա�")
mygrade=rst("�ȼ�")
moral=rst("����")\255
predicament=rst("״̬")
if rst("��Աʱ��")>=date() and rst("��Ա")=true then
	fellow=true
else
	fellow=false
end if		
allowlogintime=rst("����¼ʱ��")
if predicament="����" and datediff("s",allowlogintime,lastlogintime)<0 then Response.Redirect "error.asp?id=018&allowtime="&allowlogintime
if predicament="����" and datediff("s",allowlogintime,lastlogintime)<0 then Response.Redirect "error.asp?id=022&allowtime="&allowlogintime
if predicament="����" and datediff("s",allowlogintime,lastlogintime)<0 then Response.Redirect "error.asp?id=022&allowtime="&allowlogintime
if predicament="��" and datediff("s",allowlogintime,lastlogintime)<0 then Response.Redirect "error.asp?id=036"
if moral>255 then moral=255
if moral<-255 then moral=-255
if moral>0 then
	mycolor=rgb(0,0,moral)
else
	mycolor=rgb(abs(moral),0,0)
end if
mycolor=cstr(hex(mycolor))
colorlen=len(mycolor)
select case colorlen
	case 1 
		mycolor="00000"&mycolor
	case 2
		mycolor="0000"&mycolor
end select			
rst.Close
set rst=nothing
conn.Execute "update �û� set ����¼IP='"&lastip&"',����¼ʱ��="&lastlogintimetype&",״̬='����',��Ա="&fellow&" where ����='"&username&"'"
conn.Close
set conn=nothing
if session("Ba_jxqy_username")<>"" then Response.Redirect "myhome.asp"
if instr(Application("Ba_jxqy_allonline"),";"&username&";")<>0 then Response.Redirect "error.asp?id=017"
Application.Lock
application("Ba_jxqy_allonline")=application("Ba_jxqy_allonline")&username&";"
session("Ba_jxqy_username")=username
session("Ba_jxqy_usersex")=mysex
session("Ba_jxqy_usercorp")=mycorp
Session("Ba_jxqy_useridentity")=myidentity
session("Ba_jxqy_usergrade")=mygrade
session("Ba_jxqy_userip")=lastip
session("Ba_jxqy_usermoral")=mycolor
session("Ba_jxqy_userchatroomsn")=1
session("Ba_jxqy_userlasttalktime")=now()
Response.Redirect "myhome.asp"
%>
