<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
username=session("Ba_jxqy_username")
mygrade=session("Ba_jxqy_usergrade")
lastsavetime=session("Ba_jxqy_userlastsavetime")
paycent=Application("Ba_jxqy_paycent")
nowtime=now()
if not isdate(lastsavetime) then lastsavetime=nowtime
expvalue=clng(datediff("n",lastsavetime,nowtime))*paycent
set conn=server.CreateObject("adodb.connection")
conn.Open Application("Ba_jxqy_connstr")
set rst=server.CreateObject("adodb.recordset")
rst.Open "select �ȼ�,����,��Ա from �û� where ����='"&username&"'",conn
if rst("����")>Application("Ba_jxqy_grade"&mygrade) and mygrade<10 then
	conn.Execute "update �û� set �ȼ�=�ȼ�+1 where ����='"&username&"'"
	session("Ba_jxqy_usergrade")=mygrade+1
end if
if rst("��Ա")=true and Application("Ba_jxqy_fellow")=true then	expvalue=expvalue*2
rst.Close
set rst=nothing
conn.Execute "update �û� set ����=����+"&expvalue&",����=����+"&expvalue&",����=����+"&expvalue&",����=����+"&expvalue&",����=����+"&expvalue*10&",����=����+"&expvalue&",����=����+"&expvalue&" where ����='"&username&"'"
conn.Close
set conn=nothing
session("Ba_jxqy_userlastsavetime")=nowtime
bgcolor=Application("Ba_jxqy_backgroundcolor")
bgimage=Application("Ba_jxqy_backgroundimage")
chatroomnum=Application("Ba_jxqy_chatroomnum")
maxnosaytime=Application("Ba_jxqy_nosaytime")
onlinelist=Application("Ba_jxqy_onlinelist")
onlinelistubd=ubound(onlinelist)
dim newonlinelist()
dim newonlinename()
newonlinenum=0
for i=1 to chatroomnum
	redim preserve newonlinename(i)
	newonlinename(i)=";"
next
j=1
for i=1 to onlinelistubd step 8
	if datediff("s",onlinelist(i+4),nowtime)<=maxnosaytime or onlinelist(i+3)="127.0.0.1" then
		newonlinenum=newonlinenum+1
		newonlinename(onlinelist(i+5))=newonlinename(onlinelist(i+5))&onlinelist(i)&";"
		redim preserve newonlinelist(j),newonlinelist(j+1),newonlinelist(j+2),newonlinelist(j+3),newonlinelist(j+4),newonlinelist(j+5),newonlinelist(j+6),newonlinelist(j+7)
		newonlinelist(j)=onlinelist(i)
		newonlinelist(j+1)=onlinelist(i+1)
		newonlinelist(j+2)=onlinelist(i+2)
		newonlinelist(j+3)=onlinelist(i+3)
		if onlinelist(i)<>username then
			newonlinelist(j+4)=onlinelist(i+4)
		else
			newonlinelist(j+4)=nowtime
		end if	
		newonlinelist(j+5)=onlinelist(i+5)
		newonlinelist(j+6)=onlinelist(i+6)
		newonlinelist(j+7)=onlinelist(i+7)
		j=j+8
	else
		Application.Lock
		Application("Ba_jxqy_allonline")=replace(Application("Ba_jxqy_allonline"),";"&onlinelist(i)&";",";")
		Application.UnLock
	end if
next
Application.Lock
if newonlinenum=0 then
	dim index(0)
	Application("Ba_jxqy_onlinelist")=index
else
	Application("Ba_jxqy_onlinelist")=newonlinelist
end if
Application("Ba_jxqy_allonlinenum")=newonlinenum
for i=1 to chatroomnum
	Application("Ba_jxqy_onlinename"&i)=newonlinename(i)
next
Application.UnLock
erase newonlinelist
erase newonlinename
Response.Write "<script language=javascript>setTimeout('this.location.href="&chr(34)&"onlinelist.asp"&chr(34)&"',10000);</script><div align=center><font size='4' color='#CC0000' face='��Բ'><b>"&Application("Ba_jxqy_systemname"&chatroomsn)&"</b></font><br><font color=008800>"&Application("Ba_jxqy_allonlinenum")&"</font>������<hr></div><body bgcolor='"&bgcolor&"' background='"&bgimage&"' oncontextmenu=self.event.returnValue=false><table  width='100%'><tr align=center><td><font color=0000FF size=2>�������</font></td></tr><tr><td align=center><font size=2 color=FF0000><b>��ǰ�ȼ���"&session("Ba_jxqy_usergrade")&"</font></td></tr><tr><td align=left><font size=2>������и��Ŭ����������������������������������֣������ֱ�����<font color=FF0000><b>"&expvalue&"</b></font>��������<font color=FF0000><b>"&expvalue*10&"</b></font></font></td></tr></table><body>"
%>
