<!--#include file="pass.asp"--><%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
username=trim(Request.Form("username"))
password=trim(Request.Form("password"))
say=Request.form("say")
say=trim(say)
say=server.HTMLEncode(say)
if username="" or password="" then Response.Redirect "error.asp?id=002"
if len(username)<2 or len(password)<6 or len(say)>50 then Response.Redirect "error.asp?id=004"
if len(username)>7 or len(password)>15  then Response.Redirect "error.asp?id=004"
if Request.form("say")="" then response.redirect "error.asp?id=484"
if InStr(pass,"=")<>0 or InStr(pass,"`")<>0 or InStr(pass,"'")<>0 or InStr(pass," ")<>0 or InStr(pass,"��")<>0 or InStr(pass,"'")<>0 or InStr(pass,chr(34))<>0 or InStr(pass,"\")<>0 or InStr(pass,",")<>0 or InStr(pass,"<")<>0 or InStr(pass,">")<>0 then Response.Redirect "error.asp?id=524"
if InStr(say,"=")<>0 or InStr(say,"`")<>0 or InStr(say,"'")<>0 or InStr(say," ")<>0 or InStr(say,"��")<>0 or InStr(say,"'")<>0 or InStr(say,chr(34))<>0 or InStr(say,"\")<>0 or InStr(say,",")<>0 or InStr(say,"<")<>0 or InStr(say,">")<>0 then Response.Redirect "error.asp?id=524"
for i=1 to len(username)
usernamechr=mid(username,i,1)
if asc(usernamechr)>0 then Response.Redirect "error.asp?id=003"
next
for i=1 to len(password)
passwordchr=asc(mid(password,i,1))
if passwordchr<48 or (passwordchr>57 and passwordchr<65) or (passwordchr>90 and passwordchr<97) or passwordchr>122 then Response.Redirect "error.asp?id=005"
next
password=md5(password)
set conn=server.CreateObject("adodb.connection")
conn.Open Application("yx8_mhjh_connstr")
set rst=server.CreateObject("adodb.recordset")
rst.Open "select ��ż from �û� where ����='"&username&"' and ����='"&password&"'",conn
if rst.EOF or rst.BOF then Response.Redirect "error.asp?id=019"
mate=rst("��ż")
rst.Close
set rst=nothing
if mate<>"��" then conn.Execute "update �û� set ��ż='��' where ����='"&mate&"'"
conn.Execute "delete * from �û� where ����='"&username&"'"
conn.Execute "delete * from ��Ʒ where ������='"&username&"'"
conn.Execute "delete * from ���� where ����='"&username&"'"%>
<!--#include file="data.asp"-->
<%sqla="delete * from ý�� where ������='"&username&"'"
 Set Rs=connt.Execute(sqla)%>
<!--#include file="data1.asp"-->
<%sqlb="update ���� set ����='��',����='��',����=0,����=0,��������=0 where ����='"&username&"' or ����='"&username&"'"
 Set Rs=conntt.Execute(sqlb)
mess="["&username & "]��"&say&"��"
dim newtalkarr(600) 
   talkarr=Application("yx8_mhjh_talkarr") 
		talkpoint=Application("yx8_mhjh_talkpoint") 
		j=1 
		for i=11 to 600 step 10 
			newtalkarr(j)=talkarr(i) 
			newtalkarr(j+1)=talkarr(i+1) 
			newtalkarr(j+2)=talkarr(i+2) 
			newtalkarr(j+3)=talkarr(i+3) 
			newtalkarr(j+4)=talkarr(i+4) 
			newtalkarr(j+5)=talkarr(i+5) 
			newtalkarr(j+6)=talkarr(i+6) 
			newtalkarr(j+7)=talkarr(i+7) 
			newtalkarr(j+8)=talkarr(i+8) 
			newtalkarr(j+9)=talkarr(i+9) 
			j=j+10 
		next 
		newtalkarr(591)=talkpoint+1 
		newtalkarr(592)=1 
		newtalkarr(593)=0 
		newtalkarr(594)=username 
		newtalkarr(595)="���" 
		newtalkarr(596)="" 
		newtalkarr(597)="000000" 
		newtalkarr(598)="000000" 
		newtalkarr(599)="<font color=red>����ɱ��</font><b><font color=red>"& mess &"</font></b>" 
newtalkarr(600)=Session("yx8_mhjh_userchatroomsn")
Application.Lock
Application("yx8_mhjh_talkpoint")=talkpoint+1
Application("yx8_mhjh_talkarr")=newtalkarr
Application.UnLock
erase newtalkarr
erase talkarr
session.Abandon
conn.close
set conn=nothing
Response.Write "<head><title>"&Application("yx8_mhjh_systemname")&"</title><link rel='stylesheet' href='../style.css'></head><body background='../chatroom/bg1.gif'oncontextmenu=self.event.returnValue=false><p align=center><font color=ffffff size=5>�� �� ·</font><br>��ɱ��ɣ�����ʹ�죡<br>16�������һ���ú�����������<br><input type=button onclick=javascript:top.window.close(); value=' �� �� '></p></body>"
%>
