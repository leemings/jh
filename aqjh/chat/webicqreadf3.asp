<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0

aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
useronlinename=Application("aqjh_useronlinename"&session("nowinroom"))
if aqjh_name="" or session("aqjh_roomin")="" or session("aqjh_inthechat")<>"1" or Instr(useronlinename," "&aqjh_name&" ")=0 then Response.Redirect "chaterr.asp?id=001"
n=Year(date())
y=Month(date())
r=Day(date())
s=Hour(time())
f=Minute(time())
m=Second(time())
if len(y)=1 then y="0" & y
if len(r)=1 then r="0" & r
if len(s)=1 then s="0" & s
if len(f)=1 then f="0" & f
if len(m)=1 then m="0" & m
sj=n & "-" & y & "-" & r & " " & s & ":" & f & ":" & m
wbq=Application("aqjh_webicq")
wbqub=UBound(wbq)
Dim show(),wbqnew()
xx=0
yy=0
if wbqub>0 then
for i=1 to wbqub step 4
if CStr(wbq(i+1))=CStr(aqjh_name) then
Redim Preserve show(xx+1),show(xx+2),show(xx+3),show(xx+4)
show(xx+1)=wbq(i)
show(xx+2)=wbq(i+1)
show(xx+3)=wbq(i+2)
show(xx+4)=wbq(i+3)
if LCase(show(xx+4))="ping agent" then
ping1=sj
ping2=wbq(i+2)
ping3=aqjh_name
ping4="�ҵĻ��������ǣ�" & Request.ServerVariables("HTTP_USER_AGENT") & "<br>�ҵ�IP�ǣ�" & Request.ServerVariables("REMOTE_ADDR")
pingok="1"
show(xx+4)="^_^"
end if
xx=xx+4
else
if DateDiff("n",wbq(i),sj)<=10 then
Redim Preserve wbqnew(yy+1),wbqnew(yy+2),wbqnew(yy+3),wbqnew(yy+4)
wbqnew(yy+1)=wbq(i)
wbqnew(yy+2)=wbq(i+1)
wbqnew(yy+3)=wbq(i+2)
wbqnew(yy+4)=wbq(i+3)
yy=yy+4
end if
end if
next
if pingok="1" then
Redim Preserve wbqnew(yy+1),wbqnew(yy+2),wbqnew(yy+3),wbqnew(yy+4)
wbqnew(yy+1)=ping1
wbqnew(yy+2)=ping2
wbqnew(yy+3)=ping3
wbqnew(yy+4)=ping4
yy=yy+4
end if
if yy>=4 then
wbq=wbqnew
else
Dim wbqnull(0)
wbq=wbqnull
end if
wbqub=UBound(wbq)
webicqname=""
for i=1 to wbqub step 4
webicqname=webicqname & " " & wbq(i+1)
next
webicqname=webicqname&" "
Application.Lock
Application("aqjh_webicq")=wbq
Application("aqjh_webicqname")=webicqname
Application.UnLock
end if
chatbgcolor=Session("afa_chatbgcolor")
chatimage=Application("aqjh_chatimage")
chatroomname=Application("aqjh_chatroomname")
if chatbgcolor="" then chatbgcolor="008888"%><html>
<head>
<title><%=aqjh_name%>�������ں�����</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="readonly/style.css">
</head>
<body oncontextmenu=self.event.returnValue=false bgcolor="<%=chatbgcolor%>" background="<%=chatimage%>" bgproperties="fixed">
<table width="100%" border="0" height="100%">
<tr>
<td>
<table width="100%" border="1" bgcolor="F0F0F0" bordercolorlight="#000000" bordercolordark="#FFFFFF" cellspacing="0" cellpadding="4">
<tr>
<td>
<p align=center>����<font color=red><%=int(xx/4)%></font>����Ϣ</p>
<%for i=1 to xx step 4%><p><font color=red>������</font><br><a href="#" onClick="javascript:window.open('webicq.asp?who=<%=Server.URLEncode(show(i+2))%>','hqtwebicq','Status=no,scrollbars=no,resizable=no,width=380,height=320')"><%=show(i+2)%></a></p>
<p><font color=red>ʱ�䣺</font><br><%=show(i)%></p>
<p><font color=red>��Ϣ��</font><br><%=show(i+3)%></p><%next%>
</td>
</tr>
</table>
</td>
</tr>
</table>
<Script Language=Javascript>
parent.m.location.href='about:blank';
</Script>
</body>
</html>