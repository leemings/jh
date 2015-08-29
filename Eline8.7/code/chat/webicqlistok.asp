<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0

sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
useronlinename=Application("sjjh_useronlinename"&session("nowinroom"))
if sjjh_name="" or Session("sjjh_inthechat")<>"1" or Instr(useronlinename," "&sjjh_name&" ")=0 then Response.Redirect "chaterr.asp?id=001"
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
wbq=Application("sjjh_webicq")
wbqub=UBound(wbq)
Dim show(),wbqnew()
xx=0
yy=0
if wbqub>0 then
for i=1 to wbqub step 4
if DateDiff("n",wbq(i),sj)<=10 then
Redim Preserve wbqnew(yy+1),wbqnew(yy+2),wbqnew(yy+3),wbqnew(yy+4),show(xx+1),show(xx+2),show(xx+3)
wbqnew(yy+1)=wbq(i)
wbqnew(yy+2)=wbq(i+1)
wbqnew(yy+3)=wbq(i+2)
wbqnew(yy+4)=wbq(i+3)
yy=yy+4
show(xx+1)=wbq(i+2)
show(xx+2)=wbq(i+1)
show(xx+3)=wbq(i)
xx=xx+3
end if
next
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
Application("sjjh_webicq")=wbq
Application("sjjh_webicqname")=webicqname
Application.UnLock
end if
chatroombgimage=Application("sjjh_chatimage")
chatroombgcolor=Application("sjjh_chatbgcolor")%><html>
<head>
<title>消息列表</title>
<META http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="readonly/style.css">
</head>
<body oncontextmenu=self.event.returnValue=false bgcolor="#FFFFFF">
<div align="center">共有 <font color=red><%=int((xx)/3)%></font> 条信息未被领取　<a href=javascript:history.go(0)>【刷新】</a></div>
<table border="1" align="center" cellspacing="0" bordercolorlight="#CCCCFF" bordercolordark="#FFFFFF" width="100%">
<tr align="center" bgcolor="#CCFFCC">
<td>序</td>
<td width="66">发送者</td>
<td width="66">接收者</td>
<td>时　间</td>
</tr><%for i=1 to xx step 3%>
<tr>
<td class=p9><%=int((i+2)/3)%></td>
<td class=p9 width="66"><%=show(i)%></td>
<td class=p9 width="66"><%=show(i+1)%></td>
<td class=p9><%=show(i+2)%></td>
</tr><%next%>
</table>
</body>
</html>