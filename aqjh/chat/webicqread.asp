<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
useronlinename=Application("aqjh_useronlinename"&session("nowinroom"))
if aqjh_name="" or Session("aqjh_inthechat")<>"1" or Instr(useronlinename," "&aqjh_name&" ")=0 then Response.Redirect "chaterr.asp?id=001"
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
ping4="我的机器配置是：" & Request.ServerVariables("HTTP_USER_AGENT") & "<br>我的IP是：" & Request.ServerVariables("REMOTE_ADDR")
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
chatroombgimage=Application("aqjh_chatimage")
chatroombgcolor=Session("afa_chatbgcolor")%><html>
<head>
<title><%=aqjh_name%>，聊友在呼叫你</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="readonly/style.css">
</head>
<body oncontextmenu=self.event.returnValue=false bgcolor=<%=chatroombgcolor%> background=<%=chatroombgimage%> leftmargin="0" topmargin="0">
<table width="100%" border="0" height="100%">
<tr>
<td>
<table width=400 border=1 align=center cellspacing=0 bordercolorlight=000000 bordercolordark=FFFFFF bgcolor=E0E0E0>
<tr>
<td>
<table border=0 bgcolor=009900 cellspacing=0 cellpadding=2 width=394>
<tr>
<td width=376><font color=FFFFFF face=Wingdings>*</font><font color=FFFFFF>Web ICQ 1.50 - <%=Application("aqjh_chatroomname")%></font></td>
<td width=18>
<table border=1 bordercolorlight=666666 bordercolordark=FFFFFF cellpadding=0 bgcolor=E0E0E0 cellspacing=0 width=18>
<tr>
<td width=16><b><a href="javascript:window.close()" onMouseOver="window.status='';return true" onMouseOut="window.status='';return true" title="退出"><font color="000000">×</font></a></b></td>
</tr>
</table>
</td>
</tr>
</table>
<table width="100%" border="0" cellspacing="0" bordercolorlight="#000000" bordercolordark="#FFFFFF" bgcolor="E0E0E0" cellpadding="5">
<tr valign="middle" align="center">
<td class=p9>共收到 <font color=red><%=int(xx/4)%></font> 条信息(点击消息前的姓名即可进行回复) <font color="009900">[100秒后关闭]</font>
<table border="1" align="center" cellspacing="0" bordercolorlight="#CCCCFF" bordercolordark="#FFFFFF" width="100%">
<tr align="center" bgcolor="#CCFFCC">
<td width="66">姓　名</td>
<td width="120">时　间</td>
<td>消　息</td>
</tr>
<%for i=1 to xx step 4%><tr>
<td class=p9 width="66"><a href="#" onClick="javascript:window.open('webicq.asp?who=<%=Server.URLEncode(show(i+2))%>','hqtwebicq','Status=no,scrollbars=no,resizable=no,width=380,height=320')"><%=show(i+2)%></a></td>
<td class=p9 width="120"><%=show(i)%></td>
<td class=p9><%=show(i+3)%></td>
</tr><%next%>
</table>
</td>
</tr>
</table>
</td>
</tr>
</table>
</td>
</tr>
</table>
<bgsound src="readonly/ri.mid" loop="1">
<script Language="JavaScript">
var tid=null;tid=setTimeout('window.close()',100000);
</script>
</body>
</html>