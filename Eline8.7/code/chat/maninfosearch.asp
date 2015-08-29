<%@ LANGUAGE=VBScript codepage ="936" %><%Response.Expires=0

sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
nowinroom=session("nowinroom")
if sjjh_name="" then Response.Redirect "../error.asp?id=440"
grade=Int(sjjh_grade)
id=Trim(Request.QueryString("id"))
id2=LCase(id)
if sjjh_name="" then Response.Redirect "../error.asp?id=420"
if grade<7 then Response.Redirect "../error.asp?id=482"
if id2="" then Response.Redirect "../error.asp?id=483"
'total=Application("sjjh_chatrs"&nowinroom)
onlinelist=Application("sjjh_onlinelist"&nowinroom)
ubo=UBound(onlinelist)
dim show()
js=1
for i=1 to ubo step 6
if Instr(LCase(onlinelist(i+1)),id2)<>0 or Instr(onlinelist(i+2),id2)<>0 then
Redim Preserve show(js),show(js+1),show(js+2),show(js+3)
show(js)=onlinelist(i+1)
show(js+1)=onlinelist(i+2)
show(js+2)=onlinelist(i+4)
show(js+3)=onlinelist(i+5)
js=js+4
end if
next
set onlinelist=nothing
totalrec=Int((js-1)/4)
p=1
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
sj=n & "-" & y & "-" & r & " " & s & ":" & f & ":" & m%><html>
<head>
<title>聊友信息查找</title>
<META http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="readonly/style.css">
<script language=javascript>window.moveTo(100,50);window.resizeTo(screen.availWidth*2/3,screen.availHeight*3/4);</script>
</head>
<body oncontextmenu=self.event.returnValue=false bgcolor="#FFFFFF" class=p150>
<div align="center">
<h1><font color="#000000" face="隶书" size="6">炸 人 操 作</font></h1>
<a href="javascript:history.go(0)">刷新</a> </div>
<hr noshade size="1" color=009900 width="70%">
<div align="center"><b>〔在线人员〕</b>　共 <font color=red><%=total%></font> 人，找到包含 <font color=red><%=id%></font>
的用户名 <font color=red><%=totalrec%></font> 个。 </div>
<hr noshade size="1" color=009900 width="70%">
<div align=center>本页刷新时间：<font color="#FF0000"><%=sj%></font><br></div>
<table border="1" cellspacing="0" bordercolorlight="#999999" bordercolordark="#FFFFFF" cellpadding="4" bgcolor="E0E0E0" align="center" width="416">
<tr bgcolor="#000000" align="center">
<td><font color="#FFFFFF">序号</font></td>
<td><font color="#FFFFFF">用 户 昵 称</font></td>
<td><font color="#FFFFFF">性别</font></td>
<td><font color="#FFFFFF">进 入 时 间</font></td>
<td><font color="#FFFFFF">未存点</font></td>
<td><font color="#FFFFFF">炸弹</font></td>
</tr>
<%if grade<3 then
if totalrec>0 then
j=0
for i=1 to UBound(show) step 4%> <%next
end if
end if
if grade=3 then
if totalrec>0 then
j=0
for i=1 to UBound(show) step 4%> <%next
end if
end if
if grade=4 then
if totalrec>0 then
j=0
for i=1 to UBound(show) step 4%> <%next
end if
end if
if grade=5 then
if totalrec>0 then
j=0
for i=1 to UBound(show) step 4%> <%next
end if
end if
if grade=6 then
if totalrec>0 then
j=0
for i=1 to UBound(show) step 4%> <%next
end if
end if
if grade>6 then
if totalrec>0 then
j=0
for i=1 to UBound(show) step 4%>
<tr align="center">
<td><%Response.Write p+j
j=j+1%></td>
<td><%=show(i)%></td>
<td><%=show(i+1)%></td>
<td><%=show(i+2)%></td>
<td><%=DateDiff("n",show(i+3),sj)%> 分</td>
<td><a href="manbomb.asp?id=<%=server.URLEncode(show(i))%>&ip=<%=show(i+1)%>">轰炸</a></td>
</tr>
<%next
end if
end if%>
</table>
<div align="center"><br>
注意：此功能只用于那些经常捣乱却不悔改的聊友!<br>
此记录将记载于数据库中，任何网管瞎炸人的都将会受到处罚 </div>
<hr noshade size="1" color=009900 width="70%">
<div align=center class=cp><%Response.Write "序列号：<font color=blue>" & Application("sjjh_sn") & "</font>，授权给：<font color=blue>" & Application("sjjh_user") & "</font><br><font color=999999>" & Application("sjjh_copyright") & "</font>"%></div>
</body>
</html>