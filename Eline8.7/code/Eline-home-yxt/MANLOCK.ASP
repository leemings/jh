<%@ LANGUAGE=VBScript codepage ="936" %><%Response.Expires=0
Response.Buffer=true

sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
if sjjh_name="" then Response.Redirect "../error.asp?id=440"
if session("sjjh_adminok")<>true then Response.Redirect "../chat/readonly/bomb.htm"
if sjjh_grade<>10 or instr(Application("sjjh_admin"),sjjh_name)=0  then Response.Redirect "../error.asp?id=439"
inthechat=Session("sjjh_inthechat")
if inthechat<>"1"  then Response.Redirect "manerr.asp?id=211"
lockname=Trim(Request.QueryString("id"))
lockip=Trim(Request.QuerySTring("ip"))
if lockname<>"" and lockip="" then Response.Redirect "manerr.asp?id=212"
if CStr(lockname)=CStr(sjjh_name) then Response.Redirect "manerr.asp?id=213"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")

conn.open Application("sjjh_usermdb")
sql="SELECT a,b,c FROM i"
rs.open sql,conn,1,1
totalrec=rs.RecordCount
if totalrec>0 then
Dim show()
i=1
Do while Not rs.Eof
Redim Preserve show(i),show(i+1),show(i+2)
show(i)=rs("a")
show(i+1)=rs("b")
show(i+2)=rs("c")
i=i+3
rs.MoveNext
Loop
end if
rs.close
conn.close
set rs=nothing
set conn=nothing
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
t=s & ":" & f & ":" & m
sj=n & "-" & y & "-" & r & " " & t%><html>
<head>
<title>ＩＰ管理♀wWw.51eline.com♀</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="readonly/style.css">
<Script Language=JavaScript>function ulk(ip){document.forms[1].unlockip.value=ip;document.forms[1].unlockwhy.focus();}</Script>
</head>
<body oncontextmenu=self.event.returnValue=false bgcolor="#FFFFFF" class=p150>
<div align="center">
<h1><font color="0099FF">【ＩＰ管理】</font></h1>
</div>
<hr noshade size="1" color=009900>
<b>［注意事项］</b> <br>
封锁、解锁 IP 的操作会被记录在“聊务公开”栏中，供广大聊友监督！<br>
被封锁的 IP 在 <font color=red><%=Application("sjjh_iplocktime")%></font> 分钟内不能登录。之后系统会自动解锁该 IP。
<hr noshade size="1" color=009900>
<table border="1" cellpadding="5" cellspacing="0" bgcolor="E0E0E0" bordercolorlight="#000000" bordercolordark="#FFFFFF" align="center">
<form method="post" action="manlockok.asp">
<tr>
<td>
<table border="0">
<tr>
<td>封锁ＩＰ：</td>
<td>
<input type="text" name="lockip" value="<%=lockip%>" size="15" maxlength="15"<%if lockname<>"" then Response.Write " readonly"%>>
<%if lockname<>"" then Response.Write " <font color=red>(" & lockname & ")</font>"%>
<input type="hidden" name="lockname" value="<%=lockname%>">
</td>
</tr>
<tr>
<td>封锁原因：</td>
<td>(&lt;=60字符)</td>
</tr>
<tr>
<td>
<select name="select" onChange="javascript:document.forms[0].lockwhy.value=this.value;document.forms[0].lockwhy.focus();">
<option value="" selected>自填</option>
<option value="不受欢迎人。">不欢迎</option>
<option value="所取的名字十分不雅。">不雅ID</option>
<option value="乱刷屏，警告又不听。">乱刷屏</option>
<option value="在聊天室散布有悖伦理道德的言论。">不道德</option>
<option value="不遵守聊天规则，进行人身攻击。">骂人</option>
<option value="在聊天室散布违反国家法律法规的言论。">违法</option>
</select>
</td>
<td>
<input type="text" name="lockwhy" maxlength="60" size="50">
</td>
</tr>
</table>
<div align="center">
<input type="submit" name="Submit" value="封锁">
<input type="reset" name="reset" value="重写">
<input type="button" value="放弃" onclick="javascript:history.go(-1)">
</div>
</td>
</tr>
</form>
</table>
<hr noshade size="1" color=009900>
<table border="1" cellpadding="5" cellspacing="0" bgcolor="E0E0E0" bordercolorlight="#000000" bordercolordark="#FFFFFF" align="center">
<form method="post" action="manunlockok.asp">
<tr>
<td>
<table border="0">
<tr>
<td>解锁ＩＰ：</td>
<td>
<input type="text" name="unlockip" size="15" maxlength="15" readonly>
</td>
</tr>
<tr>
<td>解锁原因：</td>
<td>(&lt;=60字符)</td>
</tr>
<tr>
<td>
<select name="select" onChange="javascript:document.forms[1].unlockwhy.value=this.value;document.forms[1].unlockwhy.focus();">
<option value="" selected>自填</option>
<option value="封锁时不小心搞错了。">误操作</option>
<option value="对方已经改过，请求解锁IP。">悔过</option>
<option value="这个IP是网吧的，封锁后都上不来了。">网吧</option>
<option value="这个IP是代理地址，封锁后都上不来了。">代理</option>
</select>
</td>
<td>
<input type="text" name="unlockwhy" maxlength="60" size="50">
</td>
</tr>
</table>
<div align="center">
<input type="submit" name="Submit" value="解锁">
<input type="reset" name="reset" value="重写">
<input type="button" value="放弃" onclick="javascript:history.go(-1)">
</div>
</td>
</tr>
</form>
</table>
<hr noshade size="1" color=009900>
<div align=center>共 <font color=red><%=totalrec%></font> 个　<a href=javascript:history.go(0)>刷新</a></div>
<table border="1" cellpadding="4" cellspacing="0" bordercolorlight="#999999" bordercolordark="#FFFFFF" bgcolor="E0E0E0" align="center">
<tr bgcolor="#0099FF">
<td><font color="#FFFFFF">序号</font></td>
<td><font color="#FFFFFF">封 锁 者</font></td>
<td><font color="#FFFFFF">封 锁 时 间</font></td>
<td><font color="#FFFFFF">被封锁的IP地址</font></td>
<td><font color="#FFFFFF">已封锁时长</font></td>
<td><font color="#FFFFFF">立即解锁</font></td>
</tr><%if totalrec>0 then
for i=1 to UBound(show) step 3%><tr>
<td><%=(i+2)/3%></td>
<td><%=show(i+2)%></td>
<td><%=show(i+1)%></td>
<td><%=show(i)%></td>
<td><%=DateDiff("n",show(i+1),sj)&" "%>分钟</td>
<td><a href="javascript:ulk('<%=show(i)%>')">Ｕ</a></td>
</tr><%next
end if%></table>
<hr noshade size="1" color=009900>
<div align=center class=cp><%Response.Write "序列号：<font color=blue>" & Application("sjjh_sn") & "</font>，授权给：<font color=blue>" & Application("sjjh_user") & "</font><br><font color=999999>" & Application("sjjh_copyright") & "</font>"%></div>
</body>
</html>