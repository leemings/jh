<%@ LANGUAGE=VBScript codepage ="936" %><%Response.Expires=0
Response.Buffer=true

sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
if sjjh_name="" then Response.Redirect "../error.asp?id=440"
if session("sjjh_adminok")<>true then Response.Redirect "../chat/readonly/bomb.htm"
if sjjh_grade<>10 or instr(Application("sjjh_admin"),sjjh_name)=0  then Response.Redirect "../error.asp?id=439"
inthechat=Session("sjjh_inthechat")
userip=Request.ServerVariables("REMOTE_ADDR")
if inthechat<>"1"  then Response.Redirect "manerr.asp?id=211"
lockname=Server.HTMLEncode(Trim(Request.Form("lockname")))
lockip=Server.HTMLEncode(Trim(Request.Form("lockip")))
if lockip="" then Response.Redirect "manerr.asp?id=212"
lockwhy=Server.HTMLEncode(Trim(Request.Form("lockwhy")))
if lockwhy="" then Response.Redirect "manerr.asp?id=214"
if CStr(lockname)=CStr(sjjh_name) then Response.Redirect "manerr.asp?id=213"
if len(lockwhy)>60 then lockwhy=Left(lockwhy,60)
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
sj=n & "-" & y & "-" & r & " " & t
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
sql="SELECT a FROM i WHERE a='" & lockip & "'"
rs.open sql,conn,1,1
if Not(rs.Eof and rs.Bof) then
	rs.close
	conn.close
	set rs=nothing
	set conn=nothing
	Response.Redirect "manerr.asp?id=215"
end if
rs.close
Application.UnLock
'sql="SELECT lastkick FROM 用户 WHERE and 姓名='" & lockname & "'"
'rs.open sql,conn,1,3
'if Not(rs.Eof and rs.Bof) then
'rs("lastkick")=sj
'rs.Update
'end if
'rs.close
'end if
'set rs=nothing
Function SqlStr(data)
SqlStr="'" & Replace(data,"'","''") & "'"
End Function
sql="INSERT INTO i (a,b,c) VALUES ("
sql=sql & SqlStr(lockip) & ","
sql=sql & SqlStr(sj) & ","
sql=sql & SqlStr(sjjh_name) & ")"
conn.Execute sql
conn.execute "insert into l(a,b,c,d,e) values (now(),'"& sjjh_name &"','"& Request.ServerVariables("REMOTE_ADDR") &"','管理记录','封ip30分钟："&lockip&"')"
locklog="封锁IP：<font color=red>" & lockip & "</font>(<font color=blue>" & lockname & "</font>)！<br><font color=009900>【原因：" & lockwhy & "】</font>"
conn.close
set conn=nothing
Session("sjjh_lasttime")=sj
says="<font color=black>【封锁】</font><font color=8800FF><font color=red>" & sjjh_name & "</font>封锁IP：<font color=red>" & lockip & "</font>(" & lockname & ")……〖原因：" & lockwhy & "〗</font><font class=t></font>"			'聊天数据
says=replace(says,"'","\'")
says=replace(says,chr(34),"\"&chr(34))
act="消息"
towhoway=0
towho="大家"
addwordcolor="660099"
saycolor="008888"
addsays="对"
saystr="<script>parent.sh("& chr(39) & addwordcolor & chr(39) &","& chr(39) & saycolor & chr(39) &","& chr(39) & act & chr(39) &","& chr(39) & sjjh_name & chr(39) &","& chr(39) & addsays & chr(39) &","& chr(39) & towho & chr(39) &"," & chr(39) & says & chr(39) &"," & towhoway &  ","& nowinroom & ");<"&"/script>"
addmsg saystr
Function Yushu(a)
	Yushu=(a and 31)
End Function
Sub AddMsg(Str)
Application.Lock()
Application("SayCount")=Application("SayCount")+1
i="SayStr"&YuShu(Application("SayCount"))
Application(i)=Str
Application.UnLock()
End Sub
%><html>
<head>
<title>封锁ＩＰ</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="readonly/style.css">
</head>
<body oncontextmenu=self.event.returnValue=false bgcolor="#FFFFFF" class=p150>
<div align="center">
<h1><font color="0099FF">【封锁ＩＰ】</font></h1>
</div>
<hr noshade size="1" color=009900>
<b>［操作完成］</b> <br>
被封锁的 IP 在 <font color=red><%=Application("sjjh_iplocktime")%></font> 分钟内不能登录。你刚才的操作没有被记录在公开栏中。
<hr noshade size="1" color=009900>
<br>
<table border="1" cellspacing="0" cellpadding="10" bgcolor="E0E0E0" bordercolorlight="#000000" bordercolordark="#FFFFFF" align="center">
<form>
<tr>
<td>
<p><%=sj%></p>
<p><%=sjjh_name&"("&userip&")"%></p>
<p><%=locklog%></p>
<div align="center">
<input type="button" value="返回" onclick="javascript:history.go(-2)">
</div>
</td>
</tr>
</form>
</table>
<br>
<hr noshade size="1" color=009900>
<div align=center class=cp><%Response.Write "序列号：<font color=blue>" & Application("sjjh_sn") & "</font>，授权给：<font color=blue>" & Application("sjjh_user") & "</font><br><font color=999999>" & Application("sjjh_copyright") & "</font>"%></div>
</body>
</html>