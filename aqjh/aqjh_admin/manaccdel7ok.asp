<%@ LANGUAGE=VBScript codepage ="936" %><%Response.Expires=0
Response.Buffer=true
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
nowinroom=session("nowinroom")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
if session("aqjh_adminok")<>true then Response.Redirect "../chat/readonly/bomb.htm"
if aqjh_grade<>10 or instr(Application("aqjh_admin"),aqjh_name)=0  then Response.Redirect "../error.asp?id=439"
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
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")

conn.open Application("aqjh_usermdb")
sql="SELECT 姓名 FROM 用户 WHERE times=1 and CDate(lasttime)<now()-7"
rs.open sql,conn,1,1
if rs.Eof and rs.Bof then
rs.close
conn.close
set rs=nothing
set conn=nothing
Response.Redirect "manerr.asp?id=131"
end if
totalrec=rs.RecordCount
rs.close
set rs=nothing
sql="DELETE * FROM 用户 WHERE times=1 and CDate(lasttime)<now()-7 and 会员等级=0"
conn.Execute(sql)
conn.close
set conn=nothing
says="<font size=2 color=red>【清理数据】<font color=green>站长考虑到江湖运行速度,已经将7内只登陆一次的非会员用户全部删除了!</font></font>"
act="消息"
towhoway=0
towho="大家"
addwordcolor="660099"
saycolor="008888"
addsays="对"
saystr="<script>parent.sh("& chr(39) & addwordcolor & chr(39) &","& chr(39) & saycolor & chr(39) &","& chr(39) & act & chr(39) &","& chr(39) & aqjh_name & chr(39) &","& chr(39) & addsays & chr(39) &","& chr(39) & towho & chr(39) &"," & chr(39) & says & chr(39) &"," & towhoway &  ","& nowinroom & ");<"&"/script>"
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
<title>帐号管理</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href="css/css.css" type=text/css rel=stylesheet>
</head>
<body bgColor=#dcdec6 text=#000000 link=#003366 vLink=#006666>
<div align="center">
<h1><font color="0099FF">【帐号管理】</font></h1>
<font color="#FF0000">【清理７天只用过１次的帐号】</font></div>
<hr noshade size="1" color=009900>
<b>［操作完成］</b><br>
共清除只用过１次，且超过７天的帐号 <font color=red><%=totalrec%></font> 个。
<div align=center><a href="manacc.asp">返回</a></div>
<hr noshade size="1" color=009900>
<div align=center class=cp><%Response.Write "序列号：<font color=blue>" & Application("aqjh_sn") & "</font>，授权给：<font color=blue>" & Application("aqjh_user") & "</font><br><font color=999999>" & Application("aqjh_copyright") & "</font>"%></div>
</body>
</html>