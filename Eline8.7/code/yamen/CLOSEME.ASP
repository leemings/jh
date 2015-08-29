<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="../pass.asp"-->
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
myid=Request.form("id")
name=trim(request.form("name"))
pass=trim(request.form("pass"))
for each element in Request.Form
if instr(element,"'")<>0 or instr(element,"|")<>0 or instr(element," ")<>0 or instr(Request.Form(element),"'")<>0 or instr(Request.Form(element)," ")<>0 or instr(Request.Form(element),"|")<>0 then 
		Response.Write "<script Language=Javascript>alert('提示：输入数据有问题，请查看！');location.href = 'javascript:history.go(-1)';</script>"
		Response.End
end if
next
if name="" or pass="" or myid="" then
	Response.Write "<script Language=Javascript>alert('提示：是不是想开玩了？连大名和口令都不报？！');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
sjjh_roominfo=split(Application("sjjh_room"),";")
chatroomnum=ubound(sjjh_roominfo)-1
for i=0 to chatroomnum	
	ydl=1
	if Instr(LCase(Application("sjjh_useronlinename"&i))," "&LCase(name)&" ")=0 then ydl=0
	if ydl=1 then exit for
next 
if ydl=0 then
	Response.Write "<script Language=Javascript>alert('提示：你在搞什么呀？你并没有卡在江湖里！看看是不是选择错了！');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
inroom=i
pass=md5(pass)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "select id,密码 from 用户 where id="& myid &" and 姓名='" & name & "'",conn
if rs.eof or rs.bof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：您输入的名字我们找不到!\n请查看你的id、姓名、密码！');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
id=rs("id")
if trim(pass)<>rs("密码") then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：密码不对呀，让我怎么办！');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
rs.close
set rs=nothing
conn.close
set conn=nothing
Response.Write "<script Language=Javascript>alert('提示：您已经通过了我们的认正！');</script>"
'sjjh_name=name
userip=Request.ServerVariables("REMOTE_ADDR")
'kickname=name
'kickwhy="我不小线掉线了，没办法！"
s=Hour(time())
f=Minute(time())
m=Second(time())
if len(s)=1 then s="0" & s
if len(f)=1 then f="0" & f
if len(m)=1 then m="0" & m
t=s & ":" & f & ":" & m
'*************************
Application.Lock
onlinelist=Application("sjjh_onlinelist"&inroom)
onno=ubound(onlinelist)
for i=1 to onno
if InStr(onlinelist(i),name & "|")=1 then
  for j=i to onno-1
   onlinelist(j)=onlinelist(j+1)
  next
  ReDim Preserve onlinelist(onno-1)
  exit for
 end if
next
Application("sjjh_onlinelist"&inroom)=onlinelist
Application("sjjh_useronlinename"&inroom)=Replace(Application("sjjh_useronlinename"&inroom)," " & name & " ","")
Application.UnLock
'****************
says="<font color=black>【掉线】</font><font color=blue>江湖ID["& id &"]大侠自己踢自己的一脚，为什么？掉线了……</font><font class=t>(" & t & ")</font>"			'聊天数据
says=replace(says,"'","\'")
says=replace(says,chr(34),"\"&chr(34))
act="消息"
towhoway=0
towho="大家"
addwordcolor="660099"
saycolor="008888"
addsays="对"
saystr="<script>parent.sh("& chr(39) & addwordcolor & chr(39) &","& chr(39) & saycolor & chr(39) &","& chr(39) & act & chr(39) &","& chr(39) & "系统机器" & chr(39) &","& chr(39) & addsays & chr(39) &","& chr(39) & towho & chr(39) &"," & chr(39) & says & chr(39) &"," & towhoway &  ",0);<"&"/script>"
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

%>
<html>
<head>
<title>掉线管理</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href="../css.css" rel=stylesheet>
</head>
<body bgcolor="#006699" class=p150>
<div align="center">
  <table width="360" border="1" cellpadding="10" cellspacing="13" bgcolor="#000000">
    <tr bgcolor="#FFFFFF">
      <td bgcolor="#006699"> 
        <div align="center"><font color="#FF6600">【 掉 线 管 理 】</font> </div>
<b><br>
        <font color="#FFFFFF">［操作完成］</font></b> <br>
        <font color="#FFFFFF">现在你要以点返回进行登陆了，下次一定要注意使用！一般10分钟不说话会掉线的，要注意！ </font><br>
<br>
<table border="1" cellspacing="0" cellpadding="10" bordercolorlight="ff6600" bordercolordark="#FFFFFF" align="center">
<form>
<tr>
<td> <%=now()%><br>
<%=name&"("&userip&")"%><br>
                <font color="#FFFFFF">现在你已经可以使用了！</font> 
                <div align="center">
<input type="button" value="关闭窗口" onClick="window.close()"
name="button">
</div>
</td>
</tr>
</form>
</table>
<div align=center><%Response.Write "序列号：" & Application("sjjh_sn") & "，授权给：" & Application("sjjh_user") & "<br>" & Application("sjjh_copyright") & ""%></div>
</td>
</tr>
</table>
</div>
</body>
</html>