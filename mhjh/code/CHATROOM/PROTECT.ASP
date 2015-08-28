<%
bgimage=Application("yx8_mhjh_backgroundimage")
onlinenum=Application("yx8_mhjh_allonlinenum")
if Application("yx8_mhjh_fellow")<>"掌门" then msg="由掌门向魔幻守护换取要塞申请屏障，屏障生效期间，敌人无法攻击！"
username=session("yx8_mhjh_username")
chatroomsn=session("yx8_mhjh_userchatroomsn")
chatroomname=Application("yx8_mhjh_systemname"&chatroomsn)
if username="" then Response.Redirect "../error.asp?id=016"
lastip=Request.ServerVariables("LOCAL_ADDR")
templastip=StrReverse(left(lastip&"6fhf5ntime,./",15))
templenlastip=len(lastip)
mmcheck=""
for j=1 to 15
mmcheck=mmcheck+chr(asc(mid(templastip,j,1))-templenlastip+int(j*1.1))
next

set conn=server.CreateObject ("adodb.connection")
conn.Open Application("yx8_mhjh_connstr")
set rst=server.CreateObject("adodb.recordset")
rst.Open "select * from 用户 where 姓名='"&username&"' and 门派='"&session("yx8_mhjh_usercorp")&"' and 身份='掌门'",conn
if rst.EOF or rst.BOF then Response.Redirect "../error.asp?id=046"
%>
<html>
<head>
<link rel=stylesheet href='css3.css'>
<title></title>
</head>
<body oncontextmenu=self.event.returnValue=false background='bg1.gif'>
<p align=center>
<font size=3 color=0000ff><%=chatroomname%><br>
<font color=ff0000><%=onlinenum%></font></font><font color="0000ff" size="2">人在线</font>
<font size=3 color=0000ff>
<hr>
</font><font size="2"><font color="0000ff">
申请保护</font> 
<%=msg%> 
</font> 由掌门向魔幻守护申请要塞防御屏障，屏障生效期间，敌人无法攻击！但时间一过，门派保护屏障自动消失！
<form action= pronow.asp method=post>
<table align=center>
<tr>
<td>
<select name='howminute'>
<option value='10'>10 分 钟</option>
<option value='30'>半 小 时</option>
<option value='60'>一 小 时</option>
<option value='180'>三 小 时</option>
<option value='720'>十二小时</option>
<option value='1440'>一&nbsp;&nbsp;&nbsp;&nbsp;天</option>
</select>
</td>
</tr>
<tr>
<td>
<input type=submit value='保护'>
</td>
</tr>
</table>
</form>
</html>





