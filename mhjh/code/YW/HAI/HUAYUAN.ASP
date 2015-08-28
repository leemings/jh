
<!--#include file="func213.asp"-->
<!--#include file="sub.asp" -->
<%
Response.Buffer=true
sh=request.form("sh")
sy=request.form("sy")
if request.form("h")="1" then
my=session("yx8_mhjh_username")
if session("yx8_mhjh_username")="" then Response.Redirect "../../error.asp?id=016"
if session("advtime")<>"" then
if session("advtime")>now()-0.0001 then  Response.Redirect "../../error.asp?id=490"
end if
session("advtime")=now()
set conn=server.createobject("adodb.connection") 
conn.open Application("yx8_mhjh_connstr")
set rs=server.CreateObject ("adodb.recordset")
sql="select 体力,任务,任务时间 from 用户 where 姓名='" & my & "'"
set rs=conn.execute(sql)
tl=rs("体力")
cz=rs("任务时间")
if tl<0 or rs("任务")<>"孤岛探宝" then 
	Response.Write "<script Language=Javascript>alert('提示：你由于体力已经不够或者任务已经改变,你将马上离开孤岛！');</script>"
	Response.Write "<script Language=Javascript>location.href ='javascript:window.close()';</script>"
	Response.end
end if
rwtm=DateDiff("s",""&cz&"",now()) 
if rwtm>600 then
sql="update 用户 set 任务='无' where 姓名='" & my & "'"
conn.execute sql
	Response.Write "<script Language=Javascript>alert('提示：你本次孤岛冒险所携带的风铃已经失去效力，你的次元真气将自动送你安全返回！');</script>"
	Response.Write "<script Language=Javascript>location.href ='javascript:window.close()';</script>"
	Response.end
end if
conn.Close
Set conn=Nothing
message=huayuan(my,sh,sy)
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
		newtalkarr(595)="大家" 
		newtalkarr(596)="" 
		newtalkarr(597)="000000" 
		newtalkarr(598)="000000" 
		newtalkarr(599)="<font color=red>【孤岛探宝】</font><font color=blue>"&message&"</font>" 
newtalkarr(600)=Session("yx8_mhjh_userchatroomsn")
Application.Lock
Application("yx8_mhjh_talkpoint")=talkpoint+1
Application("yx8_mhjh_talkarr")=newtalkarr
Application.UnLock
erase newtalkarr
erase talkarr
%>
<html>
<head>
<LINK href="../../style.css" rel=stylesheet></head>
<body background='../../chatroom/bg1.gif' oncontextmenu=self.event.returnValue=false leftmargin="5" marginwidth="5">
<center>
<%
if tl<20 then
%>
<script language=vbscript>
MsgBox "你疲劳程度已超出范围，为防不测，还是离开孤岛为上！"
location.href = "javascript:history.back()"
</script>
<%
else
message1=message

end if
end if%>
<table border=1 align=center width=350 cellpadding="0" cellspacing="1" bordercolor="#000000" bgcolor="#FFFFFF">
<tr>
<td>
<table height="225" align="center" width="100%">
<tr>
<td height="37">
<div align="center"><strong><font color="#FF0000">突发事件</font></strong></div>

<tr>
<td height="147" valign="top">
  <p align="center">  <img border="0" src="../../IMAGE/SPINSTAR.GIF" width="47" height="47"><br><%=message1%>

</p>
</td>
</tr>
<td align=center height="29">
<div align="right">
<p align="center">
<input type=button value="返 回 孤 岛" onClick="location.href='index.asp'" name="button">&nbsp; 
<font color="#FF0000"><b>[ </b></font> <a href="javascript:self.close()"><font color="#FF0000"><b>离开孤岛</b></font></a><font color="#FF0000"><b> ]</b></font> 
</div> 
</td> 
</table> 
</td> 
</tr> 
</table> 
</center> 
</body> 
</html> 
 
