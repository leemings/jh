<%
Response.Buffer =true
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
username=session("yx8_mhjh_username")
st=server.htmlencode(trim(Request.Form("username")))
mn=server.htmlencode(trim(Request.Form("month")))
opt=server.htmlencode(trim(Request.Form("submit")))
if instr(username,"'")<>0 then Response.Redirect "../error.asp?id=024"
if not isnumeric(1) then Response.Redirect "../error.asp?id=024"
today=date()
set conn=server.CreateObject("adodb.connection")
conn.Open Application("yx8_mhjh_connstr")
set rst=server.CreateObject("adodb.recordset")
rst.Open "select 会员时间 from 用户 where 姓名='"&username&"' and 积分>5000 and 积分<10000 and 会员=False",conn
if rst.EOF or rst.BOF then Response.Redirect "../error.asp?id=228"
if rst("会员时间")>today then today=rst("会员时间")
if opt="申请" then
fellowdate=dateadd("m",1,today)
else
fellowdate=dateadd("m",-1,today)
end if
fellowdatetype="#"&month(fellowdate)&"/"&day(fellowdate)&"/"&year(fellowdate)&"#"
rst.Close
set rst=nothing
conn.Execute "update 用户 set 会员=true,会员时间="&fellowdatetype&" where 姓名='"&st&"'"
conn.Close
set conn=nothing
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
		newtalkarr(599)="<font color=red>【江湖通报】</font><font color=blue>"&st&"经过几天的艰苦努力,终于获得了免费会员资格,自己在会员申请功能里成功获得了会员身份!</font>" 
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
<link rel=stylesheet href='../chatroom/css.css'>
</head>
<body oncontextmenu=self.event.returnValue=false background='../chatroom/bg1.gif'>
<p align=center><font color=0000ff size=4>会员自动申请</font></p>
<hr>
<table width=80% align=center>
<tr><td>会员</td><td><%=st%></td></tr>
<tr><td>时限</td><td><%=formatdatetime(fellowdate,1)%></td></tr>
<tr align=center><td colspan=2><a href='showfell.asp' onmouseover="window.status='返回';return true;" onmouseout="window.status='';return true;">返回</a></td></tr>
</table>
</body>
</html>
