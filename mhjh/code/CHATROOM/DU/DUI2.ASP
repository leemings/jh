<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
un=session("yx8_mhjh_username")
if un="" then Response.Redirect "../../error.asp?id=016"
mg=Request.QueryString("mg")
bb=Request.QueryString("bb")
changci=Request.QueryString("changci")
if instr(mg,"'")<>0 then Response.Redirect "../../error.asp?id=024"
if instr(bb,"'")<>0 then Response.Redirect "../../error.asp?id=024"
if un<>mg then Response.Redirect "../../error.asp?id=016"
chatroomsn=Session("yx8_mhjh_userchatroomsn")
set conn=server.CreateObject("adodb.connection")
conn.Open Application("yx8_mhjh_connstr")
set rst=server.CreateObject("adodb.recordset")
rst.Open "select 姓名,银两 from 用户 where 姓名='"&mg&"'",conn
if rst.EOF or rst.BOF then
	usecurative="你并没有中奖！"
else
	conn.execute "update 用户 set 银两=银两+"&bb&" where 姓名='"&mg&"'"
	%><!--#include file="data.asp"--> <%
    sql="update 投注 set 领奖=True where 场次="&changci&" and 下注者='"&mg&"'"
    connt.execute sql
    usecurative="你领取了"&bb&"两银子的奖金"

   	    talkarr=Application("yx8_mhjh_talkarr") 
		talkpoint=Application("yx8_mhjh_talkpoint") 
		dim newtalkarr(600) 
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
		newtalkarr(599)="<font color=red>【消息】</font><b><font color=red>"&usecurative&"</font></b>" 
		newtalkarr(600)=chatroomsn 
		Application("yx8_mhjh_talkarr")=newtalkarr
end if
rst.Close
set rst=nothing
conn.close
set conn=nothing
%>
<html>
<head>
<title></title>
<link rel=stylesheet href='../css.css'>
<script language=javascript>
setTimeout("location.href='dui.asp'",3000);
</script>
</head>
<body oncontextmenu=self.event.returnValue=false background='../../chatroom/bg1.gif'>
<div align=center>
<font size="2" face="宋体"><b>领奖结果</b></font>
<p>
三秒钟自动返回<br>
<input type=button value='返回' onclick="javascript:location.href='dui.asp'">
</div>
<%=usecurative%>
</body>
</html>

