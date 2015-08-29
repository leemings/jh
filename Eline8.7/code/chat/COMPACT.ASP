<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc/chatconfig.asp"-->
<%
if Instr(LCase(Application("sjjh_useronlinename"&session("nowinroom")))," "&LCase(sjjh_name)&" ")=0 then
	Response.Write "<script language=JavaScript>{alert('你不能进行操作，进行此操作必须进入聊天室！');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End 
end if

sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
if sjjh_name="" then Response.Redirect "../error.asp?id=440"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
useronlinename=Application("sjjh_useronlinename"&session("nowinroom"))
sjjh_name=sjjh_name
who=Trim(Request.Form("who"))
if who="" then who=sjjh_name
show=Split(Trim(useronlinename),"  ",-1)
x=UBound(show)
rs.open "SELECT * FROM 用户 where 姓名='" & sjjh_name &"'",conn
if rs.bof or rs.eof then
	response.write "你不是江湖中人"
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	response.end
end if
if rs("grade")>5 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('你们夫妻有一个是管理员，禁止使用！');history.go(-1);</script>"
	response.end
end if
peiou=rs("配偶")
rs.close
rs.open "SELECT * FROM 用户 where 姓名='" & peiou & "'",conn
if rs.bof or rs.eof then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('没有配偶呀！！');history.go(-1);</script>"
	response.end
end if
if rs("grade")>5 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('你们夫妻有一个是管理员，禁止使用！');history.go(-1);</script>"
	response.end
end if
rs.close
chatbgcolor=Application("sjjh_chatbgcolor")
chatimage=Application("sjjh_chatimage")
if chatbgcolor="" then chatbgcolor="008888"
%>
<html>
<head>
<title>合体技</title>
<META http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="../setup.css">
<style>
body{
CURSOR: url('3.ani');
scrollbar-track-color:#ffffff;
SCROLLBAR-FACE-COLOR:#effaff;
SCROLLBAR-HIGHLIGHT-COLOR:#ffffff;
SCROLLBAR-SHADOW-COLOR:#eeeeee;
SCROLLBAR-3DLIGHT-COLOR:#eeeeee;
SCROLLBAR-ARROW-COLOR:#dddddd;
SCROLLBAR-DARKSHADOW-COLOR:#ffffff
}
</style>
</head>
<body bgcolor="#006699" leftmargin="0" topmargin="2" bgproperties="fixed" oncontextmenu=self.event.returnValue=false>
<div align="center">
<p><br>
<font color="#FFFFFF"><span style='font-size:9pt'><font size="3">【</font></span>合体技<font size="3"><span style='font-size:9pt'>】</span></font></font><font size="3"><br>
</font> <font color="#FFFFFF" size="2"><br>
攻击力=男方杀伤力+女方杀伤力</font></p>
<form method=POST action='compact1.asp'>
<font color="#FFFFFF" size="2"> </font><font size="2" color="#FF0033">攻击对像</font><br>
<select name="to1" style='BORDER-RIGHT: #000000 1px solid; BORDER-TOP: #000000 1px solid; FONT-SIZE: 12px; BACKGROUND: #666666; BORDER-LEFT: #000000 1px solid; COLOR: #f6e700; BORDER-BOTTOM: #000000 1px solid; HEIGHT: 16px;font-size:12px'>
<%
for i=0 to x
if show(i)<>peiou and  show(i)<>sjjh_name Then
%>
<option value="<%=show(i)%>"<%if CStr(show(i))=CStr(who) then Response.Write " selected"%>><%=show(i)%></option>
<%
end if
next
%>
</select>
<font color="#FFFFFF" size="2"><br>
<br>
<font color="#FF0033">合体技一览</font><br>
<%
rs.open "SELECT * FROM t where b='" & sjjh_name & "' or c='" & sjjh_name & "'"
if rs.eof and rs.bof then
	response.write "<p align='center'>你没有任何合体技！<br>请自行到合体特技处创建。</p>"
else%>
	<select class="smallSel" name="at" size="1"><%
	do while not rs.eof
	sel="selected"
	response.write "<option " & sel & " value='"+CStr(rs("a"))+"' >"+rs("a")+""+chr(13)+chr(10)
	rs.movenext
	loop
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	%>
	</select>
	<br>
	<br>
	</font><font color=#ffffff size=2>
	<input type=submit value="发 招" name=submit style="BACKGROUND-COLOR: #009933; BORDER-BOTTOM: #ffffff 1px solid; BORDER-LEFT: #ffffff 1px solid; BORDER-RIGHT: #ffffff 1px solid; BORDER-TOP: #ffffff 1px solid; CURSOR: hand; COLOR: #FFFFFF; FONT-SIZE: 9pt; WIDTH:40;HEIGHT: 18px;">
	</font>
<%end if%>
<br></form>
  <div align="center"> <a href=javascript:history.go(-1)><font color=#00FF00 size=-1>返回上级</font></a></div>
	</div></body></html>
