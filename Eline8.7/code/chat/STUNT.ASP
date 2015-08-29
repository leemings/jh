<%@ LANGUAGE=VBScript codepage ="936" %><%if Instr(LCase(Application("sjjh_useronlinename"&session("nowinroom")))," "&LCase(sjjh_name)&" ")=0 then
	Response.Write "<script language=JavaScript>{alert('你不能进行操作，进行此操作必须进入聊天室！');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End 
end if

sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
if sjjh_name="" then Response.Redirect "../error.asp?id=440"
useronlinename=Application("sjjh_useronlinename"&session("nowinroom"))
who=Trim(Request.Form("who"))
if who="" then who=sjjh_name
show=Split(Trim(useronlinename),"  ",-1)
x=UBound(show)
chatroombgimage=Application("sjjh_chatimage")
chatroombgcolor=Application("sjjh_chatbgcolor")
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "SELECT 门派 FROM 用户 where 姓名='" & sjjh_name &"'",conn
if rs.bof or rs.eof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('你不是江湖人物！');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End 
end if
pai=rs("门派")
chatbgcolor=Application("sjjh_chatbgcolor")
chatimage=Application("sjjh_chatimage")
if chatbgcolor="" then chatbgcolor="008888"
%>
<html>
<head>
<title>连续技</title>
<META http-equiv="Content-Type" content="text/html; charset=gb2312">
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
<p>
<font color="#FFFFFF"><span style='font-size:9pt'><font size="3">【</font></span>连续技<span style='font-size:9pt'><font size="3">】</font></span></font><font size="3"><br>
</font> </p>
<p class=p1 align="center"><font color="#FFFFFF" size="2">最多可连三招<br>
伤害=三招伤害总和</font></p>
<form method=POST action='stunt1.asp'>
<font color="#FFFFFF" size="2"> </font><font size="2" color="#FF0033">攻击对像</font><br>
<select name="to1" style='BORDER-RIGHT: #000000 1px solid; BORDER-TOP: #000000 1px solid; FONT-SIZE: 12px; BACKGROUND: #666666; BORDER-LEFT: #000000 1px solid; COLOR: #f6e700; BORDER-BOTTOM: #000000 1px solid; HEIGHT: 16px;font-size:12px'>
<%
for i=0 to x
	if show(i)<>sjjh_name and show(i)<>peiou then
		%><option value="<%=show(i)%>"<%if CStr(show(i))=CStr(who) then Response.Write " selected"%>><%=show(i)%></option><%
	end if
next
%>
</select>
<font color="#FFFFFF" size="2"><br>
<br>
连招1<br>
<select style='BORDER-RIGHT: #000000 1px solid; BORDER-TOP: #000000 1px solid; FONT-SIZE: 12px; BACKGROUND: #666666; BORDER-LEFT: #000000 1px solid; COLOR: #f6e700; BORDER-BOTTOM: #000000 1px solid; HEIGHT: 16px;font-size:12px' name="at1" size="1">
<%
rs.close
rs.open "SELECT * FROM y where b='" & pai & "'",conn
do while not rs.eof
sel="selected"
response.write "<option " & sel & " value='"+CStr(rs("a"))+"' >"+rs("a")+""+chr(13)+chr(10)
rs.movenext
loop
%>
</select>
<br>
<br>
<br>
连招2<br>
<select style='BORDER-RIGHT: #000000 1px solid; BORDER-TOP: #000000 1px solid; FONT-SIZE: 12px; BACKGROUND: #666666; BORDER-LEFT: #000000 1px solid; COLOR: #f6e700; BORDER-BOTTOM: #000000 1px solid; HEIGHT: 16px;font-size:12px' name="at2" size="1">
<%
rs.close
rs.open "SELECT * FROM y where b='" & pai & "'",conn
do while not rs.eof
sel="selected"
response.write "<option " & sel & " value='"+CStr(rs("a"))+"'>"+rs("a")+""+chr(13)+chr(10)
rs.movenext
loop
%>
</select>
<br>
<br>
<br>
连招3<br>
<select style='BORDER-RIGHT: #000000 1px solid; BORDER-TOP: #000000 1px solid; FONT-SIZE: 12px; BACKGROUND: #666666; BORDER-LEFT: #000000 1px solid; COLOR: #f6e700; BORDER-BOTTOM: #000000 1px solid; HEIGHT: 16px;font-size:12px' name="at3" size="1">
<%
rs.close
rs.open "SELECT * FROM y where b='" & pai & "'",conn
do while not rs.eof
sel="selected"
response.write "<option " & sel & " value='"+CStr(rs("a"))+"'>"+rs("a")+""+chr(13)+chr(10)
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
</font><br>
</form>
  <div align="center"> <a href=javascript:history.go(-1)><font color=#00FF00 size=-1>返回上级</font></a></div>
</div>
</body>
</html>
