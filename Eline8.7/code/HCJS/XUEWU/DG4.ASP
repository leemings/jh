<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"


sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
if sjjh_name="" then Response.Redirect "../../error.asp?id=440"


Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
my=sjjh_name
rs.open "select * from 用户 where 姓名='"&sjjh_name&"'",conn
if rs.eof or rs.bof then
	response.write "你不是江湖中人，不能进入江湖武馆"
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	response.end
else
tl=rs("体力")
nl=rs("内力")
%>
<html>
<head>
<style>
body{font-size:9pt;color:#000000;}
p{font-size:16;color:#000000;}
</style>
</head>
<body background="by.gif" bgproperties="fixed" bgcolor="#000000" vlink="#000000">
<%
if tl<1000 or nl<50 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
%>
<script language=vbscript>
MsgBox "错误！你目前的体力或内力值不够，回去多加练练马步蹲裆再来吧！"
location.href = "javascript:history.back()"
</script>
<%
else
conn.execute "update 用户 set 体力=体力-1000,内力=内力-50,武功=武功+150 where 姓名='"&sjjh_name&"'"
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing

message="" & my & "辛辛苦苦学了一天，你获得武功150！消耗体力1000内力50！"
		
end if
end if%>
<table border=1 bgcolor="#948754" align=center width=350 cellpadding="10" cellspacing="13">
<tr><td bgcolor=#C6BD9B>
<table height="260">
<tr><td height="37">
<font color="#000000"><strong>江湖武馆:</strong></font>
<tr>
<td height="182" valign="top"><%=message%><Br><Br>
</td>
</tr>
<td align=center height="29">&nbsp;
<div align="right">
<input type=button value="返 回" onclick="location.href='xuetang.htm'">
</div>
</td></tr>
</table>
</td></tr>
</table>
</body>
</html>
