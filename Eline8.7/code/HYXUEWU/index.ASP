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
end if
hy=rs("会员等级")
pdhy=rs("会员")
if hy=0 and pdhy=False then
 rs.close
 set rs=nothing
 conn.close
 set conn=nothing
 Response.Write "<script Language=javascript>alert('天呀，["&sjjh_name &"]你可不是会员，不能练武林绝学！');window.close();</script>"
 response.end
end if
if rs("银两")<1000000 then
		Response.Write "<script language=javascript>alert('抱歉你的条件不够，银子100万不够');history.back();</script>"
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		Response.End	
end if
sj=DateDiff("n",rs("操作时间"),now())
if sj<5 then
ss=5-sj
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('刚刚修炼过，请等："&ss&"分后再来修炼吧！');location.href = 'javascript:history.back()';}</script>"
	Response.End 
end if
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
<center>
<%
if tl<10000 or nl<5000 then
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
conn.execute "update 用户 set 体力=体力-10000,内力=内力-5000,银两=银两-1000000,武功=武功+5000 where 姓名='"&sjjh_name&"'"
rs.close
set rs=nothing
conn.close
set conn=nothing
message="" & my & "辛辛苦苦学了一天，你获得武功5000！消耗体力10000内力5000银两100万！"
end if
%>
<table border=1 bgcolor="#948754" align=center width=350 cellpadding="10" cellspacing="13">
<tr><td bgcolor=#C6BD9B>
<table height="260">
<tr><td height="37">
<font color="#000000"><strong>江湖武馆:</strong></font>
<tr>
<td height="182" valign="top"><%=message%><Br><Br><center>
</td>
</tr>
<td align=center height="29">&nbsp;
<div align="right">
<input type=button value="返 回" onclick="location.href='index.htm'">
</div>
</td></tr>
</table>
</td></tr>
</table>
</body>
</html>


