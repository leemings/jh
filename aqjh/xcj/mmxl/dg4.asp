<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="../../showchat.asp"-->
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../../error.asp?id=440"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select * from 用户 where 姓名='"&aqjh_name&"'",conn
if rs.eof or rs.bof then
	response.write "你不是江湖中人，不能进入修炼武功"
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	response.end
end if
mp=rs("门派")
sj=DateDiff("n",rs("操作时间"),now())
if sj<8 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	ss=8-sj
	Response.Write "<script language=JavaScript>{alert('提示：请你等上["& ss &"]分,再操作！');window.close();}</script>"
	Response.End
end if
if rs("银两")<500000 then
		Response.Write "<script language=javascript>alert('抱歉！你的银两不够50万！');history.back();</script>"
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		Response.End	
	end if
if rs("体力")<10000 then
		Response.Write "<script language=javascript>alert('抱歉！你的体力不够1万~想走火入魔死啊！');history.back();</script>"
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		Response.End	
	end if
gg=rs("根骨")
if rs("内力")<10000 then
		Response.Write "<script language=javascript>alert('抱歉！你的内力不够1万,想走火入魔死啊！');history.back();</script>"
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		Response.End	
	end if
rs.close
rs.open "select * from 用户 where 姓名='"&aqjh_name&"'",conn
yy=15+gg*2
conn.execute "update 用户 set 体力=体力-10000,内力=内力-10000,银两=银两-500000,allvalue=allvalue+"&yy&",操作时间=now() where 姓名='"&aqjh_name&"'"
%>
<%
rs.close
set rs=nothing
conn.close
set conn=nothing
says="<font color=red>【投拜高人】["&aqjh_name&"]</font><font color=blue>经过三天三天夜的恳求（使得体力内力降低1万,花了50万白银），终于求得武林神话[<font color=red>无名</font>]答应指点武功，但由于<font color=red>["&aqjh_name&"]</font>知质有限，非练武材料，只令自己总积分上涨<font color=red>"&yy&"</font>点，相信经过武林神话的耐心教导，终究也能成为武林高手！！</font>"
call showchat(says)
%>
<html>
<head>
<style>
body{font-size:9pt;color:#000000;}
p{font-size:16;color:#000000;}
</style>
</head>
<body bgproperties="fixed" bgcolor="#000000" vlink="#000000">
<table border=1 bgcolor="#948754" align=center width=350 cellpadding="10" cellspacing="13">
<tr><td bgcolor=#C6BD9B>
<table height="260">
<tr>
          <td height="37"> <strong><font color="#800080">爱情江湖独创--门派修炼:</font></strong> 
        <tr>
          <td height="182" valign="top"><font color="#FF0000">你在门派修炼圣地修炼武功，经过8分钟的修炼，小有成就，战斗力提高20点（积分）！</font><font color="#FF00FF"><Br>
            </font><Br>
          </td>
</tr>
<td align=center height="29">&nbsp;
<div align="right">
<input type=button value="返 回" onclick="location.href='index.asp'">
</div>
</td></tr>
</table>
</td></tr>
</table>
</body>
</html>
