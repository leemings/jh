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
	response.write "你不是江湖中人，不能进入修炼金币"
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	response.end
end if
hy=rs("会员等级")
pdhy=rs("会员")
if hy<>1 and pdhy=False then
 rs.close
 set rs=nothing
 conn.close
 set conn=nothing
 Response.Write "<script Language=javascript>alert('天呀，["&aqjh_name &"]你会员不是1级，不能修炼金币！');window.close();</script>"
 response.end
end if
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
if rs("银两")<800000 then
		Response.Write "<script language=javascript>alert('抱歉！你的银两不够80万！');history.back();</script>"
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		Response.End	
	end if
if rs("金")<50 then
		Response.Write "<script language=javascript>alert('抱歉！你的金属性不够50点！');history.back();</script>"
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		Response.End	
	end if
if rs("等级")<25 then
		Response.Write "<script language=javascript>alert('抱歉！你等级太低了，请25级后再来！');history.back();</script>"
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		Response.End	
	end if
if rs("体力")<100000 then
		Response.Write "<script language=javascript>alert('抱歉！你的体力不够10万~！');history.back();</script>"
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		Response.End	
	end if
if rs("内力")<100000 then
		Response.Write "<script language=javascript>alert('抱歉！你的内力不够10万！');history.back();</script>"
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		Response.End	
	end if
if rs("木")<50 then
		Response.Write "<script language=javascript>alert('抱歉！你的木属性不够50点！~');history.back();</script>"
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		Response.End	
	end if
if rs("智力")<20 then
		Response.Write "<script language=javascript>alert('抱歉！你的智力不够20点');history.back();</script>"
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		Response.End	
end if
rs.close
rs.open "select * from 用户 where 姓名='"&aqjh_name&"'",conn
conn.execute "update 用户 set 体力=体力-100000,内力=内力-100000,银两=银两-800000,金币=金币+2,金=金-50,智力=智力-20,木=木-50,操作时间=now() where 姓名='"&aqjh_name&"'"
%>
<%
rs.close
set rs=nothing
conn.close
set conn=nothing
says="<font color=red>【炼制金币】["&aqjh_name&"]</font><font color=blue>在会员练金处练得了金币<font color=red>2</font>个,金木属性减少<font color=red>50</font>点，体力内力减少<font color=red>10</font>万点，智力减少<font color=red>20</font>点，银两减少<font color=red>80</font>万!</font>"
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
          <td height="37"> <strong><font color="#800080">快乐江湖--修炼金币:</font></strong> 
        <tr>
          <td height="182" valign="top"><img src="../../chat/img/menoy.gif" width="50" height="40" border="0"><font color="#FF0000">辛辛苦苦修炼了一天，你获得金币2个！消耗体力10万内力10万银两80万，五行之金木减少50，智力减少20点！</font><font color="#FF00FF"><Br>
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
