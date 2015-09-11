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
if hy<>4 and pdhy=False then
 rs.close
 set rs=nothing
 conn.close
 set conn=nothing
 Response.Write "<script Language=javascript>alert('天呀，["&aqjh_name &"]你会员不是4级，不能修炼金币！');window.close();</script>"
 response.end
end if
sj=DateDiff("n",rs("操作时间"),now())
if sj<20 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	ss=20-sj
	Response.Write "<script language=JavaScript>{alert('提示：请你等上["& ss &"]分,再操作！');window.close();}</script>"
	Response.End
end if
if rs("银两")<20000000 then
		Response.Write "<script language=javascript>alert('抱歉！你的银两不够2000万！');history.back();</script>"
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		Response.End	
	end if
if rs("金")<150 then
		Response.Write "<script language=javascript>alert('抱歉！你的金属性不够150点！');history.back();</script>"
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		Response.End	
	end if
if rs("火")<150 then
		Response.Write "<script language=javascript>alert('抱歉！你的火属性不够150点！');history.back();</script>"
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		Response.End	
	end if
if rs("水")<150 then
		Response.Write "<script language=javascript>alert('抱歉！你的水属性不够150点！');history.back();</script>"
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		Response.End	
	end if
if rs("会员金卡")>4500 then
		Response.Write "<script language=javascript>alert('抱歉！你个人金卡大于4500块~不能来了！');history.back();</script>"
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		Response.End	
	end if
if rs("轻功")<300000 then
		Response.Write "<script language=javascript>alert('抱歉！你的轻功不够30万~！');history.back();</script>"
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		Response.End	
	end if
if rs("法力")<300000 then
		Response.Write "<script language=javascript>alert('抱歉！你的法力不够30万！');history.back();</script>"
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		Response.End	
	end if
if rs("木")<150 then
		Response.Write "<script language=javascript>alert('抱歉！你的木属性不够150点！~');history.back();</script>"
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		Response.End	
	end if
if rs("土")<150 then
		Response.Write "<script language=javascript>alert('抱歉！你的土属性不够150点！~');history.back();</script>"
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		Response.End	
	end if
if rs("智力")<200 then
		Response.Write "<script language=javascript>alert('抱歉！你的智力不够200点');history.back();</script>"
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		Response.End	
end if
rs.close
rs.open "select * from 用户 where 姓名='"&aqjh_name&"'",conn
conn.execute "update 用户 set 防御=防御+100,轻功=轻功-300000,法力=法力-300000,银两=银两-20000000,会员金卡=会员金卡+4,金币=金币+5,金=金-150,水=水-150,火=火-150,智力=智力-200,木=木-150,土=土-150,操作时间=now(),道德=道德+2000 where 姓名='"&aqjh_name&"'"
%>
<%
rs.close
set rs=nothing
conn.close
set conn=nothing
says="<font color=red>【炼制金卡】["&aqjh_name&"]</font><font color=blue>在会员练金处练得了金卡<font color=red>4</font>块和金币<font color=red>5</font>个,五行属性减少<font color=red>150</font>点，轻功法力减少<font color=red>30</font>万点，智力减少<font color=red>200</font>点，银两减少<font color=red>2000</font>万!</font>"
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
          <td height="37"> <strong><font color="#800080">爱情江湖--修炼金卡:</font></strong> 
        <tr>
          <td height="182" valign="top"><img src="jk.gif" width="50" height="40" border="0"><font color="#FF0000">辛辛苦苦修炼了一天，你获得金币5个，金卡4块！消耗轻功20万法力20万银两2000万，五行减少150，防御增加100，道德增加2000，智力减少200点！</font><font color="#FF00FF"><Br>
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
