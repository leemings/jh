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
if hy<>2 and pdhy=False then
 rs.close
 set rs=nothing
 conn.close
 set conn=nothing
 Response.Write "<script Language=javascript>alert('天呀，["&aqjh_name &"]你会员不是2级，不能修炼金币！');window.close();</script>"
 response.end
end if
sj=DateDiff("n",rs("操作时间"),now())
if sj<12 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	ss=12-sj
	Response.Write "<script language=JavaScript>{alert('提示：请你等上["& ss &"]分,再操作！');window.close();}</script>"
	Response.End
end if
if rs("银两")<5000000 then
		Response.Write "<script language=javascript>alert('抱歉！你的银两不够500万！');history.back();</script>"
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		Response.End	
	end if
if rs("金")<100 then
		Response.Write "<script language=javascript>alert('抱歉！你的金属性不够100点！');history.back();</script>"
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		Response.End	
	end if
if rs("等级")>400 then
		Response.Write "<script language=javascript>alert('抱歉！你等级太高了，请去别的地方修炼吧！');history.back();</script>"
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		Response.End	
	end if
if rs("体力")<500000 then
		Response.Write "<script language=javascript>alert('抱歉！你的体力不够50万~！');history.back();</script>"
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		Response.End	
	end if
if rs("法力")<100000 then
		Response.Write "<script language=javascript>alert('抱歉！你的法力不够10万！');history.back();</script>"
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		Response.End	
	end if
if rs("木")<100 then
		Response.Write "<script language=javascript>alert('抱歉！你的木属性不够100点！~');history.back();</script>"
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		Response.End	
	end if
if rs("土")<100 then
		Response.Write "<script language=javascript>alert('抱歉！你的土属性不够100点！~');history.back();</script>"
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		Response.End	
	end if
if rs("水")<100 then
		Response.Write "<script language=javascript>alert('抱歉！你的水属性不够100点！~');history.back();</script>"
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		Response.End	
	end if
if rs("智力")<80 then
		Response.Write "<script language=javascript>alert('抱歉！你的智力不够80点');history.back();</script>"
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		Response.End	
end if
rs.close
rs.open "select * from 用户 where 姓名='"&aqjh_name&"'",conn
conn.execute "update 用户 set 体力=体力-500000,法力=法力-100000,银两=银两-5000000,会员金卡=会员金卡+1,金币=金币+1,金=金-100,智力=智力-80,木=木-100,水=水-100,土=土-100,操作时间=now() where 姓名='"&aqjh_name&"'"
%>
<%
rs.close
set rs=nothing
conn.close
set conn=nothing
says="<font color=red>【炼制金卡】["&aqjh_name&"]</font><font color=blue>在会员练金处练得了会员金卡<font color=red>1</font>块，金币<font color=red>1</font>个,金木土水属性减少<font color=red>100</font>点，体力减少<font color=red>50</font>万点，法力减少<font color=red>10</font>万点，智力减少<font color=red>80</font>点，银两减少<font color=red>500</font>万!</font>"
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
          <td height="37"> <strong><font color="#800080">快乐江湖--修炼会员金卡:</font></strong> 
        <tr>
          <td height="182" valign="top"><img src="jk.gif" width="50" height="40" border="0"><font color="#FF0000">辛辛苦苦修炼了一天，你获得金币1个,金卡1个！消耗体力50万法力10万银两500万，五行之金木土水减少100，智力减少80点！</font><font color="#FF00FF"><Br>
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
