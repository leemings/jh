<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="../showchat.asp"-->
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
%>
<title>职业转换</title>
<body bgcolor="#CCCCCC">
<%ziye=trim(request.form("jiu"))
if len(ziye)>3 then
	Response.Write "<script language=JavaScript>{alert('提示：你要作什么？！');window.close();}</script>"
	Response.End 
end if
if InStr(ziye,"or")<>0 or InStr(ziye,"'")<>0 or InStr(ziye,"`")<>0 or InStr(ziye,"=")<>0 or InStr(ziye,"-")<>0 or InStr(ziye,",")<>0 then 
	Response.Write "<script language=JavaScript>{alert('滚吧，你想做什么？想捣乱吗？！');window.close();}</script>"
	Response.End 
end if
if ziye<>"采冰" and ziye<>"采矿" and ziye<>"伐木" and ziye<>"冒险家" and ziye<>"炼金师" and ziye<>"男贩" and ziye<>"女贩" and ziye<>"武师" and ziye<>"大夫" and ziye<>"小偷" and ziye<>"气功师" and ziye<>"算命师" and ziye<>"魔法师" and ziye<>"倒夜香" and ziye<>"渔夫" then
	Response.Write "<script language=JavaScript>{alert('滚吧，你想做什么？想捣乱吗？！');window.close();}</script>"
	Response.End 
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select * from 用户 where 姓名='"&aqjh_name&"'",conn,2,2
if rs("职业")=ziye then 
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：你原来就是这个职业，还转什么呀！');window.close();}</script>"
	response.end
end if

if ziye="算命师" and (aqjh_jhdj<20 or aqjh_grade>=11 or pi="出家") then
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		Response.Write "<script language=JavaScript>{alert('提示：想转职为算命师，必须战斗等级最低为20级，并且是非出家人，非官府人员！');window.close();}</script>"
		response.end
end if
if ziye="倒夜香" and (aqjh_jhdj<15 or aqjh_grade>=11 or pi="出家") then
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		Response.Write "<script language=JavaScript>{alert('提示：想转职为倒夜香，必须是战斗等级最低为15级，并且是非出家人，非官府人员！');window.close();}</script>"
		response.end
end if
if ziye="小偷" and (aqjh_jhdj<20 or aqjh_grade>=6 or pi="出家") then
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		Response.Write "<script language=JavaScript>{alert('提示：想转职为小偷，必须是战斗等级最低为20级，并且是非出家人，非官府人员！');window.close();}</script>"
		response.end
end if
if ziye="冒险家" and (aqjh_jhdj<50 or pi="出家") then
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		Response.Write "<script language=JavaScript>{alert('提示：想转职为冒险家，必须是战斗等级最低为50级，并且是非出家人！');window.close();}</script>"
		response.end
end if
if ziye="炼金师" and (aqjh_jhdj<60 or pi="出家" or pi="天网") then
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		Response.Write "<script language=JavaScript>{alert('提示：想转职为炼金师，必须是战斗等级最低为60级，并且是非出家人！');window.close();}</script>"
		response.end
end if
if ziye="魔法师" and hy=rs("会员等级") and (aqjh_jhdj<80 or pi="出家" or hy<1) then
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		Response.Write "<script language=JavaScript>{alert('提示：想转职为魔法师，必须是会员，战斗等级最低为80级，并且是非出家人！');window.close();}</script>"
		response.end
end if
pi=rs("门派")
sex=rs("性别")
if ziye="男贩" and (aqjh_jhdj<20 or aqjh_grade>=6 or pi="出家" or sex<>"男") then
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：想转职为男贩，必须是男的，战斗等级最低为20级，并且是非出家人，非官府人员！');window.close();}</script>"
		response.end
		end if
		if ziye="女贩" and (aqjh_jhdj<20 or aqjh_grade>=6 or pi="出家" or sex<>"女") then
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		Response.Write "<script language=JavaScript>{alert('提示：想转职为女贩，必须是女的，战斗等级最低为20级，并且是非出家人，非官府人员！');window.close();}</script>"
		response.end
		end if
if rs("银两")<500000 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：你的钱不够50万！');window.close();}</script>"
	response.end
end if
if rs("金币")<10 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：你连10个金币都没有，还想要职业！');window.close();}</script>"
	response.end
end if
if rs("智力")<1000 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：转换职业需要1000智力的根基，去提高吧！');window.close();}</script>"
	response.end
end if
pi=rs("门派")
sex=rs("性别")
if ziye="采冰" or ziye="采矿" or ziye="伐木" or ziye="男贩" or ziye="女贩" or ziye="大夫" or ziye="武师" or ziye="算命师" or ziye="气功师" or ziye="倒夜香" or ziye="渔夫" or ziye="小偷"  then
	conn.execute "update 用户 set 智力=智力-100,金币=金币-5,银两=银两-500000,职业='"& ziye &"' where 姓名='"&aqjh_name&"'",conn
	else
end if
if ziye="魔法师" or ziye="冒险家" or ziye="炼金师"  then
	conn.execute "update 用户 set 智力=智力-500,金币=金币-10,银两=银两-500000,职业='"& ziye &"' where 姓名='"&aqjh_name&"'",conn
	else
end if
rs.close
set rs=nothing
'conn.close
set conn=nothing
says="<font color=red>【职业转换】["&aqjh_name&"]</font><font color=blue>转换成职业<font color=red>["&ziye&"]</font>,希望这个职业能令他开心！</font>"
call showchat(says)
Response.Write "<script language=JavaScript>{alert('提示：您职业转换成了："& ziye &"工作，点击确定关闭当前窗口！');window.close();}</script>"
response.end
%>