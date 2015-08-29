<%@ LANGUAGE=VBScript codepage ="936" %>

<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
if sjjh_name="" then Response.Redirect "../error.asp?id=440"
%>
<title>职业转换♀wWw.51eline.com♀</title>
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
if ziye<>"采冰" and ziye<>"采矿" and ziye<>"伐木" and ziye<>"男贩" and ziye<>"女贩" and ziye<>"郎中" and ziye<>"武师" and ziye<>"大夫" and ziye<>"小偷" and ziye<>"气功师" and ziye<>"算命师" and ziye<>"魔法师" and ziye<>"倒夜香" and ziye<>"渔夫" then
	Response.Write "<script language=JavaScript>{alert('滚吧，你想做什么？想捣乱吗？！');window.close();}</script>"
	Response.End 
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "select * from 用户 where 姓名='"&sjjh_name&"'",conn,2,2
if rs("职业")=ziye then 
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：你原来就是这个职业，还转什么呀！');window.close();}</script>"
	response.end
end if

if ziye="算命师" and (sjjh_jhdj<20 or sjjh_grade>=11 or pi="出家") then
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		Response.Write "<script language=JavaScript>{alert('提示：想转职为算命师，必须战斗等级最低为20级，并且是非出家人，非官府人员！');window.close();}</script>"
		response.end
end if
if ziye="倒夜香" and (sjjh_jhdj<15 or sjjh_grade>=11 or pi="出家") then
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		Response.Write "<script language=JavaScript>{alert('提示：想转职为倒夜香，必须是战斗等级最低为15级，并且是非出家人，非官府人员！');window.close();}</script>"
		response.end
end if
if ziye="小偷" and (sjjh_jhdj<50 or sjjh_grade>=6 or pi="出家") then
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		Response.Write "<script language=JavaScript>{alert('提示：想转职为小偷，必须是战斗等级最低为50级，并且是非出家人，非官府人员！');window.close();}</script>"
		response.end
end if
if ziye="魔法师" and hy=rs("会员等级") and (sjjh_jhdj<20 or sjjh_grade>=11 or pi="出家" or hy<1) then
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		Response.Write "<script language=JavaScript>{alert('提示：想转职为魔法师，必须是会员，战斗等级最低为20级，并且是非出家人，非官府人员！');window.close();}</script>"
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
pi=rs("门派")
sex=rs("性别")


if ziye="采冰" or ziye="采矿" or ziye="伐木" or ziye="郎中" or ziye="大夫" or ziye="武师" or ziye="算命师" or ziye="魔法师" or ziye="气功师" or ziye="倒夜香" or ziye="渔夫" or ziye="小偷" then
	conn.execute "update 用户 set 银两=银两-500000,职业='"& ziye &"' where 姓名='"&sjjh_name&"'",conn
	else
	
	if ziye="男贩" and (sjjh_jhdj<50 or (sjjh_grade>=6 and sjjh_name<>"一线天") or pi="出家" or sex<>"男") then
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		Response.Write "<script language=JavaScript>{alert('提示：想转职为人贩，必须是男的，战斗等级最低为50级，并且是非出家人，非官府人员！');window.close();}</script>"
		response.end
	else
		if ziye="男贩" then
			conn.execute "update 用户 set 银两=银两-500000,职业='"& ziye &"' where 姓名='"& sjjh_name &"'",conn
		end if
	end if
	if ziye="女贩" and (sjjh_jhdj<50 or (sjjh_grade>=6 and sjjh_name<>"一线天") or pi="出家" or (sex<>"女" and sjjh_name<>"一线天")) then
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		Response.Write "<script language=JavaScript>{alert('提示：想转职为女贩，必须是女的，战斗等级最低为50级，并且是非出家人，非官府人员！');window.close();}</script>"
		response.end
	else
		if ziye="女贩" then
			conn.execute "update 用户 set 银两=银两-500000,职业='"& ziye &"' where 姓名='"& sjjh_name &"'",conn
		end if
	end if
end if
rs.close
set rs=nothing
'conn.close
set conn=nothing
Response.Write "<script language=JavaScript>{alert('提示：您职业转换成了："& ziye &"工作，点击确定关闭当前窗口！');window.close();}</script>"
response.end
%>
