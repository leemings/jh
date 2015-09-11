<%@ LANGUAGE=VBScript codepage ="936" %>
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
<title>领取任务</title>
<body bgcolor="#CCCCCC">
<%renwu=trim(request.form("jiu"))
if InStr(renwu,"or")<>0 or InStr(renwu,"'")<>0 or InStr(renwu,"`")<>0 or InStr(renwu,"=")<>0 or InStr(renwu,"-")<>0 or InStr(renwu,",")<>0 then 
	Response.Write "<script language=JavaScript>{alert('滚吧，你想做什么？想捣乱吗？！');window.close();}</script>"
	Response.End 
end if
if renwu<>"神族任务" and renwu<>"人族任务" and renwu<>"魔族任务"  then
	Response.Write "<script language=JavaScript>{alert('滚吧，你想做什么？想捣乱吗？！');window.close();}</script>"
	Response.End 
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select * from 用户 where 姓名='"&aqjh_name&"'",conn,2,2
if rs("任务")=renwu then 
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：你原来就是这个任务，还换什么呀！');window.close();}</script>"
	response.end
end if

if renwu="神族任务" and (aqjh_jhdj<120 or aqjh_grade>=11 or pi="出家") then
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		Response.Write "<script language=JavaScript>{alert('提示：神族任务，必须战斗等级最低为120级，并且是非出家人');window.close();}</script>"
		response.end
end if
if renwu="人族任务" and (aqjh_jhdj<60 or aqjh_grade>=11 or pi="出家") then
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		Response.Write "<script language=JavaScript>{alert('提示：人族任务，必须是战斗等级最低为60级，并且是非出家人');window.close();}</script>"
		response.end
end if
if renwu="魔族任务" and (aqjh_jhdj<90 or aqjh_grade>=11 or pi="出家") then
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		Response.Write "<script language=JavaScript>{alert('提示：魔族任务，必须是战斗等级最低为90级，并且是非出家人');window.close();}</script>"
		response.end
		end if
if rs("银两")<5000000 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：你的钱不够500万！');window.close();}</script>"
	response.end
end if
pi=rs("门派")
sex=rs("性别")
if renwu="神族任务" or renwu="人族任务" or renwu="魔族任务"  then
	conn.execute "update 用户 set 银两=银两-50000000,任务='"& renwu &"',任务时间='"& time &"' where 姓名='"&aqjh_name&"'",conn
	else
end if
rs.close
set rs=nothing
'conn.close
set conn=nothing
Response.Write "<script language=JavaScript>{alert('提示：您接受了："& renwu &"，点击确定关闭当前窗口！');window.close();}</script>"
response.end
%>