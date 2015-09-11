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
<title>种族转换</title>
<body bgcolor="#CCCCCC">
<%zhongzu=trim(request.form("jiu"))
if len(zhongzu)>3 then
	Response.Write "<script language=JavaScript>{alert('提示：你要作什么？！');window.close();}</script>"
	Response.End 
end if
if InStr(zhongzu,"or")<>0 or InStr(zhongzu,"'")<>0 or InStr(zhongzu,"`")<>0 or InStr(zhongzu,"=")<>0 or InStr(zhongzu,"-")<>0 or InStr(zhongzu,",")<>0 then 
	Response.Write "<script language=JavaScript>{alert('滚吧，你想做什么？想捣乱吗？！');window.close();}</script>"
	Response.End 
end if
if zhongzu<>"神族" and zhongzu<>"人族" and zhongzu<>"魔族"  then
	Response.Write "<script language=JavaScript>{alert('滚吧，你想做什么？想捣乱吗？！');window.close();}</script>"
	Response.End 
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select * from 用户 where 姓名='"&aqjh_name&"'",conn,2,2
if rs("种族")=zhongzu then 
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：你原来就是这个种族，还转什么呀！');window.close();}</script>"
	response.end
end if

if zhongzu="神族" and (aqjh_jhdj<150 or aqjh_grade>=11 or pi="出家") then
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		Response.Write "<script language=JavaScript>{alert('提示：想转职为神族，必须战斗等级最低为150级，并且是非出家人');window.close();}</script>"
		response.end
end if
if zhongzu="人族" and (aqjh_jhdj<90 or aqjh_grade>=11 or pi="出家" ) then
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		Response.Write "<script language=JavaScript>{alert('提示：想转职为人族，必须是战斗等级最低为90级，并且是非出家人');window.close();}</script>"
		response.end
end if
if zhongzu="魔族" and (aqjh_jhdj<100 or aqjh_grade>=11 or pi="出家" ) then
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		Response.Write "<script language=JavaScript>{alert('提示：想转职为魔族，必须是战斗等级最低为100级，并且是非出家人');window.close();}</script>"
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
if rs("威望")<2 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：种族转换需要威望2点！');window.close();}</script>"
	response.end
end if
if rs("转生")<3 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：种族转换至少转生3次以上！');window.close();}</script>"
	response.end
end if
pi=rs("门派")
sex=rs("性别")
if zhongzu="神族" or zhongzu="人族" or zhongzu="魔族"  then
	conn.execute "update 用户 set 银两=银两-50000000,威望=威望-1,种族='"& zhongzu &"' where 姓名='"&aqjh_name&"'",conn
	else
end if
rs.close
set rs=nothing
'conn.close
set conn=nothing
Response.Write "<script language=JavaScript>{alert('提示：您转换成了："& zhongzu &"种族，威望扣除一点，点击确定关闭当前窗口！');window.close();}</script>"
response.end
%>