<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="../mywp.asp"-->
<%Response.Buffer=true
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.Expires = 0
sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
if sjjh_name="" then Response.Redirect "../error.asp?id=440"
cz=LCase(trim(Request.QueryString("cz")))
lb=LCase(trim(Request.QueryString("lb")))
toname=LCase(trim(Request.QueryString("toname")))
name=LCase(trim(Request.QueryString("name")))
if InStr(toname,"'")<>0 or InStr(toname,"=")<>0 then 
	Response.Write "<script language=JavaScript>{alert('滚吧，你想做什么？想捣乱吗？！');window.close();}</script>"
	Response.End 
end if
if sjjh_grade<>10 then
	Response.Write "<script Language=Javascript>alert('提示：你来这里作什么，想死呀！');window.close();</script>"
	response.end
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "select "&lb&" from 用户 where 姓名='"&toname&"'",conn,2,2
if cz="1" then
	zstemp=del(rs(lb),name)
else
	zstemp=add(rs(lb),name,10)
end if
conn.execute "update 用户 set "&lb&"='"&zstemp&"' where  姓名='"&toname&"'"
rs.close
set rs=nothing
conn.close
set conn=nothing
if cz="1" then
	Response.Write "<script Language=Javascript>alert('提示：删除["&toname&"]物品["&name&"]成功');location.href = 'towupin.asp?toname="&toname&"';</script>"
else
	Response.Write "<script Language=Javascript>alert('提示：增加["&toname&"]物品["&name&"]10个成功');location.href = 'towupin.asp?toname="&toname&"';</script>"
end if
%>
