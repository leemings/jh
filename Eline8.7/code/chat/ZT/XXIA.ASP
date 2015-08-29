<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="../../mywp.asp"-->
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
sjjh_name=Session("sjjh_name")
if sjjh_name="" then Response.Redirect "../../error.asp?id=440"
http = Request.ServerVariables("HTTP_REFERER") 
wpname=trim(Request.Form("wpname"))
if wpname="" then
	Response.Write "<script language=javascript>{alert('提示：你没有这样的武器！');parent.history.go(-1);}</script>" 
	Response.End 
end if
sl=1
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "SELECT b,h,l,m FROM b where a='"&wpname&"'",conn
if rs("b")<>"右手" and rs("b")<>"左手" and rs("b")<>"盔甲" and rs("b")<>"头盔" and rs("b")<>"双脚" and rs("b")<>"装饰" then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Redirect "../../error.asp?id=72"
	response.end
end if
tj=rs("l")
yin=rs("h")
jinbi=rs("m")
rs.close
sql="select w3 from 用户 where 姓名='"&sjjh_name&"'"
set rs=conn.execute(sql)
zstemp=add(rs("w3"),wpname,1)
rs.close
rs.open "select b,f,g,j,k FROM b WHERE a='" & wpname &"'",conn,3,3
lb=rs("b")
select case lb
	case "头盔"
		lb="z1"
	case "装饰"
		lb="z2"
	case "盔甲"
		lb="z3"
	case "左手"
		lb="z4"
	case "右手","锻造"
		lb="z5"
	case "双脚"
		lb="z6"
	case else
			rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		Response.Write "<script language=JavaScript>{alert('提示：类别数据出错，与程序开发商联系！');location.href = 'javascript:history.go(-1)';}</script>"
		Response.End
end select
conn.execute "update 用户 set w3='"&zstemp&"',"&lb&"='' where  姓名='"&sjjh_name&"'"
rs.close
set rs=nothing
conn.close
set conn=nothing
Response.Write "<script Language=Javascript>alert('提示：卸下物品["&wpname&"]完成\n按确定返回');location.href = 'javascript:history.go(-1)';</script>"
response.end
%>

