<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="../../../mywp.asp"-->
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
if aqjh_name="" then Response.Redirect "../../../error.asp?id=440"
http = Request.ServerVariables("HTTP_REFERER") 
if InStr(http,"jhhy/workauto/ice")=0 then 
Response.Write "<script language=javascript>{alert('对不起，程序拒绝您的操作！！！\n     按确定返回！');parent.history.go(-1);}</script>" 
Response.End 
end if 
wpsl=int(Session("aqjh_jl")/4)
if wpsl<=0 then
	Response.Write "<script Language=Javascript>alert('提示：你只采到0个怎么存？');</script>"
	response.end
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select w6,银两,体力,内力 from 用户 where 姓名='"&aqjh_name&"'",conn,2,2
zswp=add(rs("w6"),"冰水",wpsl)
mcsl=mywpsl(zswp,"冰水")
if rs("银两")<wpsl*700000 then
	if rs("体力")>1000 and rs("内力")>500 then
		aatl=int(rs("体力")/2)
		aanl=int(rs("内力")/2)
		conn.execute "update 用户 set 体力="&aatl&",内力="&aanl&" where id="&aqjh_id
	else
		conn.execute "update 用户 set 体力=100,内力=100 where id="&aqjh_id
	end if
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：你钱没带够，也想要什么冰水呀？想让我们替你白干？回去带钱来！！！伙计们，大家一起上去狠揍这家伙！');</script>"	
else
conn.execute "update 用户 set w6='"&zswp&"',银两=银两-" & wpsl*700000 & " where  姓名='"&aqjh_name&"'"
set rs=nothing
conn.close
set conn=nothing
Session("aqjh_jl")=0
Response.Write "<script Language=Javascript>;parent.show.fm1.innerHTML="&chr(34)&"<div align='center'>"&mcsl&"<img src='images/tong4.gif'></div>"&chr(34)&";parent.show.fm2.innerHTML="&chr(34)&"<div align='center'>0<img src='images/tong4.gif'></div>"&chr(34)&";parent.show.fm3.innerHTML="&chr(34)&"<div align='center'><img src='images/tong0.gif'></div>"&chr(34)&"</script>"
Response.Write "<script Language=Javascript>alert('提示：采到：["&wpsl&"]个，总数:["&mcsl&"]个？注意：如果你的银两不够支付"& wpsl*700000 &" 的工人工资，那你不仅得不到冰水，还会损失很多体力和内力！');</script>"
response.end
end if
%>