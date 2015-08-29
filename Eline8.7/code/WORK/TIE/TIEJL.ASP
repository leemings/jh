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
if InStr(http,"work/tie")=0 then 
Response.Write "<script language=javascript>{alert('对不起，程序拒绝您的操作！！！\n     按确定返回！');parent.history.go(-1);}</script>" 
Response.End 
end if 
wpsl=int(Session("sjjh_jl")/3)
if wpsl<=0 then
	Response.Write "<script Language=Javascript>alert('提示：你只采到0个怎么存？');</script>"
	response.end
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "select w6 from 用户 where 姓名='"&sjjh_name&"'",conn,2,2
zswp=add(rs("w6"),"铁块",wpsl)
mcsl=mywpsl(zswp,"铁块")
conn.execute "update 用户 set w6='"&zswp&"' where  姓名='"&sjjh_name&"'"
set rs=nothing
conn.close
set conn=nothing
Session("sjjh_jl")=0
Response.Write "<script Language=Javascript>;parent.show.fm1.innerHTML="&chr(34)&mcsl&chr(34)&";parent.show.fm2.innerHTML="&chr(34)&"0"&chr(34)&";parent.show.fm3.innerHTML="&chr(34)&"0"&chr(34)&";parent.show.shiform.fmsl.value=0</script>"
Response.Write "<script Language=Javascript>alert('提示：采到：["&wpsl&"]个，总数:["&mcsl&"]个？');</script>"
response.end
%>
