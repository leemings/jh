<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
sjjh_name=Session("sjjh_name")
if sjjh_name="" then Response.Redirect "../../error.asp?id=440"
http = Request.ServerVariables("HTTP_REFERER") 
if InStr(http,"work/ice")=0 then 
Response.Write "<script language=javascript>{alert('对不起，程序拒绝您的操作！！！\n     按确定返回！');parent.history.go(-1);}</script>" 
Response.End 
end if 
sj=DateDiff("s",session("sjjh_sj"),now())
if sj<10 then
	Response.Write "<script language=JavaScript>{alert('提示：请不要使用变速软件！');window.close();}</script>"
	Response.End
end if	
session("sjjh_sj")=now()
Session("sjjh_jl")=Session("sjjh_jl")+1
Response.Write "<script Language=Javascript>parent.show.fm2.innerHTML="&chr(34)&"<div align='center'>"&int(Session("sjjh_jl")/4)&"<img src='images/tong4.gif'></div>"&chr(34)&";parent.show.fm3.innerHTML="&chr(34)&"<div align='center'><img src='images/tong"&(Session("sjjh_jl") mod 4)&".gif'></div>"&chr(34)&"</script>"
%>
