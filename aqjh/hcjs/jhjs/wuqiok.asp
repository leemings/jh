<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="../../mywp.asp"-->
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
if aqjh_name="" then Response.Redirect "../../error.asp?id=440"
http = Request.ServerVariables("HTTP_REFERER") 
if InStr(http,"hcjs/jhjs")=0 then 
	Response.Write "<script language=javascript>{alert('对不起，程序拒绝您的操作！！！\n     按确定返回！');parent.history.go(-1);}</script>" 
	Response.End 
end if 
wpname=trim(Request.Form("wpname"))
if wpname="" then
	Response.Write "<script language=javascript>{alert('提示：请正确选择你要买的武器！');parent.history.go(-1);}</script>" 
	Response.End 
end if
sl=1
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
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
rs.open "select 会员等级,银两,操作时间,w3,金币 from 用户 where 姓名='"&aqjh_name&"' and "&tj,conn,3,3
If Rs.Bof OR Rs.Eof then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：你不具备购买条件\n["&replace(tj,chr(39),"")&"]！');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End
end if
hy=rs("会员等级")
select case hy
case 0
hygm=1
case 1
hygm=0.8
case 2
hygm=0.7
case 3
hygm=0.6
case 4
hygm=0.5
case 5
hygm=0.4
case 6
hygm=0.3
case 7
hygm=0.2
case 8
hygm=0.1
end select
yin=int(yin*hygm)
sj=DateDiff("s",rs("操作时间"),now())
if sj<10 then
	s=10-sj
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('对不起系统限制，请等["&s&"秒钟]再操作！');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End
end if	
if yin>rs("银两") or jinbi>rs("金币") then
	Response.Write "<script Language=Javascript>alert('提示：购买不成功，原因：你的银两不够了、或金币不够');location.href = 'javascript:history.go(-1)';</script>"
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	response.end
end if
zstemp=add(rs("w3"),wpname,1)
conn.execute "update 用户 set 银两=银两-" & yin & ",金币=金币-"&jinbi&",操作时间=now(),w3='"&zstemp&"' where 姓名='"&aqjh_name&"'"
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
<body topmargin="6" leftmargin="0" bgcolor="#FFFFFF" background="../../bg.gif">
<%
if hy>0 then
	Response.Write "<script Language=Javascript>alert('提示：你是"&hy&"级会员，购买物品打"&hygm*10&"折,购买"&wpname&"1个成功！');location.href = 'javascript:history.go(-1)';</script>"
else
	Response.Write "<script Language=Javascript>alert('提示：购买"&wpname&"1个成功,如果您是会员可以打折！');location.href = 'javascript:history.go(-1)';</script>"
end if
%>