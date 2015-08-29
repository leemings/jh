<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="../../mywp.asp"-->
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
if sjjh_name="" then Response.Redirect "../../error.asp?id=440"
http = Request.ServerVariables("HTTP_REFERER") 
if InStr(http,"hcjs/jhjs")=0 then 
Response.Write "<script language=javascript>{alert('对不起，程序拒绝您的操作！！！\n     按确定返回！');parent.history.go(-1);}</script>" 
Response.End 
end if
huaname=trim(Request.form("huaname"))
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "SELECT b,h,l,m FROM b where a='"&huaname&"'",conn
if rs("b")<>"宠物" then
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
rs.open "select 会员等级,银两,操作时间,cw,金币 from 用户 where 姓名='"&sjjh_name&"' and "&tj,conn,3,3
If Rs.Bof OR Rs.Eof then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：你不具备购买条件\n["&replace(tj,chr(39),"")&"]！');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End
end if
if rs("cw")<>"" then
zt=split(rs("cw"),"|")
if UBound(zt)=7 then
	Response.Write "<script Language=Javascript>alert('提示：你已经有宠物了，不可以购买.');location.href = 'javascript:history.go(-1)';</script>"
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	response.end
end if
end if
hy=rs("会员等级")
if hy>0 then
	yin=int(yin/2)
end if
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
'zstemp=add(rs("cw"),huaname,huasl)
zstemp=huaname&"|"&huaname&"|"&now()&"|100|100|100|0"&"|"&now()
'金龍|金龍|2002-6-28 11:04:29|100|100|100|0|2002-6-28 11:04:29
conn.execute "update 用户 set 银两=银两-" & yin & ",金币=金币-"&jinbi&",操作时间=now(),cw='"&zstemp&"' where 姓名='"&sjjh_name&"'"
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
<body topmargin="6" leftmargin="0" bgcolor="#FFFFFF" background="../JHZB/bg.gif">
<%
if hy>0 then
	Response.Write "<script Language=Javascript>alert('提示：会员购买物品打5折,购买"&huaname&"成功！');location.href = 'javascript:history.go(-1)';</script>"
else
	Response.Write "<script Language=Javascript>alert('提示：购买"&huaname&"成功,如果您是会员可以打5折！');location.href = 'javascript:history.go(-1)';</script>"
end if
%>
