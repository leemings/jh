<%@ LANGUAGE=VBScript codepage ="936" %>
<script>
window.moveTo(30,30);
</script>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
if aqjh_name="" then Response.Redirect "../../error.asp?id=440"
session("aqjh_sj")=now()
Session("aqjh_jl")=0
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select 职业 from 用户 where 姓名='"&aqjh_name&"'",conn
if trim(rs("职业"))<>"采冰" then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：您的职业不是采冰工，所以您不能在这里进行木材开采！！');window.close();</script>"
	response.end
end if
rs.close
set rs=nothing	
conn.close
set conn=nothing
%>
<HTML>
<HEAD>
<TITLE>极地采冰</TITLE>
<META content="text/html; charset=gb2312" http-equiv=Content-Type>
<META content="MSHTML 5.00.2920.0" name=GENERATOR>
</HEAD>
<FRAMESET border=0 cols=*,100 frameBorder=no frameSpacing=0 rows=*>
<FRAME name=fs scrolling=no src="" frameSpacing=0 scrolling=no frameBorder=no>
<FRAMESET border=0 cols=* frameBorder=no frameSpacing=0 rows=50,0>
<FRAME border=0 frameBorder=no frameSpacing=0 name=show scrolling=no src="show.asp">
<FRAME border=0 frameBorder=no frameSpacing=0 name=cz scrolling=no src="about:blank">
</FRAMESET>
</FRAMESET>
</HTML>
