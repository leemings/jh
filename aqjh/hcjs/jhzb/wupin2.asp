<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="../../mywp.asp"-->
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../../error.asp?id=440"
name=lcase(trim(request("name")))
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "SELECT 操作时间,w3,z1,z2,z3,z4,z5,z6 FROM 用户 WHERE 姓名='"&aqjh_name&"'",conn,3,3
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
duyao=abate(rs("w3"),name,1)
z11=rs("z1")
z21=rs("z2")
z31=rs("z3")
z41=rs("z4")
z51=rs("z5")
z61=rs("z6")
rs.close
rs.open "select b,f,g,j,k FROM b WHERE a='" & name &"'",conn,3,3
If Rs.Bof OR Rs.Eof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：此类物品在数据库中不存在！');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End
end if
nai=rs("j")
lb1=rs("b")
wpfy=rs("g")
wpgj=rs("f")
wptx=rs("k")
select case lb1
	case "头盔"
		lb="z1"
		zb=z11
	case "装饰"
		lb="z2"
		zb=z21
	case "盔甲"
		lb="z3"
		zb=z31
	case "左手"
		lb="z4"
		zb=z41
	case "右手","锻造"
		lb="z5"
		zb=z51
	case "双脚"
		lb="z6"
		zb=z61
	case else
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		Response.Write "<script language=JavaScript>{alert('提示：类别数据出错，与程序开发商联系！');location.href = 'javascript:history.go(-1)';}</script>"
		Response.End
end select
if trim(zb)<>"" then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('你的["&lb1&"]装备了武器，如要更换请先抛弃旧的武器再装备新的！');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End
end if
wh=0
conn.execute "update 用户 set 攻击加=攻击加+"& wpgj &",防御加=防御加+"&wpfy&" where 姓名='"&aqjh_name&"'"
conn.execute "update 用户 set w3='"&duyao&"',"&lb&"='"&name&"|"&nai&"|"&wpgj&"|"&wpfy&"|"&wptx&"|"&wh&"' where  姓名='"&aqjh_name&"'"
rs.close
set rs=nothing
conn.close
set conn=nothing
Response.Write "<script Language=Javascript>alert('提示：装备物品["&name&"]完成\n按确定返回');location.href = 'javascript:history.go(-1)';</script>"
response.end
%>