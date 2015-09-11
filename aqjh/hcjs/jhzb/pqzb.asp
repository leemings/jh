<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"

aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../../error.asp?id=440"

lb=lcase(trim(request("wqlb")))
if lb="" or (lb<>"z1" and lb<>"z2" and lb<>"z3" and lb<>"z4" and lb<>"z5" and lb<>"z6") then
	Response.Write "<script Language=Javascript>alert('物品类别出错！！');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
if InStr(lb,"or")<>0 or InStr(lb,"=")<>0 or InStr(lb,"`")<>0 or InStr(lb,"'")<>0 or InStr(lb," ")<>0 or InStr(lb,"　")<>0 or InStr(lb,"'")<>0 or InStr(lb,chr(34))<>0 or InStr(lb,"\")<>0 or InStr(lb,",")<>0 or InStr(lb,"<")<>0 or InStr(lb,">")<>0 then Response.Redirect "../../error.asp?id=54"
select case lb
	case z1
		lbmc="头盔"
	case z2
		lbmc="装饰"
	case z3
		lbmc="盔甲"
	case z4
		lbmc="左手"
	case z5
		lbmc="右手"
	case z6
		lbmc="双脚"
end select
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "SELECT "&lb&" FROM 用户 WHERE 姓名='"&aqjh_name&"'",conn
nnn=trim(rs(lb))
if trim(nnn)="" or trim(nnn)=null then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：你的["&lbmc&"]并没有任何装备，抛弃什么？？！！');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
rs.close
data=split(nnn,"|")
ygj=data(2)
yfy=data(3)
conn.execute "update 用户 set 攻击加=攻击加-"&ygj&",防御加=防御加-"&yfy&" where 姓名='"&aqjh_name&"'"
conn.execute "update 用户 set "&lb&"='' where 姓名='"&aqjh_name&"'"
set rs=nothing
conn.close
set conn=nothing
%>
<html><head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>丢弃装备</title>
</head>
<body background="back.gif">
<p align="center"><font size="6" color="#FF0000"><b>丢弃装备</b></font></p>
<p align="center">你丢弃了<font color=green><%=lbmc%></font>的装备<font color=red>[<%=data(0)%>]</font>，如要重新装备，请购买新的装备。<br>
  <br>
  <a href="index.asp">返回</a> <br>
  <br>
  <font size="2">版权所有『爱情江湖总站』</font> </p>
</body></html>