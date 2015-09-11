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
lb=lcase(trim(request("lb")))
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
rs.open "SELECT "&lb&",w6,银两 FROM 用户 WHERE 姓名='"&aqjh_name&"'",conn
nnn=trim(rs(lb))
xj=rs("银两")
yywpsl=trim(rs("w6"))
rs.close
if yywpsl="" or yywpsl=null then
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('你的什么物品也没有，不能维修装备！！');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
if trim(nnn)="" or trim(nnn)="无" or trim(nnn)=Null then
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('你的"&lbmc&"并没有装备！！');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
data111=split(nnn,"|")
if ubound(data111)<>5 then
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('你的"&lbmc&"装备数据有误！！');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
if data111(5)>=3 then
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('你的"&data111(0)&"已维修过3次了，无法修复！！');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
rs.open "select * from b where a='"&data111(0)&"'",conn
if rs.eof or rs.bof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('物品数据库中无法找到["&data111(0)&"]，装备数据出错！！');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
wqlx=rs("b")
nai=rs("j")
newnai=data111(1)
ztz=int((newnai/nai)*100)
rs.close
if ztz>=95 then
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('你的["&data111(0)&"]状态良好，不需要维修！！');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
if ztz<15 then
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('你的["&data111(0)&"]已报废，无法修复！！');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
if data111(5)=1 and ztz>=75 then
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('武器已修理过一次，最高只能达到原耐久度的75％，不需要修理！！');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
if data111(5)=2 and ztz>=55 then
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('武器已修理过两次，最高只能达到原耐久度的55％，不需要修理！！');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
select case data111(5)
	case 0
		data111(1)=nai
	case 1
		data111(1)=int(nai*0.75)
	case 2
		data111(1)=int(nai*0.55)
end select
yin=(data111(1)-newnai)*10000
if xj<yin then
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('你没有["&yin&"]两银子，没人肯给你修装备！！');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
liang=int(((data111(1)-newnai)/10)+1)
xhwpmc=""
if wqlx="锻造" then
	rs.open "select * from x where a='"&data111(0)&"'",conn
	if rs.eof or rs.bof then
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		Response.Write "<script Language=Javascript>alert('锻造武器数据库中找不到你要维修的武器！！');location.href = 'javascript:history.go(-1)';</script>"
		Response.End
	end if
	sxwp=rs("b")
	rs.close
	xdata1=split(sxwp,"|")
	xdata2=UBound(xdata1)
	for ii=0 to xdata2-1
		xadata2=split(xdata1(ii),":")
		myxlwp=trim(xadata2(0))
		yywpsl=abate(yywpsl,myxlwp,liang)
		xhwpmc=xhwpmc&"["&xadata2(0)&"]"&liang&"个  "
	next
	xhwpmc=xhwpmc&","
else
	yywpsl=abate(yywpsl,"虎肉",liang)
	yywpsl=abate(yywpsl,"冰水",liang)
	yywpsl=abate(yywpsl,"矿石",liang)
	yywpsl=abate(yywpsl,"铁块",liang)
	yywpsl=abate(yywpsl,"木头",liang)
	xhwpmc="[虎肉]"&liang&"个  [冰水]"&liang&"个  [矿石]"&liang&"个  [铁块]"&liang&"个  [木头]"&liang&"个,"
end if
data111(5)=clng(data111(5)+1)
wqnewzt=data111(0)&"|"&data111(1)&"|"&data111(2)&"|"&data111(3)&"|"&data111(4)&"|"&data111(5)

conn.execute "update 用户 set "&lb&"='"&wqnewzt&"',w6='"&yywpsl&"',银两=银两-"&yin&" where 姓名='"&aqjh_name&"'"
set rs=nothing
conn.close
set conn=nothing
%>
<html>
<head>
<title>装备维护</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>

<body bgcolor="#FFFFFF" text="#000000" background="back.gif">
<p align="center"><font color="#FF0000" size="7">装备维护</font></p>
<p align="center">装备名称：<%=data111(0)%> 耐久度：<font color="#408080"><%=data111(1)%></font>/<font color="#FF0000"><%=nai%></font>  
  这是你第<font color="#FF0000"><%=data111(5)%></font>次维修此装备。</p> 
<p align="center">你消耗<%=xhwpmc%>花费银两：<%=yin%>，将<%=data111(0)%>的耐久度恢复至<%=data111(1)%>。<br>
  <br>
  说明：每样装备只能维修三次，三次后此装备将无法修复。</p>
<p align="center"><a href="index.asp">返回</a></p>
<p align="center"><font size="2">版权所有『爱情江湖总站』</font></p>
</body>
</html>
