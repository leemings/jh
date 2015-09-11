<%@ LANGUAGE=VBScript codepage ="936" %>
<%
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../../error.asp?id=440"
acpage=request("acpage")
if not isnumeric(acpage) then acpage=1
acpage=cint(acpage)
if acpage<1 then acpage=1
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open application("aqjh_usermdb")
Set rst= Server.CreateObject("ADODB.Recordset")
sqlstr="select * from 寄售 where 数量>0 order by 价格 "
rst.open sqlstr,conn,1,2
if  not (rst.EOF or rst.BOF) then
	rst.PageSize=10
	if acpage>rst.PageCount then acpage=rst.PageCount
	rst.AbsolutePage=acpage
	msg="<table width='95%' border=1 bordercolorlight=000000 cellspacing=0 cellpadding=2 bordercolordark=FFFFFF align='center'><tr bgcolor=ffff00 align=center><td  width=40>类型</td><td width=100>名称</td><td>价格</td><td>数量</td><td>出售人</td><td width=40>操作</td></TR>"
	for j=1 to rst.Pagesize
		if rst.EOF or rst.BOF then
			exit for
		else
		select case rst("类型")
		case "w6"
			lx="物品w6"
		case "w8"
			lx="配药w8"
		case "w9"
			lx="魔宝w9"
		case "w10"
			lx="法器w10"
		case "w3"
			lx="装备w3"
		end select
			msg=msg&"<tr><td>"&lx&"</td><td>"&rst("名称")&"</td><td>"&rst("价格")&"</td><td>"&rst("数量")&"</td><td>"&rst("所有者")&"</td><td align='center'><a href='goumai.asp?id="&rst("id")&"'>购买</a></td></tr>"
			rst.MoveNext
		end if
	next
	msg=msg+"</tr></table><form action='market.asp' method=post id=form1 name=form1><table width='95%' border=1 bordercolorlight=000000 cellspacing=0 cellpadding=2 bordercolordark=FFFFFF align='center' bgcolor=ffFF00><tr align=center>"
	if acpage>1 then
		msg=msg&"<td><a href='market.asp?acpage=1' onmouseover="&chr(34)&"window.status='第一页';return true;"&chr(34)&" onmouseout="&chr(34)&"window.status='';return true;"&chr(34)&">第一页</a></td><td><a href='market.asp?acpage="&acpage-1&"' onmouseover="&chr(34)&"window.status='前一页';return true;"&chr(34)&" onmouseout="&chr(34)&"window.status='';return true;"&chr(34)&">前一页</a></td>"
	else
		msg=msg&"<td>第一页</td><td>前一页</td>"
	end if
	if acpage<rst.PageCount then
		msg=msg&"<td><a href='market.asp?acpage="&acpage+1&"' onmouseover="&chr(34)&"window.status='下一页';return true;"&chr(34)&" onmouseout="&chr(34)&"window.status='';return true;"&chr(34)&">下一页</a></td><td><a href='market.asp?acpage="&rst.PageCount&"' onmouseover="&chr(34)&"window.status='最后一页';return true;"&chr(34)&" onmouseout="&chr(34)&"window.status='';return true;"&chr(34)&">最后一页</a></td>"
	else
		msg=msg&"<td>后一页</td><td>最后一页</td>"
	end if
	msg=msg&"<td>第<input type=text name=acpage size=4 value='"&acpage&"'>页共"&rst.PageCount&"页<input type=submit value='GO' id=submit1 name=submit1></td></tr></table></form>"
end if
rst.Close
set rst=nothing 	
conn.Close 
set conn=nothing
if msg="" then msg="目前集市中没有东西可供出售！"
%>
<head><title>自由集市</title><LINK href="../../css.css" rel=stylesheet><meta http-equiv="Content-Type" content="text/html; charset=gb2312"></head>
<body bgcolor='<%=bgcolor%>' background='../../bg.gif' oncontextmenu=self.event.returnValue=false>
<table width=100% border=1 bordercolorlight=000000 cellspacing=0 cellpadding=2 bordercolordark=FFFFFF bgcolor=FFB442>
  <tr align=center> 
    <td bgcolor=#FFFF00><A href="market.asp" onmouseover="window.strtus='购买集市中流通的物品';return true;" onmouseout="window.status='';return true">购买</a></td>
    <td><A href="sale.asp" onmouseover="window.strtus='标价出售自己拥有的物品';return true;" onmouseout="window.status='';return true">寄售</a></td>
    </tr></table>
<p align=center><font color="#CC0000" size="4" face="幼圆"><b><%=Application("aqjh_chatroomname")%>自由集市</b></font></p>
<div align=center><%=msg%></div>
<p align=center>
  <input type="button" value=" 返 回 " onClick="javascript:location.href='zysc.asp'" name="button"> 
  <input type="button" value=" 关 闭 " onClick="javascript:window.close();" name="button">
</p>
</body>