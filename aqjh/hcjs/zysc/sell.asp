<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="../../mywp.asp"-->
<!--#include file="../../showchat.asp"-->
<%
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../../error.asp?id=440"
lx=request("lx")
name=request("name")
sl=request("sl")
danjia=request("danjia")
select case lx
case "物品"
dang="w6"
case "配药"
dang="w8"
case "魔宝"
dang="w9"
case "法器"
dang="w10"
case "装备"
dang="w3"
case else
Response.Write "<script language=javascript>alert('对不起，类型出错,想作弊吗?'); window.close();</script>"
Response.End
end select
if name="" then
Response.Write "<script language=javascript>alert('对不起，物品名出错!'); window.close();</script>"
Response.End
end if
if not isnumeric(sl) or not isnumeric(danjia) then
Response.Write "<script language=javascript>alert('对不起，数量或单价出错!'); window.close();</script>"
Response.End
end if
if sl<1 then
Response.Write "<script language=javascript>alert('对不起，你什么都不卖来这里干什么，滚出去！'); window.close();</script>"
Response.End
end if
if danjia<10000 then
Response.Write "<script language=javascript>alert('对不起，还不到10000两啊？你是不想卖呢？还是想白送人？不想卖就出去，想送人就去聊天室里送！'); window.close();</script>"
Response.End
end if
if danjia>1000000000 then
Response.Write "<script language=javascript>alert('对不起，都超出10亿了，谁能买得起？'); window.close();</script>"
Response.End
end if
set conn=server.createobject("adodb.connection")
conn.open Application("aqjh_usermdb")
set rst=server.CreateObject ("adodb.recordset")
sqlstr="select * from 用户 where 姓名='"&aqjh_name&"'"
rst.open sqlstr,conn
if rst.EOF or rst.BOF then
Response.Write "<script language=javascript>alert('对不起，你不是江湖中人！');window.close();</script>"
Response.End
end if
dangwp=rst(dang)
if IsNull(dangwp) then
Response.Write "<script language=javascript>alert('对不起，你没有这种类型的物品啊！');window.close();</script>"
Response.End
end if
if instr(dangwp,name)=0 then
Response.Write "<script language=javascript>alert('对不起，你没有这种物品啊！');window.close();</script>"
Response.End
end if
data1=split(dangwp,";")
data2=UBound(data1)
for y=0 to data2-1
	data3=split(data1(y),"|")
	if data3(0)=name then
		if sl>data3(1) then
			Response.Write "<script language=javascript>alert('对不起，你的物品数量不够啊！');window.close();</script>"
			Response.End
		end if
	'删除物品
	wpsl=sl
	mydang=abate(dangwp,name,wpsl)
	conn.execute "update 用户 set "&dang&"='"&mydang&"' where  姓名='"&aqjh_name&"'"
	Set connb = Server.CreateObject("ADODB.Connection")
	connb.Open Application("aqjh_usermdb")
	Set rst= Server.CreateObject("ADODB.Recordset")
	sql="select * from 寄售 where 名称='"&name&"' and 所有者='"&aqjh_name&"'"
	Set rst=connb.Execute(sql) 
	if rst.EOF or rst.BOF then
		sqlstr1="insert into 寄售(名称,类型,价格,数量,所有者) values ('"&name&"','"&dang&"',"&danjia&","&wpsl&",'"&aqjh_name&"')"
	else
		sqlstr1="update 寄售 set 数量=数量+"&wpsl&",价格="&danjia&" where 名称='"&name&"' and 所有者='"&aqjh_name&"'"
	end if
	connb.Execute(sqlstr1)
	says="<font color=red>【市场信息】</font><font color=blue>"&aqjh_name&"把自己多余的"&wpsl&"个"&name&"在江湖的自由市场出售了，单价为"&danjia&"银子，需要的大侠赶紧去抢购哦，过这村就没这店了！</font>"
	call showchat(says)
	exit for
	end if
next
%>
<head><title>自由集市</title><link rel="stylesheet" href="../../css.css"><script language=javascript>setTimeout("location.replace('sale.asp');",3000);</script></head>
<body background="../../bg.gif" oncontextmenu=self.event.returnValue=false>
<table width=100% border=1 bordercolorlight=000000 cellspacing=0 cellpadding=2 bordercolordark=FFFFFF bgcolor=FFB442>
<tr align=center>
<td><A href="market.asp" onmouseover="window.strtus='购买集市中流通的物品';return true;" onmouseout="window.status='';return true">购买</a></td>
<td bgcolor=#005b00><A href="sale.asp" onmouseover="window.strtus='标价出售自己拥有的物品';return true;" onmouseout="window.status='';return true">寄售</a></td>
</tr></table>
<p align=center><font color="#CC0000" size="4" face="幼圆"><b><%=Application("aqjh_chatroomname")%>自由集市</b></font></p>
<div align="center">你的货物我们收下了，如果卖出我们将马上将钱给你，恕不另行通知了！服务费是10% </div>
<p align=center><a href="javascript:location.replace('sale.asp');" onmouseover="window.status='返回';return true;" onmouseout="window.status='';return true;">返回</a></p>
</body>