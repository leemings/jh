<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="../../mywp.asp"-->
<!--#include file="../../showchat.asp"-->
<%
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../../error.asp?id=440"
cmid=Request.QueryString("id")
sl=Request.Form("sl")
if not (isnumeric(cmid) and isnumeric(sl)) then
Response.Write "<script language=javascript>alert('对不起，数量请使用数字！'); history.back();</script>"
Response.End
end if
sl=clng(sl)
Set connb = Server.CreateObject("ADODB.Connection")
connb.Open application("aqjh_usermdb")
Set rst= Server.CreateObject("ADODB.Recordset")
sqlstr="select * from 寄售 where id="&cmid
rst.open sqlstr,connb
if rst.EOF or rst.BOF then
Response.Write "<script language=javascript>alert('对不起，市场上没有这种商品出售或者已经卖完！'); history.back();</script>"
Response.End
end if
syz=rst("所有者")
danjia=rst("价格")
name=rst("名称")
lx=rst("类型")
shul=rst("数量")
yin=sl*danjia
if sl<1 then
Response.Write "<script language=javascript>alert('对不起，你什么都不买来这里干什么，想骗我东西啊？滚出去！'); history.back();</script>"
Response.End
end if
'shul为总数,sl为买的数量
if sl>shul then
Response.Write "<script language=javascript>alert('对不起，没有"&sl&"这么多"&name&"卖哦，目前还剩下"&shul&"个！'); history.back();</script>"
Response.End
end if
rst.Close
set conn=server.CreateObject("adodb.connection")
conn.Open Application("aqjh_usermdb")
set rst=server.CreateObject("adodb.recordset")
sqlstr="select 银两 from 用户 where 姓名='"&aqjh_name&"' "
rst.open sqlstr,conn
if rst.EOF or rst.BOF then
Response.Write "<script language=javascript>alert('对不起，你已经断开连接！'); history.back();</script>"
Response.End
end if
if rst("银两")<yin then
Response.Write "<script language=javascript>alert('对不起，总共需要银两"&yin&"，你钱不够哦！'); history.back();</script>"
Response.End
end if
rst.Close
sqlstr="select * from 用户 where 姓名='"&aqjh_name&"'"
rst.open sqlstr,conn
dangwp=rst(lx)
'增加物品
addsl=sl
zstemp=add(dangwp,name,addsl)
conn.execute "update 用户 set "&lx&"='"&zstemp&"' where  姓名='"&aqjh_name&"'"
if syz<>aqjh_name then
conn.execute "update 用户 set 银两=银两-"&yin&" where  姓名='"&aqjh_name&"'"
conn.execute "update 用户 set 银两=银两+"&yin&" where  姓名='"&syz&"'"
says="<font color=red>【市场信息】</font><font color=blue>"&aqjh_name&"花费了"&yin&"银子在自由市场里购买了"&addsl&"个"&name&"，还算挺便宜的哦！</font>"
else
says="<font color=red>【市场信息】</font><font color=blue>"&aqjh_name&"在自由市场吆喝了半天,可是没人买["&name&"]，天黑已晚,只好收摊回来了！</font>"
end if
sqlstr1="update 寄售 set 数量=数量-"&addsl&" where 名称='"&name&"' and 所有者='"&syz&"'"
connb.execute (sqlstr1)
sqlstr2="delete * from 寄售 where 数量=0"
connb.execute (sqlstr2)
call showchat(says)
%>
<head><title>自由集市</title><link rel="stylesheet" href="../../css.css"><script language=javascript>setTimeout("location.replace('Market.asp');",2000);</script></head>
<body background="../../bg.gif" oncontextmenu=self.event.returnValue=false>
<table width=100% border=1 bordercolorlight=000000 cellspacing=0 cellpadding=2 bordercolordark=FFFFFF bgcolor=FFB442>
<tr align=center>
<td><A href="market.asp" onmouseover="window.strtus='购买集市中流通的物品';return true;" onmouseout="window.status='';return true">购买</a></td>
<td bgcolor=#ffff00><A href="sale.asp" onmouseover="window.strtus='标价出售自己拥有的物品';return true;" onmouseout="window.status='';return true">寄售</a></td>
</tr></table>
<p align=center><font color="#CC0000" size="4" face="幼圆"><b><%=Application("aqjh_chatroomname")%>自由集市</b></font></p>
<p align=center>购买物品：您需要带够现金，购买成功后，银两直接划出</p>
<div align="center">我们的货物你收好了啊，如果不想要了，再来这里我帮你卖出去吧！ </div>
<p align=center><a href="javascript:location.replace('Market.asp');" onmouseover="window.status='返回';return true;" onmouseout="window.status='';return true;">返回</a></p>
</body>