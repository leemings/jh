<%
Response.Expires=-1
bgcolor=Application("Ba_jxqy_backgroundcolor")
bgimage=Application("Ba_jxqy_backgroundimage")
cmid=Request.QueryString("id")
cmprice=Request.Form("price")
if not (isnumeric(cmid) and  isnumeric(cmprice))then Response.Redirect "../error.asp?id=024"
cmprice=clng(cmprice)
if cmprice<1 then cmprice=1
username=session("Ba_jxqy_username")
if username="" then Response.Redirect "../error.asp?id=016"
set conn=server.createobject("adodb.connection") 
conn.open Application("Ba_jxqy_connstr")
set rst=server.CreateObject ("adodb.recordset")
sqlstr="select * from 物品 where id="&cmid&" and 所有者='"&username&"' and 数量>0 and 寄售=False "
rst.open sqlstr,conn
if  rst.EOF or rst.BOF then Response.Redirect "../error.asp?id=028"
rst.Close
set rst=nothing 	
sqlstr="update 物品 set 数量=数量-1,价格="&cmprice&",寄售=True where id="&cmid&" and 所有者='"&username&"'"
conn.Execute sqlstr
conn.Close
set conn=nothing
%>
<head><title>自由集市</title><link rel="stylesheet" href="../chatroom/css.css"><script language=javascript>setTimeout("location.replace('sale.asp');",3000);</script></head>
<body bgcolor='<%=bgcolor%>' background='<%=bgimage%>' oncontextmenu=self.event.returnValue=false>
<table width=100% border=1 bordercolorlight=000000 cellspacing=0 cellpadding=2 bordercolordark=FFFFFF bgcolor=FFB442>
  <tr align=center> 
    <td><A href="market.asp" onmouseover="window.strtus='购买集市中流通的物品';return true;" onmouseout="window.status='';return true">购买</a></td>
    <td bgcolor=#FFFF00><A href="sale.asp" onmouseover="window.strtus='标价出售自己拥有的物品';return true;" onmouseout="window.status='';return true">寄售</a></td>
    <td><A href="chgprice.asp" onmouseover="window.strtus='变更已寄售物品的价格';return true;" onmouseout="window.status='';return true">改价</a></td>
</tr></table>
<p align=center><font color="#990000" size="4" face="幼圆"><b><%=Application("Ba_jxqy_systemname")%>自由集市</b></font></p>
<div align="center">你的货物我们收下了，如果卖出我们将马上将钱给你，恕不另行通知了！服务费是10% </div>
<p align=center><a href="javascript:location.replace('sale.asp');" onmouseover="window.status='返回';return true;" onmouseout="window.status='';return true;">返回</a></p>
</body>