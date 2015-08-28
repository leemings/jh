<%@ codepage ="936" %>
<%
Response.Expires=-1
bgcolor=Application("yx8_mhjh_backgroundcolor")
bgimage=Application("yx8_mhjh_backgroundimage")
cmid=Request.QueryString("id")
if not isnumeric(cmid) then Response.Redirect "../error.asp?id=024"
username=session("yx8_mhjh_username")
if username="" then Response.Redirect "../error.asp?id=016"
set conn=server.createobject("adodb.connection") 
conn.Mode=16
conn.IsolationLevel=256
conn.open Application("yx8_mhjh_connstr")
set rst=server.CreateObject ("adodb.recordset")
sqlstr="select * from 物品 where id="&cmid&" and 寄售=True "
rst.open sqlstr,conn
if  rst.EOF or rst.BOF then Response.Redirect "../error.asp?id=029"
cmname=rst("名称")
cmtype=rst("属性")
cmhp=rst("体力")
cmmp=rst("内力")
cmattack=rst("攻击")
cmdefence=rst("防御")
cmespecial=rst("特效")
cmprice=rst("价格")
cmhost=rst("所有者")
rst.Close
rst.Open "select * from 用户 where 姓名='"&username&"' and 银两>="&cmprice,conn
if rst.EOF or rst.BOF then Response.Redirect "../error.asp?id=030"
rst.Close
conn.BeginTrans
conn.Execute "update 物品 set 寄售=False where id="&cmid
conn.CommitTrans
rst.Open "select * from 物品 where 名称='"&cmname&"' and 所有者='"&username&"'",conn
if rst.EOF or rst.BOF then
	conn.Execute "insert into 物品(名称,属性,体力,内力,攻击,防御,特效,价格,数量,所有者,寄售,装备) values('"&cmname&"','"&cmtype&"',"&cmhp&","&cmmp&","&cmattack&","&cmdefence&",'"&cmespecial&"',"&cmprice&",1,'"&username&"',False,False)"
else
	conn.Execute "update 物品 set 数量=数量+1 where 名称='"&cmname&"' and 所有者='"&username&"'"
end if
conn.Execute "update 用户 set 银两=银两+"&cmprice*0.9&" where 姓名='"&cmhost&"'"
conn.Execute "update 用户 set 银两=银两-"&cmprice&" where 姓名='"&username&"'"
conn.Execute sqlstr
conn.Close
set conn=nothing
dim newtalkarr(600) 
   talkarr=Application("yx8_mhjh_talkarr") 
		talkpoint=Application("yx8_mhjh_talkpoint") 
		j=1 
		for i=11 to 600 step 10 
			newtalkarr(j)=talkarr(i) 
			newtalkarr(j+1)=talkarr(i+1) 
			newtalkarr(j+2)=talkarr(i+2) 
			newtalkarr(j+3)=talkarr(i+3) 
			newtalkarr(j+4)=talkarr(i+4) 
			newtalkarr(j+5)=talkarr(i+5) 
			newtalkarr(j+6)=talkarr(i+6) 
			newtalkarr(j+7)=talkarr(i+7) 
			newtalkarr(j+8)=talkarr(i+8) 
			newtalkarr(j+9)=talkarr(i+9) 
			j=j+10 
		next 
		newtalkarr(591)=talkpoint+1 
		newtalkarr(592)=1 
		newtalkarr(593)=0 
		newtalkarr(594)=username 
		newtalkarr(595)="大家" 
		newtalkarr(596)="" 
		newtalkarr(597)="000000" 
		newtalkarr(598)="000000" 
		newtalkarr(599)="<font color=red>【江湖通报】</font><font color=blue>"&username&"在自由市场看见"&cmhost&"的"&cmname&"卖得很便宜，花了"&cmprice*0.9&"银子，给抢购去了，看来只要多花心思，还是能得到便宜的！</font>" 
newtalkarr(600)=Session("yx8_mhjh_userchatroomsn")
Application.Lock
Application("yx8_mhjh_talkpoint")=talkpoint+1
Application("yx8_mhjh_talkarr")=newtalkarr
Application.UnLock
erase newtalkarr
erase talkarr

%>
<head><title>自由集市</title><LINK href="../chatroom/css.css" rel=stylesheet><script language=javascript>setTimeout("location.replace('market.asp');",3000);</script></head>
<body bgcolor='<%=bgcolor%>' background='<%=bgimage%>' oncontextmenu=self.event.returnValue=false>
<table width=100% border=1 bordercolorlight=000000 cellspacing=0 cellpadding=2 bordercolordark=FFFFFF bgcolor=FFB442>
  <tr align=center> 
    <td bgcolor=#FFFF00><A href="market.asp" onmouseover="window.strtus='购买集市中流通的物品';return true;" onmouseout="window.status='';return true">购买</a></td>
    <td><A href="sale.asp" onmouseover="window.strtus='标价出售自己拥有的物品';return true;" onmouseout="window.status='';return true">寄售</a></td>
    <td><A href="chgprice.asp" onmouseover="window.strtus='变更已寄售物品的价格';return true;" onmouseout="window.status='';return true">改价</a></td>
</tr></table>
<p align=center><font color="#CC0000" size="4" face="幼圆"><b><%=Application("Ba_jxqy_systemname")%>自由集市</b></font></p>
<div align="center">钱货两讫，你收好了，谢谢您的惠顾 </div>
<p align=right><a href="javascript:location.replace('market.asp');" onmouseover="window.status='返回';return true;" onmouseout="window.status='';return true;">返回</a></p>
</body>

