<%
Response.Expires=-1
bgcolor=Application("yx8_mhjh_backgroundcolor")
cmid=Request.QueryString("id")
cmprice=Request.Form("price")
if not (isnumeric(cmid) and  isnumeric(cmprice))then Response.Redirect "../error.asp?id=024"
cmprice=clng(cmprice)
if cmprice<1 then cmprice=1
username=session("yx8_mhjh_username")
if username="" then Response.Redirect "../error.asp?id=016"
set conn=server.createobject("adodb.connection")
conn.open Application("yx8_mhjh_connstr")
set rst=server.CreateObject ("adodb.recordset")
sqlstr="select * from 物品 where id="&cmid&" and 所有者='"&username&"' and 数量>0 and 寄售=False "
rst.open sqlstr,conn
if  rst.EOF or rst.BOF then Response.Redirect "../error.asp?id=028"
jg=rst("价格")
rst.Close
set rst=nothing
if cmprice>jg then Response.Redirect "../error.asp?id=032"
sqlstr="update 物品 set 数量=数量-1,价格="&cmprice&",寄售=True where id="&cmid&" and 所有者='"&username&"'"
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
		newtalkarr(599)="<font color=red>【江湖通报】</font><font color=blue>"&username&"把自己多余的一个神秘宝物在自由市场出售了，价格为"&cmprice&"银子，需要的大侠赶紧去抢购哦，过这村就没这店了！</font>" 
newtalkarr(600)=Session("yx8_mhjh_userchatroomsn")
Application.Lock
Application("yx8_mhjh_talkpoint")=talkpoint+1
Application("yx8_mhjh_talkarr")=newtalkarr
Application.UnLock
erase newtalkarr
erase talkarr

%>
<head><title>自由集市</title><link rel="stylesheet" href="../style4.css"><script language=javascript>setTimeout("location.replace('sale.asp');",3000);</script></head>
<body bgcolor='<%=bgcolor%>' background='<%=bgimage%>' oncontextmenu=self.event.returnValue=false>
<table width=100% border=1 bordercolorlight=000000 cellspacing=0 cellpadding=2 bordercolordark=FFFFFF bgcolor=FFB442>
<tr align=center>
<td><A href="market.asp" onmouseover="window.strtus='购买集市中流通的物品';return true;" onmouseout="window.status='';return true">购买</a></td>
<td bgcolor=#005b00><A href="sale.asp" onmouseover="window.strtus='标价出售自己拥有的物品';return true;" onmouseout="window.status='';return true">寄售</a></td>
<td><A href="chgprice.asp" onmouseover="window.strtus='变更已寄售物品的价格';return true;" onmouseout="window.status='';return true">改价</a></td>
</tr></table>
<p align=center><b><font color="#FFFF00" size="6" face="方正舒体">自由集市</font></b></p>
<div align="center">你的货物我们收下了，如果卖出我们将马上将钱给你，恕不另行通知了！服务费是10% </div>
<p align=center><a href="javascript:location.replace('sale.asp');" onmouseover="window.status='返回';return true;" onmouseout="window.status='';return true;">返回</a></p>
</body>
