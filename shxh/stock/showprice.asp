<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no cache"
Response.Expires=-1
bgcolor=Application("Ba_jxqy_backgroundcolor")
bgimage=Application("Ba_jxqy_backgroundimage")
msg="<head><link rel='stylesheet' href='../chatroom/css.css'><script language=javascript>function selectstock(stock){parent.stockfrm.location.replace('showstock.asp?stock='+stock);parent.infofrm.location.replace('info.asp?stock='+stock);}setTimeout('location.reload();',180000);</script></head><body oncontextmenu=self.event.returnValue=false topmargin=0 background='"&bgimage&"' bgcolor='"&bgcolor&"' text='FF0000'><table align=center width=100% border=0 cellspacing=0 cellpadding=2 bgcolor=ffcc00><tr align=center><td width=100>股票名称</td><td width=60>现价</td><td width=60>升降</td></tr></table><marquee direction=up bgcolor=ffcc00 scrolldelay=500 scrollamout=100><table align=center width=100% border=0 cellspacing=0 cellpadding=2 bgcolor=ffcc00>"
set conn=server.CreateObject("adodb.connection")
conn.Open Application("Ba_jxqy_connstr")
set rst=server.CreateObject("adodb.recordset")
rst.Open "select * from 股票",conn
do while not(rst.EOF or rst.BOF)
	increa =(rst("现价")/rst("最近成交价")-1)*100
	if increa>0 then 
		increa ="<font color=0000FF>"&formatnumber(increa,2,-1)&"</font>"
	elseif increa<0 then
		increa ="<font color=FF0000>"&formatnumber(increa,2,-1)&"</font>"
	else
		increa ="<font color=000000>"&formatnumber(increa,2,-1)&"</font>"
	end if		
	msg=msg&"<tr><td width=100 align=center><a href='#' onclick="&chr(34)&"javascript:selectstock('"&rst("股票名称")&"');"&chr(34)&">"&rst("股票名称")&"</a></td><td width=60 align=center>"&rst("现价")&"</td><td width=60 align=center>"&increa&"</td></tr>"
	rst.MoveNext
loop
msg=msg&"</table></marquee></body>"
Response.write msg
%>
