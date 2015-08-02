<%
bgcolor=Application("Ba_jxqy_backgroundcolor")
bgimage=Application("Ba_jxqy_backgroundimage")
msg=Request.QueryString("msg")
if msg="" then msg="武器"
msgarr=array("头盔","盔甲","武器","防具","饰品","暗器")
msgtmp="<table width=100% border=1 bordercolorlight=000000 cellspacing=0 cellpadding=2 bordercolordark=FFFFFF bgcolor=FFB442><tr align=center>"
for i=0 to 5
	if msg=msgarr(i) then
		msgtmp=msgtmp&"<td bgcolor=ffFF00><a href='armorshop.asp?msg="&msgarr(i)&"' onmouseover=""window.status='"&msgarr(i)&"类';return true;"" onmouseout=""window.status='';return true;"">"&msgarr(i)&"</a></td>"
	else
		msgtmp=msgtmp&"<td><a href='armorshop.asp?msg="&msgarr(i)&"' onmouseover=""window.status='"&msgarr(i)&"类';return true;"" onmouseout=""window.status='';return true;"">"&msgarr(i)&"</a></td>"
	end if	
next
msgtmp=msgtmp&"</tr></table>"
set conn=server.createobject("adodb.connection") 
conn.open Application("Ba_jxqy_connstr")
set rst=server.CreateObject ("adodb.recordset")
sqlstr="select ID,名称,攻击,防御,特效,价格 from 商店 where 属性='"&msg&"' order by 价格 "
rst.open sqlstr,conn
msg="<table width='100%' border=1 bordercolorlight=000000 cellspacing=0 cellpadding=2 bordercolordark=FFFFFF><tr bgcolor=FFFF00 align=center><td>名称</td><td>攻击</td><td>防御</td><td>特效</td><td>价格</td></TR>"
do while not(rst.EOF or rst.BOF)
	msg=msg&"<tr><td width=100><a href='buy.asp?id="&rst("id")&"' onmouseover="&chr(34)&"window.status='购买"&rst("名称")&"';return true;"&chr(34)&" onmouseout="&chr(34)&"window.status='';return true;"&chr(34)&">"&rst("名称")&"</td><td> "&rst("攻击")&"</td><td> "&rst("防御")&"</td><td> "&rst("特效")&"</td><td> "&rst("价格")&"</td></tr>"
	rst.MoveNext
loop
msg=msg+"</table>"
rst.Close
set rst=nothing 	
conn.Close 
set conn=nothing
%>
<head><LINK href="../chatroom/css.css" rel=stylesheet><title>武器店</title></head>
<body bgcolor='<%=bgcolor%>' background='../images/bg.gif' oncontextmenu=self.event.returnValue=false>
<div align=center>
<%=msgtmp%>
  <p align=center><b><font color=CC0000 face="幼圆" size="4"><%=Application("Ba_jxqy_systemname")%>武器铺</font></b></p>
<%=msg%></div>
<p align=center>
  <input type="button" value=" 返 回 " onClick="javascript:location.href='street.asp'" name="button"> 
  <input type="button" value=" 关 闭 " onClick="javascript:window.close();" name="button">
</p>
</body>