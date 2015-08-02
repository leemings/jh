<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no cache"
Response.Expires=-1
stock=Request.QueryString("stock")
if stock="" then stock="梦想成真"
bgcolor=Application("Ba_jxqy_backgroundcolor")
bgimage=Application("Ba_jxqy_backgroundimage")
msg="<head><link rel='stylesheet' href='../style.css'><script language=javascript>function petition(stock){parent.infofrm.location.replace('petition.asp?stock='+stock);}</script></head><body oncontextmenu=self.event.returnValue=false topmargin=0 background='"&bgimage&"' bgcolor='"&bgcolor&"' text='FF0000'><center><font color=red><h3>"&stock&"</h3></font></center><table width=80% border=1 bordercolorlight=000000 cellspacing=0 cellpadding=2 bordercolordark=FFFFFF align=center>"
set conn=server.CreateObject("adodb.connection")
conn.Open Application("Ba_jxqy_connstr")
set rst=server.CreateObject("adodb.recordset")
rst.Open "select * from 股票 where 股票名称='"&stock&"'",conn
if rst.EOF or rst.BOF then
	msg=msg&"<tr><td>你要找的股票没有找到</td></tr>"
else
	for i=1 to rst.Fields.Count -1
		msg=msg&"<tr><td bgcolor=00ff00>"&rst.Fields(i).name&"</td><td align=right>"&rst.Fields(i).value&"</td></tr>"
	next
	if rst("剩余股份")<>0 then msg=msg&"<tr bgcolor=ffff00><td colspan=2 align=center><input type=button onclick="&chr(34)&"javascript:petition('"&stock&"');"&chr(34)&" value=' 申 购 '></td></tr>"
end if
msg=msg&"</table><hr bgcolor=red size=1><table width=80% align=center border=1 bordercolorlight=000000 cellspacing=0 cellpadding=2 bordercolordark=FFFFFF ><tr><td align=center colspan=2>主要持股人</td></tr>"
rst.Close
rst.Open "select top 10 * from 持股 where 名称='"&stock&"' and 股权>0 order by 股权 desc",conn
do while not (rst.EOF or rst.BOF)
	msg=msg&"<tr><td>"&rst("持股人")&"</td><td align=right>"&rst("股权")&"股</td></tr>"
	rst.MoveNext
loop
msg=msg&"</table></body>"
rst.Close
set rst=nothing
conn.Close
set conn=nothing
Response.Write msg
%>