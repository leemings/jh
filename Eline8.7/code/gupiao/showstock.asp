<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
if sjjh_name="" then Response.Redirect "../error.asp?id=440"
stock=Request.QueryString("stock")
if stock="" then stock="『快乐江湖』"
msg="<head><link rel='stylesheet' href='style.css'><script language=javascript>function petition(stock){parent.infofrm.location.replace('petition.asp?stock='+stock);}</script></head><body oncontextmenu=self.event.returnValue=false topmargin=0 bgcolor='#339966' text='FF0000'><center><font color=red><h3>"&stock&"</h3></font></center><table width=80% border=1 bordercolorlight=000000 cellspacing=0 cellpadding=2 bordercolordark=FFFFFF align=center>"
set conn=server.CreateObject("adodb.connection")
set rs=server.CreateObject("adodb.recordset")
conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("stock.mdb")
rs.Open "select * from 股票 where 股票名称='"&stock&"'",conn,2,2
if rs.EOF or rs.BOF then
	msg=msg&"<tr><td>你要找的股票没有找到</td></tr>"
else
	for i=1 to rs.Fields.Count -1
		msg=msg&"<tr><td bgcolor=00ff00>"&rs.Fields(i).name&"</td><td align=right>"&rs.Fields(i).value&"</td></tr>"
	next
	if rs("剩余股份")<>0 then msg=msg&"<tr bgcolor=ffff00><td colspan=2 align=center><input type=button onclick="&chr(34)&"javascript:petition('"&stock&"');"&chr(34)&" value=' 申 购 '></td></tr>"
end if
msg=msg&"</table><hr bgcolor=red size=1><table width=80% align=center border=1 bordercolorlight=000000 cellspacing=0 cellpadding=2 bordercolordark=FFFFFF ><tr><td align=center colspan=2>主要持股人</td></tr>"
rs.Close
rs.Open "select top 10 * from 持股 where 名称='"&stock&"' and 股权>0 order by 股权 desc",conn
do while not (rs.EOF or rs.BOF)
	msg=msg&"<tr><td>"&rs("持股人")&"</td><td align=right>"&rs("股权")&"股</td></tr>"
	rs.MoveNext
loop
msg=msg&"</table></body>"
rs.Close
set rs=nothing
conn.Close
set conn=nothing
Response.Write msg
%>