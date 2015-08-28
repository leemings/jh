<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
username=session("yx8_mhjh_username")
adminname=Application("yx8_mhjh_admin")
if username<>adminname then Response.Redirect "../error.asp?id=046"
if session("jxqy_adminbok")="" or session("yx8_mhjh_usercorp")<>"官府" then Response.Redirect "../exit.asp"
msg="<head><link rel='stylesheet' href='../chatroom/css.css'></head><body oncontextmenu=self.event.returnValue=false topmargin=0 background='../chatroom/bg1.gif'><table width=100% border=1 bordercolorlight=000000 cellspacing=0 cellpadding=2 bordercolordark=FFFFFF bgcolor=FFB442><tr align=center>"
pattern=Request.QueryString("pattern")
if pattern="" then pattern="武器"
patternarr=array("头盔","盔甲","武器","防具","饰品","暗器")
for i=0 to 5
if pattern=patternarr(i) then
msg=msg&"<td bgcolor=ffFF00><a href='modarm1.asp?pattern="&patternarr(i)&"' onmouseover=""window.status='"&patternarr(i)&"类';return true;"" onmouseout=""window.status='';return true;"">"&patternarr(i)&"</a></td>"
else
msg=msg&"<td><a href='modarm1.asp?pattern="&patternarr(i)&"' onmouseover=""window.status='"&patternarr(i)&"类';return true;"" onmouseout=""window.status='';return true;"">"&patternarr(i)&"</a></td>"
end if
next
if username=adminname  then opt="<td><a href='modarm2.asp?opt=新增&id=-1&pattern="&pattern&"'>增加新"&pattern&"</a></td>"
msg=msg&"</tr></table><p align=center>武器管理</p><hr><table border=3 width=80% align=center><tr bgcolor=FFFF00 align=center><td>名称</td><td>攻击</td><td>防御</td><td>特效</td><td>价格</td>"&opt&"</tr>"
set conn=server.CreateObject("adodb.connection")
conn.Open Application("yx8_mhjh_connstr")
set rst=server.CreateObject("adodb.recordset")
rst.Open "select ID,名称,攻击,防御,特效,价格 from 商店 where 属性='"&pattern&"' order by 价格",conn
do while not (rst.EOF or rst.BOF)
id=rst("id")
if username=adminname then opt="<TD><a href='modarm2.asp?opt=修改&id="&id&"'>修改</a> | <a href='modarm2.asp?opt=删除&id="&id&"'>删除</a></td>"
msg=msg&"<tr><td>"&rst("名称")&"</td><td>"&rst("攻击")&"</td><td>"&rst("防御")&"</td><td>"&rst("特效")&"</td><td>"&rst("价格")&"</td>"&opt&"</tr>"
rst.MoveNext
loop
rst.Close
set rst=nothing
conn.Close
set conn=nothing
msg=msg&"</table></body>"
Response.Write msg
%>
