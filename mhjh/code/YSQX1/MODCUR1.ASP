<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
username=session("yx8_mhjh_username")
adminname=Application("yx8_mhjh_admin")
if username<>adminname then Response.Redirect "../error.asp?id=046"
if session("jxqy_adminbok")="" or session("yx8_mhjh_usercorp")<>"官府" then Response.Redirect "../exit.asp"
if username=adminname then opt="<td><a href='modcur2.asp?opt=新增&id=-1'>新&nbsp;增&nbsp;药&nbsp;品</a></td>"
msg="<head><link rel='stylesheet' href='../chatroom/css.css'></head><body oncontextmenu=self.event.returnValue=false topmargin=0 background='../chatroom/bg1.gif'><p align=center>药品管理</p><hr><table border=3 width=80% align=center><tr bgcolor=FFFF00 align=center><td>名称</td><td>体力</td><td>内力</td><td>特效</td><td>价格</td>"&opt&"</tr>"
set conn=server.CreateObject("adodb.connection")
conn.Open Application("yx8_mhjh_connstr")
set rst=server.CreateObject("adodb.recordset")
rst.Open "select ID,名称,体力,内力,特效,价格 from 商店 where 属性='药品' order by 价格",conn
do while not (rst.EOF or rst.BOF)
id=rst("id")
if username=adminname then opt="<TD><a href='modcur2.asp?opt=修改&id="&id&"'>修改</a> | <a href='modcur2.asp?opt=删除&id="&id&"'>删除</a></td>"
msg=msg&"<tr><td>"&rst("名称")&"</td><td>"&rst("体力")&"</td><td>"&rst("内力")&"</td><td>"&rst("特效")&"</td><td>"&rst("价格")&"</td>"&opt&"</tr>"
rst.MoveNext
loop
rst.Close
set rst=nothing
conn.Close
set conn=nothing
msg=msg&"</table></body>"
Response.Write msg
%>
