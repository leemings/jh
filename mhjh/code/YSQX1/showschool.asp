<%
username=session("yx8_mhjh_username")
adminname=Application("yx8_mhjh_admin")
if username<>adminname then Response.Redirect "../error.asp?id=046"
if session("jxqy_adminbok")="" or session("yx8_mhjh_usercorp")<>"官府" then Response.Redirect "../exit.asp"
if username=adminname then opt="<td><a href='modschool.asp?opt=新增&id=-1'>新&nbsp;增&nbsp;书&nbsp;籍</a></td>"
msg="<head><link rel='stylesheet' href='../chatroom/css.css'></head><body oncontextmenu=self.event.returnValue=false topmargin=0 background='../chatroom/bg1.gif'><p align=center>私塾管理</p><hr><table border=3 width=80% align=center><tr bgcolor=FFFF00 align=center><td>书名</td><td>作者</td><td>说明</td><td>银两</td>"&opt&"</tr>"
set conn=server.CreateObject("adodb.connection")
conn.Open Application("yx8_mhjh_connstr")
set rst=server.CreateObject("adodb.recordset")
rst.Open "select * from 私塾 order by 银两",conn
do while not (rst.EOF or rst.BOF)
id=rst("id")
if username=adminname then opt="<TD><a href='modschool.asp?opt=修改&id="&id&"'>修改</a> | <a href='modschool.asp?opt=删除&id="&id&"'>删除</a></td>"
msg=msg&"<tr><td>"&rst("书名")&"</td><td>"&rst("作者")&"</td><td>"&rst("说明")&"</td><td>"&rst("银两")&"</td>"&opt&"</tr>"
rst.MoveNext
loop
rst.Close
set rst=nothing
conn.Close
set conn=nothing
msg=msg&"</table></body>"
Response.Write msg
%>
