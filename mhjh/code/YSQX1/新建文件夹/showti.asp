<%
username=session("yx8_mhjh_username")
usergrade=session("yx8_mhjh_usergrade")
usercorp=session("yx8_mhjh_usercorp")
bgcolor=Application("yx8_mhjh_backgroundcolor")
bgimage=Application("yx8_mhjh_backgroundimage")
if usergrade=Application("yx8_mhjh_allright") and usercorp="官府" then opt="<td><a href='modti.asp?opt=新增&id=-1'>新&nbsp;增&nbsp;答&nbsp;题</a></td>"
msg="<head><link rel='stylesheet' href='../chatroom/css.css'></head><body oncontextmenu=self.event.returnValue=false topmargin=0 background='"&bgimage&"' bgcolor='"&bgcolor&"'><p align=center>答题管理</p><hr><table border=3 width=90% align=center><tr bgcolor=FFFF00 align=center><td>问题</td><td>答案</td><td>提供者</td><td>奖金</td><td>抢答</td>"&opt&"</tr>"
set conn=server.CreateObject("adodb.connection")
conn.Open Application("yx8_mhjh_connstr")
set rst=server.CreateObject("adodb.recordset")
rst.Open "select * from 答题",conn
do while not (rst.EOF or rst.BOF)
	id=rst("id")
	if usergrade=Application("yx8_mhjh_allright") and usercorp="官府" then opt="<TD><a href='modti.asp?opt=修改&id="&id&"'>修改</a> | <a href='modti.asp?opt=删除&id="&id&"'>删除</a></td>"
	msg=msg&"<tr><td>"&rst("问题")&"</td><td>"&rst("答案")&"</td><td>"&rst("提供者")&"</td><td>"&rst("奖金")&"</td><td>"&rst("抢答")&"</td>"&opt&"</tr>"
	rst.MoveNext
loop
rst.Close 
set rst=nothing
conn.Close
set conn=nothing
msg=msg&"</table></body>"
Response.Write msg
%>