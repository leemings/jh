<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
username=session("Ba_jxqy_username")
usergrade=session("Ba_jxqy_usergrade")
usercorp=session("Ba_jxqy_usercorp")
bgcolor=Application("Ba_jxqy_backgroundcolor")
bgimage=Application("Ba_jxqy_backgroundimage")
if usergrade=Application("Ba_jxqy_allright") and usercorp="官府" then opt="<td><a href='modmill2.asp?opt=新增&id=-1'>新增打工项目</a></td>"
msg="<head><link rel='stylesheet' href='../chatroom/css.css'></head><body oncontextmenu=self.event.returnValue=false topmargin=0 background='"&bgimage&"' bgcolor='"&bgcolor&"'><p align=center>打工管理</p><hr><table border=3 width=80% align=center><tr bgcolor=FFFF00 align=center><td>工作</td><td>说明</td><td>银两</td>"&opt&"</tr>"
set conn=server.CreateObject("adodb.connection")
conn.Open Application("Ba_jxqy_connstr")
set rst=server.CreateObject("adodb.recordset")
rst.Open "select * from 打工 order by 银两",conn
do while not (rst.EOF or rst.BOF)
	id=rst("id")
	if usergrade=Application("Ba_jxqy_allright") and usercorp="官府" then opt="<TD><a href='modmill2.asp?opt=修改&id="&id&"'>修改</a> | <a href='modmill2.asp?opt=删除&id="&id&"'>删除</a></td>"
	msg=msg&"<tr><td>"&rst("工作")&"</td><td>"&rst("说明")&"</td><td>"&rst("银两")&"</td>"&opt&"</tr>"
	rst.MoveNext
loop
rst.Close 
set rst=nothing
conn.Close
set conn=nothing
msg=msg&"</table></body>"
Response.Write msg
%>