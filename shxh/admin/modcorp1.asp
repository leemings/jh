<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
username=session("Ba_jxqy_username")
usergrade=session("Ba_jxqy_usergrade")
usercorp=session("Ba_jxqy_usercorp")
bgcolor=Application("Ba_jxqy_backgroundcolor")
bgimage=Application("Ba_jxqy_backgroundimage")
if usergrade=Application("Ba_jxqy_allright") and usercorp="官府" then opt="<td><a href='modcorp2.asp?opt=新增&corpid=-1'>新&nbsp;增&nbsp;门&nbsp;派</a></td>"
msg="<head><link rel='stylesheet' href='../chatroom/css.css'></head><body oncontextmenu=self.event.returnValue=false topmargin=0 background='"&bgimage&"' bgcolor='"&bgcolor&"'><p align=center>门派管理</p><hr><table border=3 width=100% align=center><tr bgcolor=FFFF00 align=center><td>门派</td><td>工资</td><td width='50%'>简介</td><td>弟子数</td><td>掌门</td>"&opt&"</tr>"
set conn=server.CreateObject("adodb.connection")
conn.Open Application("Ba_jxqy_connstr")
set rst=server.CreateObject("adodb.recordset")
set rst2=server.CreateObject("adodb.recordset")
rst.Open "门派",conn
do while not (rst.EOF or rst.BOF)
	corpid=rst("id")
	corp=rst("门派")
	if usergrade=Application("Ba_jxqy_allright") and usercorp="官府" then opt="<TD><a href='modcorp2.asp?opt=修改&corpid="&corpid&"'>修改</a> | <a href='modcorp2.asp?opt=删除&corpid="&corpid&"'>删除</a></td>"
	if usercorp=corp then 
		msg=msg&"<tr><td><a href='showcorp.asp'>"&corp&"</a></td>"
	else
		msg=msg&"<tr><td>"&corp&"</td>"
	end if		
	rst2.Open "select count(*) as cnumber from 用户 where 门派='"&corp&"'",conn
	cnumber=rst2("cnumber")
	rst2.Close
	rst2.Open "select 姓名 from 用户 where 门派='"&corp&"' and 身份='掌门'"
	if rst2.EOF or rst2.BOF then 
		cname="&nbsp;"
	else
		cname=rst2("姓名")
	end if
	msg=msg&"<td align=right>"&rst("工资系数")&"</td><td>"&rst("简介")&"</td><td>"&cnumber&"</td><td>"&cname&"</td>"&opt&"</tr>"
	rst2.Close
	rst.MoveNext
loop
rst.Close
set rst2=nothing
set rst=nothing
conn.Close
set conn=nothing
msg=msg&"</table></body>"
Response.Write msg
%>