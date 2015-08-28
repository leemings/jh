<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
username=session("yx8_mhjh_username")
usergrade=session("yx8_mhjh_usergrade")
usercorp=session("yx8_mhjh_usercorp")
msg="<head><link rel='stylesheet' href='../style.css'><title>七门八派</title></head><body oncontextmenu=self.event.returnValue=false topmargin=0 background='../chatroom/bg1.gif'><p align=center size=4>七门八派(10天内10级以上的弟子不达到10个以上的门派，将被系统删除)</p><table border=1 cellspacing=0 cellpadding=2  bordercolorlight=#993300 bordercolordark=#FFFFFF align=center><tr bgcolor=ffffff align=center><td>门派</td><td>工资</td><td width='50%'>简介</td><td>弟子数</td><td>掌门</td>"&opt&"</tr>"
set conn=server.CreateObject("adodb.connection")
conn.Open Application("yx8_mhjh_connstr")
set rst=server.CreateObject("adodb.recordset")
set rst2=server.CreateObject("adodb.recordset")
rst.Open "门派",conn
do while not (rst.EOF or rst.BOF)
corpid=rst("id")
corp=rst("门派")
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

