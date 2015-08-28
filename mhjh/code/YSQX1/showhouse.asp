<%
username=session("yx8_mhjh_username")
adminname=Application("yx8_mhjh_admin")
if username<>adminname then Response.Redirect "../error.asp?id=046"
if session("jxqy_adminbok")="" or session("yx8_mhjh_usercorp")<>"官府" then Response.Redirect "../exit.asp"
if username=adminname  then opt="<td><a href='modhouse.asp?opt=新增&id=-1'>新&nbsp;增&nbsp;房&nbsp;屋</a></td>"
msg="<head><link rel='stylesheet' href='../chatroom/css.css'></head><body oncontextmenu=self.event.returnValue=false topmargin=0 background='../chatroom/bg1.gif'><p align=center>房屋管理</p><hr><table border=3 width=90% align=center><tr bgcolor=FFFF00 align=center><td>户主</td><td>伴侣</td><td>图片</td><td>类型</td><td>售价</td><td>位置</td><td>街道</td><td>状态</td><td>银两</td><td>伴侣银两</td>"&opt&"</tr>"
%>
<!--#include file="../fw/data1.asp"--><%
sql="select * from 房屋 order by 售价"
set rst=conn.execute(sql)
 do while not (rst.EOF or rst.BOF)
	id=rst("id")
	if username=adminname then opt="<TD><a href='modhouse.asp?opt=修改&id="&id&"'>修改</a> | <a href='modhouse.asp?opt=删除&id="&id&"'>删除</a></td>"
	msg=msg&"<tr><td>"&rst("户主")&"</td><td>"&rst("伴侣")&"</td><td>"&rst("图片")&"</td><td>"&rst("类型")&"</td><td>"&rst("售价")&"</td><td>"&rst("位置")&"</td><td>"&rst("街道")&"</td><td>"&rst("状态")&"</td><td>"&rst("银两")&"</td><td>"&rst("伴侣银两")&"</td>"&opt&"</tr>"
	rst.MoveNext
loop
rst.Close 
set rst=nothing
conn.Close
set conn=nothing
msg=msg&"</table></body>"
Response.Write msg
%>