<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
username=session("yx8_mhjh_username")
if username<>Application("yx8_mhjh_admin") then Response.Redirect "../../error.asp?id=046"
if username=Application("yx8_mhjh_admin")  then opt="<td><a href='modsong.asp?opt=新增&id=-1'>新&nbsp;增&nbsp;歌&nbsp;曲</a></td>"
msg="<head><link rel='stylesheet' href='../../chatroom/css.css'></head><body oncontextmenu=self.event.returnValue=false topmargin=0 background='../../chatroom/bg1.gif'><p align=center>歌曲管理</p><hr><table border=3 width=90% align=center><tr bgcolor=FFFF00 align=center><td>歌曲网址</td><td>歌曲名字</td><td>歌手</td><td>添加人</td><td>点击</td>"&opt&"</tr>"
%>
<!--#include file="data.asp"-->
<%
rs.Open"SELECT * FROM 歌曲  order by 点击 desc",conn,3
 do while not (rs.EOF or rs.BOF)
	id=rs("id")
	if username=Application("yx8_mhjh_admin") then opt="<TD><a href='modsong.asp?opt=修改&id="&id&"'>修改</a> | <a href='modsong.asp?opt=删除&id="&id&"'>删除</a></td>"
	msg=msg&"<tr><td>"&rs("地址")&"</td><td>"&rs("歌名")&"</td><td>"&rs("歌手")&"</td><td>"&rs("添加人")&"</td><td>"&rs("点击")&"</td>"&opt&"</tr>"
	rs.MoveNext
loop
rs.Close 
set rs=nothing
conn.Close
set conn=nothing
msg=msg&"</table></body>"
Response.Write msg
%>




