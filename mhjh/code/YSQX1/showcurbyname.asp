<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
username=session("yx8_mhjh_username")
adminname=Application("yx8_mhjh_admin")
if username<>adminname then Response.Redirect "../error.asp?id=046"
if session("jxqy_adminbok")="" or session("yx8_mhjh_usercorp")<>"官府" then Response.Redirect "../exit.asp"
uname=Request("uname")
if uname="" then Response.Redirect "../error.asp?id=024"
ucur=request("ucur")
if ucur="" then ucur="药品"
set conn=server.CreateObject("adodb.connection")
conn.Open Application("yx8_mhjh_connstr")
set rst=server.CreateObject("adodb.recordset")
rst.Open "select * from 物品 where 所有者='"&uname&"' and 属性='"&ucur&"' and 数量>0 order by 数量 desc",conn,1,3
Response.Write "<head><link rel='stylesheet' href='../chatroom/css.css'></head><body oncontextmenu=self.event.returnValue=false topmargin=0 background='../chatroom/bg1.gif' text='FF0000'><p align=center><h4>"&ucur&"所有人:<font color=FF0000>"&uname&"</font><h4></p><hr><table width=50% align=center border=3><tr bgcolor=FFFF00 align=center><td>名称</td><td>属性</td><td>数量</td><td>操作</td></tr>"
do while not(rst.EOF or rst.BOF)
Response.Write "<tr><td>"&rst("名称")&"</td><td>"&ucur&"</td><td align=right>"&rst("数量")&"</td><td><a href='chgcur.asp?id="&rst("ID")&"&uname="&uname&"&ucur="&ucur&"'>清零</a></td></tr>"
rst.MoveNext
loop
rst.Close
set rst=nothing
conn.Close
set conn=nothing
Response.Write "</table><p align=center><a href="&chr(34)&"javascript:location.replace('showuser.asp?search="&uname&"');"&chr(34)&">返回</a></p></body>"
%>
