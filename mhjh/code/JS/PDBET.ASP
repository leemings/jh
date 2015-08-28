<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
username=Session("yx8_mhjh_username")
if username="" then Response.Redirect "../error.asp?id=016"
set conn=server.CreateObject("adodb.connection")
conn.Open Application("yx8_mhjh_connstr")
set rst=server.CreateObject("adodb.recordset")
rst.Open "select 银两 from 用户 where 姓名='"&username&"'",conn
if rst.EOF or rst.BOF then Response.Redirect "../error.asp?id=100"
silver=rst("银两")
rst.Close
set conn=nothing
bet=100
if silver<bet then silver=bet
response.write "<html><head><link rel=stylesheet href='../style.css'></head><body oncontextmenu=self.event.returnValue=false background='../chatroom/bg1.gif'><p align=center><form action='pdstart.asp' method=post >"&username&" 你的剑侠金币是："&silver&"两 下注：<input type=text maxlength=4 size=4 name=pcbet value='"&bet&"'> <input type=submit value=' 下 注 '> <input type=button value=' 返 回 ' onclick=javascript:parent.location.replace('dubo.asp');></form></p></body></html>"
%>

