<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no cache"
Response.Expires=-1
stock=Request.QueryString("stock")
if instr(stock,"'")<>0 then Response.Redirect "../error.asp?id=024"
username=session("Ba_jxqy_username")
if username="" then Response.Redirect "../error.asp?id=016"
bgcolor=Application("Ba_jxqy_backgroundcolor")
bgimage=Application("Ba_jxqy_backgroundimage")
set conn=server.CreateObject("adodb.connection")
conn.Open Application("Ba_jxqy_connstr")
set rst=server.CreateObject("adodb.recordset")
rst.Open "select * from 持股 where 持股人='"&username&"' and 名称='"&stock&"' and 数量>0",conn
if not (rst.EOF or rst.BOF) then  Response.Redirect "../error.asp?id=067"
rst.Close
rst.Open "select * from 股票 where 股票名称='"&stock&"'",conn
if rst.EOF or rst.BOF then Response.Redirect "../error.asp?id=100"
uprice=rst("现价")
rst.Close
rst.Open "select 存款 from 用户 where 姓名='"&username&"'",conn
umoney=rst("存款")
rst.Close
unum=umoney\uprice
set rst=nothing
conn.Close 
set conn=nothing
Response.Write "<head><link rel='stylesheet' href='../chatroom/css.css'></head><body oncontextmenu=self.event.returnValue=false topmargin=0 background='"&bgimage&"' bgcolor='"&bgcolor&"' text='FF0000'><h3>实时行情</h3><hr><table bgcolor=CCCCCC><form action=buynow.asp method=post><table><tr><td bgcolor=FFFF00 align=center>收购股票</td></tr><tr><td><input type=text size=14 value='"&stock&"' readonly name='stock'></tr><tr><td>可购:1-"&unum&"股</td></tr><tr><td><input type=text value='"&unum&"' size=9 name='unum'></td></tr><tr><td align=center>收购价</td></tr><tr><td><input type=text value='"&uprice&"' size=9 name='uprice'></td></tr><tr><td align=center><input type=submit value=' 购 入 ' id=submit1 name=submit1> <input type=button value=' 返 回 ' onclick='javascript:history.back();'></td></tr></table></form></body>"
%>