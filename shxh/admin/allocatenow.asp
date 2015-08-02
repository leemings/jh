<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
bgcolor=Application("Ba_jxqy_backgroundcolor")
bgimage=Application("Ba_jxqy_backgroundimage")
username=session("Ba_jxqy_username")
usercorp=session("Ba_jxqy_usercorp")
usergrade=session("Ba_jxqy_usergrade")
stock=Trim(Request.Form("stock"))
money=Request.Form("money")
if username="" then Response.Redirect "../error.asp?id=016"
if not(usergrade=Application("Ba_jxqy_allright") and usercorp="官府") then Response.Redirect "../error.asp?id=046"
set conn=server.CreateObject("adodb.connection")
conn.Open Application("Ba_jxqy_connstr")
set rst=server.CreateObject("adodb.recordset")
rst.Open "select * from 持股 where 名称='"&stock&"'",conn
do while not (rst.EOF or rst.BOF)
	conn.Execute "update 用户 set 存款=存款+"&rst("股权")*money&" where 姓名='"&rst("持股人")&"'"
	rst.MoveNext
loop
rst.Close
set rst=nothing
conn.Execute "insert into 证交公告(时间,标题,内容) values(now(),'"&stock&"股现金分红','"&stock&"经营报表分析显示赢利,董事会决定现金分红,每股派息"&money&"')"
conn.Close
set conn=nothing
Response.Write "<head><title>"&Application("Ba_jxqy_systemname")&"</title><script language=javascript>setTimeout('history.go(-1);',3000);</script><link rel='stylesheet' href='../chatroom/css.css'></head><body oncontextmenu=self.event.returnValue=false  background='"&bgimage&"' bgcolor='"&bgcolor&"'><p align=center><font color=0000ff size=4>股息分红</font><br>股票分红完成.三秒钟后自动返回<br><a href='javascript:history.go(-1);'>返回</a></p>"
%>