<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no cache"
Response.Expires=-1
stock=Request.form("stock")
num=Request.Form("num")
if instr(stock,"'")<>0 then Response.Redirect "../error.asp?id=024"
if not isnumeric(num) then Response.Redirect "../error.asp?id=024"
num=clng(num)
if num<1 then num=1
bgcolor=Application("Ba_jxqy_backgroundcolor")
bgimage=Application("Ba_jxqy_backgroundimage")
username=session("Ba_jxqy_username")
msg="<head><link rel='stylesheet' href='../style.css'><script language=javascript>setTimeout('history.back();',3000);</script></head><body oncontextmenu=self.event.returnValue=false topmargin=0 background='"&bgimage&"' bgcolor='"&bgcolor&"' text='FF0000'><h3>申购原始股</h3><hr>"
set conn=server.CreateObject("adodb.connection")
conn.Open Application("Ba_jxqy_connstr")
set rst=server.CreateObject("adodb.recordset")
rst.Open "select 存款 from 用户 where 姓名='"&username&"'",conn
if rst.EOF or rst.BOF then Response.Redirect "../error.asp?id=016"
umoney=rst("存款")
rst.Close
rst.Open "select * from 股票 where 股票名称='"&stock&"'",conn
if rst.EOF or rst.BOF then Response.Redirect "../error.asp?id=100"
uprice=rst("发行价")
unum=umoney\uprice
uallownum=rst("剩余股份")
if uallownum<unum then uallownum=unum
if num>unum then Response.Redirect "../error.asp?id=030"
if num>uallownum then num=uallownum
umoney=num*uprice
rst.Close
rst.Open "select * from 持股 where 持股人='"&username&"' and 名称='"&stock&"'",conn
conn.BeginTrans
if rst.EOF or rst.BOF then
	conn.Execute "insert into 持股(持股人,时间,名称,股权,出售,标价,数量) values('"&username&"',now(),'"&stock&"',"&num&",0,"&uprice&",0)"
else
	conn.Execute "update 持股 set 股权=股权+"&num&" where 持股人='"&username&"' and 名称='"&stock&"'"
end if
conn.Execute "update 股票 set 剩余股份=剩余股份-"&num&" where 股票名称='"&stock&"'"
conn.Execute "update 用户 set 存款=存款-"&umoney&" where 姓名='"&username&"'"
rst.Close
set rst=nothing
if conn.Errors.Count>0 then
	msg=msg&"<p align=center>交易<font color=FF0000>失败</font>,三秒钟后自动返回<br><a href='javascript:history.back();'>返回</a></p><script language=javascript>parent.stockfrm.location.reload();</script></body>"
	conn.RollbackTrans
else
	msg=msg&"<p align=center>交易<font color=0000FF>完成</font>,三秒钟后自动返回<br><a href='javascript:history.back();'>返回</a></p><script language=javascript>parent.stockfrm.location.reload();</script></body>"
	conn.CommitTrans
end if	
conn.Close
set conn=nothing
Response.Write msg
%>