<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no cache"
Response.Expires=-1
username=session("Ba_jxqy_username")
if username="" then Response.Redirect "../error.asp?id=016"
stock=Request.form("stock")
unum=Request.Form("unum")
uprice=Request.Form("uprice")
if instr(stock,"'")<>0 then Response.Redirect "../error.asp?id=024"
if not(isnumeric(unum) and isnumeric(uprice)) then Response.Redirect "../error.asp?id=024"
unum=clng(unum) 
if unum<1 then unum=1
uprice=clng(uprice)
if uprice<1 then uprice=1
umoney=uprice*unum
uallnum=unum
bgcolor=Application("Ba_jxqy_backgroundcolor")
bgimage=Application("Ba_jxqy_backgroundimage")
set conn=server.CreateObject("adodb.connection")
conn.Open Application("Ba_jxqy_connstr")
set rst=server.CreateObject("adodb.recordset")
rst.open "select * from 用户 where 姓名='"&username&"' and 存款>="&umoney,conn
if rst.EOF or rst.BOF then Response.Redirect "../error.asp?id=030"
rst.Close
rst.Open "select * from 持股 where 持股人='"&username&"' and 名称='"&stock&"' and 数量>0",conn
if not (rst.EOF or rst.BOF) then  Response.Redirect "../error.asp?id=067"
rst.Close
rst.Open "select * from 持股 where 出售=True and 数量>0 and 名称='"&stock&"' and 标价<="&uprice&" order by 时间"
conn.BeginTrans
do while not rst.EOF
	num=rst("数量")
	st=rst("持股人")
	if num<unum then
		unum=unum-num
	elseif num>=unum then
		num=unum
		unum=0	
	end if	
	conn.Execute "update 持股 set 股权=股权-"&num&",数量=数量-"&num&" where 持股人='"&st&"' and 名称='"&stock&"'"
	conn.Execute "update 用户 set 存款=存款+"&num*uprice&" where 姓名='"&st&"'"
	if unum=0 then exit do
	rst.MoveNext 
loop
rst.Close
uright=uallnum-unum
if uright<>0 then conn.Execute "update 股票 set 最近成交价=现价,现价="&uprice&" where 股票名称='"&stock&"'"
rst.Open "select * from 持股 where 持股人='"&username&"' and 名称='"&stock&"'",conn
if rst.EOF or rst.BOF then
	conn.Execute "insert into 持股(持股人,时间,名称,股权,出售,标价,数量) values('"&username&"',now(),'"&stock&"',"&uright&",False,"&uprice&","&unum&")"
else
	conn.Execute "update 持股 set 股权=股权+"&uright&",出售=False,标价="&uprice&",数量="&unum&" where 持股人='"&username&"' and 名称='"&stock&"'"
end if
conn.Execute "update 用户 set 存款=存款-"&umoney&" where 姓名='"&username&"'"
conn.CommitTrans
set rst=nothing
conn.Close
set conn=nothing
Response.Write "<head><link rel='stylesheet' href='../chatroom/css.css'><script language=javascript>parent.pricefrm.location.reload();</script></head><body oncontextmenu=self.event.returnValue=false topmargin=0 background='"&bgimage&"' bgcolor='"&bgcolor&"'><h2>实时行情</h2><hr><table width='100%'><tr><td colspan=2 align=center bgcolor=00ff00>"&stock&"</td></tr><tr><td >单价</td><td align=right>"&uprice&"</td></tr><tr><td>要求</td><td align=right>"&uallnum&"</td></tr><tr><td>成交</td><td align=right>"&uright&"</td></tr><tr><td>金额</td><td align=right>"&uright*uprice&"</td></tr><tr><td>押金</td><td align=right>"&unum*uprice&"</td><tr><td>实收</td><td align=right>"&umoney&"</td></tr></tr></table><p align=center><a href='javascript:history.go(-2)'> 返 回 </a></p></body>"
%>