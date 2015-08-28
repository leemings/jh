
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
username=session("yx8_mhjh_username")
if username="" then Response.Redirect "../../error.asp?id=016"
stock=Request.form("stock")
unum=Request.Form("unum")
uprice=Request.Form("uprice")
if instr(stock,"'")<>0 then
	Response.Write "<script Language=Javascript>alert('提示：参数错误！');location.href = 'javascript:history.go(-1)'</script>"
	Response.end 
end if
if not(isnumeric(unum) and isnumeric(uprice)) then 
	Response.Write "<script Language=Javascript>alert('提示：参数错误！');location.href = 'javascript:history.go(-1)'</script>"
	Response.end 
end if
unum=clng(unum)
if unum<1 then unum=1
uprice=clng(uprice)
if uprice<1 then uprice=1
uallnum=unum
set conn=server.CreateObject("adodb.connection")
set rs=server.CreateObject("adodb.recordset")
conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("stock.mdb")
set conn1=server.CreateObject("adodb.connection")
conn1.open Application("yx8_mhjh_connstr")
rs.Open "select * from 持股 where 持股人='"&username&"' and 名称='"&stock&"' and 股权>="&unum,conn
if rs.EOF or rs.BOF then 
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	conn1.close
	set conn1=nothing
	Response.Write "<script Language=Javascript>alert('提示：对不起,你没有足够数量的此种股票可供出售！');location.href = 'javascript:history.go(-1)'</script>"
	Response.end 
end if
rs.Close
rs.Open "select * from 持股 where 出售=False and 数量>0 and 名称='"&stock&"' and 标价>="&uprice&" order by 时间"
do while not rs.EOF
	num=rs("数量")
	st=rs("持股人")
	if num<unum then
		unum=unum-num
	elseif num>=unum then
		num=unum
		unum=0 
	end if 
	conn.Execute "update 持股 set 股权=股权+"&num&",数量=数量-"&num&" where 持股人='"&st&"' and 名称='"&stock&"'"
	if unum=0 then exit do
	rs.MoveNext
loop
rs.Close
set rs=nothing
uright=uallnum-unum
if uright<>0 then conn.Execute "update 股票 set 最近成交价=现价,现价="&uprice&" where 股票名称='"&stock&"'"
conn.Execute "update 持股 set 股权=股权-"&uright&",出售=True,标价="&uprice&",数量="&unum&" where 持股人='"&username&"' and 名称='"&stock&"'"
conn1.Execute "update 用户 set 银两=银两+"&uright*uprice&" where 姓名='"&username&"'"
conn.Close
set conn=nothing
Response.Write "<head><link rel='stylesheet' href='../../style.css'><script language=javascript>top.location.reload();</script></head><body oncontextmenu=self.event.returnValue=false topmargin=0 bgcolor='#339966'><h2>实时行情</h2><hr><p align=center>交易完成三秒钟后自动返回<br><a href='javascript:history.go(-2)'> 返 回 </a></p></body>"
%>