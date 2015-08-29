<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
if sjjh_name="" then Response.Redirect "../error.asp?id=440"
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
unum=int(abs(clng(unum)))
if unum<1 then unum=1
uprice=int(abs(clng(uprice)))
if uprice<1 then uprice=1
umoney=uprice*unum
uallnum=unum
set conn=server.CreateObject("adodb.connection")
set rs=server.CreateObject("adodb.recordset")
conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("stock.mdb")
rs.Open "select * from 股票 where 股票名称='"&stock&"'",conn
if rs.EOF or rs.BOF then 
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：对不起，没有此种股票！');location.href = 'javascript:history.go(-1)'</script>"
	Response.end 
end if
if rs("流通股总量")<unum then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：对不起，你要购买数已经超过了流通总量！');location.href = 'javascript:history.go(-1)'</script>"
	Response.end 
end if
rs.close
Set conn1=Server.CreateObject("ADODB.CONNECTION")
Set rs1=Server.CreateObject("ADODB.RecordSet")
conn1.open Application("sjjh_usermdb")
rs1.open "select * from 用户 where 姓名='"&sjjh_name&"' and 银两>="&umoney,conn1,2,2
if rs1.EOF or rs1.BOF then 
	rs1.close
	set rs1=nothing
	conn1.close
	set conn1=nothing
	Response.Write "<script Language=Javascript>alert('提示：对不起，你的钱好象不够哟！');location.href = 'javascript:history.go(-1)'</script>"
	Response.end 
end if
rs1.Close
rs.Open "select * from 持股 where 持股人='"&sjjh_name&"' and 名称='"&stock&"' and 数量>0",conn,2,2
if not (rs.EOF or rs.BOF) then 
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：对不起，无法支持对同一只股票进行两次操作，请先撤销上次的操作要求！');location.href = 'javascript:history.go(-1)'</script>"
	Response.end 
end if
rs.Close
rs.Open "select * from 持股 where 出售=True and 数量>0 and 名称='"&stock&"' and 标价<="&uprice&" order by 时间",conn
do while not rs.EOF
	num=rs("数量")
	st=rs("持股人")
	if num<unum then
		unum=unum-num
	elseif num>=unum then
		num=unum
	unum=0 
	end if 
	conn.Execute "update 持股 set 股权=股权-"&num&",数量=数量-"&num&" where 持股人='"&st&"' and 名称='"&stock&"'"
	conn1.Execute "update 用户 set 银两=银两+"&num*uprice&" where 姓名='"&st&"'"
	if unum=0 then exit do
	rs.MoveNext
loop
rs.Close
uright=uallnum-unum
if uright<>0 then conn.Execute "update 股票 set 最近成交价=现价,现价="&uprice&" where 股票名称='"&stock&"'"
rs.Open "select * from 持股 where 持股人='"&sjjh_name&"' and 名称='"&stock&"'",conn,2,2
if rs.EOF or rs.BOF then
	conn.Execute "insert into 持股(持股人,时间,名称,股权,出售,标价,数量) values('"&sjjh_name&"',now(),'"&stock&"',"&uright&",False,"&uprice&","&unum&")"
else
	conn.Execute "update 持股 set 股权=股权+"&uright&",出售=False,标价="&uprice&",数量="&unum&" where 持股人='"&sjjh_name&"' and 名称='"&stock&"'"
end if
conn1.Execute "update 用户 set 银两=银两-"&umoney&" where 姓名='"&sjjh_name&"'"
set rs=nothing
set rs1=nothing
conn.Close
conn1.Close
set conn=nothing
set conn1=nothing
Response.Write "<head><link rel='stylesheet' href='../chat/readonly/style.css'><script language=javascript>parent.pricefrm.location.reload();</script></head><body oncontextmenu=self.event.returnValue=false topmargin=0 bgcolor='#339966'><h2>实时行情</h2><hr><table width='100%'><tr><td colspan=2 align=center bgcolor=00ff00>"&stock&"</td></tr><tr><td >单价</td><td align=right>"&uprice&"</td></tr><tr><td>要求</td><td align=right>"&uallnum&"</td></tr><tr><td>成交</td><td align=right>"&uright&"</td></tr><tr><td>金额</td><td align=right>"&uright*uprice&"</td></tr><tr><td>押金</td><td align=right>"&unum*uprice&"</td><tr><td>实收</td><td align=right>"&umoney&"</td></tr></tr></table><p align=center><a href='javascript:history.go(-2)'> 返 回 </a></p></body>"
%>