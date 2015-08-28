<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
username=session("yx8_mhjh_username")
if username="" then Response.Redirect "../../error.asp?id=016"
stock=Request.QueryString("stock")
if stock="" then stock="爱火神话"
msg="<head><link rel='stylesheet' href='style.css'></head><body oncontextmenu=self.event.returnValue=false topmargin=0 bgcolor='#339966' text='FF0000'><h3>申购原始股</h3><hr>"
set conn=server.CreateObject("adodb.connection")
set rs=server.CreateObject("adodb.recordset")
conn.open Application("yx8_mhjh_connstr")
set conn1=server.CreateObject("adodb.connection")
set rs1=server.CreateObject("adodb.recordset")
conn1.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("stock.mdb")
rs.Open "select 银两 from 用户 where 姓名='"&username&"'",conn,2,2
umoney=rs("银两")
rs.close
rs1.Open "select * from 股票 where 股票名称='"&stock&"'",conn1
if rs1.EOF or rs1.BOF then 
	rs1.close
	set rs1=nothing
	conn1.close
	set conn1=nothing
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：数据库查询错误？！');location.href = 'javascript:history.go(-1)'</script>"
	Response.end 
end if
unum=umoney\rs1("发行价")
uallownum=unum
if uallownum>rs1("剩余股份") then uallownum=rs1("剩余股份")
	msg=msg&stock&"的原始股份还剩余"&rs1("剩余股份")&"股,你拥有江湖银两"&umoney&",<font color=ffffff>可申购"&uallownum&"股</font><form action='petitionnow.asp' method=post><table><tr><td>股票</td><td><input type=text value='"&stock&"' size=14 readonly name='stock'></td></tr><tr><td>数量</td><td><input type=text value='"&uallownum&"' size=9 name='num'></td></tr><tr><td align=center colspan=2><input type=submit value=' 申 购 '></td></tr></table></form>"
	rs1.Close
	conn.Close
	conn1.Close
	set rs=nothing
	set rs1=nothing
	set conn=nothing
	set conn1=nothing
	Response.Write msg
%>