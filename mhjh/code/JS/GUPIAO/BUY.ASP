<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
username=session("yx8_mhjh_username")
if username="" then Response.Redirect "../../error.asp?id=016"
stock=Request.QueryString("stock")
if instr(stock,"'")<>0 then
	Response.Write "<script Language=Javascript>alert('提示：参数错误！！');location.href = 'javascript:history.go(-1)'</script>"
	Response.end 
end if
set conn=server.CreateObject("adodb.connection")
set rs=server.CreateObject("adodb.recordset")
conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("stock.mdb")
rs.Open "select * from 持股 where 持股人='"&username&"' and 名称='"&stock&"' and 数量>0",conn,2,2
if not (rs.EOF or rs.BOF) then 
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：对不起，无法支持对同一只股票进行两次操作，请先撤销上次的操作要求！');location.href = 'javascript:history.go(-1)'</script>"
	Response.end
end if 
rs.close
rs.open "select * from 股票 where 股票名称='"&stock&"'",conn,2,2
if rs.EOF or rs.BOF then 
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：数据库查询错误！');location.href = 'javascript:history.go(-1)'</script>"
	Response.end 
end if
uprice=rs("现价")
gpmc=rs("股票名称")
lt=rs("流通股总量")
rs.Close
set rs=nothing
conn.close
set conn=nothing
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("yx8_mhjh_connstr")
rs.Open "select 银两 from 用户 where 姓名='"&username&"'",conn,2,2
umoney=rs("银两")
unum=umoney\uprice
rs.close
set rs=nothing
conn.Close
set conn=nothing
Response.Write "<head><link rel='stylesheet' href='../../style.css'></head><body oncontextmenu=self.event.returnValue=false topmargin=0 bgcolor='#339966' text='FF0000'><h3>实时行情</h3><hr><table bgcolor=CCCCCC><form action=buynow.asp method=post><table><tr><td bgcolor=FFFF00 align=center>收购股票</td></tr><tr><td><input type=text size=14 maxlength=7 value='"&gpmc&"' readonly name='stock'></tr><tr><td>可购:1"&unum&"股<br>流通总量："&lt&"股</td></tr><tr><td><input type=text value='"&unum&"' maxlength=9 size=9 name='unum'></td></tr><tr><td align=center>收购价</td></tr><tr><td><input type=text value='"&uprice&"' maxlength=8 size=9 name='uprice'></td></tr><tr><td align=center><input type=submit value=' 购 入 ' id=submit1 name=submit1> <input type=button value=' 返 回 ' onclick='javascript:history.back();'></td></tr></table></form></body>"
%>