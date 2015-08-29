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
num=Request.Form("num")
if instr(stock,"'")<>0 then 
	Response.Write "<script Language=Javascript>alert('提示：参数错误！');location.href = 'javascript:history.go(-1)'</script>"
	Response.end 
end if
if not isnumeric(num) then 
	Response.Write "<script Language=Javascript>alert('提示：参数错误！');location.href = 'javascript:history.go(-1)'</script>"
	Response.end 
end if
num=abs(int(clng(num)))
if num<1 then num=1
msg="<head><link rel='stylesheet' href='../style.css'><script language=javascript>setTimeout('history.back();',3000);</script></head><body oncontextmenu=self.event.returnValue=false topmargin=0 bgcolor='#339966' text='FF0000'><h3>申购原始股</h3><hr>"
set conn=server.CreateObject("adodb.connection")
set rs=server.CreateObject("adodb.recordset")
conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("stock.mdb")
set conn1=server.CreateObject("adodb.connection")
set rs1=server.CreateObject("adodb.recordset")
conn1.open Application("sjjh_usermdb")
rs1.Open "select 银两 from 用户 where 姓名='"&sjjh_name&"'",conn1,2,2
if rs1.EOF or rs1.BOF then 
	set rs=nothing
	conn.close
	set conn=nothing
	rs1.close
	set rs1=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：数据库查询错误！');location.href = 'javascript:history.go(-1)'</script>"
	Response.end 
end if
umoney=rs1("银两")
rs1.Close
rs.Open "select * from 股票 where 股票名称='"&stock&"'",conn,2,2
if rs.EOF or rs.BOF then 
	set rs1=nothing
	conn1.close
	set conn1=nothing
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：对不起，没有此种股票');location.href = 'javascript:history.go(-1)'</script>"
	Response.end 
end if
uprice=rs("发行价")
unum=umoney\uprice
uallownum=rs("剩余股份")
if num>uallownum then
	set rs1=nothing
	conn1.close
	set conn1=nothing
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：剩余股票只有:"&uallownum&"股！');location.href = 'javascript:history.go(-1)'</script>"
	Response.end 
end if
if uallownum<unum then uallownum=unum
	if num>unum then 
		set rs1=nothing
		conn1.close
		set conn1=nothing
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		Response.Write "<script Language=Javascript>alert('提示：对不起，你的钱好象不够哟！');location.href = 'javascript:history.go(-1)'</script>"
		Response.end 
	end if
if num>uallownum then num=uallownum
umoney=num*uprice
rs.Close
rs.Open "select * from 持股 where 持股人='"&sjjh_name&"' and 名称='"&stock&"'",conn,2,2
if rs.EOF or rs.BOF then
	conn.Execute "insert into 持股(持股人,时间,名称,股权,出售,标价,数量) values('"&sjjh_name&"',now(),'"&stock&"',"&num&",0,"&uprice&",0)"
else
	conn.Execute "update 持股 set 股权=股权+"&num&" where 持股人='"&sjjh_name&"' and 名称='"&stock&"'"
end if
	conn.Execute "update 股票 set 剩余股份=剩余股份-"&num&" where 股票名称='"&stock&"'"
	conn1.Execute "update 用户 set 银两=银两-"&umoney&" where 姓名='"&sjjh_name&"'"
rs.Close
set rs=nothing
if conn.Errors.Count>0 then
msg=msg&"<p align=center>交易<font color=FF0000>失败</font>,三秒钟后自动返回<br><a href='javascript:history.back();'>返回</a></p><script language=javascript>parent.stockfrm.location.reload();</script></body>"
conn.RollbackTrans
else
msg=msg&"<p align=center>交易<font color=0000FF>完成</font>,三秒钟后自动返回<br><a href='javascript:history.back();'>返回</a></p><script language=javascript>parent.stockfrm.location.reload();</script></body>"
end if 
conn.Close
set conn=nothing
Response.Write msg
%>