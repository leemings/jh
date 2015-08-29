<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.Buffer=true
sjjh_id=Session("sjjh_id")
sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
if sjjh_name="" then Response.Redirect "../error.asp?id=440"
if session("sjjh_adminok")<>true then Response.Redirect "../chat/readonly/bomb.htm"
if sjjh_grade<>10 or instr(Application("sjjh_admin"),sjjh_name)=0  then Response.Redirect "../error.asp?id=439"
if sjjh_name<>Application("sjjh_user") then Response.Redirect "../chat/readonly/bomb.htm"
if sjjh_name<>Application("sjjh_user") then 
	Response.Write "<script Language=Javascript>alert('提示：你不是正站长，你不能操作！');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
stock=Trim(Request.Form("stock"))
price=Request.Form("price")
num=Request.Form("num")
on error resume next
set conn=server.CreateObject("adodb.connection")
set rs=server.CreateObject("adodb.recordset")
conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("../../gupiao/stock.mdb")
rs.Open "select * from 股票 where 股票名称='"&stock&"'",conn
if not(rs.EOF or rs.BOF) then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：对不起,请重新为您的新股取个名字好吗?');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
rs.Close
set rs=nothing
conn.BeginTrans
conn.Execute "Insert Into 股票(股票名称,流通股总量,剩余股份,发行价,历史高价,历史低价,最近成交价,现价) values('"&stock&"',"&num&","&num&","&price&","&price&","&price&","&price&","&price&")"
if conn.Errors.Count>0 then
	conn.RollbackTrans
	Response.Write "<script Language=Javascript>alert('提示：数据库错误?');location.href = 'javascript:history.go(-1)';</script>"
	Response.end 
else
	conn.CommitTrans
end if
conn.Close
set conn=nothing
Response.Write "<script Language=Javascript>alert('提示：股票增加完成！');location.href = 'javascript:history.go(-1)';</script>"
%>