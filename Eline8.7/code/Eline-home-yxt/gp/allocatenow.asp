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
money=Request.Form("money")
set conn=server.CreateObject("adodb.connection")
set rs=server.CreateObject("adodb.recordset")
conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("../../gupiao/stock.mdb")
set conn1=server.CreateObject("adodb.connection")
conn1.Open Application("sjjh_usermdb")
rs.Open "select * from 持股 where 名称='"&stock&"'",conn
do while not (rs.EOF or rs.BOF)
conn1.Execute "update 用户 set 银两=银两+"&rs("股权")*money&" where 姓名='"&rs("持股人")&"'"
rs.MoveNext
loop
rs.Close
set rs=nothing
Response.Write "<script Language=Javascript>alert('提示：股票分红完成！');location.href = 'javascript:history.go(-1)';</script>"
%>