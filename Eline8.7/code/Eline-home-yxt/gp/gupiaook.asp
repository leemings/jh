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
gupiao=request.form("gupiao")
zl=request.form("zl")
sy=request.form("sy")
fx=request.form("fx")
gj=request.form("gj")
dj=request.form("dj")
cjj=request.form("cjj")
xj=request.form("xj")
if gupiao="" or zl="" or sy="" or fx="" or gj="" or dj="" or cjj="" or xj="" then
Response.write "资料未填写完整!请重新输入"
Response.end 
end if
if request.form("submit1")="删除" then
Set conn=Server.CreateObject("ADODB.CONNECTION")
conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("../../gupiao/stock.mdb")
conn.execute("delete from 股票 where 股票名称='"&gupiao&"'")
conn.execute("delete from 持股 where 名称='"&gupiao&"'")
conn.close
set conn=nothing
Response.write "数据库操作完成！你已经成功删除了股票"
Response.end 
end if
if request.form("submit")="更新" then
Set conn=Server.CreateObject("ADODB.CONNECTION")
conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("../../stock/stock.mdb")
conn.execute("update 股票 set 流通股总量='"&zl&"',剩余股份='"&sy&"',发行价='"&fx&"',历史高价='"&gj&"',历史低价='"&dj&"',最近成交价='"&cjj&"',现价='"&xj&"' where 股票名称='"&gupiao&"'")
conn.close
set conn=nothing
Response.write "数据库操作完成！你已经成功更新了股票"
Response.end 
end if
%>