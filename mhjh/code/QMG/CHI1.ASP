<%
id=request("id")
my=session("yx8_mhjh_username")
%><!--#include file="dadata.asp"-->
<%
set conn=server.createobject("adodb.connection")
conn.open Application("yx8_mhjh_connstr")
set rst=server.CreateObject ("adodb.recordset")
sql="select 银两 from 用户 where 姓名='" & my & "'"
set rs=conn.execute(sql)
if rs.eof or rs.bof then
response.write "你不是剑侠中人，不能吃面"
conn.close
response.end
else
xianjin=rs("银两")
sql="SELECT * FROM 卖面 where ID=" & id
Set rs=connt.Execute(sql)
yin=rs("售价")
tl=rs("体力")
yy=rs("拥有者")
if yin <= xianjin then
sql="update 用户 set 银两=银两-" & yin & ",体力=体力+" & tl & "   where  姓名='" & my & "'"
conn.execute(sql)
connt.execute("delete * from 卖面 where ID="&id)
response.write "你吃了巧面馆师傅"& yy &"做的面，增加体力"&tl&""
else
response.write "不能吃面，原因：你的银两不够了"
conn.close
response.end
end if
set rs=nothing
conn.close
connt.close%>
<%end if%>
