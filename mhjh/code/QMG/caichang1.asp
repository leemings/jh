<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
id=request("id")
my=session("yx8_mhjh_username")
%>
<!--#include file="dadata.asp"-->
<%
set conn=server.createobject("adodb.connection")
conn.open Application("yx8_mhjh_connstr")
set rst=server.CreateObject ("adodb.recordset")
sql="select 银两 from 用户 where 姓名='" & my & "'"
set rs=conn.execute(sql)
if rs.eof or rs.bof then
response.write "你不是剑侠中人，不能买原料"
conn.close
response.end
else
sql="SELECT * FROM 菜场 where ID=" & id
Set Rs=connt.Execute(sql)
wu=rs("物品名")
yin=rs("银两")
lx=rs("类型")
sm=rs("说明")
gn=rs("功能")
mw=rs("美味度")
set conn=server.createobject("adodb.connection")
conn.open Application("yx8_mhjh_connstr")
set rst=server.CreateObject ("adodb.recordset")
sql="select * from 用户 where 姓名='" & my & "'"
rs=conn.execute(sql)
if yin<=rs("银两") then
sql="update 用户 set 银两=银两- " & yin & "  where 姓名='" & my & "'"
rs=conn.execute(sql)
set conn=server.createobject("adodb.connection")
conn.open Application("yx8_mhjh_connstr")
set rst=server.CreateObject ("adodb.recordset")
sql="select * from 物品 where 物品名='" & wu & "' and 拥有者='" & my & "'"
set rs=connt.execute(sql)
if rs.eof or rs.bof then
sql="insert into 物品(物品名,拥有者,类型,说明,功能,美味度,银两) values ('"& wu &"','"& my &"','"& lx &"','" & sm &"','"& gn &"',"& mw &","& yin &")"
rs=connt.execute(sql)
connt.close
Response.Redirect "caichang.asp"
else
response.write "由于你已购买了这个物品，所以不能再买！"
connt.close
response.end
end if
else
response.write "不能买东西，原因：你的银两不够了"
connt.close
response.end
end if
rs.close
set rs=nothing
end if
%>
