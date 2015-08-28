<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
id=request("id")
if InStr(id,"=")<>0 or InStr(id,"`")<>0 or InStr(id,"'")<>0 or InStr(id," ")<>0 or InStr(id,"　")<>0 or InStr(id,"'")<>0 or InStr(id,chr(34))<>0 or InStr(id,"\")<>0 or InStr(id,",")<>0 or InStr(id,"<")<>0 or InStr(id,">")<>0 then Response.Redirect "../error.asp?id=120"
my=session("yx8_mhjh_username")
set conn=server.createobject("adodb.connection")
conn.open Application("yx8_mhjh_connstr")
set rs=server.CreateObject ("adodb.recordset")
sql="SELECT * FROM 礼品 where ID=" & id
Set Rs=conn.Execute(sql)
wu=rs("名称")
yin=rs("价格")
lx=rs("属性")
nl=rs("美貌")
tl=rs("个性")
sql="select * from 用户 where 姓名='" & my & "'"
rs=conn.execute(sql)
if yin<=rs("银两") then
sql="update 用户 set 银两=银两-" & yin & " where 姓名='" & my & "'"
rs=conn.execute(sql)
sql="select * from 物品 where 名称='" & wu & "' and 所有者='" & my & "'"
set rs=conn.execute(sql)
if rs.eof or rs.bof then
sql="insert into 物品(名称,所有者,属性,美貌,个性,价格) values ('"&wu&"','"&my&"','"&lx&"',"&nl&","&tl&","&yin&")"
rs=conn.execute(sql)
sql="update 物品 set 数量=1 where 名称='" & wu & "' and 所有者='" & my & "'"
rs=conn.execute(sql)
conn.close
Response.Redirect "ok.asp"
else
if rs("数量")<10 then
sql="update 物品 set 数量=数量+1  where 名称='" & wu & "' and 所有者='" & my & "'"
rs=conn.execute(sql)
conn.close
Response.Redirect "ok.asp"
else
response.write "购买不成功，原因：你买的物品太多了，没有地方放了"
conn.close
response.end
end if
end if
else
response.write "购买不成功，原因：你的银两不够了"
conn.close
response.end
end if
rs.close
set rs=nothing
%>
