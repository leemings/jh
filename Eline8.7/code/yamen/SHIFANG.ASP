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
name=request("name")
my=sjjh_name
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "select * from 用户 where 姓名='"&sjjh_name&"'",conn
if rs.eof or rs.bof then
	response.write "你不是江湖中人，不能救人"
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	response.end
end if
if rs("银两")<50000000 then
response.write my & "你的银两不够"
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	response.end
end if
rs.close
rs.open "select * from 用户 where 姓名='" & name & "'",conn
if rs("状态")="监禁" then
	response.write name & "用户是被监禁的不能释放"
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	response.end
end if
conn.execute "update 用户 set 银两=银两-1000000 where 姓名='"&sjjh_name&"'"
conn.execute "update 用户 set 状态='正常',登录=now() where 姓名='" & name & "'"
rs.close
set rs=nothing
conn.close
set conn=nothing
Response.Redirect "laofan.asp"
%>
