<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"

aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
name=request("name")
my=aqjh_name
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select * from 用户 where 姓名='"&aqjh_name&"'",conn
if rs.eof or rs.bof then
	response.write "你不是江湖中人，不能救人"
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	response.end
end if
if rs("银两")<80000000 then
response.write my & "你的银两不够怎么救人？"
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	response.end
end if
if rs("法力")<5000 then
response.write my & "你的法力不够怎么救人？"
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
conn.execute "update 用户 set 银两=银两-80000000 where 姓名='"&aqjh_name&"'"
conn.execute "update 用户 set 法力=法力-5000 where 姓名='"&aqjh_name&"'"
conn.execute "update 用户 set 状态='正常',登录=now() where 姓名='" & name & "'"
rs.close
set rs=nothing
conn.close
set conn=nothing
Response.Redirect "laofan1.asp"
%>
