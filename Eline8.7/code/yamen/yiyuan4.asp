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
if rs("银两")<10000000 then
response.write my & "你的银两也学去去医院看病？"
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	response.end
end if
rs.close
rs.open "select * from 用户 where 姓名='" & name & "'",conn
if rs("状态")="监禁" then
	response.write name & "你是重犯,想逃跑？"
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	response.end
end if
if rs("状态")="狱" then
	response.write name & "这是医院，你来错地方了吧？"
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	response.end
end if
if rs("状态")="牢" then
	response.write name & "这是医院，你来错地方了吧？"
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	response.end
end if
if rs("状态")="眠" then
	response.write name & "这是医院，你来错地方了吧？"
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	response.end
end if
conn.execute "update 用户 set 银两=银两-10000000 where 姓名='"&sjjh_name&"'"
conn.execute "update 用户 set 状态='正常',登录=now() where 姓名='" & name & "'"
rs.close
set rs=nothing
conn.close
set conn=nothing
Response.Redirect "yiyuan.asp"
%>
