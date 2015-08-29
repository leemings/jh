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
a=trim(lcase(request("a")))
wg=trim(request.form("wg"))
nei=int(abs(request.form("nl")))
if instr(wg,",")<>0then
	response.write "你好呀！黑客先生，这回不灵了吧？！"
	response.end
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "SELECT * FROM 用户 where 配偶<>'无' and 配偶<>姓名 and 姓名='"&sjjh_name&"'" ,conn
if rs.bof or rs.eof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	response.write "你并没有配偶，不能创建合体技！"
	response.end
end if
ff=rs("配偶")
sex=rs("性别")
rs.close
rs.open "update 用户 set 银两=银两-10000 where 姓名='"&sjjh_name&"'",conn
if a="m" then
	id=int(abs(request.form("id")))
	conn.Execute "update t set a='" & wg & "', d=" & nei & " where id="&id
end if
if a="n" then
	if sex="男" then
		conn.Execute "insert into t(a,b,c,d) values ('" & wg & "','" & sjjh_name & "','" & ff & "','" & nei & "')"
	else
		conn.Execute "insert into t(a,b,c,d) values ('" & wg & "','" & ff & "','" & sjjh_name & "','" & nei & "')"
	end if
end if
set rs=nothing
conn.close
set conn=nothing
Response.Redirect "stunt.asp"
%>
