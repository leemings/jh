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
name=trim(request("name"))
my=sjjh_name
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "SELECT * FROM h where a='" & name & "' and b='"&my&"'",conn
if rs.eof or rs.bof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Redirect "../error.asp?id=481"
	response.end
end if
rs.close
rs.open "select * from 用户 where 配偶='无' and 配偶<>姓名 and 姓名='"&name&"'",conn
if rs.eof or rs.bof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Redirect "../error.asp?id=430"
	response.end
end if
if rs("门派")="出家" and sjjh_grade<10 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：你是出家人不可以操作！');}</script>"
	Response.End
end if
sex=rs("性别")
rs.close
rs.open "SELECT * FROM 用户 where 姓名='"&sjjh_name&"'",conn
if rs("门派")="出家" and sjjh_grade<10 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：你是出家人不可以操作！');}</script>"
	Response.End
end if
if rs("性别")=sex then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Redirect "../error.asp?id=428"
	response.end
end if			
if rs("配偶")<>"无" then
	conn.execute "delete * from h where a='" & name & "' or b='" & name & "'"
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Redirect "../error.asp?id=429"
	response.end
end if
	conn.execute "update 用户 set 配偶='" & name & "',结婚记念日=date(),结婚次数=结婚次数+1 where 姓名='"&sjjh_name&"'"
	conn.execute "update 用户 set 配偶='" & my & "',结婚记念日=date(),结婚次数=结婚次数+1 where 姓名='" & name & "'"
	conn.execute "delete * from h where a='" & name & "' or b='" & name & "'"
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Redirect "yuelao.asp"
%>