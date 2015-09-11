<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../../error.asp?id=440"
name=trim(Request.form("name"))
mess=trim(Request.form("mess"))
money=Request.form("money")
if not isnumeric(money) then 
	Response.Write "<script language=JavaScript>{alert('提示：["&money&"]输入错误，ID请使用数字！');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End 
end if
money=int(money)
if name=aqjh_name or name="" or mess="" or money<=100000 then Response.Redirect "../../error.asp?id=433"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "SELECT * FROM h WHERE a='"&aqjh_name&"'",conn
If not(Rs.Bof OR Rs.Eof) Then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Redirect "../../error.asp?id=431"
	response.end
end if
rs.close
'校验用户 魅力大于100，钱大于1000
rs.open "SELECT * FROM 用户 WHERE 姓名='"&aqjh_name&"'"&" and 配偶='无' and 等级>5 and 银两>"&(money+100000),conn
If Rs.Bof OR Rs.Eof Then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Redirect "../../error.asp?id=433"
	response.end
end if
if rs("门派")="出家" and aqjh_grade<10 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：你是出家人不可以操作！');}</script>"
	Response.End
end if
sex=trim(rs("性别"))
rs.close
rs.open "SELECT * FROM 用户 WHERE 等级>18 and 姓名='" & name & "' and 配偶='无' and 性别<>'"&sex&"'",conn
if rs.bof or rs.eof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Redirect "../../error.asp?id=430" 
end if
	conn.execute "update 用户 set 银两=银两-"&(money+100000)&" where 姓名='"&aqjh_name&"'"
	conn.Execute "INSERT INTO h (a,b,c,d,e,f) VALUES('"&aqjh_name&"','"&name&"','结婚',"&money&",now(),'"&mess&"')"
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Redirect "../../ok.asp?id=600"
rs.close
set rs=nothing
conn.close
set conn=nothing
%>