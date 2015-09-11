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
money=abs(request.form("money"))
mess=trim(Request.form("mess"))
if name="" or mess="" or money<10000 or name=aqjh_name then
	Response.Write "<script Language=Javascript>alert('提示：姓名为空，信息为空或分手费太少!');location.href = 'lihun.asp';</script>"	
	response.end
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "SELECT * FROM h WHERE a='" & aqjh_name & "' or b='" & name & "'",conn
If not(Rs.Bof OR Rs.Eof) Then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：["&aqjh_name&"你在江湖上申请过求婚或离婚了！！');location.href = 'lihun.asp';</script>"	
	response.end
end if
rs.close
rs.open "SELECT * FROM 用户 WHERE 姓名='"&aqjh_name&"'" ,conn
If DateDiff("d",rs("结婚记念日"),date())<30 Then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：结婚还不到一个月不能离婚！');location.href = 'lihun.asp';</script>"	
	response.end
end if
peiou=trim(rs("配偶"))
if peiou<>name then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示："& name &"也不是你的配偶！!');location.href = 'lihun.asp';</script>"	
	response.end
end if
if rs("银两")<money then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：你自己有那么多的钱吗？！!');location.href = 'lihun.asp';</script>"
	response.end
end if
rs.close
rs.open "SELECT * FROM 用户 WHERE trim(姓名)='" & name & "'",conn
if rs.bof or rs.eof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：["&name&"在江湖上找不到，请找管理员解决！！');location.href = 'lihun.asp';</script>"	
	response.end
end if
rs.close
conn.execute "update 用户 set 银两=银两-"&money&" where 姓名='"&aqjh_name&"'"
conn.Execute "INSERT INTO h (a,b,c,d,e,f) VALUES('"&aqjh_name&"','"&name&"','离婚',"&money&",now(),'"&mess&"')"
Response.Write "<script Language=Javascript>alert('提示："& aqjh_name &"你的离婚登记成功！！');location.href = 'lihun.asp';</script>"	
response.end
%>