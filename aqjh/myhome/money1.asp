<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
if isonline(session("aqjh_name"),session("inroom"))=false then
	Response.Write "<script Language=Javascript>alert('提示：使用本功能必须要进入聊天室！');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
money=clng(trim(Request("money")))
money1=clng(trim(Request("money1")))
fs=clng(trim(Request("fs")))
if fs<>1 and fs<>2 then
	Response.Write "<script Language=Javascript>alert('提示：选择数据有错，请重新选择！');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "SELECT 配偶 FROM 用户 WHERE 姓名='" & aqjh_name &"'",conn,1,1
mywife=rs("配偶")
rs.close
select case fs
case 1
	rs.open "SELECT 银两 FROM 用户 WHERE 姓名='" & aqjh_name &"'",conn,1,3
	if rs.Eof or rs.Bof then
		rs.Close
		Set rs = Nothing
		conn.Close
		Set conn = Nothing
		Response.Write "<script Language=Javascript>alert('提示：你的账号在江湖并不存在！');location.href = 'javascript:history.go(-1)';</script>"
		Response.End
	end if
	if rs("银两")<money then
		rs.Close
		Set rs = Nothing
		conn.Close
		Set conn = Nothing
		Response.Write "<script Language=Javascript>alert('提示：你哪里有这么多的银两！');location.href = 'javascript:history.go(-1)';</script>"
		Response.End	
	end if
	rs("银两")=rs("银两")-money
	rs.update
	rs.close
	rs.open "SELECT h_夫妻钱柜 FROM house WHERE "& session("myroom") &"='" & aqjh_name &"'",conn,1,3
	rs("h_夫妻钱柜")=rs("h_夫妻钱柜")+money
	rs.update
	if session("myroom")="h_拥有者" then
		show=session("aqjh_name") &"在" &now() &"向自己钱柜中存入:"& money &"两"
	else
		show=session("aqjh_name") &"在" &now() &"向配偶钱柜中存入:"& money &"两"
	end if
case 2
	rs.open "SELECT h_夫妻钱柜 FROM house WHERE "& session("myroom") &"='" & aqjh_name &"'",conn,1,3
	if rs("h_夫妻钱柜")<money then
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		Response.Write "<script Language=Javascript>alert('提示：你的钱柜中没有这么多的钱!');javascript:history.go(-1);</script>"
		response.end
	end if
	rs("h_夫妻钱柜")=rs("h_夫妻钱柜")-money
	rs.update
	if session("myroom")="h_拥有者" then
		show=session("aqjh_name") &"在" &now() &"从自己钱柜中取出:"& money &"两!"
	else
		show=session("aqjh_name") &"在" &now() &"从配偶钱柜中取出:"& money &"两!"
	end if
	rs.close
	rs.open "SELECT 银两 FROM 用户 WHERE 姓名='" & aqjh_name &"'",conn,1,3
	if rs.Eof or rs.Bof then
		rs.Close
		Set rs = Nothing
		conn.Close
		Set conn = Nothing
		Response.Write "<script Language=Javascript>alert('提示：你的账号在江湖并不存在！');location.href = 'javascript:history.go(-1)';</script>"
		Response.End
	end if
	rs("银两")=rs("银两")+money
	rs.update
end select
conn.execute "insert into m(b,a,c,e,d) values ('" & session("aqjh_name") & "',"& application("sysdate") &",'"& mywife &"','" & show & "','夫妻钱柜')"
rs.close
set rs=nothing
conn.close
set conn=nothing
Response.Write "<script Language=""Javascript"">alert('"& show &"');location.href = 'money.asp';</script>"
%>