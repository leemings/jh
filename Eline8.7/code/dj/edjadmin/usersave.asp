<!--#include file="conn.asp"-->
<!--#include file="checkadmin.inc"-->

<%
id=request("id")
act=request("act")

set rs=server.CreateObject("ADODB.RecordSet")
if act="edit" then
	username=trim(request("username"))
	password=trim(request("password"))
	Sex=trim(request("Sex"))
	Email=trim(request("Email"))
	OICQ=trim(request("oicq"))
	homepage=trim(request("homepage"))
	Address=trim(request("address"))
	
	if username="" or password="" or Sex="" then
		Response.Write "用户名称，密码和性别都不能为空！"
		Response.End 
	end if
	sql="select * from user where id="& request("id")
	rs.open sql,conn,1,3
	if not rs.eof then
		rs("Username")=username
		rs("Password")=password
		rs("quesion")=request("quesion")
		rs("answer")=request("answer")
		rs("Sex")=Sex
		rs("email")=email
       	rs("sex")=request("sex")
       	rs("Address")=request("Address")
       	rs("homepage")=request("homepage")
       	rs("Tel")=request("Tel")
       	rs("oicq")=request("oicq")
		rs.update
	end if
	rs.close

elseif act="del" then
set rs=conn.execute("delete * from user where id="&request("id")) 
set rs=nothing
conn.close
set conn=nothing
response.redirect "UserMana.asp"

elseif act="deldate" then
	selectdate=trim(Request.Form ("selectdate"))
	selecttimes=trim(Request.Form ("selecttimes"))
	if selectdate<>"" or selecttimes<>"" then
		thisdate=date()
		set rs=server.CreateObject("ADODB.RecordSet")
		if selectdate="" and selecttimes<>"" then
			sql="select username from [user] where lockuser<>true and logins<"&selecttimes
		elseif selectdate<>"" and selecttimes="" then
			sql="select username from [user] where lockuser<>true and datediff('y',datevalue(lastlogin),#"&thisdate&"#)>"&selectdate
		elseif selectdate<>"" and selecttimes<>"" then
			sql="select username from [user] where lockuser<>true and datediff('y',datevalue(lastlogin),#"&thisdate&"#)>"&selectdate&" and logins<"&selecttimes
		end if
		rs.open sql,conn,1,1
		if not rs.eof then
			arrRow=rs.GetRows
			arrCon=rs.RecordCount
			rs.close
			set rs=nothing
			for i=0 to arrCon-1
				conn.execute("delete * from [user] where lockuser<>true and username='"&arrRow(0,i)&"'")
			next
			arrCon=null
			arrRow=null

		else
			rs.close
			set rs=nothing
		end if
	end if
end if
conn.close
set conn=nothing
response.redirect "UserMana.asp"
%>