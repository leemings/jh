<!--#include file="conn.asp"-->
<!--#include file="checkadmin.inc"-->
<%

act=request("act")
ClassID=request.QueryString("ClassID")
if act="rename" then
	FunRename
elseif act="del" then
	FunDel
elseif act="add" then
	FunAdd
else
		response.write("����ʧ��!<a href=### onclick=window.history.go(-1);>�����ﷵ��</a>")
		response.end
end if
Response.Redirect "ClassMana.asp"
function FunRename
if trim(request.form("Class"))="" then
		response.write("����ʧ��<font color=red size=8>������</font>����Ϊ��!<a href=### onclick=window.history.go(-1);>�����ﷵ��</a>")
		response.end
else
	set rs=server.createobject("adodb.recordset")
	sql = "SELECT * FROM class where classid=" & ClassID
	rs.Open sql,conn,1,3
	if err.Number<>0 then
		err.clear
		response.write("����ʧ��<font color=red size=8>δ֪!</font>!<a href=### onclick=window.history.go(-1);>�����ﷵ��</a>")
		response.end
	else
		rs("class") = request.form("Class")
		rs.Update
	end if
	rs.Close
	set rs=nothing
	conn.close
	set conn=nothing
end if
end function

function FunDel
	sql="delete from class where classid=" & ClassID
	conn.execute sql
end function

function FunAdd
if trim(request("Class"))="" then
		response.write("����ʧ��<font color=red size=8>������</font>����Ϊ��!<a href=### onclick=window.history.go(-1);>�����ﷵ��</a>")
		response.end
else
	set rs=server.createobject("adodb.recordset")
	sql = "SELECT * FROM class where (classid IS NULL)"
	rs.Open sql,conn,1,3
	if err.Number<>0 then
		err.clear
		response.write("����ʧ��<font color=red size=8>δ֪</font>!<a href=### onclick=window.history.go(-1);>�����ﷵ��</a>")
		response.end
	else
		rs.AddNew
		rs("class") = request.form("Class")
		rs.Update
	end if
	rs.Close
	set rs=nothing
	conn.close
	set conn=nothing
end if
end function
%>