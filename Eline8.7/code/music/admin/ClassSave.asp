<!--#include file="function.asp"-->
<%CheckAdmin1%>
<!--#include file="conn.asp"-->
<%
founderr=false
act=request("act")
ClassID=request.QueryString("ClassID")
if act="rename" then
	FunRename
elseif act="del" then
	FunDel
elseif act="add" then
	FunAdd
else
	errmsg=errmsg+"<br>"+"<li>������������ϵ����Ա"
	founderr=true
end if
if founderr=true then
	call error()
else

end if

function FunRename
if trim(request.form("Class"))="" then
	errmsg=errmsg+"<br>"+"<li>����������Ϊ�գ�"
	founderr=true
else
	set rs=server.createobject("adodb.recordset")
	sql = "SELECT * FROM class where classid=" & ClassID
	rs.Open sql,conn,1,3
	if err.Number<>0 then
		err.clear
		errmsg=errmsg+"<br>"+"<li>������������ϵ����Ա"
		founderr=true
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
'ɾ�������Ŀ�¶�����Ŀ
	sql="delete from Sclass where classid=" & ClassID
	conn.execute sql
'ɾ�������Ŀ��������Ŀ
	sql="delete from Nclass where classid=" & ClassID
	conn.execute sql
'ɾ�������Ŀ��Ӧ��Ʒ
	sql="delete from Shangpin where classid=" & ClassID
	conn.execute sql
	if err.Number<>0 then
		err.clear
		errmsg=errmsg+"<br>"+"<li>������������ϵ����Ա"
		founderr=true
	end if
end function

function FunAdd
if trim(request("Class"))="" then
	errmsg=errmsg+"<br>"+"<li>����������Ϊ�գ�"
	founderr=true
else
	set rs=server.createobject("adodb.recordset")
	sql = "SELECT * FROM class where (classid IS NULL)"
	rs.Open sql,conn,1,3
	if err.Number<>0 then
		err.clear
		errmsg=errmsg+"<br>"+"<li>������������ϵ����Ա"
		founderr=true
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
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="RSHOP" content="k666Ӱ������http://www.vv66.com">
<meta Author="Recall Star" content="k666Ӱ������http://www.vv66.com">
<title>k666������</title>
<!--#include file="style.asp"-->
</head>
<body topmargin="111" leftmargin="0">
<div align="center">
  <center>
  <table border="0" cellspacing="0" width="60%">
    <tr>
      <td width="100%" bgcolor="#CC0066">
        <div align="center">
          <table border="0" cellpadding="0" cellspacing="0" width="100%">
            <tr>
              <td width="100%" bgcolor="#FFFFFF" height="80" align="center">
                <b>O K !&nbsp; �� �� �� �� !&nbsp; ^_^</b>
                <p><b><a href="javascript:history.go(-1)">...::: �� �� �� �� 
                :::...</a></b>
              </td>
            </tr>
          </table>
        </div>
      </td>
    </tr>
  </table>
  </center>
</div>
</body>                    
</html>           


