<!--#include file="function.asp"-->
<%CheckAdmin1%>
<!--#include file="conn.asp"-->
<!--#include file="char.inc"-->
<%
act=request("act")
ClassID=request.form("Classid")
founderr=false
if act="rename" then
	FunRename
elseif act="del" then
	FunDel
elseif act="edit" then
	FunEdit
elseif act="add" then
	FunAdd
else
	errmsg=errmsg+"������������ϵ����Ա"
	founderr=true
end if

if founderr=true then
	call error()
else

end if

function FunRename
	if trim(request.form("NClass"))="" then
		errmsg=errmsg+"<br>"+"<li>�������Ʋ���Ϊ�գ�"
		call error()
		Response.End 
	end if
'�޸���Ŀ�е���Ŀ��
	set rs=server.createobject("adodb.recordset")
	sql = "SELECT * FROM Nclass where Nclassid=" & request.QueryString("NClassid")
	rs.Open sql,conn, 1, 3
	if err.Number<>0 then
		err.clear
		errmsg=errmsg+"������������ϵ����Ա"
		founderr=true
	else
		rs("Nclass") = trim(request.form("Nclass"))
		rs("Abcd") = trim(request.form("Abcd"))
		rs.Update
	end if
	rs.Close
'�޸ĸ����е���Ŀ��
	sql= "SELECT * FROM MusicList where Nclassid=" & request.QueryString("NClassid")
	rs.Open sql, conn, 1, 3
	if err.Number<>0 then
		err.clear
		errmsg=errmsg+"������������ϵ����Ա"
		founderr=true
	else
		do while not rs.eof
			rs("singer") = trim(request.form("NClass"))
			rs.Update
		rs.movenext
		loop
	end if
	rs.Close
'�޸�ר���е���Ŀ��
	sql = "SELECT * FROM special where Nclassid=" & request.QueryString("NClassid")
	rs.Open sql,conn, 1, 3
	if err.Number<>0 then
		err.clear
		errmsg=errmsg+"������������ϵ����Ա"
		founderr=true
	else
		do while not rs.eof
			rs("Nclass") = trim(request.form("NClass"))
			rs.Update
		rs.movenext
		loop
	end if
	rs.Close
'�޸��ղ��е���Ŀ��
	sql = "SELECT * FROM Box where Musicid=" & request.QueryString("NClassid")
	rs.Open sql,conn, 1, 3
	if err.Number<>0 then
		err.clear
		errmsg=errmsg+"������������ϵ����Ա"
		founderr=true
	else
		do while not rs.eof
			rs("MusicName") = trim(request.form("NClass"))
			rs.Update
		rs.movenext
		loop
	end if
	rs.Close
'�����޸�
	set rs=nothing
	conn.close
	set conn=nothing
end function

function FunDel
	Rem ɾ������Ŀ
	sql="delete from Nclass where Nclassid=" & request.QueryString("NClassid")
	conn.execute sql
	Rem ɾ�������Ʒ
	sql="delete from shangpin where Nclassid=" & request.QueryString("NClassid")
	conn.execute sql
	if err.Number<>0 then
		err.clear
		errmsg=errmsg+"������������ϵ����Ա"
		founderr=true
	end if
end function
	
function FunAdd
	NClass=trim(request.form("NClass"))
	if NClass="" then
		errmsg=errmsg+"������Ʋ���Ϊ�գ�"
		call error()
		Response.End 
	end if
	set rs=server.createobject("adodb.recordset")
	sql = "SELECT * FROM Nclass where (Nclassid IS NULL)"
	rs.Open sql,conn, 1, 3
	if err.Number<>0 then
		err.clear
		errmsg=errmsg+"������������ϵ����Ա"
		founderr=true
	else
		rs.AddNew
		rs("Nclass") = NClass
		rs("Classid") = request.QueryString("Classid")
		rs("SClassid") = request.QueryString("SClassid")
        if trim(request("Abcd"))<>"" then rs("Abcd") = request.form("Abcd")
		rs.Update
	end if
	rs.Close
	set rs=nothing
	conn.close
	set conn=nothing
end function

function FunEdit
	if trim(request("NClass"))="" then
		errmsg=errmsg+"������Ʋ���Ϊ�գ�"
		call error()
		Response.End 
	end if
	set rs=server.createobject("adodb.recordset")
	sql = "SELECT * FROM Nclass where NClassid="&request.QueryString("NClassid")
	rs.Open sql,conn, 1, 3
	if err.Number<>0 then
		err.clear
		errmsg=errmsg+"������������ϵ����Ա"
		founderr=true
	else

'�޸���Ŀ�е�����
		rs("Nclass") = trim(request.form("NClass"))
		rs("Classid") = request.form("Classid")
		rs("SClassid") = request.form("SClassid")
		rs.Update
	end if
	rs.Close
	set rs=nothing
	conn.close
	set conn=nothing
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


