<!--#include file="conn.asp"-->
<!--#include file="checkadmin.inc"-->
<%
SClassid=request.querystring("SClassid")
Classid=request.form("Classid")
Sclass=request.form("Sclass")
act=request("act")

if act<>"del" and act<>"add" then
	if SClassid="" then
		response.write("����ʧ������<font color=red size=8>����������</font>!<a href=### onclick=window.history.go(-1);>�����ﷵ��</a>")
		response.end
	end if

		if ClassID="" then
		response.write("����ʧ����ѡ��˶�������������<font color=red size=8>һ������</font>!<a href=### onclick=window.history.go(-1);>�����ﷵ��</a>")
		response.end
	end if
end if
if act="add" or act="edit" then
set rs=server.createobject("adodb.recordset")
sql="select * from Class where Classid="&Classid
rs.open sql,conn,1,1
if not rs.eof then
ClassName=rs("class")
else
		response.write("����ʧ��û��IDΪ  "&Classid&"  ��<font color=red size=8>һ������</font>!<a href=### onclick=window.history.go(-1);>�����ﷵ��</a>")
		response.end
end if
rs.close
set rs=nothing
end if

set rs=server.createobject("adodb.recordset")
	select case act
	case "add"
		call add()
	case "edit"
		call edit()
	case "del"
		call del()
	case else
	rs.close
	conn.close
	set conn=nothing
	response.write "��������ԭ��δ֪!"
	Response.End
end select


sub add()
'ִ����Ӷ������࿪ʼ
		sql="select * from SClass where (SClassid is null)" 
		rs.open sql,conn,1,3
		rs.addnew
		rs("Sclass")=Sclass
		rs("Classid")=Classid
		rs.update
		rs.close
		call Success
		Response.End 
'��ӽ���
end sub


sub edit()
'�༭�������࿪ʼ
		sql="select * from SClass where SClassid="&SClassid
		rs.open sql,conn,1,3
		rs("Sclass")=Sclass
		rs("Classid")=Classid
		rs.update
		rs.close
		call Success
		Response.End 
		
'�༭����
end sub


sub del()
'ɾ���������࿪ʼ
	sql="delete from SClass where SClassid="&request.QueryString("SClassID")
	rs.open sql,conn,1,1
	conn.close
	set conn=nothing
	response.redirect "SClassmana.asp"

'ɾ������
end sub




sub Success
%>
<body>

      <div align="center">
        <center>
          <table width="400" border="1" cellspacing="3" cellpadding="3" class="Tableline" style="border-collapse: collapse" bordercolor="#111111">
            <tr>
              <td width="100%" align="center" bgcolor="#FF9933">��������<%if act="add" then%>���<%else%>�޸�<%end if%>�ɹ�</td>
            </tr>
            <tr>
              <td width="100%">������������<%=SClass%></td>
            </tr>
            
            <tr>
              <td>����һ�����ࣺ<%=ClassName%> </td>
            </tr>
            
            <tr>
              <td height="15" align=center bgcolor="#FF9933">
              <%
              	Sclassid=request("Sclassid")
				page=request("page")
              %>
              <input type="button" name="button1" value="��  ��" onclick=window.open("SClassmana.asp","_self")>&nbsp;<input type="button" name="button2" value="����<%if act="add" then%>���<%else%>�޸�<%end if%>" onclick="javascript:history.go(-1)"></td>
            </tr>
          </table>
        </center>
    </div>

</body>
</html>
<%end sub%>