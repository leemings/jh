<!--#include file="conn.asp"-->
<!--#include file="../const.asp"-->
<!--#include file="checkadmin.inc"-->
<%
page=trim(request.querystring("page"))
AskSClassid=trim(request.querystring("AskSClassid"))
LF_Path=request.form("LF_path")
MusicName=request.form("MusicName")
MusicType=request.form("MusicType")
DJUser=request.form("DJUser")
SClassid=request.form("SClassid")
act=request("act")

if act<>"del" and act<>"istop" then
	if SClassid="" then
		response.write("����ʧ����ѡ��<font color=red size=8>����</font>!<a href=### onclick=window.history.go(-1);>�����ﷵ��</a>")
		response.end
	end if
		if MusicName="" then
		response.write("����ʧ��������<font color=red size=8>��������</font>!<a href=### onclick=window.history.go(-1);>�����ﷵ��</a>")
		response.end
	end if
		if LF_Path="" then
		response.write("����ʧ��������<font color=red size=8>�������ӵ�ַ</font>!<a href=### onclick=window.history.go(-1);>�����ﷵ��</a>")
		response.end
	end if
end if

set rs=server.createobject("adodb.recordset")
	select case act
	case "add"
		call add()
	case "edit"
		call edit()
	case "del"
		call del()
	case "istop"
		call istop()
	case else
	rs.close
	conn.close
	set conn=nothing
	response.write "��������ԭ��δ֪!"
	Response.End
end select


sub add()
'ִ�����������ʼ
		sql="select * from MusicDJ where (id is null)" 
		rs.open sql,conn,1,3
		rs.addnew
		rs("LF_Path")=LF_Path
		rs("MusicName")=MusicName
		rs("MusicType")=MusicType
		rs("DJUser")=DJUser
		rs("Sclassid")=Sclassid
		rs("dateandtime")=now()
		rs.update
		rs.close
		call Success
		Response.End 
'��ӽ���
end sub


sub edit()
'�༭������ʼ
		sql="select * from MusicDJ where id="&request("id")
		rs.open sql,conn,1,3
		rs("LF_Path")=LF_Path
		rs("MusicName")=MusicName
		rs("MusicType")=MusicType
		rs("DJUser")=DJUser
		rs("Sclassid")=Sclassid
		rs("dateandtime")=now()
		rs.update
		rs.close
		call Success
		Response.End 
		
'�༭����
end sub


sub istop()
'�����ö�
		sql="select istop from MusicDJ where id="&request("id") 
		rs.open sql,conn,1,3
		if not rs.EOF then
			if rs("istop")=0 then
				rs("istop")=1
			else
				rs("istop")=0
			end if
			rs.update
		end if
		rs.close
		Response.Write "����<font color=red size=8>�ö�(����)</font>�����ɹ���<a href=### onclick=window.history.go(-1);>�����ﷵ��!</a>"
		Response.End 
'�ö�����
end sub


sub del()
'ɾ��������ʼ
	sql="delete from MusicDJ where id="&request.QueryString("ID")
	rs.open sql,conn,1,1
	conn.close
	set conn=nothing
	Sclassid=request("Sclassid")
	page=request("page")
	response.redirect "songmana.asp?Sclassid="&Sclassid&"&page="&page&""

'ɾ������
end sub




sub Success
%>
<body>

      <div align="center">
        <center>
          <table width="400" border="1" cellspacing="3" cellpadding="3" class="Tableline" style="border-collapse: collapse" bordercolor="#111111">
            <tr>
              <td width="100%" align="center" bgcolor="#FF9933">����<%if act="add" then%>���<%else%>�޸�<%end if%>�ɹ�</td>
            </tr>
            <tr>
              <td width="100%">�� �� ����<%=MusicName%></td>
            </tr>
            <tr>
              <td>�� �� �ˣ�DJ ���</td>
            </tr>
            
            <tr>
              <td>���ŷ�ʽ��<%=MusicType%></td>
            </tr>
            
            <tr>
              <td>������ַ��<a href=<%=djserver%><%=LF_Path%> title=��һ�������Ƿ���ȷ><%=djserver%><%=LF_Path%></a></td>
            </tr>
            
            <tr>
              <td height="15" align=center bgcolor="#FF9933">
              <%
              	Sclassid=request("Sclassid")
				page=request("page")
              %>
              <input type="button" name="button1" value="��  ��" onclick=window.open("songmana.asp?Sclassid=<%=Sclassid%><%if act<>"add" then%>&page=<%=page%><%end if%>","_self")>&nbsp;<input type="button" name="button2" value="����<%if act="add" then%>���<%else%>�޸�<%end if%>" onclick="javascript:history.go(-1)"></td>
            </tr>
          </table>
        </center>
    </div>

</body>
</html>
<%end sub%>