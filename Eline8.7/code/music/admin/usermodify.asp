<%PageName="admin_UserModify"%>
<!--#include file="function.asp"-->
<%CheckAdmin3%>
<!--#include file="conn.asp"-->
<%
id=request.QueryString("id")
set rs=server.createobject("adodb.recordset")
sql="select * from [user] where id="&id
rs.open sql,conn,1,1
if rs.eof then
	errmsg="<br>������������ϵ����Ա"
	call error()
	Response.end
else
	Username=rs("Username")
	Password=rs("Password")
	Sex=rs("Sex")
	Email=rs("Email")
	Tel=rs("Tel")
	Name=rs("Name")
	Address=rs("Address")
	Youbian=rs("Youbian")
	Shenfenzheng=rs("Shenfenzheng")
rs.close
end if
%>
<!--#include file="top.asp"-->
<div align="center">
<table border="0" width="88%" cellspacing="1" cellpadding="1">
  <tr>
    <td align=center valign=top>
      <table border="1" width="100%" cellspacing="0" cellpadding="0" class="TableLine" bordercolor="#FF1171" bordercolordark="#FFFFFF">
        <form method="POST" action="UserSave.asp?id=<%=id%>" id=form2 name=form2>
          <tr>
            <td width="100%" height="20" colspan=2 bgcolor="#FF1171" align=center><font color="white"><b>�� �� �� �� �� ��</b></font></td>
          </tr>
          <tr>
            <td width="30%" height="20" align="right">�û�ID��</td>
            <td width="70%"><input type="text" name="username" value="<%=Username%>" size="20"></td>
          </tr>
          <tr>
            <td width="30%" height="20" align="right" valign="top">���룺</td>
            <td width="70%"><input type="password" name="password" value="<%=Password%>" size="20"></td>
          </tr>
          <tr>
            <td width="30%" height="20" align="right">�ƺ���</td>
            <td width="70%"><input type="text" name="Sex" value="<%=Sex%>" size="20"></td>
            </td>
          </tr>
          <tr>
            <td width="30%" height="20" align="right">E-mail��</td>
            <td width="70%"><input type="text" name="Email" value="<%=Email%>" size="20"></td>
          </tr>
          <tr>
            <td width="30%" height="20" align="right">��ʵ������</td>
            <td width="70%"><input type="text" name="Name" value="<%=Name%>" size="20"></td>
          </tr>

          <tr>
            <td width="30%" height="20" align="right">OICQ��</td>
            <td width="70%"><input type="text" name="Shenfenzheng" value="<%=Shenfenzheng%>" size="20"></td>
          </tr>
          <tr>
            <td width="30%" height="20" align="right">��ϵ�绰��</td>
            <td width="70%"><input type="text" name="Tel" value="<%=Tel%>" size="20"></td>
          </tr>
          <tr>
            <td width="30%" height="20" align="right">��ϸ��ַ��</td>
            <td width="70%"><input type="text" name="Address" value="<%=Address%>" size="20"></td>
          </tr>
          <tr>
            <td width="30%" height="20" align="right">�������룺</td>
            <td width="70%"><input type="text" name="Youbian" value="<%=Youbian%>" size="20"></td>
          </tr>
          <tr align="center">
            <td colspan=2>
              <input type="hidden" value="edit" name="act"> 
              <input type="submit" value=" �� �� " name="cmdok2">&nbsp; 
              <input type="reset" value=" �� �� " name="cmdcance2l">
            </td>
          </tr>
        </form>
      </table>
    </td>
  </tr>
</table>
</div>
<%
set rs=nothing
conn.close
set conn=nothing
%></body></html>








