<%@ LANGUAGE="VBSCRIPT" %>
<!--#include file="session.asp"-->
<!--#include file="conn.asp"-->
<!--#include file="level.asp"-->
<%
dim ID
ID=request("ID")
set rs=server.createobject("adodb.recordset")
sql="select * from admin where ID="&request("ID")
rs.open sql,conn,1,1
%>
<html>
<head>
<title>�޸��û�</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="../css/style.css" type="text/css">
<Script Language=Javascript>
function check(){
var f=document.add
if(f.username.value.length==0){alert("�û�������Ϊ��");f.username.focus();return false}
if(f.password.value.length==0){alert("�û����벻��Ϊ��");f.password.focus();return false}
}
function over(locate){
locate.style.border='1px solid black';locate.style.backgroundColor='#FFFFFF';locate.style.cursor='hand'
}
function out(locate){
locate.style.border='';locate.style.backgroundColor='#F7F7F7'
}
</script>
</head>

<body bgcolor="#FFFFFF" text="#000000" background="bkg.gif">
<form name="add" method="post" action="save_user.asp" onsubmit="return check()">
        
  <table width="40%" border="0" cellspacing="0" cellpadding="0" align="center">
    <tr> 
            <td height="27" width="120" align="center">�û�������</td>
            
      <td height="27" width="209"> 
        <input type="text" name="username" size="20" maxlength="20" class="form" value="<%=rs("username")%>">
            </td>
          </tr>
          <tr> 
            <td height="27" width="120" align="center">�û����룺</td>
            
      <td height="27" width="209"> 
        <input type="password" name="password" size="20" maxlength="20" class="form" value="<%=rs("password")%>">
            </td>
          </tr>
          <tr> 
            <td height="27" width="120" align="center">�û��ȼ���</td>
            
      <td height="27" width="209"> 
        <select name="level" class="form">
                <option value="">��ѡ���û��ȼ�</option>
                <option value="ȫ��" <% if rs("Level")="ȫ��" then%>selected<%end if%>>ȫ��</option>
                <option value="����" <% if rs("Level")="����" then%>selected<%end if%>>����</option>
                <option value="Ӱ��" <% if rs("Level")="Ӱ��" then%>selected<%end if%>>Ӱ��</option>
                <option value="����" <% if rs("Level")="����" then%>selected<%end if%>>����</option>
		<option value="����" <% if rs("Level")="����" then%>selected<%end if%>>����</option>
		<option value="��Ϸ" <% if rs("Level")="��Ϸ" then%>selected<%end if%>>��Ϸ</option>
		<option value="����" <% if rs("Level")="����" then%>selected<%end if%>>����</option>
		<option value="Ů��" <% if rs("Level")="Ů��" then%>selected<%end if%>>Ů��</option>
		<option value="����" <% if rs("Level")="����" then%>selected<%end if%>>����</option>
		<option value="����" <% if rs("Level")="����" then%>selected<%end if%>>����</option>
		<option value="���" <% if rs("Level")="���" then%>selected<%end if%>>���</option>
		<option value="�˲�" <% if rs("Level")="�˲�" then%>selected<%end if%>>�˲�</option>
		<option value="����" <% if rs("Level")="����" then%>selected<%end if%>>����</option>
		<option value="��ֽ" <% if rs("Level")="��ֽ" then%>selected<%end if%>>��ֽ</option>
              </select>
            </td>
          </tr>
          <tr> 
            <td height="37" colspan="2" align="center">
		<input type="hidden" name="ID" value="<%=rs("id")%>"> 
              <input type="hidden" name="action" value="edit">
              <input type="submit" name="Submit" value="�ύ�޸�">
              �� 
              <input type="reset" name="reset" value="�����">
            </td>
          </tr>
        </table>
      </form>
<%
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
</body>
</html>