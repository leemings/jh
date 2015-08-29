<%PageName="admin_AdminModify"%>
<!--#include file="function.asp"-->
<%CheckAdmin3%>
<!--#include file="conn.asp"-->
<%
founderr=false

id=request.QueryString("id")
set rs=server.createobject("adodb.recordset")
sql="select * from admin where id="&id
rs.open sql,conn,1,1
if rs.eof then
	errmsg=errmsg+"<br>"+"<li>服务器出错！请联系管理员"
	founderr=true
else
	Username=rs("Username")
	Password=rs("Password")
	oskey=rs("oskey")
rs.close
end if

if founderr=true then
	call error()
else
%>
<!--#include file="top.asp"-->
<div align="center">
<center>
<table border="0" width="80%" cellspacing="0" cellpadding="5">
  <tr>
    <td align=center valign=top>
      <table border="1" width="100%" cellspacing="0" cellpadding="0" class="TableLine" bordercolor="#FF1171" bordercolordark="#FFFFFF">
        <form method="POST" action="AdminSave.asp?id=<%=id%>" id=form2 name=form2>
          <tr>
            <td width="100%" height="20" colspan=2 bgcolor="#FF1171" align=center><font color="white"><b>修 改 管 理 员 资 料</b></td>
          </tr>
          <tr>
            <td width="30%" align="right" height="30">用户名：</td>
            <td width="70%">
            <input type="text" name="username" value="<%=Username%>" size="20"></td>
          </tr>
          <tr>
            <td width="30%" align="right" valign="top" height="20">密码：</td>
            <td width="70%">
            <input type="password" name="password" value="<%=Password%>" size="20"></td>
          </tr>
          <tr>
            <td width="30%" align="right" height="30">权限：</td>
            <td width="70%" height="30">
              <select name="oskey" style="font-size:9pt">
                <option value=super<%if oskey="super" then%> selected<%end if%>>高级管理员</option>
                <option value=check<%if oskey="check" then%> selected<%end if%>>中级管理员</option>
                <option value=input<%if oskey="input" then%> selected<%end if%>>初级管理员</option>
              </select>
            </td>
          </tr>
          <tr align="center">
            <td colspan=2>
              <input type="hidden" value="edit" name="act"> 
              <input type="submit" value=" 修 改 " name="cmdok2">&nbsp; 
              <input type="reset" value=" 清 除 " name="cmdcance2l">
            </td>
          </tr>
        </form>
      </table>
    </td>
  </tr>
</table>
</body>
</html>
<%end if%>




