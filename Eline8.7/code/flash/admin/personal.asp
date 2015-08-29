<%@ LANGUAGE="VBSCRIPT" %>
<!--#include file="session.asp"-->
<!--#include file="conn.asp"-->
<%
set rs=server.createobject("adodb.recordset")
sql="select * from admin where username='"&session("username")&"'"
rs.open sql,conn,1,1
%>
<html>
<head>
<title>修改用户</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="../css/style.css" type="text/css">
<Script Language=Javascript>
function check(){
var f=document.add
if(f.password.value.length==0){alert("用户密码不能为空");f.password.focus();return false}
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
<form name="add" method="post" action="save_personal.asp" onsubmit="return check()">
        
  <table width="40%" border="0" cellspacing="0" cellpadding="0" align="center">
    <tr> 
            <td height="27" width="120" align="center">用户姓名：</td>
            
      <td height="27" width="209"> <b><%=rs("username")%></b>
            </td>
          </tr>
          <tr> 
            <td height="27" width="120" align="center">用户密码：</td>
            
      <td height="27" width="209"> 
        <input type="password" name="password" size="20" maxlength="20" class="form" value="<%=rs("password")%>">
            </td>
          </tr>
          <tr> 
            <td height="37" colspan="2" align="center">
              <input type="hidden" name="action" value="edit">
              <input type="submit" name="Submit" value="提交修改">
              <input type="reset" name="reset" value="重设表单">
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