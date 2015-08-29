<!--#include file="admin_top.asp"-->
<!--#include file="checkadmin.inc"-->
<!--#include file="../function.asp"-->
<%
founderr=false
id=request.QueryString("id")
set rs=server.createobject("adodb.recordset")
sql="select * from admin where id="&id
rs.open sql,conn,1,1
if rs.eof then
	errmsg=errmsg+"<br>"+"<li>操作错误，没有这个用户！"
	founderr=true
else
	Username=rs("Username")
	Password=rs("Password")
	oskey=rs("oskey")
rs.close
conn.close
set rs=nothing
set conn=nothing
end if

if founderr=true then
	call error()
else
%>
<meta http-equiv="Content-Language" content="zh-cn">
<div align="center">
  <center>
  <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="760" id="AutoNumber2">
    <tr>
      <td width="100%"><img border="0" src="补间.gif" width="1" height="3"></td>
    </tr>
  </table>
  </center>
</div>
<div align="center">
  <center>
<table border="0" width="760" cellspacing="0" cellpadding="0" style="border-collapse: collapse">
  <tr>
    <td valign=top width=150>
    <!--#include file="admin_left.asp"-->
    　</td>
    <td valign=top width=10>　</td>
    <td valign=top width="600">
    <table border="0" cellpadding="3" cellspacing="3" style="border-collapse: collapse" width="100%" id="AutoNumber3" bgcolor="#FF9933">
      <tr>
        <td width="100%" bgcolor="#CC6600" align="center">
        管理员修改</td>
      </tr>
      <tr>
        <td width="100%">
      　<div align="center">
        <center>
      <table border="1" width="50%" cellspacing="0" cellpadding="0" class="TableLine" bordercolor="#CC6600" style="border-collapse: collapse">
        <form method="POST" action="AdminSave.asp?id=<%=id%>" id=form2 name=form2>
          <tr>
            <td width="100%" height="20" colspan=2 align=center bgcolor="#CC6600"><b>修 改 管 理 员 资 料</b></td>
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
          <tr align="center">
            <td colspan=2 bgcolor="#CC6600">
              <input type="hidden" value="edit" name="act"> 
              <input type="submit" value=" 修 改 " name="cmdok2">&nbsp; 
              <input type="reset" value=" 清 除 " name="cmdcance2l">
            </td>
          </tr>
        </form>
      </table>

        </center>
      </div>
<br>
</td>
      </tr>
      <tr>
        <td width="100%" bgcolor="#CC6600">　</td>
      </tr>
    </table>
    </td>
  </tr>
  </table>
  </center>
</div>
<%end if%>