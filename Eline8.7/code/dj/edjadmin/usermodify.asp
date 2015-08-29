<!--#include file="admin_top.asp"-->
<!--#include file="checkadmin.inc"-->
<%
id=request("id")
set rs=server.createobject("adodb.recordset")
sql="select * from user where id="&id
rs.open sql,conn,1,1
if rs.eof then
	response.write "操作错误！"
	Response.end
else
	Username=rs("Username")
	Password=rs("Password")
	Sex=rs("Sex")
	email=rs("Email")
	oicq=rs("oicq")
	homepage=rs("homepage")
	Address=rs("Address")
	answer=rs("answer")
	quesion=rs("quesion")
	tel=rs("tel")
rs.close
end if
%>
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
<table border="0" width="760" cellspacing="0" cellpadding="0" style="border-collapse: collapse" height="155">
  <tr>
    <td valign=top width=150 height="155">
    <!--#include file="admin_left.asp"-->
    　</td>
    <td valign=top width=10 height="155">　</td>
    <td valign=top width="600" height="155">
    <table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="100%" id="AutoNumber3" height="100%" bordercolor="#000000">
      <tr>
        <td width="100%" bgcolor="#FF9933" valign="top" height="84">
    <table border="0" cellpadding="0" style="border-collapse: collapse" bordercolor="#FF9933" width="100%" id="AutoNumber3">
        <form method="POST" action="UserSave.asp?id=<%=id%>" name=reg2>
          <tr>
            <td width="100%" colspan=2  align=center bgcolor="#CC6600" height="23"><b>修 改 用 户 资 料</b></td>
          </tr>
          <tr>
            <td width="30%" align="right">用户名：</td>
            <td width="70%"><input type="text" name="username" value="<%=Username%>" size="20"></td>
          </tr>
          <tr>
            <td width="30%" align="right" valign="top">密码：</td>
            <td width="70%">
            <input type="password" name="password" value="<%=Password%>" size="20"></td>
          </tr>
          <tr>
                <TD align=right>密码提示问题：</TD>
                <TD>
                  	 <input type="password" name="quesion" value="<%=quesion%>" size="20"></TD>
              </tr>
          <tr>
                <TD align=right>问题答案：</TD>
                <TD>
                  	 <input type="password" name="answer" value="<%=answer%>" size="20"></TD>
              </tr>
          <tr>
                <TD align=right>性 别：</TD>
                <TD>
                  	 <font color="#FFFFFF">
                  	 <select size="1" name="sex">
                     <option value="0"<%if sex=false then%> selected<%end if%>>女孩</option>
                     <option value="1"<%if sex=true then%> selected<%end if%>>男孩</option>
                     </select></font></TD>
              </tr>
          <tr>
            <td width="30%" align="right" valign="top" height="20">信箱：</td>
            <td width="70%"><input type="text" name="Email" value="<%=Email%>" size="20"></td>
          </tr>
          <tr>
            <td width="30%" align="right" valign="top" height="20">OICQ号码：</td>
            <td width="70%"><input type="text" name="oicq" value="<%=oicq%>" size="20"></td>
          </tr>
          <tr>
          <TD align=right>联系地址：</TD>
          <TD> 
              <input size=40 name=Address value="<%=Address%>"></TD>
              </tr>
          <tr>
          <TD align=right>电话或手机：</TD>
          <TD><INPUT size=20 name=TEL value="<%=TEL%>"></TD>
              </tr>
          <tr>
          <TD align=right>主页地址：</TD>
          <TD><INPUT size=20 name=homepage value="<%=homepage%>">注意要填写完整如：http://www.xn163.com/dj</TD>
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

</td>
      </tr>
    </table>
    </td>
  </tr>
  </table>
  </center>
</div>