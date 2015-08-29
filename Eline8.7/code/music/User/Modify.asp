<%PageName="UserModify"%>
<!--#include file="conn.asp"-->
<!--#include file="function.asp"-->
<!--#include file="Star.INC"-->

<%
if session("username")="" or isnull(session("username")) or session("Password")="" or isnull(session("Password")) then
	errmsg="你没有登陆！"
	call error()
	Response.end 
end if
username=session("username")
Password=session("Password")
set rs=server.createobject("adodb.recordset")
sql="select * from [user] where username='"&username&"' and password='"&password&"'"
rs.open sql,conn,1,1
if rs.eof then
	errmsg="<br>操作错误。请重新登陆！"
	call error()
	Response.end
else
	id=rs("id")
	Sex=rs("Sex")
	Email=rs("Email")
	Name=rs("Name")
	Shenfenzheng=rs("Shenfenzheng")
	Address=rs("Address")
	Youbian=rs("Youbian")
	Tel=rs("Tel")
rs.close
end if
%>
<!--#include file="Top.asp"-->
<div align="center">
  <table width="766" border="0" cellspacing="0" cellpadding="0">
    <tr> 
      <td bgcolor="#FFFFFF"> <br>
        <table border="0" cellspacing="0" cellpadding="10" align="center" bgcolor="#9CD7FF" width="323">
          <tr> 
            <td bgcolor="#E6F4FF"> 
              <table border="0" cellspacing="0" width="502">
                <tr> 
                  <td width="100%"> 
                    <div align="center"> 
                      <table border="0" cellpadding="4" cellspacing="1" width="502">
                        <tr> 
                          <td colspan="2" bgcolor="#9FD6F8" align="center">用户注册资料修改</td>
                        </tr>
                        <form name=UserModify action="Save.asp" method=post onSubmit='return validate(this)'>
                          <tr> 
                            <td width="180" bgcolor="#9FD6F8" align="right">登录ID：</td>
                            <td width="320" bgcolor="#9FD6F8"><%=Username%></td>
                          </tr>
                          <tr> 
                            <td width="180" bgcolor="#b4def8" align="right">用户密码：</td>
                            <td width="320" bgcolor="#b4def8"> 
                              <input type=password maxlength=20 size=16 name=password value="<%=password%>" style="color: #000000; background-color: #ffffff; border: 1 solid #000000">
                              *</td>
                          </tr>
                          <tr> 
                            <td width="180" bgcolor="#9FD6F8" align="right">验证密码：</td>
                            <td width="320" bgcolor="#9FD6F8"> 
                              <input type=password maxlength=16 size=16 name=password2 value="<%=password%>" style="color: #000000; background-color: #ffffff; border: 1 solid #000000">
                              *</td>
                          </tr>
                          <tr> 
                            <td width="180" bgcolor="#b4def8" align="right">E-mail：</td>
                            <td width="320" bgcolor="#b4def8"> 
                              <input maxlength=50 size=20 name=Email value="<%=Email%>" style="color: #000000; background-color: #ffffff; border: 1 solid #000000">
                              *</td>
                          </tr>
                          <tr> 
                            <td width="180" bgcolor="#9FD6F8" align="right">称呼：</td>
                            <td width="320" bgcolor="#9FD6F8"> 
                              <select style="FONT-SIZE: 9pt" name="sex">
                                <option value="先生">先生</option>
                                <option value="女士">女士</option>
                              </select>
                            </td>
                          </tr>
                          <tr> 
                            <td width="180" bgcolor="#b4def8" align="right">真实姓名：</td>
                            <td width="320" bgcolor="#b4def8"> 
                              <input maxlength=20 size=20 name=Name value="<%=Name%>" style="color: #000000; background-color: #ffffff; border: 1 solid #000000">
                              * </td>
                          </tr>
                          <tr> 
                            <td width="180" bgcolor="#9FD6F8" align="right">OICQ：</td>
                            <td width="320" bgcolor="#9FD6F8"> 
                              <input maxlength=20 size=20 name=Shenfenzheng value="<%=Shenfenzheng%>" style="color: #000000; background-color: #ffffff; border: 1 solid #000000">
                              *</td>
                          </tr>
                          <tr> 
                            <td width="180" bgcolor="#b4def8" align="right">联系电话：</td>
                            <td width="320" bgcolor="#b4def8"> 
                              <input maxlength=20 size=20 name=Tel value="<%=Tel%>" style="color: #000000; background-color: #ffffff; border: 1 solid #000000">
                              *</td>
                          </tr>
                          <tr> 
                            <td width="180" bgcolor="#9FD6F8" align="right">详细地址：</td>
                            <td width="320" bgcolor="#9FD6F8"> 
                              <input maxlength=20 size=20 name=Address value="<%=Address%>" style="color: #000000; background-color: #ffffff; border: 1 solid #000000">
                              *</td>
                          </tr>
                          <tr> 
                            <td width="180" bgcolor="#b4def8" align="right">邮政编码：</td>
                            <td width="320" bgcolor="#b4def8"> 
                              <input maxlength=20 size=20 name=Youbian value="<%=Youbian%>" style="color: #000000; background-color: #ffffff; border: 1 solid #000000">
                              *</td>
                          </tr>
                          <tr> 
                            <td width="500" bgcolor="#67BDF1" align="center" colspan="2"> 
                              <input value="修 改" name=Submit type=submit>
                              &nbsp; 
                              <input type=reset value="重 写" name=Submit2>
                            </td>
                          </tr>
                        </form>
                      </table>
                    </div>
                  </td>
                </tr>
                <tr> 
                  <td width="100%" height="16"></td>
                </tr>
              </table>
            </td>
          </tr>
        </table>
      </td>
    </tr>
  </table>
  
</div>        
<div align="center">        
  <center>
  </center>     
</div>        
<script language="javascript">
<!--
function validate(theform) {
if (theform.password.value==""|| theform.password2.value==""|| theform.Email.value==""|| theform.Name.value==""|| theform.Shenfenzheng.value==""|| theform.Address.value==""|| theform.Tel.value=="") {
alert("错 误 ! 必 须 认 真 填 写 所 有 项 !");
return false; }
}
//-->
</script>                 
<!--#include file="Bottom.asp"-->  