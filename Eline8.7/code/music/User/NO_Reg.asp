<div align="center"> 
  <table width="766" border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td bgcolor="#FFFFFF"> <br>
        <table border="0" cellspacing="0" cellpadding="10" align="center" bgcolor="#9CD7FF">
          <tr>
            <td bgcolor="#E6F4FF"> 
              <table border="0" cellspacing="0" width="502" align="center">
                <tr> 
                  <td width="100%"> 
                    <div align="center"> 
                      <table border="0" cellpadding="4" cellspacing="1" width="502">
                        <tr align="center"> 
                          <td colspan="2" bgcolor="#9FD6F8">用户注册</td>
                        </tr>
                        <form name=reg action=RegPost.asp method=post onSubmit='return validate(this)'>
                          <tr> 
                            <td width="180" bgcolor="#9FD6F8" align="right">登录ID：</td>
                            <td width="320" bgcolor="#9FD6F8"> 
                              <input maxlength=20 size=16 name=UserName style="color: #000000; background-color: #ffffff; border: 1 solid #000000">
                              * 可以是数字或字母或者中文</td>
                          </tr>
                          <tr> 
                            <td width="180" bgcolor="#b4def8" align="right">用户密码：</td>
                            <td width="320" bgcolor="#b4def8"> 
                              <input type=password maxlength=20 size=16 name=password style="color: #000000; background-color: #ffffff; border: 1 solid #000000">
                              *</td>
                          </tr>
                          <tr> 
                            <td width="180" bgcolor="#9FD6F8" align="right">验证密码：</td>
                            <td width="320" bgcolor="#9FD6F8"> 
                              <input type=password maxlength=16 size=16 name=password2 style="color: #000000; background-color: #ffffff; border: 1 solid #000000">
                              *</td>
                          </tr>
                          <tr> 
                            <td width="180" bgcolor="#b4def8" align="right">E-mail：</td>
                            <td width="320" bgcolor="#b4def8"> 
                              <input maxlength=50 size=20 name=Email style="color: #000000; background-color: #ffffff; border: 1 solid #000000">
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
                              <input maxlength=20 size=20 name=Name style="color: #000000; background-color: #ffffff; border: 1 solid #000000">
                              * </td>
                          </tr>
                          <tr> 
                            <td width="180" bgcolor="#9FD6F8" align="right">OICQ：</td>
                            <td width="320" bgcolor="#9FD6F8"> 
                              <input maxlength=20 size=20 name=Shenfenzheng style="color: #000000; background-color: #ffffff; border: 1 solid #000000">
                              *</td>
                          </tr>
                          <tr> 
                            <td width="180" bgcolor="#b4def8" align="right">联系电话：</td>
                            <td width="320" bgcolor="#b4def8"> 
                              <input maxlength=20 size=20 name=Tel style="color: #000000; background-color: #ffffff; border: 1 solid #000000">
                              *</td>
                          </tr>
                          <tr> 
                            <td width="180" bgcolor="#9FD6F8" align="right">详细地址：</td>
                            <td width="320" bgcolor="#9FD6F8"> 
                              <input maxlength=20 size=20 name=Address style="color: #000000; background-color: #ffffff; border: 1 solid #000000">
                              *</td>
                          </tr>
                          <tr> 
                            <td width="180" bgcolor="#b4def8" align="right">邮政编码：</td>
                            <td width="320" bgcolor="#b4def8"> 
                              <input maxlength=20 size=20 name=Youbian style="color: #000000; background-color: #ffffff; border: 1 solid #000000">
                              *</td>
                          </tr>
                          <tr> 
                            <td width="500" bgcolor="#67BDF1" align="center" colspan="2"> 
                              <input value="注 册" name=Submit type=submit>
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
                  <td width="100%" height="12"></td>
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
if (theform.UserName.value=="" || theform.password.value==""|| theform.password2.value==""|| theform.Email.value==""|| theform.Name.value==""|| theform.Shenfenzheng.value==""|| theform.Address.value==""|| theform.Youbian.value==""|| theform.Tel.value=="") {
alert("错误! 必须认真填写所有项 !");
return false; }
}
//-->
</script>                 