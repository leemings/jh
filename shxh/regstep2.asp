<head>
<title><%=Application("Ba_jxqy_systemname")%></title>
<LINK href="style.css" rel=stylesheet>
</head>
<body bgcolor="<%=Application("Ba_jxqy_backgroundcolor")%>" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" background="<%=Application("Ba_jxqy_backgroundimage")%>">
<form action="regstep3.asp" method="post" id=form1 name=form1>
  <table width="100%" border="2" align="center" cellspacing="3" bordercolor="#666666" height="100%">
    <tr>
      <td align="center"> 
        <hr size="2" width="97%" color="#800000" noshade>
        
        <table cellpadding="3" cellspacing="2" width="400">
          <tr>
            <td>帐号**</td>
            <td>
              <input name="account" maxlength=14 size=14>
              6-14位数字或字母</td>
          </tr>
          <tr>
            <td>姓名**</td>
            <td>
              <input name="username" maxlength=6 size="12">
              2位以上中文或片假名</td>
          </tr>
          <tr>
            <td>密码**</td>
            <td>
              <input type="password" name="password" maxlength=14 size=14>
              6-14位数字或字母</td>
          </tr>
          <tr>
            <td>密码确认**</td>
            <td>
              <input type="password" name="repassword" maxlength=14 size=14>
              同上</td>
          </tr>
          <tr>
            <td>性别**</td>
            <td>
              <select name="sex">
                <option value="" selected>请选择</option>
                <option value="男">男</option>
                <option value="女">女</option>
              </select>
            </td>
          </tr>
          <tr>
            <td>E_Mail**</td>
            <td>
              <input name="e_mail" maxlength="30" size="25">
              重要：密码找回需要</td>
          </tr>
		  <tr>
            <td>联系方式</td>
            <td>
              <input name="contact" maxlength="20" size="20"> QQ/手机
			</td>
          </tr>
		  <tr>
            <td>介绍人</td>
            <td>
              <input name="recommend" maxlength="20" size="20"> 没有请留空
              </td>
          </tr>
          <tr>
            <td>签名档</td>
            <td>
              <input name="sign" maxlength="100" size="40" value="快乐江湖">
            </td>
          </tr>
          <tr>
            <td colspan=2 align=middle>
              <input type="submit" value=" 注 册 " name="submit">
              <input type="reset" value=" 重 填 " name="reset">
              <input type="button" value=" 关 闭 " onClick="javascript:top.window.close();" name="button">
            </td>
          </tr>
        </table>
       <hr size="2" width="97%" color="#800000" noshade>
      </td>
    </tr>
  </table>
</FORM></BODY></HTML>
