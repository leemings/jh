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
            <td>�ʺ�**</td>
            <td>
              <input name="account" maxlength=14 size=14>
              6-14λ���ֻ���ĸ</td>
          </tr>
          <tr>
            <td>����**</td>
            <td>
              <input name="username" maxlength=6 size="12">
              2λ�������Ļ�Ƭ����</td>
          </tr>
          <tr>
            <td>����**</td>
            <td>
              <input type="password" name="password" maxlength=14 size=14>
              6-14λ���ֻ���ĸ</td>
          </tr>
          <tr>
            <td>����ȷ��**</td>
            <td>
              <input type="password" name="repassword" maxlength=14 size=14>
              ͬ��</td>
          </tr>
          <tr>
            <td>�Ա�**</td>
            <td>
              <select name="sex">
                <option value="" selected>��ѡ��</option>
                <option value="��">��</option>
                <option value="Ů">Ů</option>
              </select>
            </td>
          </tr>
          <tr>
            <td>E_Mail**</td>
            <td>
              <input name="e_mail" maxlength="30" size="30">
              ����ȷ����</td>
          </tr>
          <tr>
            <td>ǩ����</td>
            <td>
              <input name="sign" maxlength="100" size="40">
            </td>
          </tr>
          <tr>
            <td colspan=2 align=middle>
              <input type="submit" value=" ע �� " name="submit">
              <input type="reset" value=" �� �� " name="reset">
              <input type="button" value=" �� �� " onClick="javascript:top.window.close();" name="button">
            </td>
          </tr>
        </table>
       <hr size="2" width="97%" color="#800000" noshade>
      </td>
    </tr>
  </table>
</FORM></BODY></HTML>
