<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
randomize timer
regjm=int(rnd*9998)+1
%>
<html>
<head>
<title>�޸������wWw.happyjh.com��</title>
<LINK href="../css.css" rel=stylesheet>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312"></head>
<body bgcolor="#006699">
<center>
  <table width=360 border=1 align=center cellpadding="5" cellspacing="10" bgcolor="#000000">
    <tr bgcolor="#FFFFFF" align="center">
      <td height="209" bgcolor="#006699"> 
        <p><font color="#FFFFFF">�� �� �� ��</font></p>
        <table border="0" width="288" height="257">
          <tr>
<td>
<form method=POST action='modifyok.asp'>
                <div align="center"><font color="#FFFFFF">����ID��</font> 
                  <input type=password name=id size=12 maxlength="10">
                  <br>
                  <font color="#FFFFFF">�ա�����</font> 
                  <input type=text name=name size=12 maxlength="10">
                  <br>
                  <font color="#FFFFFF">ԭ���룺</font> 
                  <input type=password name=oldpass size=12 maxlength="10">
                  <br>
                  <font color="#FFFFFF">�����룺</font> 
                  <input type=password name=pass size=12 maxlength="10">
                  <br>
                  <font color="#FFFFFF">ȷ���ϣ�</font> 
                  <input type=password name=repass size=12 maxlength="10">
                  <br>
                  <br>
                  <br>
                  <p> 
                    <input type=submit value=�޸� name="submit">
                    <input type=button value=�ر� onClick="window.close()" name="button">
                  </p>
                </div>
</form>
</td>
</tr>
<tr>
<td style='color:red;font-size:9pt'> <font color="#FF6600">˵����<br>
����IE5���¼��������룬<br>
������������������ͨ���������ʷ��¼��<br>
��ɾ����¼�������ʺű����� </font></td>
</tr>
</table>
</table>

</center>
</body>
</html>
