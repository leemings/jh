<!--#include file="conn.asp"-->
<!--#include file="const.asp"-->
<%
if newreg=false then
response.write("�Բ��𣬱�վ��ͣע��!")
response.end
end if
%>
<html>

<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>�û�ע��</title>
</head>
<style>
td
{
FONT-WEIGHT: normal; FONT-SIZE: 9pt; FONT-STYLE: normal; FONT-VARIANT: normal; color: #000000
}
A {
	TEXT-DECORATION: none
}
A:link {
	COLOR: #FF0000
}
A:visited {
	COLOR: #FF0000
}
A:active {
	COLOR: #FF0000
}
A:hover {
	COLOR: #FF0000; TEXT-DECORATION: underline overline
}


</style>
<body text="#FFFFFF" topmargin="0" leftmargin="0"  bgcolor="#E7E7E7">
<%
if request("act")="yes" then
call reg_2()
else
call reg_1()
end if
sub reg_1()
%>
<div align="center">
  <center>
  <table border="1" cellpadding="3" bordercolor="#FF6600" width="350" id="AutoNumber3" height="211" cellspacing="3" style="border-collapse: collapse">
<form action="usereg.asp" method=post name=reg1>
    <tr>
      <td width="100%" bgcolor="#FF9966" height="25">
      <p align="center">���û�ע��Э��</td>
    </tr>
    <tr>
      <td width="100%" height="135" valign="top">
      ע��ǰ����ϸ�Ķ�����ע��Э�飺<ol>
        <li>��վ����������Ϊԭ����Ʒ��δ����Ȩ�κλ�Ա�������û�Ա�ʺ�Ȩ�޵�����վ��Ʒ��������Ϊ�ֽ���</li>
        <li>��ֹ�ظ�ע���ʺţ��벻Ҫ��ζ��ע�������ʺš�һ������IP��ɾ���ʺš�</li>
        <li>���÷������Ρ�ɫ�顢���������衢�̰�����������Υ���ҷ��͵���Ϣ��</li>
        <li>��Ա������Ϣ��Ȩ���ܣ�δ��ͬ�⣬�κ�����Ȩ��������Ա�ʺš����롢����ȸ�����Ϣ�����й©��</li>
        <li>�������ó���©�����⹥����վ������ע���û��������ɵ�һ�к�����������Լ�����</li>
        <li>����˽�Խ������ʺ��������ʹ�ã�</li>
        <li>��վ����������ʱ��û�е�½����վ���û����벻Ҫ���κ���Ҫ��Ϣ�����ڱ�վ�Ķ������С�������˸�����ɵ��κ���ʧ�Ͳ��㱾վ�����κ����Ρ�</li>
      </ol>
      </td>
    </tr>
    <tr>
      <td width="100%" bgcolor="#FF9966" height="25">
      <p align="center">
                   <font color="#FFFFFF">
                   <input type="hidden" value="yes" name="act">
                   <input type="submit" value="ͬ  ��" name="B3">
                   <input type="button" value="��ͬ��" name="Close" onclick="javascript:self.close()"></font></td>
    </tr>
    </form>
  </table>
  </center>
</div>
<%
end sub
sub reg_2()
%>
<div align="center">
  <center>
  <table border="1" bordercolor="#FF6600" width="350" id="AutoNumber1" height="100%" cellpadding="0" style="border-collapse: collapse">
  <form action="sendreg.asp" method=post name=reg2 onSubmit="return formCheck()">
    <tr>
            <td width="100%" align="right" bgcolor="#FF9966" height="20" colspan="2">
            <p align="center">�û�ע��</td>
          </tr>
    <tr>
            <td width="28%" align="right" height="23">�û�����</td>
            <td width="72%" height="23">
            <input type="text" name="username" size="15">*</td>
          </tr>
          <tr>
            <td width="28%" align="right" height="23">�Ա�</td>
            <td width="72%" height="23"><select size="1" name="sex">
            <option value="True">��</option>
            <option value="false">Ů</option>
            </select>*</td>
          </tr>
          <tr>
            <td width="28%" align="right" height="23">���룺</td>
            <td width="72%" height="23">
            <input type="password" name="password" size="18">*</td>
          </tr>
          <tr>
            <td width="28%" align="right" height="23">ȷ�����룺</td>
            <td width="72%" height="23">
            <input type="password" name="password2" size="18">*</td>
          </tr>
          <tr>
            <td width="28%" align="right" height="23">������ʾ���⣺</td>
            <td width="72%" height="23">
            <input type="text" name="quesion" size="29">*</td>
          </tr>
          <tr>
            <td width="28%" align="right" height="23">�𰸣�</td>
            <td width="72%" height="23">
            <input type="text" name="answer" size="29">*</td>
          </tr>
          <tr>
            <td width="28%" align="right" height="23">���䣺</td>
            <td width="72%" height="23">
            <input type="text" name="email" size="20" value="@">*</td>
          </tr>
          <tr>
            <td width="28%" align="right" height="23">OICQ��</td>
            <td width="72%" height="23">
            <input type="text" name="oicq" size="15" value= >*</td>
          </tr>
          <tr>
            <td width="28%" align="right" height="23">��ϵ��ַ��</td>
            <td width="72%" height="23">
            <input type="text" name="address" size="33"></td>
          </tr>
          <tr>
            <td width="28%" align="right" height="23">�绰��</td>
            <td width="72%" height="23">
            <input type="text" name="tel" size="15"></td>
          </tr>
          <tr>
            <td width="28%" align="right" height="23">������ַ��</td>
            <td width="72%" height="23">
            <input type="text" name="homepage" size="33" value="HTTP://WWW.XXDJ.NET">
      </td>
    </tr>
    <tr>
            <td width="100%" align="right" bgcolor="#FF9966" height="6" colspan="2">
            <p align="center">
                   <font color="#FFFFFF">
                   <input type="submit" value="ע����" name="reg">
                   <input type="reset" value="�� д" name="reset"> </font></td>
    </tr>

              </form>
  </table>
<script language="VBScript">
function lenX(byval uStr)
dim theLen,x,testuStr
theLen = 0
for x = 1 to len(uStr)
testuStr=mid(uStr,x,1)
if asc(testuStr) < 0 then
theLen=theLen + 2
else
theLen=theLen + 1
end if
next
lenX=theLen
end function
  </script>
<script LANGUAGE="JavaScript">
function formCheck()

{
if (document.reg2.username.value == "")
{
alert("����д�����û���!");
reg2.username.focus();
return false;
}
if (lenX(document.reg2.username.value)>12)
{
alert("�û����Ϊ12���ֽڣ�����6��ȫ�Ǻ��֣�����");
reg2.username.focus();
return false;
}
if (document.reg2.sex.value =="")
{
alert("��ѡ���Ա�!");
reg2.sex.focus();
return false;
}


if (document.reg2.password.value == "")
{
alert("����д��������!");
reg2.password.focus();
return false;
}
if (document.reg2.password2.value != document.reg2.password.value)
{
alert("������������벻һ�£�����!");
reg2.password2.focus();
return false;
}
if (document.reg2.quesion.value =="")
{
alert("������������ʾ����");
reg2.quesion.focus();
return false;
}
if (document.reg2.answer.value =="")
{
alert("�����������!");
reg2.answer.focus();
return false;
}
if (document.reg2.email.value =="")
{
alert("����������Email!");
reg2.email.focus();
return false;
}
if (document.reg2.email.value =="@")
{
alert("����������Email!");
reg2.email.focus();
return false;
}
if (document.reg2.oicq.value =="")
{
alert("����������oicq�ţ���վ��Ҫ��ʵ��¼!");
reg2.oicq.focus();
return false;
}

if (document.reg2.sex.value =="")
{
alert("��ѡ���Ա�!");
reg2.sex.focus();
return false;
}
}
  </script>

  </center>
</div>
<%end sub%>
</body>

</html>