<!--#include file="conn.asp"-->
<!--#include file="const.asp"-->
<%
if newreg=false then
response.write("对不起，本站暂停注册!")
response.end
end if
%>
<html>

<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>用户注册</title>
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
      <p align="center">用用户注册协议</td>
    </tr>
    <tr>
      <td width="100%" height="135" valign="top">
      注册前请仔细阅读以下注册协议：<ol>
        <li>本站所有舞曲均为原创作品，未经授权任何会员不可利用会员帐号权限盗链本站作品，否则视为侵仅！</li>
        <li>禁止重复注册帐号，请不要无味的注册垃圾帐号。一经发封IP、删除帐号。</li>
        <li>不得发表含政治、色情、暴力、污辱、诽谤、反动等有违国家法纪的信息。</li>
        <li>会员个人信息有权保密，未经同意，任何人无权将其他会员帐号、密码、邮箱等个人信息向外界泄漏。</li>
        <li>不得利用程序漏洞恶意攻击本站其他的注册用户，因此造成的一切后果将由您来自己负责。</li>
        <li>不得私自将个人帐号外借他人使用！</li>
        <li>本站将定期清理长时间没有登陆过本站的用户，请不要将任何重要信息保留在本站的短信箱中。否则因此给您造成的任何损失和不便本站不负任何责任。</li>
      </ol>
      </td>
    </tr>
    <tr>
      <td width="100%" bgcolor="#FF9966" height="25">
      <p align="center">
                   <font color="#FFFFFF">
                   <input type="hidden" value="yes" name="act">
                   <input type="submit" value="同  意" name="B3">
                   <input type="button" value="不同意" name="Close" onclick="javascript:self.close()"></font></td>
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
            <p align="center">用户注册</td>
          </tr>
    <tr>
            <td width="28%" align="right" height="23">用户名：</td>
            <td width="72%" height="23">
            <input type="text" name="username" size="15">*</td>
          </tr>
          <tr>
            <td width="28%" align="right" height="23">性别：</td>
            <td width="72%" height="23"><select size="1" name="sex">
            <option value="True">男</option>
            <option value="false">女</option>
            </select>*</td>
          </tr>
          <tr>
            <td width="28%" align="right" height="23">密码：</td>
            <td width="72%" height="23">
            <input type="password" name="password" size="18">*</td>
          </tr>
          <tr>
            <td width="28%" align="right" height="23">确认密码：</td>
            <td width="72%" height="23">
            <input type="password" name="password2" size="18">*</td>
          </tr>
          <tr>
            <td width="28%" align="right" height="23">密码提示问题：</td>
            <td width="72%" height="23">
            <input type="text" name="quesion" size="29">*</td>
          </tr>
          <tr>
            <td width="28%" align="right" height="23">答案：</td>
            <td width="72%" height="23">
            <input type="text" name="answer" size="29">*</td>
          </tr>
          <tr>
            <td width="28%" align="right" height="23">邮箱：</td>
            <td width="72%" height="23">
            <input type="text" name="email" size="20" value="@">*</td>
          </tr>
          <tr>
            <td width="28%" align="right" height="23">OICQ：</td>
            <td width="72%" height="23">
            <input type="text" name="oicq" size="15" value= >*</td>
          </tr>
          <tr>
            <td width="28%" align="right" height="23">联系地址：</td>
            <td width="72%" height="23">
            <input type="text" name="address" size="33"></td>
          </tr>
          <tr>
            <td width="28%" align="right" height="23">电话：</td>
            <td width="72%" height="23">
            <input type="text" name="tel" size="15"></td>
          </tr>
          <tr>
            <td width="28%" align="right" height="23">个人网址：</td>
            <td width="72%" height="23">
            <input type="text" name="homepage" size="33" value="HTTP://WWW.XXDJ.NET">
      </td>
    </tr>
    <tr>
            <td width="100%" align="right" bgcolor="#FF9966" height="6" colspan="2">
            <p align="center">
                   <font color="#FFFFFF">
                   <input type="submit" value="注　册" name="reg">
                   <input type="reset" value="重 写" name="reset"> </font></td>
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
alert("请填写您的用户名!");
reg2.username.focus();
return false;
}
if (lenX(document.reg2.username.value)>12)
{
alert("用户名最长为12个字节，等于6个全角汉字！！！");
reg2.username.focus();
return false;
}
if (document.reg2.sex.value =="")
{
alert("请选择性别!");
reg2.sex.focus();
return false;
}


if (document.reg2.password.value == "")
{
alert("请填写您的密码!");
reg2.password.focus();
return false;
}
if (document.reg2.password2.value != document.reg2.password.value)
{
alert("两次输入的密码不一致，请检查!");
reg2.password2.focus();
return false;
}
if (document.reg2.quesion.value =="")
{
alert("请输入密码提示问题");
reg2.quesion.focus();
return false;
}
if (document.reg2.answer.value =="")
{
alert("请输入问题答案!");
reg2.answer.focus();
return false;
}
if (document.reg2.email.value =="")
{
alert("请输入您的Email!");
reg2.email.focus();
return false;
}
if (document.reg2.email.value =="@")
{
alert("请输入您的Email!");
reg2.email.focus();
return false;
}
if (document.reg2.oicq.value =="")
{
alert("请输入您的oicq号，本站需要真实记录!");
reg2.oicq.focus();
return false;
}

if (document.reg2.sex.value =="")
{
alert("请选择性别!");
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