<!--#include file="conn.asp"-->
<%
 if request.cookies("username")="" or request.cookies("password")="" or	request.cookies("okerer")=""then
   response.write"<script>alert('尚未登陆无法修改！');window.close();</Script>"
   response.end
 end if

set rs=server.createobject("adodb.recordset")
sql="select * from user where username='"&request.cookies("username")&"' and password='"&request.cookies("password")&"'"
rs.open sql,conn,1,3
username=rs("username")
Password=rs("Password")
email=rs("email")
sex=rs("sex")
quesion=rs("quesion")
answer=rs("answer")
tel=rs("tel")
oicq=rs("oicq")
address=rs("address")
homepage=rs("homepage")
%>
<html>

<head>
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>修改资料</title>
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

<body topmargin="0" leftmargin="0">
<div align="center">
  <center>
  <table border="1" bordercolor="#FF6600" width="350" id="AutoNumber1" height="100%" cellpadding="0" style="border-collapse: collapse">
  <form action="usesave.asp" method=post name=edit onSubmit="return formCheck()">
    <tr>
            <td width="100%" align="right" bgcolor="#FF9966" height="20" colspan="2">
            <p align="center">用户资料修改</td>
          </tr>
    <tr>
            <td width="28%" align="right" height="23">用户名：</td>
            <td width="72%" height="23"><%=username%></td>
          </tr>
          <tr>
            <td width="28%" align="right" height="23">性别：</td>
            <td width="72%" height="23"><select size="1" name="sex">
            <option value="True"<%if sex=true then%> selectd<%end if%>>男</option>
            <option value="false"<%if sex=false then%> selected<%end if%>>女</option>
            </select>*</td>
          </tr>
          <tr>
            <td width="28%" align="right" height="23">密码：</td>
            <td width="72%" height="23">
            <input type="password" name="password" size="18" value="<%=password%>">*</td>
          </tr>
          <tr>
            <td width="28%" align="right" height="23">密码提示问题：</td>
            <td width="72%" height="23">
            <input type="text" name="quesion" size="29" disabled value="<%=quesion%>">*</td>
          </tr>
          <tr>
            <td width="28%" align="right" height="23">答案：</td>
            <td width="72%" height="23">
            <input type="text" name="answer" size="29" disabled value="<%=answer%>">*</td>
          </tr>
          <tr>
            <td width="28%" align="right" height="23">邮箱：</td>
            <td width="72%" height="23">
            <input type="text" name="email" size="20" value="<%=email%>">*</td>
          </tr>
          <tr>
            <td width="28%" align="right" height="23">OICQ：</td>
            <td width="72%" height="23">
            <input type="text" name="oicq" size="15" value="<%=oicq%>" onkeyup='this.value=this.value.replace(/\D/gi,"")'>*</td>
          </tr>
          <tr>
            <td width="28%" align="right" height="23">联系地址：</td>
            <td width="72%" height="23">
            <input type="text" name="address" size="33" value="<%=address%>"></td>
          </tr>
          <tr>
            <td width="28%" align="right" height="23">电话：</td>
            <td width="72%" height="23">
            <input type="text" name="tel" size="26" value="<%=tel%>"></td>
          </tr>
          <tr>
            <td width="28%" align="right" height="23">个人网址：</td>
            <td width="72%" height="23">
            <input type="text" name="homepage" size="33" value="<%=homepage%>">
      </td>
    </tr>
    <tr>
            <td width="100%" align="right" bgcolor="#FF9966" height="6" colspan="2">
            <p align="center">
                   <font color="#FFFFFF">
                   <input type="submit" value="修 改" name="edit">
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

if (document.edit.sex.value =="")
{
alert("请选择性别!");
edit.sex.focus();
return false;
}


if (document.edit.password.value == "")
{
alert("请填写您的密码!");
edit.password.focus();
return false;
}
if (document.edit.email.value =="")
{
alert("请输入您的Email!");
edit.email.focus();
return false;
}
if (document.edit.email.value =="@")
{
alert("请输入您的Email!");
edit.email.focus();
return false;
}
if (document.edit.oicq.value =="")
{
alert("请输入您的oicq号，本站需要真实记录!");
edit.oicq.focus();
return false;
}

if (document.edit.sex.value =="")
{
alert("请选择性别!");
edit.sex.focus();
return false;
}
}
  </script>

  </center>
</div>

</body>

</html>