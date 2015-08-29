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
<title>注册结果</title>
</head>
<style>
td
{
FONT-WEIGHT: normal; FONT-SIZE: 10pt; FONT-STYLE: normal; FONT-VARIANT: normal; color: #000000
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
<body topmargin="0" leftmargin="0" bgcolor="#E7E7E7">
<%
function IsValidEmail(email)

'Check for valid syntax in an email address.

IsValidEmail = true
names = Split(email, "@")
if UBound(names) <> 1 then
   IsValidEmail = false
   exit function
end if
for each name in names
   if Len(name) <= 0 then
     IsValidEmail = false
     exit function
   end if
   for i = 1 to Len(name)
     c = Lcase(Mid(name, i, 1))
     if InStr("abcdefghijklmnopqrstuvwxyz_-.", c) <= 0 and not IsNumeric(c) then
       IsValidEmail = false
       exit function
     end if
   next
   if Left(name, 1) = "." or Right(name, 1) = "." then
      IsValidEmail = false
      exit function
   end if
next
if InStr(names(1), ".") <= 0 then
   IsValidEmail = false
   exit function
end if
i = Len(names(1)) - InStrRev(names(1), ".")
if i <> 2 and i <> 3 then
   IsValidEmail = false
   exit function
end if
if InStr(email, "..") > 0 then
   IsValidEmail = false
end if

end function

sub error()
%> <div align="center">
  <center>
      <table border="1" cellspacing="3" width="350" id="AutoNumber1" bordercolor="#FF6600" style="border-collapse: collapse" cellpadding="0">
          <tr>
                <td width="100%" bgcolor="#FF9966"><p align="center">
                   --==发生错误==--</td>
          </tr>
          <tr>
                <td width="100%"><b>产生错误的可能原因：<br><%=errmsg%></td>
          </tr>
          <tr>
                <td width="100%" bgcolor="#FF9966"><p align="center">
          <INPUT value="返回上一步" name=return type=button onclick="javascript:history.go(-1)"></td>
          </tr>
      </table>
  </center></div>
<%
end sub
founderr=false
if request("UserName")="" or len(request("UserName"))>12 then
	errmsg=errmsg+"<br>"+"<li>用户名最长12个字节，等于6个全角汉字！！！"
	founderr=true
else
	UserName=trim(request("UserName"))
	end if
if request("password")="" or Len(request("password"))>10 then
	errmsg=errmsg+"<br>"+"<li>请输入您的密码(长度不能大于10)。"
	founderr=true
else
	password=request("password")
end if
if password<>request("password2") then
	errmsg=errmsg+"<br>"+"<li>您输入的密码和确认密码不一致。"
	founderr=true
end if
if IsValidEmail(trim(request("Email")))=false then
	errmsg=errmsg+"<br>"+"<li>您的Email有错误。"
	founderr=true
else
	Email=trim(request("Email"))
	
end if
if request("answer")="" or request("quesion")="" then
errmsg=errmsg+"<br>"+"<li>密码提示问题或密码提示问题答案不能为空!!!"
founderr=true
end if
if Instr(request("UserName"),"=")>0 or Instr(request("UserName"),"%")>0 or Instr(request("UserName"),chr(32))>0 or Instr(request("UserName"),"?")>0 or Instr(request("UserName"),"&")>0 or Instr(request("UserName"),";")>0 or Instr(request("UserName"),",")>0 or Instr(request("UserName"),"'")>0 or Instr(request("UserName"),",")>0 or Instr(request("UserName"),chr(34))>0 or Instr(request("UserName"),chr(9))>0 or Instr(request("UserName"),"")>0 or Instr(request("UserName"),"$")>0 or Instr(request("UserName"),"<")>0 or Instr(request("UserName"),">")>0 then
	errmsg=errmsg+"<br>"+"<li>用户名中含有非法字符，您只能使用英文字母和数字！！！"
	founderr=true

end if
if request("oicq")<>""then
if not isnumeric(request("oicq")) or len(request("oicq"))>20 then
			errmsg=errmsg+"<br>"+"<li>Oicq号码只能是4-20位数字，如果没有您可以选择不输入。"
			founderr=true
end if
else
end if
if founderr=true then
	call error()
	Response.Write "错误！！！"
else

	set rs=server.createobject("adodb.recordset")
	sql="select * from user where username='"&username&"'"
	rs.open sql,conn,1,3
	if not rs.eof or username=WebName then
		errmsg="<br>"+"<li>对不起，您输入的用户名已经被注册，请重新输入。"
		founderr=true
	else
		rs.addnew
		rs("username")=username
		rs("password")=password
		rs("email")=email
       	rs("sex")=request("sex")
		rs("quesion")=request("quesion")
		rs("answer")=request("answer")
		rs("oicq")=request("oicq")
       	rs("Address")=request("Address")
       	rs("homepage")=request("homepage")
       	rs("Tel")=request("Tel")
		rs("regDate")=NOW()
		rs("loginIP")=Request.ServerVariables("REMOTE_ADDR")
		rs.update
		rs.close
	end if
	if founderr=true then
		call error()
	else
%>
<div align="center">
  <center>
      <table border="1" cellspacing="3" width="350" id="AutoNumber2" bordercolor="#FF6600" height="291" style="border-collapse: collapse" cellpadding="0">
          <tr>
                <td width="348" colspan="2" bgcolor="#FF9966" height="31"><p align="center">
                   --==恭喜您成功完成注册==--</td>
          </tr>
          <tr>
                <td width="123" align="right" height="16">用户名：</td>
                <td width="224" height="16"><%=request("username")%></td>
          </tr>
          <tr>
                <td width="123" align="right" height="16">密码：</td>
                <td width="224" height="16"><%=request("password")%></td>
          </tr>
          <tr>
                <td width="123" align="right" height="16">密码提示问题：</td>
                <td width="224" height="16"><%=request("quesion")%></td>
          </tr>
          <tr>
                <td width="123" align="right" height="16">答案：</td>
                <td width="224" height="16"><%=request("answer")%></td>
          </tr>
          <tr>
                <td width="123" align="right" height="16">邮箱：</td>
                <td width="224" height="16"><%=request("email")%></td>
          </tr>
          <tr>
                <td width="123" align="right" height="16">性别：</td>
                <td width="224" height="16"><%if request("sex")="True" then%>男<%else%>女<%end if%></td>
          </tr>
          <tr>
                <td width="123" align="right" height="16">联系地址：</td>
                <td width="224" height="16"><%if request("Address")="" then%>未注册<%else%><%=request("Address")%><%end if%></td>
          </tr>
          <tr>
                <td width="123" align="right" height="16">电话号码:</td>
                <td width="224" height="16"><%if request("Tel")="" then%>未注册<%else%><%=request("Tel")%><%end if%></td>
          </tr>
          <tr>
                <td width="123" align="right" height="16">OICQ：</td>
                <td width="224" height="16"><%=request("oicq")%></td>
          </tr>
          <tr>
                <td width="123" align="right" height="16">网站网址：</td>
                <td width="224" height="16"><%if request("homepage")="" then%>未注册<%else%><%=request("homepage")%><%end if%></td>
          </tr>
          <tr>
                <td width="348" colspan="2" height="16"><p align="center">
                   请牢记您的注册信息！</td>
          </tr>
          <tr>
                <td width="348" colspan="2" bgcolor="#FF9966" height="33"><p align="center">
          <INPUT value="关闭此窗口" name=closewindows type=button onclick="javascript:self.close()"></td>
          </tr>
      </table>
  </center></div>
 <%
	end if
	set rs=nothing
end if
conn.close
set conn=nothing
%> 

</body>

</html>