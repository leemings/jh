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
<title>ע����</title>
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
                   --==��������==--</td>
          </tr>
          <tr>
                <td width="100%"><b>��������Ŀ���ԭ��<br><%=errmsg%></td>
          </tr>
          <tr>
                <td width="100%" bgcolor="#FF9966"><p align="center">
          <INPUT value="������һ��" name=return type=button onclick="javascript:history.go(-1)"></td>
          </tr>
      </table>
  </center></div>
<%
end sub
founderr=false
if request("UserName")="" or len(request("UserName"))>12 then
	errmsg=errmsg+"<br>"+"<li>�û����12���ֽڣ�����6��ȫ�Ǻ��֣�����"
	founderr=true
else
	UserName=trim(request("UserName"))
	end if
if request("password")="" or Len(request("password"))>10 then
	errmsg=errmsg+"<br>"+"<li>��������������(���Ȳ��ܴ���10)��"
	founderr=true
else
	password=request("password")
end if
if password<>request("password2") then
	errmsg=errmsg+"<br>"+"<li>������������ȷ�����벻һ�¡�"
	founderr=true
end if
if IsValidEmail(trim(request("Email")))=false then
	errmsg=errmsg+"<br>"+"<li>����Email�д���"
	founderr=true
else
	Email=trim(request("Email"))
	
end if
if request("answer")="" or request("quesion")="" then
errmsg=errmsg+"<br>"+"<li>������ʾ�����������ʾ����𰸲���Ϊ��!!!"
founderr=true
end if
if Instr(request("UserName"),"=")>0 or Instr(request("UserName"),"%")>0 or Instr(request("UserName"),chr(32))>0 or Instr(request("UserName"),"?")>0 or Instr(request("UserName"),"&")>0 or Instr(request("UserName"),";")>0 or Instr(request("UserName"),",")>0 or Instr(request("UserName"),"'")>0 or Instr(request("UserName"),",")>0 or Instr(request("UserName"),chr(34))>0 or Instr(request("UserName"),chr(9))>0 or Instr(request("UserName"),"��")>0 or Instr(request("UserName"),"$")>0 or Instr(request("UserName"),"<")>0 or Instr(request("UserName"),">")>0 then
	errmsg=errmsg+"<br>"+"<li>�û����к��зǷ��ַ�����ֻ��ʹ��Ӣ����ĸ�����֣�����"
	founderr=true

end if
if request("oicq")<>""then
if not isnumeric(request("oicq")) or len(request("oicq"))>20 then
			errmsg=errmsg+"<br>"+"<li>Oicq����ֻ����4-20λ���֣����û��������ѡ�����롣"
			founderr=true
end if
else
end if
if founderr=true then
	call error()
	Response.Write "���󣡣���"
else

	set rs=server.createobject("adodb.recordset")
	sql="select * from user where username='"&username&"'"
	rs.open sql,conn,1,3
	if not rs.eof or username=WebName then
		errmsg="<br>"+"<li>�Բ�����������û����Ѿ���ע�ᣬ���������롣"
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
                   --==��ϲ���ɹ����ע��==--</td>
          </tr>
          <tr>
                <td width="123" align="right" height="16">�û�����</td>
                <td width="224" height="16"><%=request("username")%></td>
          </tr>
          <tr>
                <td width="123" align="right" height="16">���룺</td>
                <td width="224" height="16"><%=request("password")%></td>
          </tr>
          <tr>
                <td width="123" align="right" height="16">������ʾ���⣺</td>
                <td width="224" height="16"><%=request("quesion")%></td>
          </tr>
          <tr>
                <td width="123" align="right" height="16">�𰸣�</td>
                <td width="224" height="16"><%=request("answer")%></td>
          </tr>
          <tr>
                <td width="123" align="right" height="16">���䣺</td>
                <td width="224" height="16"><%=request("email")%></td>
          </tr>
          <tr>
                <td width="123" align="right" height="16">�Ա�</td>
                <td width="224" height="16"><%if request("sex")="True" then%>��<%else%>Ů<%end if%></td>
          </tr>
          <tr>
                <td width="123" align="right" height="16">��ϵ��ַ��</td>
                <td width="224" height="16"><%if request("Address")="" then%>δע��<%else%><%=request("Address")%><%end if%></td>
          </tr>
          <tr>
                <td width="123" align="right" height="16">�绰����:</td>
                <td width="224" height="16"><%if request("Tel")="" then%>δע��<%else%><%=request("Tel")%><%end if%></td>
          </tr>
          <tr>
                <td width="123" align="right" height="16">OICQ��</td>
                <td width="224" height="16"><%=request("oicq")%></td>
          </tr>
          <tr>
                <td width="123" align="right" height="16">��վ��ַ��</td>
                <td width="224" height="16"><%if request("homepage")="" then%>δע��<%else%><%=request("homepage")%><%end if%></td>
          </tr>
          <tr>
                <td width="348" colspan="2" height="16"><p align="center">
                   ���μ�����ע����Ϣ��</td>
          </tr>
          <tr>
                <td width="348" colspan="2" bgcolor="#FF9966" height="33"><p align="center">
          <INPUT value="�رմ˴���" name=closewindows type=button onclick="javascript:self.close()"></td>
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