<!--#include file="conn.asp"-->
<%
 if request.cookies("username")="" or request.cookies("password")="" or	request.cookies("okerer")=""then
   response.write"<script>alert('��δ��½�޷��޸ģ�');window.close();</Script>"
   response.end
 end if
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
if request("password")="" or Len(request("password"))>10 then
		response.write("���볤�Ȳ��ɴ���10��<a href=### onclick=window.history.go(-1);>�����ﷵ��</a>")
		response.end
else
	password=request("password")
end if

if IsValidEmail(trim(request("email")))=false then
		response.write("�����ʽ����<a href=### onclick=window.history.go(-1);>�����ﷵ��</a>")
		response.end
else
	email=trim(request("email"))
	
end if
if request("oicq")<>""then
if not isnumeric(request("oicq")) or len(request("oicq"))>20 then
		response.write("OICQ������������20λ��<a href=### onclick=window.history.go(-1);>�����ﷵ��</a>")
		response.end
end if
end if
%>

<%
set rs=server.createobject("adodb.recordset")
sql="select * from user where username='"&request("username")&"'"
rs.open sql,conn,1,3
rs("password")=request.form("password")
rs("email")=request.form("email")
rs("sex")=request("sex")
rs("homepage")=request.form("homepage")
rs("tel")=request.form("tel")
rs("address")=request.form("address")
rs("oicq")=request.form("oicq")
rs.update
response.cookies("password")=rs("password")
rs.close
set rs=nothing
conn.close
set conn=nothing
response.redirect "usedit.asp"
%>