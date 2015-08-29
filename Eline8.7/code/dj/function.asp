<%
function HTMLEncode(fString)
if not isnull(fString) then
    fString = replace(fString, ">", "&gt;")
    fString = replace(fString, "<", "&lt;")

    fString = Replace(fString, CHR(32), "&nbsp;")
    fString = Replace(fString, CHR(34), "&quot;")
    fString = Replace(fString, CHR(39), "&#39;")
    fString = Replace(fString, CHR(13), "")
    fString = Replace(fString, CHR(10) & CHR(10), "</P><P> ")
    fString = Replace(fString, CHR(10), "<BR> ")
else
	fString=""
end if
HTMLEncode = fString
end function

function HTMLEncode2(fString)
if not isnull(fString) then
	fString = replace(fString,"<BR>",chr(13))
	fString = replace(fString,"&nbsp;","&nbsp;")
		fString = replace(fString,"'","’")
else
	fString=""
end if
HTMLEncode2 = fString
end function

function Checkstr(str)
str=replace(str,"'","''")
Checkstr=str
End function

function HTMLcode(fString)
if not isnull(fString) then
    fString = Replace(fString, CHR(13), "")
    fString = Replace(fString, CHR(10) & CHR(10), "</P><P>")
    fString = Replace(fString, CHR(10), "<BR>")
    HTMLcode = fString
end if
HTMLEncode = fString
end function

function Checkstr(str)
str=replace(str,"'","''")
Checkstr=str
End function

'---------------------检查用户名密码-------------------------------
function Checkin(s) 
s=trim(s) 
s=replace(s," ","&amp;nbsp;") 
s=replace(s,"'","&amp;#39;") 
s=replace(s,"""","&amp;quot;") 
s=replace(s,"&lt;","&amp;lt;") 
s=replace(s,"&gt;","&amp;gt;") 
Checkin=s 
end function 

'---------------------检查管理员-------------------------------
function CheckAdmin
	if Session("IsAdmin")<>true then response.redirect "admin_login.asp"
end function

'---------------------检查用户Email-------------------------------
function IsValidEmail(email)
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

'---------------------错误输出-------------------------------
sub error()
%>
<html>
<head><title>错误信息提示==<%=Webname%></title></head>
<style>
td
{
FONT-WEIGHT: normal; FONT-SIZE: 9pt; FONT-STYLE: normal; FONT-VARIANT: normal; color: #000000
}
A {
	TEXT-DECORATION: none
}
A:link {
	COLOR: #000000
}
A:visited {
	COLOR: #333333
}
A:active {
	COLOR: #FF0000
}
A:hover {
	COLOR: #333333; TEXT-DECORATION: underline overline
}


</style>
<body topmargin="0" leftmargin="0" text="#111111" bgcolor="#FFFFFF">
<div align="center">
  <table cellpadding=0 cellspacing=0 border=0 width="100%" style="border-collapse: collapse" bordercolor="#111111">
    <tr>
      <td>
        <div align="center">
          <center>
        <table cellspacing=0 border=1 width="100%" bordercolor="#000000" height="79" cellpadding="0" style="border-collapse: collapse">
          <tr align="center"> 
            <td width="100%" height="11" bgcolor="#FFCC66">错误信息：</td>
          </tr>
          <tr> 
            <td width="100%" height="32" bgcolor="#FF9933"><b>产生错误的可能原因：<br><%=errmsg%></td>
          </tr>
          <tr align="center">
            <td width="100%" height="16" bgcolor="#FFCC66"><p><b><font color="#FF0000"><a href="javascript:this.location.reload()"><font color="#FF0000">刷新重试</font></a></a> <a href="#" ONCLICK="window.history.go(-1);">点这里后退</a></font></b></td>
          </tr>  
        </table>
          </center>
        </div>
      </td>
    </tr>
  </table>
</div>
</html>
<%
end sub

%>