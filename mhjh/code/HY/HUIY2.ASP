<!--#include file="../config.asp"--><%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
username=session("yx8_mhjh_username")
if username="" then Response.Redirect "../error.asp?id=016"
set conn=server.createobject("adodb.connection")
conn.open Application("yx8_mhjh_connstr")
set rst=server.CreateObject ("adodb.recordset")
sql="SELECT 会员 FROM 用户 where 姓名='" &username& "'"
Set Rs=conn.Execute(sql)
if rs.bof or rs.eof then
response.write "你不是剑侠中人或者连接超时"
conn.Close             
set conn=nothing 
response.end
else
hy=rs("会员")
if yx8_mhjh_fellow=false then
%>
<script language=vbscript>
MsgBox "错误！会员功能尚未开放！"
location.href = "javascript:history.back()"
</script>
<%
else
if hy=false then
%>
<script language=vbscript>
MsgBox "错误！你不是会员，请回吧！"
location.href = "javascript:history.back()"
</script>
<%
else
e_mail=Request.Form("e_mail")
sign=Request.Form("qianming")
faceimg=Request.Form("faceimg")
if instr(e_mail,".")=0 or instr(e_mail,"@")=0 or instr(e_mail,"'")<>0 or instr(e_mail,chr(34))<>0 or instr(e_mail,"\")<>0 or instr(e_mail,"&")<>0 then Response.Redirect "error.asp?id=055"
if instr(sign,"file:")<>0 or instr(sign,"script:")<>0 or instr(sign,"'")<>0 or instr(sign,chr(34))<>0 or instr(sign,"\")<>0 or instr(sign,"&")<>0  then Response.Redirect "error.asp?id=037"
if instr(faceimg,"file:")<>0 or instr(faceimg,"script:")<>0 or instr(faceimg,"'")<>0 or instr(faceimg,chr(34))<>0 or instr(faceimg,"&")<>0  then Response.Redirect "error.asp?id=037"
conn.Execute "update 用户 set 电子邮箱='"&e_mail&"',签名档='"&sign&"',头像='"&faceimg&"' where 姓名='"&username&"'"
conn.Close
set conn=nothing
%>
<head>
<link rel="stylesheet" href="style.css">
</head>
<body background='../chatroom/bg1.gif'  oncontextmenu=self.event.returnValue=false topmargin="100">
<p align=center><font color=ff0000 size=5>会员资料更新完成，请重新登陆</font></p>
</body>
<%
end if
end if
end if%>