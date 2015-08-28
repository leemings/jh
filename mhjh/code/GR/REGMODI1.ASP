<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
username=session("yx8_mhjh_username")
usergrade=session("yx8_mhjh_usergrade")
if username="" then
%>
<script language=vbscript>
MsgBox "对不起，你还没有登录！"
window.close()
</script>
<%
elseif usergrade<5 then
%>
<script language=vbscript>
MsgBox "对不起，你等级不够5级！"
window.close()
</script>
<%
else
%>
<%
if len(touxian)>4 then Response.Redirect "../error.asp?id=004"
name=session("yx8_mhjh_username")
email=server.HTMLEncode(trim(request.form("email")))
oicq=server.HTMLEncode(trim(request.form("oicq")))
touxiang=request.form("touxiang")
touxian=server.HTMLEncode(trim(request.form("touxian")))
qianming=server.HTMLEncode(trim(request.form("qianming")))
for each element in Request.Form
elevalue=Request.Form(element)
if instr(elevalue,"'")<>0 or instr(elevalue,chr(34))<>0 or instr(elevalue,"\")<>0 or instr(elevalue,";") then Response.Redirect "../error.asp?id=056"
next
if len(touxian)>4 then Response.Redirect "../error.asp?id=507"
if len(oicq)>11 or len(qianming)>50 or len(email)>20 then Response.Redirect "../error.asp?id=527"
if len(email)=0 or len(oicq)=0 then Response.Redirect "../error.asp?id=508"
if instr(email,".")=0 or instr(email,"@")=0 or instr(email,"'")<>0 or instr(email,chr(34))<>0 or instr(email,"\")<>0 or InStr(email,",")<>0 or InStr(email,"<")<>0 or InStr(email,">")<>0 or InStr(email,"#&")<>0 or InStr(email,"&")<>0 or InStr(email,"?")<>0 then Response.Redirect "../error.asp?id=523"
if instr(qianming,"file:")<>0 or instr(qianming,"script:")<>0 or instr(qianming,"'")<>0 or instr(qianming,chr(34))<>0 or instr(qianming,"\")<>0 or InStr(qianming,",")<>0 or InStr(qianming,"<")<>0 or InStr(qianming,">")<>0 or InStr(qianming,"#&")<>0 or InStr(qianming,"&")<>0 or InStr(qianming,"?")<>0 then Response.Redirect "../error.asp?id=524"
if instr(touxian,"file:")<>0 or instr(touxian,"script:")<>0 or instr(touxian,"'")<>0 or instr(touxian,chr(34))<>0 or instr(touxian,"\")<>0  or InStr(touxian,",")<>0 or InStr(touxian,"<")<>0 or InStr(touxian,">")<>0  or InStr(touxian,"#&")<>0 or InStr(touxian,"&")<>0 or InStr(touxian,"?")<>0  then Response.Redirect "../error.asp?id=524"
if instr(touxiang,"file:")<>0 or instr(touxiang,"script:")<>0 or instr(touxiang,"'")<>0 or instr(touxiang,chr(34))<>0 or instr(touxiang,"\")<>0  or InStr(touxiang,",")<>0 or InStr(touxiang,"<")<>0 or InStr(touxiang,">")<>0  or InStr(touxiang,"#&")<>0 or InStr(touxiang,"&")<>0 or InStr(touxiang,"?")<>0  then Response.Redirect "../error.asp?id=524"
for i=1 to len(touxian)
touxianchr=mid(touxian,i,1)
if asc(touxianchr)>0 then Response.Redirect "../error.asp?id=003"
next
for i=1 to len(oicq)
	oicqchr=asc(mid(oicq,i,1))
	if oicqchr<48 or (oicqchr>57 and oicqchr<65) or (oicqchr>90 and oicqchr<97) then Response.Redirect "../error.asp?id=006"
next	
set conn=server.createobject("adodb.connection")
conn.open Application("yx8_mhjh_connstr")
set rst=server.CreateObject ("adodb.recordset")
conn.BeginTrans
sql="update 用户 set 电子邮箱='"&email&"',oicq='"&oicq&"',头像='"&touxiang&"',头衔='"&touxian&"',签名档='"&qianming&"' where 姓名='"&Name&"'"
conn.Execute(sql)
if conn.Errors.Count=0 then
conn.CommitTrans
else
conn.RollbackTrans
if conn.Errors(0).Number =-2147467259 then Response.Redirect "../error.asp?id=009"
Response.Redirect "../error.asp?id=104&errormsg="&conn.Errors(0).Description
end if
conn.Close
set conn=nothing
%>
<html>
<title>资料修改</title>
<link rel=stylesheet href='../style.css'>
<body oncontextmenu=self.event.returnValue=false background="../chatroom/bg1.gif">
<table border=1 bgcolor="#BEE0FC" align=center width=336 cellpadding="10" cellspacing="13" height="293">
<tr valign="top">
<td bgcolor=#cccccc height="226">
<div align="center">
<p>　</p>
<p><font color="#000000" size="3">恭喜您成功修改资料!</font><font color="#FF3333" size="3"><br>
</font></p>
</div>
<table width=100%>
<tr>
<td>
<p align=center style='font-size:14;color:red'><font color="#000000"><br>
</font><font color="#000000">你的修改必须要重新登陆才能生效，谢谢！</font></p>
<p align=center>
<a href="REGMODI.ASP"><font color="#0000FF">返回</font></a>
</p>
</td>
</tr>
</table>
</td>
</tr>
</table>
</html>
<%
end if
%>

