<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="../pass.asp"-->
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
myid=Request.form("id")
name=Request.form("name")
oldpass=Request.form("oldpass")
pass=request.form("pass")
repass=request.form("repass")
for each element in Request.Form
if instr(element,"'")<>0 or instr(element,"|")<>0 or instr(element," ")<>0 or instr(Request.Form(element),"'")<>0 or instr(Request.Form(element)," ")<>0 or instr(Request.Form(element),"|")<>0 then 
		Response.Write "<script Language=Javascript>alert('提示：输入数据有问题，请查看！');location.href = 'javascript:history.go(-1)';</script>"
		Response.End
end if
next
if not isnumeric(myid) then 
	Response.Write "<script language=JavaScript>{alert('提示：["&myid&"]输入错误，ID请使用数字！');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End 
end if
if name="" then
message="帐号不能为空"
elseif oldpass="" or pass="" or repass="" then
message="密码不能为空"
elseif pass<>repass then
message="密码与确认密码必须一致"
else
oldpass=md5(oldpass)
pass=md5(pass)
'ip=Request.ServerVariables("REMOTE_ADDR")
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
'校验用户
rs.open "SELECT 状态 FROM 用户 WHERE id="&myid&" and 姓名='" & name & "' and 密码='" & oldpass& "'",conn
If Rs.Bof OR Rs.Eof Then
message="对不起，你的密码不对"
else
	if  rs("状态")<>"正常" then
		message="你的状态不正常不可以改密码！"
	else
		conn.Execute "update 用户 set 密码='" & pass & "' where 姓名='" & name & "'"
		conn.execute "insert into l(a,b,c,d,e) values (now(),'"& name &"','无','操作','改换用户密码！')"
		message="恭喜您成功地修改了密码！"
	end if
end if
	conn.close
	set conn=nothing
end if
%>
<head><LINK href="../css.css" rel=stylesheet>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312"></head>
<body bgcolor="#006699">
<table width="360" border="1" align="center" cellpadding="10"
cellspacing="13" bgcolor="#000000">
  <tr bgcolor="#FFFFFF">
<td>
<table width="100%">
<tr>
<td>
<p align="center" style="font-size:14;color:red"><%=message%></p>
<p align="center"><a href="#" onclick="window.close()">[ 关 闭 窗 口 ]</a></p>
</td>
</tr>
</table>
</td>
</tr>
</table>

</body>