<!-- #include file="conn.asp" -->
<!-- #include file="inc/const.asp" -->
<title><%=Forum_info(0)%>--管理页面</title>
<link rel="stylesheet" href="forum_admin.css" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<BODY leftmargin="0" bottommargin="0" rightmargin="0" topmargin="0" marginheight="0" marginwidth="0" bgcolor="#DDEEFF">
<%
dim admin_flag
admin_flag="04"
if not master or instr(session("flag"),admin_flag)=0 then
	Errmsg=Errmsg+"<br>"+"<li>本页面为管理员专用，请<a href=admin_index.asp target=_top>登录</a>后进入。<br><li>您没有管理本页面的权限。"
	call dvbbs_error()
else
	call main()
	conn.close
	set conn=nothing
end if

sub main()
%>
<table width="95%" border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
<tr> 
<th width="100%" class="tableHeaderText" colspan=2 height=25>论坛初始信息设置
</th>
</tr>
<tr>
<td class="forumrow" colspan=2>
<p><B>注意</B>：在这里您可以调整部分论坛信息提示，其中部分功能可能需要FSO的支持，如果不支持FSO文件写入功能，请直接手动修改相关文件。
</td>
</tr>
<tr>
<td class="forumrow">
<B>操作选项</B></td>
<td class="forumrow"><a href="admin_inputfile.asp">注册条款修改</a> | <a href="?action=email">注册邮件内容修改</a> | <a href="?action=sms">注册发送短信修改</a>
</td>
</tr>
</table>
<p></p>
<%
select case request("action")
case "email"
	call email()
case "sms"
	call sms()
case else
	call reg()
end select
end sub

sub reg()
dim objFSO
dim fdata
dim objCountFile
on error resume next
Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
if request("save")="" then
	Set objCountFile = objFSO.OpenTextFile(Server.MapPath("inc/reg_txt.asp"),1,True)
	If Not objCountFile.AtEndOfStream Then fdata = objCountFile.ReadAll
else
	fdata=request("fdata")
	Set objCountFile=objFSO.CreateTextFile(Server.MapPath("inc/reg_txt.asp"),True)
	objCountFile.Write fdata
end if
objCountFile.Close
Set objCountFile=Nothing
Set objFSO = Nothing
if err.number<>0 then
response.write "您的空间不支持FSO，请同您的空间商联系，或者查看相关权限设置<br>"&Err.Description&""
response.end
end if
%> 
<form method=post action="?action=reg">
<table width="95%" border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
<tr> 
<th width="100%" class="tableHeaderText" height=25>注册条款修改
</th>
</tr>
<tr>
<td class="forumrow">
注意：文件将被保存在你安装目录下的<B>inc/reg_txt.asp</B>里，如果您的空间不支持<font color="red">FSO</font>，请直接编辑该文件！内容中带有《%%》的部分不要进行编辑
</td>
</tr>
<tr>
<td class="forumrow">
<textarea name="fdata" cols="100" rows="15"><%=fdata%></textarea>
</td>
</tr>
<tr>
<td class="forumrow">
<input type="reset" name="Reset" value="重新修改">
<input type="submit" name="save" value="提交修改"> 当前文件路径：<b><%=Server.MapPath("inc/reg_txt.asp")%></b>
</td>
</tr>
</form>
</table>
<%
end sub

sub email()
dim objFSO
dim fdata
dim objCountFile
on error resume next
Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
if request("save")="" then
	Set objCountFile = objFSO.OpenTextFile(Server.MapPath("inc/email_txt.asp"),1,True)
	If Not objCountFile.AtEndOfStream Then fdata = objCountFile.ReadAll
else
	fdata=request("fdata")
	Set objCountFile=objFSO.CreateTextFile(Server.MapPath("inc/email_txt.asp"),True)
	objCountFile.Write fdata
end if
objCountFile.Close
Set objCountFile=Nothing
Set objFSO = Nothing
if err.number<>0 then
response.write "您的空间不支持FSO，请同您的空间商联系，或者查看相关权限设置<br>"&Err.Description&""
response.end
end if
%> 
<form method=post action="?action=email">
<table width="95%" border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
<tr> 
<th width="100%" class="tableHeaderText" height=25>注册条款修改
</th>
</tr>
<tr>
<td class="forumrow">
注意：文件将被保存在你安装目录下的<B>inc/email_txt.asp</B>里，如果您的空间不支持<font color="red">FSO</font>，请直接编辑该文件！<BR>
<B>修改说明</B>：<BR>
①修改本文件需要一定HTML和ASP代码知识，如果您不会，请不要随意改动<BR>
②文件中带有《%%》和"&&"之类的字符不要随意做改动<BR>
③mailbody变量是邮件内容变量，htmlencode(username)是用户名，getpass是密码，Copyright是版权，Version版本信息<BR>
④需要增加行和内容可以回车后采用这种方式：mailbody=mailbody & "添加内容"<BR>
⑤HTML中的双引号建议改成单引号，也可以不要引号或者加两个双引号都可以
</td>
</tr>
<tr>
<td class="forumrow">
<textarea name="fdata" cols="100" rows="15"><%=fdata%></textarea>
</td>
</tr>
<tr>
<td class="forumrow">
<input type="reset" name="Reset" value="重新修改">
<input type="submit" name="save" value="提交修改"> 当前文件路径：<b><%=Server.MapPath("inc/email_txt.asp")%></b>
</td>
</tr>
</form>
</table>
<%
end sub

sub sms()
dim objFSO
dim fdata
dim objCountFile
on error resume next
Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
if request("save")="" then
	Set objCountFile = objFSO.OpenTextFile(Server.MapPath("inc/sms_txt.asp"),1,True)
	If Not objCountFile.AtEndOfStream Then fdata = objCountFile.ReadAll
else
	fdata=request("fdata")
	Set objCountFile=objFSO.CreateTextFile(Server.MapPath("inc/sms_txt.asp"),True)
	objCountFile.Write fdata
end if
objCountFile.Close
Set objCountFile=Nothing
Set objFSO = Nothing
if err.number<>0 then
response.write "您的空间不支持FSO，请同您的空间商联系，或者查看相关权限设置<br>"&Err.Description&""
response.end
end if
%> 
<form method=post action="?action=sms">
<table width="95%" border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
<tr> 
<th width="100%" class="tableHeaderText" height=25>注册条款修改
</th>
</tr>
<tr>
<td class="forumrow">
注意：文件将被保存在你安装目录下的<B>inc/email_txt.asp</B>里，如果您的空间不支持<font color="red">FSO</font>，请直接编辑该文件！<BR>
<B>修改说明</B>：<BR>
①修改本文件需要一定HTML和ASP代码知识，如果您不会，请不要随意改动<BR>
②文件中带有《%%》和"&&"之类的字符不要随意做改动<BR>
③body变量是短信内容变量，Forum_info(0)是论坛名称变量，chr(10)是<B>换行</B>符号<BR>
④<B>不可用回车</B>，增加内容请直接添加或者使用换行符号后添加<BR>
⑤<B>不可用HTML代码</B>
</td>
</tr>
<tr>
<td class="forumrow">
<textarea name="fdata" cols="100" rows="15"><%=fdata%></textarea>
</td>
</tr>
<tr>
<td class="forumrow">
<input type="reset" name="Reset" value="重新修改">
<input type="submit" name="save" value="提交修改"> 当前文件路径：<b><%=Server.MapPath("inc/sms_txt.asp")%></b>
</td>
</tr>
</form>
</table>
<%
end sub
%>