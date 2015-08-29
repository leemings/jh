<!--#include file=conn.asp-->
<!-- #include file="inc/const.asp" -->
<head>
<title><%=Forum_info(0)%>--管理页面</title>
<link rel="stylesheet" href="forum_admin.css" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>

<BODY leftmargin="0" bottommargin="0" rightmargin="0" topmargin="0" marginheight="0" marginwidth="0" bgcolor="#DDEEFF">
<%
	if not master or session("flag")="" then
		Errmsg=Errmsg+"<br>"+"<li>本页面为管理员专用，请<a href=admin_index.asp target=_top>登录</a>后进入。<br><li>您没有管理本页面的权限。"
		call dvbbs_error()
	else
		call main()
	end if
sub main()
%>
<table width="100%" border="0" cellspacing="1" cellpadding="3"  align=center >
<tr>
<th align=center colspan=3   width="717" height=1>**论坛数据库管理**
</th>
<tr width="95%">
<td width="185" height="62"></td>
<td width="399" height="62"><b>SOL语句执行操作</b>：本操作仅限高级、对SQL编程比较熟悉的用户，您可以直接输入SQL执行语句，比如DELETE FROM bbs1 WHERE username='test'进行删除某用户帖子操作，在操作前请慎重考虑您的执行语句是否正确和完整，执行后不可恢复。</td>
<td width="113" height="62"></td>
</tr>
                                                                               
<%
if request("action") = "save" then
dim SQL_Statement
SQL_Statement=Request("SQL_Statement")
if SQL_Statement<>Empty then
On Error Resume Next 
conn.Execute(SQL_Statement)
if err.number="0" then%>
<tr><th align=center colspan=3  width="717" height=1><b>执行结果</b></th>  
</tr>
<tr><td align=center colspan=3  width="717" height=1><%response.write "执行成功"%>
</td>                                                                                 
</tr>
<%
else%>
<tr><th align=center colspan=3  width="717" height=1><b>执行结果</b></th>  
</tr>
<tr><td align=center colspan=3  width="717" height=1><%response.write "语句有问题，具体出错如下：<br>"
response.write Err.Description%>
</td>                                                                                 
</tr>
<%
err.clear
end if
end if
else
%>
<tr><th align=center colspan=3  width="717" height=1><b></b>请输入SQL语句</th>  
</tr>
<tr><td align=center colspan=3  width="717" height=1><Form Name=FormPst Method=Post Action="z_admin_connsql.asp?action=save">
<FieldSet>
<Legend>请输入SQL语句</Legend>
指令：<Input type="text" name="SQL_Statement" size=80> <p>
<Input type="Submit" Value="送出"> <p>
</FieldSet>
</Form></td>                                                                                 
</tr>
<%end if%>
<%
	end sub
%>