<!--#include file=conn.asp-->
<!-- #include file="inc/const.asp" -->
<head>
<title><%=Forum_info(0)%>--����ҳ��</title>
<link rel="stylesheet" href="forum_admin.css" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>

<BODY leftmargin="0" bottommargin="0" rightmargin="0" topmargin="0" marginheight="0" marginwidth="0" bgcolor="#DDEEFF">
<%
	if not master or session("flag")="" then
		Errmsg=Errmsg+"<br>"+"<li>��ҳ��Ϊ����Աר�ã���<a href=admin_index.asp target=_top>��¼</a>����롣<br><li>��û�й���ҳ���Ȩ�ޡ�"
		call dvbbs_error()
	else
		call main()
	end if
sub main()
%>
<table width="100%" border="0" cellspacing="1" cellpadding="3"  align=center >
<tr>
<th align=center colspan=3   width="717" height=1>**��̳���ݿ����**
</th>
<tr width="95%">
<td width="185" height="62"></td>
<td width="399" height="62"><b>SOL���ִ�в���</b>�����������޸߼�����SQL��̱Ƚ���Ϥ���û���������ֱ������SQLִ����䣬����DELETE FROM bbs1 WHERE username='test'����ɾ��ĳ�û����Ӳ������ڲ���ǰ�����ؿ�������ִ������Ƿ���ȷ��������ִ�к󲻿ɻָ���</td>
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
<tr><th align=center colspan=3  width="717" height=1><b>ִ�н��</b></th>  
</tr>
<tr><td align=center colspan=3  width="717" height=1><%response.write "ִ�гɹ�"%>
</td>                                                                                 
</tr>
<%
else%>
<tr><th align=center colspan=3  width="717" height=1><b>ִ�н��</b></th>  
</tr>
<tr><td align=center colspan=3  width="717" height=1><%response.write "��������⣬����������£�<br>"
response.write Err.Description%>
</td>                                                                                 
</tr>
<%
err.clear
end if
end if
else
%>
<tr><th align=center colspan=3  width="717" height=1><b></b>������SQL���</th>  
</tr>
<tr><td align=center colspan=3  width="717" height=1><Form Name=FormPst Method=Post Action="z_admin_connsql.asp?action=save">
<FieldSet>
<Legend>������SQL���</Legend>
ָ�<Input type="text" name="SQL_Statement" size=80> <p>
<Input type="Submit" Value="�ͳ�"> <p>
</FieldSet>
</Form></td>                                                                                 
</tr>
<%end if%>
<%
	end sub
%>