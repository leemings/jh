<!-- #include file="conn.asp" -->
<!-- #include file="inc/const.asp" -->
<title><%=Forum_info(0)%>--����ҳ��</title>
<link rel="stylesheet" href="forum_admin.css" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<BODY leftmargin="0" bottommargin="0" rightmargin="0" topmargin="0" marginheight="0" marginwidth="0" bgcolor="#DDEEFF">
<%
dim admin_flag
admin_flag="04"
if not master or instr(session("flag"),admin_flag)=0 then
	Errmsg=Errmsg+"<br>"+"<li>��ҳ��Ϊ����Աר�ã���<a href=admin_index.asp target=_top>��¼</a>����롣<br><li>��û�й���ҳ���Ȩ�ޡ�"
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
<th width="100%" class="tableHeaderText" colspan=2 height=25>��̳��ʼ��Ϣ����
</th>
</tr>
<tr>
<td class="forumrow" colspan=2>
<p><B>ע��</B>�������������Ե���������̳��Ϣ��ʾ�����в��ֹ��ܿ�����ҪFSO��֧�֣������֧��FSO�ļ�д�빦�ܣ���ֱ���ֶ��޸�����ļ���
</td>
</tr>
<tr>
<td class="forumrow">
<B>����ѡ��</B></td>
<td class="forumrow"><a href="admin_inputfile.asp">ע�������޸�</a> | <a href="?action=email">ע���ʼ������޸�</a> | <a href="?action=sms">ע�ᷢ�Ͷ����޸�</a>
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
response.write "���Ŀռ䲻֧��FSO����ͬ���Ŀռ�����ϵ�����߲鿴���Ȩ������<br>"&Err.Description&""
response.end
end if
%> 
<form method=post action="?action=reg">
<table width="95%" border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
<tr> 
<th width="100%" class="tableHeaderText" height=25>ע�������޸�
</th>
</tr>
<tr>
<td class="forumrow">
ע�⣺�ļ������������㰲װĿ¼�µ�<B>inc/reg_txt.asp</B>�������Ŀռ䲻֧��<font color="red">FSO</font>����ֱ�ӱ༭���ļ��������д��С�%%���Ĳ��ֲ�Ҫ���б༭
</td>
</tr>
<tr>
<td class="forumrow">
<textarea name="fdata" cols="100" rows="15"><%=fdata%></textarea>
</td>
</tr>
<tr>
<td class="forumrow">
<input type="reset" name="Reset" value="�����޸�">
<input type="submit" name="save" value="�ύ�޸�"> ��ǰ�ļ�·����<b><%=Server.MapPath("inc/reg_txt.asp")%></b>
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
response.write "���Ŀռ䲻֧��FSO����ͬ���Ŀռ�����ϵ�����߲鿴���Ȩ������<br>"&Err.Description&""
response.end
end if
%> 
<form method=post action="?action=email">
<table width="95%" border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
<tr> 
<th width="100%" class="tableHeaderText" height=25>ע�������޸�
</th>
</tr>
<tr>
<td class="forumrow">
ע�⣺�ļ������������㰲װĿ¼�µ�<B>inc/email_txt.asp</B>�������Ŀռ䲻֧��<font color="red">FSO</font>����ֱ�ӱ༭���ļ���<BR>
<B>�޸�˵��</B>��<BR>
���޸ı��ļ���Ҫһ��HTML��ASP����֪ʶ����������ᣬ�벻Ҫ����Ķ�<BR>
���ļ��д��С�%%����"&&"֮����ַ���Ҫ�������Ķ�<BR>
��mailbody�������ʼ����ݱ�����htmlencode(username)���û�����getpass�����룬Copyright�ǰ�Ȩ��Version�汾��Ϣ<BR>
����Ҫ�����к����ݿ��Իس���������ַ�ʽ��mailbody=mailbody & "�������"<BR>
��HTML�е�˫���Ž���ĳɵ����ţ�Ҳ���Բ�Ҫ���Ż��߼�����˫���Ŷ�����
</td>
</tr>
<tr>
<td class="forumrow">
<textarea name="fdata" cols="100" rows="15"><%=fdata%></textarea>
</td>
</tr>
<tr>
<td class="forumrow">
<input type="reset" name="Reset" value="�����޸�">
<input type="submit" name="save" value="�ύ�޸�"> ��ǰ�ļ�·����<b><%=Server.MapPath("inc/email_txt.asp")%></b>
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
response.write "���Ŀռ䲻֧��FSO����ͬ���Ŀռ�����ϵ�����߲鿴���Ȩ������<br>"&Err.Description&""
response.end
end if
%> 
<form method=post action="?action=sms">
<table width="95%" border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
<tr> 
<th width="100%" class="tableHeaderText" height=25>ע�������޸�
</th>
</tr>
<tr>
<td class="forumrow">
ע�⣺�ļ������������㰲װĿ¼�µ�<B>inc/email_txt.asp</B>�������Ŀռ䲻֧��<font color="red">FSO</font>����ֱ�ӱ༭���ļ���<BR>
<B>�޸�˵��</B>��<BR>
���޸ı��ļ���Ҫһ��HTML��ASP����֪ʶ����������ᣬ�벻Ҫ����Ķ�<BR>
���ļ��д��С�%%����"&&"֮����ַ���Ҫ�������Ķ�<BR>
��body�����Ƕ������ݱ�����Forum_info(0)����̳���Ʊ�����chr(10)��<B>����</B>����<BR>
��<B>�����ûس�</B>������������ֱ����ӻ���ʹ�û��з��ź����<BR>
��<B>������HTML����</B>
</td>
</tr>
<tr>
<td class="forumrow">
<textarea name="fdata" cols="100" rows="15"><%=fdata%></textarea>
</td>
</tr>
<tr>
<td class="forumrow">
<input type="reset" name="Reset" value="�����޸�">
<input type="submit" name="save" value="�ύ�޸�"> ��ǰ�ļ�·����<b><%=Server.MapPath("inc/sms_txt.asp")%></b>
</td>
</tr>
</form>
</table>
<%
end sub
%>