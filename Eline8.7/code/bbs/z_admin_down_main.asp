<!--#include file="conn.asp"-->
<!-- #include file="inc/const.asp" -->
<head>
<title><%=Forum_info(0)%>--����ҳ��</title>
<link rel="stylesheet" href="forum_admin.css" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<base target="footer">
</head>
<BODY leftmargin="0" bottommargin="0" rightmargin="0" topmargin="0" marginheight="0" marginwidth="0">
<%if not master or session("flag")="" then
	Errmsg=Errmsg+"<br>"+"<li>��ҳ��Ϊ����Աר�ã���<a href=admin_index.asp target=_top>��¼</a>����롣<br><li>��û�й���ҳ���Ȩ�ޡ�"
	call dvbbs_error()
else%>
<table border="0" cellspacing="1" width="80%" align=center>
	<tr>
		<td width="100%"><br><p align=center><b>�������Ĺ���ҳ��</b></p><b>����Ա�ɽ��в���˵��</b>��<br>
			<br>1��ͨ��Web�ϴ��������������ѡ������ϴ�����Լ�׼�����ϴ��ļ���<br><br>
			<br>2���Ѿ��ϴ��ļ�������������ѡ������������������ӡ�<br>
			<br>3�����Ѿ��������޸Ļ�ɾ����������������ӽ��в�����<br>
			<br>4���������Ŀ��������޸�ɾ����������������ӽ��в�����<br>
			<br>5���������Է�Ϊ��<font color="#800080">��������</font>�����зÿͼ���Ա���������أ���<font color="#800080">��Ա����</font>��������̳��Ա���أ���<font color="#800080">VIP����</font>��������̳VIP��Ա���أ���<font color="#800080">��Լ����</font>�����ް����͹�������������͹���Ա���أ�<br>
			<br>6������ϵͳ����Ա˵����<font color="#800080">����̳����Ա</font>�����Թ���������Ŀ����<font color="#800080">��������Ա</font>�����Խ����ϴ�����������������ɾ�����޸ģ���<font color="#800080">�������Ա</font>�����Խ����ϴ��������������<font color="#800080">��ͨ�û�</font>�����������<br>
		</td>
	</tr>
</table>
<%end if%>