<!--#include file="conn.asp"-->
<!-- #include file="inc/const.asp" -->
<head>
<title><%=Forum_info(0)%>--����ҳ��</title>
<link rel="stylesheet" href="forum_admin.css" type="text/css">
<base target="top">
</head>
<BODY leftmargin="0" bottommargin="0" rightmargin="0" topmargin="0" marginheight="0" marginwidth="0">
<%if not master or session("flag")="" then
	Errmsg=Errmsg+"<br>"+"<li>��ҳ��Ϊ����Աר�ã���<a href=admin_index.asp target=_top>��¼</a>����롣<br><li>��û�й���ҳ���Ȩ�ޡ�"
	call dvbbs_error()
else%>
	<br>
	&nbsp;��<a href="z_admin_down_softadd.asp" target="footer" title="��׼�����ϴ��ļ�"><b>����ϴ����</b></a>&nbsp;
	&nbsp;��<a href="z_admin_down_freeadd.asp" target="footer" title="ֱ��������ص�ַ����Ϣ"><b>����������</b></a>&nbsp; 
	&nbsp;��<a href="z_admin_down_adminedit.asp" target="footer"><b>�޸�ɾ��</b></a>&nbsp;  
	&nbsp;��<a href="z_admin_down_class.asp" target="footer"><b>��Ŀ����</b></a>&nbsp; 
	&nbsp;��<a href="z_admin_down_unite.asp" target="footer"><b>��������</b></a>&nbsp; 
	&nbsp;��<a href="z_admin_down_adminuser.asp" target="footer"><b>����Ա</b></a>&nbsp;
	&nbsp;��<a href="z_admin_down_adminevent.asp" target="footer"><b>�鿴�¼�</b></a>&nbsp;
	&nbsp;��<a href="z_admin_down_paymanage.asp" target="footer"><b>���ѹ���</b></a>
<%end if%>
</body>
</html>