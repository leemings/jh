<!--#include file="conn.asp"-->
<!-- #include file="inc/const.asp" -->
<head>
<title><%=Forum_info(0)%>--����ҳ��</title>
<link rel="stylesheet" href="forum_admin.css" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>
<%if not master or session("flag")="" then
	Errmsg=Errmsg+"<br>"+"<li>��ҳ��Ϊ����Աר�ã���<a href=admin_index.asp target=_top>��¼</a>����롣<br><li>��û�й���ҳ���Ȩ�ޡ�"
	call dvbbs_error()
else%>
	<frameset rows="40,*" frameborder="yes" border="1" framespacing="2">
		<noframes>
			<BODY leftmargin="0" bottommargin="0" rightmargin="0" topmargin="0" marginheight="0" marginwidth="0">
			<p>����ҳʹ���˿�ܣ��������������֧�ֿ�ܡ�</p>
  		</body>
		</noframes>
		<frame name="top" target="footer" scrolling="no" marginwidth="0" marginheight="0" src="z_admin_down_top.asp" noresize>
		<frame name="footer" scrolling="auto" marginwidth="0" marginheight="0" src="z_admin_down_main.asp">
	</frameset>
<%end if%>