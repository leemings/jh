<%
bgcolor=Application("Ba_jxqy_backgroundcolor")
bgimage=Application("Ba_jxqy_backgroundimage")
%>
<html>
<head>
<title><%=Application("Ba_jxqy_systemname")%></title>
<link rel=stylesheet href=chatroom/css.css>
</head>
<body  oncontextmenu=self.event.returnValue=false bgcolor='<%=bgcolor%>' background='<%=bgimage%>'>
<table  bgcolor="#FFE4CA" width=75% align=center border=1 bordercolorlight=000000 cellspacing=0 cellpadding=1 bordercolordark=FFFFFF >
<tr align=center><td><a href='#' onclick="javascript:mainwin=window.open('main.asp','main','left=0,top=0,status=yes,scrollbars=yes,resizable=yes');mainwin.resizeTo(screen.availWidth,screen.availHeight);" onmouseover="window.status='ʹ�õ�ͼ����ϵͳ';return true;" onmouseout="window.status='';return true;" title="ʹ�õ�ͼ����ϵͳ">��ͼ����</a></td></tr>
<tr align=center><td bgcolor="#FFFFFF"><a href='myhome/myhome.asp' target=rulfrm onmouseover="window.status='�ҵİ��ҵ����ҵļ�';return true;" onmouseout="window.status='';return true;" title="�ҵİ��ҵ����ҵļ�">�� �� ��</a></td></tr>
<tr align=center><td><a href='chatroom/announce.htm' target=rulfrm onmouseover="window.status='�ۿ��ٸ���ʱ����';return true;" onmouseout="window.status='';return true;" title="�ۿ��ٸ���ʱ����">��Ա��֪</a></td></tr>
<tr align=center><td bgcolor="#FFFFFF" ><a href='#' onclick="javascript:admwin=window.open('admin/administrator.asp','admin','left=0,top=0,status=yes,scrollbars=yes,resizable=yes');admwin.resizeTo(screen.availWidth,screen.availHeight);" onmouseover="window.status='����ٸ�';return true;" onmouseout="window.status='';return true;" title="����ٸ�">��&nbsp;&nbsp;&nbsp;&nbsp;��</a></td></tr>
<tr align=center><td><a href='#' onclick="javascript:chatwin=window.open('chatroom/chatroom.asp','chatroom','left=0,top=0,status=yes,scrollbars=yes,resizable=yes');chatwin.resizeTo(screen.availWidth,screen.availHeight);" onmouseover="window.status='����������';return true;" onmouseout="window.status='';return true;" title="����������">�� �� ��</a></td></tr>
<tr align=center><td bgcolor="#FFFFFF"><a href='#' onclick="javascript:bbswin=window.open('bbs/index.asp','bbs','left=0,top=0,status=yes,scrollbars=yes,resizable=yes');bbswin.resizeTo(screen.availWidth,screen.availHeight);" onmouseover="window.status='�������ɽׯ';return true;" onmouseout="window.status='';return true;" title="�������ɽׯ">����ɽׯ</a></td></tr>
<tr align=center><td ><a href='#' onclick="javascript:advwin=window.open('adventure/adventure.asp','adventure','left=100,top=50,status=yes,scrollbars=yes,resizable=yes');advwin.resizeTo(screen.availWidth*2/3,screen.availHeight*3/4);" onmouseover="window.status='��ʼ���ð������';return true;" onmouseout="window.status='';return true;" title="��ʼ���ð������">̽ �� ��</a></td></tr>
<tr align=center><td bgcolor="#FFFFFF"><a href='street/street.asp' target=rulfrm onmouseover="window.status='�����Ƽ����˽��ദ';return true;" onmouseout="window.status='';return true;" title="�����������Ʒ">ʮ�ﳤ��</a></td></tr>
<tr align=center><td ><a href='other/dubo.asp' target=rulfrm onmouseover="window.status='ʮ�ľ�թ,С��Ϊ��';return true;" onmouseout="window.status='';return true;" title="ʮ�ľ�թ,С��Ϊ��">��&nbsp;&nbsp;&nbsp;&nbsp;��</a></td></tr>
<tr align=center><td bgcolor="#FFFFFF"><a href='#' onclick="javascript:stowin=window.open('stock/stock.htm','stock','left=0,top=0,status=yes,scrollbars=yes,resizable=yes');stowin.resizeTo(screen.availWidth,screen.availHeight);" onmouseover="window.status='����֤ȯ�����ĳ���';return true;" onmouseout="window.status='';return true;" title="����֤ȯ�����ĳ���">֤ �� ��</a></td></tr>
<tr align=center><td><a href='street/hotel.asp' target=rulfrm onmouseover="window.status='����,��Ҫ˯����';return true;" onmouseout="window.status='';return true;" title="����,��Ҫ˯����">������ջ</a></td></tr>
<tr align=center><td bgcolor="#FFFFFF"><a href='other/top.asp' target=rulfrm onmouseover="window.status='��Ϊ��������,����ԴȪ';return true;" onmouseout="window.status='';return true;" title="��Ϊ��������,����ԴȪ">�� �� ��</a></td></tr>
<tr align=center><td><a href='street/herostele.asp' target=rulfrm onmouseover="window.status='�ù�����������ӥ����ãԶȥ';return true;" onmouseout="window.status='';return true;" title="�ù�����������ӥ����ãԶȥ">Ӣ �� ��</a></td></tr>
<tr align=center><td bgcolor="#FFFFFF"><a href='suicide.asp' target=rulfrm onmouseover="window.status='�ҵ������������һ��';return true;" onmouseout="window.status='';return true;" title="�ҵ������������һ��">�� �� ��</a></td></tr>
<tr align=center><td><a href='#' onclick="javascript:chatwin=window.open('diaoyu/diaoyu.asp','chatroom','left=0,top=0,status=yes,scrollbars=yes,resizable=yes');chatwin.resizeTo(screen.availWidth,screen.availHeight);" onmouseover="window.status='����������';return true;" onmouseout="window.status='';return true;" title="����������">�� 
    �� ̨</a></td></tr>
<tr align=center><td bgcolor="#FFFFFF"><a href='#' onclick="javascript:chatwin=window.open('Qmg/Qmg.htm','chatroom','left=0,top=0,status=yes,scrollbars=yes,resizable=yes');chatwin.resizeTo(screen.availWidth,screen.availHeight);" onmouseover="window.status='����������';return true;" onmouseout="window.status='';return true;" title="����������">�� 
    �� ��</a></td></tr>
<tr align=center><td bgcolor="#FFE4CA"><a href='#' onclick="javascript:chatwin=window.open('jiudian/jd.asp','chatroom','left=0,top=0,status=yes,scrollbars=yes,resizable=yes');chatwin.resizeTo(screen.availWidth,screen.availHeight);" onmouseover="window.status='����������';return true;" onmouseout="window.status='';return true;" title="����������">�� 
    �� ��</a></td></tr>
<tr align=center><td bgcolor="#FFFFFF"><a href='#' onclick="javascript:chatwin=window.open('hktk/xuetang.asp','chatroom','left=0,top=0,status=yes,scrollbars=yes,resizable=yes');chatwin.resizeTo(screen.availWidth,screen.availHeight);" onmouseover="window.status='����������';return true;" onmouseout="window.status='';return true;" title="����������">ѵ 
    �� ��</a></td></tr>
<tr align=center><td bgcolor="#FFE4CA"><a href='#' onclick="javascript:chatwin=window.open('myjinmao/index.asp','chatroom','left=0,top=0,status=yes,scrollbars=yes,resizable=yes');chatwin.resizeTo(screen.availWidth,screen.availHeight);" onmouseover="window.status='����������';return true;" onmouseout="window.status='';return true;" title="����������">����α��</a></td></tr>
<tr align=center><td bgcolor="#FFFFFF"><a href='#' onclick="javascript:chatwin=window.open('yhy/index.asp','chatroom','left=0,top=0,status=yes,scrollbars=yes,resizable=yes');chatwin.resizeTo(screen.availWidth,screen.availHeight);" onmouseover="window.status='����������';return true;" onmouseout="window.status='';return true;" title="����������">������Ժ</a></td></tr>
<tr align=center><td bgcolor="#FFE4CA">�� �� ��</td></tr>
<tr align=center><td bgcolor="#FFFFFF">�� �� ��</td></tr>
<tr align=center><td bgcolor="#FFE4CA">�� �� ��</td></tr>
<tr align=center><td bgcolor="#FFFFFF">�� �� ��</td></tr>
<tr align=center><td bgcolor="#FFE4CA"><a href='welcome.asp' target=rulfrm onmouseover="window.status='�ط���ҳ';return true;" onmouseout="window.status='';return true;" title="�ط���ҳ">��&nbsp;&nbsp;&nbsp;&nbsp;��</a></td></tr>
<tr align=center><td bgcolor="#FFFFFF"><a href='exit.asp' target=rulfrm onmouseover="window.status='�˳������½���';return true;" onmouseout="window.status='';return true;" title="�˳������½���">��&nbsp;&nbsp;&nbsp;&nbsp;��</a></td></tr>
</table>
</body>
</html>