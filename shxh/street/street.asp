<%
bgcolor=Application("Ba_jxqy_backgroundcolor")
bgimage=Application("Ba_jxqy_backgroundimage")
set conn=server.CreateObject("adodb.connection")
conn.open Application("BA_jxqy_connstr")
conn.execute "update ϵͳ���� set ����ֵ='"&Application("Ba_jxqy_visitor")&"' where ����='��������'"
conn.close
set conn=nothing
%>
<html>
<head>
<title><%=Application("Ba_jxqy_systemname")%></title>
<link rel=stylesheet href=../chatroom/css.css>
</head>
<body  oncontextmenu=self.event.returnValue=false bgcolor='<%=bgcolor%>' background='<%=bgimage%>'>
<p align=center><font color=0000ff size=5>ʮ�ﳤ��</font></p><hr>
<table  bgcolor="#FFE4CA" width=30% align=center border=1 bordercolorlight=000000 cellspacing=0 cellpadding=1 bordercolordark=FFFFFF >
<tr align=center><td ><a href='../other/factory.asp' onmouseover="window.status='���Լ��ĺ�,���Լ��ķ�,���쿿�ؿ�����,�����Ǻú�';return true;" onmouseout="window.status='';return true;" title="���Լ��ĺ�,���Լ��ķ�,���쿿�ؿ�����,�����Ǻú�">�� �� ��</a></td></tr>
<tr align=center><td bgcolor="#FFFFFF"><a href='bank.asp' target=rulfrm onmouseover="window.status='�������ǮȡǮ';return true;" onmouseout="window.status='';return true;" title="�������ǮȡǮ">Ǯ&nbsp;&nbsp;&nbsp;&nbsp;ׯ</a></td></tr>
<tr align=center><td ><a href='armorshop.asp' target=rulfrm onmouseover="window.status='���������ͷ��,����,����,����,��Ʒ,����';return true;" onmouseout="window.status='';return true;" title="���������ͷ��,����,����,����,��Ʒ,����">�� �� ��</a></td></tr>
<tr align=center><td bgcolor="#FFFFFF"><a href='curativeshop.asp' target=rulfrm onmouseover="window.status='���������ҩƷ';return true;" onmouseout="window.status='';return true;" title="���������ҩƷ">ɽ ҩ ��</a></td></tr>
<tr align=center><td><a href='market.asp' target=rulfrm onmouseover="window.status='������ʹ�����������Լ����õ���Ʒ';return true;" onmouseout="window.status='';return true;" title="������ʹ�����������Լ����õ���Ʒ">���ɼ���</a></td></tr>
<tr align=center><td bgcolor="#FFFFFF"><a href='pawnshop.asp' target=rulfrm onmouseover="window.status='Ҫ����?���õ���,����Ҳ������ʱ!';return true;" onmouseout="window.status='';return true;" title="Ҫ����?���õ���,����Ҳ������ʱ!">��&nbsp;&nbsp;&nbsp;&nbsp;��</a></td></tr>
<tr align=center><td ><a href='hotel.asp' target=rulfrm onmouseover="window.status='����,��Ҫ˯����';return true;" onmouseout="window.status='';return true;" title="����,��Ҫ˯����">������ջ</a></td></tr>
<tr align=center><td bgcolor="#FFFFFF"><a href='cardshop.asp' target=rulfrm onmouseover="window.status='����е��߿���';return true;" onmouseout="window.status='';return true;" title="���ߵ�">�� �� ��</a></td></tr>
</table>
</body>
</html>