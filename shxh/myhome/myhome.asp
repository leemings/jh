<%
username=Session("Ba_jxqy_username")
if username="" then Response.Redirect "../error.asp?id=016"
bgcolor=Application("Ba_jxqy_backgroundcolor")
bgimage=Application("Ba_jxqy_backgroundimage")
%>
<html>
<head>
<title>�ҵļ�</title>
<link rel=stylesheet href='../chatroom/css.css'>
</head>
<body bgcolor='<%=bgcolor%>' background='<%=bgimage%>' oncontextmenu=self.event.returnValue=false>
<p align=center><font color=0000ff size=5>�ҵİ��ҵ����ҵļ�</font><br>�����ĵĺ���,���ǰ��Ľ���</p><hr>
<table width=60% border=0 align=center>
<tr><td><a href='seeabout.asp' onmouseover="window.status='�鿴�ҵ�״̬';return true;" onmouseout="window.status='';return true;" title="�鿴�ҵ�״̬">״&nbsp;&nbsp;&nbsp;&nbsp;̬</a></td></tr>
<tr><td><a href='pet.asp' onmouseover="window.status='���������ĵĺ���,���ǰ��밮�Ľ���';return true;" onmouseout="window.status='';return true;" title="���������ĵĺ���,���ǰ��밮�Ľ���">����֮��</a></td></tr>
<tr><td><a href='../bbs/getpm.asp' target=rulfrm onmouseover="window.status='�����Ƿ������Ѹ��Լ�����';return true;" onmouseout="window.status='';return true;" title="�����Ƿ������Ѹ��Լ�����">�ҵ�����</a></td></tr>
<tr><td><a href='seeequip.asp' onmouseover="window.status='�ҵ�����װ��';return true;" onmouseout="window.status='';return true;" title="�ҵ�����װ��">��&nbsp;&nbsp;&nbsp;&nbsp;��</a></td></tr>
<tr><td><a href='seecurative.asp' onmouseover="window.status='�ҵ�ҩƷ';return true;" onmouseout="window.status='';return true;" title="�ҵ�ҩƷ">ҩ&nbsp;&nbsp;&nbsp;&nbsp;Ʒ</a></td></tr>
<tr><td><a href='../ballot/ballot.asp' onmouseover="window.status='�μ�����ͶƱ';return true;" onmouseout="window.status='';return true;" title="�μ�����ͶƱ">Ͷ&nbsp;&nbsp;&nbsp;&nbsp;Ʊ</a></td></tr>
<tr><td><a href='../other/marriage.asp' onmouseover="window.status='�μ���������';return true;" onmouseout="window.status='';return true;" title="�μ���������">��&nbsp;&nbsp;&nbsp;&nbsp;��</a></td></tr>
</table>
</body>
</html>