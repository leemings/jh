<%
chatroomsn=Session("Ba_jxqy_userchatroomsn")
bgcolor=Application("Ba_jxqy_backgroundcolor")
bgimage=Application("Ba_jxqy_backgroundimage")
mycorp=session("Ba_jxqy_usercorp")
username=session("Ba_jxqy_username")
Set conn=server.CreateObject("adodb.connection")
conn.Open Application("Ba_jxqy_connstr")
set rst=server.CreateObject("adodb.recordset")
sqlstr="select ��ʽ,�������� from ���� where ����='"&username&"' order by ��������"
rst.Open sqlstr,conn
msgarr=";"
do while not (rst.EOF or rst.BOF)
	attackname=rst("��ʽ")
	msgarr=msgarr&attackname&";"
	msg=msg+"<tr><td bgcolor=FEE3AB><a href=javascript:settalk('//����','"&attackname&"'); target='talkfrm' onmouseover=""window.status='ʹ���书��������';return true;"" onmouseout=""window.status='';return true;"" title='ʹ���书��������'><font color=FF0000 valign='middle'>"&attackname&"</font></a></td><td>"&rst("��������")&"</td></tr>"
	rst.movenext
loop
rst.Close
set rst=nothing
conn.Close
set conn=nothing
if msg="" then msg="<tr valign=middle height='80%'><td colspan=2>�Բ����ˣ���û��ѧ���κ���ʽ�������޷�����</td></tr>"
%>
<head><link rel="stylesheet" href="css.css"></head>
<body bgcolor='<%=bgcolor%>' background='<%=bgimage%>' oncontextmenu=self.event.returnValue=false leftmargin="5" topmargin="2" marginwidth="5" marginheight="2">
<div align=center><font size="4" color="#CC0000" face="��Բ"><b><%=Application("Ba_jxqy_systemname"&chatroomsn)%></b></font><br>
  <font color=FF0000><%=Application("Ba_jxqy_onlinenum"&chatroomsn)%></font>/<font color=008800><%=Application("Ba_jxqy_allonlinenum")%></font>������
  <hr noshade size="1" color=darkred>
<font color=0000FF>��ʽ����</font><br>
  <table width='99%' border=1 cellspacing="1" cellpadding="1" align="right" bordercolor="#FF9933">
    <tr align=center>
      <td height="19" width="65%">��ʽ</td>
      <td height="19" width="36%">����</td>
    </tr>
<%=msg%>
</table>
</div>
</body>