<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
chatroomsn=Session("Ba_jxqy_userchatroomsn")
bgcolor=Application("Ba_jxqy_backgroundcolor")
bgimage=Application("Ba_jxqy_backgroundimage")
username=session("Ba_jxqy_username")
msg="<table border='1' cellspacing='0' cellpadding='1' bordercolor='#800000' bordercolorlight='#800000' bordercolordark='#FFFFFF' width='100%'>"
set conn=server.CreateObject("adodb.connection")
conn.Open Application("Ba_jxqy_connstr")
set rst=server.CreateObject("adodb.recordset")
rst.Open "select ����,���� from ��Ʒ where ������='"&username&"' and װ��=true order by ����",conn
do while not rst.EOF
	msg=msg&"<tr><td width=35 bgcolor=#FFff00><font color=red>"&rst("����")&"</font></td><td ><a href='discharge.asp?cmtype="&rst("����")&"' onmouseover=""window.status='ж��';return true;"" onmouseout=""window.status='';return true;"">"&rst("����")&"</a></td></tr>"
	rst.MoveNext
loop
msg=msg&"</table><font color='#006600'>��ӵ��װ����</font><table border='1' cellspacing='0' cellpadding='1' bordercolor='#800000' bordercolorlight='#800000' bordercolordark='#FFFFFF' width='100%'>"
rst.Close
sqlstr="select ����,����,���� from ��Ʒ where ���� in('ͷ��','����','����','����','��Ʒ') and ������='"&username&"' and ����>0 order by �۸�"
rst.Open sqlstr,conn
do while not (rst.BOF or rst.EOF)
	msg=msg&"<tr><td width=35 bgcolor=#FFCC00><font color=red>"&rst("����")&"</font></td><td><a href='equip.asp?commodityname="&rst("����")&"' onmouseover=""window.status='װ��';return true;"" onmouseout=""window.status='';return true;"">"&rst("����")&"</a>(x"&rst("����")&")</td></tr>"
	rst.MoveNext
loop
msg=msg&"</table>"
rst.Close 
set rst=nothing
conn.Close
set conn=nothing
%>
<head><link rel="stylesheet" href="style1.css"></head>
<body bgcolor='<%=bgcolor%>' background='<%=bgimage%>' oncontextmenu=self.event.returnValue=false leftmargin="5" topmargin="0" marginwidth="5" marginheight="0">
<div align=center><font size="4" color="#990000"><b><font face="��Բ"><%=Application("Ba_jxqy_systemname"&chatroomsn)%></font></b></font><br>
  <font color=FF0000><%=Application("Ba_jxqy_onlinenum"&chatroomsn)%></font>/<font color=008800><%=Application("Ba_jxqy_allonlinenum")%></font>������
  <hr noshade size="1">
  <font color="#006600">���鿴װ����</font><br>
<%=msg%>
</div>
</body>
