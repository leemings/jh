<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
username=session("yx8_mhjh_username")
set conn=server.CreateObject("adodb.connection")
conn.Open Application("yx8_mhjh_connstr")
set rst=server.CreateObject("adodb.recordset")
sqlstr="select ����,��Ч,���� from ��Ʒ where ����='����' and ������='"&username&"' and ����>0 order by �۸�"
rst.Open sqlstr,conn
msg="<table width=95% border=1 bordercolorlight=000000 cellspacing=0 cellpadding=1 bordercolordark=FFFFFF><tr></tr>"
do while not (rst.BOF or rst.EOF)
msg=msg&"<tr><td><a href=javascript:parent.talkfrm.settalk('//Ͷ��','"&rst("����")&"') target='talkfrm' onmouseover=""window.status='ʹ�ð���';return true;"" onmouseout=""window.status='';return true;"" title="&chr(34)&rst("��Ч")&chr(34)&">"&rst("����")&"</a></td><td width=15  bgcolor=ffff00>"&rst("����")&"</td></tr>"
rst.MoveNext
loop
msg=msg&"</table>"
rst.Close
set rst=nothing
conn.Close
set conn=nothing
%>
<head><link rel="stylesheet" href="css3.css"></head>
<body  oncontextmenu=self.event.returnValue=false leftmargin="4" marginwidth="4" background='bg1.gif'>
<div align=center><br>
<font color=008800><%=Application("yx8_mhjh_allonlinenum")%></font>������
<hr noshade size="1" color=red>
<font color=0000FF>ӵ�а���</font><br>
<%=msg%>
</div>
<br>&nbsp;&nbsp;&nbsp;<a href="#" onClick="window.open('../js/armorshop.asp','reg','width=580,height=550,scrollbars=yes')"><font color="#F71004">ȥ����������</font></a> 
</body>
