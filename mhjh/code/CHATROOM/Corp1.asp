<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
set conn=server.CreateObject("adodb.connection")
conn.Open Application("yx8_mhjh_connstr")
set rst=server.CreateObject("adodb.recordset")
sqlstr="select ����,���,���� from ���� where ��������<>'False'"
rst.Open sqlstr,conn
do while not (rst.EOF or rst.BOF)
corp=rst("����")
tl=rst("����")
msg=msg+"<tr><td><a href='javascript:parent.chgsendto("&chr(34)&corp&chr(34)&");' target='talkfrm' onmouseover=""window.status='��������ɵ�Ҫ��';return true;"" onmouseout=""window.status='';return true;"" title='��������ɵ�Ҫ��'><font color=#CC0000>"&corp&"</font></a></td><td>"&tl&"</td></tr>"
rst.movenext
loop
rst.Close
set rst=nothing
conn.Close
set conn=nothing
if msg="" then msg="<tr valign=middle height='80%'><td colspan=2>���۾���ë�����Բ���û���������������㹥�����ǲ�����ưԽ�������ˣ�ȥ��㿹����ҩƷ�԰ɣ�������</td></tr>"
%>
<head><link rel="stylesheet" href="css2.css">
</head>
<body  oncontextmenu=self.event.returnValue=false leftmargin="5" marginwidth="5" background='bg1.gif'>
<div align=center><br><font color=008800><%=Application("yx8_mhjh_allonlinenum")%></font>������
<hr noshade size="1" color=red>
<font color=0000FF><b>��������Ҫ��</b></font><br>
<font color="#FF0000"><b>[</b>���������֣��ٷ��ԣ��Ϳ��Թ���Ҫ����һͳ�����Ĵ�ҵ��������ʵ�֣�<b>]</b></font><br>
<table width=98% border=1 bordercolorlight=000000 cellspacing=0 cellpadding=2 bordercolordark=FFFFFF>
<%=msg%>
</table>
</div>
</body>
