<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
chatroomsn=Session("Ba_jxqy_userchatroomsn")
bgcolor=Application("Ba_jxqy_backgroundcolor")
bgimage=Application("Ba_jxqy_backgroundimage")
mycorp=session("Ba_jxqy_usercorp")
username=session("Ba_jxqy_username")
Set conn=server.CreateObject("adodb.connection")
conn.Open Application("Ba_jxqy_connstr")
set rst=server.CreateObject("adodb.recordset")
sqlstr="select ��ʽ,�ȼ� from ���� where ����='"&username&"' order by �ȼ� desc"
rst.Open sqlstr,conn
msgarr=";"
do while not (rst.EOF or rst.BOF)
	attackname=rst("��ʽ")
	msgarr=msgarr&attackname&";"
	msg=msg+"<tr><td><a href='learn2.asp?mg="&attackname&"' onmouseover=""window.status='����ҵ��ѧ��Ĺ���';return true;"" onmouseout=""window.status='';return true;"" title='����ҵ��ѧ��Ĺ���'>"&attackname&"</a></td><td>"&rst("�ȼ�")&"��</td></tr>"
	rst.movenext
loop
rst.Close
sqlstr="select ��ʽ from ��ʽ where ����='"&mycorp&"'"
rst.Open sqlstr,conn
do while not (rst.EOF or rst.BOF)
	attackname=rst("��ʽ")
	if instr(msgarr,";"&attackname&";")=0 then msg=msg+"<tr><td bgcolor=FEE3AB><a href='learn2.asp?mg="&attackname&"' onmouseover=""window.status='��ϰ�µĹ���';return true;"" onmouseout=""window.status='';return true;"" title='��ϰ�µĹ���'><font color=FF0000>"&attackname&"</font></a></td><td>0��</td></tr>"
	rst.MoveNext
loop
rst.Close
set rst=nothing
conn.Close
set conn=nothing
if msg="" then msg="<tr valign=middle height='80%'><td colspan>�Բ����ˣ�û��������������书</td></tr>"
%>
<head><link rel="stylesheet" href="css.css"></head>
<body bgcolor='<%=bgcolor%>' background='<%=bgimage%>' oncontextmenu=self.event.returnValue=false>
<div align=center><font size="4" color="#CC0000" face="��Բ"><b><%=Application("Ba_jxqy_systemname"&chatroomsn)%></b></font><br>
  <font color=FF0000><%=Application("Ba_jxqy_onlinenum"&chatroomsn)%></font>/<font color=008800><%=Application("Ba_jxqy_allonlinenum")%></font>������<hr>
<font color=0000FF>��ϰ�书</font><br>
  <table width="95%" border="1" cellpadding="2" cellspacing="1" bordercolor="#FF9933">
        <%=msg%>
   </table>
</div>
</body>
