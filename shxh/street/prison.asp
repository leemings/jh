<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
bgcolor=Application("Ba_jxqy_backgroundcolor")
bgimage=Application("Ba_jxqy_backgroundimage")
mycorp=session("Ba_jxqy_usercorp")
mygrade=session("Ba_jxqy_usergrade")
if mygrade="" then Response.Redirect "../error.asp?id=016"
nowtime=now()
set conn=server.CreateObject("adodb.connection")
conn.Open Application("Ba_jxqy_connstr")
set rst=server.CreateObject("adodb.recordset")
sqlstr="select ����¼ʱ��,����,״̬ from �û� where ״̬ in('����','����')"
rst.Open sqlstr,conn
do while not (rst.EOF or rst.BOF )
	if rst("״̬")="����" then
		if mycorp="�ٸ�" and mygrade>=Application("Ba_jxqy_arrestright") then
			msg=msg&"<tr><td>"&rst("����")&"</td><td>"&rst("����¼ʱ��")&"</td><td><a href='assoil.asp?un="&rst("����")&"' onmouseover="&chr(34)&"window.status='�ͷ�"&rst("����")&"';return true;"&chr(34)&" onmouseout="&chr(34)&"window.status='';return true;"&chr(34)&">�ͷ�</a></td></tr>"
		else
			msg=msg&"<tr><td>"&rst("����")&"</td><td>"&rst("����¼ʱ��")&"</td><td><a href='mainprise.asp?un="&rst("����")&"'  onmouseover="&chr(34)&"window.status='����"&rst("����")&"';return true;"&chr(34)&" onmouseout="&chr(34)&"window.status='';return true;"&chr(34)&">����</a></td></tr>"	
		end if
	else
		if mycorp="�ٸ�" and mygrade>=Application("Ba_jxqy_gaolright") then
			msg=msg&"<tr><td>"&rst("����")&"</td><td>"&rst("����¼ʱ��")&"</td><td><a href='assoil.asp?un="&rst("����")&"'  onmouseover="&chr(34)&"window.status='�ͷ�"&rst("����")&"';return true;"&chr(34)&" onmouseout="&chr(34)&"window.status='';return true;"&chr(34)&">�ͷ�</a></td></tr>"
		else
			msg=msg&"<tr><td>"&rst("����")&"</td><td>"&rst("����¼ʱ��")&"</td><td>���ɱ���</td></tr>"
		end if
	end if		
	rst.MoveNext
loop
rst.Close
set rst=nothing 
conn.Close 
set conn=nothing
%>
<head><title><%=Application("Ba_jxqy_systemname")%></title><LINK href="../chatroom/css.css" rel=stylesheet></head>
<body bgcolor='<%=bgcolor%>' background='<%=bgimage%>' oncontextmenu=self.event.returnValue=false>
<p align=center><font color="#CC0000" size="4" face="��Բ"><b><%=Application("Ba_jxqy_systemname")%>����</b></font></p>
<div align=center>
  <table width='95%' border=1 bordercolorlight=000000 cellspacing=0 cellpadding=2 bordercolordark=FFFFFF>
    <tr align=center bgcolor=#FFFF00> 
      <td>����</td>
      <td>����ʱ��</td>
      <td>����</td>
    </tr>
<%=msg%>
</table>
</div>
<p align=center><input type="button" value=" �� �� " onClick="javascript:top.window.close();" name="button"></p>
</body>