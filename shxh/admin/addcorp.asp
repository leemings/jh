<%
bgcolor=Application("Ba_jxqy_backgroundcolor")
bgimage=Application("Ba_jxqy_backgroundimage")
username=session("Ba_jxqy_username")
if username="" then Response.redirect "../error.asp?id=016"
set conn=server.CreateObject ("adodb.connection")
conn.Open Application("Ba_jxqy_connstr")
set rst=server.CreateObject("Adodb.recordset")
rst.Open "select * from �û� where ����='"&username&"'",conn
msg="<head><link rel='stylesheet' href='../chatroom/style1.css'></head><body oncontextmenu=self.event.returnValue=false background='"&bgimage&"' bgcolor='"&bgcolor&"' topmargic=100><p align=center><font color=ff0000 size=5>�����Ż�</font><hr>����:����,����,����,������10��,����100��,����1��<br>ע�⣺�����ɹ������Ժ��ǲ������޸ĵ�,��������д�������<br>����״̬,����:"&rst("����")&";����:"&rst("����")&";����:"&rst("����")&";����:"&rst("����")&";����:"&rst("����")&";����:"&rst("����")
if rst("����")>100000 and rst("����")>100000 and rst("����")>100000 and rst("����")>100000 and rst("����")>1000000  and rst("���")<>"����" and rst("����")<>"�ٸ�" and rst("����")>10000 then msg=msg&"<form action=addcorp2.asp method=post><table border=3 width='80%' align=center><tr><td>��������</td><td><input type=text name='corpname' value='' maxlength=7 size=7></td></tr><tr><td>����˵��</td><td><input type=text name='corpcond' value=''��size=50 maxlength=100></td><tr><td>�����ʽ�</td><td><input type=text value='1000000' name='corpsilver' size=10 maxlength=10></td></tr></tr><tr><td>��������</td><td><input type=radio name='tj1' value='True' checked>����<br><input type=radio name=tj1 value='False'>ȫ��<br><input type=radio name=tj1 value='Other'>����<ul><dd>�Ա�<input type=checkbox name=sexchk><input type=radio name=sex value='��' checked>��<input type=radio name=sex value='Ů'>Ů<dd>����<input type=checkbox name='hpchk'><input type=radio name=hpopt value='>='  checked>���ڵ���<input type=radio name=hpopt value='<='>С�ڵ���<input type=text name=hp value=500 size=8 maxlength=8><dd>����<input type=checkbox name='attackchk'><input type=radio name=attackopt value='>='  checked>���ڵ���<input type=radio name=attackopt value='<='>С�ڵ���<input type=text name=attack value=500 size=8 maxlength=8><dd>����<input type=checkbox name='defencechk'><input type=radio name=defenceopt value='>='  checked>���ڵ���<input type=radio name=defenceopt value='<='>С�ڵ���<input type=text name=defence value=500 size=8 maxlength=8><dd>����<input type=checkbox name='moralchk'><input type=radio name=moralopt value='>='  checked>���ڵ���<input type=radio name=moralopt value='<='>С�ڵ���<input type=text name=moral value=500 size=8 maxlength=8><dd>����<input type=checkbox name='silverchk'><input type=radio name=silveropt value='>='  checked>���ڵ���<input type=radio name=silveropt value='<='>С�ڵ���<input type=text name=silver value=500 size=8 maxlength=8></ul></td></tr><tr><td colspan=2 align=center><input type=submit value='��������'></td></tr></table></form>"
if rst("���")="����" then msg=msg&"<br><a href='corpman.asp?corp="&rst("����")&"'>�����б�</a> <a href='corpatt.asp?corp="&rst("����")&"'>�书����</a>"
rst.Close
set rst=nothing
conn.Close
set conn=nothing
Response.Write msg&"</body>"
%>