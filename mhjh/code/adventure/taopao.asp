<%
username=session("yx8_mhjh_username")
mygrade=Session("yx8_mhjh_usergrade")
biological=session("yx8_mhjh_userfight")
if biological="none" then
msg="<FONT color=#ff0000>����������</FONT>û�ж��������,���ʲô,��С��!<br>"
elseif username="" then
msg="<FONT color=#ff0000>����������</FONT>��û�е�¼��ʱ�Ͽ�����,�����½���<br>"
else
randomize()
nowtime=now()
nowtimetype="#"&month(nowtime)&"/"&day(nowtime)&"/"&year(nowtime)&" "&hour(nowtime)&":"&minute(nowtime)&":"&second(nowtime)&"#"
set conn=server.CreateObject("adodb.connection")
conn.Open Application("yx8_mhjh_connstr")
set rst=server.CreateObject("adodb.recordset")
rst.Open "select ����,���� from �û� where ����='"&username&"' and ����>=100000",conn
if rst.EOF or rst.BOF then
msg="<FONT color=#ff0000>����������</FONT>��������ۿ�Ҫ���Ĺ���,û����������,�Ͽ�ʹ��ǧ�ﴫ�������ֽ���ȥ�����������!<br>"
else
fhp=rst("����")
fdefence=rst("����")
rst.Close
rst.Open "select * from biological where biological='"&biological&"'",conn
maxsinew=rst("hp")
tattack=rst("attack")
tdefence=rst("defence")
encourage=rst("encourage")
encourage_m=Mid(encourage,1,2)
encourage_c=Mid(encourage,4)
defence=tattack-fdefence
Maxattack=Application("yx8_mhjh_Maxattack")*mygrade\50
randomize()
if defence>0 then
defence=Maxattack-clng(rnd()*1000)
elseif defence<=0 then
defence=clng(rnd()*10000)
end if
fhp=fhp-defence
if fhp<=0 then
conn.Execute "update �û� set ״̬='����',����¼ʱ��="&nowtimetype&" where ����='"&username&"'"
session.Abandon
Response.Write "<script language=javascript>top.location.replace('../error.asp?id=054');</script>"
msg="<FONT color=#ff0000>�������Ұ���</FONT>�㱻"&biological&"������!<br>"
else
conn.Execute "update �û� set ����="&fhp&",����=����+10 where ����='"&username&"'"
rndcatch=rnd()*99+1
if rndcatch<25 then
rst.Close
rst.Open "select * from pet where username='"&username&"' and exist=true"
if rst.EOF or rst.BOF then
msg="<FONT color=#ff0000>������ʧ�ܡ�</FONT>�㱿�ױ���,ˤ��һ����ͷ,"&biological&"�ķ���ʹ���������½�"&defence&"<br>"
else
session("yx8_mhjh_userfight")="none"
msg="<FONT color=#ff0000>�����ܳɹ���</FONT>�����������������,��"&biological&"�ķ���ʹ���������½�"&defence&"<br>"
end if
else
msg="<FONT color=#ff0000>������ʧ�ܡ�</FONT>"&biological&"�����ʹ���������½�"&defence&"<br>"
end if
end if
end if
rst.Close
set rst=nothing
conn.Close
set conn=nothing
end if
Response.Write "<script language=javascript>parent.msgfrm.document.writeln('"&msg&"');parent.confrm.document.form1.move.disabled=false;parent.behfrm.location.replace('Action.asp');</script>"
%>
