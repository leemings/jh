<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
username=Session("yx8_mhjh_username")
set conn=server.CreateObject("adodb.connection")
conn.Open Application("yx8_mhjh_connstr")
set rst=server.CreateObject("adodb.recordset")
if Application("yx8_mhjh_klt4")=0 then
	Response.Write "<script Language=Javascript>alert('��ʾ���㶯�����߷�Ӧ̫��������ս��Ʒ�Ѿ����������ˣ�');</script>"
	Response.Write "<script Language=Javascript>location.href ='javascript:window.close()';</script>"
	Response.end
end if
conn.execute "update �û� set ����=����-"&Application("yx8_mhjh_klt4")&" where ����='"&username&"'"
rst.open "select ����,���� FROM �û� WHERE ����='"&username&"'",conn
if rst("����")="ʮ�˵���" then
	Response.Write "<script Language=Javascript>alert('��ʾ����С���ǲ���û������ѽ����껹��ñ��˰���');</script>"
	Response.Write "<script Language=Javascript>location.href ='javascript:window.close()';</script>"
	Application("yx8_mhjh_klt4")=0
	Response.end
end if
if rst("����")<0 then
	conn.execute "update �û� set ״̬='����',��ɱ=��ɱ+1 where ����='"&username&"'"
conn.execute "insert into Ӣ����(ʱ��,����,����,����) values(now(),'"&username&"','���������','������������')"
	kl="<img src='data/kl.gif'>̫�����ˣ�<font color=0000FF>"&username&"</font>����ս��Ʒʱ����ǿ��,���������������,�й��ҵ�!"
	Response.Write "<script Language=Javascript>alert('��ʾ����̰�Ʋ�Ҫ��������̫��������С����ȥ����ɣ�û���ͱ��Ҵ���');</script>"
	session.Abandon
	response.write "<script language=javascript>parent.top.location.href='chaterror.asp?id=009';</script>"
else
rst.close 
set conn=server.CreateObject("adodb.connection")
conn.Open Application("yx8_mhjh_connstr")
set rst=server.CreateObject("adodb.recordset")
rst.cursorlocation=3 
rst.Open "select * from �̵�",conn
randomize       
r=150
id=int((r-1)*rnd)+1 
rst.close
set rst=server.CreateObject("adodb.recordset")
rst.Open "select * from �̵� where id="&id,conn
ming=rst("����")
rst.Close
set rst=server.CreateObject("adodb.recordset")
rst.Open "select * from ��Ʒ where ������='"&username&"' and ����='"&ming&"'",conn
if rst.EOF or rst.BOF then
conn.Execute "insert into ��Ʒ(����,����,����,����,����,����,��Ч,����,�۸�,������,����,װ��) values('"&rst("����")&"','"&rst("����")&"',"&rst("����")&","&rst("����")&","&rst("����")&","&rst("����")&",'"&rst("��Ч")&"',1,"&rst("�۸�")\2&",'"&username&"',false,false)"
else
         conn.Execute "update ��Ʒ set ����=����+1 where ������='"&username&"' and ����='"&ming&"'"
	kl="<font color=0000FF>"&username&"</font>����Ѹ�٣��ɹ����������ҵ��µ�<img src=IMAGE/IMAGE/wg7.gif>" & ming &"��"
	Application("yx8_mhjh_klt4")=0
rst.Close
set rst=nothing
conn.Close
set conn=nothing
end if
end if
says="<font color=red><b>����ս��Ʒ��</b></font>"&kl	
dim newtalkarr(600) 
talkarr=Application("yx8_mhjh_talkarr") 
		talkpoint=Application("yx8_mhjh_talkpoint") 
		j=1 
		for i=11 to 600 step 10 
			newtalkarr(j)=talkarr(i) 
			newtalkarr(j+1)=talkarr(i+1) 
			newtalkarr(j+2)=talkarr(i+2) 
			newtalkarr(j+3)=talkarr(i+3) 
			newtalkarr(j+4)=talkarr(i+4) 
			newtalkarr(j+5)=talkarr(i+5) 
			newtalkarr(j+6)=talkarr(i+6) 
			newtalkarr(j+7)=talkarr(i+7) 
			newtalkarr(j+8)=talkarr(i+8) 
			newtalkarr(j+9)=talkarr(i+9) 
			j=j+10 
		next 
		newtalkarr(591)=talkpoint+1 
		newtalkarr(592)=1 
		newtalkarr(593)=0 
		newtalkarr(594)=name
		newtalkarr(595)="���" 
		newtalkarr(596)="" 
		newtalkarr(597)="000000" 
		newtalkarr(598)="008000" 
		newtalkarr(599)=""&says&""
        newtalkarr(600)=Session("yx8_mhjh_userchatroomsn")
        Application.Lock
        Application("yx8_mhjh_talkpoint")=talkpoint+1
        Application("yx8_mhjh_talkarr")=newtalkarr
        Application.UnLock
        erase newtalkarr
        erase talkarr
%>
<script Language=Javascript>alert('��ʾ������ս��Ʒ�ɹ�,ע�⿴�������ڵĽ���ͨ����');</script>"
	Response.Write "<script Language=Javascript>location.href ='javascript:window.close()';</script>
