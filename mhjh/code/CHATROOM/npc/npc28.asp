<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
username=Session("yx8_mhjh_username")
set conn=server.CreateObject("adodb.connection")
conn.Open Application("yx8_mhjh_connstr")
set rs=server.CreateObject("adodb.recordset")
if Application("yx8_mhjh_klt28")=0 then
	Response.Write "<script Language=Javascript>alert('��ʾ�������Ѿ������ˣ�');</script>"
	Response.Write "<script Language=Javascript>location.href = '../option.asp';</script>"
	Response.end
end if
tempjs=int(abs(clng(Application("yx8_mhjh_klt28"))))
'if tempjs<>r then
	'Response.Write "<script Language=Javascript>alert('��ʾ�����������ʲô�����س���������!');</script>"
	'Response.Write "<script Language=Javascript>location.href = '../option.asp';</script>"
	'response.end
'end if
conn.execute "update �û� set ����=����-"&Application("yx8_mhjh_klt28")&" where ����='"&username&"'"
rs.open "select ����,���� FROM �û� WHERE ����='"&username&"'",conn
if rs("����")="ʮ�˵���" then
	Response.Write "<script Language=Javascript>alert('��ʾ����С���ǲ���û������ѽ����껹��ñ��˰���');</script>"
	Response.Write "<script Language=Javascript>location.href = '../option.asp';</script>"
	Application("yx8_mhjh_klt28")=0
	Response.end
end if
if rs("����")<0 then
	conn.execute "update �û� set ״̬='����',��ɱ=��ɱ+1 where ����='"&username&"'"
conn.execute "insert into Ӣ����(ʱ��,����,����,����) values(now(),'"&username&"','����','������������')"
	kl="<img src='data/kl.gif'>̫�����ˣ�["&username&"]Ϊ�˵�������Ϯ��������ȡ�壬��������Զ�������˲ʺ�֮�£��������ֻ��ǧ�������ڱ���������â�Ĵ�˵������"
	Response.Write "<script Language=Javascript>alert('��ʾ����Ӣ�¿ɼΣ���������̫��������С����ȥ����ɣ�û���ͱ��Ҵ���');</script>"
	session.Abandon
	response.write "<script language=javascript>parent.top.location.href='../chaterror.asp?id=009';</script>"
else
randomize
r=128 
id=int((r-1)*rnd)+1 
set conn=server.CreateObject("adodb.connection")
conn.Open Application("yx8_mhjh_connstr")
set rs = server.createobject("adodb.recordset") 
rs.open "select * from �̵� where id="&id,conn
if rs.EOF or rs.BOF then
Response.Write "<script Language=Javascript>alert('�ܲ��ң���ʲô��û�õ���');</script>"
	Response.Write "<script Language=Javascript>location.href ='javascript:window.close()';</script>"
Application("yx8_mhjh_klt28")=0
	Response.end
end if
ming=rs("����")
rs.Close
rs.Open "select * from ��Ʒ where ������='"&username&"' and ����='"&ming&"'",conn
If Rs.Bof OR Rs.Eof then
rs.close 
rs.open "select * from �̵� where ����='"&ming&"'",conn
conn.Execute "insert into ��Ʒ(����,����,����,����,����,����,��Ч,����,�۸�,������,����,װ��) values('"&ming&"','"&rs("����")&"',"&rs("����")&","&rs("����")&","&rs("����")&","&rs("����")&",'"&rs("��Ч")&"',1,"&rs("�۸�")\2&",'"&username&"',false,false)"
	kl="<font color=#0000FF>"&username&"</font>�ɹ������˹���,����ù������ϵ�"&ming&"��"
Application("yx8_mhjh_klt28")=0
else
         conn.Execute "update ��Ʒ set ����=����+1 where ������='"&username&"' and ����='"&ming&"'"
	kl="<font color=#0000FF>"&username&"</font>�ɹ������˹���,����ù������ϵ�"&ming&"��"
	Application.Lock
	Application("yx8_mhjh_klt28")=0
	Application.UnLock
	rs.close
set rs=nothing
conn.close
set conn=nothing
end if
end if
says="<font color=red><b>����ֳɹ���</b></font>"&kl	
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
<script Language=Javascript>location.href = "../option.asp";</script>
