<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
jstl=Application("yx8_mhjh_klt")
un=Session("yx8_mhjh_username")
if Application("yx8_mhjh_klt")=0 then
	Response.Write "<script Language=Javascript>alert('��ʾ���㶯�����߷�Ӧ̫���������Ѿ�����ˣ�')</script>"
	Response.Write "<script Language=Javascript>location.href ='javascript:window.close()';</script>"
	Response.end
end if
tempjs=int(abs(clng(Application("yx8_mhjh_klt"))))
'if tempjs<>jstl then
	'Response.Write "<script Language=Javascript>alert('��ʾ�����������ʲô�����س���������!');</script>"
	'Response.Write "<script Language=Javascript>location.href ='javascript:window.close()';</script>"
	'response.end
'end if
randomize          
r=rst.recordcount 
id=int((r-1)*rnd)+1 
Set conn=server.CreateObject("adodb.connection")
conn.Open Application("yx8_mhjh_connstr")
set rst = server.createobject("adodb.recordset") 
rst.open "select * from �̵�" where id="&id
rst.Open sqlstr,conn
if rst.EOF or rst.BOF then
says="<font color=red><b>����������,ʲôҲû�õ�</b></font>"
else
commodityname=rst("����")
commoditytype=rst("����")
commodityhealth=rst("����")
magic=rst("����")
attack=rst("����")
defence=rst("����")
especial=rst("��Ч")
commodityprice=rst("�۸�")
rst.Close
rst.Open "select ���� from �û� where ����='"&username&"' and ����>10000,conn
if rst.EOF or rst.BOF then
msg="�������Ǯ����û�д���Ӵ��"
else
rst.Close
sqlstr="select * from ��Ʒ where ����='"&commodityname&"' and ������='"&un&"'"
rst.Open sqlstr,conn
if rst.EOF or rst.BOF then
sqlstr="insert into ��Ʒ(����,����,����,����,����,����,��Ч,�۸�,����,������,����,װ��) values('"&commodityname&"','"&commoditytype&"',"&commodityhealth&","&magic&","&attack&","&defence&",'"&especial&"',"&commodityprice/2&",1,'"&un&"',False,False)"
else
sqlstr="update ��Ʒ set ����=����+1 where ����='"&commodityname&"' and ������='"&un&"'"
	kl="����ǧ�����İ��棬["&un&"]���ڵִ���ϣ���ı˰����ɹ�ȡ�òʺ�֮��ı��أ���Ȼ�ķ���1������ȴ�õ�һ�����ر�ʯ��["&un&"]���ڹ���ƶ�����ҹٸ������˰���<img src=data/251.gif>��������10000�㣡"
end if
end if
rst.Close
set rst=nothing
conn.Close
set conn=nothing
end if
says="<font color=red><b>������ս��Ʒ��</b></font>"&kl	
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
		newtalkarr(594)=username
		newtalkarr(595)="���" 
		newtalkarr(596)="" 
		newtalkarr(597)="000000" 
		newtalkarr(598)="000000" 
		newtalkarr(599)=""&says&""
        newtalkarr(600)=Session("yx8_mhjh_userchatroomsn")
        Application.Lock
        Application("yx8_mhjh_talkpoint")=talkpoint+1
        Application("yx8_mhjh_talkarr")=newtalkarr
        Application.UnLock
        erase newtalkarr
        erase talkarr
%>
<script Language=Javascript>location.href ='javascript:window.close()';</script>
