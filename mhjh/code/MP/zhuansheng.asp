<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
username=session("yx8_mhjh_username")
if username="" then Response.Redirect "../error.asp?id=016"
if session("yx8_mhjh_inchat")="" then
Response.write "<script language='javascript'>alert ('�㲻�ܽ���,���Ƚ���������������лл����'); window.close();</script>"
Response.End 
        end if
set conn=server.createobject("adodb.connection")
conn.open Application("yx8_mhjh_connstr")
set rs=server.CreateObject ("adodb.recordset")
sql="select ����,����,�ȼ�,����,�������� from �û� where ����='"&username&"'"
set rs=conn.execute(sql)
if rs.eof or rs.bof then Response.Redirect "../error.asp?id=016"
jf=rs("����")
zz=rs("����")
yl=rs("����")
dj=rs("�ȼ�")
if dj<100 then
mess=""&username&"�㻹����100����,��ô���,����������ʲô,��ȥ������˵�ɣ�"
elseif jf<4000000 then
mess=""&username&"����ֲ���400��,Ŀǰ���޷�ʵ��ת����"
elseif yl<100000000 then
mess=""&username&"���Ǯ����������ô�߼�����,�����Ǯ���ܲ��£�������˵��׻���ô��Ŷ��"
elseif zz<10000 then
mess=""&username&"������̫��,����Ҫ�ﵽ100���ˮƽ,���ﲻ�ܽ������������ǵ��ˣ�"
else
yl1=yl-100000000
dd=rs("����")-10000
if rs("��������")="" or rs("��������")="��" then
mess="��ϲ��"&username&"�Ѿ�ʵ����ת��������ͻ��ֶ����0�����ô����Ŷ���Ժ��Լ���·�Լ��߰�<br><br>���˳������µ�½"
conn.execute "update �û� set ����="&yl1&",����="&dd&",����=0,�ȼ�=0,��������='ת����',������=5000000,������=500000 where ����='"&username&"'"
elseif rs("��������")="ת����" then
mess="��ϲ��"&username&"�Ѿ�ʵ������ת��������Ϊǿ���ˣ�����ͻ��ֶ����0�����ô����Ŷ���Ժ��Լ���·�Լ��߰�<br><br>���˳������µ�½"
conn.execute "update �û� set ����="&yl1&",����="&dd&",����=0,�ȼ�=0,��������='ǿ����',������=10000000,������=1000000 where ����='"&username&"'"
elseif rs("��������")="ǿ����" then
mess="��ϲ��"&username&"�Ѿ�ʵ������ǿ��������Ϊ�����ˣ�����ͻ��ֶ����0�����ô����Ŷ���Ժ��Լ���·�Լ��߰�<br><br>���˳������µ�½"
conn.execute "update �û� set ����="&yl1&",����="&dd&",����=0,�ȼ�=0,��������='������',������=15000000,������=1500000 where ����='"&username&"'"
elseif rs("��������")="������" then
mess="��ϲ��"&username&"�Ѿ�ʵ�����ɾ���������Ϊ��*�޵У�����ͻ��ֶ����0�����ô����Ŷ���Ժ��Լ���·�Լ��߰�<br><br>���˳������µ�½"
conn.execute "update �û� set ����="&yl1&",����="&dd&",����=0,�ȼ�=0,��������='��*�޵�',������=20000000,������=2000000 where ����='"&username&"'"
end if
end if
rs.Close             
set rs=nothing             
conn.Close             
set conn=nothing       
says="<font color=red><b>������ת����</b></font>"&mess	
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
<script Language=Javascript>alert('<%=mess%>��');</script>"
	Response.Write "<script Language=Javascript>location.href ='javascript:window.close()';</script>
