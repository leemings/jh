<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
st=Session("yx8_mhjh_username")
set conn=server.CreateObject("adodb.connection")
conn.Open Application("yx8_mhjh_connstr")
set rs=server.CreateObject("adodb.recordset")
if Application("yx8_mhjh_klt1")=0 then
	Response.Write "<script Language=Javascript>alert('��ʾ���㶯�����߷�Ӧ̫��������ս��Ʒ�Ѿ����������ˣ�');</script>"
	Response.Write "<script Language=Javascript>location.href ='javascript:window.close()';</script>"
	Response.end
end if
tempjs=int(abs(clng(Application("yx8_mhjh_klt1"))))
'if tempjs<>jstl1 then
	'Response.Write "<script Language=Javascript>alert('��ʾ�����������ʲô�����س���������!');</script>"
	'Response.Write "<script Language=Javascript>location.href ='javascript:window.close()';</script>"
	'response.end
'end if
conn.execute "update �û� set ����=����-"&Application("yx8_mhjh_klt1")&" where ����='"&st&"'"
rs.open "select ����,���� FROM �û� WHERE ����='"&st&"'",conn
if rs("����")="ʮ�˵���" then
	Response.Write "<script Language=Javascript>alert('��ʾ����С���ǲ���û������ѽ����껹��ñ��˰���');</script>"
	Response.Write "<script Language=Javascript>location.href ='javascript:window.close()';</script>"
	Application("yx8_mhjh_klt1")=0
	Response.end
end if
if rs("����")<0 then
	conn.execute "update �û� set ״̬='����',��ɱ=��ɱ+1 where ����='"&st&"'"
conn.execute "insert into Ӣ����(ʱ��,����,����,����) values(now(),'"&username&"','����','������������')"
	kl="<img src='data/kl.gif'>̫�����ˣ�<font color=0000FF>"&st&"</font>����ս��Ʒʱ����ǿ��,���������������,�й��ҵ�!"
	Response.Write "<script Language=Javascript>alert('��ʾ����̰�Ʋ�Ҫ��������̫��������С����ȥ����ɣ�û���ͱ��Ҵ���');</script>"
	session.Abandon
	response.write "<script language=javascript>parent.top.location.href='chaterror.asp?id=009';</script>"
else
	conn.execute "update �û� set ����=����+"& Application("yx8_mhjh_klt1") *2000&",����=����+"&Application("yx8_mhjh_klt")*1000&" where ����='" & st & "'"
	kl="<font color=0000FF>"&st&"</font>����Ѹ�٣��ɹ����������ҵ��µĵ���<img src=IMAGE/IMAGE/wg8.gif>"&Application("yx8_mhjh_klt1")*2000&"�㣬����"&Application("yx8_mhjh_klt1")*1000&"�㣡"
	Application.Lock
	Application("yx8_mhjh_klt1")=0
	Application.UnLock
	rs.close
set rs=nothing
conn.close
set conn=nothing
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
		newtalkarr(594)=username
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
