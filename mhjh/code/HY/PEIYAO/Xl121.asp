<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
name=session("yx8_mhjh_username")
if name="" then Response.Redirect "../../error.asp?id=016"
set conn=server.createobject("adodb.connection") 
conn.open Application("yx8_mhjh_connstr")
set rs=server.CreateObject ("adodb.recordset")
sql="SELECT * FROM ��Ʒ WHERE ������='" & name & "' and ����>=100 and ����='��ʯ'  "
Set Rs=conn.Execute(sql)
if rs.eof or rs.bof then %>
<script language="vbscript">alert("����ʯ����100�������Բ��ܲ�����")
window.close()
</script><%
response.end
end if
sql="SELECT * FROM ��Ʒ WHERE ������='" & name & "' and ����>=100 and ����='��ˮ'  "
Set Rs=conn.Execute(sql)
if rs.eof or rs.bof then %>
<script language="vbscript">alert("����ˮ����100�������Բ��ܲ�����")
window.close()
</script><%
response.end
end if
sql="SELECT * FROM ��Ʒ WHERE ������='" & name & "' and ����>=40 and ����='�����'  "
Set Rs=conn.Execute(sql)
if rs.eof or rs.bof then %>
<script language="vbscript">alert("������㲻��40�������Բ��ܲ�����")
window.close()
</script><%
response.end
end if
sql="update ��Ʒ set ����=����-100 WHERE ������='" & name & "' and ����='��ʯ'  "
Set Rs=conn.Execute(sql)
sql="update ��Ʒ set ����=����-100 WHERE ������='" & name & "' and ����='��ˮ'  "
Set Rs=conn.Execute(sql)
sql="update ��Ʒ set ����=����-40 WHERE ������='" & name & "' and ����='�����'  "
Set Rs=conn.Execute(sql)
sql="select * from ��Ʒ where ����='ʥ�齣' and ������='" & name & "' and ����='����'"
Set Rs=conn.Execute(sql)
If Rs.Bof OR Rs.Eof then
			sql="insert into ��Ʒ(����,������,����,����,����,��Ч,����) values ('ʥ�齣','" & name & "','����',5000000,5000000,'�����',1)"
			conn.execute(sql)
                        else 
				sql="update ��Ʒ set ����=����+1 where ����='ʥ�齣' and ������='" & name & "'"
				conn.execute(sql)
end if
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
		newtalkarr(598)="000000" 
		newtalkarr(599)="<font color=red>������������</font><font color=blue>���򲻸������ˣ�"&name&"��ʥ�齣������ɣ�װ���������ӹ���500�򣬷���500�򣡿����������ֶ���һλ����,Ѫ���ȷ��������Ҫ������!</font>" 
newtalkarr(600)=Session("yx8_mhjh_userchatroomsn")
Application.Lock
Application("yx8_mhjh_talkpoint")=talkpoint+1
Application("yx8_mhjh_talkarr")=newtalkarr
Application.UnLock
erase newtalkarr
erase talkarr
%>
<%
set rs=nothing
conn.close
set conn=nothing
%>
<script Language="Javascript">alert("���򲻸������ˣ�����ʥ�齣������ɣ�װ���������ӹ���200�򣬷���300������Ժ��н����ˣ���")
parent.cz1.location.reload();
parent.ig.location.reload();
</script>