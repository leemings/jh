<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
name=request("name")
my=session("yx8_mhjh_username")
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
connstr=Application("yx8_mhjh_connstr")
conn.open connstr
sql="select ����¼IP from �û� where ����='" & my & "'"
set rs=conn.execute(sql)
ip=rs("����¼IP")
if rs.eof or rs.bof then
%><script language=vbscript>
MsgBox "�㲻�ǽ�������!"
location.href = "../exit.ASP"
</script><%
conn.close
response.end
else
sql="select * from �û� where ����='" & name & "'"
set rs=conn.execute(sql)
yin=rs("ɱ��")*20000
if rs("ɱ����")<>my or rs("ɱ����")="�ѽ᰸" then%>
<script language=vbscript>
MsgBox "�������ͨ������������ɱ�Ļ����Ѿ��᰸�ˡ�"
location.href = "tongji.asp"
</script>
<%else
if rs("����")=my then%>
<script language=vbscript>
MsgBox "�������Լ�ɱ�Լ�������û�и��"
location.href = "tongji.asp"
</script><%
else
if rs("����¼IP")=ip then
sql="update �û� set ɱ����='�Ⱥ򴦿�' where ����='" & name & "'"
rs=conn.execute(sql)
%>
<script language=vbscript>
MsgBox "������û�и����ɱ����������������Ա��أ��ǲ����뵽������ƭǮ��"
location.href = "tongji.asp"
</script><%
conn.close
response.end
else
sql="update �û� set ����=����+'" & yin & "' where ����='" & my & "'"
rs=conn.execute(sql)
sql="update �û� set ɱ����='�ѽ᰸',ɱ��=0 where ����='" & name & "'"
rs=conn.execute(sql)
mess=my & " ����ͨ���Ľ������<font color=ff0000> "& name &" </font>�����ˣ��쵽���ͽ�"&yin&"����"
talkarr=Application("yx8_mhjh_talkarr")
talkpoint=Application("yx8_mhjh_talkpoint")
dim newtalkarr(600)
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
newtalkarr(599)="<font color=red>���¼���</font><b><font color=red>" & mess & "</font></b>"
newtalkarr(600)=chatroomsn
Application.Lock
Application("yx8_mhjh_talkpoint")=talkpoint+1
Application("yx8_mhjh_talkarr")=newtalkarr
Application.UnLock
erase newtalkarr
erase talkarr
conn.close
Response.write ""&mess&"<b><p align=center><input type=button value=' �� �� ' onClick=javascript:window.close();></p>"
response.end
end if
end if
end if
rs.close
set rs=nothing
end if
%>



