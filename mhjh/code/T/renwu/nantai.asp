<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
username=session("yx8_mhjh_username")
sj=DateDiff("s",Application("yx8_mhjh_dg"),now()) 
if sj<5 then 
s=5-sj 
Response.write "<script language='javascript'>alert ('���ǽ���֮�񣬷�ֹˢǮ����������������ˢ\n�������õģ���������["&s&"����]�ٲ�����');setTimeout('history.back();',1000);</script>"
Response.End 
end if 
Application.Lock 
Application("yx8_mhjh_dg")=now() 
Application.UnLock 
set conn=server.createobject("adodb.connection")
conn.open Application("yx8_mhjh_connstr")
set rst=server.CreateObject ("adodb.recordset")
sql="SELECT * FROM �û� where ����='" &username& "'"
Set Rs=conn.Execute(sql)
if rs.bof or rs.eof then
response.write "�㲻�ǽ������˻������ӳ�ʱ"
conn.close
response.end
else
mymoney=rs("����")
dj=rs("�ȼ�")
rs.close
if username="" then
%>
<script language=vbscript>
MsgBox "�Բ����㻹û�е�¼"
location.href = "../index.asp"
</script>
<%
else
if mymoney="" then
%>
<script language=vbscript>
MsgBox "����ˣ��㻹û�е������峤�����������������ʲô��"
location.href = "javascript:history.back()"
</script>
<%
else
if dj<10 then
%>
<script language=vbscript>
MsgBox "������Ŀǰ�ĵȼ�����10,���������������ĵȴ������ɡ���"
location.href = "javascript:history.back()"
</script>
<%
else
randomize timer
r=int(rnd*10)

if r<=2 then
randomize timer
r=int(rnd*50)
s=1
nl=0
sm=0
nu=int(rnd*18)+1
s=int(rnd*100)
mess="��̫��ɽ��Ұ�ǹ�Ȼ�������㱻Ⱥ�ǹ����������ˡ���Ĺ����ͷ�����˼�����"& s &""
sql="update �û� set ����=����-'"& s &"',����=����-'"& s &"'  where ����='" &username& "'"
conn.execute sql
conn.close
elseif r=3 then
randomize timer
r=int(rnd*50)
s=1
nl=0
sm=0
nu=int(rnd*18)+1
s=int(rnd*100)
mess="������̫��ɽ������ɽկ��Ӣ�����Ӷ�ȥ������"& s &"��"
sql="update �û� set ����=����-'"& s &"'  where ����='" &username& "'"
conn.execute sql
conn.close
elseif r=5 or r=6 then
mess="��̫��ɽ�յ����һ����������һ����ԭ��������壬����͵ѧ�˼��У�����������20"
sql="update �û� set ����=����+20  where ����='" &username& "'"
conn.execute sql
conn.close
elseif r=7 or r=8 then
mess="������Ⱥ��ս������Ȼ������ʧ10000������õ�����Ƥһ�ţ��쵽�����������ܻ���ɽ����ɣ�"
sql="update �û� set ����='��Ƥ',����=����-10000  where ����='" &username& "'"
conn.execute sql
conn.close
else
mess="��ϲ��!��������̫��ɽɽկ��Ȼû�������κ��ˣ�ֱ�Ӵ�ɽկͷ�ӱ����ϵõ��˷����壬���Ե����ݳ�����ѿ��������ǹ�ˣ�"
sql="update �û� set ����='������'  where ����='" &username& "'"
conn.execute sql
conn.close
end if
set conn=nothing
%>
<script language=vbscript>
MsgBox "<%=mess%>"
location.href = "taihang.asp"
</script>

<%end if
end if
end if
end if%>
