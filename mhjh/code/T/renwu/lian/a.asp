<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
username=session("yx8_mhjh_username")
sj=DateDiff("s",Application("yx8_mhjh_dg"),now()) 
if sj<3 then 
s=3-sj 
Response.write "<script language='javascript'>alert ('����ħ��֮�񣬷�ֹˢǮ����������������ˢ\n�������õģ���������["&s&"����]�ٲ����������\n��Ȱ�棬�㽫�˷ѵ�һ�����壡��');setTimeout('history.back();',1000);</script>"
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
rwtm=rs("����ʱ��")
dj=rs("����")
yijuan=rs("����")
rs.close
if now()-rwtm>10 or now()-rwtm<-10 then
%>
<script language=vbscript>
MsgBox "�Բ��𣬷��幦��ʧЧ,���뿪��"
location.href = "javascript:self.close()"
</script>
<%
else
if dj<>"ɽ��Ѱ��" then
%>
<script language=vbscript>
MsgBox "����ˣ����������" & dj & "��������ʲô��"
location.href ="javascript:self.close()"
</script>
<%else
if username="" then%>
<script language=vbscript>
MsgBox "�Բ����㻹û�е�¼"
location.href = "../index.asp"
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
mess="����������һ��ľ׮��û��ʲô�ջ񣬷����˷��˲���������������˼�����"& s &""
sql="update �û� set ����=����-'"& s &"'  where ����='" &username& "'"
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
mess="���ľ׮�����Ʋ��ԣ�������ʦָ�㣬�������������"& s &"��"
sql="update �û� set ����=����+'"& s &"' where ����='" &username& "'"
conn.execute sql
conn.close
elseif r=5 or r=6 then
mess="����ʱ����ר��һ�£��������ʦѵ��һͨ����˼�������ȥ�ˡ�������˼���200"
sql="update �û� set ����=����-200 where ����='" &username& "'"
conn.execute sql
conn.close
elseif r=7 or r=8 then
mess="������������һ��ľ׮���������ľ׮������п�������һ�飬��������ʱ��ʦ�������ˣ�����ֻ�еȵ��Ժ���˵�ˡ�������������������500"
sql="update �û� set ����=����-500,����=����-500 where ����='" &username& "'"
conn.execute sql
conn.close
else
mess="��ϲ"&username&"�ҵ�һ����Թ��������8000������1000"
sql="select * from ��Ʒ where ����='��Թ��' and ������='" & username & "'"
			set rs=conn.execute(sql)
			if rs.eof and rs.bof then
			sql="insert into ��Ʒ(����,������,����,����,����,����) values ('��Թ��','" & username & "','����',8000,1000,1)"
			conn.execute(sql)
                        else 
				sql="update ��Ʒ set ����=����+1 where ����='��Թ��' and ������='" & username & "'"
				conn.execute(sql)
		        end if
                        rs.close
                        set rs=nothing
end if
set conn=nothing
%>
<script language=vbscript>
MsgBox "<%=mess%>"
location.href = "d2.asp"
</script>

<%end if
end if
end if
end if%>
