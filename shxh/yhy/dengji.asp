<%
Response.Buffer=true
username=session("Ba_jxqy_username")
grade=session("�ȼ�")
if username="" then Response.Redirect "error.asp?id=016"

Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
connstr=Application("Ba_jxqy_connstr")
conn.open connstr
sql="SELECT * FROM �û� WHERE ����='"&username&"'"
set rs=conn.execute(sql)
sex=rs("�Ա�")
meimao=rs("����")
if sex<>"Ů" then 
response.write "����û�и��ѽ������Ժ��ɲ�Ҫ�е�Ŷ��"
response.write "<br><a href=../myhome.asp>����</a>"
response.end
end if
if meimao<1000 then 
response.write "����û�и��ѽ���㳤����ô�����������ĵ���Ҫ����1000�ſ��Եġ�"
response.write "<br><a href=../myhome.asp>����</a>"
response.end
conn.colse
set rs=nothing
end if%>
<!--#include file="jiu.asp"-->
<% 
sql="select * from ���� where ����='" & username & "'"
set rs=connt.execute(sql)
if rs.bof or rs.eof then
sql="insert into ����(����,��ò��) values ('" & username & "'," & meimao & ")"
			connt.execute sql
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
connstr=Application("Ba_jxqy_connstr")
conn.open connstr
sql="SELECT * FROM �û� WHERE ����='"&username&"'"
set rs=conn.execute(sql)
sql="update �û� set ����=����+12000 where ����='"& username &"'"
set rs=conn.execute(sql)
response.write "��ϲ������ʽ��Ϊ��Ժ�Ĺ���ȵ���������12000��"
response.end
response.write "<br><a href=index.asp>����</a>"
else
response.write "���Ѿ��Ǳ�����Ժ�Ĺ����ˣ���ô�����Ǽ�ѽ��"
response.write "<br><a href=../myhome.asp>����</a>"
response.end
connt.close
conn.close
set rs=nothing
  end if
%>