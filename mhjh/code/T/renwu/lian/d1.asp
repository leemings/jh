<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
name=session("yx8_mhjh_username")
if name="" then Response.Redirect "../../error.asp?id=016"
set conn=server.createobject("adodb.connection") 
conn.open Application("yx8_mhjh_connstr")
set rs=server.CreateObject ("adodb.recordset")
sql="SELECT * FROM ��Ʒ WHERE ������='" & name & "' and ����>=3 and ����='�𻨾�'  "
Set Rs=conn.Execute(sql)
if rs.eof or rs.bof then %>
<script language="vbscript">alert("���𻨾Ʋ���3ƿ�����Բ��ܲ�����")
window.close()
</script><%
response.end
end if
sql="update ��Ʒ set ����=����-3 WHERE ������='" & name & "' and ����='�𻨾�'  "
Set Rs=conn.Execute(sql)
sql="select * from ��Ʒ where ����='������' and ������='" & name & "' and ����='����'"
Set Rs=conn.Execute(sql)
If Rs.Bof OR Rs.Eof then
sql="insert into ��Ʒ(����,������,����,����,��Ч,����,����,����,����,����,װ��) values ('������','" & name &"',1,'����','��',0,10000,2000,0,False,False)"
conn.execute sql
else
sql="update ��Ʒ set ����=����+1 where ����='������' and ������='"&name&"'"
conn.execute sql
end if
set rs=nothing
conn.close
set conn=nothing
%>
<script Language="Javascript">alert("������һ�Ѻ��������������յĺ�������")
parent.cz1.location.reload();
parent.ig.location.reload();
</script>