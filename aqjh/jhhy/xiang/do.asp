<%
username=session("aqjh_name")
set conn=server.createobject("adodb.connection")
conn.open Application("aqjh_usermdb")
set rst=server.CreateObject ("adodb.recordset")
sql="SELECT * FROM �û� where ����='" &username& "'"
Set Rs=conn.Execute(sql)
if rs.bof or rs.eof then
response.write "�㲻�ǽ������˻������ӳ�ʱ"
conn.close
response.end
else
hy=rs("��Ա")
hydj=rs("��Ա�ȼ�")
tl=rs("����")
nl=rs("����")
gj=rs("����")
fy=rs("����")
if username="" then
%>
<script language=vbscript>
MsgBox "�Բ����㻹û�е�¼"
location.href ="../../exit.asp"
</script>
<%
else
if Application("aqjh_hy")=true then
%>
<script language=vbscript>
MsgBox "���󣡻�Ա������δ���ţ�"
location.href ="javascript:history.back()"
</script>
<%
else
if hy=false and hydj=0 then
%>
<script language=vbscript>
MsgBox "�����㲻�ǻ�Ա����ذɣ�"
location.href ="javascript:history.back()"
</script>
<%
else
sj=DateDiff("n",rs("����ʱ��"),now())
if sj<10 then
ss=10-sj
rs.close
set rs=nothing
conn.close
set conn=nothing
Response.Write "<script language=JavaScript>{alert('��ող�������ô����������ȣ�"&ss&"�ֺ������ɣ�');location.href = 'javascript:history.back()';}</script>"
Response.End 
end if
jiu=request.form("jiu")
Jname=session("aqjh_name")
select case jiu
case "lao"
tili=40
yin=20000
case "wu"
tili=60
yin=30000
case "du"
tili=80
yin=40000
case "mao"
tili=100
yin=50000
end select
%>
<!--#include file="data.asp"-->
<%
sql="select * from �û� where ����='" & Jname & "' and ״̬<>'����' and ״̬<>'����' and ״̬<>'����'"
set rs=conn.execute(sql)
if rs.eof or rs.bof then
mess=Jname & "���㲻�ܲ�����"
else
if rs("����")<yin then
mess=Jname & "���������������"
else
daode=rs("����")
jiuliang=int((daode)/100)
if jiuliang>1 then
sql="update �û� set ����=����-" & yin & ",����=����+"& tili &" where ����='" & Jname & "'"
conn.execute sql
mess=Jname & "������³����󣬾����Լ������ǳ��ͬʱҲ������ٵ���!"
else 
sql="update �û� set ����=����-" & yin & ",����=����+"& tili &",��¼=now(),״̬='��' where ����='" & Jname & "'"
conn.execute sql
mess=Jname & "�����������������̸��Ͷ�����������������ڳ��˷����ߣ�����1Сʱ��ʹ�ø��ʺţ�"
response.cookies("Jname")=""
response.cookies("Jpai")=""
conn.close
%>
<script>
confirm('<%=mess%>')
top.menu.location.href="../../index.asp target=_top"
</script>
<%
end if
end if
end if
%>
<head>
<link rel="stylesheet" href="../../css.css" type="text/css">
</head>
<body bgcolor="#000000">
<table border="1" bgcolor="#00A200" align="center" width="350" cellpadding="10"
cellspacing="13">
<tr>
<td bgcolor="#005b00">
<table width="100%">
<tr>
<td>
<p align="center" style="font-size:14;color:red"><%=mess%></p>
<p align="center"><a href="index.asp">����</a></p>
</td>
</tr>
</table>
</td>
</tr>
</table>
</body>
<%
end if
end if
end if
end if
%>