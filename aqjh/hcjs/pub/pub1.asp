<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../../error.asp?id=440"
time1=request.form("time")+0
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
if time1=0 then
mess="��ѡ����ȷ��ʱ�䣡"
elseif  time1<0 then
mess="��ѡ����ȷ��ʱ�䣡"
elseif  time1>24 then
mess="��ѡ����ȷ��ʱ�䣡"
else
Jname=aqjh_name
Jpass=request.cookies("pass")
yin=abs((time1)*10)
shi=0.0416*time1
rs.open "select * from �û� where ����='"&aqjh_name&"'" & "  and ״̬<>'��' and ״̬<>'��'",conn
if rs.eof or rs.bof then
mess=Jname & "���㲻�ܲ�����"
else
if rs("����")<yin then
mess=Jname & "���������������"
else
conn.execute "update �û� set ����=����-" & yin & ",����=����+"& yin &",����=����+"&yin&",��¼=now()+" & shi & ",״̬='��' where ����='"&aqjh_name&"'"
mess=Jname & "�����Ѿ��ɹ��ĵ���ջ��Ϣ�ˣ�����" &time1& "Сʱ��ʹ�ø��ʺţ�"
session.Abandon
%>
<script>
confirm('<%=mess%>')
top.menu.location.href="../../index.htm"
</script>
<%
end if
end if
rs.close
set rs=nothing
conn.close
set conn=nothing
end if
%>

<head>
<style>td           { font-size: 14 }
</style>
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title><%=Application("aqjh_chatroomname")%>��ջ</title></head>

<body bgcolor="#000000" background="../../jhimg/bk_hc3w.gif" text="#000000" link="#0000FF">
<table border="1" bgcolor="#BEE0FC" align="center" width="350" cellpadding="10"
cellspacing="13">
<tr>
<td bgcolor="#CCE8D6">
<table width="100%">
<tr>
<td>
<p align="center" style="font-size:14;color:red"><%=mess%></p>
<p align="center"><a href="pub.asp">����</a></p>
</td>
</tr>
</table>
</td>
</tr>
</table>

</body>