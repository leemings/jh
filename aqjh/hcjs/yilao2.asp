<%@ LANGUAGE=VBScript codepage ="936" %>
<SCRIPT LANGUAGE=JavaScript>if(window.name!='aqjh_win'){var i=1;while(i<=50){window.alert('ˢǮ����ϲ�����𣿵㰡��ˢ������');i=i+1;}top.location.href='../../exit.asp'}</script>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
name=aqjh_name
yilao=request.form("yilao")
if yilao<>"��" then
'У���û�
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "SELECT * FROM �û� WHERE ����='"&aqjh_name&"'" &" and ����<1000",conn
sj=DateDiff("n",rs("����ʱ��"),now())
if sj<10 then
	s=10-sj
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('�Բ���ϵͳ���ƣ����["&s&"����]�ٲ�����');}</script>"
	Response.End
end if	

If Rs.Bof OR Rs.Eof Then
mess="�㲻�ܽ�������"
else
sm=rs("����")
yl=rs("����")
select case yilao
case "��Щ��Ʒ"
bl=1.5
cd=1000-sm
case "һ������"
bl=1
cd=1000-sm
case "��������"
bl=0.5
cd=1000-sm
end select
fy=int(cd/bl)
sm=1000
if yl<fy then
fy=yl
sm=yl*bl
end if
conn.execute "update �û� set ����=����+" & sm & ",����=����-" & fy & ",����ʱ��=now() where ����='"&aqjh_name&"'"
mess="������ҽ�ɵ����ƣ��㻨��" & fy & "����������������" & sm & "��"
end if
conn.close
set conn=nothing
else
mess="�㲻������"
end if
%>

<head>
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
</head>

<body bgcolor="#000000">

<table border="1" bgcolor="#BEE0FC" align="center" width="350" cellpadding="10"
cellspacing="13">
<tr>
<td bgcolor="#CCE8D6">
<table width="100%">
<tr>
<td>
<p align="center" style="font-size:14;color:red"><%=mess%></p>
<p align="center"><a href="YILAO.ASP">����</a></p>
</td>
</tr>
</table>
</td>
</tr>
</table>

</body>
