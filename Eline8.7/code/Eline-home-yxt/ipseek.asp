<%@ LANGUAGE=VBScript codepage ="936" %><%Response.Expires=0
Response.Buffer=true

sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
if sjjh_name="" then Response.Redirect "../error.asp?id=440"
if session("sjjh_adminok")<>true then Response.Redirect "../chat/readonly/bomb.htm"
if sjjh_grade<>10 or instr(Application("sjjh_admin"),sjjh_name)=0  then Response.Redirect "../error.asp?id=439"
ipdizi=request.form("ipdizi")
if ipdizi="" then%>
<html>
<head>
<title><%=Application("sjjh_chatroomname")%></title>
<style></style>
<link rel="stylesheet" href="../chat/READONLY/Style.css">
</head>
<body bgcolor=#FFFFFF background="../JHIMG/Bk_hc3w.gif">
<center>
<font color="#000000"><b><font size="+2">
<%ip=Request.ServerVariables("REMOTE_ADDR")
sip=split(ip,".")
num=cint(sip(0))*256*256*256+cint(sip(1))*256*256+cint(sip(2))*256+cint(sip(3))-1
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
'�����ip�ĵ���
sql="select * FROM z WHERE a<="& num &" and b>="&num
Set Rs=conn.Execute(sql)
if rs.eof or rs.bof then
country="����"
city="δ֪"
else
country=rs("c")
city=rs("d")
end if
if country="" then country="�й�"
if city="" then city="δ֪"
last="����:"& country &" ����:"& city
set rs=nothing	
set conn=nothing
%>
<%=Application("sjjh_chatroomname")%>ip���ҳ���</font></b></font><br>
<br>
��л׷�������ṩ��ip���ݿ����(11.1��ǰ)<br>
ע���ó�����������ip��ַ��<br>
����ip��ַΪ��<%=ip%> <%=last%>
<form action=ipseek.asp method=post>
������ip��ַ���س�:
<input  size=20 name=ipdizi><input type=submit value='ȷ��'>
</form>
</center>
</body>
</html>
<%else
ip=ipdizi
sip=split(ip,".")
num=cint(sip(0))*256*256*256+cint(sip(1))*256*256+cint(sip(2))*256+cint(sip(3))-1
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
'�����ip�ĵ���
sql="select top 1 c,d FROM z WHERE a<="& num &" and b>="&num
Set Rs=conn.Execute(sql)
if rs.eof or rs.bof then
country="����"
city="δ֪"
else
country=rs("c")
city=rs("d")
end if
if country="" then country="�й�"
if city="" then city="δ֪"
last="����:"& country &" ����:"& city
set rs=nothing	
set conn=nothing%>
<html>
<head>
<title><%=Application("sjjh_chatroomname")%></title>
<style></style>
<link rel="stylesheet" href="../chat/READONLY/Style.css">
</head>
<body bgcolor=#FFFFFF background="../JHIMG/Bk_hc3w.gif">
�������ҵ�ip��ַΪ��<%=ip%><br>����Ϊ��<%=last%>
<br>
<br>
<a href="ipseek.asp">����</a>
</html>
<%end if%>