<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.Buffer=true
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
if session("aqjh_adminok")<>true then Response.Redirect "../chat/readonly/bomb.htm"
if aqjh_grade<>10 or instr(Application("aqjh_admin"),aqjh_name)=0  then Response.Redirect "../error.asp?id=439"
id=request("id")
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select * from �� where id="&id,conn
tp=rs("ͼƬ")
rs.close
conn.execute "delete * from �� where id="&id
conn.execute "update �û� set ��='��' where ��='" & tp & "'"
set rs=nothing
conn.close
set conn=nothing
%>
<html>
<head>
<title>�ޱ����ĵ�</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>

<body bgcolor="#FFFFFF" text="#000000" background="../JHIMG/bk_Hc3w.gif">
<p align="center">ɾ����ͨ���߳ɹ�</p>
<p align="center"><a href="manjtgj.asp">����</a> </p>
</body>
</html>
