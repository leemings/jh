<!--#include file="conn.asp"-->
<%if session("aqjh_name")="" then
response.redirect "../login.asp"
response.end
else
sql= "select * from ��Ʊ "
set rss=conn.execute(sql)
if rss("״̬")="��" then 
DO while not RSS.EOF   
sql="update ��Ʊ set ״̬='��' ,������=0 "
conn.execute sql
rss.movenext
loop

end if


%>
<html>
<head>
<title>������</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>

<body bgcolor="#FFFEF4" text="#000000">
<div align="center">
  <p>�Բ��𣬹�Ʊ�г��Ѿ�����,����ÿ��� ���� 5:00���У� 24:00 ������ </p>
  <p><OBJECT  id=closes  type="application/x-oleobject"  classid="clsid:adb880a6-d8ff-11cf-9377-00aa003b7a11" width="14" height="14">   
<param  name="Command"  value="Close">   
</object> <a href="#" onclick="closes.Click()">�رձ�ҳ </a>
</p>
</div>
</body>
</html>
<%
end if
CloseDatabase		'�ر����ݿ�

%>