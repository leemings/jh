<!--#include file="data1.asp"-->
<%set rs=conn.execute("select * from sheep where user='"&session("yx8_mhjh_username")&"'")
if not rs.eof then
rs.close
conn.execute("delete from sheep where user='"&session("yx8_mhjh_username")&"'")
conn.close
set conn=server.createobject("adodb.connection") 
set rst=server.CreateObject ("adodb.recordset")
conn.open Application("yx8_mhjh_connstr")
set rs=conn.execute("select * from �û� where ����='"&session("yx8_mhjh_username")&"'")
money=rs("����")+800
conn.execute("update �û� set ����='"&money&"'  where ����='"&session("yx8_mhjh_username")&"'")%>

<head>
<title></title>
</head>

<body bgcolor="#000000" text="#FFFFFF">

����ʹ��������ĺ��ӣ���ֻ�õ�800��Ǯ��


<%else%>
��û�к��Ӳ�������
<%rs.close%>
<%end if%>