<!--#include file="data1.asp"-->
<%set rs=conn.execute("select * from sheep where user='"&session("yx8_mhjh_username")&"'")
if not rs.eof then
rs.close
conn.execute("delete from sheep where user='"&session("yx8_mhjh_username")&"'")
conn.close
set conn=server.createobject("adodb.connection") 
set rst=server.CreateObject ("adodb.recordset")
conn.open Application("yx8_mhjh_connstr")
set rs=conn.execute("select * from 用户 where 姓名='"&session("yx8_mhjh_username")&"'")
money=rs("银两")+800
conn.execute("update 用户 set 银两='"&money&"'  where 姓名='"&session("yx8_mhjh_username")&"'")%>

<head>
<title></title>
</head>

<body bgcolor="#000000" text="#FFFFFF">

你忍痛卖出了你的孩子，但只得到800块钱！


<%else%>
你没有孩子不能卖出
<%rs.close%>
<%end if%>