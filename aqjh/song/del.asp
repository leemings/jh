<!-- #include file="conn.asp"-->
<%
 set rs=server.createobject("adodb.recordset")
   sql="delete from music where id="&request("ID")
   rs.open sql,conn,1,1
   rs.close
response.write "<script Language=Javascript>alert('��ʾ��ɾ���ɹ�!');location.href = '.';</script>"
%>