<!-- #include file="conn.asp"-->
<%
IF Session("aqjh_grade")<7 THEN
	response.write "<script>alert('��ʾ:�޸ĵ����Ϣ��Ҫ7�����Ϲ���!');window.close();</script>"
	response.end
END IF
%> 
<%
name = request.form("name")
zhufu = request.form("zhufu")
toname = request.form("toname")
songurl = request.form("songurl")
dim sql
set rs=server.createobject("adodb.recordset")
sql="select * from music where id="&request("id")
rs.open sql,conn,3,3
rs("name")=name
rs("zhufu")=zhufu
rs("toname")=toname
rs("songurl")=songurl
rs.update
rs.close
conn.close
response.write "<script Language=Javascript>alert('��ʾ���޸ĳɹ�!');window.close();</script>"
%>
