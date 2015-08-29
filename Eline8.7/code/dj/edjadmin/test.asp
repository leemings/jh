<!--#include file="admin_top.asp"-->              <select name="classid" size="1">
<%
set Trs=server.createobject("adodb.recordset")
sql="select * from class"
Trs.open sql,conn,1,1
do while not rs.eof
%>
                <option value="<%=Trs("classID")%>" name=classid><%=Trs("class")%></option>
<%
Trs.movenext
loop
Trs.close
%>
              </select>