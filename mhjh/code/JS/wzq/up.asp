<%
Set cn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
datestr=Application("yx8_mhjh_connstr")
cn.open datestr
sql="delete * from game where username<>'нч'"
cn.Close
set cn=nothing
%>
