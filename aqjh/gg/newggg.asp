<%
mdbsj="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("gg.asa")
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open mdbsj
rs.open "select top 5 * from ���� order by ʱ�� desc",conn,1,2
response.write "document.write('<table><tr><td width=270>');"
do while not(rs.eof or rs.bof)
	response.write "document.write('<FONT color=#b70000><B>��</B></FONT><a href=gg/disp.asp?id="&rs("id")&" target=_blank title="&rs("����")&">');"
	if len((rs("����")))>16 then
		response.write "document.write('"&left(rs("����"),16)&"...');"
	else
		response.write "document.write('"&rs("����")&"');"
	end if
	response.write "document.write('</a></td></tr><tr><td width=270>');"
	'<td width=15>"&rs("ʱ��")&"</td>
	rs.movenext
loop
response.write "document.write('</td></tr><tr><td width=15></td></tr><tr><td colspan=2 align=right></a></td></tr></table>');"
rs.close
set rs=nothing
conn.close
set conn=nothing
%>