<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "select top 10 * from �û� where ״̬='����' order by -���",conn
x=0
%>
document.write ("<marquee scrollamount='1' scrolldelay='1' direction= 'up' width='130' height='160' id=paiheng onmouseover=paiheng.stop() onmouseout=paiheng.start()><div align=center>");
document.writeln("<font color=blue><b>����10����</b></font><br>");
<%do while not rs.eof%>
document.writeln("<font color=#333333><%=rs("����")%></font><br>");
<%
x=x+1
if x=10 then exit do
rs.movenext
loop
rs.close
rs.open "select top 10 ���� from �û� where ״̬='����' order by -(�书+����+����)",conn
x=0%>
document.writeln("<font color=blue><b>����10�����</b></font><br>");
<%do while not rs.eof%>
document.writeln("<font color=#333333><%=rs("����")%></font><br>");
<%x=x+1
if x=10 then exit do
rs.movenext
loop
rs.close
rs.open "select top 10 ���� from �û� where ����<>'' order by ����",conn
x=0%>
document.writeln("<font color=blue><b>����10�����</b></font><br>");
<%do while not rs.eof%>
document.writeln("<font color=#333333><%=rs("����")%></font><br>");
<%x=x+1
if x=10 then exit do
rs.movenext
loop
rs.close
rs.open "select top 10 ���� from �û� where ״̬='����' order by -mvalue",conn
x=0%>
document.writeln("<font color=blue><b>����10������</b></font><br>");
<%do while not rs.eof%>
document.writeln("<font color=#333333><%=rs("����")%></font><br>");
<%x=x+1
if x=10 then exit do
rs.movenext
loop
rs.close
rs.open "select top 10 ���� from �û� where ״̬='����' order by -������",conn
x=0%>
document.writeln("<font color=blue><b>����10��ɫħ</b></font><br>");
<%do while not rs.eof%>
document.writeln("<font color=#333333><%=rs("����")%></font><br>");
<%x=x+1
if x=10 then exit do
rs.movenext
loop
rs.close
rs.open "select top 10 ���� from �û� where �Ա�='��' and ��ż='��' and ������=0 order by -allvalue",conn
x=0%>
document.writeln("<font color=blue><b>����10���ͯ</b></font><br>");
<%do while not rs.eof%>
document.writeln("<font color=#333333><%=rs("����")%></font><br>");
<%x=x+1
if x=10 then exit do
rs.movenext
loop
rs.close
rs.open "select top 10 ���� from �û� where �Ա�='Ů' and ��ż='��' and ������=0 order by -allvalue",conn
x=0%>
document.writeln("<font color=blue><b>����10����Ů</b></font><br>");
<%do while not rs.eof%>
document.writeln("<font color=#333333><%=rs("����")%></font><br>");
<%x=x+1
if x=10 then exit do
rs.movenext
loop
rs.close
rs.open "select top 10 ���� from �û� where ״̬='����' and �Ա�='��' order by -(����+����)",conn
x=0%>
document.writeln("<font color=blue><b>����10��˧��</b></font><br>");
<%do while not rs.eof%>
document.writeln("<font color=#333333><%=rs("����")%></font><br>");
<%x=x+1
if x=10 then exit do
rs.movenext
loop
rs.close
rs.open "select top 10 ���� from �û� where ״̬='����' and �Ա�='Ů' order by -(����+����)",conn
x=0%>
document.writeln("<font color=blue><b>����10����Ů</b></font><br>");
<%do while not rs.eof%>
document.writeln("<font color=#333333><%=rs("����")%></font><br>");
<%x=x+1
if x=10 then exit do
rs.movenext
loop
rs.close
rs.open "select top 10 ���� from �û� where ״̬='����' order by -��ɱ��",conn
x=0%>
document.writeln("<font color=blue><b>����10��ɱ��</b></font><br>");
<%do while not rs.eof%>
document.writeln("<font color=#333333><%=rs("����")%></font><br>");
<%x=x+1
if x=10 then exit do
rs.movenext
loop
rs.close
rs.open "select top 10 ����,��ż from �û� where ״̬='����' and �Ա�='��' and ��ż<>'��' order by -��������",conn
x=0%>
document.writeln("<font color=blue><b>����10�Խ��</b></font><br>");
<%do while not rs.eof%>
document.writeln("<font color=#333333><%=rs("����")%></font><br>");
document.writeln("<font color=#333333>(<%=rs("��ż")%>)</font><br>");
<%
x=x+1
if x=10 then exit do
rs.movenext
loop
rs.close
rs.open "select top 10 a from p where b<>'δ��' order by - c",conn
x=0%>
document.writeln("<font color=blue><b>����10������</b></font><br>");
<%do while not rs.eof%>
document.writeln("<font color=#333333><%=rs("a")%></font><br>");
<%x=x+1
if x=10 then exit do
rs.movenext
loop
rs.close
rs.open "select top 10 a from p where b<>'δ��' order by - h",conn
x=0%>
document.writeln("<font color=blue><b>����10����</b></font><br>");
<%do while not rs.eof%>
document.writeln("<font color=#333333><%=rs("a")%></font><br>");
<%x=x+1
if x=10 then exit do
rs.movenext
loop
rs.close
set rs=nothing
conn.close
set conn=nothing
%>

document.write ("</div></marquee>")