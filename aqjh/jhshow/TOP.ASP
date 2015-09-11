<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("showmdb")
rs.open "select top 10 * from use  order by e desc",conn,1,1
x=0
%>
document.write ("<marquee scrollamount='1' scrolldelay='1' direction= 'up' width='120' height='150' id=paiheng onmouseover=paiheng.stop() onmouseout=paiheng.start()><div align=center>");
//document.writeln("<font color=000000><b>20大富翁</b></font><br>");
<%do while not rs.eof
x=x+1
%>
document.writeln("<%=x%>.<%=rs("a")%>&nbsp;&nbsp;<IMG alt=财富 src='images/qian.gif' border=0>&nbsp;<%=rs("e")%><br>");
<%
if x=10 then exit do
rs.movenext
loop
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
document.write ("</div></marquee>")