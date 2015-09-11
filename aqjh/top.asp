<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select top 10 * from 用户 where 状态='正常' order by -存款",conn
x=0
%>
document.write ("<marquee scrollamount='1' scrolldelay='1' direction= 'up' width='130' height='200' id=paiheng onmouseover=paiheng.stop() onmouseout=paiheng.start()><div align=center>");
document.writeln("<font color=blue><b>江湖10大富翁</b></font><br>");
<%do while not rs.eof%>
document.writeln("<font color=#000000><%=rs("姓名")%></font><br>");
<%
x=x+1
if x=10 then exit do
rs.movenext
loop
rs.close
rs.open "select top 10 姓名 from 用户 where 状态='正常' order by -(武功+攻击+防御)",conn
x=0%>
document.writeln("<font color=blue><b>江湖10大高手</b></font><br>");
<%do while not rs.eof%>
document.writeln("<font color=#000000><%=rs("姓名")%></font><br>");
<%x=x+1
if x=10 then exit do
rs.movenext
loop
rs.close
rs.open "select top 10 姓名 from 用户 where 姓名<>'' order by 道德",conn
x=0%>
document.writeln("<font color=blue><b>江湖10大恶人</b></font><br>");
<%do while not rs.eof%>
document.writeln("<font color=#000000><%=rs("姓名")%></font><br>");
<%x=x+1
if x=10 then exit do
rs.movenext
loop
rs.close
rs.open "select top 10 姓名 from 用户 where 状态='正常' order by -mvalue",conn
x=0%>
document.writeln("<font color=blue><b>本月10大泡手</b></font><br>");
<%do while not rs.eof%>
document.writeln("<font color=#000000><%=rs("姓名")%></font><br>");
<%x=x+1
if x=10 then exit do
rs.movenext
loop
rs.close
rs.open "select top 10 姓名 from 用户 where 状态='正常' order by -结婚次数",conn
x=0%>
document.writeln("<font color=blue><b>江湖10大色魔</b></font><br>");
<%do while not rs.eof%>
document.writeln("<font color=#000000><%=rs("姓名")%></font><br>");
<%x=x+1
if x=10 then exit do
rs.movenext
loop
rs.close
rs.open "select top 10 姓名 from 用户 where 性别='男' and 配偶='无' and 结婚次数=0 order by -allvalue",conn
x=0%>
document.writeln("<font color=blue><b>江湖10大金童</b></font><br>");
<%do while not rs.eof%>
document.writeln("<font color=#000000><%=rs("姓名")%></font><br>");
<%x=x+1
if x=10 then exit do
rs.movenext
loop
rs.close
rs.open "select top 10 姓名 from 用户 where 性别='女' and 配偶='无' and 结婚次数=0 order by -allvalue",conn
x=0%>
document.writeln("<font color=blue><b>江湖10大玉女</b></font><br>");
<%do while not rs.eof%>
document.writeln("<font color=#000000><%=rs("姓名")%></font><br>");
<%x=x+1
if x=10 then exit do
rs.movenext
loop
rs.close
rs.open "select top 10 姓名 from 用户 where 状态='正常' and 性别='男' order by -(道德+魅力)",conn
x=0%>
document.writeln("<font color=blue><b>江湖10大帅男</b></font><br>");
<%do while not rs.eof%>
document.writeln("<font color=#000000><%=rs("姓名")%></font><br>");
<%x=x+1
if x=10 then exit do
rs.movenext
loop
rs.close
rs.open "select top 10 姓名 from 用户 where 状态='正常' and 性别='女' order by -(道德+魅力)",conn
x=0%>
document.writeln("<font color=blue><b>江湖10大美女</b></font><br>");
<%do while not rs.eof%>
document.writeln("<font color=#000000><%=rs("姓名")%></font><br>");
<%x=x+1
if x=10 then exit do
rs.movenext
loop
rs.close
rs.open "select top 10 姓名 from 用户 where 状态='正常' order by -总杀人",conn
x=0%>
document.writeln("<font color=blue><b>江湖10大杀手</b></font><br>");
<%do while not rs.eof%>
document.writeln("<font color=#000000><%=rs("姓名")%></font><br>");
<%x=x+1
if x=10 then exit do
rs.movenext
loop
rs.close
rs.open "select top 10 姓名,配偶 from 用户 where 状态='正常' and 性别='男' and 配偶<>'无' order by -结婚记念日",conn
x=0%>
document.writeln("<font color=blue><b>江湖10对金婚</b></font><br>");
<%do while not rs.eof%>
document.writeln("<font color=#000000><%=rs("姓名")%></font><br>");
document.writeln("<font color=#000000>(<%=rs("配偶")%>)</font><br>");
<%
x=x+1
if x=10 then exit do
rs.movenext
loop
rs.close
rs.open "select top 10 a from p where b<>'未定' order by - c",conn
x=0%>
document.writeln("<font color=blue><b>江湖10大门派</b></font><br>");
<%do while not rs.eof%>
document.writeln("<font color=#000000><%=rs("a")%></font><br>");
<%x=x+1
if x=10 then exit do
rs.movenext
loop
rs.close
rs.open "select top 10 a from p where b<>'未定' order by - h",conn
x=0%>
document.writeln("<font color=blue><b>江湖10大金库</b></font><br>");
<%do while not rs.eof%>
document.writeln("<font color=#000000><%=rs("a")%></font><br>");
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