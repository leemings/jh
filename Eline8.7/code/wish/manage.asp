<%@ LANGUAGE=VBScript codepage ="936" %><!--#include file="config.asp"-->
<!--#include file="conn.asp"-->
<%
if session("admin")="ture" then

Set rs=Server.CreateObject("ADODB.Recordset")
sql="select * from wish order by id desc"
rs.open sql,conn,1,1
if rs.eof or rs.bof then
 response.write "<br><br><center><font color=red>��û���κ���Ը���Թ���!</font></center>"
else
%>
<html>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<head>
<style type="text/css">
<!--
a {  text-decoration: none}  
a:hover {  text-decoration: underline} 
table {  font-size: 9pt}
body,table,p,td,input {  font-size: 9pt} 
-->
</style>
<title><%=title%></title>
</head>
<body bgcolor="<%=bgcolor%>" text="<%=textcolor%>" link="<%=linkcolor%>">

<div align="center"><center>

<table border="1" cellspacing="1" width="550" bordercolor="<%=titlelightcolor%>">
  <tr>
    <td width="668" height="15" bgcolor="<%=titledarkcolor%>" colspan="6"><p align="center"><font
    color="<%=titlelightcolor%>">ɾ���û����� (����<font color="<%=linkcolor%>"><b><%=rs.recordcount%></b></font>����Ը)</td>
  </tr>
  <tr>
    <td width="108" height="12" bgcolor="<%=titledarkcolor%>" align="center"><font color="<%=titlelightcolor%>">�û���</font></td>
    <td width="40" height="12" bgcolor="<%=titledarkcolor%>" align="center"><font color="<%=titlelightcolor%>">�Ա�</font></td>
    <td width="80" height="12" bgcolor="<%=titledarkcolor%>" align="center"><font color="<%=titlelightcolor%>">Ը�����</font></td>
    <td width="106" height="12" bgcolor="<%=titledarkcolor%>" align="center"><font color="<%=titlelightcolor%>">�Ǽ�IP</font></td>
    <td width="133" height="12" bgcolor="<%=titledarkcolor%>" align="center"><font color="<%=titlelightcolor%>">�Ǽ�ʱ��</font></td>
    <td width="46" height="12" bgcolor="<%=titledarkcolor%>" align="center"><font color="<%=titlelightcolor%>">ɾ��</font></td>
  </tr>
<%
    page=request.querystring("page")
    pmcount=10    'ÿҳ��ʾ��¼��

    rs.pagesize=pmcount
    maxpage=rs.pagecount
    total=rs.recordcount

    if isempty(page) or cint(page)<1 or cint(page)>maxpage then
        page=1
    end if

    rs.absolutepage=page

rc=rs.pagesize
re=1
do while not rs.eof and rc>0
%>
  <tr>
    <td align="center"><%=rs("name")%></td>
    <td align="center"><%=rs("sex")%></td>
    <td align="center"><%=rs("wishtype")%></td>
    <td align="center"><%=rs("ip")%></td>
    <td align="center"><%=rs("date")%></td>
    <td align="center"><a href=del.asp?did=<%=rs("id")%>>ɾ��</a></td>
  </tr>
<%
rs.movenext
loop
%>
</table>
</center>
</div>

<p align="center">��ҳ�� <% for num=1 to maxpage
if num=cint(page) then%>
      <%=num%> 
      <%else%>
      <b><a href='admin.asp?page=<%=num%>'><%=num%></a></b> 
      <%end if%>
      <%next%>��</p>
<div align="center"><hr width=550 color=<%=titlelightcolor%> size=1></div>
<!--#include file="copyright.asp"-->
</body>
</html>
<%end if
else
    response.redirect "error.asp?msg=��ûȨ���鿴��ҳ"
end if
%>