<%@ LANGUAGE="VBSCRIPT" %>
<!--#include file="session.asp"-->
<!--#include file="level.asp"-->
<!--#include file="conn.asp"-->
<%
set rs=Server.CreateObject("ADODB.Recordset")
sql="Select * from admin"
rs.open sql,Conn,1,1
%>
<html>
<head>
<title>用户管理</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="../css/style.css" type="text/css">
<Script Language=Javascript>
function over(locate){
locate.style.border='1px solid black';locate.style.backgroundColor='#FFFFFF';locate.style.cursor='hand'
}
function out(locate){
locate.style.border='';locate.style.backgroundColor='#F7F7F7'
}
function del(ID){
if(confirm('您确定要删除吗？'))
 {
 location.href="del_user.asp?ID="+ID+""
 }
}
function edit(ID){
location.href="edit_user.asp?ID="+ID+""
}
</Script>
</head>

<body bgcolor="#FFFFFF" text="#000000" topmargin="0" background="bkg.gif">
<table width="90%" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr> 
          <td height="27"></td>
        </tr>
        <tr> 
          <td height="27">
            <table width="100%" border="0" cellspacing="1">
              <tr bgcolor="#5F8AC5"> 
                <td height="23" width="45" align="center"><font color="#FFFFFF">编号</font></td>
                <td height="21" align="center" width="80"><font color="#FFFFFF">用户等级</font></td>
                <td height="21" align="center" width="164"><font color="#FFFFFF">用户名称</font></td>
                <td height="21" align="center" width="165" bgcolor="#5F8AC5"><font color="#FFFFFF">用户密码</font></td>
                <td colspan="2" align="center"><font color="#FFFFFF">操作</font></td>
              </tr>
              <%
dim record,pgsz,total,page,rowcount
if rs.RecordCount=0 then
response.write"<tr><td class=font align=center><font color=#FF0000>暂无任何用户</font></td></tr>"
else
record=rs.recordcount
PgSz=10
rs.PageSize = Cint(PgSz) 
total=int(record/PgSz*-1)*-1
page=Request("page")
if page="" Then
page = 1
else
page=page+1
page=page-1
end if

if page<1 Then 
page=1
end if

rs.absolutePage = page

dim start
start=PgSz*Page-PgSz+1

sub showpage()
if page > 1 Then 
response.write "<a href=manage.asp?page="&page-1&"><img src=../images/page1.gif border=0 align=absmiddle></a>&nbsp;"
else
response.write ""
end if 
if RowCount = 0 and page <>Total then 
response.write "<a href=manage.asp?page="&page+1&"><img src=../images/page2.gif border=0 align=absmiddle></a>"
else
response.write ""
end if
response.write"&nbsp;&nbsp;共"&total&"页，每页显示"&PgSz&"个 共"&record&"个用户"
end sub
   rowCount = rs.PageSize
   do while Not rs.eof and rowcount>0
%>
              <tr> 
                <td height="21" bgcolor="#F9F9F9" align="center" class="small" width="45"><%=rs("ID")%></td>
                <td height="21" bgcolor="#F7F7F7" align="center" class="small"><%=rs("Level")%></td>
                <td height="21" bgcolor="#F9F9F9" align="center"><%=rs("username")%></td>
                <td height="21" bgcolor="#F9F9F9" align="center"><%=rs("password")%></td>
                <td width="33" align="center" bgcolor="#F7F7F7"><a href="javascript:edit(<%=rs("ID")%>)">编辑</a></td>
                <td width="33" align="center" bgcolor="#F7F7F7"><%if rs("id")="1" then%><font color=CCCCCC>删除</font><%else%><a href="javascript:del(<%=rs("ID")%>)">删除</a><%end if%></td>
              </tr>
              <%
   RowCount=RowCount-1
   rs.movenext
   loop

%>
              <tr> 
                <td align="right" height="24" colspan="6"> 
                  <%
showpage
end if
set rs=nothing
conn.close
set conn=nothing
%>
                </td>
              </tr>
            </table>
          </td>
        </tr>
      </table>
</body>
</html>