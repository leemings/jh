<%@ LANGUAGE=VBScript codepage ="936" %>
<%
response.expires=0
sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
if sjjh_name="" then Response.Redirect "../error.asp?id=440"
chatbgcolor=Application("sjjh_chatbgcolor")
chatimage=Application("sjjh_chatimage")
if chatbgcolor="" then chatbgcolor="008888"%><html>
<head>
<title>菜单♀wWw.happyjh.com♀</title>
<META http-equiv="Content-Type" content="text/html; charset=gb2312">
<style type='text/css'>
body{font-size:9pt;
CURSOR: url('48.ani');
scrollbar-track-color:#ffffff;
SCROLLBAR-FACE-COLOR:#effaff;
SCROLLBAR-HIGHLIGHT-COLOR:#ffffff;
SCROLLBAR-SHADOW-COLOR:#eeeeee;
SCROLLBAR-3DLIGHT-COLOR:#eeeeee;
SCROLLBAR-ARROW-COLOR:#dddddd;
SCROLLBAR-DARKSHADOW-COLOR:#ffffff}
td{font-size:9pt;}
input{font-size:9pt;}
a{font-size:9pt; color:black;text-decoration:none;}
a:hover{color:red;text-decoration:none;}
</style>
</head>
<%
const MaxPerPage=10
dim totalPut
dim CurrentPage
dim TotalPages
dim i,j
%>
<body bgcolor="#006699" leftmargin="0" topmargin="0" bgproperties="fixed" oncontextmenu=self.event.returnValue=false>
<table border="1" width="140" cellspacing="0" cellpadding="0" bgcolor="#000000" height="16" align="center" bordercolorlight="#000000" bordercolordark="#FFFFFF">
  <tr>
<td width="100%" height="28">
<div align="center"><font color="#CCCCFF" size="2"><strong>十大恶人</strong></font></div>
</td>
</tr>
</table>
<%
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
dim sql
dim rs
dim filename
rs.open "select top 10 * from 用户 where 道德<0 and 状态<>'无' and 门派<>'官府' order by 道德",conn,1,1
if rs.eof and rs.bof then
	response.write "<p align='center'>没有可排行的对象 </p>"
else
%>
<table border="1" cellspacing="0" width="140" bordercolorlight="#000000"
bordercolordark="#FFFFFF" cellpadding="4" align="center">
  <tr bgcolor="#336633">
    <td width="62" align="center" bgcolor="#000000"><font color="#FFFFFF">恶 人 
      谷</font></td>
    <td width="36" align="center" bgcolor="#000000"><font color="#FFFFFF">道 德</font></td>
</tr>
<%do while not rs.eof%>
<tr>
    <td align="center" bgcolor="#9999CC" width="62"><a href='../yamen/mt.asp?action=<%=rs("姓名")%>' target='_blank'><%=rs("姓名")%></a></td>
    <td align="center" bgcolor="#9999CC" width="36"><%=rs("道德")%> </td>
</tr>
<%
rs.movenext
filename=filename+1
if filename>19 then Exit Do
loop
end if
rs.Close
set rs=nothing
conn.close
set conn=nothing
%>
</table>
</body>
</html>
