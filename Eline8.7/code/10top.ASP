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
<title>菜单-::一线网络www.happyjh.com::</title>
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
<table border="1" width="140" cellspacing="0" cellpadding="0" bgcolor="#006699" height="16" align="center" bordercolorlight="#000000" bordercolordark="#FFFFFF">
  <tr>
<td width="100%" height="28">
<div align="center"><font color="#CCCCFF" size="2"><strong>十 大 排 行</strong></font></div>
</td>
</tr>
</table>
<table width="140" height="350" border="1" align="center" cellpadding="1" cellspacing="0" bordercolorlight="#000000"
bordercolordark="#FFFFFF" background="bg2.gif">
  <tr bgcolor="#336633"> 
    <td align="left" valign="top" background="bg2.gif" bgcolor="#efefef"> 
      <div align="left"><img src="blank.gif" width="1" height="1"><script language=javascript src="top.asp"></script></div></td>
  </tr>
</table>
</body>
</html>
