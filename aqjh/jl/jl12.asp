<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
dim page
page=request.querystring("page")
PageSize = 15
rs.open "delete * from l where a<now()-3",conn,3,3
rs.open "Select * From l where d='����' Order by a DESC",conn,3,3
rs.PageSize = PageSize
pgnum=rs.Pagecount
if page="" or clng(page)<1 then page=1
if clng(page) > pgnum then page=pgnum
if pgnum>0 then rs.AbsolutePage=page
%>
<html><head>
<title><%=Application("aqjh_chatroomname")%>-��������</title>
<LINK href="../css.css" rel=stylesheet type=text/css>
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
</head>
<body leftmargin="0" topmargin="0" background=../bg.gif oncontextmenu=window.event.returnValue=false>
<table border="0" height="24" width="91%" cellspacing="0" cellpadding="0"
align="center">
<tbody>
<tr>
<td height="15" width="100%"><font color="red"><b>��������</b></font></font></td>
</tr>
</tbody>
</table>
<div align="center">
<table width="91%" align="center" cellspacing="0" border="0"
cellpadding="0">
<tr>
<td width="100%">
<table width=670 border=1 bordercolorlight=000000 cellspacing=0 cellpadding=2 bordercolordark=FFFFFF align='center'>
<tr bgcolor="#FFFFFF">
<td width="11%">
<div align="center"><font color="#FF6600"> �� �� �� </font></div>
</td>
<td width="18%">
<div align="center"><font color="#FF6600"> ʱ �� </font></div>
</td>
<td width="12%">
<div align="center"><font color="#FF6600"> �������� </font></div>
</td>
<td width="59%">
<div align="center"><font color="#FF6600"> �� �� �� ��</font></div>
</td>
</tr>
<%
count=0
do while not (rs.eof or rs.bof) and count<rs.PageSize
%>
<tr bgcolor="#85C2E0" onmouseout="this.bgColor='#85C2E0';"onmouseover="this.bgColor='#DFEFFF';">
<td width="11%">
<div align="center"> <font color="#0000FF"><%=rs("b")%></font>
</div>
</td>
<td width="18%">
<div align="center"> <%=rs("a")%> </div>
</td>
<td width="12%">
<div align="center"> <%=rs("c")%> </div>
</td>
<td width="59%">
<div align="center"> <%=Replace(rs("e"),"red","blue")%> </div>
</td>
</tr>
<%rs.movenext%>
<%count=count+1%>
<%loop
pa=rs.pagecount
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
</table>
<table border="0" cellspacing="1" cellpadding="1" width="650" bordercolorlight="#EFEFEF" align=center>
<tr>
<td align="left" width="37%" height="2">[��<font color="red"><b><%=pa%></b></font>ҳ]
</td>
<td align="right" width="63%" height="2">
<div align="center">[<a href="jl12.asp?page=<%=page-1%>">��һҳ</a>][��<%=page%>ҳ][<a href="jl12.asp?page=<%=page+1%>">��һҳ</a>]</div>
</td>
</tr>
</table>
</table>
</div></body>