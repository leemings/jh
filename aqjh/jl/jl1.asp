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
dim page
page=request.querystring("page")
PageSize = 15
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "delete * from l where a<now()-5",conn,3,3
rs.open "Select * From l where d='人命' Order by a DESC",conn,3,3
rs.PageSize = PageSize
pgnum=rs.Pagecount
if page="" or clng(page)<1 then page=1
if clng(page) > pgnum then page=pgnum
if pgnum>0 then rs.AbsolutePage=page
%>
<html>
<head>
<title><%=Application("aqjh_chatroomname")%>-江湖命案</title>
<LINK href="../css.css" rel=stylesheet type=text/css>
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
</head>
<body leftmargin="0" topmargin="0" background=../bg.gif oncontextmenu=window.event.returnValue=false>
<table border="0" height="24" width="91%" cellspacing="0" cellpadding="0" align="center">
<tbody>
<tr>
<td height="15" width="100%"><font color="#669966"> <font color="red"><b>江湖命案</b></font></font></td>
</tr>
</tbody>
</table>
<div align="center">
<table width="100%" align="center" cellspacing="0" border="0"
cellpadding="0">
<tr>
<td width="100%">
<table width=670 border=1 bordercolorlight=000000 cellspacing=0 cellpadding=2 bordercolordark=FFFFFF align='center'>
<tr bgcolor="#FFFFFF">
<td width="20%">
<div align="center"><font color="#FF6600"> 死 亡 者 </font></div>
</td>
<td width="25%">
<div align="center"><font color="#FF6600"> 时 间 </font></div>
</td>
<td width="19%">
<div align="center"><font color="#FF6600"> 杀 人 凶 手 </font></div>
</td>
<td width="36%">
<div align="center"><font color="#FF6600"> 死 亡 原 因 </font></div>
</td>
</tr>
<%
count=0
do while not (rs.eof or rs.bof) and count<rs.PageSize
%>
<tr bgcolor="#85C2E0" onmouseout="this.bgColor='#85C2E0';"onmouseover="this.bgColor='#DFEFFF';">
<td width="20%">
<div align="center"> <font color="#0000FF"><%=rs("b")%></font>
</div>
</td>
<td width="25%">
<div align="center"> <%=rs("a")%> </div>
</td>
<td width="19%">
<div align="center"> <%=rs("c")%> </div>
</td>
<td width="36%">
<div align="center"> <%=rs("e")%> </div>
</td>
</tr>
<%rs.movenext%>
<%count=count+1%>
<%loop
pa=rs.pagecount
sw=musers()
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
</table>
<table border="0" cellspacing="1" cellpadding="1" width="650" bordercolorlight="#EFEFEF" align=center>
<tr>
<td align="left" width="37%" height="2">[共<font color="red"><b><%=pa%></b></font>页][<font
color="red"><b><%=sw%></b></font>人死亡]</td>
<td align="right" width="63%" height="2">
<div align="center">[<a href="jl1.asp?page=<%=page-1%>">上一页</a>][第<%=page%>页][<a href="jl1.asp?page=<%=page+1%>">下一页</a>]</div>
</td>
</tr>
</table>
</table>
</div>
</body>
<%
function musers()
dim tmprs
tmprs=conn.execute("Select count(*) As 杀人 from l where d='人命'")
musers=tmprs("杀人")
set tmprs=nothing
if isnull(musers) then musers=0
end function
%>

