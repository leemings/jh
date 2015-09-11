<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Buffer=true
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.Expires = 0
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
chatbgcolor=Application("aqjh_chatbgcolor")
chatimage=Application("aqjh_chatimage")
if chatbgcolor="" then chatbgcolor="008888"
%>
<html>
<head>
<title>基本命令</title>
<script language="JavaScript">
function s(list){parent.f2.document.af.sytemp.value=list;parent.f2.document.af.sytemp.focus();}
function rc(list){if(list!="0"){parent.f2.document.af.sytemp.value=list;parent.f2.document.af.sytemp.focus();}}
</script>
<style>
td{font-size:9pt}
</style>
</head>
<body bgcolor="<%=chatbgcolor%>" background="<%=chatimage%>">
<center>
<p style="font-size:12pt; color:red">江湖动作命令</p>
<%
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "SELECT * FROM d",conn%>
<table border="1" cellspacing="0" cellpadding="0" width="134" bordercolorlight="#000000" bordercolordark="#FFFFFF" bgcolor="#CCCCCC" align="center" >
 <tr bgcolor="#3399CC" onmouseout="this.bgColor='#3399CC';"onmouseover="this.bgColor='#DFEFFF';"> 
<td width=40%><div align="center">命令</div></td>
<td width=20%>自己</td>
<td width=20%>大家</td>
<td width=20%>别人</td>
</tr>
<%do while not rs.eof and not rs.bof
%>
  <tr bgcolor="#3399CC" onmouseout="this.bgColor='#3399CC';"onmouseover="this.bgColor='#DFEFFF';"> 
<td valign=top><div align="center"><%=rs("a")%></div></td>
<td valign=top><a href="javascript:s('//<%=replace(rs("b"),"##","")%>')" title="<%=rs("b")%>">自己</a></td>
<td valign=top><a href="javascript:s('//<%=replace(rs("c"),"##","")%>')" title="<%=rs("c")%>">大家</a></td>
<td valign=top><a href="javascript:s('//<%=replace(rs("d"),"##","")%>')" title="<%=rs("d")%>">别人</a></td>
</tr>
<%
rs.movenext
l=l+1
loop
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
</table>
</center>
</body>
</html>