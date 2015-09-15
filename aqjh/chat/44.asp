<%@ LANGUAGE=VBScript%>
<%

Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.Expires = 0
id=trim(Request.QueryString ("id"))
mode=trim(Request.QueryString ("mode"))
chatbgcolor=Session("afa_chatbgcolor")
chatimage=Application("aqjh_chatimage")
if chatbgcolor="" then chatbgcolor="008888"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
nowinroom=session("nowinroom")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")

rs.Open ("Select count(*) As 数量 from 用户 where times<3 and CDate(lasttime)<now()-30"),conn
sn=rs(0)
rs.Close
s60=1000000
s61=800000
s62=600000
s63=400000
s64=200000
s65=100000
s66=50000

rs.Open ("select top 500 姓名 from 用户 where 身份='弟子'  and 等级<70"),conn
%>
<html>
<head>
<title>快乐江湖</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="../css.CSS">
</head>
<body oncontextmenu=self.event.returnValue=false bgcolor="<%=chatbgcolor%>" background="<%=chatimage%>" bgproperties="fixed" >
<div align="left">
<table width="147" border="1" cellpadding="3" cellspacing="0" bordercolorlight="#000000" bordercolordark="#FFFFFF" bgcolor="#FFFFFF" height="1">
<tr align="center"> 
<td width="93" height="1" valign="middle" >目前共有<br>
</td>
</tr><tr>  
<td width="93" height="1" valign="top" align="left" > <%
i=0
do while not rs.EOF
i=i+1
if i=1 then 
p="npc"
elseif i=2 then 
p="npc"
elseif i=3 then
p="npc"
elseif i=4 then
p="npc"
end if
%> <p><%=rs("姓名") %> (<a href=440.asp?id=<%=rs("姓名")%> target=ps><font color="#0000FF">上线</font></a><font color="#0000FF">)</font></p> 
 
 </tr>  
<tr>
<td width="93" height="1" > 
<%
if i=300 then exit do
rs.MoveNext
loop
%> 
 </td>
</tr>
</table>
</div>
</body>
</html>
<%
rs.Close
set rs=nothing
conn.Close
set conn=nothing
%>