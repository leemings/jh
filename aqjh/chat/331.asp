<%@ LANGUAGE=VBScript codepage ="936" %>
<%
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.Expires = 0
id=trim(Request.QueryString ("id"))
mode=trim(Request.QueryString ("mode"))
chatbgcolor=Application("aqjh_chatbgcolor")
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



rs.Open ("select top 500 召唤兽1 from 用户 where 姓名='"&aqjh_name&"'"),conn
%>
<html>
<head>
<title>爱情江湖</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="../css.CSS">
</head>
<body oncontextmenu=self.event.returnValue=false bgcolor="<%=chatbgcolor%>" background="<%=chatimage%>" bgproperties="fixed" >
<div align="left">
<table width="159" border="1" cellpadding="3" cellspacing="0" bordercolorlight="#000000" bordercolordark="#FFFFFF" bgcolor="#FFFFFF" height="1">
<tr align="center"> 
<%if rs("召唤兽1")="无" then %>   
<td width="149" height="1" align="center" valign="middle"><a href="#" onClick="window.open('../yamen/shenshou.asp','','scrollbars=yes,resizable=yes,width=700,height=400')" title="建立自己的神兽"><font color="#0000FF">建立神兽</font></a></td>
<%end if%>
</tr>
<tr align="center"> 
<td width="149" height="1" valign="middle" >您的神兽</td>
</tr><tr>  
<td width="149" height="1" valign="top" align="center" > <%
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
%> <p><%if rs("召唤兽1")<>"无" then %><%=rs("召唤兽1") %> (<a href=attack3.asp?id=<%=rs("召唤兽1")%> target=ps><font color="#0000FF">释放</font></a><font color="#0000FF">)</font><%end if%></p>  
 
 </tr>  
<tr>
<td width="149" height="1" > 
<%
if i=150 then exit do
rs.MoveNext
loop
%> 
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
