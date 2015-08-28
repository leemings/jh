<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
username=session("yx8_mhjh_username")
if username="" then response.redirect "error.asp?id=016"
mypai=session("yx8_mhjh_usercorp")
if mypai="" or mypai="无" or mypai="出家" then
response.write "你还没有加入任何门派！"
response.end
end if
set conn=server.createobject("adodb.connection")
conn.open Application("yx8_mhjh_connstr")
set rst=server.CreateObject ("adodb.recordset")
sql="SELECT * FROM 用户 WHERE 门派='"&mypai&"'"
Set Rs=conn.Execute(sql)
sql="Select count(*) from 用户 where 门派='"&mypai&"'"
set rst=conn.execute(sql)
renshu=rst(0)
%>

<head>
<title>本门兄弟</title>
<link rel="stylesheet" href="../STYLE.CSS">
</head>
<body background='../chatroom/bg1.gif' oncontextmenu=self.event.returnValue=false>
<p align="center"><b><font color="#FF3333" size="4" >本门兄弟</font></b></p>
<table width="99%" border="1" cellspacing="0" cellpadding="2" align="center" bordercolorlight="#000000" bordercolordark="#FFFFFF">
<tr align="center">
<td  width="110" bgcolor="#FFFFFF"> 【<%=mypai%>】[<%=renshu%>]人</td>
<td  width="61" bgcolor="#FFFFFF">性 别</td> 
<td  width="61" bgcolor="#FFFFFF"> 攻击</td> 
<td  width="61" bgcolor="#FFFFFF"> 防 御</td> 
<td  width="49" bgcolor="#FFFFFF"> 体 力</td> 
<td  width="56" bgcolor="#FFFFFF">内 力</td> 
<td  width="61" bgcolor="#FFFFFF">身 份</td> 
<td  width="49" bgcolor="#FFFFFF">积 分</td> 
<td  width="63" bgcolor="#FFFFFF">师 傅</td> 
</tr> 
<tr> <%do while not rs.bof and not rs.eof%> 
<td width="110"><%=rs("姓名")%></td> 
<td align="center" width="61"><%=rs("性别")%></td> 
<td align="center" width="61"><%=rs("攻击")%></td> 
<td align="center" width="61"><%=rs("防御")%></td> 
<td align="center" width="49"><%=rs("体力")%></td> 
<td align="center" width="56"><%=rs("内力")%></td> 
<td align="center" width="61"><%=rs("身份")%></td> 
<td align="center" width="49"><%=rs("积分")%></td> 
<td align="center" width="63"><%=rs("师傅")%></td> 
</tr> 
<% 
rs.movenext 
loop 
rs.Close 
set rs=nothing 
rst.Close 
set rst=nothing 
conn.Close 
set conn=nothing 
%> 
</table> 
