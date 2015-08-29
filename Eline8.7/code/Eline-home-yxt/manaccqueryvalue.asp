<%@ LANGUAGE=VBScript codepage ="936" %>
<%if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
%>
<%Response.Expires=0
if session("sjjh_adminok")<>true then Response.Redirect "../chat/readonly/bomb.htm"
page=Request.QueryString("page")
if page="" then page=1
if Not(IsNumeric(page)) then page=1
if page<1 then page=1
page=int(page)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")

conn.open Application("sjjh_usermdb")
sql="SELECT 姓名,grade,times,allvalue,regip,lastip,lasttime FROM 用户 WHERE kill='0' ORDER BY allvalue DESC"
rs.open sql,conn,1,1
rs.PageSize=15
totalrec=rs.RecordCount
totalpage=rs.PageCount
if page>totalpage then page=totalpage
if totalrec>0 then
rs.AbsolutePage=page
p=1+(page-1)*rs.PageSize
Dim show()
i=0
j=1
Do while (Not rs.Eof) and (i<rs.PageSize)
Redim Preserve show(j),show(j+1),show(j+2),show(j+3),show(j+4),show(j+5),show(j+6)
show(j)=rs("姓名")
show(j+1)=rs("grade")
show(j+2)=rs("times")
show(j+3)=rs("allvalue")
show(j+4)=rs("regip")
show(j+5)=rs("lastip")
show(j+6)=rs("lasttime")
i=i+1
j=j+7
rs.MoveNext
Loop
end if
rs.close
conn.close
set rs=nothing
set conn=nothing%>
<html>
<head>
<title>帐号管理♀wWw.51eline.com♀</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>
<body bgcolor="#FFFFFF" class=p150 background="../jhimg/bk_hc3w.gif">
<div align="center">
<h1><font color="0099FF">【帐号管理】</font></h1>
<font color="#FF0000">【按累积分列出帐号】</font> <br>
<a href="manacc.asp">返回</a></div>
<hr noshade size="1" color=009900>
<b>［注意事项］</b><br>
共找到 <font color=red><%=totalrec%></font> 个帐号。<%if totalrec>0 then%>
<hr noshade size="1" color=009900>
<div align=center>
<%for i=1 to totalpage
if page=i then
Response.Write " [" & i & "]"
else
Response.Write " <a href=manaccqueryvalue.asp?page=" & i & ">[" & i & "]</a>"
end if
next%></div>
<hr noshade size="1" color=009900>
<table border="1" cellspacing="0" cellpadding="6" bordercolorlight="#999999" bordercolordark="#FFFFFF" bgcolor="E0E0E0" align="center" width="100%">
<tr bgcolor="#3399FF">
<td width="6%" height="23">
<div align="center"><font color="#FFFFFF" size="2">序</font></div>
</td>
<td width="13%" height="23">
<div align="center"><font color="#FFFFFF" size="2">用户名</font></div>
</td>
<td width="10%" height="23">
<div align="center"><font color="#FFFFFF" size="2">等级</font></div>
</td>
<td width="10%" height="23">
<div align="center"><font color="#FFFFFF" size="2">次数</font></div>
</td>
<td bgcolor="FF6666" width="13%" height="23">
<div align="center"><font color="#FFFFFF" size="2">累积分</font></div>
</td>
<td width="13%" height="23">
<div align="center"><font color="#FFFFFF" size="2">注册ＩＰ</font></div>
</td>
<td width="13%" height="23">
<div align="center"><font color="#FFFFFF" size="2">最后ＩＰ</font></div>
</td>
<td width="22%" height="23">
<div align="center"><font color="#FFFFFF" size="2">最后时间</font></div>
</td>
</tr>
<%for i=1 to UBound(show) step 7%>
<tr>
<td width="6%"><%=(page-1)*15+(i+6)/7%></td>
<td width="13%"><%=show(i)%></td>
<td width="10%"><%=show(i+1)%></td>
<td width="10%"><%=show(i+2)%></td>
<td width="13%"><%=show(i+3)%></td>
<td width="13%"><%=show(i+4)%></td>
<td width="13%"><%=show(i+5)%></td>
<td width="22%"><%=show(i+6)%></td>
</tr>
<%next%>
</table>
<%end if%>
<hr noshade size="1" color=009900>
<div align=center class=cp><%Response.Write "序列号：<font color=blue>" & Application("sjjh_sn") & "</font>，授权给：<font color=blue>" & Application("sjjh_user") & "</font><br><font color=999999>" & Application("sjjh_copyright") & "</font>"%></div>
</body>
</html>