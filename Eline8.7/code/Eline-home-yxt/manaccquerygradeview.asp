<%@ LANGUAGE=VBScript codepage ="936" %>
<%if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
%>
<%Response.Expires=0
if session("sjjh_adminok")<>true then Response.Redirect "../chat/readonly/bomb.htm"
nickname=info(0)
grade=Int(info(2))
if info(5)<>"官府" and info(2)<10 or info(0)<>Application("sjjh_admin") then Response.Redirect "manerr.asp?id=242"
if nickname="" then Response.Redirect "manerr.asp?id=100"'
'if nickname="" then Response.Redirect "manerr.asp?id=100"
'if grade<>10 then Response.Redirect "manerr.asp?id=130"
n=Year(date())
y=Month(date())
r=Day(date())
s=Hour(time())
f=Minute(time())
m=Second(time())
if len(y)=1 then y="0" & y
if len(r)=1 then r="0" & r
if len(s)=1 then s="0" & s
if len(f)=1 then f="0" & f
if len(m)=1 then m="0" & m
sj=n & "-" & y & "-" & r & " " & s & ":" & f & ":" & m
querygrade=Request.QueryString("grade")
if querygrade="" then querygrade=1
if Not(IsNumeric(querygrade)) then querygrade=1
if querygrade<0 then page=0
if querygrade>10 then querygrade=10
querygrade=int(querygrade)
page=Request.QueryString("page")
if page="" then page=1
if Not(IsNumeric(page)) then page=1
if page<1 then page=1
page=int(page)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")

conn.open Application("sjjh_usermdb")
sql="SELECT 姓名,grade,times,allvalue,lasttime,lastip FROM 用户 WHERE kill='0' AND grade=" & querygrade
rs.open sql,conn,1,1
rs.PageSize=30
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
Redim Preserve show(j),show(j+1),show(j+2),show(j+3),show(j+4),show(j+5)
show(j)=rs("姓名")
show(j+1)=rs("grade")
show(j+2)=rs("times")
show(j+3)=rs("allvalue")
show(j+4)=rs("lasttime")
show(j+5)=rs("lastip")
i=i+1
j=j+6
rs.MoveNext
Loop
end if
rs.close
conn.close
set rs=nothing
set conn=nothing%><html>
<head>
<title>帐号管理♀wWw.happyjh.com♀</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="../chat/readonly/style.css">
</head>
<body bgcolor="#FFFFFF" class=p150 background="../jhimg/bk_hc3w.gif">
<div align="center">
<h1><font color="0099FF">【帐号管理】</font></h1>
<font color="#FF0000">【按级别列出所有帐号】</font> <br>
<a href="manaccquerygrade.asp">返回</a></div>
<hr noshade size="1" color=009900>
<b>［注意事项］</b><br>
共有 <font color="#FF0000"><%=querygrade%></font> 级的帐号 <font color=red><%=totalrec%></font>
个。<%if totalrec>0 then%>
<hr noshade size="1" color=009900>
<div align=center>
<%for i=1 to totalpage
if page=i then
Response.Write " [" & i & "]"
else
Response.Write " <a href=manaccquerygradeview.asp?grade=" & querygrade & "&page=" & i & ">[" & i & "]</a>"
end if
next%></div>
<hr noshade size="1" color=009900>
<table border="1" cellspacing="0" cellpadding="6" bordercolorlight="#999999" bordercolordark="#FFFFFF" bgcolor="E0E0E0" align="center" width="460">
<tr bgcolor="#3399FF">
<td><font color="#FFFFFF">序</font></td>
<td><font color="#FFFFFF">用户名</font></td>
<td><font color="#FFFFFF">等级</font></td>
<td><font color="#FFFFFF">次数</font></td>
<td><font color="#FFFFFF">经验值</font></td>
<td><font color="#FFFFFF">最后时间</font></td>
<td><font color="#FFFFFF">最后ＩＰ</font></td>
</tr>
<%for i=1 to UBound(show) step 6%>
<tr>
<td><%=(page-1)*30+(i+5)/6%></td>
<td><%=show(i)%></td>
<td><%=show(i+1)%></td>
<td><%=show(i+2)%></td>
<td><%=show(i+3)%></td>
<td><%=show(i+4)%></td>
<td><%=show(i+5)%></td>
</tr>
<%next%>
</table><%end if%>
<hr noshade size="1" color=009900>
<div align=center class=cp><%Response.Write "序列号：<font color=blue>" & Application("sjjh_sn") & "</font>，授权给：<font color=blue>" & Application("sjjh_user") & "</font><br><font color=999999>" & Application("sjjh_copyright") & "</font>"%></div>
</body>
</html>