<%@ LANGUAGE=VBScript codepage ="936" %><%Response.Expires=0
Response.Buffer=true

sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
if sjjh_name="" then Response.Redirect "../error.asp?id=440"
if session("sjjh_adminok")<>true then Response.Redirect "../chat/readonly/bomb.htm"
if sjjh_grade<>10 or instr(Application("sjjh_admin"),sjjh_name)=0  then Response.Redirect "../error.asp?id=439"
ipdizi=request.form("ipdizi")
if ipdizi="" then%>
<html>
<head>
<title><%=Application("sjjh_chatroomname")%></title>
<style></style>
<link rel="stylesheet" href="../chat/READONLY/Style.css">
</head>
<body bgcolor=#FFFFFF background="../JHIMG/Bk_hc3w.gif">
<center>
<font color="#000000"><b><font size="+2">
<%ip=Request.ServerVariables("REMOTE_ADDR")
sip=split(ip,".")
num=cint(sip(0))*256*256*256+cint(sip(1))*256*256+cint(sip(2))*256+cint(sip(3))-1
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
'查最后ip的地区
sql="select * FROM z WHERE a<="& num &" and b>="&num
Set Rs=conn.Execute(sql)
if rs.eof or rs.bof then
country="亚洲"
city="未知"
else
country=rs("c")
city=rs("d")
end if
if country="" then country="中国"
if city="" then city="未知"
last="地区:"& country &" 城市:"& city
set rs=nothing	
set conn=nothing
%>
<%=Application("sjjh_chatroomname")%>ip查找程序</font></b></font><br>
<br>
感谢追捕作者提供的ip数据库程序(11.1日前)<br>
注：该程序仅收入国内ip地址！<br>
您的ip地址为：<%=ip%> <%=last%>
<form action=ipseek.asp method=post>
请输入ip地址并回车:
<input  size=20 name=ipdizi><input type=submit value='确认'>
</form>
</center>
</body>
</html>
<%else
ip=ipdizi
sip=split(ip,".")
num=cint(sip(0))*256*256*256+cint(sip(1))*256*256+cint(sip(2))*256+cint(sip(3))-1
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
'查最后ip的地区
sql="select top 1 c,d FROM z WHERE a<="& num &" and b>="&num
Set Rs=conn.Execute(sql)
if rs.eof or rs.bof then
country="亚洲"
city="未知"
else
country=rs("c")
city=rs("d")
end if
if country="" then country="中国"
if city="" then city="未知"
last="地区:"& country &" 城市:"& city
set rs=nothing	
set conn=nothing%>
<html>
<head>
<title><%=Application("sjjh_chatroomname")%></title>
<style></style>
<link rel="stylesheet" href="../chat/READONLY/Style.css">
</head>
<body bgcolor=#FFFFFF background="../JHIMG/Bk_hc3w.gif">
您所查找的ip地址为：<%=ip%><br>鉴定为：<%=last%>
<br>
<br>
<a href="ipseek.asp">返回</a>
</html>
<%end if%>