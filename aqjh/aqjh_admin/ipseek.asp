<%@ LANGUAGE=VBScript codepage ="936" %><%Response.Expires=0
Response.Buffer=true
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
if session("aqjh_adminok")<>true then Response.Redirect "../chat/readonly/bomb.htm"
if aqjh_grade<>10 or instr(Application("aqjh_admin"),aqjh_name)=0  then Response.Redirect "../error.asp?id=439"
ipdizi=request.form("ipdizi")
if ipdizi="" then%>
<html><head>
<LINK href=css/css.css type=text/css rel=stylesheet>
</head>
<body bgColor=#dcdec6 text=#000000 link=#003366 vLink=#006666>
<center>
<font color="#000000"><b>
<%ip=Request.ServerVariables("REMOTE_ADDR")
sip=split(ip,".")
num=cint(sip(0))*256*256*256+cint(sip(1))*256*256+cint(sip(2))*256+cint(sip(3))-1
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
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
<%=Application("aqjh_chatroomname")%>ip查找程序</b><br>
<br><br>
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
conn.open Application("aqjh_usermdb")
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
<head><LINK href=css/css.css type=text/css rel=stylesheet>
<title><%=Application("aqjh_chatroomname")%></title>
</head>
<body bgColor=#dcdec6 text=#000000 link=#003366 vLink=#006666>
您所查找的ip地址为：<%=ip%><br>鉴定为：<%=last%>
<br>
<br>
<a href="ipseek.asp">返回</a>
</html>
<%end if%>