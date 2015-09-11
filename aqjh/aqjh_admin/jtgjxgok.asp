<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.Buffer=true
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
if session("aqjh_adminok")<>true then Response.Redirect "../chat/readonly/bomb.htm"
if aqjh_grade<>10 or instr(Application("aqjh_admin"),aqjh_name)=0  then Response.Redirect "../error.asp?id=439"
id=request("id")
lx=request.form("lx")
tp=request.form("tp")
jg=request.form("jg")
zddj=request.form("zddj")
gldj=request.form("gldj")
hyzy=request.form("hyzy")
sm=request.form("sm")
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select * from 车 where id="& id,conn
oldtp=rs("图片")
rs.close
conn.execute "update 车 set 类型='" & lx & "',图片='" & tp & "',战斗等级=" & zddj & ",管理等级=" & gldj & ",价格=" & jg & ",会员专用=" & hyzy & ",说明='" & sm & "' where id="&id
set rs=nothing
conn.close
set conn=nothing
%>
<html>
<head>
<title>交通工具修改成功</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>

<body bgcolor="#FFFFFF" text="#000000" background="../JHIMG/bk_Hc3w.gif">
<p align="center">交通工具修改成功</p>
<p align="center">类型：<%=lx%><br>
  图片：<%=tp%> <br>
  购买条件：<%=gmtj%> <br>
  价格： <%=jg%><br>
  会员专用：<%=hyzy%> <br>
  进入说明： <%=sm%></p>
<p align="center"><a href="manjtgj.asp">返回</a> </p>
</body>
</html>
