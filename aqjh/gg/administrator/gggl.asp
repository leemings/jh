<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.Buffer=true
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../../error.asp?id=440"
if session("aqjh_adminok")<>true then Response.Redirect "../../chat/readonly/bomb.htm"
if aqjh_grade<>10 or aqjh_name<>application("aqjh_user") then Response.Redirect "../../error.asp?id=439"
mdbsj="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("../gg.asa")
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open mdbsj
rs.open "select * from l",conn
%>
<html>
<head>
<title>公告类别管理</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>
<body bgcolor="#FFFFFF" text="#000000" background="../../JHIMG/bk_Hc3w.gif">
<p align="center">&nbsp;</p>
<p align="center"><b><font size="4">站长公告类别设定</font></b></p>
<p> 
  <%maxid=0
  do while not (rs.eof or rs.bof)%>
</p>
<form name="form1" method="post" action="gglbgl.asp">
  <div align="center"><font size="2">ID： 
    <input type="text" name="id" size="6" readonly value="<%=rs("id")%>">
    类别名称： 
    <input type="text" name="lbmc" value="<%=rs("lb")%>" size="10">
    <input type="submit" name="tj" value="修改">
    <input type="submit" name="tj" value="删除">
    </font></div>
</form>
<%if maxid<rs("id") then
		maxid=rs("id")
	end if
rs.movenext
loop
rs.close
set rs=nothing
conn.close
set conn=nothing
maxid=maxid+1
%>
<form name="form2" method="post" action="gglbgl.asp">
  <div align="center"><font size="2">ID： 
    <input type="text" name="id" size="6" value="<%=maxid%>">
    类别名称： 
    <input type="text" name="lbmc" size="10">
    <input type="submit" name="tj" value="新增">
    </font> </div>
</form>
<p align="center"><a href="ggadmin.asp">返回</a></p>
</body>
</html>
