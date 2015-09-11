<html>
<head>
<meta http-equiv="Content-Language" content="zh-cn">
<%@ LANGUAGE=VBScript codepage ="936" %>
<%
Response.Expires=0
Response.Buffer=true
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
if aqjh_grade<9 then
   response.write "<script>alert('提示:需要九级以上管理!');window.close();</script>"
   response.end
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
Set Rs=conn.Execute("select  * FROM song")
if request("cz")="del" then
conn.Execute("delete * from song where name='"&request("name")&"'")
rs.close
rs.open "select  * FROM song",conn
end if
if request("cz")="ud" then
conn.Execute("update song set url='"&request("name1")&"' where name='"&name&"'")
rs.close
rs.open "select  * FROM song",conn
end if
if request("cz")="inst" then
rs.close
Rs.OPEN "select * FROM song where name='"&request("name")&"'",CONN
if rs.bof or rs.eof then
conn.Execute("insert into song(name,url) values ('" &request("name")& "','" &request("name2")& "')")
else
conn.Execute("update song set url='"&request("name1")&"' where name='"&request("name")&"'")
end if
rs.close
rs.open "select  * FROM song",conn
end if

%>
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>歌曲管理</title>
</head>
<script Language="JavaScript">
    function addurl()
    {
	var temp;
	var name;
	var url;
	temp = '请输入新歌曲名称';
	name =prompt(temp, '');
	if (name == ''){
		return;
	}
	temp = '请输入新歌曲的URL地址';
	url =prompt(temp, '');
	if (url == '')
	{
		return;
	}
sendurl="songadd.asp?cz=inst&name="+name+" &name2="+url;window.location.href=sendurl;
    }
 function udurl(name)
    {
	var temp;
	var url;
	temp = '请输入歌曲的新URL地址';
	url =prompt(temp, '');
	if (url == null || url == '')
	{
		return;
	}
sendurl="songadd.asp?cz=ud&name="+name+"&name2="+url;window.location.href=sendurl;
    }
</script>
<body background=bg.gif>
<center>
<table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="49%" height="8">
  <tr>
    <td width="65%" align="center" height="8"><font size="2">名称</font></td>
    <td width="35%" align="center" height="8"><font size="2">
    <font color="#333300"><a href="javaScript:addurl();">手工添加</font></a>    <font color="#000066">
</font></font></td>
  </tr>
<% if rs.bof or rs.eof then %>
  <tr>
    <td width="100%" align="center" height="8" colspan="2">
    <font size="2" color="#FF3300">没有任何歌曲</font></td>
  </tr>
<%
else
do while not rs.bof and not rs.eof 
%>
  <tr>
    <td width="65%" align="center" height="8"><font size="2"><a href="javaScript:window.open('<%=rs("url")%>');"><%=rs("name")%></a></font></td>
    <td width="35%" align="center" height="8"><%if aqjh_grade=10 then%><font size="2"><a href="javaScript:udurl('<%=rs("name")%>');">修改</a>   <a href=songadd.asp?cz=del&name=<%=rs("name")%>>删除</a><%else%><font size="2">无权</font><%end if%></td>
  </tr>
<%
rs.movenext
loop
end if
rs.close
conn.close
set rs=nothing
set conn=nothing
%>
</table><br>
<font size="2">九级以上可以添加<br>十级以上可以修删</font>
<center>
<p></p>
</body>
</html>
