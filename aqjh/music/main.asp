<!--#include file="conn.asp"-->
<%
if Session("open")<>True then
Response.Redirect "login.asp"
end if

if Request("action")="exit" then
	response.cookies("users")("username")=""
	response.cookies("users")("userpass")=""
	Session("open")=False
	Response.redirect"login.asp"
end if

if Request("action")="admin" then
Set rs=Server.CreateObject("ADODB.RecordSet") 
sql="select * from admin where id="&request("id")
rs.open sql,conn,1,3 
rs("username")=trim(request("username"))
rs("password")=trim(request("password"))
rs.Update
rs.Close
set rs=nothing
response.write "<script language=JavaScript>alert('修改管理员资料成功！');top.location.href='main.asp';</script>"
end if

if Request("action")="edit" then
	IF Request.form("del")<>""  Then
		Sql = "Delete From music Where id="&Request.form("id")
		Conn.Execute(Sql)
		response.write "<script language=JavaScript>alert('删除音乐条目成功！');top.location.href='main.asp';</script>"
		Response.end
	else
Set rs=Server.CreateObject("ADODB.RecordSet") 
sql="select * from music where id="&request("id")
rs.open sql,conn,1,3 
rs("musicname")=trim(request("musicname"))
rs("singer")=trim(request("singer"))
rs("file")=trim(request("file"))
if trim(request("checkup"))<>"" then
	rs("checkup")=true
else
	rs("checkup")=false
end if
playorder=cint(trim(request("playorder")))
if playorder="" then playorder=1
rs("playorder")=playorder
rs.Update
rs.Close
set rs=nothing
response.write "<script language=JavaScript>alert('修改歌曲资料完成！');top.location.href='main.asp';</script>"
end if
end if

if Request("action")="add" then
set rs=server.CreateObject("adodb.recordset")
sql="select * from music"
rs.open sql,conn,3,2
rs.addnew
if trim(request.form("musicname"))="" then
  response.write "<script language=javascript>"	
		response.write "alert('请填写名称');"	
		response.write "</script>"
		response.write "<script language=javascript>location='javascript:history.back(1)'</script>"
   Response.End
   end if
   if trim(request.form("file"))="" then
  response.write "<script language=javascript>"	
		response.write "alert('请填写地址');"	
		response.write "</script>"
		response.write "<script language=javascript>location='javascript:history.back(1)'</script>"
   Response.End
   end if
rs("musicname")=trim(request("musicname"))
rs("singer")=trim(request("singer"))
rs("file")=trim(request("file"))
if trim(request("checkup"))<>"" then
	rs("checkup")=true
else
	rs("checkup")=false
end if
playorder=trim(request("playorder"))
if playorder="" then playorder=1
rs("playorder")=playorder
rs.update
rs.close
response.redirect "main.asp"
end if
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>首页音乐播放器 - 管理</title>
<style>
body         { color: #12463b; FONT-FAMILY 宋体; font-size: 9pt }
td           { color: #12463b; FONT-FAMILY 宋体; font-size: 9pt }
a            { color: #12463b; font-size: 9pt; text-decoration: none }
a:hover      { color: red; font-size: 9pt; text-decoration: none }
a.linkblue   { color: #2222cc; font-size: 9pt; text-decoration: none }
a.linkgr     { color: #666622; font-size: 9pt; text-decoration: none }
a.linkblue:hover { background-color: #000066; color: white; font-size: 9pt; text-decoration: none }

table {
	font-size: 9pt;
}
input {
	background-color: #FFFFFF;
	border: 1px solid #CCCCCC;
}
</style>
</head>

<body bgcolor="#32404B">
<font color="#EE0000">欢迎管理员<%=Request.cookies("users")("username")%>登陆 <a href="?action=exit"><font color="#FFFF33">[退出]</font></a></font><br/>
<br/>
<table width="750"  border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#000000">
  <tr bgcolor=#001100> 
    <td> <font color="#EEEEEE"><strong>歌手姓名</strong></font></td>
    <td> <font color="#EEEEEE">歌曲名字</font></td>
    <td> <font color="#EEEEEE">歌曲地址</font></td>
    <td> <font color="#EEEEEE">播放</font></td>
    <td> <font color="#EEEEEE">顺序</font></td>
    <td colspan="2"> 
      <div align="center"><font color="#EEEEEE">操作</font></div></td>
  </tr>
  <%
sql="select * from music order by id desc" 
Set rs=Server.CreateObject("ADODB.RecordSet") 
rs.Open sql,conn,1,3
if rs.eof and rs.bof then
	response.write "<div align=center><font size=2>还没有</font></div>" 
else
	const maxperpage = 10 '每页显示记录数
	strFileName="?"
	page=Trim(request("Page"))  'page值为接受值
	if not isnumeric(page) then page = 1 else page = cint(page)
	totalPut=rs.recordcount
	rs.pagesize = maxperpage
	pages = rs.pagecount
	if page < 1 then page = 1
	if page > pages then page = pages
	rs.absolutepage = page
	for i = 1 to maxperpage
		if rs.eof then exit for
		if rs("checkup")=True then
			checktoplay="<input name=checkup type=checkbox value=checkup checked>"
		else
			checktoplay="<input name=checkup type=checkbox value=checkup>"
		end if
%>
  <tr> 
    <form action="main.asp?action=edit" method="post">
      <td bgcolor="#FFFFFF" > <input name="singer" type="text" id="singer" value="<%=rs("singer")%>" size="10"> 
      </td>
      <td bgcolor="#FFFFFF" > <div align="center"> 
          <input name="musicname" type="text" id="musicname" value="<%=rs("musicname")%>" size="15">
          <input name="id" type="hidden" id="id" value="<%=rs("id")%>">
        </div></td>
      <td bgcolor="#FFFFFF"> <div align="center"> 
          <input name="file" type="text" id="file" value="<%=rs("file")%>" size="50">
        </div></td>
      <td bgcolor="#FFFFFF"> <div align="center"> <%=checktoplay%> </div></td>
      <td bgcolor="#FFFFFF"> <div align="center"> 
          <input name="playorder" type="text" id="playorder" value="<%=rs("playorder")%>" size="2">
        </div></td>
      <td nowrap bgcolor="#FFFFFF"> <div align="center"> 
          <input name="Submit2" type="submit" class="button1" value="修改">
        </div></td>
      <td nowrap bgcolor="#FFFFFF"> <input name="del" type="submit" class="button1" value="删除"> 
      </td>
    </form>
  </tr>
  <%
	      rs.movenext
	next
end if
  %>
  <tr bgcolor="#FFFFFf"> 
    <td colspan="7"> 
      <%
	if TotalPut>0 then
		call showpage(strFileName,totalPut,MaxPerPage,Page,Pages,"歌曲数","Page")
	end if
rs.close   
set rs=nothing 
%>
    </td>
  </tr>
</table>
<br>
<table width="750"  border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#000000">
  <form action="main.asp?action=add" method="post" name="form" >
    <tr> 
      <td bgcolor="#000000" align="right"><font color=#ffffff>歌曲信息：</font></td>
      <td bgcolor="#FFFFFF">
	  歌手:
	  <input name="singer" type="text" class="input01"  size="20"　title="singer">
      歌曲名字:
	  <input name="musicname" type="text" class="input01"  size="20" title="music">
     </td>
    </tr>
	<tr> 
      <td bgcolor="#000000" align="right"><font color=#ffffff>播放选项：</font></td>
      <td bgcolor="#FFFFFF">
	  是否播放:
	  <input name=checkup type=checkbox value=checkup checked>
      播放顺序:
	  <input name="playorder" type="text" id="playorder" value="1" size="2">
     </td>
    </tr>
    <tr> 
      <td bgcolor="#000000" align="right"><font color=#ffffff>歌曲地址：</font></td>
      <td bgcolor="#FFFFFF">
	  <input name="file" type="text" class="input01" id="file" value="http://" size="50"> 
      </td>
    </tr>
    <tr> 
      <td bgcolor="#000000"></td>
      <td bgcolor="#FFFFFF">
	  <input name="Submit" type="submit" class="button1" value=" 增 加 "> 
        &nbsp;&nbsp;&nbsp; 
	  <input name="Submit1" type="reset" class="button1" value=" 重 设 "></td>
    </tr>
  </form>
</table>
<br>
<%
Set rs=Server.CreateObject("ADODB.RecordSet") 
sql="select id,username,password from admin"
rs.open sql,conn,1,3
%>
<table width="750"  border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#000000">
  <form action="main.asp?action=admin" method="post" name="form" >
    <tr>
      <td bgcolor="#000000" align="right"><font color=#ffffff>管理员信息：</font></td>
      <td bgcolor="#FFFFFF">
	  原管理员名称:
	  <input name="oldusername" type="text" class="input"  size="20" value="<%=rs("username")%>" readonly>
      原管理员密码:
	  <input name="oldpassword" type="password" class="input"  size="20" value="<%=rs("password")%>" readonly>
     </td>
    </tr>
	<tr>
      <td bgcolor="#000000" align="right"><font color=#ffffff>管理员修改：</font></td>
      <td bgcolor="#FFFFFF">
	  新管理员名称:
	  <input name="username" type="text" class="input"  size="20">
      新管理员密码:
	  <input name="password" type="password" class="input"  size="20">
	  <input name="id" type="hidden" id="id" value="<%=rs("id")%>">
     </td>
    </tr>
    <tr> 
      <td bgcolor="#000000"></td>
      <td bgcolor="#FFFFFF">
	  <input name="Submit" type="submit" class="button1" value=" 修 改 "></td>
    </tr>
  </form>
</table>
<p><CENTER>
    <font color="#EEEEEE">以上网址是管理登陆入口，可设置多个管理员</font><font color="#000000"><br>
    </font></CENTER></p>
</body>
</html>
<%
rs.close   
set rs=nothing   
conn.close   
set conn=nothing

sub showpage(sfilename,totalnumber,maxperpage,nPage,totalpages,strUnit,sPage)
	strTemp="<table width=100% border=0 cellspacing=0 cellpadding=0><tr><td>"
	strTemp=strTemp&"总"&strUnit&" <b>"&totalnumber&"</b>　　"
	strTemp=strTemp&"每页显示 <b>"&maxperpage&"</b>　　总页数 <b>"&totalpages&"</b></td>" 
	strTemp=strTemp&"<td align=center>分页： <a href="&sfilename&sPage&"=1><font face=webdings>9</font></a> "
		For i=1 to totalpages
		    if i=npage then
			strTemp=strTemp & "<a href="&sfilename&sPage&"="&i&"><font color=#ff6600><b>"&i&"</b></font></a> "
			else
			strTemp=strTemp & "<a href="&sfilename&sPage&"="&i&"><b>"&i&"</b></a> "
			end if
		Next
	strTemp=strTemp & " <a href="&sfilename&sPage&"="&totalpages&"><font face=webdings>:</font></a></td></tr></table>"
	response.write strTemp
end sub
%>