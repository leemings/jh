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
response.write "<script language=JavaScript>alert('�޸Ĺ���Ա���ϳɹ���');top.location.href='main.asp';</script>"
end if

if Request("action")="edit" then
	IF Request.form("del")<>""  Then
		Sql = "Delete From music Where id="&Request.form("id")
		Conn.Execute(Sql)
		response.write "<script language=JavaScript>alert('ɾ��������Ŀ�ɹ���');top.location.href='main.asp';</script>"
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
response.write "<script language=JavaScript>alert('�޸ĸ���������ɣ�');top.location.href='main.asp';</script>"
end if
end if

if Request("action")="add" then
set rs=server.CreateObject("adodb.recordset")
sql="select * from music"
rs.open sql,conn,3,2
rs.addnew
if trim(request.form("musicname"))="" then
  response.write "<script language=javascript>"	
		response.write "alert('����д����');"	
		response.write "</script>"
		response.write "<script language=javascript>location='javascript:history.back(1)'</script>"
   Response.End
   end if
   if trim(request.form("file"))="" then
  response.write "<script language=javascript>"	
		response.write "alert('����д��ַ');"	
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
<title>��ҳ���ֲ����� - ����</title>
<style>
body         { color: #12463b; FONT-FAMILY ����; font-size: 9pt }
td           { color: #12463b; FONT-FAMILY ����; font-size: 9pt }
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
<font color="#EE0000">��ӭ����Ա<%=Request.cookies("users")("username")%>��½ <a href="?action=exit"><font color="#FFFF33">[�˳�]</font></a></font><br/>
<br/>
<table width="750"  border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#000000">
  <tr bgcolor=#001100> 
    <td> <font color="#EEEEEE"><strong>��������</strong></font></td>
    <td> <font color="#EEEEEE">��������</font></td>
    <td> <font color="#EEEEEE">������ַ</font></td>
    <td> <font color="#EEEEEE">����</font></td>
    <td> <font color="#EEEEEE">˳��</font></td>
    <td colspan="2"> 
      <div align="center"><font color="#EEEEEE">����</font></div></td>
  </tr>
  <%
sql="select * from music order by id desc" 
Set rs=Server.CreateObject("ADODB.RecordSet") 
rs.Open sql,conn,1,3
if rs.eof and rs.bof then
	response.write "<div align=center><font size=2>��û��</font></div>" 
else
	const maxperpage = 10 'ÿҳ��ʾ��¼��
	strFileName="?"
	page=Trim(request("Page"))  'pageֵΪ����ֵ
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
          <input name="Submit2" type="submit" class="button1" value="�޸�">
        </div></td>
      <td nowrap bgcolor="#FFFFFF"> <input name="del" type="submit" class="button1" value="ɾ��"> 
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
		call showpage(strFileName,totalPut,MaxPerPage,Page,Pages,"������","Page")
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
      <td bgcolor="#000000" align="right"><font color=#ffffff>������Ϣ��</font></td>
      <td bgcolor="#FFFFFF">
	  ����:
	  <input name="singer" type="text" class="input01"  size="20"��title="singer">
      ��������:
	  <input name="musicname" type="text" class="input01"  size="20" title="music">
     </td>
    </tr>
	<tr> 
      <td bgcolor="#000000" align="right"><font color=#ffffff>����ѡ�</font></td>
      <td bgcolor="#FFFFFF">
	  �Ƿ񲥷�:
	  <input name=checkup type=checkbox value=checkup checked>
      ����˳��:
	  <input name="playorder" type="text" id="playorder" value="1" size="2">
     </td>
    </tr>
    <tr> 
      <td bgcolor="#000000" align="right"><font color=#ffffff>������ַ��</font></td>
      <td bgcolor="#FFFFFF">
	  <input name="file" type="text" class="input01" id="file" value="http://" size="50"> 
      </td>
    </tr>
    <tr> 
      <td bgcolor="#000000"></td>
      <td bgcolor="#FFFFFF">
	  <input name="Submit" type="submit" class="button1" value=" �� �� "> 
        &nbsp;&nbsp;&nbsp; 
	  <input name="Submit1" type="reset" class="button1" value=" �� �� "></td>
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
      <td bgcolor="#000000" align="right"><font color=#ffffff>����Ա��Ϣ��</font></td>
      <td bgcolor="#FFFFFF">
	  ԭ����Ա����:
	  <input name="oldusername" type="text" class="input"  size="20" value="<%=rs("username")%>" readonly>
      ԭ����Ա����:
	  <input name="oldpassword" type="password" class="input"  size="20" value="<%=rs("password")%>" readonly>
     </td>
    </tr>
	<tr>
      <td bgcolor="#000000" align="right"><font color=#ffffff>����Ա�޸ģ�</font></td>
      <td bgcolor="#FFFFFF">
	  �¹���Ա����:
	  <input name="username" type="text" class="input"  size="20">
      �¹���Ա����:
	  <input name="password" type="password" class="input"  size="20">
	  <input name="id" type="hidden" id="id" value="<%=rs("id")%>">
     </td>
    </tr>
    <tr> 
      <td bgcolor="#000000"></td>
      <td bgcolor="#FFFFFF">
	  <input name="Submit" type="submit" class="button1" value=" �� �� "></td>
    </tr>
  </form>
</table>
<p><CENTER>
    <font color="#EEEEEE">������ַ�ǹ����½��ڣ������ö������Ա</font><font color="#000000"><br>
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
	strTemp=strTemp&"��"&strUnit&" <b>"&totalnumber&"</b>����"
	strTemp=strTemp&"ÿҳ��ʾ <b>"&maxperpage&"</b>������ҳ�� <b>"&totalpages&"</b></td>" 
	strTemp=strTemp&"<td align=center>��ҳ�� <a href="&sfilename&sPage&"=1><font face=webdings>9</font></a> "
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