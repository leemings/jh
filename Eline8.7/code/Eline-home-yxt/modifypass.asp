<%@ LANGUAGE=VBScript codepage ="936" %>
<html>
<head>
<title>��������</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="READONLY/Style.css">
</head>
<body bgcolor="#FFFFFF" background="readonly/Gauze.jpg">
<% 
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if Session("jiutian_name")="" or Session("jiutian_id")="" then response.end
nickname=Session("jiutian_name")
grade=Session("jiutian_grade")
myid=session("jiutian_id")	
If Session("ADMIN") <> True Then Response.Redirect "../error.asp?id=439"
if grade<9 then Response.Redirect "../error.asp?id=439"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
connstr=Application("lbjhhg_connstr")
conn.open connstr
user=Request.Form("user")
pass=Request.Form("pass")
if user<>"" and  pass<>"" then 
user=Request.Form("user")
if InStr(user,"=")<>0 or InStr(user,"`")<>0 or InStr(user,"'")<>0 or InStr(user," ")<>0 or InStr(user,"��")<>0 or InStr(user,"'")<>0 or InStr(user,chr(34))<>0 or InStr(user,"\")<>0 or InStr(user,",")<>0 or InStr(user,"<")<>0 or InStr(user,">")<>0 then Response.Redirect "error.asp?id=470"
pass=Request.Form("pass")
if InStr(pass,"=")<>0 or InStr(pass,"`")<>0 or InStr(pass,"'")<>0 or InStr(pass," ")<>0 or InStr(pass,"��")<>0 or InStr(pass,"'")<>0 or InStr(pass,chr(34))<>0 or InStr(pass,"\")<>0 or InStr(pass,",")<>0 or InStr(pass,"<")<>0 or InStr(pass,">")<>0 then Response.Redirect "error.asp?id=470"
pass=CStr(Replace(pass,chr(13)&chr(10),""))
sql="select * from �û� where ����='"&user&"'"
set rs=conn.Execute(sql)
if rs.EOF or rs.BOF then
Response.Write "<p align=center>�Ҳ������û�</p>"
Response.Write "<p align=center>[ <a href=modifypass.asp>����</a> ]</p>"
Response.End 
else
temppass=StrReverse(left(pass&"xzcvbmn,./",10))
templen=len(pass)
mmpassword=""
for j=1 to 10
 mmpassword=mmpassword+chr(asc(mid(temppass,j,1))-templen+int(j*1.1))
next
psw=replace(mmpassword,"'","B")


sql="update �û� set ����='"&psw&"' where ����='"&user&"'"
set rs=conn.Execute(sql)
conn.Close
set conn=nothing

ip=Request.ServerVariables("REMOTE_ADDR")
act="["&username&"]������["&user&"]������!"

Response.Write "<p align=center>"&user&"�����뼺�ɹ��޸�!</p>"
Response.Write "<p align=center>[ <a href=modifypass.asp>����</a> ]</p>"
Response.End 
end if
else
if user="" then
%>
<table width="400" border="1" cellspacing="0" cellpadding="10" align="center" bordercolorlight="#000000" bordercolordark="#FFFFFF">
  <tr align="center"> 
    <td height="47"><b><font color="#FFFFFF">��������ʾ��</font></b><font color="#FFFFFF">�˹����������ܸ�����������û���������</font><br>
    </td>
  </tr>
  <tr align="center" bordercolor="#FFFFFF"> 
    <td height="34"> 
      <form method="post" action="modifypass.asp">
        <font color="#CCCCCC">�û�����</font> 
        <input type="text" name="user" size="10" maxlength="10">
        <input type="submit" name="Submit" value="����">
      </form>
      
    </td>
  </tr>
  <tr align="center">
    <td height="18">&nbsp;</td>
  </tr>
</table>
<p align="center"><font color="#FFFFFF">[ <a href=modifypass.asp>����</a> ]</font></p>
<%
else

if InStr(user,"=")<>0 or InStr(user,"`")<>0 or InStr(user,"'")<>0 or InStr(user," ")<>0 or InStr(user,"��")<>0 or InStr(user,"'")<>0 or InStr(user,chr(34))<>0 or InStr(user,"\")<>0 or InStr(user,",")<>0 or InStr(user,"<")<>0 or InStr(user,">")<>0 then Response.Redirect "error.asp?id=470"
sql="select * from �û� where ����='"&user&"'"
set rs=conn.Execute(sql)
if rs.EOF or rs.BOF then
Response.Write "<p align=center>�Ҳ������û�</p>"
Response.Write "<p align=center>[ <a href=modifypass.asp>����</a> ]</p>"
Response.End 
end if
%>
<table width="400" border="1" cellspacing="0" cellpadding="10" align="center" bordercolorlight="#000000" bordercolordark="#FFFFFF">
  <tr align="center"> 
    <td height="18"><b><font color="#FFFFFF">��������ʾ��</font></b><font color="#FFFFFF">�˹����������ܸ�����������û���������</font></td>
  </tr>
  <tr align="center" bordercolor="#FFFFFF"> 
    <td height="34"> 
      <form method="post" action="modifypass.asp" id=form1 name=form1>
        �û����� 
        <input type="text" name="user" size="10" maxlength="10" style="text-align:center;" value='<%=user%>' readonly>
        �������Ϊ�� 
        <input name="pass" size="10" maxlength="10" type="password">
        <input type="submit" name="Submit" value="����">
      </form>
      
    </td>
  </tr>
  <tr align="center">
    <td height="18">&nbsp;</td>
  </tr>
</table>
<p align="center"><font color="#FFFFFF">[ <a href=modifypass.asp>����</a> ]</font></p>
<%
end if
end if
%>
</body>
</html>
