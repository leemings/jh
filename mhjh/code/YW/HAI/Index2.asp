<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
timetmp=now()
username=session("yx8_mhjh_username")
if username="" then Response.Redirect "../error.asp?id=016"
set conn=server.createobject("adodb.connection") 
conn.open Application("yx8_mhjh_connstr")
set rs=server.CreateObject ("adodb.recordset")
name=Session("yx8_mhjh_userName")
sql="select ����ʱ��,����,���� from �û� where ����='" & name & "'"
set rs=conn.execute(sql)
if rs("����")<1 then
%>
<script language=vbscript>
MsgBox "�ҵ�������һ�����嶼û��ѽ,��취Ūһ�������ɣ�"
location.href = "javascript:self.close()"
</script>
<%
else
if rs("����")>0 then
sql="update �û� set ����=����-1,����ʱ��='" & timetmp & "',����='�µ�̽��' where ����='" & name & "'"
conn.execute sql
dim newtalkarr(600) 
   talkarr=Application("yx8_mhjh_talkarr") 
		talkpoint=Application("yx8_mhjh_talkpoint") 
		j=1 
		for i=11 to 600 step 10 
			newtalkarr(j)=talkarr(i) 
			newtalkarr(j+1)=talkarr(i+1) 
			newtalkarr(j+2)=talkarr(i+2) 
			newtalkarr(j+3)=talkarr(i+3) 
			newtalkarr(j+4)=talkarr(i+4) 
			newtalkarr(j+5)=talkarr(i+5) 
			newtalkarr(j+6)=talkarr(i+6) 
			newtalkarr(j+7)=talkarr(i+7) 
			newtalkarr(j+8)=talkarr(i+8) 
			newtalkarr(j+9)=talkarr(i+9) 
			j=j+10 
		next 
		newtalkarr(591)=talkpoint+1 
		newtalkarr(592)=1 
		newtalkarr(593)=0 
		newtalkarr(594)=username 
		newtalkarr(595)="���" 
		newtalkarr(596)="" 
		newtalkarr(597)="000000" 
		newtalkarr(598)="000000" 
		newtalkarr(599)="<font color=red>������ͨ����</font><font color=blue>"&name&"ʹ����һ������,�ɹ����뺣��µ�,�ж���ʼʱ����" & timetmp & "������ף�����˰�!</font>" 
newtalkarr(600)=Session("yx8_mhjh_userchatroomsn")
Application.Lock
Application("yx8_mhjh_talkpoint")=talkpoint+1
Application("yx8_mhjh_talkarr")=newtalkarr
Application.UnLock
erase newtalkarr
erase talkarr
end if
end if
%>
<html>
<head>
<title>ħ�ñ��������</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="../../style.css">
</head>

<body  background='../../chatroom/bg1.gif' oncontextmenu=self.event.returnValue=false leftmargin="5" marginwidth="5">
<div align="center"><font color="#FFFFFF"><br>
</font>
<font size="2"> [��һ������������ħ�ú�������������������и��ֽ�����ʿ�Ӳ�ͬ�ĵط���̽���Ĳر���Ϣ]<br>
</font><font size="2"><br>
</font> </div>
<div align="center">
<table width="600" border="1" cellspacing="0" cellpadding="3" align="center" bordercolor="#000000" height="26" bordercolordark="#FFFFFF">
<tr bgcolor="#0066CC">
<td height="10" colspan="4" bgcolor="#B88230">
<div align="center"><span style="letter-spacing: 1"><font color="#FFFFFF">���±�����Ϣ</font></span></div>
</td>
</tr>
<tr>
<td width="165" height="9">
<div align="center"><span style="letter-spacing: 1"><font color="red">������</font></span>
</div>
</td>
<td colspan="3" height="9" width="429">
<div align="center"><font color="red">���Ӳ���</font></div>
</td>
</tr>
<!--#include file="data.asp"-->
<%
sql="SELECT * FROM ���� order by id desc"
Set Rs=connt.Execute(sql)
do while not rs.bof and not rs.eof
%>
<tr>
<td width="165" height="19">
<div align="center"><font color="#FF0000"><%=rs("������")%></font></div>
</td>
<td colspan="3" height="19" width="429"><span style="letter-spacing: 1">
</span>
<p align="center"><span style="letter-spacing: 1"><font color="red">������+<%=rs("������")%>
������+<%=rs("������")%>�۸�+<%=rs("�۸�")%></font></span></p>   
</td>   
</tr>   
<%   
rs.movenext   
loop   
%>   
</table>   
<table width="600" border="1" cellspacing="0" cellpadding="3" align="center" height="26" bordercolor="#000000" bordercolordark="#FFFFFF">   
<tr>   
<td height="10" bgcolor="#B88230" colspan="4" valign="middle">   
<div align="center"><font size="2" color="#FFFFFF">���±��﷢����</font></div>   
</td>   
</tr>   
  
<tr>   
<td width="126" height="9">   
<div align="center"><span style="letter-spacing: 1"><font color="#FF0033">������</font></span>   
</div>   
</td>   
<td width="218" height="9">   
<div align="center"><font color="#FF0033">���Ӳ���</font></div>   
</td>   
<td width="115" height="9">   
<div align="center"><span style="letter-spacing: 1"><font color="red">������</span></font></div>   
</td>   
<td width="131" height="9">   
<div align="center"><span style="letter-spacing: 1"><font color="red">����ʱ��</span></font></div>   
</td>   
</tr>   
<%   
sql="SELECT * FROM ���� where ����='1' order by id desc"   
Set Rs=connt.Execute(sql)   
do while not rs.bof and not rs.eof   
%>  
<tr>   
<td width="126" height="19">   
<div align="center"><font color="#FF0033"><%=rs("������")%></font></div>  
</td>  
<td width="218" height="19">  
<div align="center"><span style="letter-spacing: 1"><font color="red">������+<%=rs("������")%>  
������+<%=rs("������")%> </span></div>   
</font>   
</td>   
<td width="115" height="19">   
<div align="center"><span style="letter-spacing: 1"><%=rs("����")%></span></div>   
</td>   
<td width="131" height="19">   
<p align="center"><span style="letter-spacing: 1"><%=rs("ʱ��")%></span></p>   
</td>   
</tr>   
<%   
rs.movenext   
loop   
conn.close 
set conn=nothing 
connt.close 
set connt=nothing 
%>   
</table>   
<br>   
<br>   
[ <b><a href="index.asp">�� ½ �� ������ ʼ Ѱ ��</a></b> ]<br>   
<table width="551" border="1" cellspacing="1" cellpadding="0" bordercolor="#000000">   
<tr>   
<td height="94">   
<p style="margin-left: 5; margin-right: 5"><br>   
<font color="#000000">  
&lt;&lt;</font><font color="#0066CC">Ѱ���ض�</font><font color="#000000">&gt;&gt;</font>  
</p>  
<p style="margin-left: 5; margin-right: 5; margin-bottom: 8">1����˵���ﶼ����������ŷ�ɥ����<font color="#FF0000">�����µ�</font>������Ѱ�ҵ��Ļ������˵��΢����΢��������ʱ�п��ܱ���û�õ��־��������ڹµ��ˡ�<br>  
<br>  
2�����ﶼ�ǽ�����ϡ���䱦�������ڰ���������õ��ֺǣ����������˭�õ�����Ļ���Ҫ���ܺúǣ���Ϊ���ܿ���Ϊ�ڰ��<font color="#FF0000">��������</font>�ǡ�  
</p>  
</td>  
</tr>  
</table>  
  
<br>  
</div>  
</body>  
</html>  
