<!--#include file="../../config.asp"-->
<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
timetmp=now()
my=session("yx8_mhjh_username")
if my="" then Response.Redirect "../error.asp?id=016"
if session("yx8_mhjh_usercorp")="ʮ�˵���" then%>
<script language=vbscript>
MsgBox "���󣡹���ֹ���ڣ�"
location.href = "javascript:history.back()"
</script>
<%response.end
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
connstr=Application("yx8_mhjh_connstr")
conn.open connstr
sql="select ����ʱ��,����,���� from �û� where ����='" & my & "'"
set rs=conn.execute(sql)
if rs("����")<1 then%>
<script language="vbscript">
  alert("����һ�����嶼û��,��ô��ִ���������")
location.href = "javascript:self.close()"
</script>
<%
else
if rs("����")>0 then
sql="update �û� set ����=����-1,����ʱ��='" & timetmp & "',����='ɽ��Ѱ��' where ����='" & my & "'"
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
		newtalkarr(599)="<font color=red>������ͨ����</font><font color=blue>"&my&"ʹ����һ������,�ɹ�����ħ��ɽ��,�ж���ʼʱ����" & timetmp & "������ף�����˰�!</font>" 
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
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>NPC</title>
<link rel="stylesheet" href="../../Style2.css">
</head>
<table border="0" width="100%" cellspacing="0" cellpadding="0">
  <tr>
    <td width="20%"></td>
    <td width="46%"><img border="0" src="../../pic/OLDMAN.jpg"></td>
    <td width="33%"><font color="#800000">�������ֵ�</font><font color="#FF0000">��ħ����ʦ��</font><font color="#800000"><br>
      <br>
      <br>
      &nbsp;&nbsp;&nbsp; </font><a href="lian/1.asp"><font color="#800000">���������������ϰ��ľ׮��</font></a><font color="#800000"><br> 
      <br>
      <br>
      &nbsp;&nbsp;&nbsp; </font><a href="lian/d1.asp" target="_blank"><font color="#800000">����ܸ��ҵ�þƣ�������һ������</font></a><font color="#800000"><br> 
      <br>
      &nbsp;&nbsp;&nbsp; <br>
      &nbsp;&nbsp;&nbsp; </font><a href="lian/d2.asp"><font color="#800000">�������ܲ��ܴ��Ϻ��������ѵ㶫��<br>
      <br>
      <br>
      </font></a><font color="#800000"><a href="lian/d2.asp">&nbsp;&nbsp;&nbsp; </a></font><font color="#FF0000"><b>[ÿ����һ�ζ�Ҫ�ķ�һ�����]</b></font></td>
  </tr>
</table>
