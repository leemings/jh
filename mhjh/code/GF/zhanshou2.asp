<!--#include file="../config.asp"-->
<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
username=session("yx8_mhjh_username")
if username="" then Response.Redirect "../error.asp?id=016"
usercorp=session("yx8_mhjh_usercorp")
usergrade=session("yx8_mhjh_usergrade")
mess="�ٸ����˺�������ʱ�ѵ����˷�"&Request("name")&"ն����!!!!"
if usercorp="�ٸ�" and usergrade>=lockipright then
chatroomsn=session("yx8_mhjh_userchatroomsn")
shi=0.0416*1
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
connstr=Application("yx8_mhjh_connstr")
conn.open connstr
sql="update �û� set ����¼ʱ��=now()+"&shi&",״̬='����',�ȼ�=�ȼ�-2,����=����-50000,����=����-20000,����=����-10000,��ɱ=��ɱ+1 WHERE ����='"&Request("name")&"'"
conn.execute sql
sql="insert into Ӣ����(����,ʱ��,����,����) values ('"&Request("name")&"',now(),'"&username&"','�η���ն')"
conn.execute sql
sql="insert into �����¼(��������,������Ա,��������,����ԭ��,����ʱ��,������) values ('ն��','"&username&"','"&Request("name")&"','�η���ն',now(),'ħ�ý���')"
conn.execute sql
talkarr=Application("yx8_mhjh_talkarr")
talkpoint=Application("yx8_mhjh_talkpoint")
dim newtalkarr(600)
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
newtalkarr(599)="<font color=red>����Ϣ��</font><b><font color=red>"&mess&"</font></b>"
newtalkarr(600)=chatroomsn
Application("yx8_mhjh_talkarr")=newtalkarr
end if
conn.Close             
set conn=nothing  
%>
<head>
<LINK
href="../style.css" rel=stylesheet>
</head>
<body background='../chatroom/bg1.gif'>
<div align="center">
  <center>
<table width=338 bordercolorlight="#000000" border="1" cellspacing="0" cellpadding="0" bordercolordark="#FFFFFF">
<tr><td width="330">
<p align=center style='font-size:14;color:red'><%=mess%></p>
<p align=center><a href="../STREET/PRISON.ASP">����</a></p>
</td></tr>
</table>
  </center>
</div>
</body>
