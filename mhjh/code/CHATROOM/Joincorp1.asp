<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
un=session("yx8_mhjh_username")
co=session("yx8_mhjh_usercorp")
if un="" then Response.Redirect "../error.asp?id=016"
mg=Request.QueryString("mg")
if instr(mg,"'") then Response.Redirect "../error.asp?id=024"
chatroomsn=Session("yx8_mhjh_userchatroomsn")
set conn=server.CreateObject("adodb.connection")
conn.Open Application("yx8_mhjh_connstr")
set rst=server.CreateObject("adodb.recordset")
rst.Open "select * from ���� where ����='"&mg&"'",conn
if rst.EOF or rst.BOF then
joincorp="���޷������"&mg&"����Ϊ�����в�û��������ɡ�"
else
cond=rst("��������")
rst.Close
rst.Open "select * from �û� where ����='"&un&"' and "&cond,conn
if rst.EOF or rst.BOF then
joincorp="������������,"&mg&"�ܾ����������ĵ��ӡ�"
else
if co="��" then
conn.Execute "update �û� set ����='"&mg&"',���='��' where ����='"&un&"'"
onlinelist=Application("yx8_mhjh_onlinelist")
onlinelistubd=ubound(onlinelist)
for i=1 to onlinelistubd step 7
if onlinelist(i)=un then
onlinelist(i+2)=mg
exit for
end if
next
Application.Lock
Application("yx8_mhjh_onlinelist")=onlinelist
Application.UnLock
erase onlinelist
session("yx8_mhjh_usercorp")=mg
joincorp="��ӭ������"&mg
talkarr=Application("yx8_mhjh_talkarr")
talkpoint=clng(Application("yx8_mhjh_talkpoint"))
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
newtalkarr(592)=2
newtalkarr(593)=0
newtalkarr(594)=un
newtalkarr(595)="���"
newtalkarr(596)=""
newtalkarr(597)="#660099"
newtalkarr(598)="#660099"
newtalkarr(599)="<font color=FF0000>�����롿</font>##�ɹ��ļ���"&mg&"��������Ϊ##��,ϣ�����ҵ���һ�����ʺϵ�����,��ףԸ##��Ϊ����"&mg&"��ս��!<font class=timsty>��"&time()&"��<\/font>"
newtalkarr(600)=chatroomsn
Application.Lock
Application("yx8_mhjh_talkpoint")=talkpoint+1
Application("yx8_mhjh_talkarr")=newtalkarr
Application.UnLock
erase newtalkarr
erase talkarr
elseif co=mg then
joincorp="���Ѿ���"&co&"�ĵ����ˣ������ټ���"&mg
else
joincorp="��̾ѽ�ɱ���##��Ϊ"&co&"���ӣ���Ȼ����Ͷ"&mg
end if
end if
end if
rst.Close
set rst=nothing
conn.close
set conn=nothing
%>
<html>
<head>
<title></title>
<link rel=stylesheet href='css3.css'>
<script language=javascript>
setTimeout('javascript:window.close()',1000);
</script>
</head>
<body oncontextmenu=self.event.returnValue=false background="bg1.gif" >
<div align=center>
<hr noshade size="1" color=red>
�������Զ�����<br>
<input type=button value='����' onclick='javascript:window.close()' id=button1 name=button1>
</div>
<%=joincorp%>
</body>
</html>
