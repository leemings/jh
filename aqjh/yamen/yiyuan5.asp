<%@ LANGUAGE=VBScript codepage ="936" %>
<%

aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
nowinroom=session("nowinroom")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
my=aqjh_name
id=request("id")
'У���û�
sql="SELECT * FROM �û� WHERE ����='" & my & "'"
Set Rs=conn.Execute(sql)
If Rs.Bof OR Rs.Eof Then
mess=my&"���㲻�ܲ�����"
else
if rs("����")="�ٸ�" and rs("����")>"8" then
UID=rs("ID")
sql="select * from �û� where (״̬='������' or ״̬='÷��' or ״̬='�����' or ״̬='��ʬ��') and ����<>'�ٸ�' and id=" & id
set rs=conn.execute(sql)
if not rs.eof and not rs.bof then
sql="update �û� set ״̬='����',��¼=now() where id=" & id
conn.execute sql
mess=my&"�����Ѿ������Ѳ�ҽ���ˣ�"
pos=application("chatpos")
chats=application("chats")
sj=formatdatetime(time(),3)
sj="<font color=#FF00FF size=1>(" & sj & ")</font>"
says="<font color=red><b>" & rs("����") & "����ҽ���ֻش��κ���</b>(��ҽID="&UID&")</font>"			'��������
says=replace(says,"'","\'")
says=replace(says,chr(34),"\"&chr(34))
act="��Ϣ"
towhoway=0
towho="���"
addwordcolor="660099"
saycolor="008888"
addsays="��"
saystr="<script>parent.sh("& chr(39) & addwordcolor & chr(39) &","& chr(39) & saycolor & chr(39) &","& chr(39) & act & chr(39) &","& chr(39) & aqjh_name & chr(39) &","& chr(39) & addsays & chr(39) &","& chr(39) & towho & chr(39) &"," & chr(39) & says & chr(39) &"," & towhoway &  ","& nowinroom & ");<"&"/script>"
addmsg saystr
Function Yushu(a)
	Yushu=(a and 31)
End Function
Sub AddMsg(Str)
Application.Lock()
Application("SayCount")=Application("SayCount")+1
i="SayStr"&YuShu(Application("SayCount"))
Application(i)=Str
Application.UnLock()
End Sub
else
mess="û��������˻��Ǵ����ǹٸ�����"
end if
else
mess=mye&"�����޴�Ȩ��������"
end if
end if
rs.close
set rs=nothing
conn.close
set conn=nothing
%>

<head>
<style>td           { font-size: 14 }
</style>
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title><%=Application("aqjh_chatroomname")%></title></head>

<body background="../jhimg/bk_hc3w.gif">
<div align="center">�����ɹ�^_^<br>
</div>
<table border="1" bgcolor="#BEE0FC" align="center" width="350" cellpadding="10"
cellspacing="13">
<tr>
<td bgcolor="#CCE8D6">
<table width="100%">
<tr>
<td>
<p align="center" style="font-size:14;color:red"><%=mess%></p>
<p align="center"><a href="yiyuan.asp">����</a></p>
</td>
</tr>
</table>
</td>
</tr>
</table></body>