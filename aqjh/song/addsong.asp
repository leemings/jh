<!-- #include file="conn.asp"-->
<!--#include file="../showchat.asp"-->
<%
id = request.form("id")
name = request.form("name")
songname = request.form("songname")
songurl = request.form("songurl")
zhufu = request.form("zhufu")
toname = request.form("toname")
n=Year(date())
y=Month(date())
r=Day(date())
s=Hour(time())
f=Minute(time())
m=Second(time())
if len(y)=1 then y="0" & y
if len(r)=1 then r="0" & r
if len(s)=1 then s="0" & s
if len(f)=1 then f="0" & f
if len(m)=1 then m="0" & m
sj=n & "-" & y & "-" & r & " " & s & ":" & f & ":" & m
t=s & ":" & f & ":" & m
dim sql
set rs=server.createobject("adodb.recordset")
sql="select ���� from �û� where ����='"&toname&"'"
rs.open sql,conn,1,1
if instr(toname,"���")>0 then
		toname="���"
		isfull=true
else
if rs.eof and rs.bof then
call smsg("û��[" & toname & "]����û���","-1")
end if
end if
rs.close
if len(zhufu)<2 or len(zhufu)>250 then 
call smsg("�Բ���!ף������������2������,���250������!","-1")
end if
if Instr(songname,"������")>0 or Instr(songname,"��")>0 or Instr(songname,"�����")>0 or Instr(songname,"���")>0 or Instr(songname,"����")>0 or Instr(songname,"����")>0 or Instr(songname,"���")>0 or Instr(songname,"�ҿ�")>0 or Instr(songname,"����")>0 or Instr(songname,"��")>0 or Instr(zhufu,"B")>0 or Instr(zhufu,"b")>0 or Instr(zhufu,"��")>0 or Instr(zhufu,"������")>0 or Instr(zhufu,"��")>0 or Instr(zhufu,"�����")>0 or Instr(zhufu,"���")>0 or Instr(zhufu,"����")>0 or Instr(zhufu,"����")>0 or Instr(zhufu,"���")>0 or Instr(zhufu,"�ҿ�")>0 or Instr(zhufu,"����")>0 or Instr(zhufu,"��")>0 or Instr(zhufu,"B")>0 or Instr(zhufu,"b")>0 or Instr(zhufu,"��")>0 then response.Redirect "err.asp"
sql="select * from music"
rs.open sql,conn,1,3
rs.addnew
rs("name")=name
rs("songname")=songname
rs("songurl")=songurl
rs("zhufu")=zhufu
rs("toname")=toname
rs.update
says="<img src=../chat/img/music.gif><font color=blue>"&name&"</font>Ϊ<font color=blue>"&toname&"</font><font color=#ff0000>�ͳ�һ��<font color=blue>["&songname&"]</font>����˵:<font color=blue>["&zhufu&"]</font></font> <input type=button value=���� onclick=javascript:window.open("&chr(34)&"../song/play.asp?id="&rs("id")&chr(34)&","&chr(34)&"mp3"&chr(34)&","&chr(34)&"Status=no,scrollbars=yes,resizable=no,width=250,height=80,top=20,left=20"&chr(34)&") style="&chr(34)&"background-color:#86A231;color:FFFFFF;border: 1 double"&chr(34)&"><font class=t>("&t&")</font>"
call showchat(says)
call smsg("��ϲ[" & name & "],���ɹ�!","1")
rs.colse
conn.close
Sub smsg(str1, str2)
 set rs=nothing
 conn.close
 set conn=nothing
 onend = 0
 Select Case str2
 Case "-1","-2","-3" '����
    str2 = "location.href = 'javascript:history.go(-1)';"
 Case "1" '�ر�
    str2 = "window.close();"
 Case "2" '������
    str2 = ""
 Case "3" '����
    str2 = ""
    onend = 1
 Case Else '���˵�ָ���ļ�
    str2 = "location.href = '" & str2 & "';"
 End Select
 Response.Write "<script language=JavaScript>{alert('��ʾ��" & str1 & "');" & str2 & "}</script>"
 If onend = 0 Then
    Response.End
 End If
End Sub
%>