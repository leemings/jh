<%
Sub showchat(mess)
says=mess   '��������
act="��Ϣ"
towhoway=0
towho="���"
addwordcolor="660099"
saycolor="008888"
addsays="��"
saystr="<script>parent.sh("& chr(39) & addwordcolor & chr(39) &","& chr(39) & saycolor & chr(39) &","& chr(39) & act & chr(39) &","& chr(39) & session("aqjh_name") & chr(39) &","& chr(39) & addsays & chr(39) &","& chr(39) & towho & chr(39) &"," & chr(39) & says & chr(39) &"," & towhoway &  ","&session("nowinroom")&");<"&"/script>"
addmsg saystr
end sub
Function Yushu1(a)
 Yushu1=(a and 31)
End Function
Sub AddMsg(Str)
Application.Lock()
Application("SayCount")=Application("SayCount")+1
i="SayStr"&YuShu1(Application("SayCount"))
Application(i)=Str
Application.UnLock()
End Sub
%>
<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
server_v1=Cstr(Request.ServerVariables("HTTP_REFERER"))
server_v2=Cstr(Request.ServerVariables("SERVER_NAME"))
if mid(server_v1,8,len(server_v2))<>server_v2 then
response.write "<br><br><center><table border=1 cellpadding=20 bordercolor=black bgcolor=#EEEEEE width=450>"
response.write "<tr><td style='font:9pt Verdana'>"
response.write "���ύ��·�����󣬽�ֹ��վ���ⲿ�ύ�����벻Ҫ�Ҹò�����"
response.write "</td></tr></table></center>"
response.end
end if

if Session("aqjh_name")="" then Response.Redirect "../error.asp?id=440"
chatroomsn=session("hxf_u_roomin")
nickname=Session("aqjh_name")
grade=Session("aqjh_grade")
myid=session("aqjh_id")

animalname=Request.Form ("animalname")
newname=Request.Form ("newname")
if newname="" or animalname="" then
Response.Write "<script language=JavaScript>{alert('  �Բ���\n  ��û�д�����������������������ȡ��Ϊ�գ�����\n  �� [ȷ��] �� �أ�');parent.history.go(-1);}</script>"
Response.End 
end if

namelen=0
for i=1 to len(newname)
zh=mid(newname,i,1)
zhasc=asc(zh)
if zhasc<0 then
namelen=namelen+2
else
namelen=namelen+1
if CStr(server.URLEncode(zh))<>CStr(zh) then Response.Redirect "error.asp?id=120"
end if
next
if namelen>10 then 
Response.Write "<script language=JavaScript>{alert('  �Բ���\n  ����̫�������5�����֣��뷵�ظ���������\n  �� [ȷ��] �� �أ�');parent.history.go(-1);}</script>"
Response.end
end if

namelen=0
if InStr(LCase(newname),"fuck")<>0 or InStr(LCase(newname),"sex")<>0 or InStr(newname,"��")<>0 or InStr(newname,"��")<>0 or InStr(newname,"�")<>0 or InStr(newname,"��")<>0 or InStr(newname,"��")<>0 and InStr(newname,"��")<>0 or InStr(newname,"��")<>0 or InStr(newname,"ɫ")<>0 and InStr(newname,"��")<>0 or InStr(newname,"ɫ")<>0 and InStr(newname,"��")<>0 or InStr(newname,"��")<>0 and InStr(newname,"��")<>0 or InStr(newname,"��")<>0 and InStr(newname,"��")<>0 or InStr(newname,"��")<>0 and InStr(newname,"��")<>0 or InStr(newname,"��")<>0 and InStr(newname,"��")<>0 or InStr(newname,"��")<>0 and InStr(newname,"��")<>0 or InStr(newname,"��")<>0 or InStr(newname,"��")<>0 or InStr(newname,"��")<>0 then
Response.Write "<script language=JavaScript>{alert('  �Բ���\n  ���ֺ��в������ۣ��뷵�ظ���������\n  �� [ȷ��] �� �أ�');parent.history.go(-1);}</script>"
Response.End 
end if

Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
connstr=Application("aqjh_usermdb")
conn.open connstr
rs.Open ("select * from myanimal where animalname='"&animalname&"' and username='"&nickname&"' and rest=0"),conn
if rs.BOF or rs.EOF then
Response.Write "<script language=JavaScript>{alert('  �Բ���\n  ��û�д�������������\n  �� [ȷ��] �� �أ�');parent.history.go(-1);}</script>"
rs.Close
set rs=nothing
conn.Close
set conn=nothing
Response.End 
else
id=rs("id")
rs.Close
rs.Open ("select * from myanimal where animalname='"&newname&"' and rest=0 and username='"&nickname&"'"),conn
if not(rs.BOF or rs.EOF) then
Response.Write "<script language=JavaScript>{alert('  �Բ���\n  �����д������ֵ�����������\n  �� [ȷ��] �� �أ�');parent.history.go(-1);}</script>"
rs.Close
set rs=nothing
conn.Close
set conn=nothing
Response.End 
end if
conn.Execute ("update myanimal set animalname='"&newname&"' where username='"&nickname&"' and rest=0 and id="&id)
msg="<font color=blank>["&nickname&"]���Լ�������<font color=red>["&animalname&"]</font>ȡ��һ������������<font color=red>["&newname&"]</font></font>��"
end if
rs.Close
set rs=nothing
conn.Close
set conn=nothing

says=msg   '��������
call showchat(says)

Response.Write "<script language=JavaScript>{alert('  ��ϲ����\n  ����������"&animalname&"����Ϊ"&newname&"������\n  �� [ȷ��] �� �أ�');parent.ps.location.href='about:blank';parent.f3.location.href='myanimal.asp';}</script>"
%>

</body>
</html>
