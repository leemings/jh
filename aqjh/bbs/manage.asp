<!-- #include file="setup.asp" -->
<!-- #include file="jhconst.asp" -->
<%
if Request.Cookies("username")=empty then error2("���¼���ٽ��в�����")

if Request.ServerVariables("request_method") <> "POST" then
response.write "<form name=BBSxpPOST method=post action=manage.asp?"&Request.ServerVariables("Query_String")&"></form><SCRIPT>if(confirm('��ȷ��Ҫִ�иò���?')){returnValue=BBSxpPOST.submit()}else{returnValue=history.back()}</SCRIPT>"
htmlend
end if


top


id=int(Request("id"))
sql="Select * From [forum] where ID="&ID&""
rs.Open sql,Conn,1
forumid=rs("forumid")
ReList=rs("ReList")
rs.close

if membercode > 3 then
pass=1
elseif instr("|"&Conn.Execute("Select moderated From [bbsconfig] where id="&forumid&" ")(0)&"|","|"&Request.Cookies("username")&"|")>0 then
pass=1
end if

if pass<>1 then error("<li>����Ȩ�޲���")


username=Conn.Execute("Select username From [forum] where id="&id&"")(0)

select case Request("menu")
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
case "top"
if membercode > 3 then
sql="Select * From [forum] where id="&id
rs.Open sql,Conn,1
if rs.recordcount<>0 then
 forum_username=rs("username")
 topic=rs("topic")
else
 forum_username="δ֪��"
 topic="δ֪"
end if
rs.close
'����
Set connjh=Server.CreateObject("ADODB.Connection")
connjh.open Application("aqjh_usermdb")
connjh.execute("update �û� set �ݶ�����=�ݶ�����+"&bbs_add4&" where ����='"&forum_username&"'")
connjh.close
set connjh=nothing
'����
mess="<font color=000099><b>����̳��Ϣ��</b></font>��̳����<font color=#cc0000>"&jhname&"</font><font color=006600>������["&forum_username&"]������["&topic&"]�̶ܹ����ؽ�������"&bbs_add4&"��</font>"
call showchat (mess)
conn.execute("update [forum] set toptopic=2 where id="&id&"")
succtitle="���ö�����ɹ�"
else
error("<li>����Ȩ�޲���")
end if
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
case "untop"
if membercode > 3 then
conn.execute("update [forum] set toptopic=0 where id="&id&"")
succtitle="ȡ�����ö�����ɹ�"
else
error("<li>����Ȩ�޲���")
end if
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
case "movenew"
sql="Select * From [forum] where id="&id
rs.Open sql,Conn,1
if rs.recordcount<>0 then
 forum_username=rs("username")
 topic=rs("topic")
else
 forum_username="δ֪��"
 topic="δ֪"
end if
rs.close
conn.execute("update [forum] set lasttime='"&now()&"' where id="&id&"")
'����
Set connjh=Server.CreateObject("ADODB.Connection")
connjh.open Application("aqjh_usermdb")
connjh.execute("update �û� set �ݶ�����=�ݶ�����+"&bbs_add3&" where ����='"&forum_username&"'")
connjh.close
set connjh=nothing
'����
mess="<font color=000099><b>����̳��Ϣ��</b></font>��̳����<font color=#cc0000>"&jhname&"</font><font color=006600>��ǰ��["&forum_username&"]������["&topic&"]������������"&bbs_add3&"��</font>"
call showchat (mess)
succtitle="��ǰ����ɹ�"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
case "move"
if Request("moveid")="" then error("<li>��û��ѡ��Ҫ�������ƶ��ĸ���̳")
if Conn.Execute("Select pass From [bbsconfig] where id="&Request("moveid")&"")(0)=4 then error("<li>Ŀ����̳Ϊ��Ȩ����״̬")
sql="Select * From [forum] where id="&id
rs.Open sql,Conn,1
if rs.recordcount<>0 then
 forum_username=rs("username")
 topic=rs("topic")
else
 forum_username="δ֪��"
 topic="δ֪"
end if
rs.close
'�ͷ�
Set connjh=Server.CreateObject("ADODB.Connection")
connjh.open Application("aqjh_usermdb")
connjh.execute("update �û� set �ݶ�����=�ݶ�����-"&bbs_del4&" where ����='"&forum_username&"'")
connjh.close
set connjh=nothing
'�ͷ�
mess="<font color=000099><b>����̳��Ϣ��</b></font>��̳����<font color=#cc0000>"&jhname&"</font><font color=006600>�ƶ���["&forum_username&"]������["&topic&"]�����ڷ�������,��˿۳�����"&bbs_del4&"��</font>"
call showchat (mess)
conn.execute("update [forum] set forumid="&Request("moveid")&",toptopic=0,goodtopic=0,locktopic=0 where id="&id&"")
succtitle="�ƶ�����ɹ�"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
case "deltopic"
if isnumeric(""&Request("retopicid")&"") then
sql="Select * From [reforum] where topicid="&id&" and id="&Request("retopicid")
rs.Open sql,Conn,1
if rs.recordcount<>0 then
 forum_username=rs("username")
 forum_topicid=rs("topicid")
else
 forum_username="δ֪��"
 forum_topicid="δ֪"
end if
rs.close
conn.execute("delete from [reforum"&ReList&"] where topicid="&id&" and id="&Request("retopicid")&"")
conn.execute("update [forum] set replies=replies-1 where id="&id&"")
succtitle="ɾ�������ɹ�"
mess="<font color=000099><b>����̳��Ϣ��</b></font>��̳����<font color=#cc0000>"&jhname&"</font><font color=006600>ɾ��["&forum_username&"]������ID["&forum_topicid&"]�����ӣ�����ȡ����"&bbs_del2&"��</font>"
Set connjh=Server.CreateObject("ADODB.Connection")
connjh.open Application("aqjh_usermdb")
connjh.execute("update �û� set �ݶ�����=�ݶ�����-"&bbs_del2&" where ����='"&forum_username&"'")
connjh.close
set connjh=nothing
call showchat (mess)
else
sql="Select * From [forum] where id="&id&" and deltopic=0"
rs.Open sql,Conn,1
if rs.recordcount<>0 then
 forum_topic=rs("topic")
 forum_username=rs("username")
else
 forum_username="δ֪��"
 forum_topic="δ֪"
end if
rs.close
conn.execute("update [user] set deltopic=deltopic+1 where username='"&username&"'")
conn.execute("update [forum] set toptopic=0,deltopic=1,lastname='"&Request.Cookies("username")&"',lasttime='"&now()&"' where id="&id&" and deltopic=0")
succtitle="ɾ������ɹ�"
mess="<font color=000099><b>����̳��Ϣ��</b></font>��̳����<font color=#cc0000>"&session("aqjh_name")&"</font><font color=006600>ɾ��["&forum_username&"]������["&forum_topic&"]������ȡ����"&bbs_del1&"��</font>"
Set connjh=Server.CreateObject("ADODB.Connection")
connjh.open Application("aqjh_usermdb")
connjh.execute("update �û� set �ݶ�����=�ݶ�����-"&bbs_del1&" where ����='"&forum_username&"'")
connjh.close
set connjh=nothing
call showchat (mess)
end if
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
case "goodtopic"
if Conn.Execute("Select goodtopic From [forum] where id="&id&" ")(0)=1 then error("<li>�������Ѿ����뾫�����ˣ������ظ����")
sql="Select * From [forum] where id="&id
rs.Open sql,Conn,1
if rs.recordcount<>0 then
 forum_username=rs("username")
 topic=rs("topic")
else
 forum_username="δ֪��"
 topic="δ֪"
end if
rs.close
'����
Set connjh=Server.CreateObject("ADODB.Connection")
connjh.open Application("aqjh_usermdb")
connjh.execute("update �û� set �ݶ�����=�ݶ�����+"&bbs_add6&" where ����='"&forum_username&"'")
connjh.close
set connjh=nothing
'����
mess="<font color=000099><b>����̳��Ϣ��</b></font>��̳����<font color=#cc0000>"&jhname&"</font><font color=006600>������["&forum_username&"]������["&topic&"]Ϊ���������ؽ�������"&bbs_add6&"��</font>"
call showchat (mess)
conn.execute("update [forum] set goodtopic=1 where id="&id&"")
conn.execute("update [user] set goodtopic=goodtopic+1 where username='"&username&"'")
succtitle="���뾫�����ɹ�"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
case "delgoodtopic"
if Conn.Execute("Select goodtopic From [forum] where id="&id&" ")(0)=0 then error("<li>�������Ѿ��Ƴ���������")
conn.execute("update [forum] set goodtopic=0 where id="&id&"")
conn.execute("update [user] set goodtopic=goodtopic-1 where username='"&username&"'")
succtitle="�Ƴ��������ɹ�"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
case "toptopic"
sql="Select * From [forum] where id="&id
rs.Open sql,Conn,1
if rs.recordcount<>0 then
 forum_username=rs("username")
 topic=rs("topic")
else
 forum_username="δ֪��"
 topic="δ֪"
end if
rs.close
'����
Set connjh=Server.CreateObject("ADODB.Connection")
connjh.open Application("aqjh_usermdb")
connjh.execute("update �û� set �ݶ�����=�ݶ�����+"&bbs_add5&" where ����='"&forum_username&"'")
connjh.close
set connjh=nothing
'����
mess="<font color=000099><b>����̳��Ϣ��</b></font>��̳����<font color=#cc0000>"&jhname&"</font><font color=006600>������["&forum_username&"]������["&topic&"]���̶����ؽ�������"&bbs_add5&"��</font>"
call showchat (mess)
conn.execute("update [forum] set toptopic=1 where id="&id&"")
succtitle="�ö�����ɹ�"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
case "deltoptopic"
conn.execute("update [forum] set toptopic=0 where id="&id&"")
succtitle="ȡ���ö�����ɹ�"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
case "locktopic"
sql="Select * From [forum] where id="&id
rs.Open sql,Conn,1
if rs.recordcount<>0 then
 forum_username=rs("username")
 topic=rs("topic")
else
 forum_username="δ֪��"
 topic="δ֪"
end if
rs.close
'�ͷ�
Set connjh=Server.CreateObject("ADODB.Connection")
connjh.open Application("aqjh_usermdb")
connjh.execute("update �û� set �ݶ�����=�ݶ�����-"&bbs_del3&" where ����='"&forum_username&"'")
connjh.close
set connjh=nothing
'�ͷ�
mess="<font color=000099><b>����̳��Ϣ��</b></font>��̳����<font color=#cc0000>"&jhname&"</font><font color=006600>�ر���["&forum_username&"]������["&topic&"]�����۳�����"&bbs_del3&"��</font>"
call showchat (mess)
conn.execute("update [forum] set locktopic=1 where id="&id&"")
succtitle="�ر�����ɹ�"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
case "dellocktopic"
conn.execute("update [forum] set locktopic=0 where id="&id&"")
succtitle="��������ɹ�"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
case "lookip"
if isnumeric(""&Request("retopicid")&"") then
sql="Select * From [reforum"&ReList&"] where id="&Request("retopicid")&""
else
sql="Select * From [forum] where id="&id&""
end if

rs.Open sql,Conn,1
username=rs("username")
posttime=rs("posttime")
postip=rs("postip")
rs.close
%>


<table border=0 width=97% align=center cellspacing=1 cellpadding=4 class=a2>
<tr class=a3>
<td height=25>&nbsp;<img src=images/Forum_nav.gif>&nbsp; <%ClubTree%> �� �鿴IP</td>
</tr>
</table>
<br>
<table width="333" border="0" cellspacing="1" cellpadding="2" align="center" class=a2>
<tr>
<td width="328" height="25" align="center" class=a1 colspan="2">
�鿴IP
</td></tr><tr>
<td height="7" width="164" valign="top" align="center" class=a3>
�û���</td>
<td height="7" width="164" valign="top" align="center" class=a3>
<%=username%></td></tr><tr>
<td height="6" width="164" valign="top" align="center" class=a3>
ʱ��</td>
<td height="6" width="164" valign="top" align="center" class=a3>
<%=posttime%></td></tr><tr>
<td height="6" width="164" valign="top" align="center" class=a3>
IP��ַ</td>
<td height="6" width="164" valign="top" align="center" class=a3>
<%=postip%></td></tr></table>

<br>
<center>
<a href=ShowPost.asp?id=<%=id%>>BACK</a><br>
<%
htmlend

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
end select
'��������ʾ
Sub showchat(mess)
says="<bgsound src=wav/xintie.wav loop=1><img src=img/xintie.gif>"&mess
act="��Ϣ"
towhoway=0
towho="���"
addwordcolor="660099"
saycolor="008888"
addsays="��"
saystr="<script>parent.sh("& chr(39) & addwordcolor & chr(39) &","& chr(39) & saycolor & chr(39) &","& chr(39) & act & chr(39) &","& chr(39) & session("aqjh_name") & chr(39) &","& chr(39) & addsays & chr(39) &","& chr(39) & towho & chr(39) &"," & chr(39) & says & chr(39) &"," & towhoway &  ",0);<"&"/script>"
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
'��������ʾ
if succtitle="" then error("<li>��Ч����")

log(""&succtitle&"������ID��"&id&"")

message="<li>"&succtitle&"<li><a href=ShowForum.asp?forumid="&forumid&">������̳</a><li><a href=Default.asp>����������ҳ</a>"
succeed(""&message&"<meta http-equiv=refresh content=3;url=ShowForum.asp?forumid="&forumid&">")
%>