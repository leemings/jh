<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Buffer=true
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.Expires = 0
sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
nowinroom=session("nowinroom")
if sjjh_name="" then Response.Redirect "../error.asp?id=440"
if Instr(LCase(Application("sjjh_useronlinename"&session("nowinroom")))," "&LCase(sjjh_name)&" ")=0 then
	Response.Write "<script language=JavaScript>{alert('��ʾ��������������Ҳſ��Բ�����');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End 
end if
name=LCase(trim(Request.QueryString("name")))
yn=LCase(trim(Request.QueryString("yn")))
if yn=1 and Application("sjjh_dantiao")=sjjh_name  then
'���ܵ���
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "select * FROM �û� WHERE ����='" & sjjh_name &"'",conn
myyin=rs("����")
myzd=rs("����")+rs("����")+rs("����")+rs("����")+rs("�书")
rs.close
rs.open "select * FROM �û� WHERE ����='"&name&"'",conn
toyin=rs("����")
tozd=rs("����")+rs("����")+rs("����")+rs("����")+rs("�书")
'��Ӯ��
if myzd>tozd then
	conn.execute "update �û� set ����=int(����/2),����=����+"&int(toyin/2)&" where ����='" & sjjh_name &"'" &""
	conn.execute "update �û� set ����=100,����=100,�书=100,����=int(����/2) where ����='"& name &"'"
	dantiao="����ս��һ��ս�����ذ��� {"& name &"} ���ڲ��У���Ķ���......��["&sjjh_name&"]���������˵�����֪���ҵ������˰ɣ������õ�������"&int(toyin/2)
end if
'������
if myzd<tozd then
	conn.execute "update �û� set ����=100,����=100,�书=100,����=int(����/2) where ����='" & sjjh_name &"'" &""
	conn.execute "update �û� set ����=int(����/2),����=����+"&int(myyin/2)&" where ����='"& name &"'"
	dantiao="����ս��һ��ս�����ذ��� {"& sjjh_name &"} ���ڲ��У���Ķ���......��["&name&"]���������˵�����֪���ҵ������˰ɣ������õ�������"&int(myyin/2)
end if
'����ƽ��
if myzd=tozd then
	conn.execute "update �û� set ����=int(����/2) where ����='"& name &"'"
	conn.execute "update �û� set ����=int(����/2) where ����='" & sjjh_name &"'"
	dantiao="����ս��һ��Զ�����ذ��� �������ڶ����书�������£� ����ʤ����ֻ��Լ��������ս���ݲ��նӣ�����" 
end if
dantiao="������������"&sjjh_name&"���һ�Ӹ߽е��� "& name &" �������С����������ս�����������Ҿ�ͬ���ս���ٺϣ����㳢��ĳ�ҵ�����������<br>"&dantiao
else
	if Application("sjjh_dantiao")<>sjjh_name then
		Response.Write "<script language=JavaScript>{alert('��ʾ�����糬ʱ����');}</script>"
		Response.End 
	end if
dantiao="����ս�ơ�["&sjjh_name&"]����һ��˵���ߣ� {"& name &"} ����С�ˣ�Ҳ��ͬ�ҽ��֣������þ����Լ����հɣ�����" 
end if
says=dantiao		'��������

says=replace(says,chr(39),"\'")
says=replace(says,chr(34),"\"&chr(34))
act="��Ϣ"
towhoway=0
towho="���"
addwordcolor="660099"
saycolor="008888"
addsays="��"
saystr="<script>parent.sh("& chr(39) & addwordcolor & chr(39) &","& chr(39) & saycolor & chr(39) &","& chr(39) & act & chr(39) &","& chr(39) & sjjh_name & chr(39) &","& chr(39) & addsays & chr(39) &","& chr(39) & towho & chr(39) &"," & chr(39) & says & chr(39) &"," & towhoway &  ","& nowinroom & ");<"&"/script>"
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
%>
