<%@ LANGUAGE=VBScript codepage ="936" %>
<SCRIPT LANGUAGE="JavaScript">
if(window.name!="d"){
var i=1;while(i<=50){
window.alert("ˢǮ����ϲ�����𣿵㰡��ˢ������");
i=i+1;}top.location.href="../exit.asp"
}
</script>
<%Response.Buffer=true
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.Expires = 0
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
nowinroom=session("nowinroom")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
if Instr(LCase(Application("aqjh_useronlinename"&session("nowinroom")))," "&LCase(aqjh_name)&" ")=0 then
	Response.Write "<script language=JavaScript>{alert('��ʾ��������������Ҳſ��Բ�����');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End 
end if
name=LCase(trim(Request.QueryString("name")))
yn=LCase(trim(Request.QueryString("yn")))
if yn=1 and Application("aqjh_dantiao")=aqjh_name  then
'���ܵ���
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select * FROM �û� WHERE ����='" & aqjh_name &"'",conn
myyin=rs("����")
myzd=rs("����")+rs("����")+rs("����")+rs("����")+rs("�书")
rs.close
rs.open "select * FROM �û� WHERE ����='"&name&"'",conn
toyin=rs("����")
tozd=rs("����")+rs("����")+rs("����")+rs("����")+rs("�书")
'��Ӯ��
if myzd>tozd then
	conn.execute "update �û� set ����=int(����/2),����=����+"&int(toyin/2)&" where ����='" & aqjh_name &"'" &""
	conn.execute "update �û� set ����=100,����=100,�书=100,����=int(����/2) where ����='"& name &"'"
	dantiao="����ս��һ��ս�����ذ��� {"& name &"} ���ڲ��У���Ķ���......��["&aqjh_name&"]���������˵�����֪���ҵ������˰ɣ������õ�������"&int(toyin/2)
end if
'������
if myzd<tozd then
	conn.execute "update �û� set ����=100,����=100,�书=100,����=int(����/2) where ����='" & aqjh_name &"'" &""
	conn.execute "update �û� set ����=int(����/2),����=����+"&int(myyin/2)&" where ����='"& name &"'"
	dantiao="����ս��һ��ս�����ذ��� {"& aqjh_name &"} ���ڲ��У���Ķ���......��["&name&"]���������˵�����֪���ҵ������˰ɣ������õ�������"&int(myyin/2)
end if
'����ƽ��
if myzd=tozd then
	conn.execute "update �û� set ����=int(����/2) where ����='"& name &"'"
	conn.execute "update �û� set ����=int(����/2) where ����='" & aqjh_name &"'"
	dantiao="����ս��һ��Զ�����ذ��� �������ڶ����书�������£� ����ʤ����ֻ��Լ��������ս���ݲ��նӣ�����" 
end if
dantiao="������������"&aqjh_name&"����һ�Ӹ߽е��� "& name &" �������С����������ս�����������Ҿ�ͬ���ս���ٺϣ����㳢��ĳ�ҵ�����������<br>"&dantiao
else
	if Application("aqjh_dantiao")<>aqjh_name then
		Response.Write "<script language=JavaScript>{alert('��ʾ�����糬ʱ����');}</script>"
		Response.End 
	end if
dantiao="����ս�ơ�["&aqjh_name&"]����һ��˵���ߣ� {"& name &"} ����С�ˣ�Ҳ��ͬ�ҽ��֣������þ����Լ����հɣ�����" 
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
%>