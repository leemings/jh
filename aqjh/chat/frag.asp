<%@ LANGUAGE=VBScript codepage ="936" %>
<SCRIPT LANGUAGE="JavaScript">if(window.name!="d"){var i=1;while(i<=50){window.alert("ˢǮ����ϲ�����𣿵㰡��ˢ������");
i=i+1;}top.location.href="../exit.asp"}</script>
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
fromname=LCase(trim(request.querystring("fromname")))
toname=LCase(trim(request.querystring("toname")))
if toname<>aqjh_name then
	if fromname<>toname then
		Response.Write "<script language=JavaScript>{alert('������ʲô���˼�"&fromname&"Ҳû��Ҫ�����������˵��');}</script>"
		Response.End
	end if
		Response.Write "<script language=JavaScript>{alert('��ʲô����"&fromname&" �����������������⻹�����Լ���������');}</script>"
		Response.End
	else
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select �Ա� from �û� where ����='"&aqjh_name&"'",conn,3,3
If Rs.Bof OR Rs.Eof then
	rs.close	
	set rs=nothing	
	conn.close	
	set conn=nothing	
	Response.Write "<script language=JavaScript>{alert('���������˲����ڣ�������������Ϊ�����վ����ӳ��');}</script>"	'��4
	Response.End	
end if
rs.close
rs.open "select ���,����,����,�Ա� FROM �û� WHERE ����='"&aqjh_name&"'",conn
tosex=rs("�Ա�")
Response.Write "<script language=JavaScript>{alert('Ҫ�ǵó���������Ҫ��Ȼ���ѵ�5000�����ร�');}</script>"
frag="��ҹ��ʦ����<font color=red>��"& fromname &"��</font>�����棬<font color=red>��"& aqjh_name &"��</font>������������ֳ�����.����ͷ2�������治��׬^_^<font color=red>��"& fromname &"��</font>����������ϴ��ȥ.."
	   	 conn.execute "update �û� set  ����=����-500000,���=0,����=����-1000 where ����='"&aqjh_name&"'"
	     conn.execute "update �û� set  ����=����+500000,����=����+1000 where ����='"&fromname&"'"
set rs=nothing
conn.close
set conn=nothing
Application.Lock
says="<font color=red><b>����ҹ����Ϣ��</b></font>"&frag	
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