<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Buffer=true
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
nowinroom=session("nowinroom")
'#####���䴦��#####
aqjh_roominfo=split(Application("aqjh_room"),";")
chatinfo=split(aqjh_roominfo(session("nowinroom")),"|")
if chatinfo(7)<>0 then
	Response.Write "<script language=JavaScript>{alert('��["&chatinfo(0)&"]���䲻��������������');}</script>"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "SELECT ͨ��,����,ɱ����,����,����,����ʱ��,��Ա�ȼ� FROM �û� WHERE ����='" & aqjh_name &"'",conn
hydj=rs("��Ա�ȼ�")
if rs("����")=false and aqjh_jhdj>100 and hydj=0 then
	Response.Write "<script language=JavaScript>{alert('��ĵȼ��Ѿ�����100�ˣ�ϵͳ���Ʊ�����');}</script>"
	Response.End
end if
if rs("����")="����" and rs("����")=False then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ������ɱ�ֲ����Ա�����');}</script>"
	Response.End
end if
if rs("����")="����"  and rs("����")=False then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ�����Һ�û�б�Ҫ������');}</script>"
	Response.End
end if
if rs("ͨ��")=True  and rs("����")=False then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('�ӷ�Ҳ�뱣��������ʲô���ģ�');}</script>"
	Response.End
end if
if rs("ɱ����")>=int(Application("aqjh_killman")) and rs("����")=False then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('�������»���,���뱣������');}</script>"
	Response.End
end if
if rs("����")=Application("aqjh_baowuname") and rs("����")=False then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('�������н�������["& Application("aqjh_baowuname") &"]���ܽ�������������');}</script>"
	Response.End
end if
sj=DateDiff("n",rs("����ʱ��"),now())
if sj<2 then
	s=2-sj
	if rs("����")=true then
		Response.Write "<script language=JavaScript>{alert('���������������У����["&s&"����]�ٲ�����');}</script>"
		rs.close
		set rs=nothing	
		conn.close
		set conn=nothing
		Response.End
	else
		Response.Write "<script language=JavaScript>{alert('��ոս�����������["&s&"����]�ٲ�����');}</script>"
		rs.close
		set rs=nothing	
		conn.close
		set conn=nothing
		Response.End
	end if
end if
if rs("����")=true then
	conn.Execute "update �û� set ����=false,����ʱ��=now() where ����='" & aqjh_name &"'"	
	diaox="ѧ���гɣ�<img src=xx/gif/wg13.gif>���ڿ��Դ�չȭ�ź�ɱһ���ˣ�����������"
else
	conn.Execute "update �û� set ����=true,����ʱ��=now() where ����='" & aqjh_name &"'"
	diaox="Ϊ���ӱܳ��׷ɱ����������������������˭Ҳ���ˣ�"
end if
conn.close
set rs=nothing
set conn=nothing
says="<font color=#ff0000>��������Ϣ��" & aqjh_name & ""& diaox &"</font>"			'��������
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