<%@ LANGUAGE=VBScript%>
<SCRIPT LANGUAGE=JavaScript>if(window.name!='aqjh_win'){var i=1;while(i<=50){window.alert('ˢǮ����ϲ�����𣿵㰡��ˢ������');i=i+1;}top.location.href='../../exit.asp'}</script>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
nowinroom=session("nowinroom")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
chatbgcolor=Session("afa_chatbgcolor")
chatimage=Application("aqjh_chatimage")
if chatbgcolor="" then chatbgcolor="008888"
if Instr(LCase(Application("aqjh_useronlinename"&session("nowinroom")))," "&LCase(aqjh_name)&" ")=0 then
	Response.Write "<script Language=Javascript>alert('�㲻�ܽ��в��������д˲���������������ң�');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
types=Request.QueryString("typename")
id=Request.QueryString("id")
Select Case id
Case "1"
  classes="״Ԫ"
Case "2" 
   classes="����"
Case "3"
   classes="̽��"
End Select
if classes=null or types=null then
	Response.Write "<script language=JavaScript>{alert('��ʾ�����ݴ�������ҵĻ���');}</script>"
	Response.End 
end if
Select Case types
Case "gold"
  typeses="���"
Case "silver" 
   typeses="����"
Case "copper"
   typeses="ͭ��"
End Select
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select �书,����,����,����,�Ա�,����,���,���� from �û� where ����='"&aqjh_name&"'",conn
if rs.eof or rs.bof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ���㲻�ǽ������ˡ�');window.close();}</script>"
	Response.End 
end if
mywg=rs("�书")
mygj=rs("����")
myfy=rs("����")
mynl=rs("����")
mymp=rs("����")
mysf=rs("���")
mysex=rs("�Ա�")
mytl=rs("����")
if mywg>1000000 and types="copper" then
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ��������������书�Ѳ�м����ͭ���ˡ�');window.close();}</script>"
	Response.End 
end if
if mywg>2000000 and types="silver" then
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ��������������书�Ѳ�м���������ˡ�');window.close();}</script>"
	Response.End 
end if
if mywg<10000 and types="silver" then
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ������˲���书�����������񣿻�ȥ�����ɣ�');window.close();}</script>"
	Response.End 
end if
if mywg<20000 and types="gold" then
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ������˲���书�������ҽ�񣿻�ȥ�����ɣ�');window.close();}</script>"
	Response.End 
end if
Set rs1=Server.CreateObject("ADODB.RecordSet")
rs1.open "select * from "&typeses&" where ����='"&aqjh_name&"'",conn
if not rs1.eof then
	if (rs1("id")<=3 and id1=3) or (rs1("id")<=2 and id=2) then
		rs1.close
		set rs1=nothing
		conn.close
		set conn=nothing
		Response.Write "<script language=JavaScript>{alert('��ʾ�����Ѿ�ȡ�����������⻹Ҫ�ߵĹ���������Ҫ�����ˣ�');window.close();}</script>"
		Response.End
	end if
end if
ltsay=""
boxh=""
rs1.close
Set rs1=Server.CreateObject("ADODB.RecordSet")
rs1.open "select * from "&typeses&" where id="&id,conn
if myfy>rs1("����") and mygj+mynl>rs1("����")+rs1("����") then
	wugong=mywg-rs1("�书")
	defence=myfy-rs1("����")
	force=mygj-rs1("����")
	neili=mynl-rs1("����")
	conn.execute "update "&typeses&" set ����='"&aqjh_name&"',�Ա�='"&mysex&"',����='"&mymp&"',���='"&mysf&"',�书="&mywg&",����="&mynl&",����="&mytl&",����="&mygj&",����="&myfy&" where id="&id
	sql="update �û� set ����=����+500,����=����-50,allvalue=allvalue+100, ����=����+500000 where ����='"&aqjh_name&"'"
	set rs=conn.execute(sql)
	boxh="��ϲ�㣬��Ұ�ɹ���"
	ltsay="<font color=#9900ff>��ϲ<font color=#ff0000>"& aqjh_name &"</font>�Ұ�ɹ���������<font color=#ff0000>"& typeses & classes &"</font>�ı���!վ������������5000������100�㣬����500000����</font>" 
else
	conn.execute "update �û� set  ����=����-50,����=����-50,����=����-500 where ����='"&aqjh_name&"'"
	boxh="�Բ�����Ұ�ʧ�ܣ��������ñ������׵ġ�" 
	ltsay="<font color=#9900ff>��ù��</font><font color=#ff0000>"& aqjh_name &"</font>��"&typeses&"ʧ�ܣ��������ı������ס�" 
end if
rs1.close
set rs1=nothing
conn.close
set conn=nothing
newer=ltsay
says="<font color=red><b>���ٸ����桿</b></font>" & newer
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
Response.Write "<script language=JavaScript>{alert('��ʾ��"&boxh&"');window.close();}</script>"
%>
