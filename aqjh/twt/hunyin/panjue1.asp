<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
nowinroom=session("nowinroom")
if aqjh_name="" then Response.Redirect "../../error.asp?id=440"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select ��ż,�ȼ�,����,����,boy FROM �û� WHERE ����='" & aqjh_name &"'",conn
peiou=rs("��ż")
myboy=rs("boy")
	if isnull(myboy) or myboy="" then
		myboy=""
	else
		zt=split(myboy,"|")
		if UBound(zt)<>7 then
			myboy=""
		end if
	end if
if peiou="��" then
		rs.close
		set rs=nothing	
		conn.close
		set conn=nothing
		Response.Write "<script language=javascript>alert('��"&aqjh_name&"���㻹û�н����,��ʲô�鰡��');window.close();</script>"
		Response.End
end if
if rs("�ȼ�")<18 then
Response.Write "<script language=JavaScript>{alert('����С��Ҳ�ӻ�ѽ,����18���ȼ���˵�ɣ�');window.close();}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
if rs("����")<100000 then
Response.Write "<script language=JavaScript>{alert('�����������10��ѽ����ôҲ�����ӻ鰡��');window.close();}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
if rs("����")<20000 then
Response.Write "<script language=JavaScript>{alert('��ĵ��»�����2���ѧ���ӻ飿������ͨ����ѽ��');window.close();}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
if myboy<>"" then
 conn.execute "insert into gry(boy,fm1,fm2) values ('"&myboy&"','"&peiou&"','"&aqjh_name&"')"
end if
conn.Execute ("update �û� set ��ż='��',boy='',boysex='',����=����-20000,����=����-100000 where ����='" & aqjh_name &"'")
rs.close
rs.open "select ��ż FROM �û� WHERE ����='"&peiou&"'",conn
if not(Rs.Bof OR Rs.Eof) then
conn.Execute ("update �û� set ��ż='��',boy='',boysex='' where ����='"&peiou&"'")
end if
message="[�ӻ�]" & aqjh_name & "�ó���10��������" & peiou & "���ʧ�ܣ�����ܳ���Ժ����2����������ӻ�ɹ��ˣ�"
if myboy<>"" then message=message&"˫����С���Ѿ��͵��¶�Ժ���������޹��İ�������"
rs.close
set rs=nothing
conn.close
set conn=nothing
says="<font color=red>����Ϣ��</font><font color=blank>"& message &"</font>"		'��������
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
%>
<body bgcolor="#000000" background="../../bg.gif">
<table border=1 bgcolor="#BEE0FC" align=center width=350 cellpadding="10" cellspacing="13">
<tr><td bgcolor=#CCE8D6>
<table width=100%>
<tr><td>
<p align=center style='font-size:14;color:red'><%=message%></p>
<p align=center><a href=taohun.asp>����</a></p>
</td></tr>
</table>
</td></tr>
</table></body>