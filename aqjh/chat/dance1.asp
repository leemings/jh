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
to1=LCase(trim(Request.QueryString("to1")))
if InStr(id,"or")<>0 or InStr(id,"'")<>0 or InStr(id,"`")<>0 or InStr(id,"=")<>0 or InStr(id,"-")<>0 or InStr(id,",")<>0 then 
	Response.Write "<script language=JavaScript>{alert('���ɣ�������ʲô���뵷���𣿣�');}</script>"
	Response.End 
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select ���� FROM �û� WHERE ����='"&name&"'",conn
momo=rs("����")
if yn=1 then
	if aqjh_name=to1 then
		if Application("aqjh_online")<>aqjh_name then
			Response.Write "<script language=JavaScript>{alert('��ʾ�����糬ʱ����');}</script>"
			Response.End 
		end if
		Response.Write "<script language=JavaScript>{alert('�Ҵ�Ӧ���Ҫ����.�������ģ�');}</script>"
		conn.execute "update �û� set ����=����+500,����=����-100 where ����='" & aqjh_name &"'"
		conn.execute "update �û� set ����=����-100,����=����+500 where ����='"&name&"'"
		dance="��ϲ��["& name &"]����{"& aqjh_name &"}���ĳɹ�����������������ģ���<br><a href=liao.asp target=_blank><img src=img/dt5.gif width=149 height=85></a><br>"& name &"��������500���¼�100"& aqjh_name &"���¼�500������100"
	else
		randomize()
		r = Int(7 * Rnd)+1
	Select Case r
		Case 1
			dance=aqjh_name&"˵��֧�֣�"&aqjh_name&"��֧��["&momo&"��"&to1&"]���ģ��������ĺܷ�Ȥ�ǡ�����"
		case 2 
			dance=aqjh_name&"˵��֧�֣�"&aqjh_name&"���޳�["&momo&"��"&to1&"]����,��û������������."
		case 3
			dance=aqjh_name&"˵��֧�֣�"&aqjh_name&"���޳�["&momo&"��"&to1&"]����,��û���������Ƥ���."
		case 4
			dance=aqjh_name&"˵��֧�֣�"&aqjh_name&"���޳�["&momo&"��"&to1&"]����,��û�����һ����һ���������Ź�����."
		case 5
			dance=aqjh_name&"˵��֧�֣�"&aqjh_name&"���޳�["&momo&"��"&to1&"]����,��û������������."
		case else
			dance=aqjh_name&"˵��֧�֣�"&aqjh_name&"���޳�["&momo&"��"&to1&"]����,��û������������."
	end select
	if momo=aqjh_name then
			dance=aqjh_name&"˵��֧�֣�"&aqjh_name&"��֧��["&momo&"��"&to1&"]���ģ����ǲ�Ҫ�����������Լ�֧���Լ����ĵģ�"
	end if
	end if
else
	if aqjh_name=to1 then
			if Application("aqjh_online")<>aqjh_name then
				Response.Write "<script language=JavaScript>{alert('��ʾ�����糬ʱ����');}</script>"
				Response.End 
			end if
		Response.Write "<script language=JavaScript>{alert('�ҲŲ����������������Ұ���');}</script>"
		conn.execute "update �û� set ����=����-100 where ����='"&name&"'"
		randomize()
		r = Int(6 * Rnd)+1
		Select Case r
			Case 1
				dance="���ܾ����ġ�"&aqjh_name&"��"&momo&"˵�������ȥ�ұ��������ĺ���Ҳ!["& momo &"]����{"& aqjh_name &"}���Ĳ��ɣ������½�100��!"
			case 2
				dance="���ܾ����ġ�"&aqjh_name&"��"&momo&"˵�������ȥ�ұ��������Ĳ����ҵ�ר��Ҳ.�ϴ������ҵ�ǿ��!["& momo &"]����{"& aqjh_name &"}���Ĳ��ɣ������½�100��!"
			case 3
				dance="���ܾ����ġ�"&aqjh_name&"��"&momo&"˵�������ȥ�ұ����������Ҳ�����.����˯���ҿɺ�Ը��!["& momo &"]����{"& aqjh_name &"}���Ĳ��ɣ������½�100��!"
			case 4
				dance="���ܾ����ġ�"&aqjh_name&"��"&momo&"˵�������ȥ�ұ��������ĺ���Ҳ!["& momo &"]����{"& aqjh_name &"}���Ĳ��ɣ������½�100��!"
			case else
				dance="���ܾ����ġ�"&aqjh_name&"��"&momo&"˵�������ȥ�ұ��������ĺ���Ҳ!["& momo &"]����{"& aqjh_name &"}���Ĳ��ɣ������½�100��!"
			end select
				
	else
		randomize()
		r = Int(6 * Rnd)+1
	Select Case r
		Case 1
			dance=aqjh_name&"˵�����ԣ�"&aqjh_name&"�ҷ���["&momo&"��"&to1&"]���ģ��������ĸ����Ͳ��ÿ������ˣ�"
		case 2
			dance=aqjh_name&"˵�����ԣ�"&aqjh_name&"�ҷ���["&momo&"��"&to1&"]���ģ����������.���������������˺�~~~"
		case 3
			dance=aqjh_name&"˵�����ԣ�"&aqjh_name&"�ҷ���["&momo&"��"&to1&"]���ģ����������.������Ƥ���������˺�~~~"
		case 4
			dance=aqjh_name&"˵�����ԣ�"&aqjh_name&"�ҷ���["&momo&"��"&to1&"]���ģ����������.���������������˺�~~~"
		case 5
			dance=aqjh_name&"˵�����ԣ�"&aqjh_name&"�ҷ���["&momo&"��"&to1&"]���ģ����������.���������������˺�~~~"
		case else
			dance=aqjh_name&"˵�����ԣ�"&aqjh_name&"�ҷ���["&momo&"��"&to1&"]���ģ����������.���������������˺�~~~"
	end select
	end if
end if
rs.close
set rs=nothing
conn.close
set conn=nothing
says="<font color=red><b>������������Ϣ��</b></font>"&dance	
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