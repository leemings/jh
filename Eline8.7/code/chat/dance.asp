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
to1=LCase(trim(Request.QueryString("to1")))
if InStr(id,"or")<>0 or InStr(id,"'")<>0 or InStr(id,"`")<>0 or InStr(id,"=")<>0 or InStr(id,"-")<>0 or InStr(id,",")<>0 then 
	Response.Write "<script language=JavaScript>{alert('���ɣ�������ʲô���뵷���𣿣�');}</script>"
	Response.End 
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "select ���� FROM �û� WHERE ����='"&name&"'",conn
momo=rs("����")
if yn=1 then
	if sjjh_name=to1 then
		if Application("sjjh_online")<>sjjh_name then
			Response.Write "<script language=JavaScript>{alert('��ʾ�����糬ʱ����');}</script>"
			Response.End 
		end if
		Response.Write "<script language=JavaScript>{alert('�Ҵ�Ӧ���Ҫ����.�������裡');}</script>"
		conn.execute "update �û� set ����=����+100,����=����-100 where ����='" & sjjh_name &"'"
		conn.execute "update �û� set ����=����-100,����=����+100 where ����='"&name&"'"
		dance="��ϲ��["& name &"]����{"& sjjh_name &"}����ɹ�����������������裡��<br><img src='img/181.gif'><br>"& name &"��������100���¼�100"& sjjh_name &"���¼�100������100"
	else
		randomize()
		r = Int(7 * Rnd)+1
	Select Case r
		Case 1
			dance=sjjh_name&"˵��֧�֣�"&sjjh_name&"��֧��["&momo&"��"&to1&"]���裬��������ܷ�Ȥ�ǡ�����"
		case 2 
			dance=sjjh_name&"˵��֧�֣�"&sjjh_name&"���޳�["&momo&"��"&to1&"]����,��û��������������."
		case 3
			dance=sjjh_name&"˵��֧�֣�"&sjjh_name&"���޳�["&momo&"��"&to1&"]����,��û���������Ƥ���."
		case 4
			dance=sjjh_name&"˵��֧�֣�"&sjjh_name&"���޳�["&momo&"��"&to1&"]����,��û�����һ����һ���������Ź�����."
		case 5
			dance=sjjh_name&"˵��֧�֣�"&sjjh_name&"���޳�["&momo&"��"&to1&"]����,��û��������������."
		case else
			dance=sjjh_name&"˵��֧�֣�"&sjjh_name&"���޳�["&momo&"��"&to1&"]����,��û��������������."
	end select
	if momo=sjjh_name then
			dance=sjjh_name&"˵��֧�֣�"&sjjh_name&"��֧��["&momo&"��"&to1&"]���裬���ǲ�Ҫ�����������Լ�֧���Լ�����ģ�"
	end if
	end if
else
	if sjjh_name=to1 then
			if Application("sjjh_online")<>sjjh_name then
				Response.Write "<script language=JavaScript>{alert('��ʾ�����糬ʱ����');}</script>"
				Response.End 
			end if
		Response.Write "<script language=JavaScript>{alert('�ҲŲ����������������Ұ���');}</script>"
		conn.execute "update �û� set ����=����-100 where ����='"&name&"'"
		randomize()
		r = Int(6 * Rnd)+1
		Select Case r
			Case 1
				dance="���ܾ����衿"&sjjh_name&"��"&momo&"˵�������ȥ�ұ������������Ҳ!["& momo &"]����{"& sjjh_name &"}���費�ɣ������½�100��!"
			case 2
				dance="���ܾ����衿"&sjjh_name&"��"&momo&"˵�������ȥ�ұ��������費���ҵ�ר��Ҳ.�ϴ������ҵ�ǿ��!["& momo &"]����{"& sjjh_name &"}���費�ɣ������½�100��!"
			case 3
				dance="���ܾ����衿"&sjjh_name&"��"&momo&"˵�������ȥ�ұ����������Ҳ�����.����˯���ҿɺ�Ը��!["& momo &"]����{"& sjjh_name &"}���費�ɣ������½�100��!"
			case 4
				dance="���ܾ����衿"&sjjh_name&"��"&momo&"˵�������ȥ�ұ������������Ҳ!["& momo &"]����{"& sjjh_name &"}���費�ɣ������½�100��!"
			case else
				dance="���ܾ����衿"&sjjh_name&"��"&momo&"˵�������ȥ�ұ������������Ҳ!["& momo &"]����{"& sjjh_name &"}���費�ɣ������½�100��!"
			end select
				
	else
		randomize()
		r = Int(6 * Rnd)+1
	Select Case r
		Case 1
			dance=sjjh_name&"˵�����ԣ�"&sjjh_name&"�ҷ���["&momo&"��"&to1&"]���裬������������Ͳ��ÿ������ˣ�"
		case 2
			dance=sjjh_name&"˵�����ԣ�"&sjjh_name&"�ҷ���["&momo&"��"&to1&"]���裬���������.�����������������˺�~~~"
		case 3
			dance=sjjh_name&"˵�����ԣ�"&sjjh_name&"�ҷ���["&momo&"��"&to1&"]���裬���������.������Ƥ���������˺�~~~"
		case 4
			dance=sjjh_name&"˵�����ԣ�"&sjjh_name&"�ҷ���["&momo&"��"&to1&"]���裬���������.�����������������˺�~~~"
		case 5
			dance=sjjh_name&"˵�����ԣ�"&sjjh_name&"�ҷ���["&momo&"��"&to1&"]���裬���������.�����������������˺�~~~"
		case else
			dance=sjjh_name&"˵�����ԣ�"&sjjh_name&"�ҷ���["&momo&"��"&to1&"]���裬���������.�����������������˺�~~~"
	end select
	end if
end if
rs.close
set rs=nothing
conn.close
set conn=nothing
says="<font color=red><b>��������Ϣ��</b></font>"&dance	
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
