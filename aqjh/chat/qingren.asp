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
yn=LCase(trim(request.querystring("yn")))
ns=application("aqjh_sjqr")
if ns="" then
	Response.Write "<script language=JavaScript>{alert('û�����������������');}</script>"
	Response.End 
end if
qrdata=split(ns,"|")
name=qrdata(0)
to1=qrdata(1)
tcsj=qrdata(2)
erase qrdata
if yn<>0 and yn<>1 then
	Response.Write "<script language=JavaScript>{alert('��!ˢ���ɽ��ˡ�');}</script>"
	Response.End 
end if
nowsj=DateDiff("s",tcsj,now())
if nowsj>=30 then
	Response.Write "<script language=JavaScript>{alert('�������糬ʱ�����������󳬹�30�룬��ʧЧ��');}</script>"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select ����,��Ա�� FROM �û� WHERE ����='"&name&"'",conn,2,2
if rs("��Ա��")<2 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ���Է���Ա�𿨲���2Ԫ!');}</script>"
	Response.End 
end if
if yn=1 then
	if aqjh_name=to1 then
		Application.Lock
		Application("aqjh_sjqr")=""
		Application.unLock
		Response.Write "<script language=JavaScript>{alert('�Ҵ�Ӧ���Ҫ���ˣ�');}</script>"
		conn.execute "update �û� set ��Ա��=��Ա��-2,����='"& aqjh_name &"',������=������+1 where ����='" & name &"'"
		conn.execute "update �û� set ��Ա��=��Ա��+1,����='"& name &"',������=������+1 where ����='" & aqjh_name & "'"
		hunyin="��ϲ��["& aqjh_name &"]��{"& name &"}ϲ����Ե����ұ�ʾף�أ���<img src='img/004.gif'><br><marquee width=100% behavior=alternate scrollamount=5><font color=red size=+1>ϲϲ</font></marquee>"
		Application("aqjh_sjqr")=""
	else
		randomize()
		r = Int(7 * Rnd)+1
	Select Case r
		Case 1
			hunyin=aqjh_name&"˵��֧�֣�"&aqjh_name&"��֧��["&name&"��"&to1&"]�����ˣ�һ���еġ�����"
		case 2 
			hunyin=aqjh_name&"˵��֧�֣�"&aqjh_name&"��֧��["&name&"��"&to1&"]������,���Ƿ�������ɵ,��������ٰ���,û����Ͳ�����,������Ҳ�Ͳ���ɵ,��������������,��Ҫ���˾�����."
		case 3
			hunyin=aqjh_name&"˵��֧�֣�"&aqjh_name&"��֧��["&name&"��"&to1&"]������,����壬����������Ҷ�ۻ�ɢ����ѻ�ܸ�������˼���֪���գ���ʱ��ҹ��Ϊ�顣���������Ѿ������ˣ���ͼ޸�����!"
		case 4
			hunyin=aqjh_name&"˵��֧�֣�"&aqjh_name&"��֧��["&name&"��"&to1&"]������,��̫������ˣ����Ҹ���ô˵��һ�����¸ҵ��ˡ���ƽʱ̫����һ�������Ծ����죬��û�취��"
		case 5
			hunyin=aqjh_name&"˵��֧�֣�"&aqjh_name&"��֧��["&name&"��"&to1&"]������,�ȵ�һ�ж���͸��ϣ���������㿴ϸˮ���������������Ӧ��,�ǿ�����������ҿ���"
		case else
			hunyin=aqjh_name&"˵��֧�֣�"&aqjh_name&"��֧��["&name&"��"&to1&"]�����ˣ���ĵ��ϣ�ϣ�������鲻�䡭"
	end select
	if name=aqjh_name then
			hunyin=aqjh_name&"˵��֧�֣�"&aqjh_name&"��֧��["&name&"��"&to1&"]�����ˣ����ǲ�Ҫ�����������Լ�˵�Լ��õģ�"
	end if
	end if
else
	if aqjh_name=to1 then
		Application.Lock
		Application("aqjh_sjqr")=""
		Application.unLock
		Response.Write "<script language=JavaScript>{alert('�ҲŲ����أ�');}</script>"
		conn.execute "update �û� set ����=����-100 where ����='&name&'"
		randomize()
		r = Int(6 * Rnd)+1
		Select Case r
			Case 1
				hunyin="���ܻ顿"&aqjh_name&"��"&name&"˵��������ҵ����Ҹ��õ��ˡ����Ҳ�������ƭ����Լ���ϣ����ԭ���Ұɣ��һ���Զף�����!["& name &"]��{"& aqjh_name &"}��鲻�ɣ������½�100��!"
			case 2
				hunyin="���ܻ顿"&aqjh_name&"��"&name&"˵���Ҽ���Ĺ���̫ϲ���㣬������è��̫ϲ���ң����Բ�Ҫ��˵ʲô���Ժ������������׵�!"
			case 3
				hunyin="���ܻ顿"&aqjh_name&"��"&name&"˵����֪��ʱ������������ɫ��ʧȥ���޷����.....�����Ѻ���"
			case 4
				hunyin="���ܻ顿"&aqjh_name&"��"&name&"˵�������Ĳ׺��Ѿ��˳����ҵ������Ѿ������м��飬���㻹���ұ��Ů���ɡ�"
			case else
				hunyin="���ܻ顿"&aqjh_name&"��"&name&"˵���㿴���ҵ�������Ҳ�䣬�Ҿ��Ǽ޸������ܣ�Ҳ����޸������������"
			end select
				
	else
		randomize()
		r = Int(6 * Rnd)+1
	Select Case r
		Case 1
			hunyin=aqjh_name&"˵�����ԣ�"&aqjh_name&"�ҷ���["&name&"��"&to1&"]�����ˣ����Ǹ��������ʡ������������ң�"
		case 2
			hunyin=aqjh_name&"˵�����ԣ�"&aqjh_name&"�ҷ���["&name&"��"&to1&"]�����ˣ���޸������㿴�ұ���Ӣ�����ˣ����г��ӣ����ӣ�Ʊ�ӣ����ھ�ȱ�����ˡ�Ȼ��ZZZƲƲ���룺�Ǻ�~~~��ƭƭ����˵~~~"
		case 3
			hunyin=aqjh_name&"˵�����ԣ�"&aqjh_name&"�ҷ���["&name&"��"&to1&"]�����ˣ�����ɫ�ǣ�ǰ�����һ�������������Ժ�����������У���"
		case 4
			hunyin=aqjh_name&"˵�����ԣ�"&aqjh_name&"�ҷ���["&name&"��"&to1&"]�����ˣ�����ɫ�ǣ�ǰ�����һ�������������Ժ�����������У���"
		case 5
			hunyin=aqjh_name&"˵�����ԣ�"&aqjh_name&"�ҷ���["&name&"��"&to1&"]�����ˣ���������� ����ֻ������ƨ !"
		case else
			hunyin=aqjh_name&"˵�����ԣ�"&aqjh_name&"�ҷ���["&name&"��"&to1&"]�����ˣ��޸���ǰ��Ϊ�������£��޸��������ȫ��Ϊ���ɽ�˺��ĵ����飬Ҳ���������������ƭ����������ǧ��Ҫ�޸���!"
	end select
	end if
end if
rs.close
set rs=nothing
conn.close
set conn=nothing
says="<font color=red><b>��������Ϣ��</b></font>"&hunyin	
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
