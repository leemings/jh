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
peiou=rs("����")
if yn=1 then
	if sjjh_name=to1 then
		if Application("sjjh_online")<>sjjh_name then
			Response.Write "<script language=JavaScript>{alert('��ʾ�����糬ʱ����');}</script>"
			Response.End 
		end if
		Response.Write "<script language=JavaScript>{alert('�Ҵ�Ӧ���Ҫ���ˣ�');}</script>"
		conn.execute "update �û� set ��ż='"& peiou &"',��������=date(),������=������+1 where ����='" & sjjh_name &"'"
		conn.execute "update �û� set ��ż='"& sjjh_name &"',��������=date(),������=������+1 where ����='"&name&"'"
		hunyin="��ϲ��["& sjjh_name &"]��{"& peiou &"}ϲ����Ե����ұ�ʾף�أ���<img src='img/004.gif'><br><marquee width=100% behavior=alternate scrollamount=5><font color=red size=+1>ϲϲ</font></marquee>"
	else
		randomize()
		r = Int(7 * Rnd)+1
	Select Case r
		Case 1
			hunyin=sjjh_name&"˵��֧�֣�"&sjjh_name&"��֧��["&peiou&"��"&to1&"]��飬һ���еġ�����"
		case 2 
			hunyin=sjjh_name&"˵��֧�֣�"&sjjh_name&"��֧��["&peiou&"��"&to1&"]���,���Ƿ�������ɵ,��������ٰ���,û����Ͳ�����,������Ҳ�Ͳ���ɵ,��������������,��Ҫ���˾�����."
		case 3
			hunyin=sjjh_name&"˵��֧�֣�"&sjjh_name&"��֧��["&peiou&"��"&to1&"]���,����壬����������Ҷ�ۻ�ɢ����ѻ�ܸ�������˼���֪���գ���ʱ��ҹ��Ϊ�顣���������Ѿ������ˣ���ͼ޸�����!"
		case 4
			hunyin=sjjh_name&"˵��֧�֣�"&sjjh_name&"��֧��["&peiou&"��"&to1&"]���,��̫������ˣ����Ҹ���ô˵��һ�����¸ҵ��ˡ���ƽʱ̫����һ�������Ծ����죬��û�취��"
		case 5
			hunyin=sjjh_name&"˵��֧�֣�"&sjjh_name&"��֧��["&peiou&"��"&to1&"]���,�ȵ�һ�ж���͸��ϣ���������㿴ϸˮ���������������Ӧ��,�ǿ�����������ҿ���"
		case else
			hunyin=sjjh_name&"˵��֧�֣�"&sjjh_name&"��֧��["&peiou&"��"&to1&"]��飬��ĵ��ϣ�ϣ�������鲻�䡭"
	end select
	if peiou=sjjh_name then
			hunyin=sjjh_name&"˵��֧�֣�"&sjjh_name&"��֧��["&peiou&"��"&to1&"]��飬���ǲ�Ҫ�����������Լ�˵�Լ��õģ�"
	end if
	end if
else
	if sjjh_name=to1 then
			if Application("sjjh_online")<>sjjh_name then
				Response.Write "<script language=JavaScript>{alert('��ʾ�����糬ʱ����');}</script>"
				Response.End 
			end if
		Response.Write "<script language=JavaScript>{alert('�ҲŲ����أ�');}</script>"
		conn.execute "update �û� set ����=����-100 where ����='"&name&"'"
		randomize()
		r = Int(6 * Rnd)+1
		Select Case r
			Case 1
				hunyin="���ܻ顿"&sjjh_name&"��"&peiou&"˵��������ҵ����Ҹ��õ��ˡ����Ҳ�������ƭ����Լ���ϣ����ԭ���Ұɣ��һ���Զף�����!["& peiou &"]��{"& sjjh_name &"}��鲻�ɣ������½�100��!"
			case 2
				hunyin="���ܻ顿"&sjjh_name&"��"&peiou&"˵���Ҽ���Ĺ���̫ϲ���㣬������è��̫ϲ���ң����Բ�Ҫ��˵ʲô���Ժ������������׵�!"
			case 3
				hunyin="���ܻ顿"&sjjh_name&"��"&peiou&"˵����֪��ʱ������������ɫ��ʧȥ���޷����.....�����Ѻ���"
			case 4
				hunyin="���ܻ顿"&sjjh_name&"��"&peiou&"˵�������Ĳ׺��Ѿ��˳����ҵ������Ѿ������м��飬���㻹���ұ��Ů���ɡ�"
			case else
				hunyin="���ܻ顿"&sjjh_name&"��"&peiou&"˵���㿴���ҵ�������Ҳ�䣬�Ҿ��Ǽ޸������ܣ�Ҳ����޸������������"
			end select
				
	else
		randomize()
		r = Int(6 * Rnd)+1
	Select Case r
		Case 1
			hunyin=sjjh_name&"˵�����ԣ�"&sjjh_name&"�ҷ���["&peiou&"��"&to1&"]��飬���Ǹ��������ʡ������������ң�"
		case 2
			hunyin=sjjh_name&"˵�����ԣ�"&sjjh_name&"�ҷ���["&peiou&"��"&to1&"]��飬��޸������㿴�ұ���Ӣ�����ˣ����г��ӣ����ӣ�Ʊ�ӣ����ھ�ȱ�����ˡ�Ȼ��ZZZƲƲ���룺�Ǻ�~~~��ƭƭ����˵~~~"
		case 3
			hunyin=sjjh_name&"˵�����ԣ�"&sjjh_name&"�ҷ���["&peiou&"��"&to1&"]��飬����ɫ�ǣ�ǰ�����һ�������������Ժ�����������У���"
		case 4
			hunyin=sjjh_name&"˵�����ԣ�"&sjjh_name&"�ҷ���["&peiou&"��"&to1&"]��飬����ɫ�ǣ�ǰ�����һ�������������Ժ�����������У���"
		case 5
			hunyin=sjjh_name&"˵�����ԣ�"&sjjh_name&"�ҷ���["&peiou&"��"&to1&"]��飬��������� ����ֻ������ƨ !"
		case else
			hunyin=sjjh_name&"˵�����ԣ�"&sjjh_name&"�ҷ���["&peiou&"��"&to1&"]��飬�޸���ǰ��Ϊ�������£��޸��������ȫ��Ϊ���ɽ�˺��ĵ����飬Ҳ���������������ƭ����������ǧ��Ҫ�޸���!"
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
