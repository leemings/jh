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

fromname=LCase(trim(request.querystring("fromname")))	'�Է�����
toname=LCase(trim(request.querystring("toname")))	'��������
qql=lcase(trim(request.querystring("qq"))) '��������
if instr(fromname,"or")<>0 or instr(fromname,".")<>0 or instr(fromname,"'")<>0 or instr(fromname,"OR")<>0 or instr(fromname,"%20")<>0 or instr(fromname,">")<>0 or instr(fromname,"=")<>0 or instr(fromname,"<")<>0 or instr(fromname," ")<>0 then
	Response.Write "<script language=JavaScript>{alert('������ʲô��');}</script>"
	Response.End
end if
if instr(toname,"or")<>0 or instr(toname,".")<>0 or instr(toname,"'")<>0 or instr(toname,"OR")<>0 or instr(toname,"%20")<>0 or instr(toname,">")<>0 or instr(toname,"=")<>0 or instr(toname,"<")<>0 or instr(toname," ")<>0 then
	Response.Write "<script language=JavaScript>{alert('������ʲô��');}</script>"
	Response.End
end if
if Instr(LCase(Application("sjjh_useronlinename"&session("nowinroom")))," "&LCase(fromname)&" ")=0 then
	Response.Write "<script language=JavaScript>{alert('��ʾ�������鲻���������У����������裡');location.href = 'javascript:history.go(-1)';}</script>"
	Application.Lock
	Application("sjjh_smxs")=""
	Application.unLock
	Response.End
end if
if toname<>sjjh_name then
	Response.Write "<script language=JavaScript>{alert('������ʲô���˼�"&fromname&"�������������裬�����������ˣ�');}</script>"
	Response.End
end if
if fromname=sjjh_name then
	Response.Write "<script language=JavaScript>{alert('������,������������������ģ�');}</script>"
	Response.End
end if
if qql<>"������" and qql<>"̽��" and qql<>"��˹��" then
	Response.Write "<script language=JavaScript>{alert('������ʱֻ�л����ȡ�̽��͵�˹������������');}</script>"
	Response.End
end if
Application.Lock
	Application("sjjh_tw")=""
Application.unLock
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "select ����,����,����,����,�Ա� from �û� where ����='"&fromname&"'",conn,2,2
if rs.bof or rs.eof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('������û��"&fromname&"����ˣ�');}</script>"
	Response.End
end if
if rs("����")<1000 or rs("����")<5000 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('"&fromname&" �ĵ���û�ﵽ1000����������5000����������̫û�����ˣ�');}</script>"
	Response.End
end if
if rs("����")<8000 or rs("����")<5000 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('"&fromname&" ����������8000����������5000��С�ı���������ˣ�');}</script>"
	Response.End
end if
toml=rs("����")
fromsex=rs("�Ա�")
rs.close
rs.open "select ����,����,����,����,���� from �û� where ����='"&sjjh_name&"'",conn,2,2
If Rs.Bof OR Rs.Eof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('�㲻�ǽ������ˣ�');}</script>"
	Response.End
end if
if rs("����")<50000 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��û��5�������ӣ���������Ʊ��');}</script>"
	Response.End
end if
if rs("����")<1000 or rs("����")<5000 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Application.Lock
	Application("sjjh_tw")=""
	Application.UnLock
	Response.Write "<script language=JavaScript>{alert('��ĵ���û�ﵽ1000����������5000����������̫û�����ˣ�');}</script>"
	Response.End
end if
if rs("����")<8000 or rs("����")<5000 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Application.Lock
	Application("sjjh_tw")=""
	Application.UnLock
	Response.Write "<script language=JavaScript>{alert('�����������8000����������5000��С�����ѣ�');}</script>"
	Response.End
end if
myml=rs("����")
rs.close
'��ʼ��ǩ
ab="��"&qql&"�衿<font color=red>��"& sjjh_name &"��</font>��<font color=red>��"& fromname &"��</font>������"&qql&"�衣"
wei="<font color=red>��"& sjjh_name &"��</font>"
yinliang=50000
mytl=0
mynl=0
myml=0
myallvalue=0
myjb=0
mydd=0
select case qql
	case "������"
		randomize timer
		r=int(rnd*9)+1
		select case r
			case 1
				tw="<bgsound src=mid/tg.mid loop=1><img src=img/hez.gif><font color=blue>"& fromname &"</font>����<font color=blue>"& sjjh_name &"</font>���ȶ��Ƶ������������Ĳ���Ȼ�����¸Ҵ󷽵�������������һ������Ļ����ȣ�������������ҹ������֮�С�����15��������"
				yingliang=yingliang+150000
			case 2,8
				tw="<bgsound src=mid/hez.mid loop=1><img src=img/hez.gif><font color=blue>"& fromname &"</font>������߹�����§��<font color=blue>"& sjjh_name &"</font>������������Ļ��������֣�%%Ͷ��������������񣬲���Ҳ�ã�����Ҳ�գ�������ʹ���貽����.��Ϊ�����ģ�����"& sjjh_name &"ע�ӵ�Ŀ�⣬��������Ŀ��һ�����ҵس�����һ��"
			case 3,7
				tw="<bgsound src=mid/hez.mid loop=1><img src=img/hez.gif><font color=blue>"& sjjh_name &"</font>�����§��<font color=blue>"& fromname &"</font>�����ŵ��������˵Ļ����ȣ���ӵ������Ⱥ�У�����Ͷ��������������񣬲���Ҳ�ã�����Ҳ�գ�������ʹ���貽����.��Ϊ�����ģ�����"& fromname &"ע�ӵ�Ŀ�⣬����������ģ�"
			case 4
				tw="<bgsound src=mid/hez.mid loop=1><img src=img/hez.gif><font color=blue>"& sjjh_name &"</font>ɵ�������ߵ�<font color=blue>"& fromname &"</font>��ǰ����������ѽ���㲻��ͱ�����ѽ���������̤��һ������̤��"& fromname &"�Ľţ�һ���ܷ��ˣ���"& fromname &"ʮ����������"&sjjh_name&"ʹʧ50���"
				dianjuan=-50
				yinliang=-100000
			case 5
				tw="<bgsound src=mid/hez.mid loop=1><img src=img/hez.gif><font color=blue>"& sjjh_name &"</font>ɵɵ��§��<font color=blue>"& fromname &"</font>���������������˵Ļ����ȣ���ӵ������Ⱥ�У�ʹ���貽����.��"& fromname &"ʮ����������"&sjjh_name&"ʹʧ50���"
				dianjuan=-50
				yinliang=-100000
			case 6
				tw="<bgsound src=mid/hez.mid loop=1><img src=img/hez.gif><font color=blue>"& sjjh_name &"</font>�����§��<font color=blue>"& fromname &"</font>�����ŵ��������˵Ļ����ȣ���ӵ������Ⱥ�У�ʹ���貽����.��"& fromname &"ʮ������������"&sjjh_name&"��50���"
				dianjuan=50
				yingliang=180000
			case else
				tw="<bgsound src=mid/hez.mid loop=1><img src=img/hez.gif><font color=blue>"& sjjh_name &"</font>�����§��<font color=blue>"& fromname &"</font>�����ŵ��������˵Ļ����ȣ����������費������˭���������ָ߽ţ�"& fromname &"�Ҹ��ˣ�����ô�õİ��ˣ�"
		end select
	case "̽��"
		randomize timer
		r=int(rnd*9)+1
		select case r
			case 1
				tw="<bgsound src=mid/tg.mid loop=1><img src=img/tg.gif><font color=blue>"& fromname &"</font>����<font color=blue>"& sjjh_name &"</font>��С�����������˾����̽�꣬������������ҹ������֮�С�"
			case 2,8
				tw="<bgsound src=mid/tg.mid loop=1><img src=img/tg.gif><font color=blue>"& fromname &"</font>����<font color=blue>"& sjjh_name &"</font>���������˱�����̽�꣬�����������½�1000��"
				mytl=-1000
				myml=-1000
			case 3,7
				tw="<bgsound src=mid/tg.mid loop=1><img src=img/tg.gif><font color=blue>"& sjjh_name &"</font>�������������������<font color=blue>"& fromname &"</font>����������������С�����������˾����̽�꣬�����������������⾳�м��������ꡣ"
			case 4
				tw="<bgsound src=mid/tg.mid loop=1><img src=img/tg.gif><font color=blue>"& fromname &"</font>����<font color=blue>"& sjjh_name &"</font>���ȶ��Ƶ������������Ĳ���Ȼ�����¸Ҵ󷽵�������������һ�������̽�꣬������������ҹ������֮�С����²�������500����"
				mydd=-500
			case 5
				tw="<bgsound src=mid/tg.mid loop=1><img src=img/tg.gif><font color=blue>"& fromname &"</font>����<font color=blue>"& sjjh_name &"</font>���ȶ��Ƶ������������Ĳ���Ȼ�����¸Ҵ󷽵�������������һ�������̽�꣬������������ҹ������֮�С�������200��������1000������������"
				mytl=200
				myml=1000
				yingliang=yingliang+50000
			case 6
				tw="<bgsound src=mid/tg.mid loop=1><img src=img/tg.gif><font color=blue>"& fromname &"</font>����<font color=blue>"& sjjh_name &"</font>���ȶ��Ƶ������������Ĳ���Ȼ�����¸Ҵ󷽵�������������һ�������̽�꣬�Ρ��������֮�⣬�������˭�ˣ������Ǹ��֣�������������ҹ������֮�С�"
			case else
				tw="<bgsound src=mid/tg.mid loop=1><img src=img/tg.gif><font color=blue>"& fromname &"</font>����<font color=blue>"& sjjh_name &"</font>���ȶ��Ƶ������������Ĳ���Ȼ�����¸Ҵ󷽵�������������һ�������̽�꣬������������ҹ������֮�С�"
		end select
	case "��˹��"
		randomize timer
		r=int(rnd*9)+1
		select case r
			case 1
				tw="<bgsound src=mid/disco.mid loop=1><img src=img/disco.gif><font color=blue>"& sjjh_name &"</font>����һƬ��ȹ����Ť������֫������ʱ�������׼������ǣ�Ū����������ֽ�������<font color=blue>"& fromname &"</font>Ҳ����Ť�����Ƿʴ���β���"
			case 2,8
				tw="<bgsound src=mid/disco.mid loop=1><img src=img/disco.gif><font color=blue>"& sjjh_name &"</font>����һƬ��ȹ����Ť��������ϸ֫������ʱ�������׼������ۣ�Ū����������ֽ�������<font color=blue>"& fromname &"</font>Ҳ���ذ�������ߡ�"
			case 3,7
				tw="<bgsound src=mid/disco.mid loop=1><img src=img/disco.gif><font color=blue>"& fromname &"</font>����һƬ��ȹ���Ū������DISCO����������ֽ�������<font color=blue>"& sjjh_name &"</font>Ҳ���غ���������SO IN������������"
				yingliang=yinliang+50000
			case 4
				tw="<bgsound src=mid/disco.mid loop=1><img src=img/disco.gif><font color=blue>"& sjjh_name &"</font>��<font color=blue>"& fromname &"</font>�ڵ���������裬�����ˣ����������½�500."
				mytl=-500
				mynl=-500
				yingliang=0
			case 5
				tw="<bgsound src=mid/disco.mid loop=1><img src=img/disco.gif><font color=blue>"& sjjh_name &"</font>��<font color=blue>"& fromname &"</font>�ڵ���������裬����������ƹ���ҫ��̨���ǻ����ŵ��ֵĺ���.������ʮ��"
				yingliang=yingliang+100000
			case 6
				tw="<bgsound src=mid/disco.mid loop=1><img src=img/disco.gif><font color=blue>"& sjjh_name &"</font>��<font color=blue>"& fromname &"</font>�ڵ���������裬лͥ��Ҳ������ˣ�������˭��������һ�ϸߵ�."
			case else
				tw="<bgsound src=mid/disco.mid loop=1><img src=img/disco.gif><font color=blue>"& sjjh_name &"</font>��<font color=blue>"& fromname &"</font>�ں��̵ƾư������������ƹ���ҫ�����Ǻܵ���ѽ���Ƶá�÷���˹飬���ú��Ц.������500"
				mytl=500
		end select
end select

conn.execute "update �û� set ����=����-"&yinliang&",����=����+"&mytl&",����=����+"&mynl&",����=����+"&myml&",allvalue=allvalue+"&myallvalue&",����=����+"&mydd&" where ����='"&sjjh_name&"'"
conn.execute "update �û� set ����=����+"&yinliang&",����=����-800,����=����-1000 where ����='"&fromname&"'"
set rs=nothing
conn.close
set conn=nothing
tw=ab&tw&wei
Application.Lock
says="<font color=red><b>���ȶ�������</b></font>"&tw	
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