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
if application("aqjh_qw")="" then
	Response.Write "<script language=JavaScript>{alert('��û���������㣬������');}</script>"
	Response.End
end if
wenndata=split(application("aqjh_qw"),"|")
fromname=wenndata(0)
toname=wenndata(1)
sj=wenndata(2)
erase wenndata
qql=lcase(trim(request.querystring("zl"))) '��������
if instr(fromname,"or")<>0 or instr(fromname,".")<>0 or instr(fromname,"'")<>0 or instr(fromname,"OR")<>0 or instr(fromname,"%20")<>0 or instr(fromname,">")<>0 or instr(fromname,"=")<>0 or instr(fromname,"<")<>0 or instr(fromname," ")<>0 then
	Response.Write "<script language=JavaScript>{alert('������ʲô��');}</script>"
	Response.End
end if
if instr(toname,"or")<>0 or instr(toname,".")<>0 or instr(toname,"'")<>0 or instr(toname,"OR")<>0 or instr(toname,"%20")<>0 or instr(toname,">")<>0 or instr(toname,"=")<>0 or instr(toname,"<")<>0 or instr(toname," ")<>0 then
	Response.Write "<script language=JavaScript>{alert('������ʲô��');}</script>"
	Response.End
end if
if Instr(LCase(Application("aqjh_useronlinename"&session("nowinroom")))," "&LCase(fromname)&" ")=0 then
	Response.Write "<script language=JavaScript>{alert('��ʾ��������͵����Ѿ����������һ��Ѿ����ˣ�');location.href = 'javascript:history.go(-1)';}</script>"
	Application.Lock
	Application("aqjh_smxs")=""
	Application.unLock
	Response.End
end if
if toname<>aqjh_name then
	Response.Write "<script language=JavaScript>{alert('������ʲô���˼�"&fromname&"û�����㣡');}</script>"
	Response.End
end if
if fromname=aqjh_name then
	Response.Write "<script language=JavaScript>{alert('������,�������������ѵ��в���');}</script>"
	Response.End
end if
if qql<>"�Է�" and qql<>"�����" and qql<>"�ܾ�" then
	Response.Write "<script language=JavaScript>{alert('������ʱ�޴���ͣ�"&QQ1&"');}</script>"
	Response.End
end if
nowsj=DateDiff("s",sj,now())
if nowsj>30 then
	Application.Lock
	Application("aqjh_qw")=""
	Application.unLock
	Response.Write "<script language=JavaScript>{alert('��ʾ����ʾ������30��δ�����������');}</script>"
	Response.End
end if
Application.Lock
	Application("aqjh_qw")=""
Application.unLock
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select ����,����,����,����,�Ա� from �û� where ����='"&fromname&"'",conn,2,2
if rs.bof or rs.eof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('������û��"&fromname&"����ˣ�');}</script>"
	Response.End
end if
if rs("����")<10 or rs("����")<50 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('"&fromname&" �ĵ���û�ﵽ10����������50������һ��̫û�����ˣ�');}</script>"
	Response.End
end if
toml=rs("����")
fromsex=rs("�Ա�")
rs.close
rs.open "select ����,����,����,����,�书 from �û� where ����='"&aqjh_name&"'",conn,2,2
If Rs.Bof OR Rs.Eof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('�㲻�ǽ������ˣ�');}</script>"
	Response.End
end if
if rs("����")<10 or rs("����")<50 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Application.Lock
	Application("aqjh_qw")=""
	Application.UnLock
	Response.Write "<script language=JavaScript>{alert('��ĵ���û�ﵽ10����������00������һ��̫û�����ˣ�');}</script>"
	Response.End
end if
if rs("����")<10000 or rs("����")<6000 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Application.Lock
	Application("aqjh_qw")=""
	Application.UnLock
	Response.Write "<script language=JavaScript>{alert('�����������10000����������6000��С����������ͬ���ͺ�����ͬ��');}</script>"
	Response.End
end if
myml=rs("����")
rs.close
'��ʼ����
ab="<font color=red><b>����Ǯ��͡�</b></font>"
select case qql
	case "�Է�"
		randomize timer
		r=int(rnd*2)+1
		tp="<img src=img/chifan.gif>"
		select case r
			case 1
				if fromsex="��" then
					x="<font color=blue><font color=blue>"& fromname &"</font>����Ŀ���<font color=blue>"& aqjh_name &"</font>��ڴ�ڵ�<img src=img/cf1.gif>��"& fromname &"������"& aqjh_name &"�ĳԷ�����֮�С�"
				else
					x="<font color=blue><font color=blue>"& aqjh_name &"</font>���ĵĺ�<font color=blue>"& fromname &"</font>����<img src=img/cf4.gif>�������������λ�֮�С�"
				end if
				x=x&aqjh_name&"��������<font color=red>50��</font>����������<font color=red>50��</font>��"& fromname &"��������100�㡣"
			case 2
				randomize timer
				m=int(rnd*2)+1
				if m=1 or m=3 then
					if fromsex="��" then
						x="<font color=blue>"& fromname &"</font>�ͳ����Ǯ����<font color=blue>����~~��������Զٺõġ�����<font color=blue>"& aqjh_name &"</font>���˵������Ǽ��Ƶ����һ��<img src=img/cf5.gif>���ò������ĳ�������"
					else
						x="<font color=blue>"& aqjh_name &"</font>��<font color=blue>"& fromname &"</font>˵����<font color=blue>�����ҳԷ�������ʵ��̫���ˡ�����<font color=blue>"& fromname &"</font>���ͷ�����·������һ��<img src=img/cf6.gif>��͡�"
					end if
				else
					if fromsex="Ů" then
						x="<font color=blue>"& fromname &"</font>��<font color=blue>"& aqjh_name &"</font>˵��<font color=blue>�ٺ�~~������~~������ô���⣬����ҿ��������Ǽ���Ǯ��ʲô������<img src=img/cf3.gif>���Ժ��У�"& fromname &"��ڴ�ڵĳ���"& aqjh_name &"��ĿͲ�˵����������ģ��Ժ���������ֵ���~����"
					else
						x="<font color=blue>"& fromname &"</font>��<font color=blue>"& aqjh_name &"</font>˵��<font color=blue>�ø��ǣ����˰ɣ����㵽���¥�Դ�ˣ�����ҿ������������������<img src=img/cf2.gif>����բз�治�Ǹǵģ�ζ���ü��ˣ������Ժ����ģ�"& aqjh_name &"��ڴ�ڵĳ���"& fromname &"��Ŀͣ�������ʱ���¾䣺�ø��~����"
					end if
				end if
				x=x&aqjh_name&"��������<font color=red>50��</font>����������<font color=red>50��</font>��"& fromname &"��������100�㡣"
		end select
		ab=ab&x
		conn.execute "update �û� set ����=����-20000,����=����+50,����=����+50 where ����='"&aqjh_name&"'"
		conn.execute "update �û� set ����=����+50,����=����+100 where ����='"&fromname&"'"
case "�����"
		randomize timer
		r=int(rnd*2)+1
		tp="<img src=pic/dp78.GIF>"
		select case r
			case 1
				if fromsex="��" then
					x="<font color=0088FF><font color=blue>"& fromname &"</font>�����ֵ�<font color=blue>"& aqjh_name &"</font>�ļ�򣬿������ȵ�������<font color=blue>"& aqjh_name &"</font>����˵��������Ա���ܰɣ���֪��һ�ҺܺóԵ�Ŷ~˵�ն����߽����ı���ܵ꣬����2��<img src=img/bql1.gif>��<font color=blue>"& aqjh_name &"</font>����<font color=blue>"& fromname &"</font>�����ڱ���ܵ������С�"
				else
					x="<font color=0088FF><font color=blue>"& aqjh_name &"</font>��ݺݵĺ�<font color=blue>"& fromname &"</font>���ӣ���~������~��~��ô���ȵ�����~����Ա����ȥ���㿴<img src=img/bql2.gif>����Ư�����úóԣ����ĳ԰ɣ��Ҹ�Ǯ~~~��~"
				end if
				x=x&aqjh_name&"��������<font color=red>80��</font>����������<font color=red>80��</font>��"& fromname &"��������150�㡣"
			case 2
				randomize timer
				m=int(rnd*2)+1
				if m=1 or m=3 then
					if fromsex="��" then
						x="<font color=0088FF><font color=blue>"& fromname &"</font>վ������С���ֹ�˵�ţ�������Ծ͸���˵�������<font color=blue>"& fromname &"</font>�ӱ��ı���ܵ��ó�һ����<img src=img/bql3.gif>����<font color=blue>"& aqjh_name &"</font>�Զ��ġ�"
					else
						x="<font color=0088FF><font color=blue>"& aqjh_name &"</font>��ȵ�������С���ǲ��ǿ������Ұ�~������Ը���������������ĵģ������㲻�������Ҹ���û��~������������~������<font color=blue>"& fromname &"</font>����<font color=blue>"& aqjh_name &"</font>����������<font color=blue>"& fromname &"</font>��Ե�<img src=img/bql4.gif>��"
					end if
				else
					if fromsex="Ů" then
						x="<font color=0088FF><font color=blue>"& fromname &"</font>�������ÿ���<font color=blue>"& aqjh_name &"</font>˵������ֻʣ����һ��������ˣ�����Ҫ�͸��㡭������<font color=blue>"& aqjh_name &"</font>������˵����<font color=blue>"& fromname &"</font>��������<img src=img/bql5.gif>��ֻ��<font color=blue>"& aqjh_name &"</font>����ֱ��š�"
					else
						x="<font color=0088FF><font color=blue>"& aqjh_name &"</font>վ�ڱ��Ĵ���·�ԣ������ֹ��ţ�����Ա���ܣ���Ҫ�ԡ�����,����<font color=blue>"& fromname &"</font>�����<img src=img/bql6.gif>����Ǯ����һ���ܹ���·������<font color=blue>"& fromname &"</font>�Ƽ�����������Ҫ�����������һ����<font color=blue>"& aqjh_name &"</font>�ŵ�����ϼң��ڰ�·���޶��ա�������"
					end if
				end if
				x=x&aqjh_name&"��������<font color=red>80��</font>����������<font color=red>80��</font>��"& fromname &"��������150�㡣"
			case 3
				if fromsex="��" then
					x="<font color=0088FF><font color=blue>"& fromname &"</font>վ�ڱ���ܵ��ԣ����Ÿ����<img src=img/bql3.gif>������������һ����������̰�ԣ�����<font color=blue>"& aqjh_name &"</font>��С��������֡�����<font color=blue>"& aqjh_name &"</font>������ˮ����<font color=blue>"& fromname &"</font>�Ͱ�����ı���ܻ�����ȥ��ֻ��<font color=blue>"& aqjh_name &"</font>�����ι�ȥ�ˡ�"
				else
					x="<font color=0088FF><font color=blue>"& aqjh_name &"</font>����<font color=blue>"& fromname &"</font>���<img src=img/bql2.gif>��Խ��Խ�о�����<font color=blue>"& fromname &"</font>����һ����<font color=blue>"& fromname &"</font>������Ǯ��һ�ã�ֻʣ1ëǮ��<font color=blue>"& aqjh_name &"</font>����˵�������~~~~~~~~��"
				end if
				x=x&aqjh_name&"��������<font color=red>80��</font>����������<font color=red>80��</font>��"& fromname &"��������150�㡣"
		end select
		ab=ab&x
		conn.execute "update �û� set ����=����+80,����=����+80,����=����-650,����=����-400 where ����='"&aqjh_name&"'"
		conn.execute "update �û� set ����=����+150,����=����-6500,����=����-400 where ����='"&fromname&"'"
case "����"
		randomize timer
		r=int(rnd*2)+1
		tp="<img src=pic/dp76.gif>"
		select case r
			case 1
				if fromsex="��" then
					x="<font color=blue>"& fromname &"</font>����Ŀ���<font color=blue>"& aqjh_name &"</font>���۾�"& tp &"����������������֮�С�"
				else
					x="<font color=blue>"& aqjh_name &"</font>����Ŀ���<font color=blue>"& fromname &"</font>С���"& tp &"�������������λ�֮�С�"
				end if
				x=x&aqjh_name&"��������<font color=red>120��</font>����������<font color=red>120��</font>��"& fromname &"��������200�㡣"
			case 2
				randomize timer
				m=int(rnd*2)+1
				if m=1 or m=3 then
					if fromsex="��" then
						x="<font color=blue>"& fromname &"</font>վ������С��˵�ţ�����ϲ����ѽ�����������Ǳ�����������"& tp &"<font color=blue>"& aqjh_name &"</font>����������֮�С�"
					else
						x="<font color=blue>"& aqjh_name &"</font>վ������С��˵�ţ�������á����������Ǳ�����������"& tp &"<font color=blue>"& fromname &"</font>�������λ�֮�С�"
					end if
				else
					if fromsex="Ů" then
						x="<font color=blue>"& fromname &"</font>վ������С��˵�ţ�����ϲ����ѽ�����������Ǳ�����������"& tp &"<font color=blue>"& aqjh_name &"</font>����������֮�С�"
					else
						x="<font color=blue>"& aqjh_name &"</font>վ������С��˵�ţ�������𡭡��������Ǳ�����������"& tp &"<font color=blue>"& fromname &"</font>�������λ�֮�С�"
					end if
				end if
				x=x&aqjh_name&"��������<font color=red>120��</font>����������<font color=red>120��</font>��"& fromname &"��������200�㡣"
			case 3
				if fromsex="��" then
					x="<font color=blue>"& fromname &"</font>����Ŀ���<font color=blue>"& aqjh_name &"</font>��С���"& tp &"����������������֮�С�"
				else
					x="<font color=blue>"& aqjh_name &"</font>����Ŀ���<font color=blue>"& fromname &"</font>�Ĵ��۾�"& tp &"�������������λ�֮�С�"
				end if
				x=x&aqjh_name&"��������<font color=red>120��</font>����������<font color=red>120��</font>��"& fromname &"��������200�㡣"
		end select
		ab=ab&x
		conn.execute "update �û� set ����=����+120,����=����+120,����=����-1000,����=����-6000 where ����='"&aqjh_name&"'"
		conn.execute "update �û� set ����=����+200,����=����-1000,����=����-600 where ����='"&fromname&"'"
	
	case "�ܾ�"
		randomize timer
		r=int(rnd*5)/5
		toml=int(toml*r)
		ab="<font color=red><b>���ܾ���͡�</b></font><font color=blue>"& aqjh_name &"</font>�ܾ���<font color=blue>"& fromname &"</font>����ͣ��㿴�����Ұ�������4��������Ҫ�����Լ�ûǮô��"& fromname &"����һ���ӻң��������ֺá�������ʱ�½�500��"& aqjh_name &"��������100�㣬�����½�50�㡣"
		conn.execute "update �û� set ����=����-500 where ����='"&fromname&"'"
		conn.execute "update �û� set ����=����+100,����=����-50 where ����='"&aqjh_name&"'"
end select
set rs=nothing
conn.close
set conn=nothing
Application.Lock
says=ab	
says=replace(says,chr(39),"\'")
says=replace(says,chr(34),"\"&chr(34))
act="��Ϣ"
towhoway=0
towho=fromname
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