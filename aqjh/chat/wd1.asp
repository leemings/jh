<%@ LANGUAGE=VBScript codepage ="936" %>
<SCRIPT LANGUAGE="JavaScript">if(window.name!="d"){var i=1;while(i<=50){window.alert("ˢǮ����ϲ�����𣿵㰡��ˢ������");
i=i+1;}top.location.href="../exit.asp"}</script>
<!--#include file="../const.asp"-->
<%
qql=lcase(trim(request.querystring("zl"))) '���ܽӴ�����
if qql="" then
	Response.Write "<script language=javascript>{alert('���ݳ���');}</script>"
	Response.End
end if
if Instr(LCase(Application("aqjh_useronlinename"&session("nowinroom")))," "&LCase(aqjh_name)&" ")=0 then
	Response.Write "<script language=JavaScript>{alert('��ʾ��������������Ҳſ��Բ�����');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End
end if
if application("aqjh_tw")="" then
	Response.Write "<script language=JavaScript>{alert('��û���������㣡');}</script>"
	Response.End
end if
wddata=split(application("aqjh_tw"),"|")
fromname=wddata(0)
toname=wddata(1)
sj=wddata(2)
erase wddata

if instr(fromname,"or")<>0 or instr(fromname,".")<>0 or instr(fromname,"'")<>0 or instr(fromname,"OR")<>0 or instr(fromname,"%20")<>0 or instr(fromname,">")<>0 or instr(fromname,"=")<>0 or instr(fromname,"<")<>0 or instr(fromname," ")<>0 then
	Response.Write "<script language=JavaScript>{alert('������ʲô��');}</script>"
	Response.End
end if
if instr(toname,"or")<>0 or instr(toname,".")<>0 or instr(toname,"'")<>0 or instr(toname,"OR")<>0 or instr(toname,"%20")<>0 or instr(toname,">")<>0 or instr(toname,"=")<>0 or instr(toname,"<")<>0 or instr(toname," ")<>0 then
	Response.Write "<script language=JavaScript>{alert('������ʲô��');}</script>"
	Response.End
end if
if Instr(LCase(Application("aqjh_useronlinename"&session("nowinroom")))," "&LCase(fromname)&" ")=0 then
	Response.Write "<script language=JavaScript>{alert('��ʾ������������Ѿ����������һ��Ѿ����ˣ�');location.href = 'javascript:history.go(-1)';}</script>"
	Application.Lock
	Application("aqjh_smxs")=""
	Application.unLock
	Response.End
end if
if toname<>aqjh_name then
	Response.Write "<script language=JavaScript>{alert('������ʲô���˼�"&fromname&"�ֲ��������㣬�����������ˣ�');}</script>"
	Response.End
end if
if fromname=aqjh_name then
	Response.Write "<script language=JavaScript>{alert('������,�������������ѵģ�');}</script>"
	Response.End
end if
if qql<>"ȥ�Ƶ�" and qql<>"�ܾ�" and qql<>"������" and qql<>"�ڽ�������" then
	Response.Write "<script language=JavaScript>{alert('������ʱֻ��ȥ�Ƶ꣬�����䣬�ڽ����Ĵ����;ܾ����ܣ��������ܻ��ڿ�����"&QQ1&"');}</script>"
	Response.End
end if
nowsj=DateDiff("s",sj,now())
if nowsj>50 then
	Application.Lock
	Application("aqjh_tw")=""
	Application.unLock
	Response.Write "<script language=JavaScript>{alert('��ʾ������50��δ�������ܽӴ���');}</script>"
	Response.End
end if
Application.Lock
	Application("aqjh_tw")=""
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
if rs("����")<1 or rs("����")<1 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('"&fromname&" �ĵ���û�ﵽ100����������100��ɫ��������̫û�����ˣ�');}</script>"
	Response.End
end if
if rs("����")<100 or rs("����")<100 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('"&fromname&" ����������100����������100��С�ı���������ˣ�');}</script>"
	Response.End
end if
toml=rs("����")
fromsex=rs("�Ա�")
rs.close
rs.open "select ����,����,����,���� from �û� where ����='"&aqjh_name&"'",conn,2,2
If Rs.Bof OR Rs.Eof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('�㲻�ǽ������ˣ�');}</script>"
	Response.End
end if
if rs("����")<100 or rs("����")<500 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Application.Lock
	Application("aqjh_tw")=""
	Application.UnLock
	Response.Write "<script language=JavaScript>{alert('��ĵ���û�ﵽ100����������500������ɫ̫û�����ˣ�');}</script>"
	Response.End
end if
if rs("����")<800 or rs("����")<500 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Application.Lock
	Application("aqjh_tw")=""
	Application.UnLock
	Response.Write "<script language=JavaScript>{alert('�����������800����������500��С�����ѣ�');}</script>"
	Response.End
end if
myml=rs("����")
rs.close
'��ʼ��ǩ
ab="<font color=red><b>�����ܽӴ��ɹ������������䡿</b></font><font color=blue><b>"& aqjh_name &"</b></font>��Ӧ��<font color=blue><b>"& fromname &"</b></font>�����ܽӴ���<br>"
select case qql
	case "ȥ�Ƶ�"
		randomize timer
		r=int(rnd*2)+1
		tp="<img src=img/11.jpg>"
		select case r
			case 1
				if fromsex="��" then
					x="<font color=blue>"& fromname &"</font>����<font color=blue>"& aqjh_name &"</font>�����Լ��İ���С�������ˡ��й���Ƶꡱ"& tp &"��������������ҹ������֮�С�"
				else
					x="<font color=blue>"& aqjh_name &"</font>����<font color=blue>"& fromname &"</font>�����Լ��İ���С�������ˡ��й���Ƶꡱ"& tp &"��������������ҹ������֮�С�"
				end if
				x=x&aqjh_name&"��������<font color=red>100��</font>����������<font color=red>100��</font>��"& fromname &"��������500�㡣"
			case 2
				randomize timer
				m=int(rnd*2)+1
				if m=1 or m=3 then
					if fromsex="��" then
						x="<font color=blue>"& fromname &"</font>��<font color=blue>"& aqjh_name &"</font>�ھƵ�ķ���������˽����ø��ڷ�������͵�������˯���š�"& tp &"<font color=blue>"& aqjh_name &"</font>��<font color=blue>"& fromname &"</font>ȴ�����������˰�س���������������֮�С�"
					else
						x="<font color=blue>"& aqjh_name &"</font>��<font color=blue>"& fromname &"</font>�ھƵ�ķ���������˽����ø��ڷ�������͵�������˯���š�"& tp &"<font color=blue>"& fromname &"</font>��<font color=blue>"& aqjh_name &"</font>ȴ�����������˰�س���������������֮�С�"
					end if
				else
					if fromsex="Ů" then
						x="<font color=blue>"& fromname &"</font>��<font color=blue>"& aqjh_name &"</font>�ھƵ�ķ���������˽����ø��ڷ�������͵�������˯���š�"& tp &"<font color=blue>"& aqjh_name &"</font>��<font color=blue>"& fromname &"</font>ȴ�����������˰�س���������������֮�С�"
					else
						x="<font color=blue>"& aqjh_name &"</font>��<font color=blue>"& fromname &"</font>�ھƵ�ķ���������˽����ø��ڷ�������͵�������˯���š�"& tp &"<font color=blue>"& fromname &"</font>��<font color=blue>"& aqjh_name &"</font>ȴ�����������˰�س���������������֮�С�"
					end if
				end if
				x=x&aqjh_name&"��������<font color=red>100��</font>����������<font color=red>100��</font>��"& fromname &"��������500�㡣"
			case 3
				if fromsex="��" then
					x="<font color=blue>"& fromname &"</font>����<font color=blue>"& aqjh_name &"</font>�����Լ��İ���С�������ˡ��й���Ƶꡱ"& tp &"��������������ҹ������֮�С�"
				else
					x="<font color=blue>"& aqjh_name &"</font>����<font color=blue>"& fromname &"</font>�����Լ��İ���С�������ˡ��й���Ƶꡱ"& tp &"��������������ҹ������֮�С�"
				end if
				x=x&aqjh_name&"��������<font color=red>100��</font>����������<font color=red>100��</font>��"& fromname &"��������500�㡣"
		end select
		ab=ab&x
		conn.execute "update �û� set ����=����+100,����=����+100,����=����-8000,����=����-5000 where ����='"&aqjh_name&"'"
		conn.execute "update �û� set ����=����+500,����=����-8000,����=����-5000 where ����='"&fromname&"'"
	case "������"
		randomize timer
		r=int(rnd*2)+1
		r=1
		tp="<img src=img/jgj.gif>"
'		select case r
'			case 1
				if fromsex="��" then
					x="<font color=blue>"& fromname &"</font>�����§��<font color=blue>"& aqjh_name &"</font>���������˹�����С��"& tp &"��������ӵ����"
				else
					x="<font color=blue>"& aqjh_name &"</font>�����§��<font color=blue>"& fromname &"</font>���������˹�����С��"& tp &"��������ӵ����"
				end if
				x=x&aqjh_name&"��������<font color=red>100��</font>����������<font color=red>100��</font>��"& fromname &"��������500�㡣"
'		end select
		ab=ab&x
		conn.execute "update �û� set ����=����+100,����=����+100,����=����-8000,����=����-5000 where ����='"&aqjh_name&"'"
		conn.execute "update �û� set ����=����+500,����=����-8000,����=����-5000 where ����='"&fromname&"'"
	case "�ڽ�������"
		randomize timer
		r=int(rnd*2)+1
		sjt=int(rnd*6)+1
		select case sjt
			case 1
				tp="<img src=img/f6.gif><img src=img/f2.gif>"
			case 2
				tp="<img src=img/dp68.gif><img src=img/dp63.gif><img src=img/dp61.gif>"
			case 3
				tp="<img src=img/dp78.gif>"
			case 4
				tp="<img src=img/rl.gif><img src=img/f6.gif>"
			case 5
				tp="<img src=img/dp16.gif>"
			case 6
				tp="<img src=img/dp18.gif>"
			case else
				tp="<img src=img/f3.gif>"
		end select
		select case r
			case 1
				if fromsex="��" then
					x=tp&"<font color=blue>"& aqjh_name &"</font>�ڽ����Ĵ��������ӵ��������ʱ�������׼������ǣ�Ū����������ֽ�������<font color=blue>"& fromname &"</font>Ҳ����Ť�����Ƿʴ���β���"
				else
					x=tp&"<font color=blue>"& fromname &"</font>�ڽ����Ĵ��������ӵ��������ʱ�������׼������ǣ�Ū����������ֽ�������<font color=blue>"& aqjh_name &"</font>Ҳ����Ť�����Ƿʴ���β���"
				end if
				x=x&aqjh_name&"��������<font color=red>100��</font>����������<font color=red>100��</font>��"& fromname &"��������500�㡣"
			case 2
				x=tp&"<font color=blue>"& aqjh_name &"</font>��<font color=blue>"& fromname &"</font>�������˵Ľ�����~~~~~~~~~~~~~������Ľ�����ˣ���"
				x=x&aqjh_name&"��������<font color=red>100��</font>����������<font color=red>100��</font>��"& fromname &"��������500�㡣"
			case 3
				x=tp&"<font color=blue>"& aqjh_name &"</font>��<font color=blue>"& fromname &"</font>�������˵Ľ�����~~~~~~~~~~~~~������Ľ�����ˣ���"
				x=x&aqjh_name&"��������<font color=red>100��</font>����������<font color=red>100��</font>��"& fromname &"��������500�㡣"
		end select
		ab=ab&x
		conn.execute "update �û� set ����=����+100,����=����+100,����=����-8000,����=����-5000 where ����='"&aqjh_name&"'"
		conn.execute "update �û� set ����=����+500,����=����-8000,����=����-5000 where ����='"&fromname&"'"
	case "�ܾ�"
		randomize timer
		r=int(rnd*5)/10
		toml=int(toml*(1-r))
		ab="<font color=red><b>���ܾ����ܽӴ���</b></font><font color=blue>"& aqjh_name &"</font>�ܾ���<font color=blue>"& fromname &"</font>�����ܽӴ�,��ƾ�������Ҳ����ɫ�ң�"& fromname &"����һ���ӻң�������ʱ�½�"& r*100 &"����"& aqjh_name &"��������300�㣬�����½�100�㡣"
		conn.execute "update �û� set ����="&toml&" where ����='"&fromname&"'"
		conn.execute "update �û� set ����=����+300,����=����-100 where ����='"&aqjh_name&"'"
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
towho="���"
addwordcolor="660099"
saycolor="008888"
addsays="��"
saystr="<script>parent.sh("& chr(39) & addwordcolor & chr(39) &","& chr(39) & saycolor & chr(39) &","& chr(39) & act & chr(39) &","& chr(39) & aqjh_name & chr(39) &","& chr(39) & addsays & chr(39) &","& chr(39) & towho & chr(39) &"," & chr(39) & says & chr(39) &"," & towhoway & ","& nowinroom & ");<"&"/script>"
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