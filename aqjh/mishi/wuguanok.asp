<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="../const.asp"-->
<!--#include file="../wzconfig.asp"-->
<!--#include file="../config.asp"-->
<body leftmargin="0" topmargin="0" bgcolor="#CCCCCC" background="../jhimg/bgcheetah.gif">
<%
call chkpost()
if Instr(LCase(Application("axjh_useronlinename"&nowinroom))," "&LCase(axjh_name)&" ")=0 or session("axjh_inthechat")<>"1" then
	Response.Write "<script Language=Javascript>alert('�㲻�ܽ��в��������д˲���������������ң�');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
axjh_roominfo=split(Application("axjh_room"),";")
chatinfo=split(axjh_roominfo(nowinroom),"|")
fjname=chatinfo(0)
erase chatinfo
erase axjh_roominfo
if fjname=application("axjh_dbroom") then
	erase axjh_roominfo
	erase chatinfo
	Response.Write "<script language=JavaScript>{alert('��ʾ���ᱦ�����ڼ䲻����ʹ���������䣡');location.href = 'javascript:history.back()';}</script>"
	Response.End
end if
money=int(abs(request.form("money")))
if money<>1000 and  money<>10000 and money<>100000 and money<>150000 and money<>200000 and money<>1000000 then 
	Response.Write "<script Language=Javascript>alert('���ݳ���');location.href = 'javascript:history.back()';</script>"
	response.end
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("axjh_usermdb")
rs.open "select �ȼ�,����,����,�书,�书��,��Ա�ȼ�,��Ա,����ʱ�� from �û� where ����='"&axjh_name&"'",conn
if rs.eof or rs.bof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Redirect "../error.asp?id=210"
	response.end
end if
hydj=rs("��Ա�ȼ�")
select case hydj
	case 0
		bf=1.2
		jgsj=180
	case 1
		bf=1.25
		jgsj=180
	case 2
		bf=1.3
		jgsj=120
	case 3
		bf=1.35
		jgsj=60
	case 4
		bf=1.4
		jgsj=0
end select
sj=DateDiff("s",rs("����ʱ��"),now())
if sj<jgsj then
	s=jgsj-sj
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('����"&hydj&"����Ա��"&jgsj&"��һ�Σ����["&s&"��]��');location.href = 'javascript:history.back()';}</script>"
	Response.End
end if
mytl=rs("����")
mynl=rs("����")
wg=rs("�书")
wgj=rs("�书��")
pdhy=rs("��Ա")
zddj=rs("�ȼ�")
rs.close
if pdhy<>false then
	bf=1.3
end if
wgsx=int((zddj*axjh_wgsx+3800+wgj)*bf)
xtl=money*10
xnl=money*5
newwg=wg+money
if wg>=wgsx then
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('����书�Ѵﵽ���޵�"& bf &"���ˣ��������������ɣ�');location.href = 'javascript:history.back()';</script>"
	response.end
end if
if xtl>mytl or xnl>mynl then
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('���������书������"&money&"���书����������������"&xtl&"��������"&xnl&"����Ŀǰ����ֵΪ��"&mytl&"������ֵΪ��"&mynl&"��������������');location.href = 'javascript:history.back()';</script>"
	response.end
end if
if newwg>=wgsx then
	conn.execute "update �û� set �书="&wgsx&",����=����-"& xtl &",����=����-"& xnl &",����ʱ��=now() where ����='"&axjh_name&"'"
	set rs=nothing
	conn.close
	set conn=nothing
	ltsay="<b><font color=#FF0000>���չ����䡿</font></b><font color=blue>"&axjh_name&"</font>�����ұչ���ϰ���书������书����<b><font color=#FF0000>+"&money&"</font></b>�㣬����<font color=#0000FF>����</font>��<b><font color=#FF0000>-"&xtl&"</font></b>�㣬<font color=#0000FF>����</font><font color=#FF0000><b>-"&xnl&"</b></font>�㣬�书�Ѵ�Ŀǰ״̬��߾��磬��λҪС���ˡ�"
	call lts("��Ϣ","���",ltsay,0,nowinroom)
	Response.Write "<script Language=Javascript>alert('��ͨ���չ�����������ʹ�Լ����书���,�ﵽ��ǰ��߾��磬�书���ǣ�"&money&"�㣬���Ѵﵽ���޵�"& bf &"�����������������ɣ�');window.close();</script>"
else
	conn.execute "update �û� set �书=�书+"& money &",����=����-"& xtl &",����=����-"& xnl &",����ʱ��=now() where ����='"&axjh_name&"'"
	conn.close
	set rs=nothing
	set conn=nothing
	ltsay="<b><font color=#FF0000>���չ����䡿</font></b><font color=blue>"&axjh_name&"</font>�����ұչ���ϰ���书������书����<b><font color=#FF0000>+"&money&"</font></b>�㣬����<font color=#0000FF>����</font>��<b><font color=#FF0000>-"&xtl&"</font></b>�㣬<font color=#0000FF>����</font><font color=#FF0000><b>-"&xnl&"</b></font>�㡣"
	call lts("��Ϣ","���",ltsay,0,nowinroom)
	Response.Write "<script Language=Javascript>alert('��������Ǳ����ϰ��ѧ������ʹ�Լ����书���,�书��+"& money &"�㣬����������-"&xtl&"�㣬������-"&xnl&"�㣡');location.href = 'javascript:history.back()';</script>"
end if
%>
