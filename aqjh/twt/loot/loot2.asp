<%@ LANGUAGE=VBScript codepage ="936" %>
<SCRIPT LANGUAGE=JavaScript>if(window.name!='aqjh_win'){var i=1;while(i<=50){window.alert('ˢǮ����ϲ�����𣿵㰡��ˢ������');i=i+1;}top.location.href='../../exit.asp'}</script>
<!--#include file="../../chat/sjfunc/func.asp"-->
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
'if Instr(LCase(Application("aqjh_useronlinename"&nowinroom))," "&LCase(aqjh_name)&" ")=0 then
if session("aqjh_inthechat")<>1 then
	Response.Write "<script Language=Javascript>alert('�㲻�ܽ��в��������д˲���������������ң�');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
id=request("id")
if InStr(id,"or")<>0 or InStr(id,"'")<>0 or InStr(id,"`")<>0 or InStr(id,"=")<>0 or InStr(id,"-")<>0 or InStr(id,",")<>0 then 
	Response.Write "<script language=JavaScript>{alert('���ɣ�������ʲô���뵷���𣿣�');window.close();}</script>"
	Response.End 
end if
if aqjh_jhdj<3 then
	Response.Write "<script Language=Javascript>alert��'�㻹�ǽ���С�������������ֵط���!');window.close();</script>"
	response.end
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select ID,�Ա�,����,�ȼ�,����,����,������,����ʱ�� from �û� where ����='"& aqjh_name &"'",conn,2,2
if rs.eof or rs.bof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('�㲻�ǽ������ˣ���������Ǯׯ��');window.close();</script>"
	response.end
end if
if rs("����")<10000 then
Response.Write "<script language=JavaScript>{alert('�����������10000��û�ʸ�����Ǯׯ��');window.close();}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
if rs("����")<10000 then
Response.Write "<script language=JavaScript>{alert('�����������10000��û�ʸ�����Ǯׯ��');window.close();}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
if rs("����")<1000000 then
Response.Write "<script language=JavaScript>{alert('�����������һ��������û�ʸ�����Ǯׯ��');window.close();}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
yl=rs("����") 
sex=rs("����")
nl=rs("����")
%>
<%
select case hydj
	case 0
		bf=1.2
	case 1
		bf=1.25
	case 2
		bf=1.30
	case 3
		bf=1.35
	case 4
		bf=1.40
end select
if pdhy<>false then
	bf=1.3
end if
sj=DateDiff("s",rs("����ʱ��"),now())
if sj<120 then
	ss=120-sj
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('���ǲ��Ǹմ�Ǯׯ����ѽ�������"&ss&"�������ɣ�');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
%>
<!--#include file="loot3.asp"-->
<%
rs.close
sql="select * from ���� where ID=" & id
Set Rs=connt.Execute(sql)
if rs.eof or rs.bof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	connt.close
	set connt=nothing
	Response.Write "<script Language=Javascript>alert('����û�и��ѽ���������Ǯׯ!');window.close();</script>"
	response.end
end if
jiid=rs("ID")
mingji=rs("ĬĬ����")
meimao=rs("ĬĬ����")
yin=int(meimao*3)
if yl<meimao*3 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	connt.close
	set connt=nothing
	Response.Write "<script Language=Javascript>alert('ûǮҲ��������ô�߼���Ǯׯ������̫���ˣ���ֹ����!');window.close();</script>"
	response.end
end if
yin1=int(meimao*3/2)
nlj=int(meimao/2)
randomize timer
r=int(rnd*3)+1
newnl=nl+nlj
if newnl>=nlsx then
	newnl=nlsx
end if
select case r
	case 1
	mess=aqjh_name &"�ֳ�һ��AK47,���"&mingji&"��������ʱ�����϶���һ���Ǯ,���������"&meimao&"��"
		conn.execute "update �û� set ����=����-"&meimao*3&",����=����-"&meimao*3&",����ʱ��=now() where ����='"&aqjh_name&"' "
		conn.execute "update �û� set ����=����+"&meimao*3&" where ����='"&aqjh_name&"' "
	case 2
	mess=aqjh_name &"�ֳ�һ��AK47,���"&mingji&"��������ʱ�����϶���һ���Ǯ,���������"&meimao&"�����õ�����500��"
		conn.execute "update �û� set ����=����+500,����=����-"&meimao*3&",����=����-"&meimao*3&",����ʱ��=now() where ����='"&aqjh_name&"' "
		conn.execute "update �û� set ����=����+"&meimao*3&" where ����='"&aqjh_name&"' "
	case 3
	mess=aqjh_name &"�ֳ�һ��AK47,���"&mingji&"��������ʱ�����϶���һ���Ǯ,���������"&meimao&"�����õ��Ṧ500��"
		conn.execute "update �û� set �Ṧ=�Ṧ+500,����=����-"&meimao*3&",����=����-"&meimao*3&",����ʱ��=now() where ����='"&aqjh_name&"' "
		conn.execute "update �û� set ����=����+"&meimao*3&" where ����='"&aqjh_name&"' "	
	case 4
	mess=aqjh_name &"̫̰�ģ�����"&mingji&"���ٸ���ץȥ���η���100����!"
	conn.execute "update �û� set ����=����-1000000,״̬='��',��¼=now()+1/144,����ʱ��=now(),�¼�ԭ��='"&aqjh_name&" ����:"&fn1&"' where ����='" & aqjh_name &"'"
    call boot(aqjh_name,"���Σ������ߣ�"&aqjh_name&","&fn1)
end select
says="<font color=#ff0000><b>���������ű�����</b></font>"&mess			'��������
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
Response.Write "<script Language=Javascript>alert('"&mess&"');window.close();</script>"
%>