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
rs.open "select ID,�Ա�,����,����,������,�ȼ�,��Ա,��Ա�ȼ�,����ʱ�� from �û� where ����='"& aqjh_name &"'",conn,2,2
if rs.eof or rs.bof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('�㲻�ǽ������ˣ����ɽ�������Ժ��');window.close();</script>"
	response.end
end if
yl=rs("����") 
sex=rs("�Ա�")
hydj=rs("��Ա�ȼ�")
pdhy=rs("��Ա")
nl=rs("����")
%><!--#include file="../../config.asp"--><%
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
nlsx=int((rs("�ȼ�")*aqjh_nlsx+2000+rs("������"))*bf)
if rs("����")>=rs("�ȼ�")*aqjh_nlsx+2000+rs("������") then 
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('�Բ�����������Ѵﵽ������ޣ������ٽ���԰�ˣ�');window.close();</script>"
	response.end
end if
if sex="Ů" then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('����Ů��Ҳ�����ֵط�,С��ץ��ȥ����Ů��!');window.close();</script>"
	response.end
end if
sj=DateDiff("s",rs("����ʱ��"),now())
if sj<120 then
	ss=120-sj
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('���ǲ��ǸմӺ�԰�����ѽ�������"&ss&"�������ɣ�');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
%>
<!--#include file="jiu.asp"-->
<%
rs.close
sql="select * from ��Ů where ID=" & id
Set Rs=connt.Execute(sql)
if rs.eof or rs.bof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	connt.close
	set connt=nothing
	Response.Write "<script Language=Javascript>alert('����û�и��ѽ����԰�������������ѽ���ǲ�������������!');window.close();</script>"
	response.end
end if
jiid=rs("��ŮID")
mingji=rs("����")
meimao=rs("��ò��")
yin=int(meimao*5)
if yl<meimao*5 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	connt.close
	set connt=nothing
	Response.Write "<script Language=Javascript>alert('ûǮҲ�������ָ߼��ĵط�ѽ����ֹ����!');window.close();</script>"
	response.end
end if
yin1=int(meimao*5/2)
nlj=int(meimao/2)
randomize timer
r=int(rnd*3)+1
newnl=nl+nlj
if newnl>=nlsx then
	newnl=nlsx
end if
select case r
	case 1
		mess=aqjh_name &"�ڿ�������԰��С��"&mingji&"��ϥ��̸��������ֱ����ҹ��"&aqjh_name&"����������ٵ㣬��������100�㣬��������"& nlj &"�㣬������԰С��"&mingji&"����"&int(yin)&"����԰������ȡһ��������С��"&mingji&"��ò������10�㣡"
		conn.execute "update �û� set ����=����-" & yin & ",����=����+100,����=����-500,����=" & newnl & ",����ʱ��=now() where ����='" & aqjh_name & "'"
		
		connt.execute "update ��Ů set ��ò��=��ò��+10 where id=" & id
	case 2
		mess=aqjh_name &"�ڿ�������԰��С��"&mingji&"���죬û�뵽����Ͷ�����࣬û˵����ͳ���������"& aqjh_name &"�����½�100�㣬�����½�300�㣬���װ׻���"& yin &"�����ӡ�С��"& mingji &"��ò���½�10�㣡"
		conn.execute "update �û� set ����=����-" & yin & ",����=����-100,����=����-300,����ʱ��=now() where ����='" & aqjh_name & "'"
		
		connt.execute "update ��Ů set ��ò��=��ò��-10 where id=" & id
	case 3
		mess=aqjh_name &"�ڿ�������԰��С��"&mingji&"��̸��������������һЩ��ޣ���������100�㣬��������"& int(meimao/2) &"�㣬����ʱ����"& yin &"�����ӣ�С��"& mingji &"�õ����е�һ�룬��ò������20�㣡"
		conn.execute "update �û� set ����=����-" & yin & ",����=����+100,����=����-500,����=" & newnl & ",����ʱ��=now() where ����='" & aqjh_name & "'"
		
		connt.execute "update ��Ů set ��ò��=��ò��+20 where id=" & id
	case 4
		mess=aqjh_name &"�ڿ�������԰��С��"& mingji &"�������⣬���ֶ��ţ����°ܻ�������Ժ���ڱ���һ�٣������½�100�������½�"& int(meimao/3) &"�������½�100�������½�100��"
		conn.execute "update �û� set ����=����-100,����=����-int(meimao/3),����=����-100,����=����-100,����=����-" & yin & ",����ʱ��=now() where ����='" & aqjh_name & "'"
end select
says="<font color=#ff0000><b>������Ժ��</b></font>"&mess			'��������
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
<html><head>
<title>�������</title>
<style type="text/css">
<!--
table { border: #000000; border-style: solid; border-top-width: 1px; border-right-width: 1px; border-bottom-width: 1px; border-left-width: 1px}
font { font-size: 12px}
.unnamed1 { font-size: 9pt}
-->
</style>
</head>
<body bgcolor="#FFB366">
<p>&nbsp;</p>
<table width="52%" border="0" height="156" bordercolor="#330033" cellspacing="0" cellpadding="0" align="center">
<tr> 
<td height="17" bgcolor="#996633" align="center">&nbsp;</td>
</tr>
<tr bgcolor="#66FF66"> 
<td align="center" height="378" bgcolor="#FFCC66"> 
<p><object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=5,0,0,0" width="550" height="400">
<param name=movie value="girl.swf">
<param name=quality value=high>
<embed src="girl.swf" quality=high pluginspage="http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash" type="application/x-shockwave-flash" width="550" height="400">
</embed> 
</object><font> </font></p>
</td>
</tr>
<tr bgcolor="#0033CC"> 
<td align="center" height="26" class="unnamed1" bgcolor="#FFCC66"><b></b></td>
</tr>
</table>
<p>&nbsp;</p>
</body></html>