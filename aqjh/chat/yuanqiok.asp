<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc/func.asp"-->
<!--#include file="../config.asp"-->
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
nowinroom=session("nowinroom")
aqjh_roominfo=split(Application("aqjh_room"),";")
chatinfo=split(aqjh_roominfo(nowinroom),"|")
if chatinfo(5)<>0 then
 Response.Write "<script language=javascript>{alert(''''��ʾ����["&chatinfo(0)&"]���䲻���Է��У�'''');}</script>"
 Response.End
end if
f=Minute(time()) 
%>
<div align="center"> 
  <p><font color="#FFFFFF"><span style='font-size:9pt'><font size="3">��</font></span></font><a href="yuanqi.asp" target="f3"><font size="3" color="red">Ԫ������</font></a><font size="3" color="#FFFFFF">��</font>
  </p>
</div>
<%
chatbgcolor=Application("aqjh_chatbgcolor")
chatimage=Application("aqjh_chatimage")
if chatbgcolor="" then chatbgcolor="008888"
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
http = Request.ServerVariables("HTTP_REFERER") 
if InStr(http,"chat")=0 then 
 Response.Write "<script language=javascript>{alert(''''�Բ���,����ܾ����Ĳ���������\n     ��ȷ�����أ�'''');}</script>" 
 Response.End 
end if 
if Instr(LCase(Application("aqjh_useronlinename"&session("nowinroom")))," "&LCase(aqjh_name)&" ")=0 then
 Response.Write "<script Language=javascript>alert(''''�㲻�ܽ��в���,���д˲���������������ң�'''');</script>"
 Response.End
end if
to1=trim(request.form("to1"))
hyjzzs=trim(request.form("hyjzzs"))
yuanqi=int(abs(clng(request.form("yuanqi"))))
zsdj=int(abs(clng(request.form("dj"))))
if Instr(dgjjzs,chr(39))<>0 or Instr(dgjjzs,chr(34))<>0 then
	Response.Write "<script Language=Javascript>alert('��Ҫ��ʲô��Զ�㣡');location.href = 'javascript:history.go(-1)'</script>"
	response.end
end if
if Instr(LCase(application("aqjh_zanli")),LCase("!"&aqjh_name&"!"))>0  then
	Response.Write "<script Language=Javascript>alert('�����ڡ����롱״̬��,��ʹ�á��һ����������ܽ����');location.href = 'javascript:history.go(-1)'</script>"
	response.end
end if
if Instr(LCase(application("aqjh_zanli")),LCase("!"&to1&"!"))>0  then
	Response.Write "<script Language=Javascript>alert('"&to1&"���ڡ����롱״̬��,�벻Ҫ������');location.href = 'javascript:history.go(-1)'</script>"
	response.end
end if
if yuanqi<maxyuanqi then
	Response.Write "<script Language=Javascript>alert('��ʾ��Ҫʹ��["&dgjjzs&"],����Ҫʹ��"&maxyuanqi&"Ԫ����');window.location.href='yuanqi.asp';</script>"
	Response.End
end if
if Instr(LCase(Application("aqjh_useronlinename"&session("nowinroom")))," "&LCase(to1)&" ")=0 then
	Response.Write "<script Language=Javascript>alert('��ʾ�����������˲��ڽ����У�');window.location.href='yuanqi.asp';</script>"
	Response.End
end if
if to1="���" or to1=Application("aqjh_automanname") then
	Response.Write "<script Language=Javascript>alert('��ʾ���㲻���ԶԴ�һ򽭺������˷��С�');window.location.href='yuanqi.asp';</script>"
	Response.End
end if
if InStr(";" & Application("aqjh_npc"), ";" & to1 & "|")<>0 then
	Response.Write "<script Language=Javascript>alert('�㲻�ܶ�NPCʹ�ô˲�����');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
stunt=""
stunt1=""
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select * from �û� where ����='" & aqjh_name & "'",conn,2,2
if rs("����")="����" and aqjh_grade<10 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ�����ǳ����˲����Բ�����');window.location.href='yuanqi.asp';}</script>"
	Response.End
end if
if rs("��Ա�ȼ�")<2 and aqjh_name<>"�ٸ�" then
	Response.Write "<script Language=Javascript>alert('��ʾ����Ļ�Ա�ȼ�����[2]��,������ʹ�ô��У�');window.location.href='yuanqi.asp';</script>"
	myclose()
end if
sj=DateDiff("n",rs("����ʱ��"),now())
if sj<5 and aqjh_grade<10 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ����ոձ�����ɱ��,�����������ɣ�');window.location.href='yuanqi.asp';}</script>"
	Response.End
end if
hy=rs("��Ա�ȼ�")
pdhy=rs("��Ա")
if hy=0 and pdhy=False then
 rs.close
 set rs=nothing
 conn.close
 set conn=nothing
 Response.Write "<script Language=javascript>alert('��ѽ,["&aqjh_name &"]��ɲ��ǻ�Ա,����ʹ��Ԫ��������');window.close();</script>"
 response.end
end if
mp=rs("����")
if rs("grade")>=6 and rs("grade")<10 then
	Response.Write "<script Language=Javascript>alert('��ʾ�����ǹٸ���Ա,����ʹ��["&hyjzzs&"]��');window.location.href='yuanqi.asp';</script>"
	myclose()
end if

if rs("Ԫ��")<yuanqi then
	Response.Write "<script Language=Javascript>alert('��ʾ����Ԫ������,���ܷ����У�');window.location.href='yuanqi.asp';</script>"
	myclose()
end if
if rs("����")=True and aqjh_grade<10 then
	Response.Write "<script Language=Javascript>alert('��ʾ������������������,�����Դ�ܣ�');window.location.href='yuanqi.asp';</script>"
	myclose()
end if
if rs("ɱ����")>=int(Application("aqjh_killman")) then
	Response.Write "<script Language=Javascript>alert('��ʾ�������ɱ���˹�����,������ɱ����');window.location.href='yuanqi.asp';</script>"
	myclose()
end if
sj=DateDiff("n",rs("����ʱ��"),now())
if sj<1 then
	s=1-sj
	Response.Write "<script Language=Javascript>alert('���棺���"& s &"���ٷ���,�ɱ����ţ�');window.location.href='yuanqi.asp';</script>"
	myclose()
end if
lxjwg1=rs("�书")
lxjgj1=rs("����")
mydj=rs("�ȼ�")
nn=int(mydj/10)+1
yuanqisx=rs("�ȼ�")*aqjh_yuanqisx+2000+rs("Ԫ����")
yuanqibf=int(yuanqi/yuanqisx)
rs.close
rs.open "select * from �û� where ����='" & to1 & "'",conn,2,2
zstt=rs("ת��")
if zstt>=10 and aqjh_grade<10 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ���Է���ת��10����,�����˲��˶Է���');window.location.href='yuanqi.asp';}</script>"
	Response.End
end if
if rs("����")="����" and aqjh_grade<10 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ�����ǳ����˲����Բ�����');;window.location.href='yuanqi.asp';}</script>"
	Response.End
end if
sj=DateDiff("n",rs("����ʱ��"),now())
if sj<8 and rs("����")="��" and aqjh_grade<10 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ�����ոձ�����ɱ��,�����ȷŹ����ɣ�');window.location.href='yuanqi.asp';}</script>"
	Response.End
end if
jhhy=rs("��Ա�ȼ�")
ntnt=rs("�ȼ�")
tjf=rs("ͨ��")
if rs("grade")>=6 then
	Response.Write "<script Language=Javascript>alert('��ʾ�����ǹٸ���Ա,������ʲô����������վ����Ӧ��');window.location.href='yuanqi.asp';</script>"
	myclose()
end if
if rs("�ȼ�")<70 then
	Response.Write "<script Language=Javascript>alert('��ʾ����������70��,�㲻Ҫ��ô�ݣ�');window.location.href='yuanqi.asp';</script>"
	myclose()
end if
'if rs("����")=true and hyjzzs<>"�ռ�����" then
tomp=rs("����")
lgbh=rs("����")
lxjwg2=rs("�书")
lxjgj2=rs("����")
if hyjzzs="�ռ�����" then
yiyuanqiiang2=rs("����")
else
cdyuanqi=rs("����")
if cdyuanqi>=yuanqi*100 then
yiyuanqiiang1=yuanqi*100
else
yiyuanqiiang1=cdyuanqi
end if
end if
youid=rs("id")
randomize timer
sss=int(rnd*9)+1
qq1=lxjwg1+lxjgj1+100000-lxjwg2+lxjgj2
qq2=(tl+yuanqi)*sss
qq3=(qq1+qq2)*bf
qq4=sqr(qq3)
conn.execute "update �û� set ����=����-100,Ԫ��=Ԫ��-" & yuanqi & ",����ʱ��=now() where ����='" & aqjh_name & "'"
e=""
qi2=conn.execute("select Ԫ�� from �û� where  ����='"&aqjh_name&"'")
yyuanqi=qi2(0)
if jhhy<3 then
	conn.execute "update �û� set ɱ����=ɱ����+1,��ɱ��=��ɱ��+1 where ����='" & aqjh_name & "'"
	if rs("����")=Application("aqjh_baowuname") then
		conn.execute "update �û� set ��������=0,����='"& Application("aqjh_baowuname") &"' where id="&aqjh_id
		conn.execute "update �û� set ��������=0,����='��',����=100,����=2000 where ����='"& to1 &"'"
		stunt=aqjh_name & "��"& to1 &"�ı���:"& Application("aqjh_baowuname") &"���ߡ�����������Ҫ���������ſ��Եõ�����Ķ�����"
	else
	if hyjzzs="�ռ�����" then
		if lgbh=true then
			conn.execute "update �û� set ����=false,״̬='��',����=0,����ʱ��=now(),Ԫ��=0 where ����='" & to1 & "'"
			lgbh=false
			e="������Ҫ����ù�ˡ�"
			stunt=aqjh_name & "����" & yuanqi & "��Ԫ��,��<img src='img/41.gif'><font color=blue>" & to1 & "</font>ʹ��Ԫ������֮<font color=008000>��"&hyjzzs&"��</font>,ȥ����<font color=blue>" & to1 & "</font>��<font color=red>��������<font/>��" & e
		else
			conn.execute "update �û� set ״̬='��',����=0,Ԫ��=0 where ����='" & to1 & "'"
			conn.execute "update �û� set allvalue=allvalue+50,����=����+" & yiyuanqiiang2 & " where ����='" & aqjh_name & "'"
			
			e="��," & to1 & "������<img src=xx/gif/WG8.GIF>������ȥ�����Ӵ˽�����������һֻ��Ϻ," & aqjh_name & "�õ�" & to1 & "�������ֽ�" & yiyuanqiiang2 & "��,���[50]��,����Ԫ��"&yyuanqi&"��," & to1 & "��������Ʒ�Ӵ�ɢ�佭��!"
			fn1=hyjzzs
			call boot(to1,"Ԫ������,�����ߣ�"&aqjh_name&",["&mp&"]"&fn1)
			conn.execute "insert into l(b,a,c,e,d) values ('" & to1 & "',now(),'" & aqjh_name & "','Ԫ������֮"&hyjzzs&"','����')"
		end if
	else
if lgbh=true then
	conn.execute "update �û� set ����=����+100,Ԫ��=Ԫ��+" & yuanqi & ",����ʱ��=now() where ����='" & aqjh_name & "'"
    Response.Write "<script Language=Javascript>alert('����������,��Ĺ������޷��Ƴ�������������,������������˵�ɣ�');window.location.href='javascript:history.go(-1)';</script>"
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.End
else
		conn.execute "update �û� set ״̬='��',����=0,�¼�ԭ��='"&aqjh_name&"|Ԫ������֮"&hyjzzs&"' where ����='" & to1 & "'"
		if tjf=True then
			conn.execute "update �û� set ����=0,���=int(���/2),����=0,����=0 where ����='" & to1 & "'"
		end if
		conn.execute "update �û� set allvalue=allvalue+500,����=����+" & yiyuanqiiang1 & " where ����='" & aqjh_name & "'"
		e="��," & to1 & "������<img src=xx/gif/WG8.GIF>������ȥ�����Ӵ˽���������һ�������," & aqjh_name & "�õ�" & to1 & "���ֽ�" & yiyuanqiiang1 & "�������[500]�㣡����Ԫ��"&yyuanqi&"��"
		fn1=hyjzzs
		call boot(to1,"Ԫ������,�����ߣ�"&aqjh_name&",["&mp&"]"&fn1)
		conn.execute "insert into l(b,a,c,e,d) values ('" & to1 & "',now(),'" & aqjh_name & "','Ԫ������֮"&hyjzzs&"','����')"
	end if
	end if
end if
else
 conn.execute("update �û� set Ԫ��=Ԫ��-"&yuanqi&",����ʱ��=now() where ����='"&to1&"'")
end if
if e="" then
	e="�㡣"
end if
if stunt="" then
	stunt=aqjh_name & "����" & yuanqi & "��Ԫ��,��<img src='img/021.gif'><font color=blue>" & to1 & "</font>ʹ��Ԫ������֮<font color=008000>��"&hyjzzs&"��</font>,ɱ��" & yuanqi & " "& e
end if
rs.close
set rs=nothing
conn.close
set conn=nothing
says="<font color=red><b>��Ԫ������"&hyjzzs&"��</b>"& stunt &"</font>"& t
says=replace(says,chr(39),"\'")
says=replace(says,chr(34),"\"&chr(34))
act="��Ϣ"
towhoway=0
towho="���"
addwordcolor="660099"
saycolor="008888"
addsays="��"
saystr="<script>parent.sh("& chr(39) & addwordcolor & chr(39) &","& chr(39) & saycolor & chr(39) &","& chr(39) & act & chr(39) &","& chr(39) & aqjh_name & chr(39) &","& chr(39) & addsays & chr(39) &","& chr(39) & towho & chr(39) &"," & chr(39) & says & chr(39) &"," & towhoway &  ","& session("nowinroom") & ");<"&"/script>"
addmsg saystr
Function Yushu(a)
	Yushu=(a and 31)
End Function
Sub AddMsg(Str)
Application.Lock()
Application("SayCount")=Application("SayCount")+1
i="SayStr"&YuShu(Application("SayCount"))
Application(i)=Str
Application.Unlock()
End Sub
Response.Write "<script Language=Javascript>alert('��ϲ,����"&hyjzzs&"�Ѿ�������ɣ�');window.location.href='yuanqi.asp';</script>"
function myclose()
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.End
end function
%>
