<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<!--#include file="func.asp"-->
<!--#include file="../../mywp.asp"-->
<%'ѧϰ
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
response.buffer=false
response.write " "
tfjh_name=session("tfjh_name")
if Instr(LCase(" " & application("tfjh_useronlinename"&nowinroom)),LCase(" "&tfjh_name&" "))=0 then 
	session("tfjh_inthechat")="0" 
	Response.Redirect "chaterr.asp?id=001" 
end if 
towhoway=Request.Form("towhoway")
saycolor=Request.Form("saycolor")
addwordcolor=Request.Form("addwordcolor")
addsays=Request.Form("addsays")
says=Trim(Request.Form("sy"))
'call dianzan(towho)
if len(says)>150 then says=Left(says,150)
says=lcase(says)
says=Replace(says,"&amp;","&")
if tfjh_grade<9 then
	says=bdsays(says)
end if
act=1
towhoway=0
i=instr(says," ")
fnn1=mid(says,i+1)
fnn1=trim(fnn1)
set rs=session("bug_conn_npc").execute("select ��ʽ,ħ��,w9,�ȼ�,������ from �û� where ����='" & tfjh_name & "'")
if instr(";" & rs("w9"),";" & fnn1 & "|")=0 then response.write "<script>alert('��û����ϰ[" & fnn1 & "]���鰡��');</script>" :  response.end
select case fnn1				'�ж��ǲ���ħ��
case "������","������ʾ","������","���ۿ���","ʥ����","�ջ�֮��","��������","������","������","��ڤ��","ԡ��","ʥԡ��","ԡѪ","ʥԡѪ","����"
if instr("," & rs("ħ��"),"," & fnn1 & "|")>0 then response.write "<script>alert('���Ѿ�ѧ������ħ���ˣ�û�б�Ҫ��ѧ��');</script>" :  response.end
select case fnn1
	case "������"
		if rs("�ȼ�")<30 then response.write "<script>alert('���ս���ȼ�С��30�����޷��������֮�����ħ��~��');</script>" :  response.end
		mess1="##����һ����������������ѧ����������<img src='../shenbing/mofa/zhiyu.gif'>"
	case "������ʾ"
		if rs("�ȼ�")<20 then response.write "<script>alert('���ս���ȼ�С��20�����޷��������֮�����ħ��~��');</script>" :  response.end
		mess1="##�̿�������˺ܾã�����������������ʾ<img src='../shenbing/mofa/xinling.gif'>��Ҫ��"
	case "������"
		if rs("�ȼ�")<100 then response.write "<script>alert('���ս���ȼ�С��100�����޷��������֮�����ħ��~��');</script>" :  response.end 
		mess1="##��ϸ�ذѶ������Ŀα�����һ����һ�飬���ڻ����ϸ㶮�˶�����<img src='../shenbing/mofa/dingshen.gif'>���÷�"
	case "���ۿ���"
		if rs("�ȼ�")<90 then response.write "<script>alert('���ս���ȼ�С��90�����޷��������֮�����ħ��~��');</script>" :  response.end 
		mess1="##����������ʧ�ܣ�����������+1�ε�ʱ��ѧ�������ۿ���<img src='../shenbing/mofa/haidi.gif' width=80>"
	case "ʥ����"
		if rs("�ȼ�")<150 then response.write "<script>alert('���ս���ȼ�С��150�����޷��������֮�����ħ��~��');</script>" :  response.end 
		mess1="##��ɤ�Ӷ������ˣ����˺ü�Ƭ��ɤ�Ӻ��������ܺ���ʥ����<img src='../shenbing/mofa/shengyan.gif' width=80>��ѧ����ʥ����"
	case "�ջ�֮��"
		if rs("�ȼ�")<100 then response.write "<script>alert('���ս���ȼ�С��100�����޷��������֮�����ħ��~��');</script>" :  response.end 
		mess1="##�Ծ��˵����������ϰ��N�꣬�����������ջ�֮��<img src='../shenbing/mofa/youhuo.gif' width=80>"
	case "��������"
		if rs("�ȼ�")<250 then response.write "<script>alert('���ս���ȼ�С��250�����޷��������֮�����ħ��~��');</script>" :  response.end 
		mess1="##�о��˰��켯����������һ��żȻ���ᣬʹ##һ�´������˼�������<img src='../shenbing/mofa/jiti.gif' width=80>"
	case "������"
		mess1="##�ѿα�����һ�飬����ѧ���˵�������<img src='../shenbing/mofa/dizhen.gif' width=80>"
	case "������"
		if rs("�ȼ�")<20 then response.write "<script>alert('���ս���ȼ�С��20�����޷��������֮�����ħ��~��');</script>" :  response.end 
		mess1="##�ѿα�����һ�飬����ѧ���������䣡<img src='../shenbing/mofa/xiling.gif' width=80>"
	case "��ڤ��"
		if rs("�ȼ�")<50 then response.write "<script>alert('���ս���ȼ�С��50�����޷��������֮�����ħ��~��');</script>" :  response.end 
		mess1="##�ѿα�����һ�飬����ѧ���˱�ڤ�񹦣�<img src='../shenbing/mofa/beiming.gif' width=80>"
	case "ԡ��"
		if rs("������")<>0 then response.write "<script>alert('���Ѿ����Ǵ���֮��������������ħ���ˣ�');</script>" :  response.end
		mess1="##���ſα��ܵ�����ֵ����ã�ϴ��ϴ�ţ���ѧ����ԡ�飡"
	case "ԡѪ"
		if rs("������")<>0 then response.write "<script>alert('���Ѿ����Ǵ���֮��������������ħ���ˣ�');</script>" :  response.end
		mess1="##���ſα��ܵ�����ֵ����ã�ϴ��ϴ�֣���ѧ����ԡѪ��"
	case "ʥԡ��"
		if instr(rs("ħ��"),"ԡ��")=0 then response.write "<script>alert('�������ѧԡ�飡');</script>" :  response.end
		if rs("������")=0 then response.write "<script>alert('�㻹�Ǵ���֮��������������ħ����');</script>" :  response.end
		mess1="##���ſα��ܵ��ϴ�ֵ����ã�ϴ��ϴ�ţ���ѧ����ʥԡ�飡"		
	case "ʥԡѪ"
		if instr(rs("ħ��"),"ԡѪ")=0 then response.write "<script>alert('�������ѧԡѪ��');</script>" :  response.end
		if rs("������")=0 then response.write "<script>alert('�㻹�Ǵ���֮��������������ħ����');</script>" :  response.end
		mess1="##���ſα��ܵ�����ֵ����ã�ϴ��ϴ�֣���ѧ����ʥԡѪ��"
	case "����"
		if rs("�ȼ�")<50 then response.write "<script>alert('���ս���ȼ�С��50�����޷��������֮�����ħ��~��');</script>" :  response.end 
		mess1="##�Լ��о��˿α����ص�����������ˢ��ˢ�������������[����]��"
end select
tmp=trim(cstr(rs("ħ��") & " "))
if tmp="" then 
	tmp=fnn1 & "|1"
else
	tmp=tmp & "," & fnn1 & "|1"
end if
session("bug_conn_npc").execute("update �û� set ħ��='" & tmp & "',w9='" & abate(rs("w9"),fnn1,1) & "' where ����='" & tfjh_name & "'")
case else					'ѧϰ�书��ʽ
set rsy=session("bug_conn_npc").execute("select * from y where b='NPC' and a='" & fnn1 & "'")
if rsy.eof then response.write "<script>alert('û�������书�������е㲻�ԣ�');</script>" :  response.end
if instr("," & rs("��ʽ"),"," & fnn1 & "|")>0 then response.write "<script>alert('���Ѿ�ѧ�������书�ˣ�û�б�Ҫ��ѧ��');</script>" :  response.end
tmp=trim(cstr(rs("��ʽ") & " "))
if tmp="" then 
	tmp=fnn1 & "|1"
else
	tmp=tmp & "," & fnn1 & "|1"
end if
session("bug_conn_npc").execute("update �û� set ��ʽ='" & tmp & "',w9='" & abate(rs("w9"),fnn1,1) & "' where ����='" & tfjh_name & "'")
mess1="##���˿��α���������и��Ŭ������������[" & fnn1 & "]"
end select
says="<font color=red>��ѧϰ��</font><b>" & mess1 & "</b>"  
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
%>
