<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<!--#include file="func.asp"-->
<%'ħ��
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if Instr(LCase(" " & application("aqjh_useronlinename"&nowinroom)),LCase(" "&aqjh_name&" "))=0 then 
	session("aqjh_inthechat")="0" 
	Response.Redirect "chaterr.asp?id=001" 
end if 

mc_shengyan=10				'ʥ������Ҫ��ħ������
mc_dingshen=1				'������
mc_zhiyu=1						'������
mc_haidilaoyue=5			'��������(���ۿ���)
mc_youhuozhiguang=30		'�ջ�֮��
mc_jitiliaoshang=500			'��������
mc_xinlingqishi=10				'������ʾ
mc_dizhenshu=20					'������
mc_beimingshengong=25			'��ڤ��
mc_xilingzhou=20					'������
mc_yuxue=1								'ԡѪ
mc_syuxue=5							'ʥԡѪ
mc_yuling=10								'ԡ��
mc_syuling=50								'ʥԡ��
mc_jingshen=50 
mc_jingshen=50 						'����

aqjh_name=session("aqjh_name")
towho=Trim(Request.Form("towho"))
towhoway=Request.Form("towhoway")
saycolor=Request.Form("saycolor")
addwordcolor=Request.Form("addwordcolor")
addsays=Request.Form("addsays")
says=Trim(Request.Form("sy"))
call dianzan(towho)
if len(says)>150 then says=Left(says,150)
says=lcase(says)
says=Replace(says,"&amp;","&")
if aqjh_grade<9 then
says=bdsays(says)
end if
act=1
towhoway=0
i=instr(says," ")
fnn1=trim(mid(says,i+1))

set rsuser=session("bug_conn_npc").execute("select * from �û� where ����='" & aqjh_name & "'")
strmagic=rsuser("ħ��")
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select * from �û� where ����='" & towho &"'",conn,2,2
if rs.eof then
   rs.close
   rs.open "select * from [npc] where n����='" & towho &"'",conn,2,2
   if rs.eof then
      wt=""
   else
      wt="npc"
   end if
else
   wt="ren"
end if
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
if wt="npc" then
	set rsn=session("bug_conn_npc").execute("select * from npc where ����='" & towho & "'")
else
	set rsn=session("bug_conn_npc").execute("select * from �û� where ����='" & towho & "'")
end if
if rsn.eof and wt=false and towho<>"���" then response.write "<script>alert('�㲻�ܶ�" & towho & "������');</script>" :  response.end

'response.write strmagic
if instr("," & strmagic,"," & fnn1 & "|")=0 then response.write "<script>alert('��Ŀǰ���޷�����" & fnn1 & "���ָ�����书');</script>" :  response.end

'ħ�������ȼ�����
	dim tmp
	dim tmp1
	tmp=instr("," & strmagic,"," & fnn1 & "|")
	tmp1=right(strmagic,len(strmagic)-tmp+1)
	dj=mid(tmp1 ,instr(tmp1,"|")+1,instr(tmp1 & ",",",")-instr(tmp1,"|")-1)
	if int(sqr(dj+1))-int(sqr(dj))=1 then sjaddwords=",<font color=red><b>" & aqjh_name & "��" & fnn1 & "(" & int(sqr(dj)) & "->" & int(sqr(dj+1)) & ")�����ˣ�</b></font>"
	kuandu=50+int(sqr(dj))
select case fnn1
	case "������"
		says="<font color=red>��ħ����<img src='../shenbing/mofa/zhiyu.gif' width="&kuandu&"><font color=red><b>(" & int(sqr(dj)) & "��)</b></font><font color=" & saycolor & ">" & magic_zhiyu(towho) & "</font>"
	case "ʥ����"
		says="<font color=red>��ħ����<img src='../shenbing/mofa/shengyan.gif' width="&kuandu&"><font color=red><b>(" & int(sqr(dj)) & "��)</b></font><font color=" & saycolor & ">" & magic_shengyan(towho) & "</font>"
	case "������"
		says="<font color=red>��ħ����<img src='../shenbing/mofa/dingshen.gif' width="&kuandu&"><font color=red><b>(" & int(sqr(dj)) & "��)</b></font><font color=" & saycolor & ">" & magic_dingshenshu(towho) & "</font>"
	case "���ۿ���"
		says="<font color=red>��ħ����<img src='../shenbing/mofa/haidi.gif' width="&kuandu&"><font color=red><b>(" & int(sqr(dj)) & "��)</b></font><font color=" & saycolor & ">" & magic_haidilaoyue() & "</font>"
	case "�ջ�֮��"
		says="<font color=red>��ħ����<img src='../shenbing/mofa/youhuo.gif' width="&kuandu&"><font color=red><b>(" & int(sqr(dj)) & "��)</b></font><font color=" & saycolor & ">" & magic_youhuozhiguang(towho) & "</font>"
	case "��������"
		says="<font color=red>��ħ����<img src='../shenbing/mofa/jiti.gif' width="&kuandu&"><font color=red><b>(" & int(sqr(dj)) & "��)</b></font><font color=" & saycolor & ">" & magic_jitiliaoshang() & "</font>"
	case "������ʾ"
		says="<font color=red>��ħ����<img src='../shenbing/mofa/xinling.gif' width="&kuandu&"><font color=red><b>(" & int(sqr(dj)) & "��)</b></font><font color=" & saycolor & ">" & magic_xinlingqishi(towho)  & "</font>" 
	case "������"
		says="<font color=red>��ħ����<img src='../shenbing/mofa/dizhen.gif' width="&kuandu&"><font color=red><b>(" & int(sqr(dj)) & "��)</b></font><font color=" & saycolor & ">" & magic_dizhenshu(towho)  & "</font>" 
	case "������"
		says="<font color=red>��ħ����<img src='../shenbing/mofa/xiling.gif' width="&kuandu&"><font color=red><b>(" & int(sqr(dj)) & "��)</b></font><font color=" & saycolor & ">" & magic_xilingzhou(towho)  & "</font>" 
	case "��ڤ��"
		says="<font color=red>��ħ����<img src='../shenbing/mofa/beiming.gif' width="&kuandu&"><font color=red><b>(" & int(sqr(dj)) & "��)</b></font><font color=" & saycolor & ">" & magic_beimingshengong(towho)  & "</font>" 
	case "ԡ��"
		says="<font color=red>��ħ����<img src='../shenbing/mofa/beiming.gif' width="&kuandu&"><font color=red><b>(" & int(sqr(dj)) & "��)</b></font><font color=" & saycolor & ">" & magic_yuling(towho)  & "</font>" 
	case "ʥԡ��"
		says="<font color=red>��ħ����<img src='../shenbing/mofa/beiming.gif' width="&kuandu&"><font color=red><b>(" & int(sqr(dj)) & "��)</b></font><font color=" & saycolor & ">" & magic_syuling(towho)  & "</font>" 
	case "ԡѪ"
		says="<font color=red>��ħ����<img src='../shenbing/mofa/beiming.gif' width="&kuandu&"><font color=red><b>(" & int(sqr(dj)) & "��)</b></font><font color=" & saycolor & ">" & magic_yuxue(towho)  & "</font>" 
	case "ʥԡѪ"
		says="<font color=red>��ħ����<img src='../shenbing/mofa/beiming.gif' width="&kuandu&"><font color=red><b>(" & int(sqr(dj)) & "��)</b></font><font color=" & saycolor & ">" & magic_syuxue(towho)  & "</font>" 
	case "����"
		says="<font color=red>��ħ����<img src='../shenbing/mofa/beiming.gif' width="&kuandu&"><font color=red><b>(" & int(sqr(dj)) & "��)</b></font><font color=" & saycolor & ">" & magic_jingshen(towho)  & "</font>" 
end select

says=says & sjaddwords
session("bug_conn_npc").execute("update �û� set ħ��='" & replace(strmagic,fnn1 & "|" & dj,fnn1 & "|" & cstr(dj+1)) & "' where ����='" & aqjh_name & "'")
if fnn1="������ʾ" then
	towhoway=1
else
	towhoway=0
end if
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
'ħ��

function magic_zhiyu(towho)					'������
'set rsnpc=session("bug_conn_npc").execute("select * from npc where ����='" & towho & "'")
	need_nl=int(10 * rsuser("�ȼ�") * (sqr(int(sqr(dj)))))
	if rsuser("����")<need_nl then response.write "<script>alert('����������񲻹���~');</script>" :  response.end
	if rsuser("ħ��")<zhiyu then response.write "<script>alert('���ħ������~');</script>" :  response.end
	if bug_ifnpc(towho) then
		if bug_read(towho,"����")>= rsn("����") then response.write "<script>alert('�Է��Ѿ�����Ҫ��������~');</script>" :  response.end
		rsn.close
		set rsn=nothing
		session("bug_conn_npc").execute("update �û� set ����=����-" & need_nl & " where ����='" & aqjh_name & "'")
		call bug_modify(towho,"����",10*need_nl)
	else
		if rsn("����")>=rsn("�ȼ�")*application("aqjh_tlsx")+5260+rsn("������") then response.write "<script>alert('�Է��Ѿ�����Ҫ��������~');</script>" :  response.end
		session("bug_conn_npc").execute("update �û� set ����=����-" & need_nl & " where ����='" & aqjh_name & "'")
		session("bug_conn_npc").execute("update �û� set ����=����+" & need_nl*10 & " where ����='" & towho & "'")
	end if
	
	magic_zhiyu="##ȫ����ɣ���ʼ�˹���%%���ˡ�%%�³�һ����Ѫ����ɫ�������ö���,%%����������<font color=red>" & need_nl*10 & "</font>�㡣##������<font color=red>" &  need_nl & "</font>������" 
end function


function magic_shengyan(towho)					'ʥ����	
	if wt<>"npc" then response.write "<script>alert('ʥ����ֻ�ܶ�NPCʹ��~');</script>" :  response.end
	tmp3=int(bug_read(towho,"����"))-9
	if rsuser("ħ��")<mc_shengyan*tmp3 then response.write "<script>alert('���ħ������~');</script>" :  response.end
	if tmp3>sqr(dj) then response.write "<script>alert('�Է�̫ǿ���ˣ������ǲ���');</script>" :  response.end
	session("bug_conn_npc").execute("update �û� set ħ��=ħ��-" & mc_shengyan*tmp3 & " where ����='" & aqjh_name & "'")
	randomize(timer)
	if rnd<0.8 then
		session("bug_conn_npc").execute("update �û� set ħ��=ħ��-" & mc_shengyan*tmp3 & " where ����='" & aqjh_name & "'")
		magic_shengyan="##�����˼������ͻȻ���ذ���һ����ϯ�������������һ���ҽУ�%%��Ȼ��ʧ����Ӱ����,##������<font color=red>" &  mc_shengyan*tmp3 & "</font>��ħ��" 
	else
		magic_shengyan="##�����˼��������ϧ%%ûʲô��Ӧ��##������<font color=red>" &  mc_shengyan*tmp3 & "</font>��ħ��" 
	end if
end function
	
function magic_youhuozhiguang(towho)			'�ջ�֮��
	if wt<>"npc" then response.write "<script>alert('��ħ��ֻ�ܶ�NPCʹ��~');</script>" :  response.end
	if bug_read(towho,"����")=aqjh_name then response.write "<script>alert('ιιι���ջ��Լ���NPC���');</script>" :  response.end
	tmp3=int(bug_read(towho,"����"))-9
	if rsuser("ħ��")<mc_youhuozhiguang*tmp3 then response.write "<script>alert('���ħ������~');</script>" :  response.end
	if tmp3>sqr(dj) then response.write "<script>alert('�Է�̫ǿ���ˣ������ǲ���');</script>" :  response.end
	session("bug_conn_npc").execute("update �û� set ħ��=ħ��-" & mc_youhuozhiguang*tmp3 & " where ����='" & aqjh_name & "'")
	
	randomize(timer)
	if rnd<0.8 then
		call bug_modify(towho,"����",aqjh_name)
		magic_youhuozhiguang="##����%%�����˼������ͻȻ��������â����%%,%%һ����Ŀ�⽩ֱ����Ȼ�Թ���##�ķԸ���,##������<font color=red>" &  mc_youhuozhiguang*tmp3 & "</font>��ħ��" 
	else
		magic_youhuozhiguang="##����%%�����˼��������ϧ%%ûʲô��Ӧ,##������<font color=red>" &  mc_youhuozhiguang*tmp3 & "</font>��ħ��" 
	end if
end function

function magic_dingshenshu(towho)						'������
	if bug_ifnpc(towho)=false then response.write "<script>alert('��ħ��ֻ�ܶ�NPCʹ��~');</script>" :  response.end
	tmp3=int(bug_read(towho,"����"))
	if rsuser("ħ��")<mc_dingshen*tmp3 then response.write "<script>alert('���ħ������~');</script>" :  response.end
	session("bug_conn_npc").execute("update �û� set ħ��=ħ��-" & mc_dingshen*tmp3 & " where ����='" & aqjh_name & "'")
	call bug_modify(towho,"ʱ��",cstr(now()+(sqr(sqr(dj))*60+10)/864000))
	magic_dingshenshu="##����%%�����һ���������� %%�ԹԲ�����,Ҫ��" & cstr(now()+(sqr(sqr(dj))*60+10)/864000) & "�ſ��Ի,##������<font color=red>" &  mc_dingshen*tmp3 & "</font>��ħ��" 
end function

function magic_yuling(towho)									'ԡ��
	'if aqjh_name<>towho then response.write "<script>alert('��ħ��ֻ�ܶ��Լ�ʹ��~');</script>" :  response.end
	if rsuser("ħ��")<mc_yuling*int(sqr(dj)) then response.write "<script>alert('���ħ������~');</script>" :  response.end
    session("magic_gj")=3
    session("magic_baodou")=0.1+int(int(sqr(dj))^0.3*100)/100
hchyl=0.1+int(int(sqr(dj))^0.3*100)/100
	session("bug_conn_npc").execute("update �û� set ħ��=ħ��-" & mc_yuling*int(sqr(dj))& ",����=����*" & hchyl & "  where ����='" & aqjh_name & "'")
	magic_yuling="##��ʳָѹס������˲�����һ�ɾ�����ʩչ[ԡ��]�����䵽##����,һ˲��##����������<font color=red><b>" & 0.1+int(int(sqr(dj))^0.3*100)/100 & "</b></font>��!������<font color=red>" &  mc_yuling*int(sqr(dj)) & "</font>��ħ��." 
end function

function magic_syuling(towho)									'ʥԡ��
	'if aqjh_name<>towho then response.write "<script>alert('��ħ��ֻ�ܶ��Լ�ʹ��~');</script>" :  response.end
	if rsuser("ħ��")<mc_syuling*int(sqr(dj)) then response.write "<script>alert('���ħ������~');</script>" :  response.end
    session("magic_gj")=30
    session("magic_baodou")=0.1+int(int(sqr(dj))^0.3*100)/100
hchsyl=0.1+int(int(sqr(dj))^0.3*100)/100
	session("bug_conn_npc").execute("update �û� set ħ��=ħ��-" & mc_syuling*int(sqr(dj)) & ",����=����*" & hchsyl & " where ����='" & aqjh_name & "'")
	magic_syuling="##��ʳָѹס������˲�����һ�ɾ�����ʩչ[ʥԡ��]�����䵽##����,һ˲��##����������<font color=red><b>" & 0.1+int(int(sqr(dj))^0.3*100)/100 & "</b></font>��!������<font color=red>" &  mc_syuling*int(sqr(dj)) & "</font>��ħ��." 
end function

function magic_yuxue(towho)									'ԡѪ
	if aqjh_name<>towho then response.write "<script>alert('��ħ��ֻ�ܶ��Լ�ʹ��~');</script>" :  response.end
	if rsuser("ħ��")<mc_yuxue*int(sqr(dj)) then response.write "<script>alert('���ħ������~');</script>" :  response.end
    session("magic_fy")=3
    session("magic_baodou1")=0.1+int(int(sqr(dj))^0.3*100)/100
hchyx=0.1+int(int(sqr(dj))^0.3*100)/100
	session("bug_conn_npc").execute("update �û� set ħ��=ħ��-" & mc_yuxue*int(sqr(dj)) & ",����=����*" & hchyx & " where ����='" & aqjh_name & "'")
	magic_yuxue="##ҧ����⣬�۵�ͻ��һ����Ѫ��˲�����һ������ʩչ[ԡѪ]�����䵽##����,һ˲��##����������<font color=red><b>" & 0.1+int(int(sqr(dj))^0.3*100)/100 & "</b></font>��!������<font color=red>" &  mc_yuxue*int(sqr(dj)) & "</font>��ħ��." 
end function

function magic_syuxue(towho)									'ʥԡѪ
	'if aqjh_name<>towho then response.write "<script>alert('��ħ��ֻ�ܶ��Լ�ʹ��~');</script>" :  response.end
	if rsuser("ħ��")<mc_syuxue*int(sqr(dj)) then response.write "<script>alert('���ħ������~');</script>" :  response.end
    session("magic_fy")=10
    session("magic_baodou1")=0.1+int(int(sqr(dj))^0.3*100)/100
hchsyx=0.1+int(int(sqr(dj))^0.3*100)/100
	session("bug_conn_npc").execute("update �û� set ħ��=ħ��-" & mc_syuxue*int(sqr(dj)) & ",����=����*" & hchsyx & " where ����='" & aqjh_name & "'")
	magic_syuxue="##ҧ����⣬�۵�ͻ��һ����Ѫ��˲�����һ������ʩչ[ʥԡѪ]�����䵽##����,һ˲��##����������<font color=red><b>" & 0.1+int(int(sqr(dj))^0.3*100)/100 & "</b></font>��!������<font color=red>" &  mc_syuxue*int(sqr(dj)) & "</font>��ħ��." 
end function

function magic_jingshen(towho)									'����
	if aqjh_name<>towho then response.write "<script>alert('��ħ��ֻ�ܶ��Լ�ʹ��~');</script>" :  response.end
	if rsuser("ħ��")<mc_jingshen then response.write "<script>alert('���ħ������~');</script>" :  response.end
	
	baolou=""
	for i = 1 to 6
		c1="z" & i
		c2=rsn(c1)
		if c2<>"" then 	
			c3=left(c2,instr(c2,"|")-1)
			application.lock
			application("aqjh_baowp") = application("aqjh_baowp") & c3 & "|" & aqjh_name & "|" & now() & ";"
			application.unlock
			baolou =  baolou & "<input type=submit name=id value='" & c3 & "' style='border: 1px solid;font-size: 9pt;' onclick=window.open('bao.asp?id=" & c3 & "','d');>"
		end if
	next
	session("bug_conn_npc").execute("update �û� set z1='',z2='',z3='',z4='',z5='',z6='' where ����='" & aqjh_name & "'")
	session("bug_conn_npc").execute("update �û� set ħ��=ħ��-" & mc_jingshen & " where ����='" & aqjh_name & "'")
	magic_jingshen="##�����˹������ɰ׵��죬�ɺ쵽�ϣ����ϵ��ڣ�ͻȻ���һ��:����!�� ���ϵ�װ��ȫ���������" & baolou & " ##������<font color=red>" &  mc_jingshen & "</font>��ħ��." 
	baolou=""
end function

function magic_haidilaoyue()					'�������£����ۿ��⣩
	'call aqjh_autodel()
	baowp=application("aqjh_baowp")
	wlist=split(baowp,";")
	for each j in wlist
		if trim(j)<>"" then
			zlist=split(j,"|")
			tmps = tmps & "<input type=submit name=id value='" & zlist(0) & "' style='border: 1px solid;font-size: 9pt;' onclick=window.open('bao.asp?id=" & zlist(0) & "','d');> "
		end if
	next
	session("bug_conn_npc").execute("update �û� set ħ��=ħ��-" & mc_haidilaoyue & " where ����='" & aqjh_name & "'")
	magic_haidilaoyue="##΢һ����ʩչ�����ۿ��⡹֮����˫�ֺ�Ϊ��ָ����ǰ����������˫�۾���Ϊ���ɫ���·��ܶ���һ�У������ϵ���ƷҲȫ����ʾ�˳�����" & tmps & ",##������<font color=red>" &  mc_haidilaoyue & "</font>��ħ��"
end function

function magic_jitiliaoshang()							'��������
	if rsuser("ħ��")<mc_jitiliaoshang then response.write "<script>alert('���ħ������~');</script>" :  response.end
	ql_nlsx=rsuser("�ȼ�")*application("aqjh_nlsx")+5260+rsuser("������")
	if rsuser("����")<ql_nlsx*0.9 then response.write "<script>alert('�����������ųɣ�����һ�˹���������Σ�գ�');</script>" :  response.end
	liaoed=int(ql_nlsx*(sqr(int(sqr(dj)))/30+0.1))
	if liaoed>rsuser("����") then liaoed=rsuser("����")
	tmp=trim(application("aqjh_useronlinename" & session("nowinroom")))
	while instr(tmp,"  ")>0 
		tmp=replace(tmp,"  "," ")
	wend
	tmp1=split(tmp," ")
	for each h in tmp1
		if bug_ifnpc(h)  then
			if  bug_read(h,"����")<>"��" then
				set rs1=session("bug_conn_npc").execute("select * from npc where ����='" & h & "'")
				call bug_modify(h,"����",rs1("����")-bug_read(h,"����"))
			end if
		else
			set rs1=session("bug_conn_npc").execute("select * from �û� where ����='" & h & "'")
			session("bug_conn_npc").execute("update �û� set ����=" & rs1("�ȼ�")*application("aqjh_tlsx")+5260+rs1("������") & " where ����='" & h & "'")
		end if
	next
	session("bug_conn_npc").execute("update �û� set ħ��=ħ��-" & mc_jitiliaoshang & ",����=" & ql_nlsx*0.1 & " where ����='" & aqjh_name & "'")
	magic_jitiliaoshang="##ʹ�������������������˵�������������,##������<font color=red>" &  mc_jitiliaoshang & "</font>��ħ��,�������ڷ������ȣ��Լ�������ֻʣ��<font color=red>" & liaoed & "</font>�㡣" 
end function

function magic_xinlingqishi(towho)					'������ʾ
	if rsuser("ħ��")<mc_xinlingqishi then response.write "<script>alert('���ħ������~');</script>" :  response.end
	if ifnpc then
		g_tl=bug_read(towho,"����")
		g_nl=bug_read(towho,"����")
		g_l=int(100*g_tl/rsn("����"))
		if g_l<0 then g_l=0
		if g_l>100 then g_l=100
		g_r=100-g_l
		tlt="<img height=6 border=0 src=img/11.gif width=" & g_l & "><img height=6 border=0 src=img/5.gif width=" & g_r & ">"
		strout= "##������<font color=red>" & mc_xinlingqishi  & "</font>��ħ����������<b>" & towho & "</b><img src='npc/" & bug_read(towho,"ͼƬ") & "'>����:" & tlt & "(" & g_tl & "/" & rsn("����") & ")"
		if dj>60 then 
			g_l=int(100*g_nl/rsn("����"))
			if g_l<0 then g_l=0
			if g_l>100 then g_l=100
			g_r=100-g_l
			nlt="<img height=6 border=0 src=img/11.gif width=" & g_l & "><img height=6 border=0 src=img/5.gif width=" & g_r & ">"
			strout = strout & " ����:" & nlt & "(" & g_nl & "/" & rsn("����") & ")"
		end if
		if dj>80 then strout = strout & " ����:" & bug_read(towho,"����")
		if dj>120 then strout = strout & " ����:" & bug_read(towho,"����")
		if dj>150 then strout = strout & " ����:" & bug_read(towho,"����") 
		if dj>160 then strout = strout & " ����:" & bug_read(towho,"����")
	else
		g_tl=rsn("����")
		g_nl=rsn("����")
		g_n=rsn("�ȼ�")*application("aqjh_tlsx")+5260+rsn("������")
		g_l=int(100*g_tl/g_n)
		if g_l<0 then g_l=0
		if g_l>100 then g_l=100
		g_r=100-g_l
		tlt="<img height=6 border=0 src=img/11.gif width=" & g_l & "><img height=6 border=0 src=img/5.gif width=" & g_r & ">"
		strout= "##������<font color=red>" & mc_xinlingqishi  & "</font>��ħ����������<b>" & towho & "</b><img src='" & rsn("����ͷ��") & "'>����:" & tlt & "(" & rsn("����") & "/" & g_n  & ")"
		if dj>60 then 
			g_t=rsn("�ȼ�")*application("aqjh_nlsx")+2000+rsn("������")
			g_l=int(100*g_nl/g_t)
			if g_l<0 then g_l=0
			if g_l>100 then g_l=100
			g_r=100-g_l
			nlt="<img height=6 border=0 src=img/11.gif width=" & g_l & "><img height=6 border=0 src=img/5.gif width=" & g_r & ">"
			strout = strout & " ����:" & nlt & "(" & rsn("����") & "/" & g_t & ")"
		end if
		if dj>200 then strout = strout & " ����:" & rsn("����")
		if dj>300 then strout = strout & " ����:" & rsn("����")
		if dj>380 then strout = strout & " <BR>ͷ��:" & myzb("z1") 
		if dj>460 then strout = strout & " <BR>װ��:" & myzb("z2") 
		if dj>500 then strout = strout & " <br>����:" & myzb("z3") 
		if dj>550 then strout = strout & " <br>����:" & myzb("z4") 
		if dj>600 then strout = strout & " <br>����:" & myzb("z5") 
		if dj>800 then strout = strout & " <br>˫��:" & myzb("z6") 
	end if
	session("bug_conn_npc").execute("update �û� set ħ��=ħ��-" & mc_xinlingqishi & " where ����='" & aqjh_name & "'")
	magic_xinlingqishi=strout 
end function

function magic_dizhenshu(towho)						'������
	if towho="���" or towho=aqjh_name or wt="npc" then response.write "<script>alert('�㲻�ܶԴ�һ��Լ�����NPCʹ�õ�����~');</script>" :  response.end
	if rsuser("ħ��")<mc_dizhenshu then response.write "<script>alert('���ħ�����񲻹�Ŷ');</script>" :  response.end
	session("bug_conn_npc").execute("update �û� set ħ��=ħ��-" & mc_dizhenshu & " where ����='" & aqjh_name & "'")
	
		dizhenshu_tl=int(sqr(int(sqr(dj)))*3000*rsuser("�ȼ�"))
		if dizhenshu_tl>rsuser("����")*0.1 then dizhenshu_tl=int(rsuser("����")*0.1)
		dizhenshu_nl=int(sqr(int(sqr(dj)))*3000*rsuser("�ȼ�"))
		if dizhenshu_nl>rsuser("����")*0.1 then dizhenshu_nl=int(rsuser("����")*0.1)
		session("bug_conn_npc").execute("update �û� set ����=����-" & dizhenshu_nl & ",����=����-" & dizhenshu_tl & " where ����='" & towho & "'")
	
	magic_dizhenshu="##�����۾���ĬĬ�������%%��߾�Ȼð��һ�������̣�ͻȻ��¡¡һ����%%�⵽�����Ϯ����%%�����½�<font color=red>" & dizhenshu_tl & "</font>��,�����½�<font color=red>" & dizhenshu_nl & "</font>��.##�ķ���<font color=red>" & mc_dizhenshu & "</font>��ħ��"
end function

function magic_xilingzhou(towho)							'������
	if towho="���" or towho=aqjh_name or wt="npc" then response.write "<script>alert('�㲻�ܶԴ�һ��Լ�����NPCʹ��������~');</script>" :  response.end
	if rsuser("ħ��")<mc_xilingzhou then response.write "<script>alert('���ħ�����񲻹�Ŷ');</script>" :  response.end
	
		xilingzhou_tl=int(sqr(int(sqr(dj)))*1000*rsuser("�ȼ�"))
		if xilingzhou_tl>rsuser("����")*0.1 then xilingzhou_tl=int(rsuser("����")*0.1)
		if xilingzhou_tl>rsn("����") then xilingzhou_tl=rsn("����")
		session("bug_conn_npc").execute("update �û� set ����=����-" & xilingzhou_tl & " where ����='" & towho & "'")	
		session("bug_conn_npc").execute("update �û� set ����=����+" & xilingzhou_tl & ",ħ��=ħ��-" & mc_xilingzhou & " where ����='" & aqjh_name & "'")
	magic_xilingzhou="##ߴ�ﹾ������һ�������%%����һ֧ѩ�ѣ�ƴ�����������%%����ȡ������<font color=red>" & xilingzhou_tl & "</font>�㣬##�ķ���<font color=red>" & mc_xilingzhou & "</font>��ħ����ͬʱ������������<font color=red>" & xilingzhou_tl & "</font>��"
end function

function magic_beimingshengong(towho)					'��ڤ��
	if towho="���" or towho=aqjh_name or wt="npc" then response.write "<script>alert('�㲻�ܶԴ�һ��Լ�����NPCʹ�ñ�ڤ��~');</script>" :  response.end
	if rsuser("ħ��")<mc_beimingshengong then response.write "<script>alert('���ħ�����񲻹�Ŷ');</script>" :  response.end
	
		beimingshengong_nl=int(sqr(int(sqr(dj)))*2000*rsuser("�ȼ�"))
		if beimingshengong_nl>rsuser("����")*0.1 then beimingshengong_nl=int(rsuser("����")*0.1)
		session("bug_conn_npc").execute("update �û� set ����=����-" & beimingshengong_nl & " where ����='" & towho & "'")
		session("bug_conn_npc").execute("update �û� set ����=����+" & beimingshengong_nl & ",ħ��=ħ��-" & mc_beimingshengong & " where ����='" & aqjh_name & "'")
	magic_beimingshengong="##������������һ�������%%����һ�鵰�⣬����ؽ������%%����ȡ������<font color=red>" & beimingshengong_nl & "</font>�㣬##�ķ���<font color=red>" & mc_beimingshengong & "</font>��ħ����ͬʱ������������<font color=red>" & beimingshengong_nl & "</font>��"	
end function

function myzb(lb)
if isnull(rsn(lb)) or trim(rsn(lb))="" then
	myzb="δװ���κ���Ʒ!"
	exit function
end if
data1=split(rsn(lb),"|")
set rs=session("bug_conn_npc").execute("select * FROM b where a='" & data1(0) &"'")
If Rs.Bof OR Rs.Eof then
	myzb="δװ���κ���Ʒ!"
	exit function
end if
if instr(rs("b"),"[")>0 then wid="65" else wid="40"
myzb="<img src='../hcjs/jhjs/images/" & rs("i") & "' alt='"&rs("c")&"' width=" & wid & ">["&rs("a")&"] ����:"&rs("f")&" ����:"&rs("g")&" �;�:<font color=red>"&data1(1)&"</font>/<font color=blue>"&rs("j")&"</font> ����:"&rs("k")
rs.close
end function
%>
