<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<%'���߶Ĳ�Ѻ���wWw.happyjh.com��
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
sjjh_roominfo=split(Application("sjjh_room"),";")
chatinfo=split(sjjh_roominfo(nowinroom),"|")
if chatinfo(9)<>0 then
	Response.Write "<script language=JavaScript>{alert('��ʾ����["&chatinfo(0)&"]���䲻�ܹ��Ĳ���');}</script>"
	Response.End
end if
if Instr(LCase(Application("sjjh_useronlinename"&nowinroom)),LCase(" "&sjjh_name&" "))=0 then 
	Session("sjjh_inthechat")="0" 
	Response.Redirect "chaterr.asp?id=001" 
end if 
towho=Trim(Request.Form("towho"))
towhoway=Request.Form("towhoway")
saycolor=Request.Form("saycolor")
addwordcolor=Request.Form("addwordcolor")
addsays=Request.Form("addsays")
says=Trim(Request.Form("sy"))
call dianzan(towho)
if len(says)>150 then says=Left(says,150)
says=lcase(says)
says=Replace(says,"&amp;","")
if sjjh_grade<9 then
says=bdsays(says)
end if
act=1
towhoway=0
i=instr(says,"$")
fnn1=mid(says,i+1)
says="<font color=green>�����߶Ĳ���<font color=" & saycolor & ">"+xiazhu(mid(says,i+1))+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
'���߶Ĳ�Ѻ��
function xiazhu(fn1)
randomize ()
m1 = Int(6 * Rnd + 1)
if instr(fn1,"&")=0 or right(fn1,1)="&" then
	Response.Write "<script language=JavaScript>{alert('�������󣬸�ʽ���£�[�ľֲ���&Ѻ����&����]');}</script>"
	Response.End 
end if
zt=split(fn1,"&")
abcc=trim(zt(2))
if not isnumeric(abcc) then
	Response.Write "<script language=JavaScript>{alert('["&abcc&"]��������������ʹ�����֣�');}</script>"
	Response.End 
end if
yinjindian=trim(zt(0))
yacz=zt(1)
yinliang=abs(int(zt(2)))
if yinjindian<>"��" and yinjindian<>"��" and yinjindian<>"��" then
	Response.Write "<script language=JavaScript>{alert('ѡ�����������𡢵㣡');}</script>"
	Response.End 
end if
select case yacz
	case "��"
		duboimg="<img src='../jhimg/da.gif'>"
	case "С"
		duboimg="<img src='../jhimg/xiao.gif'>"
	case "��"
		duboimg="<img src='../jhimg/dan.gif'>"
	case "˫"
		duboimg="<img src='../jhimg/shuang.gif'>"
case else
	Response.Write "<script language=JavaScript>{alert('Ѻ����Ϊ����С������˫��');}</script>"
	Response.End 
end select
	if (yinliang<1000 or yinliang>2000000) and yinjindian="��" then
		Response.Write "<script language=JavaScript>{alert('[��]����Ѻ��1000�������200��');}</script>"
		Response.End 
	end if
	if (yinliang<10 or yinliang>200) and (yinjindian="��" or yinjindian="��") then
		Response.Write "<script language=JavaScript>{alert('["&yinjindian&"]����Ѻ��10�����200"&yinjindian&"��');}</script>"
		Response.End 
	end if

'��ʼ�ж�
	Set conn=Server.CreateObject("ADODB.CONNECTION")
	Set rs=Server.CreateObject("ADODB.RecordSet")
	conn.open Application("sjjh_usermdb")
select case yinjindian
'�ԶĲ������ж�
case "��"
	randomize(timer())
	m2 = Int(6 * Rnd + 1)
		rs.open "select ���� FROM �û� WHERE ����='" & sjjh_name &"'",conn
		if rs("����")<yinliang then
			rs.close
			set rs=nothing	
			conn.close
			set conn=nothing
			Response.Write "<script language=JavaScript>{alert('������ô���Ǯ�𣿣�');}</script>"
			Response.End 		
		end if
		rs.close
	rs.open "select top 1 a FROM g WHERE c='��' and b='ׯ��' and f=true",conn
	if rs.eof or rs.bof then
		rs.close
		set rs=nothing	
		conn.close
		set conn=nothing
		Response.Write "<script language=JavaScript>{alert('����[����]û��ׯ�ң��������߶Ĳ�����');}</script>"
		Response.End 		
	end if
	rs.close
	rs.open "select top 1 * FROM g WHERE c='��' and a='"& sjjh_name &"' and f=true",conn,2,2
	if not(rs.eof) or not(rs.bof) then
		if rs("b")="ׯ��" then
			temp=sjjh_name&"��������[����]ׯ�ң���Ҫ��ʲôѽ!"
		else
			temp=sjjh_name&"����[����]Ϊ��ѹ["&rs("e")&"]"&rs("d")&"���ſ��ְɣ�"
		end if
		rs.close
		set rs=nothing	
		conn.close
		set conn=nothing
		Response.Write "<script language=JavaScript>{alert('"&temp&"');}</script>"
		Response.End 
	end if
	rs.close
	rs.open "select top 1 id,a FROM g WHERE c='��' and b='���' and f=false",conn,2,2
	tempdu=0
	if rs.eof or rs.bof then
		Response.Write "<script language=JavaScript>{alert('��ʾ��������������û�п�λ���ڿ���!');}</script>"
		tempdu=1
		rs.close
	end if
	if 	tempdu=0 then
	id=rs("id")
	conn.execute "update �û� set ����=����-"&yinliang&" where ����='" & sjjh_name &"'"
	conn.execute "update g set a='"&sjjh_name&"',f=true,d="&yinliang&",e='"&yacz&"' where id="&id&" and f=false"
	rs.close
	end if
	tmprs=conn.execute("Select count(*) As ���� from g where c='��' and e='��' and f=true")
	dars=tmprs("����")
	tmprs=conn.execute("Select count(*) As ���� from g where c='��' and e='С' and f=true")
	xiaors=tmprs("����")
	tmprs=conn.execute("Select count(*) As ���� from g where c='��' and e='��' and f=true")
	danrs=tmprs("����")
	tmprs=conn.execute("Select count(*) As ���� from g where c='��' and e='˫' and f=true")
	shuangrs=tmprs("����")
	tmprs=conn.execute("Select count(*) As ���� from g where c='��' and d<>0 and f=true")
	durs=tmprs("����")

'������
if durs>=10 then
	sj=DateDiff("s",Application("sjjh_du"),now())
	if sj<5 then
		s=5-sj
		Response.Write "<script language=JavaScript>{alert('��ʾ�����������������ݣ���������["&s&"����]�ٲ�����');}</script>"
		Response.End
	end if	
	Application.Lock
	Application("sjjh_du")=now()
	Application.UnLock
	randomize()
	m3 = Int(6 * Rnd + 1)
	sjdubo=m1+m2+m3
'����ׯ��
rs.open "select * FROM g WHERE c='��' and b='ׯ��' and f=true",conn,2,2
zhuangjia=rs("a")
rs.close

'���Ӵ���
if m1=m2 and m2=m3 and m3=m1 then
	rs.open "select * FROM g WHERE c='��' and d<>0 and b<>'ׯ��' and f=true",conn,2,2
	yingyin=0
	do while not rs.bof and not rs.eof
	yingyin=yingyin+rs("d")
	rs.movenext
	loop
	rs.close
	'��ׯ��Ǯ
	conn.execute "update �û� set ����=����+10000,���=���+"&yingyin&" where ����='"& zhuangjia &"'"
	'������жĲ�״̬
	conn.execute "update g set f=false where f=true and c='��'"
	xiazhu="[������]ׯ�ҿ���<img src='../jhimg/"&m1&".gif'><img src='../jhimg/"&m2&".gif'><img src='../jhimg/"&m3&".gif'>�ƣ�"& sjdubo &"��!ׯ�ҿ������ӡ���ͨɱ��ׯ�ң�"&zhuangjia&"���룺"&yingyin&"����������꣬����+1��㣡"
		set rs=nothing
		conn.close
		set conn=nothing
		exit function
end if

'��ʼ������
shuangyinliang=0
shuangname=""
danyinliang=0
danname=""
xiaoyinliang=0
xiaoname=""
dayinliang=0
daname=""

'����˫����
if sjdubo/2=int(sjdubo/2) then
	danshuang="<img src='../jhimg/shuang.gif'>"
	rs.open "select * FROM g WHERE c='��' and e='˫' and f=true",conn,2,2
	do while not rs.bof and not rs.eof
		shuangyinliang=shuangyinliang+rs("d")
		shuangname=shuangname&rs("a")&" "
		conn.execute "update �û� set ���=���+"&rs("d")*2&",����=����+500,����=����+500 where ����='"& rs("a") &"'"
		conn.execute "update �û� set ���=���-"& rs("d") &" where ����='"& zhuangjia &"'"	
	rs.movenext
	loop
	rs.close
	'������жĲ�״̬
	conn.execute "update g set f=false where f=true and c='��' and e='˫'"
end if
if sjdubo/2<>int(sjdubo/2) then
	danshuang="<img src='../jhimg/dan.gif'>"
	rs.open "select * FROM g WHERE c='��' and e='��' and f=true",conn
	do while not rs.bof and not rs.eof
		danyinliang=danyinliang+rs("d")
		danname=danname&rs("e")&" "
		conn.execute "update �û� set ���=���+"&rs("d")*2&",����=����+500,����=����+500 where ����='"& rs("a") &"'"
		conn.execute "update �û� set ���=���-"& rs("d") &" where ����='"& zhuangjia &"'"	
	rs.movenext
	loop
	rs.close
	conn.execute "update g set f=false where f=true and c='��' and e='��'"
end if

'�Կ���С����
if sjdubo<=10 then
	daxiao="<img src='../jhimg/xiao.gif'>"
	rs.open "select * FROM g WHERE c='��' and e='С' and f=true",conn,2,2
	do while not rs.bof and not rs.eof
		xiaoyinliang=xiaoyinliang+rs("d")
		xiaoname=xiaoname&rs("a")&" "
		conn.execute "update �û� set ���=���+"&rs("d")*2&",����=����+500,����=����+500 where ����='"& rs("a") &"'"
		conn.execute "update �û� set ���=���-"& rs("d") &" where ����='"& zhuangjia &"'"
	rs.movenext	
	loop
	rs.close
	conn.execute "update g set f=false where f=true and c='��' and e='С'"
end if
if sjdubo>10 then
	daxiao="<img src='../jhimg/da.gif'>"
	rs.open "select * FROM g WHERE c='��' and e='��' and f=true",conn,2,2
	do while not rs.bof and not rs.eof
		dayinliang=dayinliang+rs("d")
		daname=daname&rs("a")&" "
		conn.execute "update �û� set ���=���+"&rs("d")*2&",����=����+500,����=����+500 where ����='"& rs("a") &"'"
		conn.execute "update �û� set ���=���-"& rs("d") &" where ����='"& zhuangjia &"'"	
	rs.movenext
	loop
	rs.close
	conn.execute "update g set f=false where f=true and c='��' and e='��'"
end if

'��ʣ������û�����
	rs.open "select * FROM g WHERE c='��' and d<>0 and b<>'ׯ��' and f=true",conn,2,2
	yingyin=0
	do while not rs.bof and not rs.eof
	yingyin=yingyin+rs("d")
	rs.movenext
	loop
	rs.close
	conn.execute "update �û� set ���=���+"&yingyin&" where ����='"& zhuangjia &"'"
	conn.execute "update g set f=false where f=true and c='��'"
	zong=yingyin+shuangyinliang+danyinliang+xiaoyinliang+dayinliang
	pei=shuangyinliang+danyinliang+xiaoyinliang+dayinliang
	duboname=shuangname&danname&xiaoname&daname
	xiazhu="[������]ׯ�ҿ���<img src='../jhimg/"&m1&".gif'><img src='../jhimg/"&m2&".gif'><img src='../jhimg/"&m3&".gif'>�ƣ�<font color=blue>"& sjdubo &"</font>��!"&danshuang&daxiao&"����ע��<font color=blue>"&zong&"</font>����ׯ�ң�["&zhuangjia&"]���룺<font color=blue>"&yingyin &"</font>��,�����<font color=blue>"&pei&"</font>�����ϼƣ�<font color=blue>"&yingyin-pei&"</font>��,���У�<font color=red>"&duboname&"</font>�������Ѻ��,����Ѻ�������������������+500��"
	set rs=nothing
	conn.close
	set conn=nothing
	exit function

end if
xiazhu="[������]##���Լ���С�ɰ����ó�:"& yinliang &"��������"& duboimg &"����һ���еģ�������ע���£�Ѻ��<font color=blue>"& dars &"</font>����ѺС:<font color=blue>"& xiaors &"</font>��,Ѻ����<font color=blue>"&danrs&"</font>��,Ѻ˫:<font color=blue>"&shuangrs&"</font>��������:<font color=blue>"&(10-durs)&"</font>�����֣�"
	set rs=nothing	
	conn.close
	set conn=nothing

'�Դ��Ĳ��ж�
case "��"
		randomize(timer())
		m2 = Int(6 * Rnd + 1)
		rs.open "select allvalue,��Ա�ȼ� FROM �û� WHERE ����='" & sjjh_name &"'",conn
		select case rs("��Ա�ȼ�")
		case 1
			hydian=31250
		case 2
			hydian=90000
		case 3
			hydian=250000
		case 4
			hydian=490000
		case else
			hydian=0
		end select
		if (rs("allvalue")-hydian)<yinliang then
			rs.close
			set rs=nothing	
			conn.close
			set conn=nothing
			Response.Write "<script language=JavaScript>{alert('������ô��Ĵ���𣿣�');}</script>"
			Response.End 		
		end if
		if sjjh_jhdj<10 then
			rs.close
			set rs=nothing	
			conn.close
			set conn=nothing
			Response.Write "<script language=JavaScript>{alert('��ʾ�������Ĳ������Ҫ10�����ϲſ��Բ�����');}</script>"
			Response.End 		
		end if

		rs.close
	rs.open "select top 1 a FROM g WHERE c='��' and b='ׯ��' and f=true",conn
	if rs.eof or rs.bof then
		rs.close
		set rs=nothing	
		conn.close
		set conn=nothing
		Response.Write "<script language=JavaScript>{alert('����[����]û��ׯ�ң��������߶Ĳ�����');}</script>"
		Response.End 		
	end if
	rs.close
	rs.open "select top 1 * FROM g WHERE c='��' and a='"& sjjh_name &"' and f=true",conn
	if not(rs.eof) or not(rs.bof) then
		if rs("b")="ׯ��" then
			temp="##��������[����]ׯ�ң���Ҫ��ʲôѽ!"
		else
			temp="##����[����]Ϊ��ѹ["&rs("e")&"]"&rs("d")&"���ſ��ְɣ�"
		end if
		rs.close
		set rs=nothing	
		conn.close
		set conn=nothing
		Response.Write "<script language=JavaScript>{alert('"&temp&"');}</script>"
		Response.End 
	end if
	rs.close
	rs.open "select top 1 id,a FROM g WHERE c='��' and b='���' and f=false",conn,2,2
	tempdu=0
	if rs.eof or rs.bof then
		rs.close
		tempdu=1
		Response.Write "<script language=JavaScript>{alert('��ʾ���ڵ��������û�п�λ���ڿ���!');}</script>"
	end if
	if tempdu=0 then
	id=rs("id")
	conn.execute "update �û� set allvalue=allvalue-"&yinliang&" where ����='" & sjjh_name &"'"
	conn.execute "update g set a='"&sjjh_name&"',f=true,d="&yinliang&",e='"&yacz&"' where id="&id&" and f=false"
	rs.close
	end if
	tmprs=conn.execute("Select count(*) As ���� from g where c='��' and e='��' and f=true")
	dars=tmprs("����")
	tmprs=conn.execute("Select count(*) As ���� from g where c='��' and e='С' and f=true")
	xiaors=tmprs("����")
	tmprs=conn.execute("Select count(*) As ���� from g where c='��' and e='��' and f=true")
	danrs=tmprs("����")
	tmprs=conn.execute("Select count(*) As ���� from g where c='��' and e='˫' and f=true")
	shuangrs=tmprs("����")
	tmprs=conn.execute("Select count(*) As ���� from g where c='��' and d<>0 and f=true")
	durs=tmprs("����")

'������
if durs>=10 then
	sj=DateDiff("s",Application("sjjh_du"),now())
	if sj<5 then
		s=5-sj
		Response.Write "<script language=JavaScript>{alert('��ʾ�����������������ݣ���������["&s&"����]�ٲ�����');}</script>"
		Response.End
	end if	
	Application.Lock
	Application("sjjh_du")=now()
	Application.UnLock
	randomize()
	m3 = Int(6 * Rnd + 1)
	sjdubo=m1+m2+m3
'����ׯ��
rs.open "select * FROM g WHERE c='��' and b='ׯ��' and f=true",conn,2,2
zhuangjia=rs("a")
rs.close

'���Ӵ���
if m1=m2 and m2=m3 and m3=m1 then
	rs.open "select * FROM g WHERE c='��' and d<>0 and b<>'ׯ��'",conn,2,2
	yingyin=0
	do while not rs.bof and not rs.eof
	yingyin=yingyin+rs("d")
	rs.movenext
	loop
	rs.close
	'��ׯ��Ǯ
	conn.execute "update �û� set allvalue=allvalue+"&yingyin&",����=����+10000 where ����='"& zhuangjia &"'"
	'������жĲ�״̬
	conn.execute "update g set f=false where f=true and c='��'"
	xiazhu="[����]ׯ�ҿ���<img src='../jhimg/"&m1&".gif'><img src='../jhimg/"&m2&".gif'><img src='../jhimg/"&m3&".gif'>�ƣ�"& sjdubo &"��!ׯ�ҿ������ӡ���ͨɱ��ׯ�ң�"&zhuangjia&"���룺"&yingyin&"��㣡���˵�ͷ����+1��㣡"
		set rs=nothing
		conn.close
		set conn=nothing
		exit function
end if

'��ʼ������
shuangyinliang=0
shuangname=""
danyinliang=0
danname=""
xiaoyinliang=0
xiaoname=""
dayinliang=0
daname=""

'����˫����
if sjdubo/2=int(sjdubo/2) then
	danshuang="<img src='../jhimg/shuang.gif'>"
	rs.open "select * FROM g WHERE c='��' and e='˫' and f=true",conn,2,2
	do while not rs.bof and not rs.eof
		shuangyinliang=shuangyinliang+rs("d")
		shuangname=shuangname&rs("a")&" "
		conn.execute "update �û� set allvalue=allvalue+"&rs("d")*2&",����=����+500,����=����+500 where ����='"& rs("a") &"'"
		conn.execute "update �û� set allvalue=allvalue-"& rs("d") &" where ����='"& zhuangjia &"'"	
	rs.movenext
	loop
	rs.close
	conn.execute "update g set f=false where f=true and c='��' and e='˫'"
end if
if sjdubo/2<>int(sjdubo/2) then
	danshuang="<img src='../jhimg/dan.gif'>"
	rs.open "select * FROM g WHERE c='��' and e='��' and f=true",conn,2,2
	do while not rs.bof and not rs.eof
		danyinliang=danyinliang+rs("d")
		danname=danname&rs("a")&" "
		conn.execute "update �û� set allvalue=allvalue+"&rs("d")*2&",����=����+500,����=����+500 where ����='"& rs("a") &"'"
		conn.execute "update �û� set allvalue=allvalue-"& rs("d") &" where ����='"& zhuangjia &"'"	
	rs.movenext
	loop
	rs.close
	conn.execute "update g set f=false where f=true and c='��' and e='��'"
end if


'�Կ���С����
if sjdubo<=10 then
	daxiao="<img src='../jhimg/xiao.gif'>"
	rs.open "select * FROM g WHERE c='��' and e='С' and f=true",conn,2,2
	do while not rs.bof and not rs.eof
		xiaoyinliang=xiaoyinliang+rs("d")
		xiaoname=xiaoname&rs("a")&" "
		conn.execute "update �û� set allvalue=allvalue+"&rs("d")*2&",����=����+500,����=����+500 where ����='"& rs("a") &"'"
		conn.execute "update �û� set allvalue=allvalue-"& rs("d") &" where ����='"& zhuangjia &"'"
	rs.movenext	
	loop
	rs.close
	conn.execute "update g set f=false where f=true and c='��' and e='С'"
end if
if sjdubo>10 then
	daxiao="<img src='../jhimg/da.gif'>"
	rs.open "select * FROM g WHERE c='��' and e='��' and f=true",conn,2,2
	do while not rs.bof and not rs.eof
		dayinliang=dayinliang+rs("d")
		daname=daname&rs("a")&" "
		conn.execute "update �û� set allvalue=allvalue+"&rs("d")*2&",����=����+500,����=����+500 where ����='"& rs("a") &"'"
		conn.execute "update �û� set allvalue=allvalue-"& rs("d") &" where ����='"& zhuangjia &"'"	
	rs.movenext
	loop
	rs.close
	conn.execute "update g set f=false where f=true and c='��' and e='��'"
end if

'��ʣ������û��㴦��
	rs.open "select * FROM g WHERE c='��' and d<>0 and b<>'ׯ��' and f=true",conn,2,2
	yingyin=0
	do while not rs.bof and not rs.eof
	yingyin=yingyin+rs("d")
	rs.movenext
	loop
	rs.close
	conn.execute "update �û� set allvalue=allvalue+"&yingyin&" where ����='"& zhuangjia &"'"
	conn.execute "update g set f=false where f=true and c='��'"
	zong=yingyin+shuangyinliang+danyinliang+xiaoyinliang+dayinliang
	pei=shuangyinliang+danyinliang+xiaoyinliang+dayinliang
	duboname=shuangname&danname&xiaoname&daname
	xiazhu="[����]ׯ�ҿ���<img src='../jhimg/"&m1&".gif'><img src='../jhimg/"&m2&".gif'><img src='../jhimg/"&m3&".gif'>�ƣ�<font color=blue>"& sjdubo &"</font>��!"&danshuang&daxiao&"����ע��<font color=blue>"&zong&"</font>�㣬ׯ�ң�["&zhuangjia&"]���룺<font color=blue>"&yingyin &"</font>���,�����<font color=blue>"&pei&"</font>��㣬�ϼƣ�<font color=blue>"&yingyin-pei&"</font>���,���У�<font color=red>"&duboname&"</font>�������Ѻ�У�����Ѻ�������������������+500�㣡"
	set rs=nothing
	conn.close
	set conn=nothing
	exit function

end if
xiazhu="[����]##���Լ�Ŭ���ݵĴ���ó�:"& yinliang &"�㣬��Ѻ"& duboimg &"����һ���еģ�������ע���£�Ѻ��<font color=blue>"& dars &"</font>����ѺС:<font color=blue>"& xiaors &"</font>��,Ѻ����<font color=blue>"&danrs&"</font>��,Ѻ˫:<font color=blue>"&shuangrs&"</font>��������:<font color=blue>"&(10-durs)&"</font>�����֣�"

	set rs=nothing	
	conn.close
	set conn=nothing

end select
end function
%>
