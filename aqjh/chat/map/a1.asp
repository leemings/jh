<!--#include file="../sjfunc/sjfunc.asp"-->
<!--#include file="../../mywp.asp"-->
<%'�ƶ�(����)
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if Instr(LCase(Application("aqjh_useronlinename"&nowinroom)),LCase(" "&aqjh_name&" "))=0 then 
	Session("aqjh_inthechat")="0" 
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
says=Replace(says,"&amp;","&")
if aqjh_grade<9 then
	says=bdsays(says)
end if
act=1
towhoway=0
i=instr(says,"$")
fnn1=mid(says,i+1)
says="<font color=green><font color=" & saycolor & ">"+walk(mid(says,i+1))+"</font>"
towhoway=1
towho=aqjh_name
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
'�ƶ�(����)
function walk(fn1)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")

rs.open "select ����,����,�书,����,�Ṧ,�ȼ� FROM �û� WHERE ����='" & aqjh_name &"'",conn
mapai=rs("����")
if rs("����")<500000 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ�����Ǯ����50�����Բ���ʹ�����߹��ܣ�');window.close();}</script>"
	response.end
end if
if rs("�书")<1000 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ������书����1000���Բ���ʹ�����߹��ܣ�');window.close();}</script>"
	response.end
end if
if rs("����")<10000 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ�������������1�����Բ���ʹ�����߹��ܣ�');window.close();}</script>"
	response.end
end if

if rs("�Ṧ")<1000 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ������Ṧ����1000������ȥ����Ṧ���������ɣ�');window.close();}</script>"
	response.end
end if

if len(fn1)>3 then
	Response.Write "<script language=JavaScript>{alert('��ʾ�����߷��������ѡ����ȷ·�� ');}</script>"
	Response.End 
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select ����ʱ�� from �û� where ����='"&aqjh_name&"'",conn
sj=DateDiff("n",rs("����ʱ��"),now())
if sj<1 then
	ss=1-sj
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('���������۲��۵�ѽ?�����"&ss&"���������߰ɣ�');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if

rs.close
conn.execute "update �û� set ����ʱ��=now()  where ����='"&aqjh_name&"'"
randomize 
r=int(rnd*32)+1
select case r
case 1
walk="<font color=red>��������Ϣ��</font><font color=#000080>λ��"& texin &""& mapai &"<font color=red>��##��</font> ��" & fn1 & "<img src='img/ki17.gif' width='20' height='20'> ���ߵ��˽��������,������ϰ�</font><font color=red>��½�˷硿˵��</font>:<font color=#000080>������ǿ��ֽ����������������ʲô���������в�֪��"& mapai &"<font color=red>��##�� </font>����Щʲô<font color=red>��##��</font>�������治֪����ʲô��������ϰ�<font color=#ff0000>��½�˷硿</font>��<font color=#ff0000>��##��</font>������,����һ�ѷɵ���<font color=#ff0000>��##��</font>����2����������.��Ϊ������ֻ�ð������� "
	rs.open "SELECT w4 FROM �û� WHERE ����='"&aqjh_name&"'",conn
	duyao=add(rs("w4"),"�ɵ�",1)
	conn.execute "update �û� set  �Ṧ=�Ṧ-1000,����=����-20000,w4='"&duyao&"' where ����='"&aqjh_name&"'"
	rs.close
case 2
walk="<font color=red>��������Ϣ��</font><font color=#000080>λ��"& texin &""& mapai &"<font color=red>��##��</font> ��" & fn1 & "<img src='img/ki17.gif' width='20' height='20'> ���ߵ��˿�������,�������Ĵ�����һ�·�����ȫ���˴���ֻ�ü�������</font><font color=red>��" & fn1 & "/ֱ�ϡ�</font> "
	conn.execute "update �û� set �Ṧ=�Ṧ-1000 where ����='"&aqjh_name&"'" 
case 3
walk="<font color=red>��������Ϣ��</font><font color=#000080>λ��"& texin &""& mapai &"<font color=red>��##��</font> ��" & fn1 & "<img src='img/ki17.gif' width='20' height='20'> ̽��̽·,ȴ��ȫû����·��ֻ�ü�������.<font color=#ff0000>��" & fn1 & "/���¡�</font> "
	conn.execute "update �û� set �Ṧ=�Ṧ-1000 where ����='"&aqjh_name&"'" 
case 4
walk="<font color=red>��������Ϣ��</font><font color=#000080>λ��"& texin &""& mapai &"<font color=red>��##��</font> ��" & fn1 & "<img src='img/ki17.gif' width='20' height='20'> ���ߵ��˷�����֮��,<font color=red>��##�� </font>��������,�뵽�����л����<font color=red>������ǰ����</font>һ��.<font color=#ff0000>��##��</font>����һ������ſ�ϧ��ȫ���˻�Ӧ<font color=#ff0000>��##��:˵��</font>����<font color=#ff0000>������ǰ����</font>���³����˰�ֻ�õ��´������ݺ�<font color=#ff0000>������ǰ����</font>��.<font color=#ff0000>��##��</font>ֻ�ü�������. "
	conn.execute "update �û� set �Ṧ=�Ṧ-1000 where ����='"&aqjh_name&"'" 
case 5
walk="<font color=red>��������Ϣ��</font><font color=#000080>λ��"& texin &""& mapai &"<font color=red>��##��</font> ��" & fn1 & "<img src='img/ki17.gif' width='20' height='20'> ���ߵ������ֲؾ���,�ؾ����ʦ</font><font color=red>���Ƴ���ʦ��˵��</font>:<font color=#000080>ʩ��������ɫ�����Ƿ�������.ϣ��ʩ������������������<font color=red>��##�� </font>˵��:��Ϊ�⼸����ڻ�ɽ����������������,��������ϣ��������ʦ�������.<font color=red>���Ƴ���ʦ��:˵��</font>�����˴ȱ�Ϊ��ʩ�����µ�Ȼ��������.<font color=#ff0000>���Ƴ���ʦ��</font>���̴�����2000��������<font color=#ff0000>��##��</font>****<font color=#ff0000>��##��</font>����������ȫ����.<font color=#ff0000>��##��</font>������һ����������. "
	conn.execute "update �û� set  �Ṧ=�Ṧ-1000,����=����+2000 where ����='"&aqjh_name&"'"
case 6
walk="<font color=red>��������Ϣ��</font><font color=#000080>λ��"& texin &""& mapai &"<font color=red>��##��</font> ��" & fn1 & "<img src='img/ki17.gif' width='20' height='20'> ̽��̽·,ȴ��ȫû����·��ֻ�ü�������.<font color=#ff0000>��" & fn1 & "/���¡�</font> "
	conn.execute "update �û� set �Ṧ=�Ṧ-1000 where ����='"&aqjh_name&"'" 
case 7
walk="<font color=red>��������Ϣ��</font><font color=#000080>λ��"& texin &""& mapai &"<font color=red>��##��</font> ��" & fn1 & "<img src='img/ki17.gif' width='20' height='20'> ���ߵ��˳����̵�,������ϰ���</font><font color=red>�����˵��</font>:<font color=#000080>������ǿ��ֽ����������ĳ����̵��"& mapai &"<font color=red>��##�� </font>����һ�뻹��û��ʲô�����ֻ�ü�������.����ǰ<font color=red>�����</font>��������õ���<font color=#ff0000>��##��</font>˳��˵��:�´λ�Ҫ�ǵ������ҹ������. "
	rs.open "SELECT w7 FROM �û� WHERE ����='"&aqjh_name&"'",conn
	duyao=add(rs("w7"),"õ��",3)
	conn.execute "update �û� set  �Ṧ=�Ṧ-1000,w7='"&duyao&"' where ����='"&aqjh_name&"'"
	rs.close
case 8
walk="<font color=red>��������Ϣ��</font><font color=#000080>λ��"& texin &""& mapai &"<font color=red>��##��</font> ��" & fn1 & "<img src='img/ki17.gif' width='20' height='20'> ̽��̽·,ȴ��ȫû����·��ֻ�ü�������.<font color=#ff0000>��" & fn1 & "/ֱ�¡�</font> "
	conn.execute "update �û� set �Ṧ=�Ṧ-1000 where ����='"&aqjh_name&"'" 
case 9
walk="<font color=red>��������Ϣ��</font><font color=#000080>λ��"& texin &""& mapai &"<font color=red>��##��</font> ��" & fn1 & "<img src='img/ki17.gif' width='20' height='20'> ̽��̽·,ȴ��ȫû����·��ֻ�ü�������.<font color=#ff0000>��" & fn1 & "/���ϡ�</font> "
	conn.execute "update �û� set �Ṧ=�Ṧ-1000 where ����='"&aqjh_name&"'" 
case 10
walk="<font color=red>��������Ϣ��</font><font color=#000080>λ��"& texin &""& mapai &"<font color=red>��##��</font> ��" & fn1 & "<img src='img/ki17.gif' width='20' height='20'> ���ߵ��˽���֮��,�����������ڱ�������.���úþ��ʾ���ǰ������.ȴ��С�ı����˻���һ��.</font><font color=red>��##��</font>:<font color=#000080>���������½�2000<font color=red>��##�� ˵��:</font>��������˭��˭����Ե�޹ʱ�����һ��.<font color=red>��##��</font>û��˼����ȥ��ֻ�ü�������. "
	conn.execute "update �û� set  �Ṧ=�Ṧ-1000,����=����-2000 where ����='"&aqjh_name&"'"
case 11
walk="<font color=red>��������Ϣ��</font><font color=#000080>λ��"& texin &""& mapai &"<font color=red>��##��</font> ��" & fn1 & "<img src='img/ki17.gif' width='20' height='20'> ̽��̽·,ȴ��ȫû����·��ֻ�ü�������.<font color=#ff0000>��" & fn1 & "/���¡�</font> "
	conn.execute "update �û� set �Ṧ=�Ṧ-1000 where ����='"&aqjh_name&"'" 
case 12
walk="<font color=red>��������Ϣ��</font><font color=#000080>λ��"& texin &""& mapai &"<font color=red>��##��</font> ��" & fn1 & "<img src='img/ki17.gif' width='20' height='20'> ̽��̽·,ȴ��ȫû����·��ֻ�ü�������.<font color=#ff0000>��" & fn1 & "/���¡�</font> "
	conn.execute "update �û� set �Ṧ=�Ṧ-1000 where ����='"&aqjh_name&"'" 
case 13
walk="<font color=red>��������Ϣ��</font><font color=#000080>λ��"& texin &""& mapai &"<font color=red>��##��</font> ��" & fn1 & "<img src='img/ki17.gif' width='20' height='20'> ���ߵ��˷�����֮��,<font color=red>��##�� </font>��������,�뵽�����л����<font color=red>������ǰ����</font>һ��.<font color=#ff0000>��##��</font>����һ�������<font color=#ff0000>������ǰ��������˵��:</font>��λ������֪���к���<font color=#ff0000>��##��</font>����˵��:������·���˵���˵˳��ݺ������˼�.Ҳϣ��<font color=#ff0000>������ǰ����</font>��ָ������һ�����书<font color=#ff0000>������ǰ����˵��:</font>������ô�г���ʹ��������书. <font color=#ff0000>��##��</font>�书��������1000. <font color=#ff0000>��##��</font>ѧ���˳���<font color=#ff0000>������ǰ����</font>�ݱ�ͼ�������."
	conn.execute "update �û� set �Ṧ=�Ṧ-1000,�书=�书+1000 where ����='"&aqjh_name&"'" 
case 14
walk="<font color=red>��������Ϣ��</font><font color=#000080>λ��"& texin &""& mapai &"<font color=red>��##��</font> ��" & fn1 & "<img src='img/ki17.gif' width='20' height='20'> ̽��̽·,ȴ��ȫû����·��ֻ�ü�������.<font color=#ff0000>��" & fn1 & "/���¡�</font> "
	conn.execute "update �û� set �Ṧ=�Ṧ-1000 where ����='"&aqjh_name&"'" 


case 15
walk="<font color=red>��������Ϣ��</font><font color=#000080>λ��"& texin &""& mapai &"<font color=red>��##��</font> ��" & fn1 & "<img src='img/ki17.gif' width='20' height='20'> �����ߵ��˻�ɽ����.<font color=#ff0000>��##��</font>˵��:�����澰�벻����ɽ������� <font color=#ff0000>��##��</font>.����һ�����ڴ������ǿ��������������ϻ������̼���.<font color=#ff0000>��##��</font>Ц������̫���˼쵽2�黢��."
	rs.open "SELECT w6 FROM �û� WHERE ����='"&aqjh_name&"'",conn
	duyao=add(rs("w6"),"����",2)
	conn.execute "update �û� set  �Ṧ=�Ṧ-1000,w6='"&duyao&"' where ����='"&aqjh_name&"'"
	rs.close

case 16
walk="<font color=red>��������Ϣ��</font><font color=#000080>λ��"& texin &""& mapai &"<font color=red>��##��</font> ��" & fn1 & "<img src='img/ki17.gif' width='20' height='20'> �����ߵ�������ҩ��.ҩ��<img src='../pic/dz45.gif' width='30' height='30'><font color=#ff0000>�����ϰ塿</font>���̳����к� <font color=#ff0000>��##��</font>.�͹ײ�֪�������ʲô����ʲôҩ���л��ܱ����.<font color=#ff0000>��##��</font>����һ��˵�����������Ż���¶��.<font color=#ff0000>�����ϰ塿</font>���̰ѾŻ���¶���ó���.�͹�200000���ѱ��˰�<font color=#ff0000>��##��</font>���˸���Ǯ�������ڶ�����."
	rs.open "SELECT w1 FROM �û� WHERE ����='"&aqjh_name&"'",conn
	duyao=add(rs("w1"),"�Ż���¶��",1)
	conn.execute "update �û� set  �Ṧ=�Ṧ-1000,����=����-200000,w1='"&duyao&"' where ����='"&aqjh_name&"'"
	rs.close

case 17
walk="<font color=red>��������Ϣ��</font><font color=#000080>λ��"& texin &""& mapai &"<font color=red>��##��</font> ��" & fn1 & "<img src='img/ki17.gif' width='20' height='20'> �����ߵ���ȪԴ����.�����ϰ�<font color=#ff0000>����������</font>���̳����к�<font color=#ff0000>��##��</font>�͹�.�͹׽���ϴ����ɵ���û��Ů�����.<font color=#ff0000>��##��</font>��˵�ðɷ����ü���ÿϴ���˽����������.ϴ����<font color=#ff0000>��##��</font>����Ǯ��<font color=#ff0000>����������</font>֮��ͼ�����·. "
	conn.execute "update �û� set �Ṧ=�Ṧ-1000,����=����-5000 where ����='"&aqjh_name&"'" 

case 18
walk="<font color=red>��������Ϣ��</font><font color=#000080>λ��"& texin &""& mapai &"<font color=red>��##��</font> ��" & fn1 & "<img src='img/ki17.gif' width='20' height='20'> ̽��̽·,ȴ��ȫû����·��ֻ�ü�������.<font color=#ff0000>��" & fn1 & "/���¡�</font> "
	conn.execute "update �û� set �Ṧ=�Ṧ-1000 where ����='"&aqjh_name&"'" 

case 19
walk="<font color=red>��������Ϣ��</font><font color=#000080>λ��"& texin &""& mapai &"<font color=red>��##��</font> ��" & fn1 & "<img src='img/ki17.gif' width='20' height='20'> �����ߵ��˿���ҩ��.ҩ��<img src='../images/picc/man5.gif' width='30' height='30'><font color=#ff0000>�����ϰ塿</font>���̳����к� <font color=#ff0000>��##��</font>.�͹ײ�֪�������ʲô����ʲôҩ���л��ܱ����.<font color=#ff0000>��##��</font>����һ��˵������������ײ�.<font color=#ff0000>�����ϰ塿</font>���̰Ѵ�ײ��ó���.�͹�200000���ѱ��˰�<font color=#ff0000>��##��</font>���˸���Ǯ�������ڶ�����."
	rs.open "SELECT w1 FROM �û� WHERE ����='"&aqjh_name&"'",conn
	duyao=add(rs("w1"),"��ײ�",1)
	conn.execute "update �û� set  �Ṧ=�Ṧ-1000,����=����-200000,w1='"&duyao&"' where ����='"&aqjh_name&"'"
	rs.close


case 20
walk="<font color=red>��������Ϣ��</font><font color=#000080>λ��"& texin &""& mapai &"<font color=red>��##��</font> ��" & fn1 & "<img src='img/ki17.gif' width='20' height='20'> ̽��̽·,ȴ��ȫû����·��ֻ�ü�������.<font color=#ff0000>��" & fn1 & "/���ϡ�</font> "
	conn.execute "update �û� set �Ṧ=�Ṧ-1000 where ����='"&aqjh_name&"'" 

case 21
walk="<font color=red>��������Ϣ��</font><font color=#000080>λ��"& texin &""& mapai &"<font color=red>��##��</font> ��" & fn1 & "<img src='img/ki17.gif' width='20' height='20'> �����ߵ��˻�����ջ.��ջ�ϰ�<font color=#ff0000>�����ϰ塿</font>���̳����к�<font color=#ff0000>��##��</font>�͹�.�͹���Ե�ʲô����Ķ������˿��ǻ�������������.<font color=#ff0000>��##��</font>˵��:�Ǿ������ţ�������������ͷ˳������һ̳Ů����<font color=#ff0000>�����ϰ塿</font>���̰ѾƲ˶˳���<font color=#ff0000>��##��</font>������˳�㸶��5000����<font color=#ff0000>�����ϰ塿</font>*******.<font color=#ff0000>�����ϰ塿</font>˵��:�´�Ҫ�ǵ���������<font color=#ff0000>��##��</font>��ǰ��������ͻȻ����2000.���Ǿ���ٱ� "
	conn.execute "update �û� set �Ṧ=�Ṧ-1000,����=����+2000,����=����-5000 where ����='"&aqjh_name&"'" 

case 22
walk="<font color=red>��������Ϣ��</font><font color=#000080>λ��"& texin &""& mapai &"<font color=red>��##��</font> ��" & fn1 & "<img src='img/ki17.gif' width='20' height='20'> ̽��̽·,ȴ��ȫû����·��ֻ�ü�������.<font color=#ff0000>��" & fn1 & "/���¡�</font> "
	conn.execute "update �û� set �Ṧ=�Ṧ-1000 where ����='"&aqjh_name&"'" 
	
case 23
walk="<font color=red>��������Ϣ��</font><font color=#000080>λ��"& texin &""& mapai &"<font color=red>��##��</font> ��" & fn1 & "<img src='img/ki17.gif' width='20' height='20'> ̽��̽·,ȴ��ȫû����·��ֻ�ü�������.<font color=#ff0000>��" & fn1 & "/����</font> "
	conn.execute "update �û� set �Ṧ=�Ṧ-1000 where ����='"&aqjh_name&"'" 	

case 24
walk="<font color=red>��������Ϣ��</font><font color=#000080>λ��"& texin &""& mapai &"<font color=red>��##��</font> ��" & fn1 & "<img src='img/ki17.gif' width='20' height='20'> �����ߵ��˴��Ǯׯ.Ǯׯ�ϰ�<font color=#ff0000>��Ǯ���</font>���̳����к�<font color=#ff0000>��##��</font>�͹�.�͹��������Ǯ��������Ϣ�ɸ���.<font color=#ff0000>��##��</font>˵��:���ڿ�����ô������Ǯ��<font color=#ff0000>��Ǯ�ϰ塿</font>����<font color=#ff0000>��##��</font>һ�¾������ܿ����ʹ�Ǯׯ�ó�һ����Ҹ�<font color=#ff0000>��##��</font>*******.<font color=#ff0000>��##��</font>����������æ��<font color=#ff0000>��Ǯ�ϰ塿</font>��л..л���������· "
	conn.execute "update �û� set �Ṧ=�Ṧ-1000,���=���+1 where ����='"&aqjh_name&"'" 
	
case 25
walk="<font color=red>��������Ϣ��</font><font color=#000080>λ��"& texin &""& mapai &"<font color=red>��##��</font> ��" & fn1 & "<img src='img/ki17.gif' width='20' height='20'> ̽��̽·,ȴ��ȫû����·��ֻ�ü�������.<font color=#ff0000>��" & fn1 & "/���ϡ�</font> "
	conn.execute "update �û� set �Ṧ=�Ṧ-1000 where ����='"&aqjh_name&"'" 

case 26
walk="<font color=red>��������Ϣ��</font><font color=#000080>λ��"& texin &""& mapai &"<font color=red>��##��</font> ��" & fn1 & "<img src='img/ki17.gif' width='20' height='20'> ̽��̽·,ȴ��ȫû����·��ֻ�ü�������.<font color=#ff0000>��" & fn1 & "/���ϡ�</font> "
	conn.execute "update �û� set �Ṧ=�Ṧ-1000 where ����='"&aqjh_name&"'" 
	
case 27
walk="<font color=red>��������Ϣ��</font><font color=#000080>λ��"& texin &""& mapai &"<font color=red>��##��</font> ��" & fn1 & "<img src='img/ki17.gif' width='20' height='20'> �����ߵ���������.�������ϰ�<font color=#ff0000>��ҩ���ɡ�</font>���̳����к� <font color=#ff0000>��##��</font>.��λ���Ͳ�֪������®��Ҫ��ʲô����ʲô��ҩ����.<font color=#ff0000>��##��</font>����һ��˵����������������.<font color=#ff0000>��ҩ���ɡ�</font>���̰Ѵ������ó���.�͹�200000���ѱ��˰�<font color=#ff0000>��##��</font>���˸���Ǯ�������ڶ�����."
	rs.open "SELECT w8 FROM �û� WHERE ����='"&aqjh_name&"'",conn
	duyao=add(rs("w8"),"������",1)
	conn.execute "update �û� set  �Ṧ=�Ṧ-1000,����=����-200000,w8='"&duyao&"' where ����='"&aqjh_name&"'"
	rs.close
			
case 28
walk="<font color=red>��������Ϣ��</font><font color=#000080>λ��"& texin &""& mapai &"<font color=red>��##��</font> ��" & fn1 & "<img src='img/ki17.gif' width='20' height='20'> ̽��̽·,ȴ��ȫû����·��ֻ�ü�������.<font color=#ff0000>��" & fn1 & "/������</font> "
	conn.execute "update �û� set �Ṧ=�Ṧ-1000 where ����='"&aqjh_name&"'" 
	
case 28
walk="<font color=red>��������Ϣ��</font><font color=#000080>λ��"& texin &""& mapai &"<font color=red>��##��</font> ��" & fn1 & "<img src='img/ki17.gif' width='20' height='20'> ̽��̽·,ȴ��ȫû����·��ֻ�ü�������.<font color=#ff0000>��" & fn1 & "/���ϡ�</font> "
	conn.execute "update �û� set �Ṧ=�Ṧ-1000 where ����='"&aqjh_name&"'" 

case 30
walk="<font color=red>��������Ϣ��</font><font color=#000080>λ��"& texin &""& mapai &"<font color=red>��##��</font> ��" & fn1 & "<img src='img/ki17.gif' width='20' height='20'> ���ߵ��˲ؾ���,����̫���˿�����͵ѧ���书.</font><font color=red>��##��</font>:<font color=#000080>���������׾��ʼϰ��<font color=red>��##�� </font>ϰ��һ��ʱ����ò���˿���������·.<font color=red>��##��</font>��ǰ�����书����1000�㷢������ѧ��ֹ��. "
	conn.execute "update �û� set  �Ṧ=�Ṧ-1000,�书=�书+1000 where ����='"&aqjh_name&"'"

case 31
walk="<font color=red>��������Ϣ��</font><font color=#000080>λ��"& texin &""& mapai &"<font color=red>��##��</font> ��" & fn1 & "<img src='img/ki17.gif' width='20' height='20'> ���ߵ��˽���֮��,�����������ڱ�������.���úþ��ʾ���ǰ������.ȴ��С�ı����˻���һ��.</font><font color=red>��##��</font>:<font color=#000080>���������½�2000<font color=red>��##�� ˵��:</font>��������˭��˭����Ե�޹ʱ�����һ��.<font color=red>��##��</font>û��˼����ȥ��ֻ�ü�������. "
	conn.execute "update �û� set  �Ṧ=�Ṧ-1000,����=����-2000 where ����='"&aqjh_name&"'"
				
case 32
walk="<font color=red>��������Ϣ��</font><font color=#000080>λ��"& texin &""& mapai &"<font color=red>��##��</font> ��" & fn1 & "<img src='img/ki17.gif' width='20' height='20'> ���ߵ��˽���֮��,�����������ڱ�������.���úþ��ʾ���ǰ������.ȴ��С�ı����˻���һ��.</font><font color=red>��##��</font>:<font color=#000080>���������½�2000<font color=red>��##�� ˵��:</font>��������˭��˭����Ե�޹ʱ�����һ��.<font color=red>��##��</font>û��˼����ȥ��ֻ�ü�������. "
	conn.execute "update �û� set  �Ṧ=�Ṧ-1000,����=����-2000 where ����='"&aqjh_name&"'"
	
case 33
walk="<font color=red>��������Ϣ��</font><font color=#000080>λ��"& texin &""& mapai &"<font color=red>��##��</font> ��" & fn1 & "<img src='img/ki17.gif' width='20' height='20'> �����ߵ��˿���ҩ��.ҩ��<img src='../images/picc/man5.gif' width='30' height='30'><font color=#ff0000>�����ϰ塿</font>���̳����к� <font color=#ff0000>��##��</font>.�͹ײ�֪�������ʲô����ʲôҩ���л��ܱ����.<font color=#ff0000>��##��</font>����һ��˵������������ײ�.<font color=#ff0000>�����ϰ塿</font>���̰Ѵ�ײ��ó���.�͹�200000���ѱ��˰�<font color=#ff0000>��##��</font>���˸���Ǯ�������ڶ�����."
	rs.open "SELECT w1 FROM �û� WHERE ����='"&aqjh_name&"'",conn
	duyao=add(rs("w1"),"��ײ�",1)
	conn.execute "update �û� set  �Ṧ=�Ṧ-1000,����=����-200000,w1='"&duyao&"' where ����='"&aqjh_name&"'"
	rs.close
end select
set rs=nothing	
conn.close
set conn=nothing
end function
%>