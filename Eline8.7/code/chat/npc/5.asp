<%@ LANGUAGE=VBScript codepage ="936" %>

<!--#include file="../sjfunc/sjfunc.asp"-->
<!--#include file="../../mywp.asp"-->
<%'Ѱ������
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
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
says=Replace(says,"&amp;","&")
if sjjh_grade<9 then
	says=bdsays(says)
end if
act=1
towhoway=0
i=instr(says,"$")
fnn1=mid(says,i+1)
says="<font color=red>��Ѱ�����š�<font color=" & saycolor & ">"+seek(mid(says,i+1),towho)+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
'Ѱ������
function seek(fn1,to1)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "select �ȼ�,�Ṧ,����ʱ�� FROM �û� WHERE ����='" & sjjh_name &"'" ,conn,2,2
sj=DateDiff("n",rs("����ʱ��"),now())
if sj<(int(rnd*5)+1) then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	ss=(int(rnd*5)+1)-sj
	Response.Write "<script language=JavaScript>{alert('��ʾ���������["& ss &"]��,�ٲ�����');}</script>"
	Response.End
end if
dj=rs("�ȼ�")

light=rs("�Ṧ")
if rs("�Ṧ")<50000 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ������Ṧ�����޷�ʩչѽ������Ҳ��50000�㰡��');window.close();}</script>"
	response.end
end if
if dj<20 then
Response.Write "<script language=JavaScript>{alert('�˹�����Ҫ20��ս���ȼ�ѽ��');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if

rs.close
conn.execute "update �û� set �Ṧ=�Ṧ-50000,����ʱ��=now()  where ����='"&sjjh_name&"'"
randomize 
r=int(rnd*9)+1
select case r
case 1
seek=sjjh_name & "���������е���Ѱ���Ṧ����,ͻȻ�ڿ������з���һ����ʧ�Ѿõ�<img src='img/02.jpg' width='30' height='25'><font color=red>����ɽ��Ӱ��</font>����͵���Ȿ���Żؼ���ϰ."
	rs.open "SELECT w10 FROM �û� WHERE ����='"&sjjh_name&"'",conn
	duyao=add(rs("w10"),"��ɽ��Ӱ",1)
	conn.execute "update �û� set  w10='"&duyao&"' where ����='"&sjjh_name&"'"
	rs.close


case 2
seek=sjjh_name & "���������е���Ѱ���Ṧ����,���ٳ�δѰ�����ŵ�����ͻȻ" & to1 & "��������������һ��ʧ�����õ�<img src='img/02.jpg' width='30' height='25'><font color=red>�������Ӱ��</font>" & to1 & "�ʵ�˵����Ҫ���Ȿ�͸���."
	rs.open "SELECT w10 FROM �û� WHERE ����='"&sjjh_name&"'",conn
	duyao=add(rs("w10"),"�����Ӱ",1)
	conn.execute "update �û� set  w10='"&duyao&"' where ����='"&sjjh_name&"'"
	rs.close

	
case 3
seek=sjjh_name & "���������в��ϵ���Ѱ���Ṧ����,��û�뵽�㰮����������һ��<img src='img/02.jpg' width='30' height='25'><font color=red>������Ͷ�֡�</font>����������ѧ�����Ṧ�����ܾͲ�Ҫ����.��" & sjjh_name & "�İ�������ǰ����Ÿ���" & sjjh_name & "."
	rs.open "SELECT w10 FROM �û� WHERE ����='"&sjjh_name&"'",conn
	duyao=add(rs("w10"),"����Ͷ��",1)
	conn.execute "update �û� set  w10='"&duyao&"' where ����='"&sjjh_name&"'"
	rs.close
	
case 4
seek=sjjh_name & "���������в��ϵ���Ѱ���Ṧ����,���붼�벻���������Ϻ���Ѱ��<img src='img/02.jpg' width='30' height='25'><font color=red>���貨΢����</font>�������ܱ����ң��������."
	rs.open "SELECT w10 FROM �û� WHERE ����='"&sjjh_name&"'",conn
	duyao=add(rs("w10"),"�貨΢��",1)
	conn.execute "update �û� set  w10='"&duyao&"' where ����='"&sjjh_name&"'"
	rs.close
	
case 5
seek=sjjh_name & "���������в��ϵ���Ѱ���Ṧ����,��������·�������������ó���ʦ����<img src='img/02.jpg' width='30' height='25'><font color=red>����⻯�塿</font>������ؼ�ѧϰ.."
	rs.open "SELECT w10 FROM �û� WHERE ����='"&sjjh_name&"'",conn
	duyao=add(rs("w10"),"��⻯��",1)
	conn.execute "update �û� set  w10='"&duyao&"' where ����='"&sjjh_name&"'"
	rs.close

case 6
seek=sjjh_name & "�ÿ�ϧ����Ѱ���˽�����������Ҳû���ҵ�ʲô�Ṧ����!" 
	conn.execute "update �û� set �Ṧ=�Ṧ-100 where ����='" & sjjh_name &"'"

case 7
seek=sjjh_name & "�����ˣ���Ѱ���˽�����������û���ҵ�ʲô�Ṧ���ţ��Ṧ������100��!" 
	conn.execute "update �û� set �Ṧ=�Ṧ+500 where ����='" & sjjh_name &"'"
	
case 8
seek=sjjh_name & "<bgsound src=wav/m2.mid loop=1>���������в��ϵ���Ѱ���Ṧ����,����·��ΰС�����ſڿ���ΰС��.ΰС��˵��:��������̫��û��ѧ���аٱ���ϲ����ȥ��<img src='img/02.jpg' width='30' height='25'>" & sjjh_name & "��������."
	rs.open "SELECT w10 FROM �û� WHERE ����='"&sjjh_name&"'",conn
	duyao=add(rs("w10"),"���аٱ�",1)
	conn.execute "update �û� set  w10='"&duyao&"' where ����='"&sjjh_name&"'"
	rs.close


case 9
seek=sjjh_name & "<bgsound src=wav/m2.mid loop=1>���������в��ϵ���Ѱ���Ṧ����,����·����Ĺ����������С��Ů�����������������������Ṧ<img src='img/02.jpg' width='30' height='25'><font color=red>���������¡�</font>�����˸���.�����и�֮��.."
	rs.open "SELECT w10 FROM �û� WHERE ����='"&sjjh_name&"'",conn
	duyao=add(rs("w10"),"��������",1)
	conn.execute "update �û� set  w10='"&duyao&"' where ����='"&sjjh_name&"'"
	rs.close

case 10
seek=sjjh_name & "<bgsound src=wav/m2.mid loop=1>���������в��ϵ���Ѱ���Ṧ����,����·����Ĺ����������С��Ů�����������������������Ṧ<font color=red>���������¡�</font>�����˸���.�����и�֮��.."
	rs.open "SELECT w10 FROM �û� WHERE ����='"&sjjh_name&"'",conn
	duyao=add(rs("w10"),"��������",1)
	conn.execute "update �û� set  w10='"&duyao&"' where ����='"&sjjh_name&"'"
	rs.close

	
end select
set rs=nothing	
conn.close
set conn=nothing
end function
%>
