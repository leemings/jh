<%@ LANGUAGE=VBScript codepage ="936" %>

<!--#include file="../mywp.asp"-->
<%Response.Buffer=true
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.Expires = 0

sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
nowinroom=session("nowinroom")
if sjjh_name="" then Response.Redirect "../error.asp?id=440"
if Instr(LCase(Application("sjjh_useronlinename"&session("nowinroom")))," "&LCase(sjjh_name)&" ")=0 then
	Response.Write "<script language=JavaScript>{alert('��ʾ��������������Ҳſ��Բ�����');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End
end if
fromname=LCase(trim(request.querystring("fromname")))
toname=LCase(trim(request.querystring("toname")))
if toname<>sjjh_name then
	if fromname<>toname then
		Response.Write "<script language=JavaScript>{alert('������ʲô���˼�"&fromname&"Ҳ��������������˵��');}</script>"
		Response.End
	end if
		Response.Write "<script language=JavaScript>{alert('������"&fromname&" ���������Լ����������������������');}</script>"
		Response.End
	else
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "select �Ա� from �û� where ����='"&sjjh_name&"'",conn,3,3
If Rs.Bof OR Rs.Eof then
	rs.close	
	set rs=nothing	
	conn.close	
	set conn=nothing	
	Response.Write "<script language=JavaScript>{alert('��ǩ���˲����ڣ�������������Ϊ�����վ����ӳ��');}</script>"	'��4
	Response.End	
end if
rs.close
rs.open "select w9,w6,�Ա� FROM �û� WHERE ����='"&sjjh_name&"'",conn
tosex=rs("�Ա�")
randomize 
r=int(rnd*9)+1
select case r
case 1
Response.Write "<script language=JavaScript>{alert('ϣ����ֻǩ�ܹ�����������ˣ�');}</script>"
chice="����Եǩ����<font color=red>��"& sjjh_name &"��</font>�鵽һֻ��Եǩ�������������<font color=red>��"& fromname &"��</font>��ǩ˵��:����һ֧��Եǩ.ǩ����˵����꽫������Ե���Ծ����������Ե���.�����Ƿ������δ������һ��<font color=red>��"& sjjh_name &"��</font>˳��������һֻС����������Ե������.."
	   	 duyao=add(rs("w6"),"С����",1)
	     conn.execute "update �û� set  ����=����-20000,w6='"&duyao&"' where ����='"&sjjh_name&"'"
	     conn.execute "update �û� set  ����=����+20000 where ����='"&fromname&"'"

case 2
Response.Write "<script language=JavaScript>{alert('ϣ����ֻǩ�ܹ�����������ˣ�');}</script>"
chice="����ҵǩ����<font color=red>��"& sjjh_name &"��</font>�鵽һ֧��ҵǩ�������������<font color=red>��"& fromname &"��</font>��ǩ˵��:����һ֧��ҵǩ.��������ǩ.ǩ����˵�������˺�ͨ���Ჽ������,����˳����Ȼ����һ����˳<font color=red>��"& sjjh_name &"��</font>˳��������һ��ľͷף�㲽������.."
	 	 duyao=add(rs("w6"),"ľͷ",1)
	     conn.execute "update �û� set  ����=����-20000,w6='"&duyao&"' where ����='"&sjjh_name&"'"
	     conn.execute "update �û� set  ����=����+20000 where ����='"&fromname&"'"
	   
case 3
Response.Write "<script language=JavaScript>{alert('ϣ����ֻǩ�ܹ�����������ˣ�');}</script>"
chice="������ǩ����<font color=red>��"& sjjh_name &"��</font>�鵽һ֧����ǩ�������������<font color=red>��"& fromname &"��</font>��ǩ˵��:����һ֧����ǩ�൱�ĺ�.��ǩ������˵�㲻����ʲô���˶��������.���ǹ�ϲ���������ϵ�ǩ<font color=red>��"& sjjh_name &"��</font>˳��������20����.."
	   	conn.execute "update �û� set  ����=����+200000 where ����='"&sjjh_name&"'"
	    conn.execute "update �û� set  ����=����+20000 where ����='"&fromname&"'"


case 4
Response.Write "<script language=JavaScript>{alert('ϣ����ֻǩ�ܹ�����������ˣ�');}</script>"
chice="��ƽ��ǩ����<font color=red>��"& sjjh_name &"��</font>�鵽һ֧ƽ��ǩ�������������<font color=red>��"& fromname &"��</font>��ǩ˵��:����һ֧ƽ��ǩ.��������ǩ.ǩ����˵����그С��.������Ѫ��֮��.������ʦ���������ЩС��<font color=red>��"& sjjh_name &"��</font>������һ����������������ЩС��.."
	 	 duyao=add(rs("w6"),"������",1)
	     conn.execute "update �û� set  ����=����-20000,w6='"&duyao&"' where ����='"&sjjh_name&"'"
	     conn.execute "update �û� set  ����=����+20000 where ����='"&fromname&"'"
	

case 5
Response.Write "<script language=JavaScript>{alert('ϣ����ֻǩ�ܹ�����������ˣ�');}</script>"
chice="������ǩ����<font color=red>��"& sjjh_name &"��</font>�鵽һ֧����ǩ�������������<font color=red>��"& fromname &"��</font>��ǩ˵��:����һ֧����ǩ.��������ǩ.ǩ����˵���������Ƿ����Ҫ�ú�����.���ɹ��ݲ���.Ҫ��Ȼ������������ʧ<font color=red>��"& sjjh_name &"��</font>������2000�������������尲��.."
	 	 conn.execute "update �û� set  ����=����-20000,����=����+2000 where ����='"&sjjh_name&"'"
	     conn.execute "update �û� set  ����=����+20000 where ����='"&fromname&"'"
	
case 6
Response.Write "<script language=JavaScript>{alert('ϣ����ֻǩ�ܹ�����������ˣ�');}</script>"
chice="��ج��ǩ����<font color=red>��"& sjjh_name &"��</font>�鵽һ֧ج��ǩ�������������<font color=red>��"& fromname &"��</font>��ǩ˵��:����һ֧ج��ǩ.�ǳ��ǳ��Ĳ���.��������ج�������˱�����ʦֻ���������������Ϊ����ج��<font color=red>��"& sjjh_name &"��</font>��������2000.."
	 	 duyao=add(rs("w6"),"ľͷ",1)
	     conn.execute "update �û� set  ����=����-20000,����=����-2000 where ����='"&sjjh_name&"'"
	     conn.execute "update �û� set  ����=����+20000 where ����='"&fromname&"'"

case 7
Response.Write "<script language=JavaScript>{alert('ϣ����ֻǩ�ܹ�����������ˣ�');}</script>"
chice="����ɫǩ����<font color=red>��"& sjjh_name &"��</font>�鵽һ֧��ɫǩ�������������<font color=red>��"& fromname &"��</font>��ǩ˵��:����һ֧��ɫǩ.ǩ����˵�������ɫ֮ͽҲ������ǩ.���ؼ�ȥ<font color=red>��"& fromname &"��</font>˳������һȭ�����Ժ��ں�ɫ<font color=red>��"& sjjh_name &"��</font>�书������ʧ500"
	 	 conn.execute "update �û� set  ����=����-20000,�书=�书-500 where ����='"&sjjh_name&"'"
	     conn.execute "update �û� set  ����=����+20000 where ����='"&fromname&"'"
case 8
Response.Write "<script language=JavaScript>{alert('ϣ����ֻǩ�ܹ�����������ˣ�');}</script>"
chice="����ɫǩ����<font color=red>��"& sjjh_name &"��</font>�鵽һ֧��ɫǩ�������������<font color=red>��"& fromname &"��</font>��ǩ˵��:����һ֧��ɫǩ.ǩ����˵�������ɫ֮ͽҲ������ǩ.���ؼ�ȥ<font color=red>��"& fromname &"��</font>˳������һȭ�����Ժ��ں�ɫ<font color=red>��"& sjjh_name &"��</font>����������ʧ2000"
	 	 conn.execute "update �û� set  ����=����-20000,����=����-2000 where ����='"&sjjh_name&"'"
	     conn.execute "update �û� set  ����=����+20000 where ����='"&fromname&"'"
case 9
Response.Write "<script language=JavaScript>{alert('ϣ����ֻǩ�ܹ�����������ˣ�');}</script>"
chice="����Եǩ����<font color=red>��"& sjjh_name &"��</font>����һֻ��Եǩ�������������<font color=red>��"& fromname &"��</font>��ǩ˵��:�����Ե��δ��,����ȴ�.����ľ�����Ĳ�Ҫ����ǿ��.�ڵ���������Ŀ�е�Ů�ӽ������.˳����������2000��Ů�ӿ���������������."
	      conn.execute "update �û� set ����=����-20000 where ����='"&sjjh_name&"'"
	      conn.execute "update �û� set ����=����+20000 where ����='"&fromname&"'"
case 10
Response.Write "<script language=JavaScript>{alert('ϣ����ֻǩ�ܹ�����������ˣ�');}</script>"
chice="����Եǩ����<font color=red>��"& sjjh_name &"��</font>�鵽һֻ��Եǩ�������������<font color=red>��"& fromname &"��</font>��ǩ˵��:����һ֧��Եǩ.ǩ����˵����꽫������Ե���Ծ����������Ե���.�����Ƿ������δ������һ��<font color=red>��"& sjjh_name &"��</font>˳��������һֻС����������Ե������.."
	   	 duyao=add(rs("w6"),"С����",1)
	     conn.execute "update �û� set  ����=����-20000,w6='"&duyao&"' where ����='"&sjjh_name&"'"
	     conn.execute "update �û� set  ����=����+20000 where ����='"&fromname&"'"
	
end select
set rs=nothing
conn.close
set conn=nothing
Application.Lock
says="<font color=red><b>��������Ϣ��</b></font>"&chice	
says=replace(says,chr(39),"\'")
says=replace(says,chr(34),"\"&chr(34))
act="��Ϣ"
towhoway=0
towho="���"
addwordcolor="660099"
saycolor="008888"
addsays="��"
saystr="<script>parent.sh("& chr(39) & addwordcolor & chr(39) &","& chr(39) & saycolor & chr(39) &","& chr(39) & act & chr(39) &","& chr(39) & sjjh_name & chr(39) &","& chr(39) & addsays & chr(39) &","& chr(39) & towho & chr(39) &"," & chr(39) & says & chr(39) &"," & towhoway &  ","& nowinroom & ");<"&"/script>"
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

