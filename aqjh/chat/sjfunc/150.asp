<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<!--#include file="func.asp"-->
<!--#include file="../../mywp.asp"-->
<%'ʹ��
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
%>
<!--#include file="pkif.asp"-->
<%
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
'�����뿪������Ѩ�ж�
call dianzan(towho)
if len(says)>150 then says=Left(says,150)
says=lcase(says)
says=Replace(says,"&amp;","&")
'��ϵͳ��ֹ�ַ�����
if aqjh_grade<9 then
says=bdsays(says)
end if
act=1
towhoway=0
i=instr(says,"$")
fnn1=mid(says,i+1)
says=jiuqingdao(mid(says,i+1),towho)
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)

'ʹ��
function jiuqingdao(fn1,to1)
fn1=trim(fn1)
if InStr(fn1,"'")<>0 or InStr(fn1,"`")<>0 or InStr(fn1,"=")<>0 or InStr(fn1,"-")<>0 or InStr(fn1,",")<>0 then 
	Response.Write "<script language=JavaScript>{alert('���ɣ�������ʲô���뵷���𣿣�');}</script>"
	Response.End 
end if
if aqjh_name=to1 and instr(";����ȭ;��ȡ��;������;������;��ڤ��;��������;ʥ����;",fn1)<>0 then
	Response.Write "<script language=javascript>alert('��"&fn1&"�������Լ����в�����');</script>"
	Response.End
end if
if to1="���" and instr("����ȭ;��ȡ��;������;������;��ڤ��;��������;ʥ����;",fn1)<>0 then
	Response.Write "<script language=javascript>alert('��"&fn1&"�����ܴ�ҽ��в�����');</script>"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select ����,�ȼ�,����,grade,����,����ʱ��,��¼ from �û� where ����='" & to1 &"'",conn,2,2
if rs("����")="����" and aqjh_grade<10 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ������Գ�������ʲô��KKK��');}</script>"
	Response.End
end if
dlsj=DateDiff("n",rs("��¼"),now())
if dlsj<3 and aqjh_grade<10 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('�㲻�ܶ������й����������߻�û��3���ӾͲ�Ҫ������ʹ�÷����ˣ�');}</script>"
	Response.End
end if
sj=DateDiff("n",rs("����ʱ��"),now())
if sj<15 and rs("����")="��" and aqjh_grade<10 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ�����ոձ�����ɱ���������ȷŹ����ɣ�');}</script>"
	Response.End
end if
if rs("grade")>=6 and aqjh_grade<10 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ���벻Ҫ�Թٸ����й�������');}</script>"
	Response.End
end if
if rs("�ȼ�")<=30 and rs("����")="��" and aqjh_grade<10 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ���벻Ҫ�Գ��뽭�����ֲ�����');}</script>"
	Response.End
end if
f=Minute(time())
if f<00 or f>25 then
	Response.Write "<script language=JavaScript>{alert('�����ж�������Ķ���ÿСʱ��ǰ25���ӣ�����������ʱ�䣡');window.close();}</script>"
	Response.End 
end if
if rs("����")=True and aqjh_grade<10 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ���Է��������������벻Ҫ͵Ϯ!');}</script>"
	Response.End
end if
rs.close
rs.open "select ����,w10,��¼,����,���� FROM �û� WHERE ����='" & aqjh_name &"'" ,conn,2,2
fla=rs("����")
if rs("����")<3000 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ����ķ��������޷�ʩչѽ������Ҳ��3000�㰡��');window.close();}</script>"
	response.end
end if
if rs("����")="����" and aqjh_grade<10 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ�����ǳ����˰���KKK��');}</script>"
	Response.End
end if
if rs("����")<10 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ��������������޷�ʩչѽ������Ҳ��10�㰡��');window.close();}</script>"
	response.end
end if
dlsj=DateDiff("n",rs("��¼"),now())
if dlsj<3 and aqjh_grade<10 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('������߲Ų���3���ӣ��Ͳ�Ҫ������ʹ�÷����ˣ�');}</script>"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
fn1=trim(fn1)
rs.open "select w10,����,���� from �û� where  ����='"&aqjh_name&"'",conn,3,3
mycard=abate(rs("w10"),fn1,1)
select case fn1
case "����ȭ"
        rs.close
        rs.open "SELECT * FROM �û� WHERE  ����='" & to1 & "'",conn
        if rs("����")="�ٸ�" then 
            jiuqingdao="<font color=green>������ȭ��<font color=" & saycolor & ">"&aqjh_name&",�㲻�ܶԹٸ���Աʹ������ȭ!"
           	rs.close
			set rs=nothing
			conn.close
	    	set conn=nothing
	        exit function
		end if
        if rs("����")="����ѵ��Ӫ" then 
            jiuqingdao="<font color=green>������ȭ��<font color=" & saycolor & ">"&aqjh_name&",�㲻�ܶԿ���ѵ��Ӫ����ʹ��!"
           	rs.close
			set rs=nothing
			conn.close
	    	set conn=nothing
	        exit function
		end if
		jiuqingdao="<bgsound src=wav/jiuzhao7.wav loop=1>������ȭ��<font color=" & saycolor & ">"&aqjh_name&"ʹ������ȭ<img src='img/41.gif'>��"&to1&"����!�������书������һ�룬�۵ò���ѽ"
        mzky=mzk(to1)
        if mzky="ok" then   
           conn.execute "update �û� set  ����=����-10,����=int(����/2),����=int(����/2),�书=int(�书/2),����=int(����/2) where ����='"&aqjh_name&"'"
           conn.execute "update �û� set ״̬='��',����=int(����/2),��¼=now()+3 where ����='" & to1 & "'"
            call boot(to1,to1&"��"&aqjh_name&"ʹ��������ȭ")  
        else
          jiuqingdao=jiuqingdao&mzky
        end if
        
case "��ȡ��"
        rs.close
        rs.open "SELECT * FROM �û� WHERE  ����='" & to1 & "'",conn
        money=int(rs("���")/20)
        if rs("����")="�ٸ�" then 
            jiuqingdao="<font color=green>����ȡ�<font color=" & saycolor & ">"&aqjh_name&",�㲻�ܶԹٸ���Աʹ�õ�ȡ��!"
           	rs.close
			set rs=nothing
			conn.close
	    	set conn=nothing
	        exit function
		end if
        if rs("����")="����ѵ��Ӫ" then 
            jiuqingdao="<font color=green>����ȡ�<font color=" & saycolor & ">"&aqjh_name&",�㲻�ܶԿ���ѵ��Ӫ����ʹ��!"
           	rs.close
			set rs=nothing
			conn.close
	    	set conn=nothing
	        exit function
		end if
        jiuqingdao="<bgsound src=wav/Bombs020.wav loop=1>����ȡ�<font color=" & saycolor & ">"&aqjh_name&"ʹ����ȡ���"&to1&"͵ȡ���"&money&"!"
        mhky=mhk(to1)
        if mhky="ok" then   
          conn.execute "update �û� set  ����=����-10,����=����+"&money&",����=����-3000 where ����='"&aqjh_name&"'"
           conn.execute "update �û� set ���=���-"&money&" where ����='" & to1 &"'"
        else
          jiuqingdao=jiuqingdao&mhky
        end if

case "������"
        rs.close
        rs.open "SELECT * FROM �û� WHERE  ����='" & to1 & "'",conn
            if rs("����")="�ٸ�" then 
            jiuqingdao="<font color=green>�������<font color=" & saycolor & ">"&aqjh_name&",�㲻�ܶԹٸ���Աʹ��������!"
           	rs.close
			set rs=nothing
			conn.close
	    	set conn=nothing
	        exit function
		end if
        if rs("����")="����ѵ��Ӫ" then 
            jiuqingdao="<font color=green>�������<font color=" & saycolor & ">"&aqjh_name&",�㲻�ܶԿ���ѵ��Ӫ����ʹ��!"
           	rs.close
			set rs=nothing
			conn.close
	    	set conn=nothing
	        exit function
		end if
        jiuqingdao="<bgsound src=wav/Bombs020.wav loop=1>�������<font color=" & saycolor & ">"&aqjh_name&"ʹ����������"&to1&"����˯��!"
        mgky=mgk(to1)
        if mgky="ok" then   
          conn.execute "update �û� set  ����=����-10,����=����-3000 where ����='"&aqjh_name&"'"
          conn.execute "update �û� set ״̬='��',��¼=now()+1/144,�¼�ԭ��='"&aqjh_name&" ��:������' where ����='" & to1 &"'"
          call boot(to1,"�ߣ������ߣ�"&aqjh_name&","&fn1)
        else
          jiuqingdao=jiuqingdao&mgky
        end if

case "������"
        rs.close
        rs.open "SELECT * FROM �û� WHERE  ����='" & to1 & "'",conn
        if rs("����")="�ٸ�" then 
            jiuqingdao="<font color=green>�������񹦡�<font color=" & saycolor & ">"&aqjh_name&",�㲻�ܶԹٸ���Աʹ�þ�����!"
           	rs.close
			set rs=nothing
			conn.close
	    	set conn=nothing
	        exit function
		end if
        if rs("����")="����ѵ��Ӫ" then 
            jiuqingdao="<font color=green>�������񹦡�<font color=" & saycolor & ">"&aqjh_name&",�㲻�ܶԿ���ѵ��Ӫ����ʹ��!"
           	rs.close
			set rs=nothing
			conn.close
	    	set conn=nothing
	        exit function
		end if
	    jiuqingdao="<bgsound src=wav/jiuzhao7.wav loop=1>�������񹦡�<font color=" & saycolor & ">"&aqjh_name&"ʹ��������<img src='img/41.gif'>��"&to1&"�����ҹ���!��������������3000"
        mfky=mfk(to1)
        if mfky="ok" then   
           conn.execute "update �û� set  ����=����-10,����=����-3000,����=����-3000 where ����='"&aqjh_name&"'"
           conn.execute "update �û� set ����=int(����/2),�书=int(�书/2),����=int(����/2) where ����='" & to1 & "'"
           call boot(to1,czsj)
  
        else
          jiuqingdao=jiuqingdao&mfky
        end if
        
case "ʥ����"
        rs.close
        rs.open "SELECT * FROM �û� WHERE  ����='" & to1 & "'",conn
        if rs("����")="�ٸ�" then 
            jiuqingdao="<font color=green>��ʥ���<font color=" & saycolor & ">"&aqjh_name&",�㲻�ܶԹٸ���Աʹ��ʥ����!"
           	rs.close
			set rs=nothing
			conn.close
	    	set conn=nothing
	        exit function
		end if
        if rs("����")="����ѵ��Ӫ" then 
            jiuqingdao="<font color=green>��ʥ���<font color=" & saycolor & ">"&aqjh_name&",�㲻�ܶԿ���ѵ��Ӫ����ʹ��!"
           	rs.close
			set rs=nothing
			conn.close
	    	set conn=nothing
	        exit function
		end if
	    jiuqingdao="<bgsound src=wav/jiuzhao7.wav loop=1>��ʥ���<font color=" & saycolor & ">"&aqjh_name&"ʹ��ʥ����<img src='img/41.gif'>��"&to1&"�����ҹ�������"&to1&"��ɽ�ʬ��!"
        mdky=mdk(to1)
        if mdky="ok" then   
           conn.execute "update �û� set  ����=����-10,����=����-3000 where ����='"&aqjh_name&"'"
           conn.execute "update �û� set ״̬='��ʬ��',���1='��ʬ��',��¼=now()+3 where ����='" & to1 & "'"
           call boot(to1,czsj)
  
        else
          jiuqingdao=jiuqingdao&mdky
        end if

case "��������"
        rs.close
        rs.open "SELECT * FROM �û� WHERE  ����='" & to1 & "'",conn
        if rs("����")="�ٸ�" then 
            jiuqingdao="<font color=green>������������<font color=" & saycolor & ">"&aqjh_name&",�㲻�ܶԹٸ���Աʹ����������!"
           	rs.close
			set rs=nothing
			conn.close
	    	set conn=nothing
	        exit function
		end if
        if rs("����")="����ѵ��Ӫ" then 
            jiuqingdao="<font color=green>������������<font color=" & saycolor & ">"&aqjh_name&",�㲻�ܶԿ���ѵ��Ӫ����ʹ��!"
           	rs.close
			set rs=nothing
			conn.close
	    	set conn=nothing
	        exit function
		end if
	    jiuqingdao="<bgsound src=wav/jiuzhao7.wav loop=1>������������<font color=" & saycolor & ">"&aqjh_name&"ʹ����������<img src='img/41.gif'>��"&to1&"�����ҹ���"&to1&"һ�ʽ�������!�����Լ�ȷ����10000�Ļ��֣��ݵ�Ҳ�����ˣ�"
        mfky=mfk(to1)
        if mfky="ok" then   
           conn.execute "update �û� set  ����=����-10,allvalue=allvalue-10000 where ����='"&aqjh_name&"'"
           conn.execute "update �û� set ״̬='��',��¼=now()+1/144,�¼�ԭ��='"&aqjh_name&" ����:"&fn1&"' where ����='" & to1 & "'"
           call boot(to1,"���Σ������ߣ�"&aqjh_name&","&fn1)
        else
          jiuqingdao=jiuqingdao&mfky
        end if
        

case "��ڤ��"
        rs.close
        rs.open "SELECT * FROM �û� WHERE  ����='" & to1 & "'",conn
        money=int(rs("����")/5)
        if rs("����")="�ٸ�" then 
            jiuqingdao="<font color=green>����ڤ����<font color=" & saycolor & ">"&aqjh_name&",�㲻�ܶԹٸ���Աʹ�õ�ȡ��!"
           	rs.close
			set rs=nothing
			conn.close
	    	set conn=nothing
	        exit function
		end if
        if rs("����")="����ѵ��Ӫ" then 
            jiuqingdao="<font color=green>����ڤ����<font color=" & saycolor & ">"&aqjh_name&",�㲻�ܶԿ���ѵ��Ӫ����ʹ��!"
           	rs.close
			set rs=nothing
			conn.close
	    	set conn=nothing
	        exit function
		end if
	    jiuqingdao="<bgsound src=wav/Bombs020.wav loop=1>����ڤ����<font color=" & saycolor & ">"&aqjh_name&"ʹ����ڤ����"&to1&"һֱ��ȡ����.��������"&to1&"�;��ƿ�!"
        meky=mek(to1)
        if meky="ok" then   
          conn.execute "update �û� set  ����=����-10,����=����+"&money&",����=����-3000 where ����='"&aqjh_name&"'"
           conn.execute "update �û� set ����=����-"&money&" where ����='" & to1 &"'"
        else
          jiuqingdao=jiuqingdao&meky
        end if
        
  
case else

	Response.Write "<script language=JavaScript>{alert('�㲢û��["&fn1&"]���ַ���,�����ٴ�ʹ��,����ȥѰ������');}</script>"
	Response.End
end select

'ɾ���Լ���������¼
conn.execute "update �û� set w10='"&mycard&"' where  ����='"&aqjh_name&"'"
conn.execute "insert into l(a,b,c,d,e) values (now(),'"& aqjh_name &"','"& to1 &"','����','"& fn1 & "')"
set rs=nothing	
conn.close
set conn=nothing
end function
'��ɽ��Ӱ
function mzk(to1)
  Set conn=Server.CreateObject("ADODB.CONNECTION")
  Set rs=Server.CreateObject("ADODB.RecordSet")
  conn.open Application("aqjh_usermdb")
  rs.open "select w10 from �û� where ����='"&to1&"'",conn,3,3
	if iswp(rs("w10"),"��ɽ��Ӱ")=0 then
		rs.close
	    mzk="ok"
	    exit function
	else
		tocard=abate(rs("w10"),"��ɽ��Ӱ",1)
		conn.execute "update �û� set w10='"&tocard&"' where  ����='"&to1&"'"
	   'mzk="<font color=green>����ɽ��Ӱ��</font>"&to1&"���������ŷ�ɽ��Ӱ,��˲��ܶ���������ȭ!"
	   mzk="<font color=green>"&to1&"ʹ������ɽ��Ӱ��<font color=" & saycolor & ">"&"���ߺٺ�������û��ô���׹���!</font>"
	end if
	  rs.close
	  conn.close
	  set rs=nothing
	  set conn=nothing
end function

'�貨΢��
function mhk(to1)
  Set conn=Server.CreateObject("ADODB.CONNECTION")
  Set rs=Server.CreateObject("ADODB.RecordSet")
  conn.open Application("aqjh_usermdb")
  rs.open "select w10 from �û� where ����='"&to1&"'",conn,3,3
	if iswp(rs("w10"),"�貨΢��")=0 then
		rs.close
	    mhk="ok"
	    exit function
	else
		tocard=abate(rs("w10"),"�貨΢��",1)
		conn.execute "update �û� set w10='"&tocard&"' where  ����='"&to1&"'"
	   'mhk="<font color=green>���貨΢����</font>"&to1&"�����������貨΢��,��˲��ܶ����õ�ȡ��!"
	   mhk="��û͵��"&to1&"ʹ�����貨΢����<font color=" & saycolor & ">"&"���ߺٺ���͵�ҵ�Ǯ�ҿ��붼��Ҫ�������ܹ�������͵Ǯ��!</font>"
	end if
	  rs.close
	  conn.close
	  set rs=nothing
	  set conn=nothing
end function

'����Ͷ��
function mgk(to1)
  Set conn=Server.CreateObject("ADODB.CONNECTION")
  Set rs=Server.CreateObject("ADODB.RecordSet")
  conn.open Application("aqjh_usermdb")
  rs.open "select w10 from �û� where ����='"&to1&"'",conn,3,3
	if iswp(rs("w10"),"����Ͷ��")=0 then
		rs.close
	    mgk="ok"
	    exit function
	else
		tocard=abate(rs("w10"),"����Ͷ��",1)
		conn.execute "update �û� set w10='"&tocard&"' where  ����='"&to1&"'"
	   'mgk="<font color=green>������Ͷ�֡�</font>"&to1&"����������������,��˲��ܶ�����������!"
	   mgk="��û��ʹչ"&to1&"ʹ��������Ͷ�֡�<font color=" & saycolor & ">"&"���ߺٺ�������˯�����Ҿ��ǲ���˯.����������ô��!</font>"
	end if
	  rs.close
	  conn.close
	  set rs=nothing
	  set conn=nothing
end function

'��⻯��
function mfk(to1)
  Set conn=Server.CreateObject("ADODB.CONNECTION")
  Set rs=Server.CreateObject("ADODB.RecordSet")
  conn.open Application("aqjh_usermdb")
  rs.open "select w10 from �û� where ����='"&to1&"'",conn,3,3
	if iswp(rs("w10"),"��⻯��")=0 then
		rs.close
	    mfk="ok"
	    exit function
	else
		tocard=abate(rs("w10"),"��⻯��",1)
		conn.execute "update �û� set w10='"&tocard&"' where  ����='"&to1&"'"
	   'mfk="<font color=green>����⻯�塿</font>"&to1&"���������ŷ�⻯��,��˲��ܶ����þ�����!"
	   mfk="���޷�����"&to1&"ʹ��<font color=red>����⻯�塿</font><font color=" & saycolor & ">"&"���ߺٺ�����ҸϾ�ɱ��û��ô����.�����з�⻯�����!</font>"
	end if
	  rs.close
	  conn.close
	  set rs=nothing
	  set conn=nothing
end function

'���аٱ�
function mek(to1)
  Set conn=Server.CreateObject("ADODB.CONNECTION")
  Set rs=Server.CreateObject("ADODB.RecordSet")
  conn.open Application("aqjh_usermdb")
  rs.open "select w10 from �û� where ����='"&to1&"'",conn,3,3
	if iswp(rs("w10"),"���аٱ�")=0 then
		rs.close
	    mek="ok"
	    exit function
	else
		tocard=abate(rs("w10"),"���аٱ�",1)
		conn.execute "update �û� set w10='"&tocard&"' where  ����='"&to1&"'"
	   'mek="<font color=green>�����аٱ䡿</font>"&to1&"���������ŷ�⻯��,��˲��ܶ����þ�����!"
	   mek="���޷�ʹչ��ڤ��"&to1&"ʹ��<font color=red>�����аٱ䡿</font><font color=" & saycolor & ">"&"���ߺٺ��������������û��ô����.������ΰС���͵����аٱ����߹������һ���Ĺ�.!</font>"
	end if
	  rs.close
	  conn.close
	  set rs=nothing
	  set conn=nothing
end function

'��������
function mdk(to1)
  Set conn=Server.CreateObject("ADODB.CONNECTION")
  Set rs=Server.CreateObject("ADODB.RecordSet")
  conn.open Application("aqjh_usermdb")
  rs.open "select w10 from �û� where ����='"&to1&"'",conn,3,3
	if iswp(rs("w10"),"��������")=0 then
		rs.close
	    mdk="ok"
	    exit function
	else
		tocard=abate(rs("w10"),"��������",1)
		conn.execute "update �û� set w10='"&tocard&"' where  ����='"&to1&"'"
	   'mdk="<font color=green>���������¡�</font>"&to1&"���������Ŷ�������,��˲��ܶ�����ʥ����!"
	   mdk="���޷�ʹչʥ����"&to1&"ʹ��<font color=red>���������¡�</font><font color=" & saycolor & ">"&"���ߺٺ�����ұ�ɽ�ʬ��û��ô����.������С��Ů�͵Ķ����������߹������һ���Ĺ�.!</font>"
	end if
	  rs.close
	  conn.close
	  set rs=nothing
	  set conn=nothing
end function

'�����Ӱ
function mfk(to1)
  Set conn=Server.CreateObject("ADODB.CONNECTION")
  Set rs=Server.CreateObject("ADODB.RecordSet")
  conn.open Application("aqjh_usermdb")
  rs.open "select w10 from �û� where ����='"&to1&"'",conn,3,3
	if iswp(rs("w10"),"�����Ӱ")=0 then
		rs.close
	    mfk="ok"
	    exit function
	else
		tocard=abate(rs("w10"),"�����Ӱ",1)
		conn.execute "update �û� set w10='"&tocard&"' where  ����='"&to1&"'"
	   'mfk="<font color=green>������������</font>"&to1&"���������������Ӱ,��˲��ܶ�������������!"
	   mfk="���޷�ʹչ��������"&to1&"ʹ��<font color=red>�������Ӱ��</font><font color=" & saycolor & ">"&"���ߺٺ�����������û��ô����.�����������Ӱ���߹������һ���Ĺ�.!</font>"
	end if
	  rs.close
	  conn.close
	  set rs=nothing
	  set conn=nothing
end function
%>�����Ӱ��</font><font color=" & saycolor & ">"&"���ߺٺ�����������û��ô����.�����������Ӱ���߹������һ���Ĺ�.!</font>"
	end if
	  rs.close
	  conn.close
	  set rs=nothing
	  set conn=nothing
end function
%>