<!--#include file="conn.asp"-->
<!-- #include file="inc/const.asp" -->
<!--#include file="Z_bankconn.asp"-->
<!--#include file="Z_bankconfig.asp"-->
<%
dim addtili,addjinyan,addfatie,addpower,fagonggao,faduangxun,kanip,delmail
dim menu
dim fuwu_setting         '���������     ��������    2002.11.30

stats="�������� �������"
call nav()
call head_var(0,0,"��������","Z_bank.asp")

if not founduser then
	Errmsg=Errmsg+"<br>"+"<li>��û�н������������Ȩ�ޣ����ȵ�¼����ͬ����Ա��ϵ��"
	call dvbbs_error()
elseif cint(bank_setting(6))=1 and (not master) and (DatePart("h", time)<cint(Chen_BusinessTimeSlice(0)) or DatePart("h", time)>=cint(Chen_BusinessTimeSlice(1))) then '�����Ƿ����ö�ʱӪҵ
	call BankClose()	
else
    set rs=server.createobject("adodb.recordset")
    sql="select * from [fuwu]"
    rs.open sql,conn1,1,3
    addtili=rs("addtili") 			'���������ļ۸�
    addjinyan=rs("addjinyan")		'������ļ۸�
    addfatie=rs("addfatie")			'���������ļ۸�
    addpower=rs("addpower")			'���������ļ۸�
    fagonggao=rs("fagonggao")		'������ļ۸�
    faduangxun=rs("faduangxun")		'������Ϣ�ļ۸�
    kanip=rs("kanip")				'��IP�ļ۸�
	delmail=rs("delmail")			'ɾ��1�ʼ��Ļ������𵥼�   ��������� 2002.11.30
	fuwu_setting=split(rs("fuwu_setting"),",")        '���������     �Ƿ�ͨ��������    2002.11.30
	                                                  'canCP/canEP/canPower/canfatie/canGonggao/canDuangxun/canKanip/canDelmail  1-�� 0����
    rs.close

	menu=request.Form("menu")	
%>
	<table cellpadding=3 cellspacing=1 align=center class=tableborder1>
	<tr><th valign=middle colspan=2 align=center height=25>��������  �������</th></tr>
	<tr><td valign=middle class=tablebody2 height=100>
<%
	call bankhead()
	if request("action")="admin" then
		call admin()		'��������ҳ��
	elseif menu="" then
		call main()       	'����ҳ��
	elseif menu=1 then
		call jy()			'������ֵ
	elseif menu=2 then
		call ml()			'��������ֵ
	elseif menu=3 then
		call ft()			'��������
	elseif menu=4 then
		call power()		'��������     
	elseif menu=5 then
		call fgg()			'������		
	elseif menu=6 then
		call fdx()			'Ⱥ������
	elseif menu=7 then
		call kip()			'���ԱIP
	elseif menu=8 then
		call syj()			'ɾ���ʼ�
	elseif menu=10 then
		call admin1()		'�����������
	end if
   
	response.write "</td></tr></table>"
end if

call activeonline()
call footer()

'----------------------------------�޸ļ۸�-----------------------------------
sub admin1()
	if not master then
		Errmsg=Errmsg+"<br>"+"<li>��û�н���������Ĺ����Ȩ��!"
		call fuwu_err()
	else
		if request.form("tili")="" or (not isnumeric(request.form("tili"))) then
			founderr=true
			Errmsg=Errmsg+"<br>"+"<li>�����۸����Ϊ����"
		elseif request.form("tili")<0 then
			founderr=true
			Errmsg=Errmsg+"<br>"+"<li>�����۸���Ϊ����"
		end if	 
		if request.form("jinyan")="" or (not isnumeric(request.form("jinyan"))) then
			founderr=true
			Errmsg=Errmsg+"<br>"+"<li>����۸����Ϊ����" 
		elseif request.form("jinyan")<0 then
			founderr=true
			Errmsg=Errmsg+"<br>"+"<li>����۸���Ϊ����"		
		end if	
		if request.form("fatie")="" or (not isnumeric(request.form("fatie"))) then
			founderr=true
			Errmsg=Errmsg+"<br>"+"<li>�������۸����Ϊ����" 
		elseif request.form("fatie")<0 then
			founderr=true
			Errmsg=Errmsg+"<br>"+"<li>�������۸���Ϊ����"			
		end if	
		if request.form("power")="" or (not isnumeric(request.form("power"))) then			
			founderr=true
			Errmsg=Errmsg+"<br>"+"<li>�����۸����Ϊ����" 	
		elseif request.form("power")<0 then
			founderr=true
			Errmsg=Errmsg+"<br>"+"<li>�����۸���Ϊ����"		
		end if
		if request.form("gonggao")="" or (not isnumeric(request.form("gonggao"))) then			
			founderr=true
			Errmsg=Errmsg+"<br>"+"<li>����۸����Ϊ����" 	
		elseif request.form("gonggao")<=0 then
			founderr=true
			Errmsg=Errmsg+"<br>"+"<li>����۸���Ϊ����"		
		end if
		if request.form("duangxun")="" or (not isnumeric(request.form("duangxun"))) then
			founderr=true
			Errmsg=Errmsg+"<br>"+"<li>Ⱥ������Ϣ�۸����Ϊ����"
		elseif request.form("duangxun")<0 then
			founderr=true
			Errmsg=Errmsg+"<br>"+"<li>Ⱥ������Ϣ�۸���Ϊ����"		
		end if
		if request.form("kanip")="" or (not isnumeric(request.form("kanip"))) then
			founderr=true
			Errmsg=Errmsg+"<br>"+"<li>�����߻�ԱIP��Ϣ�۸����Ϊ����"
		elseif request.form("kanip")<0 then
			founderr=true
			Errmsg=Errmsg+"<br>"+"<li>�����߻�ԱIP��Ϣ�۸���Ϊ����"		
		end if
		if request.form("delmail")="" or (not isnumeric(request.form("delmail"))) then
			founderr=true
			Errmsg=Errmsg+"<br>"+"<li>ɾ���ʼ����۱���Ϊ����"
		elseif request.form("delmail")<0 then
			founderr=true
			Errmsg=Errmsg+"<br>"+"<li>ɾ���ʼ����۲���Ϊ����"		
		end if				
	end if	
	if founderr then      
		call fuwu_err()			
	else
		set rs=server.createobject("adodb.recordset")
		sql="select * from [fuwu]"
		rs.open sql,conn1,1,3
		rs("addtili")=request.form("tili")
		rs("addjinyan")=request.form("jinyan")
		rs("addfatie")=request.form("fatie")
		rs("addpower")=request.form("power")
		rs("fagonggao")=request.form("gonggao")
		rs("faduangxun")=request.form("duangxun")
		rs("kanip")=request.form("kanip")
		rs("delmail")=request.form("delmail")         
		
		'canCP/canEP/canPower/canfatie/canGonggao/canDuangxun/canKanip/canDelmail  1-�� 0����
		rs("fuwu_setting")=request.Form("canCP") & "," & request.Form("canEP")& "," & request.Form("canPower")& "," & request.Form("canfatie")& "," & request.Form("canGonggao")& "," & request.Form("canDuangxun")& "," & request.Form("canKanip")& "," & request.Form("canDelmail")
		
		rs.update
		rs.close

		if cint(log_setting(0))=1 and cint(log_setting(3))=1 then
			content="��������۸�"
			call logs("����","�����������",membername)
			sucmsg=sucmsg+"<br>"+"<li>���Ĳ�����Ϣ�Ѿ���¼�ڰ�"
		end if
		sucmsg=sucmsg+"<br>"+"<li>������ɣ��뷵��!"		
		call fuwu_suc()
	end if	
end sub

'-------------------------------ɾ���ʼ�-----------------------------------
sub syj()
	if cint(fuwu_setting(7))=0 then 
		founderr=true
		Errmsg=Errmsg+"<br>"+"<li>�÷����Ѿ���ͣ���������Ա��ϵ"
		call fuwu_err()
		exit sub
	end if
	founderr=false
	set rs=server.createobject("adodb.recordset")
	sql="select * from [message] where flag=1 and issend=1 and incept='"&membername&"' "
	rs.open sql,conn,1,3
	
	if not (rs.eof and rs.bof) then       '�Ƿ����Ѷ��Ķ���Ϣ
		rs.close
		dim cunm    '��������� ͳ�ƶ���Ϣ���� 2002.11.30
		set rs=conn.execute("select count(incept) from [message] where flag=1 and issend=1 and sender<>'"&membername&"' and incept='"&membername&"'")
		cunm=cint(rs(0))		'ͳ���Ѷ����ֲ����Լ������Լ��Ķ���Ϣ 2002.11.30
		rs.close
		
		'�������޸�  ɾ�������Ѷ�����Ϣ 2002.11.30
	  	conn.execute("delete from [message] where flag=1 and issend=1 and incept='"&membername&"'")
		
		'��������� ɾ���Լ������Լ����ż������������� 2002.11.30
		if cunm=0 then
			Errmsg=Errmsg+"<br>"+"<li>����ռ���ֻ���Լ������Լ����ż���������û�еõ���������"
			call fuwu_err()			
			exit sub		
		end if
			
		sql="select * from [user] where username='"&membername&"' "
		rs.open sql,conn,1,3
		rs("userwealth") = rs("userwealth")+delmail*cunm
		rs.update
		rs.close   
	 
		sql = "select * from [bankconfig]"
		rs.open sql,conn1,1,3            
		rs("chubei")=rs("chubei")-delmail*cunm
		rs.update
		rs.close
		set rs=nothing
		
		if cint(log_setting(0))=1 and cint(log_setting(1))=1 then
			content="����Ѷ��ռ����õ���������<font color=red>"&delmail&"Ԫ/���"&cunm&"��="&delmail*cunm&"Ԫ</font>"
			call logs("����","ɾ������Ϣ",membername)
			sucmsg=sucmsg+"<br>"+"<li>���Ĳ�����Ϣ�Ѿ���¼�ڰ�"
		end if
		sucmsg=sucmsg+"<br>"+"<li>����ռ����Ѿ���գ��õ���������<font color=red>"&delmail&"Ԫ/���"&cunm&"��="&delmail*cunm&"Ԫ</font>"		
		call fuwu_suc()		
	else
		rs.close
		set rs=nothing
		Errmsg=Errmsg+"<br>"+"<li>�����ռ�����û���Ѷ�����Ϣ!"
		call fuwu_err()		
	end if
 
end sub

'------------------------------��IP--------------------------------
sub kip()
	if cint(fuwu_setting(6))=0 then 
		founderr=true
		Errmsg=Errmsg+"<br>"+"<li>�÷����Ѿ���ͣ���������Ա��ϵ"
		call bank_err()
		exit sub
	end if
	
	dim rsip,sqlip
	set rs=server.createobject("adodb.recordset")
	set rsip=server.createobject("adodb.recordset")
	sql="select userwealth from [user] where username='"&membername&"' "
	rs.open sql,conn,1,3
	if request.form("user")="" then
		Errmsg=Errmsg+"<br>"+"<li>��Ҫ��˭��IPѽ��"
		call fuwu_err()		
	elseif rs(0)<kanip then
		Errmsg=Errmsg+"<br>"+"<li>����ֽ𲻹�!"
		call fuwu_err()		
	elseif trim(request.form("user"))="������" then
		Errmsg=Errmsg+"<br>"+"<li><font color=blue>������˵��</font><font color=red>���У���Ҫ���ҵ�IP���ðɣ������ѵ�ร���л�����ҵĹ���!</font>"
		call fuwu_err()		
	else
		
		sqlip="select IP from [online] where username='"&checkStr(trim(request.form("user")))&"' "
		rsip.open sqlip,conn,1,3
		if rsip.eof then
			Errmsg=Errmsg+"<br>"+"<li>�Բ�����û�������"
			call fuwu_err()			
		else
%>    
<table border=0 cellspacing=1 cellpadding=3 align=center class=tableborder1><tr><th height=26 >�鿴 <font color=red><%=htmlencode(request.form("user"))%></font> ��IP��Ϣ</th></tr><tr><td align=center height=70 class=tablebody1>��Ա��<font color=navy><%=htmlencode(request.form("user"))%></font> ��IP�ǣ�<font color=red><%=rsip(0)%></font><br><br>���ԣ�<font color=red><%=address(rsip(0))%></font><br></td></tr><tr><td align=center height=26 class=tablebody2><a href="Z_fuwu.asp">����</a></td></tr></table>          
<%          
			rsip.close          
			rs("userwealth")=rs("userwealth")-kanip          
			rs.update          
			rs.close          
			sql = "select * from [bankconfig]"          
			rs.open sql,conn1,1,3                      
			rs("chubei")=rs("chubei")+kanip          
			rs.update          
			rs.close 
			  
			if cint(log_setting(0))=1 and cint(log_setting(1))=1 then
				content="�鿴 <font color=navy>"&htmlencode(request.form("user"))&"</font> ��IP��Ϣ��֧��������ã�"&kanip&" Ԫ"
				call logs("����","�鿴���߻�ԱIP",membername)
			end if
		       
		end if          
	end if          
          
end sub         
            
'-------------------------------Ⱥ������------------------------------------          
sub fdx()   
	if cint(fuwu_setting(5))=0 then 
		founderr=true
		Errmsg=Errmsg+"<br>"+"<li>�÷����Ѿ���ͣ���������Ա��ϵ"
		call bank_err()
		exit sub
	end if
	       
	dim rsdx,sqldx          
	set rs=server.createobject("adodb.recordset")          
	set rsdx=server.createobject("adodb.recordset")          
	sql="select userwealth from [user] where username='"&membername&"' "          
	rs.open sql,conn,1,3          
	if request.form("title")="" then          
		Errmsg=Errmsg+"<br>"+"<li>���ⲻ��Ϊ��,���������"
		call fuwu_err()            
	elseif request.form("content")="" then          
		Errmsg=Errmsg+"<br>"+"<li>��������ǿյ�,������д����"
		call fuwu_err()              
	elseif rs("userwealth")<faduangxun then          
		Errmsg=Errmsg+"<br>"+"<li>����ֽ𲻹�!"
		call fuwu_err()     
	else          
		rs("userwealth")=rs("userwealth")-faduangxun          
		rs.update          
		rs.close          
        if request.form("stype") = 1 then 
			content="����ϢȺ�������������߻�Ա"		         
			sql="select * from [online]"          
			rs.open sql,conn,1,3          
			do while not rs.eof          
				sqldx="select * from [message]"          
				rsdx.open sqldx,conn,1,3          
				rsdx.addnew          
				rsdx("sender")=membername          
				rsdx("incept")=rs("username")          
				rsdx("title")=request.form("title")          
				rsdx("content")=request.form("content")          
				rsdx("issend")=1          
				rsdx.update          
				rsdx.close          
				rs.movenext          
			loop          
         elseif request.form("stype") = 2 then          
		 	content="����ϢȺ���������й��"		 
			sql="select * from [user] where usergroupid=8"          
			rs.open sql,conn,1,3          
			do while not rs.eof          
				sqldx="select * from [message]"          
				rsdx.open sqldx,conn,1,3          
				rsdx.addnew          
				rsdx("sender")=membername          
				rsdx("incept")=rs("username")          
				rsdx("title")=request.form("title")          
				rsdx("content")=request.form("content")          
				rsdx("issend")=1          
				rsdx.update          
				rsdx.close          
				rs.movenext          
			loop          
         elseif request.form("stype") = 3 then    
		 	content="����ϢȺ���������г��������Ͱ���"		       
			sql="select * from [user] where usergroupid=3 or usergroupid=2"          
			rs.open sql,conn,1,3          
			do while not rs.eof          
				sqldx="select * from [message]"          
				rsdx.open sqldx,conn,1,3          
				rsdx.addnew          
				rsdx("sender")=membername          
				rsdx("incept")=rs("username")          
				rsdx("title")=request.form("title")          
				rsdx("content")=request.form("content")          
				rsdx("issend")=1          
				rsdx.update          
				rsdx.close          
				rs.movenext          
			loop          
		elseif request.form("stype") = 4 then          
		 	content="����ϢȺ���������й���Ա"
			sql="select * from [user] where usergroupid=1"          
			rs.open sql,conn,1,3          
			do while not rs.eof          
				sqldx="select * from [message]"          
				rsdx.open sqldx,conn,1,3          
				rsdx.addnew          
				rsdx("sender")=membername          
				rsdx("incept")=rs("username")          
				rsdx("title")=request.form("title")          
				rsdx("content")=request.form("content")          
				rsdx("issend")=1          
				rsdx.update          
				rsdx.close          
				rs.movenext          
			loop          
		elseif request.form("stype") = 5 then 
			content="����ϢȺ���������� ����Ա���������������������"		 		         
			sql="select * from [user] where usergroupid=1 or usergroupid=2 or usergroupid=3 or usergroupid=8"          
			rs.open sql,conn,1,3          
			do while not rs.eof          
			sqldx="select * from [message]"          
			rsdx.open sqldx,conn,1,3          
			rsdx.addnew          
			rsdx("sender")=membername          
			rsdx("incept")=rs("username")          
			rsdx("title")=request.form("title")          
			rsdx("content")=request.form("content")          
			rsdx("issend")=1          
			rsdx.update          
			rsdx.close          
			rs.movenext          
			loop          
		end if          
		rs.close          
		sql = "select * from [bankconfig]"          
		rs.open sql,conn1,1,3                      
		rs("chubei")=rs("chubei")+faduangxun          
		rs.update          
		rs.close
		     
		if cint(log_setting(0))=1 and cint(log_setting(1))=1 then
			content=content&"��֧��������ã�"&faduangxun&" Ԫ"
			call logs("����","����ϢȺ��",membername)
			sucmsg=sucmsg+"<br>"+"<li>���Ĳ�����Ϣ�Ѿ���¼�ڰ�"
		end if
		sucmsg=sucmsg+"<br>"+"<li>����ϢȺ���ɹ����뷵��"		
		call fuwu_suc()		       
	end if          
end sub           
       
'---------------------------------������=-------------------------------          
sub fgg()    
	if cint(fuwu_setting(4))=0 then 
		founderr=true
		Errmsg=Errmsg+"<br>"+"<li>�÷����Ѿ���ͣ���������Ա��ϵ"
		call bank_err()
		exit sub
	end if
	      
	if trim(request.form("title"))="" then
		founderr=true 
		Errmsg=Errmsg+"<br>"+"<li>������ⲻ��Ϊ��,���������" 
	end if
	if trim(request.form("content"))="" then
		founderr=true 
		Errmsg=Errmsg+"<br>"+"<li>���������ǿյ�,������д����" 
	end if
	if mymoney<fagonggao then
		founderr=true 
		Errmsg=Errmsg+"<br>"+"<li>����ֽ𲻹�" 
	end if
	BoardID=int(request.form("boardid"))		
         
	if founderr then
		call bank_err()        
	else          
		conn.execute("update [user] set userwealth=userwealth-"&fagonggao&" where userid="&userid)
		set rs=server.createobject("adodb.recordset")		         
		sql="select * from [bbsnews]"          
		rs.open sql,conn,1,3          
		rs.addnew          
		rs("boardid")=BoardID        
		rs("username")=request.form("username")          
		rs("title")=request.form("title")          
		rs("content")=request.form("content")          
		rs.update          
		rs.close 
		
		myCache.name="AnnounceMents"&BoardID
		myCache.makeEmpty	
		         
		sql = "select * from [bankconfig]"          
		rs.open sql,conn1,1,3                      
		rs("chubei")=rs("chubei")+fagonggao          
		rs.update          
		rs.close
		
		if cint(log_setting(0))=1 and cint(log_setting(1))=1 then
			content="������̳���棬֧��������ã�"&fagonggao&" Ԫ"
			call logs("����","��������",membername)
			sucmsg=sucmsg+"<br>"+"<li>���Ĳ�����Ϣ�Ѿ���¼�ڰ�"
		end if
		sucmsg=sucmsg+"<br>"+"<li>��ϲ�㣬���淢���ɹ�" 		
		call fuwu_suc()		
	end if          
end sub           
          
'--------------------------------��������ֵ------------------------- ������ ���         
sub power() 
	if cint(fuwu_setting(2))=0 then 
		founderr=true
		Errmsg=Errmsg+"<br>"+"<li>�÷����Ѿ���ͣ���������Ա��ϵ"
		call bank_err()
		exit sub
	end if
	 
	dim buypower  
	buypower=0      
  	set rs=server.createobject("adodb.recordset")          
  	sql="select userwealth,userpower from [user] where username='"&membername&"' "          
  	rs.open sql,conn,1,3          

		if request.form("power")="" then
			founderr=true 
			Errmsg=Errmsg+"<br>"+"<li>��������Ҫ�������������"
		elseif not isnumeric(request.form("power")) then
			founderr=true 
			Errmsg=Errmsg+"<br>"+"<li>�����������������������"
		elseif request.form("power")<0 then
			founderr=true
			Errmsg=Errmsg+"<br>"+"<li>�벻Ҫ���븺��"		
		else
			buypower=abs(int(request.form("power")))
			if rs(0)<buypower*addpower then         
				founderr=true 
				Errmsg=Errmsg+"<br>"+"<li>�����ֽ𲻹�"
			end if			
		end if
			
		if founderr then
			call bank_err()		 
      	else          
			rs(0)=rs(0)-buypower*addpower          
			rs(1)=rs(1)+buypower          
			rs.update          
			rs.close          
			sql = "select * from [bankconfig]"          
			rs.open sql,conn1,1,3                      
			rs("chubei")=rs("chubei")+addpower          
			rs.update          
			rs.close 
			
			if cint(log_setting(0))=1 and cint(log_setting(1))=1 then
				content="��������ֵ��֧��������ã�"&buypower&"�� �� "&addpower&" Ԫ/�� = "& buypower*addpower & " Ԫ"
				call logs("����","��������ֵ",membername)
				sucmsg=sucmsg+"<br>"+"<li>���Ĳ�����Ϣ�Ѿ���¼�ڰ�"
			end if
			sucmsg=sucmsg+"<br>"+"<li>��������ֵ�ɹ������ѣ�"&buypower&"�� �� "&addpower&" Ԫ/�� ="& buypower*addpower & "Ԫ"		
			call fuwu_suc()			         
      end if           
end sub           
          
'--------------------------------��������-------------------------          
sub ft() 
	if cint(fuwu_setting(3))=0 then 
		founderr=true
		Errmsg=Errmsg+"<br>"+"<li>�÷����Ѿ���ͣ���������Ա��ϵ"
		call bank_err()
		exit sub
	end if
	         
	dim fatie 
	fatie=0         
	set rs=server.createobject("adodb.recordset")          
	sql="select userwealth,Article from [user] where username='"&membername&"' "          
	rs.open sql,conn,1,3  
          
		if request.form("fatie")="" then
			founderr=true 
			Errmsg=Errmsg+"<br>"+"<li>��������Ҫ�����������"
		elseif not isnumeric(request.form("fatie")) then
			founderr=true 
			Errmsg=Errmsg+"<br>"+"<li>���������������������"
		elseif request.form("fatie")<0 then
			founderr=true
			Errmsg=Errmsg+"<br>"+"<li>�벻Ҫ���븺��"		
		else
			fatie=abs(int(request.form("fatie")))
			if rs(0)<fatie*addfatie then         
				founderr=true 
				Errmsg=Errmsg+"<br>"+"<li>�����ֽ𲻹�"
			end if			
		end if
			
		if founderr then
			call bank_err()		 
		else          
			rs(0)=rs(0)-fatie*addfatie          
			rs(1)=rs(1)+fatie          
			rs.update          
			rs.close          
			sql = "select * from [bankconfig]"          
			rs.open sql,conn1,1,3                      
			rs("chubei")=rs("chubei")+fatie*addfatie          
			rs.update          
			rs.close   
			
			if cint(log_setting(0))=1 and cint(log_setting(1))=1 then
				content="����������֧��������ã�"&fatie&"�� �� "&addfatie&" Ԫ/�� = "& fatie*addfatie & " Ԫ"
				call logs("����","��������",membername)
				sucmsg=sucmsg+"<br>"+"<li>���Ĳ�����Ϣ�Ѿ���¼�ڰ�"
			end if
			sucmsg=sucmsg+"<br>"+"<li>���������ɹ������ѣ�"&fatie&"�� �� "&addfatie&" Ԫ/�� ="& fatie*addfatie & "Ԫ"		
			call fuwu_suc()			        
		end if          
end sub 
          
'--------------------------------��������ֵ-------------------------          
sub ml()  
	if cint(fuwu_setting(0))=0 then 
		founderr=true
		Errmsg=Errmsg+"<br>"+"<li>�÷����Ѿ���ͣ���������Ա��ϵ"
		call bank_err()
		exit sub
	end if
	        
	dim tili   
	tili=0       
	set rs=server.createobject("adodb.recordset")          
	sql="select userwealth,usercp from [user] where username='"&membername&"' "          
	rs.open sql,conn,1,3 
		if request.form("tili")="" then
			founderr=true 
			Errmsg=Errmsg+"<br>"+"<li>��������Ҫ���������ֵ"
		elseif not isnumeric(request.form("tili")) then
			founderr=true 
			Errmsg=Errmsg+"<br>"+"<li>���������ֵ����������"
		elseif request.form("tili")<0 then
			founderr=true
			Errmsg=Errmsg+"<br>"+"<li>�벻Ҫ���븺��"		
		else
			tili=abs(int(request.form("tili")))
		end if
      	if rs(0)<tili*addtili then         
			founderr=true 
			Errmsg=Errmsg+"<br>"+"<li>�����ֽ𲻹�"
		end if			
		if founderr then
			call bank_err()           
      	else     
			rs(0)=rs(0)-tili*addtili          
			rs(1)=rs(1)+tili 		     
			rs.update          
			rs.close          
			sql = "select * from [bankconfig]"          
			rs.open sql,conn1,1,3                      
			rs("chubei")=rs("chubei")+tili*addtili          
			rs.update          
			rs.close  
			 
			if cint(log_setting(0))=1 and cint(log_setting(1))=1 then
				content="��������ֵ��֧��������ã�"&tili&"�� �� "&addtili&" Ԫ/�� = "& tili*addtili & " Ԫ"
				call logs("����","��������ֵ",membername)
				sucmsg=sucmsg+"<br>"+"<li>���Ĳ�����Ϣ�Ѿ���¼�ڰ�"
			end if
			sucmsg=sucmsg+"<br>"+"<li>��������ֵ�ɹ������ѣ�"&tili&"�� �� "&addtili&" Ԫ/�� ="& tili*addtili & "Ԫ"		
			call fuwu_suc()			       
      	end if           
end sub 
          
'--------------------------------������ֵ-------------------------          
sub jy() 
	if cint(fuwu_setting(1))=0 then 
		founderr=true
		Errmsg=Errmsg+"<br>"+"<li>�÷����Ѿ���ͣ���������Ա��ϵ"
		call bank_err()
		exit sub
	end if
	         
	dim jinyan     
	jinyan=0     
	set rs=server.createobject("adodb.recordset")          
	sql="select userwealth,userep from [user] where username='"&membername&"' "          
	rs.open sql,conn,1,3
       
		if request.form("jinyan")="" then
			founderr=true 
			Errmsg=Errmsg+"<br>"+"<li>��������Ҫ����ľ���ֵ"
		elseif not isnumeric(request.form("jinyan")) then
			founderr=true 
			Errmsg=Errmsg+"<br>"+"<li>����ľ���ֵ����������"
		elseif request.form("jinyan")<0 then
			founderr=true
			Errmsg=Errmsg+"<br>"+"<li>�벻Ҫ���븺��"		
		else
			jinyan=abs(int(request.form("jinyan")))
		end if
      	if rs(0)<jinyan*addjinyan then         
			founderr=true 
			Errmsg=Errmsg+"<br>"+"<li>�����ֽ𲻹�"
		end if			
		if founderr then
			call bank_err()        
        else          
			rs(0)=rs(0)-jinyan*addjinyan          
			rs(1)=rs(1)+jinyan          
			rs.update          
			rs.close          
			sql = "select * from [bankconfig]"          
			rs.open sql,conn1,1,3                      
			rs("chubei")=rs("chubei")+jinyan*addjinyan          
			rs.update          
			rs.close
			
			if cint(log_setting(0))=1 and cint(log_setting(1))=1 then
				content="������ֵ��֧��������ã�"&jinyan&"�� �� "&addjinyan&" Ԫ/�� = "& jinyan*addjinyan & " Ԫ"
				call logs("����","������ֵ",membername)
				sucmsg=sucmsg+"<br>"+"<li>���Ĳ�����Ϣ�Ѿ���¼�ڰ�"
			end if
			sucmsg=sucmsg+"<br>"+"<li>������ֵ�ɹ������ѣ�"&jinyan&"�� �� "&addjinyan&" Ԫ/�� ="& jinyan*addjinyan & "Ԫ"		
			call fuwu_suc()			       
      end if           
end sub 
        
'-------------------------------------------������-----------------------------------          
sub main()    
'fuwu_setting : canCP/canEP/canPower/canfatie/canGonggao/canDuangxun/canKanip/canDelmail  1-�� 0����  ������ 2002.11.30
'fuwu_setting():  0      1      2        3         4           5         6         7      
%>          
<table cellspacing=1 cellpadding=3 align=center class=tableborder1 style="width:97%">          
<tr><th align=center height=26><b>������������</b></th></tr>          
<tr><td align=center height=26 class=tablebody1>��ӭ����������������,�����е��ֽ�<font color="#FF0000"><%=mymoney%></font>Ԫ�����ǽ������ṩ���ʵķ���,��ѡ�����ķ�����Ŀ��!</td></tr>
<tr><td align=left height=26 class=tablebody2> <b>�������</b>��<%if cint(fuwu_setting(0))=1 or cint(fuwu_setting(1))=1 or cint(fuwu_setting(2))=1 or cint(fuwu_setting(3))=1 then%><font color=blue>��</font><%else%><font color=red>��</font><%end if%></td></tr>       
          
<form name="form1" method="post" action="Z_fuwu.asp"><tr ><td align=left height=30 class=tablebody1><%if cint(fuwu_setting(1))=1 then%><font color=blue>��</font><%else%><font color=red>��</font><%end if%>��������<b>������ֵ</b> ��������Ҫ������پ���ֵ:<INPUT name=jinyan> <INPUT type=submit value=���� name=submit> ����:<font color="#FF0000"><%=addjinyan%></font>Ԫ ���������ҵľ���ֵ: <font color="#FF0000"><%=myuserep%></font></td></tr><INPUT type=hidden value="1" name=menu></form>          
          
<form name="form2" method="post" action="Z_fuwu.asp"><tr ><td align=left height=30 class=tablebody1><%if cint(fuwu_setting(0))=1 then%><font color=blue>��</font><%else%><font color=red>��</font><%end if%>��������<b>��������ֵ</b> ��������Ҫ�����������ֵ:<INPUT name=tili> <INPUT type=submit value=���� name=submit class=tablebody1> ����:<font color="#FF0000"><%=addtili%></font>Ԫ ���������ҵ�����ֵ: <font color="#FF0000"><%=myusercp%></font></td></tr><INPUT type=hidden value="2" name=menu></form>          
          
<form name="form3" method="post" action="Z_fuwu.asp"><tr ><td align=left height=30 class=tablebody1><%if cint(fuwu_setting(3))=1 then%><font color=blue>��</font><%else%><font color=red>��</font><%end if%>��������<b>��������</b> ��������Ҫ������ٷ�����:<INPUT name=fatie> <INPUT type=submit value=���� name=submit> ����:<font color="#FF0000"><%=addfatie%></font>Ԫ ���������ҵķ�����: <font color="#FF0000"><%=myArticle%></font></td></tr><INPUT type=hidden value="3" name=menu></form>          
          
<form name="form4" method="post" action="Z_fuwu.asp"><tr ><td align=left height=30 class=tablebody1><%if cint(fuwu_setting(2))=1 then%><font color=blue>��</font><%else%><font color=red>��</font><%end if%>��������<b>����<font color="#FF0000">����ֵ</font></b> ��������Ҫ�����������ֵ:<INPUT name=power> <INPUT type=submit value=���� name=submit> ����:<font color="#FF0000"><%=addpower%></font>Ԫ ���������ҵ�����ֵ: <font color="#FF0000"><%=mypower%></font></td></tr><INPUT type=hidden value="4" name=menu></form>          
          
<tr ><td align=left height=30 class=tablebody2> <b>���淢��</b> ���淢��ÿ����Ҫ<font color="#FF0000"><%=fagonggao%></font>Ԫ ��<%if cint(fuwu_setting(4))=1 then%><font color=blue>��</font><%else%><font color=red>��</font><%end if%></td></tr>          
<form name="form5" method="post" action="Z_fuwu.asp">
<tr><td align=center height=30 class=tablebody1> 
                 
<table cellspacing=1 cellpadding=3 align=center class=tableborder1 style="width:70%">          
<tr>          
<th colspan="2">���淢��</th>          
</tr>
<tr>          
<td width="25%" class=tablebody2>��������:</td>          
<td align=left class=tablebody1>          
<% set rs=server.createobject("adodb.recordset")         
sql="select boardid,boardtype from board"          
   rs.open sql,conn,1,1          
%>          
        <select name="boardid" size="1">          
          <option value="0">��̳��ҳ</option>
          <%
	do while not rs.eof
        response.write "<option value='"+CStr(rs("BoardID"))+"'>"+rs("Boardtype")+"</option>"+chr(13)+chr(10)
	rs.movenext
	loop
	rs.close
%>
        </select>          
</td></tr>          
<tr>           
      <td width="25%" valign=top class=tablebody2>�����ˣ�</td>          
      <td width="75%" class=tablebody1>          
        <input type=text name=username size=36 value="<%=membername%>" disabled>          
        <input type=hidden name=username value="<%=membername%>">          
      </td>          
    </tr>          
    <tr>           
      <td width="25%" valign=top class=tablebody2>���⣺</td>          
      <td width="75%" class=tablebody1>          
        <input type=text name=title size=36>          
      </td>          
    </tr>          
    <tr>           
      <td width="25%" valign=top class=tablebody2>���ݣ�</td>          
      <td width="75%" class=tablebody1>          
        <textarea cols=70 rows=6 name="content"></textarea>          
      </td>          
    </tr>          
    <tr>          
      <td width="100%" valign=top colspan="2" align=center class=tablebody2><input type=Submit value="�� ��" name=Submit>&nbsp;<input type="reset" name="Clear" value="�� ��"></td>          
    </tr>          
</table>          
          
</td></tr>
<INPUT type=hidden value="5" name=menu>
</form>          
          
<tr><td align=left height=30 class=tablebody2> <b>����ϢȺ��</b> ÿ��<font color=red><%=faduangxun%></font>Ԫ��<%if cint(fuwu_setting(5))=1 then%><font color=blue>��</font><%else%><font color=red>��</font><%end if%></td></tr>          
          
          
<form name="form6" method="post" action="Z_fuwu.asp"><tr  ><td align=center height=30 class=tablebody1>            
<table cellspacing=1 cellpadding=3 class=tableborder1 style="width:70%">          
<tr>          
<th colspan="2">����ϢȺ��</th>          
</tr>
<tr>           
                  <td width="30%" class=tablebody2>��Ϣ����</td>          
                  <td width="70%" class=tablebody1>           
                    <input type="text" name="title" size="50">          
                  </td>          
                </tr>          
<tr>           
                  <td width="30%" class=tablebody2>���շ�ѡ��</td>          
                  <td width="70%" class=tablebody1>           
                    <select name=stype size=1>          
					<option value="1">���������û�</option>
					<option value="2">���й��</option>
					<option value="3">���а���</option>
					<option value="4">���й���Ա</option>
					<option value="5">���/����/����Ա</option>
					</select>
                  </td>
                </tr>
                <tr> 
                  <td width="30%" height="20" valign="top" class=tablebody2>
                    <p>��Ϣ����</p>
                  </td>
                  <td width="70%" height="20" class=tablebody1> 
                    <textarea name="content" cols="70" rows="10"></textarea>
                  </td>
                </tr>
                <tr> 
                  <td colspan=2 align=center height="23" class=tablebody2><input type="submit" name="Submit" value="������Ϣ">  <input type="reset" name="Submit2" value="������д">
                  </td>
                </tr>
</table>
</td></tr>
<INPUT type=hidden value="6" name=menu>
</form>
<tr  height=30 ><td align=left height=26 class=tablebody2> <b>��������</b>��<%if cint(fuwu_setting(7))=1 or cint(fuwu_setting(6))=1 then%><font color=blue>��</font><%else%><font color=red>��</font><%end if%></td></tr>
<form name="form7" method="post" action="Z_fuwu.asp">
<tr  ><td align=left height=30 class=tablebody1><%if cint(fuwu_setting(6))=1 then%><font color=blue>��</font><%else%><font color=red>��</font><%end if%>��������<b>�鿴���߻�ԱIP</b>  ÿ��<font color=red><%=kanip%></font>Ԫ �鿴 <INPUT name=user> ��IP����Դ <INPUT type=submit value=�鿴 name=submit></td></tr>          
<INPUT type=hidden value="7" name=menu>
</form>          
          
<form name="form8" method="post" action="Z_fuwu.asp">          
<tr><td align=left height=30 class=tablebody1><%if cint(fuwu_setting(7))=1 then%><font color=blue>��</font><%else%><font color=red>��</font><%end if%>��������<b>ɾ���Ѷ�����Ϣ</b>  ÿ��ɾ�����õ����� <font color=red><%=delmail%>Ԫ������Ϣ���� </font> ��������  <INPUT type=submit value=ɾ�� name="b1" onclick="javascript:{if(confirm('��ȷ��ִ��ɾ��������?')){return true;}return false;}"> (<font color="#FF0000">ע�⣺�˹��ܽ�ɾ����������Ѷ��ʼ������ɻָ���</font>)</td></tr>          
<INPUT type=hidden value="8" name=menu>
</form>          
</table>
<br>          
<%          
end sub          

sub fuwu_err() 
%>
<table cellspacing=1 cellpadding=3 align=center class=tableborder1 style="width:75%">
	<tr><th height=25>���������Ϣ</th></tr>
	<tr><td width="100%" class=tablebody1><b>��������Ŀ���ԭ��</b><br><br><li>���Ƿ���ϸ�Ķ���<a href="boardhelp.asp?boardid=<%=boardid%>">�����ļ�</a>����������û�е�¼���߲�����ʹ�õ�ǰ���ܵ�Ȩ��
	<font color=<%=Forum_body(8)%>><%=errmsg%></font><br></td></tr>
	<tr><td align=center height=26 class=tablebody2><a href="javascript:history.go(-1)">������һҳ</a></td></tr>
</table>
<%
end sub         

sub fuwu_suc() 
%>
<table cellspacing=1 cellpadding=3 align=center class=tableborder1 style="width:75%">
	<tr><th height=25>��������ɹ�</th></tr>
	<tr><td width="100%" class=tablebody1><b>�����ɹ���</b><br><br><li>��ӭ����<%=Forum_info(0)%> ������������
	<font color=navy><%=sucmsg%></font><br></td></tr>
	<tr><td align=center height=26 class=tablebody2><a href=<%=Request.ServerVariables("HTTP_REFERER")%>>������һҳ</a></td></tr>
</table>
<%
end sub

 '-----------------------------��������---------------------------------          
sub admin() 
if not master then
	Errmsg=Errmsg+"<br>"+"<li>��û�н���������Ĺ����Ȩ��,�������Ա��ϵ"
	call fuwu_err() 	
else	
'fuwu_setting : canCP/canEP/canPower/canfatie/canGonggao/canDuangxun/canKanip/canDelmail  1-�� 0����  ������ 2002.11.30
'fuwu_setting():  0      1      2        3         4           5         6         7
%>          
	<br>
	<table border=0 cellspacing=1 cellpadding=3 align=center class=tableborder1 style="width:97%">          
	<tr  height=22><th colspan=3  height=26 >�������Ĺ���----����۸�����</th></tr>          
	<form name="form1" method="post" action="Z_fuwu.asp">           
	<tr >          
		<td align=right width="35%" class=tablebody1>ÿ�㾭��ļ۸�</td>          
		<td align=left width="35%" class=tablebody1><INPUT name=jinyan value="<%=addjinyan%>"> Ԫ</td> 
		<td align=left width="30%" class=tablebody1>�Ƿ�ͨ�÷��� 
		<input type=radio name="canEP" value=0 <%if cint(fuwu_setting(1))=0 then%>checked<%end if%>>��&nbsp;
		<input type=radio name="canEP" value=1 <%if cint(fuwu_setting(1))=1 then%>checked<%end if%>>��&nbsp;
		</td>	        
	</tr>          
	<tr >          
		<td align=right width="35%" class=tablebody1>ÿ�������ļ۸�</td>          
		<td align=left width="35%" class=tablebody1><INPUT name=tili value="<%=addtili%>"> Ԫ</td>          
		<td align=left width="30%" class=tablebody1>�Ƿ�ͨ�÷��� 
		<input type=radio name="canCP" value=0 <%if cint(fuwu_setting(0))=0 then%>checked<%end if%>>��&nbsp;
		<input type=radio name="canCP" value=1 <%if cint(fuwu_setting(0))=1 then%>checked<%end if%>>��&nbsp;
		</td>	
	</tr>          
	<tr >          
		<td align=right width="35%" class=tablebody1>ÿ�㷢���ļ۸�</td>          
		<td align=left width="35%" class=tablebody1><INPUT name=fatie value="<%=addfatie%>">  Ԫ</td>
		<td align=left width="30%" class=tablebody1>�Ƿ�ͨ�÷��� 
		<input type=radio name="canfatie" value=0 <%if cint(fuwu_setting(3))=0 then%>checked<%end if%>>��&nbsp;
		<input type=radio name="canfatie" value=1 <%if cint(fuwu_setting(3))=1 then%>checked<%end if%>>��&nbsp;
		</td>	         
	</tr>         
	<tr >         
		<td align=right width="35%" class=tablebody1>ÿ�������ļ۸�</td>        
		<td align=left width="35%" class=tablebody1><INPUT name=power value="<%=addpower%>">  Ԫ</td>        
		<td align=left width="30%" class=tablebody1>�Ƿ�ͨ�÷��� 
		<input type=radio name="canPower" value=0 <%if cint(fuwu_setting(2))=0 then%>checked<%end if%>>��&nbsp;
		<input type=radio name="canPower" value=1 <%if cint(fuwu_setting(2))=1 then%>checked<%end if%>>��&nbsp;
		</td>	
	</tr>        
	<tr >        
		<td align=right width="35%" class=tablebody1>������ļ۸�</td>        
		<td align=left width="35%" class=tablebody1><INPUT name=gonggao value="<%=fagonggao%>"> Ԫ</td>          
		<td align=left width="30%" class=tablebody1>�Ƿ�ͨ�÷��� 
		<input type=radio name="canGonggao" value=0 <%if cint(fuwu_setting(4))=0 then%>checked<%end if%>>��&nbsp;
		<input type=radio name="canGonggao" value=1 <%if cint(fuwu_setting(4))=1 then%>checked<%end if%>>��&nbsp;
		</td>	
	</tr>          
	<tr >          
		<td align=right width="35%" class=tablebody1>Ⱥ�����ŵļ۸�</td>          
		<td align=left width="35%" class=tablebody1><INPUT name=duangxun value="<%=faduangxun%>"> Ԫ</td>  
		<td align=left width="30%" class=tablebody1>�Ƿ�ͨ�÷��� 
		<input type=radio name="canDuangxun" value=0 <%if cint(fuwu_setting(5))=0 then%>checked<%end if%>>��&nbsp;
		<input type=radio name="canDuangxun" value=1 <%if cint(fuwu_setting(5))=1 then%>checked<%end if%>>��&nbsp;
		</td>	        
	</tr>          
	<tr >          
		<td align=right width="35%" class=tablebody1>��IP�ļ۸�</td>          
		<td align=left width="35%" class=tablebody1><INPUT name=kanip value="<%=kanip%>"> Ԫ</td> 
		<td align=left width="30%" class=tablebody1>�Ƿ�ͨ�÷��� 
		<input type=radio name="canKanip" value=0 <%if cint(fuwu_setting(6))=0 then%>checked<%end if%>>��&nbsp;
		<input type=radio name="canKanip" value=1 <%if cint(fuwu_setting(6))=1 then%>checked<%end if%>>��&nbsp;
		</td>	         
	</tr>          
	<tr >          
		<td align=right width="35%" class=tablebody1> ɾ��<font color=red>1</font>���ʼ����ۣ�</td>          
		<td align=left width="35%" class=tablebody1><INPUT name=delmail value="<%=Delmail%>"> Ԫ</td>  
		<td align=left width="30%" class=tablebody1>�Ƿ�ͨ�÷��� 
		<input type=radio name="canDelmail" value=0 <%if cint(fuwu_setting(7))=0 then%>checked<%end if%>>��&nbsp;
		<input type=radio name="canDelmail" value=1 <%if cint(fuwu_setting(7))=1 then%>checked<%end if%>>��&nbsp;
		</td>	         
	</tr>          
	<tr> 	       
		<td align=center colspan=3 class=tablebody2><INPUT type=submit value=�޸� name=submit></td>          
	</tr> 
	<INPUT type=hidden value="10" name=menu>         
	</form> 	         
	</table> 
         
	<% 
	end if
end sub 
%>