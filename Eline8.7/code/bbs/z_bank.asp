<!--#include file="conn.asp"-->
<!-- #include file="inc/const.asp" -->
<!--#include file="Z_bankconn.asp"-->
<!--#include file="Z_bankconfig.asp"-->
<!--#include file="z_plus_check.asp"-->
<%
dim menu,savem,rsbank,rssql
dim lockac     '�ʻ��Ƿ񱻶���  ������ 2002.10.23
dim smoney,saveday,time1,time2,daikuang,dkdu,daiday

menu=request.querystring("menu")
if menu="" then menu=0
stats="�������� ����Ӫҵ����"

select case menu
	case 1
		stats="�������� ������ʮ����"
		call nav()
		call head_var(0,0,"��������","Z_bank.asp")
	case 2
		stats="�������� ���ж�ʮ�󴢻�"
		call nav()
		call head_var(0,0,"��������","Z_bank.asp")
	case 8
		stats="�������� �����г��칫��"
		call nav()
		call head_var(0,0,"��������","Z_bank.asp")
	case 13
		stats="�������� ��������¼���¼"
		call nav()
		call head_var(0,0,"��������","Z_bank.asp")		
	case else
		call nav()
		call head_var(2,0,"","")
end select

if not founduser then
	Errmsg=Errmsg+"<br>"+"<li>��û�н���������е�Ȩ�ޣ����ȵ�¼����ͬ����Ա��ϵ��"
	call dvbbs_error()
elseif cint(bank_setting(6))=1  and (not master) then '�����Ƿ����ö�ʱӪҵ
	if DatePart("h", time)<cint(Chen_BusinessTimeSlice(0)) or DatePart("h", time)>=cint(Chen_BusinessTimeSlice(1)) then
		call BankClose()	
	else
		call start()
	end if
else
	call start()
end if
if founderr then  call dvbbs_error()
call activeonline()
call footer() 
'=========================================================
sub start()
	set rs=server.createobject("adodb.recordset")
    sql="select * from [bank] where username='"&membername&"' "
    rs.open sql,conn1,1,3
    if rs.eof then
		rs.addnew
		rs("username")=membername
		rs("bankuser_setting")="0,"&bank_setting(4)
		rs.update
		smoney=0
    end if
    daikuang=rs("daikuang")
	lockac=rs("lockac")
	smoney=rs("savemoney")
	bankuser_setting=split(rs("bankuser_setting"),",")
	
	'�������������
	if daikuang>0 then
    	dkdu=0
	elseif cint(bankuser_setting(0))=1 then
		dkdu=int(mymoney*bankuser_setting(1))         
	else
		dkdu=int(mymoney*bank_setting(4))	
	end if

    saveday=datediff("d",rs("date"),date())
    daiday=datediff("d",rs("dkdate"),date())
    rs.close
'-------------------------
%>
	<table cellpadding=3 cellspacing=1 align=center class=tableborder1>
	<tr>
		<th height=25><%=stats%></td>
	</tr>
	<tr>
		<td class=tablebody2 height=1>	
<%
	founderr=false
	
	call bankhead()
	
	select case menu
		case 0
			call main()					'����Ӫҵ����
		case 1
			call forum_top20()			'������ʮ����
		case 2
			call bank_top20()			'���ж�ʮ�󴢻�   
		case 3
			call savemoney()			'���	
		case 4
			call qumoney()				'ȡ��
		case 5
			call zmoney()				'ת��
		case 6
			call daimoney()				'����
		case 7
			call hmoney()   			'����
		case 8
			call admin()				'���й���ҳ��
		case 9 
			call admin1()				'������������
		case 10
			call hmoney2()				'ǿ�ƻ��� 
		case 11
			call jiangli()				'����
		case 12
			call savelogsetting()		'�����¼�����
		case 13
			call banklog()				'�����¼��鿴
		case else 
			call main() 											
	end select
%>
	</td></tr></table>
<%
	founderr=false
end sub   

'=======================================================================================
'-------------------------------------------������-----------------------------------
sub main()
     if daiday>daitian and daikuang>0 then
%>
		<table cellspacing=1 cellpadding=3 align=center class=tableborder1 style="width:97%"><tr><th height=26 >���б���</th></tr><tr height=200><td align=center height=26 class=tablebody1><br><font color=red>������Ĵ�����˳�������,����ǿ��ִ�У�<%=hmoney2%></font><br><br></td></tr><tr><td align=center height=26 class="tablebody1"><a href="Z_bank.asp">�������д���</a></td></tr></table>
<%   else  %>

<table cellspacing=1 cellpadding=3 align=center class=tableborder1 style="width:97%">
	<tr>
		<th height=26>��ӭ�����������й������ĲƲ�</th>
	</tr>
	<tr>
		<td align=center height=26 class=tablebody2>����Ŀǰ���������� <font color="#FF0000"><%=culi%>%</font>������������ <font color="#FF0000"><%=daili%>%</font>��������ʴӵ� <font color="#FF0000"><%=Chen_StartLixiDay%></font> ������㣬��ǰ���д����ʽ�Ϊ <font color="#FF0000"><%=chubei%></font><br>���������Ϊ <font color="#FF0000"><%=daitian%></font> �죨���ڽ�ǿ��ִ�У���ÿ�δ�ȡ���Զ�������Ϣ</td>
	</tr>
<%if lockac then%>	
	<tr>
		<td align=center height=26 class=tablebody1><font color="#FF0000">���������˻��Ѿ�������,�����ܽ���ȡ�ת�ˡ��������</font></td>
	</tr>   
<%
end if
if cint(bankstate)=1 then
%>    
	<form name="form1" method="post" action="Z_bank.asp?menu=3"><tr><td align=left height=30 class=tablebody1><%if mymoney>0 and cint(bank_setting(0))=1 then%><font color=blue>��</font><%else%><font color=red>��</font><%end if%>��������<font face=Wingdings>v</font> �� <INPUT name=saving> <INPUT type=submit value=���� name=submit> ����ֽ�Ϊ��<font color="#FF0000"><%=mymoney%></font> Ԫ���������Ϊ��<font color="#FF0000"><%=culi%>%</font> </td></tr></form>
	
	<form name="form2" method="post" action="Z_bank.asp?menu=4"><tr><td align=left height=30 class=tablebody1><%if lockac or smoney<=0 or cint(bank_setting(1))=0 then%><font color=red>��</font><%else%><font color=blue>��</font><%end if%>��������<font face=Wingdings>v</font> ȡ� <INPUT name=draw> <INPUT type=submit value=ȡ�� name=submit> �����ڵĴ��Ϊ��<font color="#FF0000"><%=smoney%></font> Ԫ�������ϢΪ��<font color="#FF0000"><% if saveday<Chen_StartLixiDay then%>0<%else%><%=clng((formatnumber(smoney)*(formatnumber(culi)/100))*saveday)%><%end if%></font> Ԫ <%if smoney>0 then%>��Ĵ��������<font color="#FF0000"><%=saveday%></font> �� <%end if%></td></tr></form>
	
	<form name="form3" method="post" action="Z_bank.asp?menu=5"><tr><td align=left height=30 class=tablebody1><%if lockac or cint(bank_setting(3))=0 or (daikuang>0 and cint(bank_setting(5))=0)  then%><font color=red>��</font><%else%><font color=blue>��</font><%end if%>��������<font face=Wingdings>v</font> ת�ˣ� <INPUT name=trans> Ԫ �� <INPUT name=transTo> <INPUT type=submit value=ת�� name=submit> ת�������ѣ�<font color="#FF0000"><%=zhuangli%>%</font> </td></tr></form>      
		  
	<form name="form4" method="post" action="Z_bank.asp?menu=6"><tr><td align=left height=30 class=tablebody1><%if lockac or daikuang>0 or cint(bank_setting(2))=0 then%><font color=red>��</font><%else%><font color=blue>��</font><%end if%>��������<font face=Wingdings>v</font> ��� <INPUT name=loan> <INPUT type=submit value=���� name=submit> ��Ĵ�����Ϊ��<font color="#FF0000"><%=dkdu%></font> Ԫ����������Ϊ��<font color="#FF0000"><%=daili%>%</font>
			  </td></tr></form>
	<%if daikuang>0 then%>
	<form name="form5" method="post" action="Z_bank.asp?menu=7"><tr><td align=left height=30 class=tablebody1>�������������� <INPUT type=submit value=���� name=submit> �����ڵĴ����Ϊ��<font color="#FF0000"><%=daikuang%></font>����������ϢΪ��<font color="#FF0000"><%if daikuang=0 then %>0<%else%><%=clng((daikuang*(formatnumber(culi)/100))*daiday)%></font> Ԫ ��Ĵ���������<font color="#FF0000"><%=daiday%></font> ��<%end if%></td></tr></form>
	<%end if%>
<%else%>
	<tr><td align=center height=25 class=tablebody1><font face=Wingdings color=blue>v</font><font color=red> ���н���/��������ʱֹͣӪҵ�����������������Ա��ϵ </font><font face=Wingdings color=blue>v</font></td></tr>
<%end if%>
	<tr><td align=center height=25 class=tablebody2><font color=gray>����Ӫҵʱ�䣺<% if cint(bank_setting(6))=0 then%>ȫ��Ӫҵ<%else%><%=Chen_BusinessTimeSlice(0)%>:00��<%=Chen_BusinessTimeSlice(1)%>:00<%end if%></font></td></tr>
</table>
<br>
<%
end if
end sub
'--------------------------------ǿ�ƻ���-------------------------------
function hmoney2()
    dim name
    if request.querystring("username")<>"" then
    	name=checkStr(trim(request.querystring("username")))
		if not master then 
			Errmsg=Errmsg+"<br>"+"<li>��û��ִ��ǿ�ƻ��������Ȩ�ޣ��������Ա��ϵ"
			call bank_err()
			exit function
		end if	
    else
    	name=membername
    end if
	set rs=server.createobject("adodb.recordset")
	sql="select * from bank where username='"&name&"' "
	rs.open sql,conn1,1,3
	if rs.eof and rs.bof then 
		errmsg=errmsg+"<br>"+"<li>�û�["&name&"]�������˺Ų�����"
		call bank_err()
	else
		dim lixi,benjin,loan
		benjin=rs("daikuang")
		lixi=clng((daikuang*(formatnumber(daili)/100))*daiday)		
		loan=lixi+benjin
		rs("daikuang")=0
		rs("dkdate")=date()
		rs.update
		rs.close
		
		sql="select * from bankconfig"
		rs.open sql,conn1,1,3
		rs("chubei")=rs("chubei")+loan
		rs.update
		rs.close
		
		sql="select * from [user] where username='"&name&"' "
		rs.open sql,conn,1,3
		rs("userwealth")=rs("userwealth")-loan
		rs.update
		rs.close

		if cint(log_setting(0))=1 and cint(log_setting(3))=1 then
			content="<font color=red> <font color=blue>"&name&"</font> ��ǿ�ƻ������д��ԭ������Ϊ��"&benjin&" Ԫ��������Ϊ��"&loan&"Ԫ</font>"
			call logs("����","ǿ�ƻ���","�����г�")
		end if
		
		if request.querystring("username")<>"" then 
			sucmsg=sucmsg+"<br>"+"<li>�û�["&name&"]�Ľ���Ѿ���ǿ�ƻ��壬����"&benjin&"Ԫ����Ϣ��"&lixi&"Ԫ�������ܽ��Ϊ��"&loan&"Ԫ"
			call bank_suc()
		else
			hmoney2="<br>"+"<font color=navy>���Ľ���Ѿ���ǿ�ƻ��壬����"&benjin&"Ԫ����Ϣ��"&lixi&"Ԫ�������ܽ��Ϊ��"&loan&"Ԫ</font>"
		end if 
		
	end if 
end function 

'--------------------------------��������-------------------------------
sub hmoney()
	if bankstate=0 then		'������ ��� 2002.11.30
		Errmsg=Errmsg+"<br>"+"<li>������ͣӪҵ���������Ա��ϵ"
		call bank_err()
		exit sub	
	end if
		dim loan
		set rs=server.createobject("adodb.recordset")
		sql="select * from bank where username='"&membername&"' "
		rs.open sql,conn1,1,3
		set rsbank=server.createobject("adodb.recordset")
		rssql="select * from [user] where username='"&membername&"'"
		rsbank.open rssql,conn,1,3
		loan=clng((daikuang*(formatnumber(daili)/100))*daiday)+rs("daikuang")
        if rs("daikuang")=0 then
			rsbank.close
			Errmsg=Errmsg+"<br>"+"<li>��û�д��"
	    	call bank_err()
		elseif loan>rsbank("userwealth") then
			rsbank.close
			rs("lockac")=true
			rs.update
			Errmsg=Errmsg+"<br>"+"<li>��û����ô���Ǯ������"
			Errmsg=Errmsg+"<br>"+"<li>���н�����ʱ�������������ʻ���ֱ�����������Ĵ���"
	    	call bank_err()		
			content="����Ǯ��������˺��Զ�������"
			call logs("����","��������",membername)	
        else
			rsbank.close
			rs("daikuang")=0
			rs("dkdate")=date()
			'rs("lockac")=false
			rs.update
			set rs=server.createobject("adodb.recordset")
			sql="select * from bankconfig"
			rs.open sql,conn1,1,3
			rs("chubei")=rs("chubei")+loan
			rs.update
			rs.close
			set rs=server.createobject("adodb.recordset")
			sql="select * from [user] where username='"&membername&"' "
			rs.open sql,conn,1,3
			rs("userwealth")=rs("userwealth")-loan
			rs.update
			rs.close
			
			if cint(log_setting(0))=1 and cint(log_setting(4))<>0 then
				if cint(log_setting(4))=1 or (cint(log_setting(4))=2 and loan>=clng(log_setting(5))) then
					content="�������д��ԭ������Ϊ��"&daikuang&" Ԫ��������Ϊ��"&loan&"Ԫ"
					call logs("����","��������",membername)
					sucmsg=sucmsg+"<br>"+"<li>���Ĳ�����Ϣ�Ѿ���¼�ڰ�"
				end if	
			end if
					   
			sucmsg=sucmsg+"<br>"+"<li>���Ľ���Ѿ�ȫ�����壬ԭ������Ϊ��"&daikuang&" Ԫ��������Ϊ��"&loan&"Ԫ"
			call bank_suc() 		   
		end if
end sub

'--------------------------------�������--------------------------------
sub daimoney()
	if bankstate=0 then			'������ ��� 2002.11.30
		Errmsg=Errmsg+"<br>"+"<li>������ͣӪҵ���������Ա��ϵ"
		call bank_err()
		exit sub	
	end if
	if cint(bank_setting(2))=0 then			'������ ��� 2002.11.30
		Errmsg=Errmsg+"<br>"+"<li>���д�������Ѿ���ͣ���������Ա��ϵ"
		call bank_err()
		exit sub	
	end if	
	if lockac then 			'������ ���
		Errmsg=Errmsg+"<br>"+"<li>�Բ��������ܽ��иò���"
		Errmsg=Errmsg+"<br>"+"<li>���������˻��Ѿ������ᣬ�������Ա��ϵ"
    	call bank_err()
		exit sub	
	end if
	
	dim loan
    if request.form("loan")="" then
		Errmsg=Errmsg+"<br>"+"<li>�����������"
    	founderr=true
	elseif not isnumeric(request.form("loan")) then
		Errmsg=Errmsg+"<br>"+"<li>���������������"
    	founderr=true		
	elseif request.form("loan")<=0 then
		Errmsg=Errmsg+"<br>"+"<li>���������������"
    	founderr=true
	else
		loan=int(request.form("loan"))
		if loan>dkdu then
			Errmsg=Errmsg+"<br>"+"<li>������Ľ��������Ĵ������ö��"
			founderr=true
		elseif loan>chubei then
			Errmsg=Errmsg+"<br>"+"<li>���Ĵ�����̫���ˣ����д������㣬�������Ա��ϵ"
			founderr=true			
		end if 		
	end if		
			
	if founderr then
		call bank_err()		
    else
		set rs=server.createobject("adodb.recordset")
		sql="select * from bank where username='"&membername&"' "
		rs.open sql,conn1,1,3
        if rs.eof and rs.bof then
			Errmsg=Errmsg+"<br>"+"<li>���������˺Ų����ڣ������Ч���ӽ���"
			call bank_err()
 		else 
			rs("daikuang")=rs("daikuang")+loan
			rs("dkdate")=date()
			rs.update
			rs.close
			set rs=server.createobject("adodb.recordset")
			sql="select * from bankconfig"
			rs.open sql,conn1,1,3
			rs("chubei")=rs("chubei")-loan
			rs.update
			rs.close
			set rs=server.createobject("adodb.recordset")
			sql="select * from [user] where username='"&membername&"' "
			rs.open sql,conn,1,3
			rs("userwealth")=rs("userwealth")+loan
			rs.update
			rs.close
			
			if cint(log_setting(0))=1 and cint(log_setting(4))<>0 then
				if cint(log_setting(4))=1 or ( cint(log_setting(4))=2 and loan>=clng(log_setting(5)) ) then
					content="�����д��������Ϊ��"&loan&" Ԫ����������Ϊ��"&daitian&"��"
					call logs("����","�����д���",membername)
					sucmsg=sucmsg+"<br>"+"<li>���Ĳ�����Ϣ�Ѿ���¼�ڰ�"
				end if	
			end if
								
			sucmsg=sucmsg+"<br>"+"<li>���Ĵ��������Ѿ�������ˣ�������Ϊ��"&loan&"Ԫ����������Ϊ��"&daitian&"��"
			call bank_suc()               
		end if
	end if
end sub

'--------------------------------ת�ʳ���------------------------------
sub zmoney()
	if bankstate=0 then			'������ ��� 2002.11.30
		Errmsg=Errmsg+"<br>"+"<li>������ͣӪҵ���������Ա��ϵ"
		call bank_err()
		exit sub	
	end if
	if cint(bank_setting(3))=0 then			'������ ��� 2002.11.30
		Errmsg=Errmsg+"<br>"+"<li>����ת�ʷ����Ѿ���ͣ���������Ա��ϵ"
		call bank_err()
		exit sub	
	end if
	if lockac then           '������ ���
		Errmsg=Errmsg+"<br>"+"<li>�Բ��������ܽ��иò���"
		Errmsg=Errmsg+"<br>"+"<li>���������˻��Ѿ������ᣬ�������Ա��ϵ"
    	call bank_err()
		exit sub	
	end if
	if daikuang>0 and cint(bank_setting(5))=0 then		'������ ��� 2002.12.03
		Errmsg=Errmsg+"<br>"+"<li>���й涨����֮����ת�ʣ��������Ա��ϵ"
    	call bank_err()
		exit sub	
	end if
	
    if request.form("trans")="" or (not isnumeric(request.form("trans"))) then
		Errmsg=Errmsg+"<br>"+"<li>����ȷ����ת�ʽ��"
		founderr=true
	elseif request.form("trans")<=0 then
		Errmsg=Errmsg+"<br>"+"<li>ת�ʽ�����������"
		founderr=true
	end if		
    if trim(request.form("transto"))="" then
		Errmsg=Errmsg+"<br>"+"<li>��Ҫת�ʵ��ĸ��˺��ϣ�"
		founderr=true
	end if
	if founderr then
		call bank_err()
    else
	   	dim trans,transto
	   	trans=clng(request.form("trans"))
	   	transto=checkStr(trim(request.form("transto")))
	   
	   	set rs=server.createobject("adodb.recordset")
	   	set rs=conn.execute("select username from [user] where username='"&transto&"'")
	   	if rs.bof and rs.eof then
			Errmsg=Errmsg+"<br>"+"<li>��̳��û��["&transto&"]����û�!"
    		call bank_err()
			rs.close
			exit sub
		else
	   		transto=rs(0)
			rs.close
		end if
        if myarticle<100 then 
  Errmsg=Errmsg+""+"<li>����ת�ʷ���涨����������100�ſ���ת��"
  call bank_err()
  exit sub 
 end if
	   	sql="select * from bank where username='"&membername&"' "
	   	rs.open sql,conn1,1,3
        if rs.eof and rs.bof then
			Errmsg=Errmsg+"<br>"+"<li>���������˺Ų����ڣ������Ч���ӽ���"
			call bank_err()
 		else 
			if rs("savemoney") < trans then
				Errmsg=Errmsg+"<br>"+"<li>����������û����ô��Ĵ��"
				call bank_err()
			elseif rs("savemoney")<int(trans*(1+formatnumber(zhuangli)/100)+0.999999) then
				Errmsg=Errmsg+"<br>"+"<li>�������д��������֧��ת��������"
				call bank_err()		
			else
				rs("savemoney")=rs("savemoney")-int(trans*(1+formatnumber(zhuangli)/100)+0.999999)
				rs.update
				rs.close
				set rs=server.createobject("adodb.recordset")
				sql="select * from bank where username='"&transto&"' "
				rs.open sql,conn1,1,3
				if rs.eof then
					rs.addnew
					rs("username")=transto
					rs("bankuser_setting")="0,"&bank_setting(4)
				end if    
				rs("savemoney")=rs("savemoney")+trans
				rs.update
				rs.close
				set rs=server.createobject("adodb.recordset")
				sql="select * from bankconfig"
				rs.open sql,conn1,1,3
				rs("chubei")=rs("chubei")+int(trans*(formatnumber(zhuangli)/100)+0.999999)
				rs.update
				rs.close
				
				if cint(log_setting(0))=1 and cint(log_setting(4))<>0 then
					if cint(log_setting(4))=1 or (cint(log_setting(4))=2 and trans>=clng(log_setting(5))) then
						content="ת���˺ţ�<font color=blue>"&HTMLEncode(transto)&"</font>, ת�ʽ��Ϊ��"&trans&"Ԫ, ת�������ѣ�"&int(trans*(formatnumber(zhuangli)/100)+0.999999)&"Ԫ"
						call logs("����","�����˺�����",membername)
						sucmsg=sucmsg+"<br>"+"<li>���Ĳ�����Ϣ�Ѿ���¼�ڰ�"
					end if	
				end if
							
				sucmsg=sucmsg+"<br>"+"<li>����ת�������Ѿ��������"
				sucmsg=sucmsg+"<br>"+"<li>ת���˺ţ�<font color=blue>"&HTMLEncode(transto)&"</font>, ת�ʽ��Ϊ��"&trans&"Ԫ, ת�������ѣ�"&int(trans*(formatnumber(zhuangli)/100)+0.999999)&"Ԫ"
					dim sender,title,body,sql2
                 sender=membername
                 title="ת��֪ͨ!"
                  body=membername&"ͨ�������������е�ת��ϵͳ�� "&trans&" Ԫ���㣬��ע�����!"
                  sql2="insert into message(incept,sender,title,content,sendtime,flag,issend) values('"&HTMLEncode(transto)&"','"&sender&"','"&title&"','"&body&"',Now(),0,1)"
                  conn.Execute(sql2)
			call bank_suc()				   
			end if
		end if
	end if
end sub

'--------------------------------ȡ�����--------------------------------
sub qumoney()
	if bankstate=0 then			'������ ��� 2002.11.30
		Errmsg=Errmsg+"<br>"+"<li>������ͣӪҵ���������Ա��ϵ"
		call bank_err()
		exit sub	
	end if
	if cint(bank_setting(1))=0 then			'������ ��� 2002.11.30
		Errmsg=Errmsg+"<br>"+"<li>�����������Ѿ���ͣ���������Ա��ϵ"
		call bank_err()
		exit sub	
	end if
	if lockac then 			'������ ���
		Errmsg=Errmsg+"<br>"+"<li>�Բ��������ܽ��иò���"
		Errmsg=Errmsg+"<br>"+"<li>���������˻��Ѿ������ᣬ�������Ա��ϵ"
    	call bank_err()
		exit sub	
	end if
    if request.form("draw")="" then
		Errmsg=Errmsg+"<br>"+"<li>������ȡ����"
		founderr=true
	elseif not isnumeric(request.form("draw")) then
		Errmsg=Errmsg+"<br>"+"<li>ȡ�������������"
		founderr=true
	elseif request.form("draw")<=0 then
		Errmsg=Errmsg+"<br>"+"<li>ȡ�������������"
		founderr=true			
    else
		dim drawm
		drawm=int(request.form("draw"))	
		if drawm>chubei then
			Errmsg=Errmsg+"<br>"+"<li>����ȡ����̫���ˣ����д������㣬�������Ա��ϵ"
			founderr=true			
		end if 
	end if	
	if founderr then
		call bank_err()	
	else
		dim saveli
		set rs=server.createobject("adodb.recordset")
		sql="select * from bank where username='"&membername&"' "
		rs.open sql,conn1,1,3
        if rs.eof and rs.bof then
			Errmsg=Errmsg+"<br>"+"<li>���������˺Ų����ڣ������Ч���ӽ���"
			call bank_err()		
        elseif drawm>rs("savemoney") then
			Errmsg=Errmsg+"<br>"+"<li>����������û����ô��Ĵ��"
	    	call bank_err()	
        else
			if saveday<Chen_StartLixiDay then
				saveli=0
			else
				saveli=clng((smoney*(formatnumber(culi)/100))*saveday)
			end if
			rs("savemoney")=rs("savemoney")-drawm+saveli
			rs("date")=date()
			rs.update
			rs.close

			sql="select userwealth from [user] where username='"&membername&"' "
			rs.open sql,conn,1,3
			if not(rs.eof and rs.bof) then
				rs(0)=rs(0)+drawm
				rs.update
			end if
			rs.close
			
			sql="select chubei from bankconfig"
			rs.open sql,conn1,1,3
			rs(0)=rs(0)-drawm
			rs.update
			rs.close 
				   
			if cint(log_setting(0))=1 and cint(log_setting(4))<>0 then
				if cint(log_setting(4))=1 or (cint(log_setting(4))=2 and drawm>=clng(log_setting(5))) then
					content="��������ȡ"&drawm&"Ԫ"
					call logs("����","���",membername)
					sucmsg=sucmsg+"<br>"+"<li>���Ĳ�����Ϣ�Ѿ���¼�ڰ�"
				end if	
			end if
			
		   sucmsg=sucmsg+"<br>"+"<li>���Ĵ������Ϣ�Ѿ����������˺ţ����õ�����Ϣ�ǣ�"&saveli&"Ԫ" 					   
		   sucmsg=sucmsg+"<br>"+"<li>����ȡ�������Ѿ�������ˣ�ȡ����Ϊ��"&drawm&"Ԫ"
		   call bank_suc()				   
 end if
end if
end sub

'--------------------------------------������------------------------------------
sub savemoney()
	if bankstate=0 then			'������ ��� 2002.11.30
		Errmsg=Errmsg+"<br>"+"<li>������ͣӪҵ���������Ա��ϵ"
		call bank_err()
		exit sub	
	end if
	if cint(bank_setting(0))=0 then			'������ ��� 2002.11.30
		Errmsg=Errmsg+"<br>"+"<li>���д������Ѿ���ͣ���������Ա��ϵ"
		call bank_err()
		exit sub	
	end if
    if request.form("saving")="" or (not isnumeric(request.form("saving"))) then
		Errmsg=Errmsg+"<br>"+"<li>����ȷ��������"
    	call bank_err()
    else
		savem=clng(request.form("saving"))
		set rs=server.createobject("adodb.recordset")
		sql="select * from [user] where username='"&membername&"' "
		rs.open sql,conn,1,3
       if savem>rs("userwealth") then
			Errmsg=Errmsg+"<br>"+"<li>��û����ô���Ǯ"
			call bank_err()	   		
 	   elseif savem<0 then 
			Errmsg=Errmsg+"<br>"+"<li>����ȷ��������"
    		call bank_err()
   	   else 
           set rsbank=server.createobject("adodb.recordset")
           rssql = "select * from bank where username='"&membername&"' "
           rsbank.open rssql,conn1,1,3
           if rsbank.eof then
              rsbank.addnew
              rsbank("username")=membername
              rsbank("savemoney")=savem
              rsbank.update
              rsbank.close
                set rsbank=server.createobject("adodb.recordset")
                rssql = "select * from bankconfig"
                rsbank.open rssql,conn1,1,3
                rsbank("chubei")=rsbank("chubei")+savem
                rsbank.update
                rsbank.close
              rs("userwealth")=rs("userwealth")-savem
              rs.update
              rs.close
           else
				dim saveli
				if saveday<Chen_StartLixiDay then
					saveli=0
				else
					saveli=clng((smoney*(formatnumber(culi)/100))*saveday)
				end if
				rsbank("savemoney")=rsbank("savemoney")+savem+saveli
				rsbank("date")=date()
				rsbank.update
				rsbank.close
                set rsbank=server.createobject("adodb.recordset")
                rssql = "select * from bankconfig"
                rsbank.open rssql,conn1,1,3            
                rsbank("chubei")=rsbank("chubei")+savem
                rsbank.update
                rsbank.close
				rs("userwealth")=rs("userwealth")-savem
				rs.update
				rs.close
           end if
		   
			if cint(log_setting(0))=1 and cint(log_setting(4))<>0 then
				if cint(log_setting(4))=1 or (cint(log_setting(4))=2 and savem>=clng(log_setting(5))) then
					content="�����д���"&savem&"Ԫ"
					call logs("����","���",membername)
					sucmsg=sucmsg+"<br>"+"<li>���Ĳ�����Ϣ�Ѿ���¼�ڰ�"
				end if	
			end if
			sucmsg=sucmsg+"<br>"+"<li>���Ĵ������Ϣ�Ѿ����������˺ţ����õ�����Ϣ�ǣ�"&saveli&"Ԫ" 		   
		   	sucmsg=sucmsg+"<br>"+"<li>����Ǯ�ɹ��Ĵ������У����δ��룺"&savem&"Ԫ"
		   	call bank_suc()
		end if
	end if
end sub

'-------------------------------------------��̳TOP20��������-------------------------------------
sub forum_top20()
    set rs=server.createobject("adodb.recordset")
    sql="select * from [user] order by userWealth desc"
    rs.open sql,conn,1,3
%>

<table cellpadding=3 cellspacing=1 align=center class=tableborder1 style="width:97%">
<tr> 
	<th>�û���</th> 
	<th>Email</th>
	<th>OICQ</th>
	<th>��ҳ</th>
	<th>����Ϣ</th>
	<th>ע��ʱ��</th>
	<th>�ȼ�״̬</th>
	<th>��������</th>
	<th>���</th>
</tr>	
<%
dim i
       for i=1 to 20
%>
<tr>
<td align=center class=tablebody2><a href="dispuser.asp?name=<%=htmlencode(rs("username"))%>" target=_blank><%=rs("username")%></a></td>
<td align=center class=tablebody1><a href=mailto:<%=rs("useremail")%>><img border=0 src="pic/email.gif"></a></td>
<td align=center class=tablebody2> 
<%if rs("oicq")="" or isnull(rs("oicq")) then%>
û��       
<%else%>      
<a href=http://search.tencent.com/cgi-bin/friend/user_show_info?ln=<%=rs("oicq")%> target=_blank><img src="pic/oicq.gif" alt="�鿴 OICQ:<%=rs("oicq")%> ������" border=0 width=16 height=16></a>       
<%end if%>      
</td>      
<td align=center class=tablebody1>       
<%if rs("homepage")="" or isnull(rs("homepage")) then%>      
û��       
<%else%>      
<a href=<%=rs("homepage")%> target=_blank><img border=0 src=pic/homepage.gif></a>       
<%end if%>      
</td>      
<td align=center class=tablebody2><a href=usersms.asp?action=new&touser=<%=htmlencode(rs("username"))%> target=_blank><img src=pic/message.gif border=0></a></td>      
<td align=center class=tablebody1><%=rs("addDate")%></td>      
<td align=center class=tablebody2><%=rs("userclass")%><br>      
</td>      
<td align=center class=tablebody1><%=rs("article")%></td>      
<td align=center class=tablebody2><%=rs("userwealth")%></td>      
</tr>      
<%      
           rs.movenext      
         if rs.eof then      
             exit for      
         end if      
       next      
%>      
</table>

</td></tr>  
<tr><td align=center colspan=9 height=26 class=tablebody1><a href=Z_bank.asp>����������ҳ</a>&nbsp;&nbsp;&nbsp;&nbsp;<a href=javascript:history.go(-1)>������һҳ</a></td></tr>    
<%      
end sub       
     
'---------------------------------��ʮ�󴢻�����-----------------------------------      
sub bank_top20()      
    set rs=server.createobject("adodb.recordset")      
    sql="select * from bank order by savemoney desc"      
    rs.open sql,conn1,1,3      
%>      

<table cellpadding=3 cellspacing=1 align=center class=tableborder1 style="width:97%">    
<tr>       
<th>�ʻ�</th>      
<th>���</th>  
<th>�����Ϣ</th>
<th>�������</th>      
<th>����</th>  
<th>������Ϣ</th>    
<th>��������</th>
<th width="65">�����ʻ�</th>  
</tr>      
<%      
dim i,savedaytmp,daidaytmp    
       for i=1 to 20  
			savedaytmp=datediff("d",rs("date"),date())
			daidaytmp=datediff("d",rs("dkdate"),date())	       
%>      
			<tr>       
			<td class=tablebody2 align=center><a href="dispuser.asp?name=<%=rs("username")%>" target=_blank><%=rs("username")%></a></td>      
			<td class=tablebody1 align=center><%=rs("savemoney")%></td>     
			<td class=tablebody1 align=center><%if savedaytmp<Chen_StartLixiDay then %>0<%else%><font color="#0066FF"><%=clng((formatnumber(rs("savemoney"))*(formatnumber(culi)/100))*savedaytmp)%></font><%end if%></td> 
			<td class=tablebody2 align=center><%=rs("date")%></td> 
			<td class=tablebody1 align=center><%=rs("daikuang")%></td> 
			<td class=tablebody1 align=center><%if rs("daikuang")=0 then %>0<%else%><font color="#0066FF"><%=clng((formatnumber(rs("daikuang"))*(formatnumber(daili)/100))*daidaytmp)%></font><%end if%></td> 
			<td class=tablebody2 align=center><%=rs("dkdate")%></td>
			<td class=tablebody1 align=center><%if rs("lockac") then%><font color="#FF0000">��</font><%else%>��<%end if%></td>     
			</tr>      
<%      
           rs.movenext      
         if rs.eof then      
             exit for      
         end if      
       next      
%>      
</table>
<br>
</td></tr>  
<tr><td align=center colspan=8 height=26 class=tablebody1><a href=Z_bank.asp>����������ҳ</a>&nbsp;&nbsp;&nbsp;&nbsp;<a href=javascript:history.go(-1)>������һҳ</a></td></tr>      
<%      
end sub 

'----------------------------���л�������--------------------------
sub admin()       
      
if not master then   
	Errmsg=Errmsg+"<br>"+"<li>��û�н������л������õ�Ȩ�ޡ�"   
  	call bank_err()      
else      
%>      
<br>
<table cellpadding=3 cellspacing=1 align=center class=tableborder1 style="width:97%">      
<tr><th colspan=4 height=20 align=left>���й�������---���л�������</th></tr>  
	<tr>      
		<td align=center colspan=4 class=tablebody1 height=10></td>      
	</tr>    
<form name="form2" method="post" action="Z_bank.asp?menu=9">      
<tr>       
	<td align=right width="20%" class=tablebody2>������Ϣ��</td>      
	<td align=left width="30%" class=tablebody1><INPUT name=saveli value="<%=culi%>"> %</td>   
	<td align=right width="20%" class=tablebody2>������</td> 
	<td align=center width="30%" class=tablebody1>
	<input type=radio name="canCunkuan" value=1 <%if cint(bank_setting(0))=1 then%>checked<%end if%>>��ͨ&nbsp;
	<input type=radio name="canCunkuan" value=0 <%if cint(bank_setting(0))=0 then%>checked<%end if%>>�ر�&nbsp;
	</td> 
</tr>      
<tr>       
	<td align=right class=tablebody2>�������Ϣ��</td>      
	<td align=left class=tablebody1><INPUT name=daili value="<%=daili%>"> %</td>  
	<td align=right class=tablebody2>�������</td> 
	<td align=center class=tablebody1>
	<input type=radio name="canDaikuan" value=1 <%if cint(bank_setting(2))=1 then%>checked<%end if%>>��ͨ&nbsp;
	<input type=radio name="canDaikuan" value=0 <%if cint(bank_setting(2))=0 then%>checked<%end if%>>�ر�&nbsp;
	</td>    
</tr>      
<tr>       
	<td align=right class=tablebody2>���еĴ����ʽ�</td>      
	<td align=left class=tablebody1><INPUT name=chubei value="<%=chubei%>"> Ԫ</td>  
	<td align=right class=tablebody2>ȡ�����</td> 
	<td align=center class=tablebody1>
	<input type=radio name="canQukuan" value=1 <%if cint(bank_setting(1))=1 then%>checked<%end if%>>��ͨ&nbsp;
	<input type=radio name="canQukuan" value=0 <%if cint(bank_setting(1))=0 then%>checked<%end if%>>�ر�&nbsp;
	</td>      
</tr>      
<tr>       
	<td align=right class=tablebody2>ת�ʵ������ѣ�</td>      
	<td align=left class=tablebody1><INPUT name=zhuangli value="<%=zhuangli%>"> %</td> 
	<td align=right class=tablebody2>ת�ʷ���</td> 
	<td align=center class=tablebody1>
	<input type=radio name="canZhuanzang" value=1 <%if cint(bank_setting(3))=1 then%>checked<%end if%>>��ͨ&nbsp;
	<input type=radio name="canZhuanzang" value=0 <%if cint(bank_setting(3))=0 then%>checked<%end if%>>�ر�&nbsp;
	</td>	     
</tr>      
<tr>       
	<td align=right class=tablebody2>�����������</td>      
	<td align=left class=tablebody1><INPUT name=daitian value="<%=daitian%>"> ��</td>    
	<td align=right class=tablebody2><font color=red>����״̬��</font></td> 
	<td align=center class=tablebody1>
	<input type=radio name="bankstate" value=1 <%if cint(bankstate)=1 then%>checked<%end if%>><font color=red>Ӫҵ</font>&nbsp;
	<input type=radio name="bankstate" value=0 <%if cint(bankstate)=0 then%>checked<%end if%>><font color=red>�ر�</font>&nbsp; 
	</td> 	  
</tr> 
<tr>       
	<td align=right class=tablebody2>��Ϣ��ʼ������</td>      
	<td align=left class=tablebody1><INPUT name="StartLixiDay" value="<%=Chen_StartLixiDay%>"> ��</td>    
	<td align=right class=tablebody2>Ӫҵʱ��</td> 
	<td align=center class=tablebody1>
	<input type=radio name="BusinessHours" value=0 <%if cint(bank_setting(6))=0 then%>checked<%end if%>>ȫ���&nbsp;
	<input type=radio name="BusinessHours" value=1 <%if cint(bank_setting(6))=1 then%>checked<%end if%>>��ʱ&nbsp; 	
	<INPUT name="BusinessTimeSlice" value="<%=bank_setting(7)%>" size=10>
	</td> 	
</tr>   
<tr>       
	<td align=right class=tablebody2>Ĭ������ϵ����</td>      
	<td align=left class=tablebody1><INPUT name=creditvalue value="<%=bank_setting(4)%>"></td>    
	<td align=right class=tablebody2>�д����Ƿ�����ת��</td> 
	<td align=center class=tablebody1>
	<input type=radio name="candaikuantrans" value=1 <%if cint(bank_setting(5))=1 then%>checked<%end if%>>��&nbsp;
	<input type=radio name="candaikuantrans" value=0 <%if cint(bank_setting(5))=0 then%>checked<%end if%>>��&nbsp; 	
	</td> 	  
</tr>     
<tr>      
	<td align=center colspan=4 class=tablebody1><INPUT type=submit value=�޸� name=submit></td>      
</tr>      
</form> 
	<tr>      
		<td align=center colspan=4 class=tablebody1 height=10></td>      
	</tr> 
<form name="form0" method="post" action="Z_bank.asp?menu=12">	
<tr><th colspan=4 height=20 align=left>���й�������---������־���� (��¼���еĸ������)</th></tr>
	<tr>      
		<td align=center colspan=4 class=tablebody1 height=10></td>      
	</tr> 
<tr>       
	<td class=tablebody2><font color=red>���в����¼���</font></td>      
	<td class=tablebody1>
	<input type=radio name="logstate" value=1 <%if cint(log_setting(0))=1 then%>checked<%end if%>><font color=red>��¼</font>&nbsp;
	<input type=radio name="logstate" value=0 <%if cint(log_setting(0))=0 then%>checked<%end if%>><font color=red>����¼</font>&nbsp; 	
	</td>    
	<td class=tablebody2>��¼�����¼�(�������)��</td> 
	<td class=tablebody1>
	<input type=radio name="logfuwu" value=1 <%if cint(log_setting(1))=1 then%>checked<%end if%>>��&nbsp; 
	<input type=radio name="logfuwu" value=0 <%if cint(log_setting(1))=0 then%>checked<%end if%>>��&nbsp; 
	</td> 	  
</tr> 
<tr>       
	<td class=tablebody2>���������¼�������<br>(���������¼���<font color=red>0</font>)</td>      
	<td class=tablebody1>
	&nbsp;<input type=text name="logrecordcount" value=<%=log_setting(2)%>>
	</td>    
	<td class=tablebody2>��¼���й�������¼���</td> 
	<td class=tablebody1>
	<input type=radio name="logadmin" value=1 <%if cint(log_setting(3))=1 then%>checked<%end if%>>��&nbsp; 
	<input type=radio name="logadmin" value=0 <%if cint(log_setting(3))=0 then%>checked<%end if%>>��&nbsp; 
	</td> 	  
</tr>
<tr>       
	<td class=tablebody2>��¼�����¼�(�����¼�)��</td> 
	<td class=tablebody1 colspan=3>    
	<input type=radio name="loguse" value=1 <%if cint(log_setting(4))=1 then%>checked<%end if%>>ȫ��&nbsp; 
	<input type=radio name="loguse" value=2 <%if cint(log_setting(4))=2 then%>checked<%end if%>>������
	<input type=text name="logminmoney" size=10 value=<%=log_setting(5)%>> Ԫ�Ĳ���
	<input type=radio name="loguse" value=0 <%if cint(log_setting(4))=0 then%>checked<%end if%>>����¼ &nbsp;
	</td> 	  
</tr> 
<tr>   
	<td class=tablebody2>�û��ɷ�鿴�¼���¼��</td> 
	<td class=tablebody1 colspan=3>    
	<input type=radio name="loguselook" value=1 <%if cint(log_setting(6))=1 then%>checked<%end if%>>����&nbsp; 
	<input type=radio name="loguselook" value=0 <%if cint(log_setting(6))=0 then%>checked<%end if%>>������&nbsp; 
	</td> 	  
</tr> 
<tr>      
	<td align=center colspan=4 class=tablebody1><INPUT type=submit value=�޸� name=submit></td>      
</tr> 
</form>   	
	<tr>      
		<td align=center colspan=4 class=tablebody1 height=10></td>      
	</tr> 		 
<tr><th colspan=4 height=20 align=left>���й�������---���д������ (ֻ�г����������ڵ��ʻ����г�)</th></tr>    
	<tr>      
		<td align=center colspan=4 class=tablebody1 height=10></td>      
	</tr> 
<%      
    set rs=server.createobject("adodb.recordset")      
    sql="select * from bank where daikuang>0 and datediff('d',dkdate,date())>"&daitian
    rs.open sql,conn1,1,3      
%>      
<tr>      
	<th>�ʻ���</th>      
	<th>������</th>   
	<th>��������</th>    
	<th>��   ��</th>      
</tr>      
<%
if  rs.eof and rs.bof then
%>
	<tr>      
		<td align=center colspan=4 class=tablebody2>��ʱ��û�г����������ڵ��ʻ�</td>      
	</tr> 
<%	
else     
	do while not rs.eof      
%>      
	<tr>      
	<td align=center class=tablebody2><a href="dispuser.asp?name=<%=rs("username")%>" target=_blank><%=rs("username")%></a></td>      
	<td align=center class=tablebody1><%=rs("daikuang")%></td>     
	<td align=center class=tablebody1><%=datediff("d",rs("dkdate"),date())%></td>  
	<td align=center class=tablebody2><a href="Z_bank.asp?menu=10&username=<%=rs("username")%>">ǿ�ƻ���</a></td>      
	</tr>      
<%       
	rs.movenext      
	loop 
end if	
%> 
	<tr>      
		<td align=center colspan=4 class=tablebody1 height=10></td>      
	</tr>      
<tr height=22><th colspan=4 height=20 align=left>���й�������---���н�������</th></tr> 
	<tr>      
		<td align=center colspan=4 class=tablebody1 height=10></td>      
	</tr>
<form name="form1" method="post" action="Z_bank.asp?menu=11">      
<tr>      
<td align=right class=tablebody2>��������</td>      
<td align=left colspan=3 class=tablebody1><select name="name">      
    <option value="20" selected>���й���Ա</option>
    <option value="19">���й��</option>
    <option value="18">���г�������</option>
     <option value="17">���а���</option>
    <option value="1">���л�Ա</option>
     <option value="24">����VIP��Ա</option>
    <option value="23">��¼20�����ϵĻ�Ա</option>
  </select>  ����������Ա����д��<INPUT name=username></td></tr>      
<tr>      
<td align=right class=tablebody2>������</td>      
<td align=left colspan=3 class=tablebody1><INPUT name=money></td>      
</tr>      
<tr>      
<td align=right class=tablebody2>�������ɣ�</td>      
<td align=left colspan=3 class=tablebody1><INPUT name=liyou size="50"></td>      
</tr>      
<tr>      
<td align=center colspan=4 class=tablebody1><INPUT type=submit value=���� name=submit></td>      
</tr>      
</form> 
	<tr>      
		<td align=center colspan=4 class=tablebody1 height=10></td>      
	</tr>      
<tr  height=22><td align=center colspan=4  height=26 class=tablebody2><a href=Z_bank.asp>����������ҳ</a>&nbsp;&nbsp;&nbsp;&nbsp;<a href=javascript:history.go(-1)>������һҳ</a></td></tr>      
</table>  
<br>    
<%       
end if      
end sub 

sub admin1() 
	dim saveli 
	if not master then
		founderr=true
		Errmsg=Errmsg+"<br>"+"<li>��û�н������л������õ�Ȩ�ޡ�"      
	else
		if request.form("saveli")="" or (not isnumeric(request.form("saveli"))) then
			founderr=true
			Errmsg=Errmsg+"<br>"+"<li>�����Ϣ����Ϊ����"
		elseif request.form("saveli")<0 then
			founderr=true
			Errmsg=Errmsg+"<br>"+"<li>�����Ϣ����Ϊ����"
		end if	 
		if request.form("daili")="" or (not isnumeric(request.form("daili"))) then
			founderr=true
			Errmsg=Errmsg+"<br>"+"<li>������Ϣ����Ϊ����" 
		elseif request.form("daili")<0 then
			founderr=true
			Errmsg=Errmsg+"<br>"+"<li>������Ϣ����Ϊ����"		
		end if	
		if request.form("chubei")="" or (not isnumeric(request.form("chubei"))) then
			founderr=true
			Errmsg=Errmsg+"<br>"+"<li>���д����ʽ����Ϊ����" 
		elseif request.form("chubei")<0 then
			founderr=true
			Errmsg=Errmsg+"<br>"+"<li>���д����ʽ���Ϊ����"			
		end if	
		if request.form("zhuangli")="" or (not isnumeric(request.form("zhuangli"))) then			
			founderr=true
			Errmsg=Errmsg+"<br>"+"<li>ת�������ѱ���Ϊ����" 	
		elseif request.form("zhuangli")<0 then
			founderr=true
			Errmsg=Errmsg+"<br>"+"<li>ת�������Ѳ���Ϊ����"		
		end if
		if request.form("daitian")="" or (not isnumeric(request.form("daitian"))) then			
			founderr=true
			Errmsg=Errmsg+"<br>"+"<li>�������������Ϊ����" 	
		elseif request.form("daitian")<=0 or int(request.form("daitian"))-request.form("daitian")<>0 then
			founderr=true
			Errmsg=Errmsg+"<br>"+"<li>������������ܱ���Ϊ������"		
		end if
		if request.form("creditvalue")="" or (not isnumeric(request.form("creditvalue"))) then
			founderr=true
			Errmsg=Errmsg+"<br>"+"<li>����ϵ������Ϊ����"
		elseif request.form("creditvalue")<0 then
			founderr=true
			Errmsg=Errmsg+"<br>"+"<li>����ϵ������Ϊ����"		
		end if
		if request.form("StartLixiDay")="" or (not isnumeric(request.form("StartLixiDay"))) then
			founderr=true
			Errmsg=Errmsg+"<br>"+"<li>�������Ϣ��ʼ����������������Ĳ�������"
		elseif request.form("StartLixiDay")<0 then
			founderr=true
			Errmsg=Errmsg+"<br>"+"<li>��Ϣ��ʼ��������Ϊ����"		
		end if
		if request.form("BusinessTimeSlice")="" then
			founderr=true
			Errmsg=Errmsg+"<br>"+"<li>������Ӫҵ��ʼʱ��"
		else
			Chen_BusinessTimeSlice=split(request.form("BusinessTimeSlice"),"||")
			if ubound(Chen_BusinessTimeSlice)<>1 then
				founderr=true
				Errmsg=Errmsg+"<br>"+"<li>Ӫҵ��ʼʱ�����벻��ȷ����ʽΪ����ʼСʱ��||����Сʱ��"
			elseif not (isnumeric(Chen_BusinessTimeSlice(0)) and isnumeric(Chen_BusinessTimeSlice(1))) then
				founderr=true
				Errmsg=Errmsg+"<br>"+"<li>Ӫҵ��ʼʱ���'||'�������߱���������"
			elseif cint(Chen_BusinessTimeSlice(0))<0 or  cint(Chen_BusinessTimeSlice(0))>24 then
				founderr=true
				Errmsg=Errmsg+"<br>"+"<li>��ʼӪҵʱ�䷶Χ����ȷ��������0��24֮�������"
			elseif cint(Chen_BusinessTimeSlice(1))<0 or  cint(Chen_BusinessTimeSlice(1))>24 then
				founderr=true
				Errmsg=Errmsg+"<br>"+"<li>����Ӫҵʱ�䷶Χ����ȷ��������0��24֮�������"
			elseif cint(Chen_BusinessTimeSlice(1))<cint(Chen_BusinessTimeSlice(0)) then
				founderr=true
				Errmsg=Errmsg+"<br>"+"<li>��ʼӪҵʱ�䲻�ܴ��ڽ���Ӫҵʱ��"												
			end if			
		end if		
	end if	
	if founderr then      
		call bank_err()
	else
		saveli=formatnumber(request.form("saveli"))      
		daili=formatnumber(request.form("daili"))      
		chubei=int(request.form("chubei"))      
		zhuangli=formatnumber(request.form("zhuangli"))      
		daitian=int(request.form("daitian"))      
		set rs=server.createobject("adodb.recordset")      
		sql="select * from [bankconfig]"      
		rs.open sql,conn1,1,3      
		rs("savedayli")=saveli      
		rs("ddayli")=daili      
		rs("chubei")=chubei      
		rs("zzli")=zhuangli      
		rs("dkday")=daitian
		rs("state")=request.form("bankstate")        '������ ��� 2002.11.30
		rs("bank_setting")=request.Form("canCunkuan") & "," & request.Form("canQukuan")& "," & request.Form("canDaikuan")& "," & request.Form("canZhuanzang")& "," & request.form("creditvalue")& "," & request.form("candaikuantrans")& "," & request.form("BusinessHours")& "," & request.form("BusinessTimeSlice")
		  
		rs("StartLixiDay")=request.form("StartLixiDay")		'��ˮ��ɽ 2003.1.6
		rs.update      
		rs.close 
		    
		if cint(log_setting(0))=1 and cint(log_setting(3))=1 then
			content="�������л�������"
			call logs("����","���й�������",membername)
			sucmsg=sucmsg+"<br>"+"<li>���Ĳ�����Ϣ�Ѿ���¼�ڰ�"
		end if
				
	   sucmsg=sucmsg+"<br>"+"<li>���л�����Ϣ�޸����"
	   call bank_suc()		  
	end if      
end sub  
     
'-----------------------------����-------------------------------       
sub jiangli()       
	dim money,liyou,classid,rs2,sql2  
	if not master then
		founderr=true
		Errmsg=Errmsg+"<br>"+"<li>��û��ִ�����н�����Ȩ�ޡ�"      
	else	
		if request.form("money")="" or (not isnumeric(request.form("money"))) then
			founderr=true
			Errmsg=Errmsg+"<br>"+"<li>����������Ϊ����"	  
		elseif request.form("money")<0 then
			founderr=true
			Errmsg=Errmsg+"<br>"+"<li>��������Ϊ����"		
		end if
		if trim(request.form("liyou"))="" then
				founderr=true
				Errmsg=Errmsg+"<br>"+"<li>�����뽱������"	
		end if 
	end if	
	
	if founderr then
		call bank_err()
		exit sub
	end if	  
		dim classuser  '���������û�
		money=int(request.form("money"))      
		liyou=trim(request.form("liyou"))
		if request.form("username") = "" then      
			classid=int(request.form("name"))     
			if classid=20 then      
				set rs=server.createobject("adodb.recordset")      
				sql="select userwealth,username from [user] where usergroupid=1"      
				rs.open sql,conn,1,3      
				do while not rs.eof      
					rs(0)=rs(0)+money      
					rs.update      
						set rs2=server.createobject("adodb.recordset")      
						sql2="select * from [message]"      
						rs2.open sql2,conn,1,3      
						rs2.addnew      
						rs2("sender")="�����г�"      
						rs2("incept")=rs(1)      
						rs2("title")="�������и�������"+cstr(money)+"Ԫ"      
						rs2("content")="�������и�����"+cstr(money)+"Ԫ�������������ɣ�"+liyou       
						rs2("flag")=0      
						rs2("issend")=1      
						rs2.update      
						rs2.close      
                rs.movenext         
            loop      
         elseif classid=19 then      
                     set rs=server.createobject("adodb.recordset")      
            sql="select userwealth,username from [user] where usergroupid=8"      
            rs.open sql,conn,1,3      
            do while not rs.eof      
                rs(0)=rs(0)+money       
                rs.update      
					set rs2=server.createobject("adodb.recordset")      
					sql2="select * from [message]"      
					rs2.open sql2,conn,1,3      
					rs2.addnew      
					rs2("sender")="�����г�"      
					rs2("incept")=rs(1)      
					rs2("title")="�������и�������"+cstr(money)+"Ԫ"      
					rs2("content")="�������и�����"+cstr(money)+"Ԫ�������������ɣ�"+liyou       
					rs2("flag")=0      
					rs2("issend")=1      
					rs2.update      
					rs2.close      
                rs.movenext      
            loop  
            elseif classid=18 then      
                set rs=server.createobject("adodb.recordset")      
				sql="select userwealth,username from [user] where usergroupid=2"      
				rs.open sql,conn,1,3      
				do while not rs.eof      
                rs(0)=rs(0)+money       
                rs.update      
					set rs2=server.createobject("adodb.recordset")      
					sql2="select * from [message]"      
					rs2.open sql2,conn,1,3      
					rs2.addnew      
					rs2("sender")="�����г�"      
					rs2("incept")=rs(1)      
					rs2("title")="�������и�������"+cstr(money)+"Ԫ"      
					rs2("content")="�������и�����"+cstr(money)+"Ԫ�������������ɣ�"+liyou       
					rs2("flag")=0      
					rs2("issend")=1      
					rs2.update      
					rs2.close      
                rs.movenext      
            loop      
    elseif classid=17 then      
            set rs=server.createobject("adodb.recordset")      
            sql="select userwealth,username from [user] where usergroupid=3"      
            rs.open sql,conn,1,3      
            do while not rs.eof      
                rs(0)=rs(0)+money       
                rs.update      
					set rs2=server.createobject("adodb.recordset")      
					sql2="select * from [message]"      
					rs2.open sql2,conn,1,3      
					rs2.addnew      
					rs2("sender")="�����г�"      
					rs2("incept")=rs(1)      
					rs2("title")="�������и�������"+cstr(money)+"Ԫ"      
					rs2("content")="�������и�����"+cstr(money)+"Ԫ�������������ɣ�"+liyou      
					rs2("flag")=0      
					rs2("issend")=1      
					rs2.update      
					rs2.close      
                rs.movenext      
            loop      

             elseif classid=1 then      
                set rs=server.createobject("adodb.recordset")      
				sql="select userwealth,username from [user]"      
				rs.open sql,conn,1,3      
				do while not rs.eof      
					rs(0)=rs(0)+money      
					rs.update      
						set rs2=server.createobject("adodb.recordset")      
						sql2="select * from [message]"      
						rs2.open sql2,conn,1,3      
						rs2.addnew      
						rs2("sender")="�����г�"      
						rs2("incept")=rs(1)      
						rs2("title")="�������и�������"+cstr(money)+"Ԫ"      
						rs2("content")="�������и�����"+cstr(money)+"Ԫ�������������ɣ�"+liyou      
						rs2("flag")=0      
						rs2("issend")=1      
						rs2.update      
						rs2.close      
					rs.movenext      
            	loop
         elseif classid=23 then       '������¼20�����ϵĻ�Ա
                     set rs=server.createobject("adodb.recordset")      
            sql="select userwealth,username from [user] where logins>20"      
            rs.open sql,conn,1,3      
            do while not rs.eof      
                rs(0)=rs(0)+money       
                rs.update      
					set rs2=server.createobject("adodb.recordset")      
					sql2="select * from [message]"      
					rs2.open sql2,conn,1,3      
					rs2.addnew      
					rs2("sender")="�����г�"      
					rs2("incept")=rs(1)      
					rs2("title")="�������и�������"+cstr(money)+"Ԫ"      
					rs2("content")="�������и�����"+cstr(money)+"Ԫ�������������ɣ�"+liyou       
					rs2("flag")=0      
					rs2("issend")=1      
					rs2.update      
					rs2.close      
                rs.movenext      
            loop  
         elseif classid=24 then       '����VIP��Ա
            set rs=server.createobject("adodb.recordset")      
            sql="select userwealth,username from [user] where vip=1"      
            rs.open sql,conn,1,3      
            do while not rs.eof      
                rs(0)=rs(0)+money       
                rs.update      
					set rs2=server.createobject("adodb.recordset")      
					sql2="select * from [message]"      
					rs2.open sql2,conn,1,3      
					rs2.addnew      
					rs2("sender")="�����г�"      
					rs2("incept")=rs(1)      
					rs2("title")="�������и�������"+cstr(money)+"Ԫ"      
					rs2("content")="�������и�����"+cstr(money)+"Ԫ�������������ɣ�"+liyou       
					rs2("flag")=0      
					rs2("issend")=1      
					rs2.update      
					rs2.close      
                rs.movenext      
            loop  
             end if      
		else      
			
        	classuser=checkStr(trim(request.form("username")))
        	set rs=server.createobject("adodb.recordset")      
     		sql="select userwealth,username from [user] where username='"&classuser&"'"      
     		rs.open sql,conn,1,3 
			if rs.bof and rs.eof then
				Errmsg=Errmsg+"<br>"+"<li>��̳��û��["&classuser&"]����û�"
				call bank_err()
				rs.close
				exit sub
			end if					     
        	rs(0)=rs(0)+money       
        	rs.update      
				set rs2=server.createobject("adodb.recordset")      
				sql2="select * from [message]"      
				rs2.open sql2,conn,1,3      
				rs2.addnew      
				rs2("sender")="�����г�"      
				rs2("incept")=rs(1)      
				rs2("title")="�������и�������"+cstr(money)+"Ԫ"      
				rs2("content")="�������и�����"+cstr(money)+"Ԫ�������������ɣ�"+liyou  
				rs2("flag")=0      
				rs2("issend")=1      
				rs2.update      
				rs2.close      
end if      
rs.close      
     set rs=server.createobject("adodb.recordset")      
     sql="select * from [bbsnews]"      
     rs.open sql,conn,1,3      
        rs.addnew      
        rs("boardid")=0      
     	if  classid=20 then      
        	rs("title")="���н������й���Ա"  
			content="�������й���Ա"    
     	elseif classid=19 then      
        	rs("title")="���н������й��" 
			content="�������й��"      
        elseif classid=18 then      
        	rs("title")="���н������г�������"  
			content="�������г�������"       
        elseif classid=17 then      
        	rs("title")="���н������а���"      
			content="�������а���"   
        elseif classid=1 then      
        	rs("title")="���н������л�Ա"      
			content="�������л�Ա"   
        else
			classuser=checkStr(trim(request.form("username")))
        	rs("title")="���н����û� "&classuser     
			content="�����û�" &classuser		
		end if      
        rs("content")="�������ɣ�"+liyou      
        rs("username")="�����г�"      
        rs.update      
        rs.close
		
		if cint(log_setting(0))=1 and cint(log_setting(3))=1 then
			content=content&" " &money&" Ԫ���������ɣ�"+liyou
			call logs("����","���й�������",membername)
			sucmsg=sucmsg+"<br>"+"<li>���Ĳ�����Ϣ�Ѿ���¼�ڰ�"
		end if
				  
	   sucmsg=sucmsg+"<br>"+"<li>���Ѿ�����˽�����Ա�Ĳ������뷵��"
	   call bank_suc()       
end sub 

sub savelogsetting() 
'������ 2002.12.01   �����¼�����
'logstate logfuwu logrecordcount logadmin loguse logminmoney  loguselook
	if not master then
		founderr=true
		Errmsg=Errmsg+"<br>"+"<li>��û�н������л������õ�Ȩ�ޡ�"      
	else
		if request.form("logrecordcount")="" or (not isnumeric(request.form("logrecordcount"))) then
			founderr=true
			Errmsg=Errmsg+"<br>"+"<li>���������¼���������Ϊ����"
		elseif request.form("logrecordcount")<0 then
			founderr=true
			Errmsg=Errmsg+"<br>"+"<li>���������¼���������Ϊ����"			
		end if	 
		if request.form("logminmoney")="" or (not isnumeric(request.form("logminmoney"))) then
			founderr=true
			Errmsg=Errmsg+"<br>"+"<li>��¼�����¼�(�����¼�)�����ƽ�����Ϊ����" 
		elseif request.form("logminmoney")<0 then
			founderr=true
			Errmsg=Errmsg+"<br>"+"<li>��¼�����¼�(�����¼�)�����ƽ���Ϊ����"		
		end if	
	end if	
	if founderr then      
		call bank_err()
	else
		dim log_settings
		log_settings=request.Form("logstate") & "," & request.Form("logfuwu")& "," & request.Form("logrecordcount")& "," & request.Form("logadmin")& "," & request.Form("loguse")& "," & request.Form("logminmoney")& "," & request.Form("loguselook")
		conn1.execute("update bankconfig set log_setting='"&log_settings&"'")
		
		if cint(log_setting(0))=1 and cint(log_setting(3))=1 then
			content="�޸��¼�����"
			call logs("����","���й�������",membername)
			sucmsg=sucmsg+"<br>"+"<li>���Ĳ�����Ϣ�Ѿ���¼�ڰ�"
		end if
				
		sucmsg=sucmsg+"<br>"+"<li>�����¼�������ɣ��뷵�ؽ�����������"
		call bank_suc()		  
	end if      
end sub

'----------------------------��������¼�------------------------------------------------
sub banklog()
'������ 2002.12.01   �鿴�¼���¼
	if (not master) and cint(log_setting(6))=0 then   
		Errmsg=Errmsg+"<br>"+"<li>��û����������¼���¼��Ȩ�ޣ��������Ա��ϵ"   
		call bank_err()      
		exit sub
	end if
	
	if clng(log_setting(2))<>0 then        '���ָ���������µ�N���¼�����ɾ��������¼���¼
		conn1.execute("delete from log where id not in (select top "&clng(log_setting(2))&" id from log order by id desc)")
	end if
		      
	if request("action")="dellog" then
		call batch()
	else
		call logeven()
	end if
	if founderr then call bank_err()
end sub
sub logeven()
'������ 2002.12.01   ��ʾ�¼�		
	dim endpage
	dim totalrec
	dim n
	dim currentpage,page_count,Pcount
		
	currentPage=request("page")
	if currentpage="" or not isInteger(currentpage) then
		currentpage=1
	else
		currentpage=clng(currentpage)
	end if	
	if master then
		response.write "<div align=center>����Ա��������ʱ���л�������״̬</div>"
		response.write "<div align=center>����״̬�µ�������߲鿴�ò����ߵ������¼�</div>"
	else
		response.write "<div align=center>������в����¼���¼</div>"
	end if
	response.write "<div align=center>�����������Ϳ��������ͬ���͵Ĳ����¼���¼</div>"
%>	
	<form action=Z_bank.asp?menu=13&action=dellog method=post name=even>
	<table cellpadding=3 cellspacing=1 align=center class=tableborder1 style="width:97%"> 
		<tr>
			<th width=5% height=25>����</th>
			<th width=15%>����</th>
			<th width=50% id=tabletitlelink>�¼�����(<a href=Z_bank.asp?menu=13 title=����鿴ȫ���¼�>�鿴ȫ���¼�</a>)</th>
			<th width=20% id=tabletitlelink><a href=Z_bank.asp?menu=13&action=batch&page=<%=currentpage%> title=����л�������״̬>����ʱ��</a></th>
			<th width=10%>������</th> 
		</tr>
<%
	set rs=server.createobject("adodb.recordset")
	if request("reaction")="����" then
		sql="select * from log where class='����' order by DateAndTime desc"
	elseif request("reaction")="����" then
		sql="select * from log where class='����' order by DateAndTime desc"
	elseif request("reaction")="����" then	
		sql="select * from log where class='����' order by DateAndTime desc"
	elseif request("reaction")="������" and trim(request("name"))<>"" then	
		sql="select * from log where UserName='"&checkStr(trim(request("name")))&"' order by DateAndTime desc"		
	else		
		sql="select * from log order by DateAndTime desc" 
	end if	
	'response.Write(sql)
	rs.open sql,conn1,1,1
	if rs.bof and rs.eof then
		response.write "<tr><td class=tablebody1 colspan=5 height=25>��ʱû���κ��¼�</td></tr></table><br>"
	else
		rs.PageSize = Forum_Setting(11)
		rs.AbsolutePage=currentpage
		page_count=0
      	totalrec=rs.recordcount
		while (not rs.eof) and (not page_count = rs.PageSize)
			response.write "<tr>"
			response.write "<td class=tablebody1 align=center height=24><a href=Z_bank.asp?menu=13&reaction="&rs("class")&" title=""����鿴����["&rs("class")&"]�¼�"">"&rs("class")&"</a></td>"
			response.write "<td class=tablebody1>"&htmlencode(rs("title"))&"</td>"
			response.write "<td class=tablebody1>"&rs("content")&"</td>"
			response.write "<td class=tablebody1>"
			if request("action")="batch" and master then
				response.write "<input type=checkbox name=lid value="&rs("id")&">"
			end if
			response.write rs("DateAndTime")
			response.write "</td>"
			response.write "<td align=center class=tablebody1>"
			if master then
				if request("action")="batch" then
					response.write "<a href=Z_bank.asp?menu=13&reaction=������&name="&htmlencode(rs("UserName"))&" title=""������IP��"&rs("IP")&"[����鿴�ò����ߵ����в�����¼]"">"
				else
					response.write "<a href=dispuser.asp?name="&htmlencode(rs("UserName"))&" target=_blank title=""������IP��"&rs("IP")&"[����鿴�����ߵ�����]"">"			
				end if
			else
				response.write "<a href=dispuser.asp?name="&htmlencode(rs("UserName"))&" target=_blank>"
			end if
			response.write htmlencode(rs("UserName"))&"</a></td>"	
			response.write "</tr>"
			page_count = page_count + 1
			rs.movenext
		wend
								
		if request("action")="batch" and master then
			response.write "<tr><td class=tablebody2 colspan=5>��ѡ��Ҫɾ�����¼���<input type=checkbox name=chkall value=on onclick=""CheckAll(this.form)"">ȫѡ <input type=submit name=Submit value=ִ��  onclick=""{if(confirm('��ȷ��ִ�д˲�����?')){this.document.even.submit();return true;}return false;}""></td></tr>"
		end if
		response.write "</table>"
		
		if totalrec mod Forum_Setting(11)=0 then
				Pcount= totalrec \ Forum_Setting(11)
		else
				Pcount= totalrec \ Forum_Setting(11)+1
		end if
		response.write "<table border=0 cellpadding=0 cellspacing=3 width=""97%"" align=center>"
		response.write "<tr><td valign=middle nowrap>"
		response.write "ҳ�Σ�<b>"&currentpage&"</b>/<b>"&Pcount&"</b>ҳ"
		response.write "&nbsp;ÿҳ<b>"&Forum_Setting(11)&"</b> ����<b>"&totalrec&"</b></td>"
		response.write "<td valign=middle nowrap align=right>��ҳ��"
		if request("reaction")="������" then
			call DispPageNum(currentpage,PCount,"""?menu=13&reaction=������&name="&request("name")&"&action="&request("action")&"&page=","""")
		else
			call DispPageNum(currentpage,PCount,"""?menu=13&action="&request("action")&"&page=","""")
		end if
		response.write "</td></tr></table>"
	end if	
	rs.close
	set rs=nothing		
end sub
sub batch()
	dim lid
	if not founduser then
		founderr=true
		Errmsg=Errmsg+"<br>"+"<li>���¼����в�����"
	end if
	if not master then
			founderr=true
			Errmsg=Errmsg+"<br>"+"<li>������ϵͳ����Ա�����ܹ���������־��"
	end if

	if request.form("lid")="" then
		founderr=true
		Errmsg=Errmsg+"<br>"+"<li>��ָ������¼���"
	else
		lid=replace(request.Form("lid"),"'","")
	end if
	if founderr then exit sub
	conn1.execute("delete from log where id in ("&lid&")")
	
	if cint(log_setting(0))=1 and cint(log_setting(3))=1 then
		content="ɾ��ָ���¼�"
		call logs("����","�����¼�����",membername)
		sucmsg=sucmsg+"<br>"+"<li>���Ĳ�����Ϣ�Ѿ���¼�ڰ�"
	end if	
	
	sucmsg="<li>ɾ��ָ���¼��ɹ�"
	call bank_suc()
end sub
%>