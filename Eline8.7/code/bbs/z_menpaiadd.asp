<!--#include file="conn.asp"-->
<!-- #include file="inc/const.asp" -->
<!-- #include file="z_menpaiconfig.asp" -->
<%
'========================������ʼ===========================
	dim menu,GroupName,chenbackasp,chentitle
	menu=request.querystring("menu")
	GroupName=checkstr(Trim(Request("GroupName")))
	if menu="" then menu=0
	
	chenbackasp="z_menpai.asp"
	chentitle="��̳����"
	stats="��̳���� "&htmlencode(GroupName)	
	call nav()
	
	select case menu
		case 2
			stats="��̳���� �������� "&htmlencode(GroupName)
		case 3
			stats="��̳���� �˳����� "&htmlencode(GroupName)
		case 4,5,6,9
			chenbackasp="z_menpaiadd.asp?GroupName="&htmlencode(GroupName)
			chentitle=htmlencode(GroupName)
			stats=htmlencode(GroupName)&" ���ɹ���"						
		case 7,8
			stats="��̳���� ����������"
	end select
	call head_var(3,0,"","")
			
	if not founduser then
		stats="��̳���� ������Ϣ"
		Errmsg=Errmsg+"<br>"+"<li>��û�н�����̳���ɵ�Ȩ�ޣ����ȵ�¼����ͬ����Ա��ϵ��"
		call dvbbs_error()
	else
		'��ȡָ�����ɵ�ֵ
		dim Menpai_setting,Zangmen,ZangmenStr,Addtime,States 
		dim CanManage		'�ܷ�����ɣ����뱾�����Ų��ܹ�����
		CanManage=false 
		
		if GroupName<>"" and (not isnull(GroupName)) then
			sql="select * from [GroupName] where groupname='"&GroupName&"'"
			set rs=conn.execute(sql)
			if not(rs.eof and rs.bof) then
				GroupName=rs("GroupName")
				Zangmen=trim(rs("Zangmen"))
				Addtime=rs("Addtime")
				States=rs("State")
				Menpai_setting=split(rs("Menpai_setting"),"|")
				
				if Zangmen="" or isnull(Zangmen) then
					ZangmenStr="<font color=#FF6666>- ��ѡ�� -</font>"
				elseif Zangmen="-������·-" then
					ZangmenStr="<font color=#CC6666>������·</font>"
				else
					if Zangmen=membername then CanManage=true
					ZangmenStr="<a href=dispuser.asp?name="&htmlencode(Zangmen)&" title=""�鿴�������˵�����"" target=_blank>"&Zangmen&"</a>"
				end if				
			else
				GroupName=""
				Menpai_setting=split("-1|0|0|0|0|0|1|","|")
			end if	
			rs.close
		else
			GroupName=""
			Menpai_setting=split("-1|0|0|0|0|0|1|","|")
		end if
		
		dim wumenwupai			'���ɲ�����������ɵ��ӵ�����
		set rs=conn.execute("select GroupName from [GroupName] where state=1")  
		if rs.eof and rs.bof then 
			wumenwupai="��������"
		else
			wumenwupai=rs(0)
		end if
		rs.close
		
		dim mymenpai		'�Լ����ڵ�����	
		mymenpai=conn.execute("select UserGroup from [user] where username='"&membername&"'")(0)				 	
		
		select case menu
			case 1
				call menpai()		'��ʾ���ɵĻ������
			case 2
				call inmenpai()		'��������
			case 3  	
    			call exitmenpai()	'�˳����� 
    		case 4  						
				call manage() 		'�������� �� �޸���������ҳ�� 
			case 5 
		    	call clearmember()		'�����Ż�  
			case 6  
			   	call changegroup()	    '�����޸ģ����ɹ���
			case 7  
				call applynewmenpai()	'���������ɣ�д�����ݿ�
			case 8
    			call newmenpai()		'���������ɣ����� 
			case 9
    			call logout()			'ע������	 				
			case else
				call menpai()
		end select	
	end if
	if founderr then call menpai_err()
	call activeonline()
	call footer() 
'========================���������===========================

'---------------------------��ʾ��������----------------------
sub menpai()
	if GroupName="" then 
		Errmsg=Errmsg+"<br>"+"<li>��������ɲ�������ָ���������"
		founderr=true
		exit sub
	end if 
	dim TotalArticle,TotalWealth,TotalUsernum,CanJoinin,CanExitout
	TotalArticle=0:	TotalWealth=0:	TotalUsernum=0
	CanJoinin=true:		CanExitout=false
	'ͳ��ָ�����ɵ�����������Ǯ�� �� ���������������ţ� 
	set rs=conn.execute("select sum(article),sum(userwealth),count(*) from [user] where usergroup='"&GroupName&"'") 
	if rs(0)<>"" then TotalArticle=rs(0)
	if rs(1)<>"" then TotalWealth=rs(1)
	if rs(2)<>"" then TotalUsernum=rs(2)
	
	dim yourmsg
%>
<table cellspacing=1 align=center class=tableborder1 width=<%=Forum_body(12)%>>
	<tr>
		<th valign=middle align=center height=24 colspan=8 id=TableTitleLink>�������� | <a href=z_menpaimember.asp?Groupname=<%=Groupname%>>�鿴���ɵ���</a></th> 
	</tr> 
	<tr> 
		<td valign=middle align=center class=tabletitle2 height="22">��������</td> 
		<td valign=middle align=center class=tabletitle2 height="22">����</td> 
		<td valign=middle align=center class=tabletitle2 height="22">����</td> 
		<td valign=middle align=center class=tabletitle2 height="22">��������</td>
		<td valign=middle align=center class=tabletitle2 height="22">������</td>
		<td valign=middle align=center class=tabletitle2 height="22">��Ǯ��</td>
		<td valign=middle align=center class=tabletitle2 height="22">������</td>
		<td valign=middle align=center class=tabletitle2 height="22">��ǰ״̬</td>
	</tr> 
	<tr> 
		<td valign=middle align=center class=tablebody1 height="20"><font color="#FF9933"><font face=Wingdings color=blue>[</font> <%=GroupName%></font></td>   	
		<td valign=middle align=center class=tablebody1 height="20"><%=ZangmenStr%></td>
		<td valign=middle align=center class=tablebody1 height="20"><%=MenpaiAttri(Menpai_setting(0))%></td> 
		<td valign=middle align=center class=tablebody1 height="20"><%=Addtime%></td>
		<td valign=middle align=center class=tablebody1 height="20"><%=TotalUsernum%></td> 
		<td valign=middle align=center class=tablebody1 height="20"><%=TotalWealth%></td> 
		<td valign=middle align=center class=tablebody1 height="20"><%=TotalArticle%></td> 
		<td valign=middle align=center class=tablebody1 height="20"><%=MenpaiState(States)%></td>
	</tr> 
	<tr><td height=8 class=tablebody2 colspan=8></td></tr>
	<tr> 
		<td valign=middle align=center class=tabletitle2 height="22">���ɼ��</td>  
		<td valign=middle align=left class=tablebody1 height="22" colspan=7>&nbsp;<font face=Wingdings color=blue>v</font>&nbsp;<%=menpai_setting(7)%></td> 
	</tr> 
	<tr><td height=8 class=tablebody2 colspan=8></td></tr>
	<tr> 
		<td valign=middle align=center class=tabletitle2 height="22">�Ź�</td>  
		<td valign=middle align=left class=tablebody1 height="22" colspan=7><ui>
<%		select case Menpai_setting(0)
			case 0
				response.Write "<li>������<u>��������</u>�����������������˳�����"
			case 1
				response.Write "<li>������<u>��������</u>�����������������ܼ�����˳�����"
				response.Write "<li>���뱾��Ҫ����һ���Ľ�Ǯ���������������飬�����Ȼ�����"
				response.Write "<li>�˳�����Ҫ�۳����������ŷѣ�����Ҫ��һ�������������������"			
			case 2
				response.Write "<li>������<u>ħ��</u>(����Щ��ν��������������)�����������������ܼ�����˳�����"
				response.Write "<li>���뱾�����Ľ�Ǯ�����ӣ����ǻ��һ����������������������˳�ʱ��֮"
			case 3
				response.Write "<li><u>������</u>���������κ����ɣ�Ⱥ�߻��ӣ�ʲô�˶���"
		end select 
		if cint(Menpai_setting(0))<>3 then
			response.Write "<li>������Ȩ���������Ź�����߳����ɣ����߳����ɵ��û����Զ�����<u>������</u>"
		end if	
%>
		</ui></td> 
	</tr> 		
	<tr><td height=15 class=tablebody2 colspan=8></td></tr>
	<th height="22" colspan=8>���뱾��������</th>
	<tr>
		<td align=center height=21 width="15%" class=tabletitle2>��Ŀ</td> 
		<td align=center height=21 class=tabletitle2>��Ǯ</td> 
		<td align=center height=21 class=tabletitle2>����</td>
		<td align=center height=21 class=tabletitle2>����</td> 
		<td align=center height=21 class=tabletitle2>����</td>
		<td align=center height=21 class=tabletitle2>����</td> 
		<td align=center height=21 class=tablebody2>�û��ȼ�</td>
		<td align=center height=21 class=tablebody2>����</td>
	</tr>	
	<tr>
		<td align=center height=21 class=tablebody2>����Ҫ��</td>
		<td align=center height=21 class=tablebody1><%=menpai_setting(1)%></td>
		<td align=center height=21 class=tablebody1><%=menpai_setting(2)%></td>
		<td align=center height=21 class=tablebody1><%=menpai_setting(3)%></td>
		<td align=center height=21 class=tablebody1><%=menpai_setting(4)%></td>
		<td align=center height=21 class=tablebody1><%=menpai_setting(5)%></td>
		<td align=center height=21 class=tablebody1><%=GetUserTitle(menpai_setting(6))%>(<%=menpai_setting(6)%>��)</td>
		<td align=center height=21 class=tablebody1><%=wumenwupai%></td>	
	</tr>
	<tr>
		<td align=center height=21 class=tablebody2>��������</td>
		<td align=center height=21 class=tablebody1><%=mymoney%></td>
		<td align=center height=21 class=tablebody1><%=myuserep%></td>
		<td align=center height=21 class=tablebody1><%=myusercp%></td>
		<td align=center height=21 class=tablebody1><%=mypower%></td>
		<td align=center height=21 class=tablebody1><%=myArticle%></td>
		<td align=center height=21 class=tablebody1><%=myClass%>(<%=GetUserClass(myClass)%>��)</td>
		<td align=center height=21 class=tablebody1><%=mymenpai%></td>		
	</tr>
	<tr>
		<td align=center height=21 class=tablebody2>�Ƿ�����</td>
		<td align=center height=21 class=tablebody1><%if int(mymoney)>=int(menpai_setting(1))   then%><font color=blue>��</font><%else%><%CanJoinin=false%><font color=red>��</font><%end if%></td>
		<td align=center height=21 class=tablebody1><%if int(myuserep)>=int(menpai_setting(2))  then%><font color=blue>��</font><%else%><%CanJoinin=false%><font color=red>��</font><%end if%></td>
		<td align=center height=21 class=tablebody1><%if int(myusercp)>=int(menpai_setting(3))  then%><font color=blue>��</font><%else%><%CanJoinin=false%><font color=red>��</font><%end if%></td>
		<td align=center height=21 class=tablebody1><%if int(mypower)>=int(menpai_setting(4))   then%><font color=blue>��</font><%else%><%CanJoinin=false%><font color=red>��</font><%end if%></td>
		<td align=center height=21 class=tablebody1><%if int(myArticle)>=int(menpai_setting(5)) then%><font color=blue>��</font><%else%><%CanJoinin=false%><font color=red>��</font><%end if%></td>
		<td align=center height=21 class=tablebody1><%if int(GetUserClass(myClass))>=int(menpai_setting(6)) then%><font color=blue>��</font><%else%><%CanJoinin=false%><font color=red>��</font><%end if%></td>
		<td align=center height=21 class=tablebody1><%if mymenpai=wumenwupai then%><font color=blue>��</font><%else%><%CanJoinin=false%><font color=red>��</font><%end if%></td>
	</tr>
	<tr><td height=8 class=tablebody2 colspan=8></td></tr>
  	<tr> 
    	<td valign=middle align=center class=tablebody2 height="22">���뱾�ɺ�</td> 
    	<td valign=middle align=center class=tablebody1 colspan=3><font color=navy> 
<%
	if cint(Menpai_setting(0))=0 then 			'���� �������� 
    	response.write "���Ľ�Ǯ-0������+0������+0������+0"                       
	elseif cint(Menpai_setting(0))=1 then		'���� ���� 
    	response.write "���Ľ�Ǯ-"&int(Menpai_setting(1)*Meipai_inout_rate(0))&"������+"&int(menpai_setting(2)*Meipai_inout_rate(1))&"������+"&int(Menpai_setting(3)*Meipai_inout_rate(2))&"������+"&int(Menpai_setting(4)*Meipai_inout_rate(3))&""
  	elseif cint(Menpai_setting(0))=2 then		'���� ħ�� 
    	response.write "���Ľ�Ǯ+"&int(Menpai_setting(1)*Meipai_inout_rate(4))&"������-"&int(menpai_setting(2)*Meipai_inout_rate(5))&"������-"&int(Menpai_setting(3)*Meipai_inout_rate(6))&"������-"&int(Menpai_setting(4)*Meipai_inout_rate(7))&""		
	else
		response.write "���Ľ�Ǯ-0������+0������+0������+0"
	end if
%>                           
  		</font></td>		
  		<td valign=middle align=center class=tablebody2 height="22" colspan=2>����״̬</td> 
  		<td valign=middle align=center class=tablebody1 height="22" colspan=2>
<%
	if mymenpai=wumenwupai then
		response.write "<font color=red>����û�м����κ�����Ӵ</font> "
	elseif mymenpai<>Groupname then
		response.write mymenpai &"��Ա"
		CanJoinin=false
	elseif CanManage then
		response.write "��������"
		CanJoinin=false
	else
		response.write "���ɵ���"
		CanJoinin=false
		CanExitout=true
	end if
%>
  		</td>                          
  	</tr> 
  	<tr> 
    	<td valign=middle align=center class=tablebody2 height="22">�˳����ɺ�</td> 
    	<td valign=middle align=center class=tablebody1 colspan=3><font color=navy> 
<%
	if cint(Menpai_setting(0))=0 then 			'�˳� ��������
    	response.write "���Ľ�Ǯ-0������+0������+0������+0"                       
	elseif cint(Menpai_setting(0))=1 then		'�˳� ����
    	response.write "���Ľ�Ǯ-"&int(Menpai_setting(1)*Meipai_inout_rate(8))&"������-"&int(menpai_setting(2)*Meipai_inout_rate(9))&"������-"&int(Menpai_setting(3)*Meipai_inout_rate(10))&"������-"&int(Menpai_setting(4)*Meipai_inout_rate(11))&""
  	elseif cint(Menpai_setting(0))=2 then		'�˳� ħ��
    	response.write "���Ľ�Ǯ-"&int(Menpai_setting(1)*Meipai_inout_rate(12))&"������+"&int(menpai_setting(2)*Meipai_inout_rate(13))&"������+"&int(Menpai_setting(3)*Meipai_inout_rate(14))&"������+"&int(Menpai_setting(4)*Meipai_inout_rate(15))&""		
	else
		response.write "���Ľ�Ǯ-0������+0������+0������+0"
	end if
%>                           
  		</font></td>		
  		<td valign=middle align=center class=tablebody2 height="22" colspan=2>���Ĳ���</td> 
  		<td valign=middle align=center class=tablebody1 height="22" colspan=2>
<%
	if CanExitout then
		yourmsg="<font color=red>���һ�ӭ���ɵ��� "&membername&" ����</font>"                            
		response.Write "<a href=""z_menpaimember.asp?action=�����б�&groupname="&htmlencode(groupname)&"""><font color=blue>�����б�</font></a> | <a href=""z_menpaiadd.asp?menu=3&groupname="&htmlencode(groupname)&""" onclick=""{if(confirm('��ȷ���˳� "&groupname&" ��?')){return true;}return false;}""><font color=blue>�˳�����</font></a> | <a href=z_menpai.asp><font color=blue>������ҳ</font></a>"  
	end if
		                                                                                                                                                                                                  
	if CanManage then                                                                                                                                                                                                                                  
		yourmsg="<font color=red>��ӭ���Ŵ�ݹ���</font>"
		response.Write "<a href=""z_menpaiadd.asp?menu=4&groupname="&htmlencode(groupname)&"""><font color=blue>������</font></a> | <a href=z_menpai.asp><font color=blue>������ҳ</font></a>"
	end if                                                                                                                                                                                                                                    
                                                                                                                                                                                                                                  
	if (not Canjoinin) and (not CanExitout) and (not CanManage) then
		if isZangmen then
			yourmsg="<font color=red>���һ�ӭ "&mymenpai&" ������ "&membername&" ���ٱ���</font>" 
		elseif mymenpai<>wumenwupai then	
			yourmsg="<font color=red>���һ�ӭ "&mymenpai&" ���� "&membername&" ���ٱ���</font>"
		else
			yourmsg="<font color=red>��������һ�����������㣬��ʱ���ܼ��뱾��Ӵ</font>"  
		end if	
		response.Write "<a href=""z_menpaiadd.asp?menu=2&groupname="&htmlencode(groupname)&""" onclick=""{if(confirm('��ȷ������ "&groupname&" ��?')){return true;}return false;}""><font color=blue>���뱾��</font></a> | <a href=z_menpai.asp><font color=blue>������ҳ</font></a>"
	elseif Canjoinin and (not CanManage) then
		if Groupname=wumenwupai then
			response.Write "<a href=z_menpai.asp><font color=blue>������ҳ</font></a>"
			yourmsg="<font color=red>�����ն񣬴���ҪС��Ŷ</font>"
		else
			yourmsg="<font color=red>��ӭ�������� "&membername&" ���� "&Groupname&"����ӭ�����뱾��</font>"
			response.Write "<a href=""z_menpaiadd.asp?menu=2&groupname="&htmlencode(groupname)&""" onclick=""{if(confirm('��ȷ������ "&groupname&" ��?')){return true;}return false;}""><font color=blue>���뱾��</font></a> | <a href=z_menpai.asp><font color=blue>������ҳ</font></a>"
		end if
	end if                                                                                                                                                                                                                                 
%>                                                                                                                                                                                                                                   
  		</td>                          
  	</tr> 	                               
  	<tr>                                
    	<td valign=middle align=center height=22 class=tablebody1 colspan=8><%=yourmsg%></td>                                                                                                                                                                                                                                   
  	</tr>                                                                                                                                                                                                                                   
</table>                                                                                                                                                                                                                                   
<%                                                                                                                                                                                                                                   
end sub                                                                                                                                                                                                                                  
'-------------------------��������----------------------------
sub inmenpai()                                                                                                                                                                                                                                   

	if Groupname="" then 
		Errmsg=Errmsg+"<br>"+"<li>��������ɲ�������ָ���������"
		founderr=true
		exit sub		
	end if 	
	if mymenpai<>wumenwupai then
		Errmsg=Errmsg+"<br>"+"<li>���Ѿ�����һ�����ɵĵ����ˣ�Ҫ���뱾�������˳�ԭ��������"
		founderr=true
		exit sub
	elseif int(mymoney)<int(menpai_setting(1)) or int(myuserep)<int(menpai_setting(2)) or int(myusercp)<int(menpai_setting(3)) or int(mypower)<int(menpai_setting(4)) or int(myArticle)<int(menpai_setting(5)) or int(GetUserClass(myClass))<int(menpai_setting(6)) then
		Errmsg=Errmsg+"<br>"+"<li>��������һ��������������뱾�ɵ�Ҫ��"
		founderr=true
		exit sub
	end if 		
	                                                                                                                                                                                                                              
    Set rs= Server.CreateObject("ADODB.Recordset")                                                                                                                                                                                                                           
	sql="select usergroup,userwealth,userep,usercp,userpower from [user] where username='"&membername&"'"                                                                                                                                                                                                                                   
	rs.open sql,conn,1,3                                                                                                                                                                                                                                  
	if rs.eof and rs.bof then
		Errmsg=Errmsg+"<br>"+"<li>����ô���ܵ���������"
		founderr=true
		exit sub
	else
		sucmsg=sucmsg+"<br>"+"<li>���ɹ������� "&groupname		
		if cint(menpai_setting(0))=0 then       	'�������� 
			rs(0)=Groupname
			rs.update
			sucmsg=sucmsg+"<br>"+"<li>���Ľ�Ǯ-0������+0������+0������+0"
		elseif cint(menpai_setting(0))=1 then       '����
			rs(0)=Groupname
			rs(1)=rs(1)-int(menpai_setting(1)*Meipai_inout_rate(0))
			rs(2)=rs(2)+int(menpai_setting(2)*Meipai_inout_rate(1))
			rs(3)=rs(3)+int(menpai_setting(3)*Meipai_inout_rate(2))
			rs(4)=rs(4)+int(menpai_setting(4)*Meipai_inout_rate(3))
			rs.update
			sucmsg=sucmsg+"<br>"+"<li>���Ľ�Ǯ-"&int(Menpai_setting(1)*Meipai_inout_rate(0))&"������+"&int(menpai_setting(2)*Meipai_inout_rate(1))&"������+"&int(Menpai_setting(3)*Meipai_inout_rate(2))&"������+"&int(Menpai_setting(4)*Meipai_inout_rate(3))
		elseif cint(menpai_setting(0))=2 then       'ħ��           
			rs(0)=Groupname
			rs(1)=rs(1)+int(menpai_setting(1)*Meipai_inout_rate(4))        
			rs(2)=rs(2)-int(menpai_setting(2)*Meipai_inout_rate(5))   
			rs(3)=rs(3)-int(menpai_setting(3)*Meipai_inout_rate(6))
			rs(4)=rs(4)-int(menpai_setting(4)*Meipai_inout_rate(7))
			rs.update
			sucmsg=sucmsg+"<br>"+"<li>���Ľ�Ǯ+"&int(Menpai_setting(1)*Meipai_inout_rate(4))&"������-"&int(menpai_setting(2)*Meipai_inout_rate(5))&"������-"&int(Menpai_setting(3)*Meipai_inout_rate(6))&"������-"&int(Menpai_setting(4)*Meipai_inout_rate(7))&""
		else
			Errmsg=Errmsg+"<br>"+"<li>��������ɲ�������ָ����ص��������Բ���"
			founderr=true
			exit sub		
		end if
	end if
	
	call menpai_suc()
end sub                                                                                                                                                                                                                                    	  
'---------------------------�˳�����------------------------------
sub exitmenpai()                                                                                                                                                                                                                                   

	if Groupname="" then 
		Errmsg=Errmsg+"<br>"+"<li>��������ɲ�������ָ���������"
		founderr=true
		exit sub
	end if
	
	if cint(Menpai_setting(0))=0 then 
	
	elseif cint(Menpai_setting(0))=1 then 
		if int(menpai_setting(2)*Meipai_inout_rate(9))>myuserep or int(Menpai_setting(3)*Meipai_inout_rate(10))>myusercp or int(Menpai_setting(4)*Meipai_inout_rate(11))>mypower then                                                                                                                                                                                                      
			Errmsg=Errmsg+"<br>"+"<li>����û���㹻���ʱ����뱾��"
			founderr=true
		end if
	elseif cint(Menpai_setting(0))=2 then
		if int(Menpai_setting(1)*Meipai_inout_rate(12))>mymoney then
			Errmsg=Errmsg+"<br>"+"<li>����û���㹻���ʱ����뱾��"
		end if
	else
		Errmsg=Errmsg+"<br>"+"<li>��������ɲ�������ָ�������������"
		founderr=true
	end if
		
	if founderr then exit sub
		
	Groupname=wumenwupai  
                                                                                                                                                                                                                   
	Set rs= Server.CreateObject("ADODB.Recordset")                                                                                                                                                                                                                                   
	sql="select usergroup,userwealth,userep,usercp,userpower from [user] where username='"&membername&"'"                                                                                                                                                                                                                                   
	rs.open sql,conn,1,3                                                                                                                                                                                                                                  
                                                                                                                                                                                                                   
	if cint(Menpai_setting(0))=0 then   		'��������                                                                                                                                                                                                                        
		rs("usergroup")=groupname                                                                                                                                                                                                                           
		rs.update     
		sucmsg=sucmsg+"<br>"+"<li>���Ľ�Ǯ+0������-0������-0������-0"                                                                                                                                                                                                                      
	elseif cint(Menpai_setting(0))=1 then     	'����	                                                                                                                                                                                                                              
		rs(0)=Groupname
		rs(1)=rs(1)-int(menpai_setting(1)*Meipai_inout_rate(8))
		rs(2)=rs(2)-int(menpai_setting(2)*Meipai_inout_rate(9))
		rs(3)=rs(3)-int(menpai_setting(3)*Meipai_inout_rate(10))
		rs(4)=rs(4)-int(menpai_setting(4)*Meipai_inout_rate(11))                                                                                                                                                                                                                         
		rs.update 
		sucmsg=sucmsg+"<br>"+"<li>���Ľ�Ǯ+"&int(Menpai_setting(1)*Meipai_inout_rate(8))&"������-"&int(menpai_setting(2)*Meipai_inout_rate(9))&"������-"&int(Menpai_setting(3)*Meipai_inout_rate(10))&"������-"&int(Menpai_setting(4)*Meipai_inout_rate(11))&""                                                                                                                                                                                                                         
	elseif cint(Menpai_setting(0))=2 then   	'ħ��
		rs(0)=Groupname
		rs(1)=rs(1)-int(menpai_setting(1)*Meipai_inout_rate(12))        
		rs(2)=rs(2)+int(menpai_setting(2)*Meipai_inout_rate(13))   
		rs(3)=rs(3)+int(menpai_setting(3)*Meipai_inout_rate(14))
		rs(4)=rs(4)+int(menpai_setting(4)*Meipai_inout_rate(15))                                                                                                                                                                                                                         
		rs.update  
		sucmsg=sucmsg+"<br>"+"<li>���Ľ�Ǯ-"&int(Menpai_setting(1)*Meipai_inout_rate(12))&"������+"&int(menpai_setting(2)*Meipai_inout_rate(13))&"������+"&int(Menpai_setting(3)*Meipai_inout_rate(14))&"������+"&int(Menpai_setting(4)*Meipai_inout_rate(15))&""                                                                                                                                                                                                                        
	end if  
	rs.close
	set rs=nothing                                                                                                                                                                                                                        
	call menpai_suc()                                                                                                                                                                                                                                
end sub                                                                                                                                                                                                                                   
'---------------------------���ɹ���====����ҳ��------------------------------
sub manage()                                                                                                                                                                                           
	if not CanManage then
		Errmsg=Errmsg+"<br>"+"<li>�����Ǳ������ţ�û�й����ɵ�Ȩ��"
		founderr=true
		exit sub	
	end if
	if Groupname="" then 
		Errmsg=Errmsg+"<br>"+"<li>��������ɲ�������ָ��������ɲ���"
		founderr=true
		exit sub
	end if	 

%>   
<table cellspacing=1 align=center class=tableborder1 width=<%=Forum_body(12)%>>                                                                                                                                                                              
  	<tr>                                                                                                                                                                                           
    	<td valign=middle align=center height=24 class=tablebody2 colspan="4">���ɹ���</td>                                                                                                                                                                                          
  	</tr> 
	<form name="form" method="post" action="z_menpaiadd.asp?menu=6">                                                                                                                                                                                            
  	<tr>                                                                                                                                                                                           
    	<th valign=middle align=left height=24 colspan="4">&nbsp;��&nbsp; ���ɻ�������</th>                                                                                                                                                                                          
  	</tr> 
	<tr><td height=8 class=tablebody2 colspan="4"></td></tr>                                                                                                                                                                                        
    <tr>                                                                                                                                                                                          
		<td valign=middle align=left class=tablebody2 height="22"><b>������֪:</b></td>                                                                                                                                                                                         
		<td valign=middle align=left class=tablebody1 height="22" colspan="3">
		 	<li><font color="#FF9966">��ӭ <font color=blue><%=Groupname%></font> ������ <font color=red><%=membername%></font> �������ɹ���</font>
			<li>���ɵ����ս���Ȩ���ڹ���Ա
			<li>��������˲������Σ�����Ա��Ȩ�����������˵�ְ�� 
			<li>������������������ɵĻ������
		</td> 
  	</tr> 	   
	<tr><td height=8 class=tablebody2 colspan="4"></td></tr> 
  	<tr> 
		<td valign=middle align=left class=tablebody2 height="22"><b>��������:</b><br> ������������֮�����Ų����޸�</td> 
		<td valign=middle align=left class=tablebody1 height="22" colspan="3">&nbsp;<INPUT name=newGroupName size="20" maxlength="20" disabled value=<%=GroupName%>> <font color=red>*</font> (��Ҫ����12��Ӣ����ĸ��6������)<INPUT name=GroupName size="20" maxlength="20" type=hidden value=<%=GroupName%>></td>
  	</tr>
  	<tr> 
		<td valign=middle align=left class=tablebody2 height="22"><b>��������:</b><br> ������������֮�����Ų����޸�</td>
		<td valign=middle align=left class=tablebody1 height="22" colspan="3">
			<input type=radio name=newMenpaiAttri value=0 disabled <%if cint(menpai_setting(0))=0 then%>checked<%end if%>>��������&nbsp;
			<input type=radio name=newMenpaiAttri value=1 disabled <%if cint(menpai_setting(0))=1 then%>checked<%end if%>>����&nbsp;
			<input type=radio name=newMenpaiAttri value=2 disabled <%if cint(menpai_setting(0))=2 then%>checked<%end if%>>ħ��
			<input type=hidden name=MenpaiAttri value=<%=cint(menpai_setting(0))%>>
		</td>
  	</tr>
	<tr><td height=8 class=tablebody2 colspan="4"></td></tr>
    <tr>
		<td valign=middle align=left class=tablebody2 height="22"><b>˵��:</b></td>                                                                                                                                                                                         
		<td valign=middle align=left class=tablebody1 height="22" colspan="3">
		 	<li>��Ǯ,����,����,����,������,�û��ȼ� �����ܳ��������ѵĵ�ǰֵ
			<li>��Ǯ,����,����,����,������,�û��ȼ� �Ǽ��뱾�ɵı����������Сֵ
			<li>��Ǯ,����,����,����,������,�û��ȼ� Ҳ�Ǽ�����˳���������/�۳����ֵʱ�Ļ���
			<li>�������ɲ�����������
			<li>���ſ��Ը���ʵ������޸��������޸ĺ�֮�� �ύ ����			
		</td> 
  	</tr> 
    <tr>                                                                                                                                                                                         
		<td valign=middle align=left class=tablebody2 height="22" width="25%"><b>��Ǯ:</b></td>                                                                                                                                                                                        
		<td valign=middle align=left class=tablebody1 height="22">&nbsp;<INPUT name=userwealth value=<%=menpai_setting(1)%> <%if cint(menpai_setting(0))=0 then%> disabled<%end if%>></td>                                                                                                                                                                                        
		<td valign=middle align=right class=tablebody2 height="22">���Ľ�Ǯ:&nbsp;</td>
		<td valign=middle align=left class=tablebody1 height="22" width="20%">&nbsp;<%=mymoney%></td>
  	</tr>                                                                                                                                                                                        
  	<tr>                                                                                                                                                                                        
		<td valign=middle align=left class=tablebody2 height="22"><b>����:</b></td>                                                                                                                                                                                        
		<td valign=middle align=left class=tablebody1 height="22">&nbsp;<INPUT name=userep value=<%=menpai_setting(2)%> <%if cint(menpai_setting(0))=0 then%> disabled<%end if%>></td>
		<td valign=middle align=right class=tablebody2 height="22">���ľ���:&nbsp;</td>                                                                                                                                                                                        
		<td valign=middle align=left class=tablebody1 height="22">&nbsp;<%=myusercp%></td>                                                                                                                                                                                        
  	</tr>                                                                                                                                                                                        
  	<tr>                                                                                                                                                                                        
		<td valign=middle align=left class=tablebody2 height="22"><b>����:</b></td>                                                                                                                                                                                        
		<td valign=middle align=left class=tablebody1 height="22">&nbsp;<INPUT name=usercp value=<%=menpai_setting(3)%> <%if cint(menpai_setting(0))=0 then%> disabled<%end if%>></td> 
		<td valign=middle align=right class=tablebody2 height="22">��������:&nbsp;</td>                                                                                                                                                                                       
		<td valign=middle align=left class=tablebody1 height="22">&nbsp;<%=myuserep%></td>                                                                                                                                                                                        
  	</tr>                                                                                                                                                                                        
  	<tr>                                                                                                                                                                                        
		<td valign=middle align=left class=tablebody2 height="22"><b>����:</b></td>                                                                                                                                                                                        
		<td valign=middle align=left class=tablebody1 height="22">&nbsp;<INPUT name=userpower value=<%=menpai_setting(4)%> <%if cint(menpai_setting(0))=0 then%> disabled<%end if%>></td> 
		<td valign=middle align=right class=tablebody2 height="22">��������:&nbsp;</td>                                                                                                                                                                                       
		<td valign=middle align=left class=tablebody1 height="22">&nbsp;<%=mypower%></td>                                                                                                                                                                                        
  	</tr>                                                                                                                                                                                        
  	<tr>                                                                                                                                                                                        
		<td valign=middle align=left class=tablebody2 height="22"><b>������:</b><br>���뱾�ɱ���ﵽ�ķ�����</td>                                                                                                                                                                                        
		<td valign=middle align=left class=tablebody1 height="22">&nbsp;<INPUT name=userarticle value=<%=menpai_setting(5)%> <%if cint(menpai_setting(0))=0 then%> disabled<%end if%>></td> 
		<td valign=middle align=right class=tablebody2 height="22">��������:&nbsp;</td>                                                                                                                                                                                       
		<td valign=middle align=left class=tablebody1 height="22">&nbsp;<%=myarticle%></td>                                                                                                                                                                                        
  	</tr>                                                                                                                                                                                        
  	<tr>                                                                                                                                                                                        
		<td valign=middle align=left class=tablebody2 height="22"><b>�û��ȼ�:</b><br>���뱾�ɱ���߱��ĵȼ�(1��<%=maxtitleid%>)</td>                                                                                                                                                                                        
		<td valign=middle align=left class=tablebody1 height="22">&nbsp;<INPUT name=userclass value=<%=menpai_setting(6)%> <%if cint(menpai_setting(0))=0 then%> disabled<%end if%>>&nbsp;<%=GetUserTitle(Menpai_setting(6))%>(<%=Menpai_setting(6)%>��)</td> 
		<td valign=middle align=right class=tablebody2 height="22">���ĵȼ�:&nbsp;</td>                                                                                                                                                                                     
		<td valign=middle align=left class=tablebody1 height="22">&nbsp;<%=myclass%>��<%=GetUserClass(myclass)%>����</td>                                                                                                                                                                                      
  	</tr>
	<tr>
		<td valign=middle align=center class=tablebody1 height=22 colspan="4"><font color="#CC9933">
		��������������������ɣ�������6��(��Ǯ,����,����,����,������,�û��ȼ�)������д</font>
		</td>
	</tr> 
	<tr><td height=8 class=tablebody2 colspan="4"></td></tr>                                                                                                                                                                                       
  	<tr>                                                                                                                                                                                        
		<td valign=middle align=left class=tablebody2 height="22"><b>���ɼ��:</b><br>�������������򵥵Ľ��ܣ����������������л���ʾ����</td>                                                                                                                                                                                        
		<td valign=middle align=left class=tablebody1 colspan="3" height="22">&nbsp;<textarea name=menpaireadme rows="4" cols="60"><%=menpai_setting(7)%></textarea>		(��Ҫ����100��Ӣ�Ļ�50������)</td>                                                                                                                                                                                      
  	</tr>  
	<tr><td height=8 class=tablebody2 colspan="4"></td></tr>                                                                                                                                                                                       
  	<tr>                                                                                                                                                                                       
    	<td valign=middle align=center class=tablebody1 colspan="4" height="22"><INPUT type=submit value=���� name=submit></td>                                                                                                                                                                                       
  	</tr>                                                                                                                                                                              
  	</form>                                                                                                                                                                                    
    <tr><td height=15 class=tablebody2 colspan="4"></td></tr>                                                                                                                                                                                                   
	<form name="form1" method="post" action="z_menpaimember.asp">                                                                                                                                                                               
	<INPUT name=Groupname size="20" value="<%=groupname%>" type=hidden>                                                                                                                                                                          
	<tr>                                                                                                                                                                                         
		<th valign=middle align=left height=24 colspan="4" >&nbsp;��&nbsp;<b>������Ա��ݹ���</b></th>                                                                                                                                                                                         
	</tr>                                                                                                                                                                               
	<tr>                                                                                                                                                                                         
		<td valign=middle align=left class=tablebody2 height="22"><b>���ɵ���</b><br>�����뱾�ɵ��ӵ��û���</td>                                                                                                                                                                                         
		<td valign=middle align=left class=tablebody1 colspan="3">
		&nbsp;<INPUT name=Susername type=text>&nbsp;&nbsp;
		<INPUT type=submit value=���� name=action>&nbsp;&nbsp;
		<INPUT type=checkbox value=1 name=search checked>��ȫƥ���û���&nbsp;&nbsp;
		<INPUT type=submit value=�����б� name=action>&nbsp;&nbsp;
	</tr>
	</form>  
	<form name="form3" method="post" action="z_menpaiadd.asp?menu=5">                                                                                                                                                                               
	<INPUT name=Groupname size="20" value="<%=groupname%>" type=hidden> 	
  	<tr>                                                                                                                                                                                       
    	<td valign=middle align=left class=tablebody2 height="22"><b>�����Ż�</b><br>�����������߳����ɵ��û���</td>
    	<td valign=middle align=left class=tablebody1 colspan="3">&nbsp;<INPUT name=Tusername type=text>&nbsp;&nbsp;                                                                                                                                                                        
      		<INPUT type=submit value=���� name=submit onclick="{if(confirm('��ȷ��ִ�����˲�����?')){this.document.form3.submit();return true;}return false;}">
		</td>                                                                                                                                                                                         
  	</tr>                                                                                                                                                                       
  	</form>  
  	<tr><td height=8 class=tablebody2 colspan="4"></td></tr>                                                                                                                                                                    
	<tr>                                                                                                                                                                                           
		<th valign=middle align=left height=24 colspan="4">&nbsp;��&nbsp;ע������</th>                                                                                                                                                                                          
	</tr>                                                                                                                                                                                 
  	<form name="form2" method="post" action="z_menpaiadd.asp?menu=9">                                                                                                                     
  	<INPUT name=Groupname size="20" value="<%=Groupname%>" type=hidden>                                                                                                                                                                     
	<tr>  
		<td valign=middle align=left class=tablebody2 height="22" colspan="2"><b>ע������</b><br>�ò�������ɾ�������ɣ��뿼���������в���</td>
		<td valign=middle align=center class=tablebody1 colspan="2">
		  	<INPUT type=submit value=ע�� name=submit onclick="{if(confirm('��ȷ��Ҫע��������?')){this.document.form2.submit();return true;}return false;}">                                                                                                                                                                  
		</td>                                                                                                                                                                                          
	</tr>                                                                                                                                                                   
  </form>                                                                                                                                                                                              
 </table>                                                                                                                                                                                            
<%                                                                                                                                                                                                                             
end sub                                                                                                                                                                                                                             
'----------------------------------�����Ż�-----------------------------------
sub clearmember()   
	if not CanManage then
		Errmsg=Errmsg+"<br>"+"<li>�����Ǳ������ţ�û�й����ɵ�Ȩ��"
		founderr=true
		exit sub	
	end if
	                                                                                                                                                                                            
	dim Tusername                                                                                                                                                                                                                                   
	Groupname=checkstr(trim(Request.form("Groupname")))
	Tusername=checkstr(trim(Request.form("Tusername")))
	
	if Groupname="" or isnull(Groupname) then
		Errmsg=Errmsg+"<br>"+"<li>��������ɲ�������ָ��������ɲ���"
		founderr=true	
	end if
	if Tusername="" or isnull(Tusername) then
		Errmsg=Errmsg+"<br>"+"<li>�����������߳����ɵ��û���"
		founderr=true	
	end if
	if founderr then exit sub
		
	Set rs= Server.CreateObject("ADODB.Recordset")                                                                                                                                                                                                                                   
	sql="select * from [user] where username='"&Tusername&"'"                                                                                                                                                                                                                                   
	rs.open sql,conn,1,3                                                                                              
	if rs.eof and rs.bof then                                                                                             
		Errmsg=Errmsg+"<br>"+"<li>��������û�������"
		founderr=true
		exit sub
	elseif rs("usergroup")<>Groupname then                                                                                             
		Errmsg=Errmsg+"<br>"+"<li>��������û����Ǳ��ɵ���"
		founderr=true
		exit sub                                                                                            
	else                                                                                             
		rs("usergroup")=wumenwupai
		rs.update                                                                                             
	end if                                                                                                                                                                                            
	rs.close                                                                                             
	set rs=nothing  
	
	call sendmessage(Tusername,Groupname)	 '�����������Ϣ���û�
	Sucmsg=Sucmsg+"<br>"+"<li>���Ѿ��ɹ���[<font color=red>"&Tusername&"</font>]���������"
	Sucmsg=Sucmsg+"<br>"+"<li>�Ѿ�����������֪ͨ�����û�" 	 
	call menpai_suc()	                                                                                                                                                                                                                          
end sub                                                                                                                                                                                                                             
'-------------------------------�����������ϵ��޸�---------------------------
sub changegroup()  
	if not CanManage then
		Errmsg=Errmsg+"<br>"+"<li>�����Ǳ������ţ�û�й����ɵ�Ȩ��"
		founderr=true
		exit sub	
	end if
	                                                                                                                                                   
	dim MenpaiAttri,userwealth,usercp,userep,userpower,userclass,userartitle
	'GroupName=checkstr(trim(Request.form("GroupName")))
	'newGroupName=checkstr(trim(Request.form("newGroupName")))
	GroupName=checkstr(trim(Request.form("GroupName")))
	MenpaiAttri=Request.form("MenpaiAttri")		'MenpaiAttri=Request.form("newMenpaiAttri")
	
	if GroupName="" then		'�����������
		Errmsg=Errmsg+"<br>"+"<li>��������������"
		founderr=true
	elseif strLength(GroupName)>12 then
		Errmsg=Errmsg+"<br>"+"<li>�������Ʋ��ܳ���12��Ӣ����ĸ(6������)"
		founderr=true		
	end if		
	if cint(MenpaiAttri)<>0 then
		if not isInteger(Request.form("userwealth")) then
			Errmsg=Errmsg+"<br>"+"<li>���������ɱ�����д��Ǯֵ����������Ĳ�����Ч����ֵ"
			founderr=true
		elseif int(Request.form("userwealth"))>mymoney then
			Errmsg=Errmsg+"<br>"+"<li>������Ľ�Ǯ��Чֵ�������Լ��ĵ�ǰֵ��Request.form(""userwealth"")="&Request.form("userwealth")&"  mymoney="&mymoney&"==="&(Request.form("userwealth")>mymoney)
			founderr=true			
		else
			userwealth=int(Request.form("userwealth"))
		end if	
		if not isInteger(Request.form("userep")) then
			Errmsg=Errmsg+"<br>"+"<li>���������ɱ�����д����ֵ����������Ĳ�����Ч����ֵ"
			founderr=true
		elseif int(Request.form("userep"))>myuserep then
			Errmsg=Errmsg+"<br>"+"<li>������ľ�����Чֵ�������Լ��ĵ�ǰֵ��"
			founderr=true			
		else 
			userep=int(Request.form("userep"))	
		end if
		if not isInteger(Request.form("usercp")) then
			Errmsg=Errmsg+"<br>"+"<li>���������ɱ�����д����ֵ����������Ĳ�����Ч����ֵ"
			founderr=true
		elseif int(Request.form("usercp"))>myusercp then
			Errmsg=Errmsg+"<br>"+"<li>�������������Чֵ�������Լ��ĵ�ǰֵ��"
			founderr=true			
		else 
			usercp=int(Request.form("usercp"))	
		end if
		if not isInteger(Request.form("userpower")) then
			Errmsg=Errmsg+"<br>"+"<li>���������ɱ�����д����ֵ����������Ĳ�����Ч����ֵ"
			founderr=true
		elseif int(Request.form("userpower"))>mypower then
			Errmsg=Errmsg+"<br>"+"<li>�������������Чֵ�������Լ��ĵ�ǰֵ��"&mypower
			founderr=true			
		else 
			userpower=int(Request.form("userpower"))	
		end if
		if not isInteger(Request.form("userarticle")) then
			Errmsg=Errmsg+"<br>"+"<li>���������ɱ�����д����������������Ĳ�����Ч����ֵ"
			founderr=true
		elseif int(Request.form("userarticle"))>myarticle then
			Errmsg=Errmsg+"<br>"+"<li>�������������Чֵ�������Լ��ĵ�ǰֵ��"
			founderr=true			
		else 
			userartitle=int(Request.form("userarticle"))	
		end if
		if not isInteger(Request.form("userclass")) then
			Errmsg=Errmsg+"<br>"+"<li>���������ɱ�����д����ֵ����������Ĳ�����Ч����ֵ"
			founderr=true
		elseif int(Request.form("userclass"))>GetUserClass(myclass) then
			Errmsg=Errmsg+"<br>"+"<li>������ĵȼ���Чֵ�������Լ��ĵ�ǰֵ��"
			founderr=true
		elseif int(Request.form("userclass"))<1 or int(Request.form("userclass"))>int(maxtitleid) then
			Errmsg=Errmsg+"<br>"+"<li>�ȼ�ֵ����С��1�����"&maxtitleid
			founderr=true						
		else 
			userclass=int(Request.form("userclass"))	
		end if
	else
		userwealth=0
		userep=0
		usercp=0
		userpower=0
		userartitle=0
		userclass=1
	end if
	
	dim menpaireadme
	menpaireadme=Request.form("menpaireadme")	
	if strLength(menpaireadme)>100 then
		Errmsg=Errmsg+"<br>"+"<li>���ɼ�鲻�ܳ���100��Ӣ�Ļ�50������"
		founderr=true		
	end if
			
	if founderr then exit sub
																																																						   
	set rs= server.createobject ("adodb.recordset")                                                                                                                                                                                                                                   
	sql="select * from [GroupName] where GroupName='"&GroupName&"'"                                                                                                                                                                                                                                   
	rs.Open sql,conn,1,3  
	 
	if rs.bof and rs.eof then   	'�ù����Ժ���չ�����޸�������������                                                                            
		Errmsg=Errmsg+"<br>"+"<li>�����ɲ����ڰ�"
		founderr=true
		rs.close
		set rs=nothing
		exit sub
	else                                  
		'rs("Groupname")=GroupName
		rs("Menpai_setting")=MenpaiAttri & "|" & userwealth & "|" & userep & "|" & usercp & "|" & userpower & "|" & userartitle & "|" & userclass & "|" & menpaireadme
		rs.update  
		'conn.execute("update [user] set usergroup='"&GroupName&"',userwealth=userwealth-"&new_menpai_setting(6)&" where username='"&membername&"'")	                                            
	end if
	Sucmsg=Sucmsg+"<br>"+"<li>���������޸ĳɹ�" 	 
	call menpai_suc()
end sub                                                                                                                                                                                                                             
'----------------------------ע������-------------------------------
sub logout()
                                                                                              
	if not CanManage then
		Errmsg=Errmsg+"<br>"+"<li>�����Ǳ������ţ�û�й����ɵ�Ȩ��"
		founderr=true
		exit sub	
	end if
	if Groupname="" then 
		Errmsg=Errmsg+"<br>"+"<li>��������ɲ�������ָ��������ɲ���"
		founderr=true
		exit sub
	elseif Groupname=wumenwupai then
		Errmsg=Errmsg+"<br>"+"<li>���������������������ɵ��˵��������ɣ�����ע��"
		founderr=true
		exit sub	
	end if
	                                                                                  
    conn.execute("update [GroupName] set state=2,Zangmen='-������·-' where GroupName='"&GroupName&"'")
	conn.execute("update [user] set usergroup='"&wumenwupai&"' where usergroup='"&GroupName&"'")

	Sucmsg=Sucmsg+"<br>"+"<li>���Ѿ�ע���˱���" 	 
	call menpai_suc()	
end sub                                                                                                                                                                                                                             
'-----------------------------����������==����------------------------------
sub newmenpai()       
	if isZangmen then
		Errmsg=Errmsg+"<br>"+"<li>���Ѿ��� "&Groupname&" ���ţ���ô����������"
		founderr=true
		exit sub
	end if
	if mymenpai<>wumenwupai then
		Errmsg=Errmsg+"<br>"+"<li>���Ѿ��� "&Groupname&" ���ӣ����ܴ���������"
		founderr=true
		exit sub		
	end if
	
	dim new_menpai_setting
	new_menpai_setting=split(conn.execute("select new_menpai_setting from [menpai]")(0),",")
		
	if int(mymoney)<(int(new_menpai_setting(0))+int(new_menpai_setting(6))) or int(myusercp)<int(new_menpai_setting(1)) or int(myuserep)<int(new_menpai_setting(2)) or int(mypower)<int(new_menpai_setting(3)) or int(myArticle)<int(new_menpai_setting(4)) or int(GetUserClass(myClass))<int(new_menpai_setting(5)) then
		Errmsg=Errmsg+"<br>"+"<li>��������һ�����������㴴��Ҫ��"
		founderr=true
		exit sub
	end if	
%>                                                                                             
<table cellspacing=1 align=center class=tableborder1 width=<%=Forum_body(12)%>>                                                                                                                                                                              
	<form name="form" method="post" action="z_menpaiadd.asp?menu=7">                                                                                                                                                                                            
  	<tr>                                                                                                                                                                                           
    	<th valign=middle align=center height=24 colspan="4">����������</th>                                                                                                                                                                                          
  	</tr> 
	<tr><td height=8 class=tablebody2 colspan="4"></td></tr>                                                                                                                                                                                        
    <tr>                                                                                                                                                                                          
		<td valign=middle align=left class=tablebody2 height="22"><b>������֪:</b></td>                                                                                                                                                                                         
		<td valign=middle align=left class=tablebody1 height="22" colspan="3">
		 	<li>��ϲ�������߱��˴������ɵ����� 
			<li>���ɴ����ɹ�֮�󣬵ȴ�����Ա����ˣ����ͨ��֮�󣬱�������ʽ���� 
			<li>������ʽ����֮���������Ǳ��ɵ������� 
			<li>��������˲������Σ�����Ա��Ȩ�����������˵�ְ�� 
			<li>���ɷ���<font color=red><%=new_menpai_setting(6)%></font>Ԫ  
			<li>��������д�������������
		</td> 
  	</tr> 	   
	<tr><td height=8 class=tablebody2 colspan="4"></td></tr> 
  	<tr> 
		<td valign=middle align=left class=tablebody2 height="22"><b>��������:</b><br> ������������֮�����Ų����޸�</td> 
		<td valign=middle align=left class=tablebody1 height="22" colspan="3">&nbsp;<INPUT name=groupname size="20" maxlength="20"> <font color=red>*</font> (��Ҫ����12��Ӣ����ĸ��6������)</td>
  	</tr>
  	<tr> 
		<td valign=middle align=left class=tablebody2 height="22"><b>��������:</b><br> ������������֮�����Ų����޸�</td>
		<td valign=middle align=left class=tablebody1 height="22" colspan="3">
			<input type=radio name=menpaiattri value=0 checked>��������&nbsp;
			<input type=radio name=menpaiattri value=1>����&nbsp;
			<input type=radio name=menpaiattri value=2>ħ��<br>
		</td>
  	</tr>
	<tr><td height=8 class=tablebody2 colspan="4"></td></tr>
    <tr>
		<td valign=middle align=left class=tablebody2 height="22"><b>˵��:</b></td>                                                                                                                                                                                         
		<td valign=middle align=left class=tablebody1 height="22" colspan="3">
		 	<li>��Ǯ,����,����,����,������,�û��ȼ� �����ܳ��������ѵĵ�ǰֵ
			<li>��Ǯ,����,����,����,������,�û��ȼ� �Ǽ��뱾�ɵı����������Сֵ
			<li>��Ǯ,����,����,����,������,�û��ȼ� Ҳ�Ǽ�����˳���������/�۳����ֵʱ�Ļ���
			<li>�������ɲ�����������
			<li>�����������ſ��������ɹ������޸�			
		</td> 
  	</tr> 
    <tr>                                                                                                                                                                                         
		<td valign=middle align=left class=tablebody2 height="22" width="25%"><b>��Ǯ:</b></td>                                                                                                                                                                                        
		<td valign=middle align=left class=tablebody1 height="22">&nbsp;<INPUT name=userwealth value=""></td>                                                                                                                                                                                        
		<td valign=middle align=right class=tablebody2 height="22">���Ľ�Ǯ:&nbsp;</td>
		<td valign=middle align=left class=tablebody1 height="22" width="20%">&nbsp;<%=mymoney%></td>
  	</tr>                                                                                                                                                                                        
  	<tr>                                                                                                                                                                                        
		<td valign=middle align=left class=tablebody2 height="22"><b>����:</b></td>                                                                                                                                                                                        
		<td valign=middle align=left class=tablebody1 height="22">&nbsp;<INPUT name=userep value=""></td>
		<td valign=middle align=right class=tablebody2 height="22">���ľ���:&nbsp;</td>                                                                                                                                                                                        
		<td valign=middle align=left class=tablebody1 height="22">&nbsp;<%=myusercp%></td>                                                                                                                                                                                        
  	</tr>                                                                                                                                                                                        
  	<tr>                                                                                                                                                                                        
		<td valign=middle align=left class=tablebody2 height="22"><b>����:</b></td>                                                                                                                                                                                        
		<td valign=middle align=left class=tablebody1 height="22">&nbsp;<INPUT name=usercp value=""></td> 
		<td valign=middle align=right class=tablebody2 height="22">��������:&nbsp;</td>                                                                                                                                                                                       
		<td valign=middle align=left class=tablebody1 height="22">&nbsp;<%=myuserep%></td>                                                                                                                                                                                        
  	</tr>                                                                                                                                                                                        
  	<tr>                                                                                                                                                                                        
		<td valign=middle align=left class=tablebody2 height="22"><b>����:</b></td>                                                                                                                                                                                        
		<td valign=middle align=left class=tablebody1 height="22">&nbsp;<INPUT name=userpower value=""></td> 
		<td valign=middle align=right class=tablebody2 height="22">��������:&nbsp;</td>                                                                                                                                                                                       
		<td valign=middle align=left class=tablebody1 height="22">&nbsp;<%=mypower%></td>                                                                                                                                                                                        
  	</tr>                                                                                                                                                                                        
  	<tr>                                                                                                                                                                                        
		<td valign=middle align=left class=tablebody2 height="22"><b>������:</b><br>���뱾�ɱ���ﵽ�ķ�����</td>                                                                                                                                                                                        
		<td valign=middle align=left class=tablebody1 height="22">&nbsp;<INPUT name=userarticle value=""></td> 
		<td valign=middle align=right class=tablebody2 height="22">��������:&nbsp;</td>                                                                                                                                                                                       
		<td valign=middle align=left class=tablebody1 height="22">&nbsp;<%=myarticle%></td>                                                                                                                                                                                        
  	</tr>                                                                                                                                                                                        
  	<tr>                                                                                                                                                                                        
		<td valign=middle align=left class=tablebody2 height="22"><b>�û��ȼ�:</b><br>���뱾�ɱ���߱��ĵȼ�(1��<%=maxtitleid%>)</td>                                                                                                                                                                                        
		<td valign=middle align=left class=tablebody1 height="22">&nbsp;<INPUT name=userclass value=""></td> 
		<td valign=middle align=right class=tablebody2 height="22">���ĵȼ�:&nbsp;</td>                                                                                                                                                                                     
		<td valign=middle align=left class=tablebody1 height="22">&nbsp;<%=myclass%>��<%=GetUserClass(myclass)%>����</td>                                                                                                                                                                                      
  	</tr>
	<tr>
		<td valign=middle align=center class=tablebody1 height=22 colspan="4"><font color="#CC9933">
		��������������������ɣ�������6��(��Ǯ,����,����,����,������,�û��ȼ�)������д</font>
		</td>
	</tr> 
	<tr><td height=8 class=tablebody2 colspan="4"></td></tr>                                                                                                                                                                                       
  	<tr>                                                                                                                                                                                        
		<td valign=middle align=left class=tablebody2 height="22"><b>���ɼ��:</b><br>�������������򵥵Ľ��ܣ����������������л���ʾ����</td>                                                                                                                                                                                        
		<td valign=middle align=left class=tablebody1 colspan="3" height="22">&nbsp;<textarea name=menpaireadme rows="4" cols="60"></textarea>		(��Ҫ����100��Ӣ�Ļ�50������)</td>                                                                                                                                                                                      
  	</tr>  
	<tr><td height=8 class=tablebody2 colspan="4"></td></tr>                                                                                                                                                                                       
  	<tr>                                                                                                                                                                                       
    	<td valign=middle align=center class=tablebody1 colspan="4" height="22"><INPUT type=submit value=���������� name=submit></td>                                                                                                                                                                                       
  	</tr>                                                                                                                                                                              
  	</form>                                                                                                                                                                                    
</table>                                                                                                                                                                                           
<%                                             
end sub                                                                                                                                                                                                                            

'------------------------------------����������====�ύ����--------------------------------
sub applynewmenpai() 
	dim MenpaiAttri,userwealth,usercp,userep,userpower,userclass,userartitle
	GroupName=checkstr(trim(Request.form("groupname")))
	MenpaiAttri=Request.form("menpaiattri")
	
	if GroupName="" then
		Errmsg=Errmsg+"<br>"+"<li>��������������"
		founderr=true
	elseif strLength(GroupName)>12 then
		Errmsg=Errmsg+"<br>"+"<li>�������Ʋ��ܳ���12��Ӣ����ĸ(6������)"
		founderr=true		
	end if		
	if cint(MenpaiAttri)<>0 then
		if not isInteger(Request.form("userwealth")) then
			Errmsg=Errmsg+"<br>"+"<li>���������ɱ�����д��Ǯֵ����������Ĳ�����Ч����ֵ"
			founderr=true
		elseif int(Request.form("userwealth"))>mymoney then
			Errmsg=Errmsg+"<br>"+"<li>������Ľ�Ǯ��Чֵ�������Լ��ĵ�ǰֵ��Request.form(""userwealth"")="&Request.form("userwealth")&"  mymoney="&mymoney&"==="&(Request.form("userwealth")>mymoney)
			founderr=true			
		else
			userwealth=int(Request.form("userwealth"))
		end if	
		if not isInteger(Request.form("userep")) then
			Errmsg=Errmsg+"<br>"+"<li>���������ɱ�����д����ֵ����������Ĳ�����Ч����ֵ"
			founderr=true
		elseif int(Request.form("userep"))>myuserep then
			Errmsg=Errmsg+"<br>"+"<li>������ľ�����Чֵ�������Լ��ĵ�ǰֵ��"
			founderr=true			
		else 
			userep=int(Request.form("userep"))	
		end if
		if not isInteger(Request.form("usercp")) then
			Errmsg=Errmsg+"<br>"+"<li>���������ɱ�����д����ֵ����������Ĳ�����Ч����ֵ"
			founderr=true
		elseif int(Request.form("usercp"))>myusercp then
			Errmsg=Errmsg+"<br>"+"<li>�������������Чֵ�������Լ��ĵ�ǰֵ��"
			founderr=true			
		else 
			usercp=int(Request.form("usercp"))	
		end if
		if not isInteger(Request.form("userpower")) then
			Errmsg=Errmsg+"<br>"+"<li>���������ɱ�����д����ֵ����������Ĳ�����Ч����ֵ"
			founderr=true
		elseif int(Request.form("userpower"))>mypower then
			Errmsg=Errmsg+"<br>"+"<li>�������������Чֵ�������Լ��ĵ�ǰֵ��"
			founderr=true			
		else 
			userpower=int(Request.form("userpower"))	
		end if
		if not isInteger(Request.form("userarticle")) then
			Errmsg=Errmsg+"<br>"+"<li>���������ɱ�����д����������������Ĳ�����Ч����ֵ"
			founderr=true
		elseif int(Request.form("userarticle"))>myarticle then
			Errmsg=Errmsg+"<br>"+"<li>�������������Чֵ�������Լ��ĵ�ǰֵ��"
			founderr=true			
		else 
			userartitle=int(Request.form("userarticle"))	
		end if
		if not isInteger(Request.form("userclass")) then
			Errmsg=Errmsg+"<br>"+"<li>���������ɱ�����д����ֵ����������Ĳ�����Ч����ֵ"
			founderr=true
		elseif int(Request.form("userclass"))>GetUserClass(myclass) then
			Errmsg=Errmsg+"<br>"+"<li>������ĵȼ���Чֵ�������Լ��ĵ�ǰֵ��"
			founderr=true
		elseif int(Request.form("userclass"))<1 or int(Request.form("userclass"))>int(maxtitleid) then
			Errmsg=Errmsg+"<br>"+"<li>�ȼ�ֵ����С��1�����"&maxtitleid
			founderr=true						
		else 
			userclass=int(Request.form("userclass"))	
		end if
	else
		userwealth=0
		userep=0
		usercp=0
		userpower=0
		userartitle=0
		userclass=1
	end if
	
	dim menpaireadme
	menpaireadme=Request.form("menpaireadme")	
	if strLength(menpaireadme)>100 then
		Errmsg=Errmsg+"<br>"+"<li>���ɼ�鲻�ܳ���100��Ӣ�Ļ�50������"
		founderr=true		
	end if
			
	if founderr then exit sub
	
	dim new_menpai_setting		'new_menpai_setting(6)   �����շ�
	new_menpai_setting=split(conn.execute("select new_menpai_setting from [menpai]")(0),",")                         
                           
	set rs= server.createobject ("adodb.recordset")                                                                                                                                                                                                                                   
	sql="select * from [GroupName] where groupname='"&groupname&"'"                                                                                                                                                                                                                                   
	rs.Open sql,conn,1,3                                                                                                                                                                                                            
	if not(rs.bof and rs.eof) then                                                                               
		Errmsg=Errmsg+"<br>"+"<li>����������������Ѿ�����"
		founderr=true
		rs.close
		set rs=nothing
		exit sub
	else                                  
		rs.addnew                                                                                                                                                          
		rs("Groupname")=GroupName
		rs("Zangmen")=membername
		rs("Addtime")=date()
		rs("Menpai_setting")=MenpaiAttri & "|" & userwealth & "|" & userep & "|" & usercp & "|" & userpower & "|" & userartitle & "|" & userclass & "|" & menpaireadme
		if cint(new_menpai_setting(7))=1 then
			 rs("State")=3		'�ȴ����
		else
			 rs("State")=0		'����Ҫ��ˣ�ֱ�ӿ�ͨ
		end if 
		rs.update  
		conn.execute("update [user] set usergroup='"&GroupName&"',userwealth=userwealth-"&new_menpai_setting(6)&" where username='"&membername&"'")	                                            
	end if
	if cint(new_menpai_setting(7))=1 then
		Sucmsg=Sucmsg+"<br>"+"<li>���ɹ������� " & GroupName & "����ȴ�����Ա�����"
	else
		Sucmsg=Sucmsg+"<br>"+"<li>�������Ѿ�����������Ϊ�����ɵ������ˣ���ϲ����������"
	end if
	Sucmsg=Sucmsg+"<br>"+"<li>�������� "&new_menpai_setting(6)&" Ԫ�Ĵ��ɷ�"
	call menpai_suc()
end sub  
'--------------------------------���Ͷ���Ϣ--------------------------------
sub sendmessage(Susername,menpainame)
	dim message,title                                        
	title=menpainame&" �����Ż�֪ͨ"                                          
	message=Susername&"��ã������㲻�����Ź棬�Ѿ�������������� "&menpainame
	sql="insert into message (incept,sender,title,content,sendtime,flag,issend) values ('"&Susername&"','"&menpainame&"������','"&title&"','"&message&"',Now(),0,1)"                                          
	conn.execute(sql)                                          
end sub                                                                                                                                                                                                                            
%>