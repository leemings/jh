<!--#include file=conn.asp-->
<!-- #include file="inc/const.asp" -->
<!-- #include file="z_menpaiconfig.asp" -->
<title><%=Forum_info(0)%>--����ҳ��</title>
<link rel="stylesheet" href="forum_admin.css" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<BODY leftmargin="0" bottommargin="0" rightmargin="0" topmargin="0" marginheight="0" marginwidth="0">
<%
	if not master or session("flag")="" then
		Errmsg=Errmsg+"<br>"+"<li>��ҳ��Ϊ����Աר�ã���<a href=admin_index.asp target=_top>��¼</a>����롣<br><li>��û�й���ҳ���Ȩ�ޡ�"
		call dvbbs_error()
	else
		dim GroupID,Groupname
		select case Request("action")
			case "add"
				call  newmenpai()    ' �������
			case "applynewmenpai"
				call applynewmenpai()
			case "edit"
				call edit()		'�޸�����
			case "savedit"
				call savedit()	
			case "del"
				call del()		'ɾ������
			case "editapply"
				call editapply()  '�޸����봴�����ɵ�����
			case "saveditapply"
				call saveditapply()
			case "pass"			'��׼�����ɳ���
				call pass()
			case "repass"			'���¿�ͨ�Ѿ�ע��������
				call repass()							
			case else			
				call main()
		end select 	
		conn.close
		set conn=nothing
	end if
	if founderr then call dvbbs_error()
'=========================================================	
sub main()
%>
<br>
<table width="95%" border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
	<tr>
		<th height=21 align="left">���ɹ��� --> <a href=z_admin_menpai.asp><font color=#FFFFFF>�����б�</font></a> | <a href=z_admin_menpai.asp?action=add><font color=#FFFFFF>�������</font></a></th>
	</tr>
</table>		
<br>
<table width="95%" border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
  <tr>
    <th height=24>��������</th>
    <th height=24>��������</th> 
    <th height=24>��������</th> 
    <th height=24>����</th> 
    <th height=24>������</th> 
    <th height=24>��Ǯ</th> 
    <th height=24>������</th> 
    <th height=24>����</th> 
  </tr>
  <%
'������ʾ  ��ʼ  
	dim rs1
	dim Zangmen,TotalArticle,TotalWealth,TotalUsernum
	dim menpai_setting	
	sql="select * from [GroupName] where state<2"
	set rs=conn.execute(sql)
	if rs.eof and rs.bof  then
		response.write "<tr><td height=21 class=forumrow colspan=8><font color=red>����̳��ʱû���κ�����</font></td></tr>"
	else 
		do while not rs.eof
			TotalArticle=0
			TotalWealth=0
			TotalUsernum=0		
			menpai_setting=split(rs("menpai_setting"),"|")
			Zangmen=trim(rs("Zangmen"))
			if Zangmen="" or isnull(Zangmen) then
				Zangmen="<font color=#FF6666>- ��ѡ�� -</font>"
			elseif Zangmen="-������·-" then
				Zangmen="<font color=#CC6666>������·</font>"
			else
				Zangmen="<a href=dispuser.asp?name="&htmlencode(Zangmen)&" title=""�鿴�������˵�����"" target=_blank>"&Zangmen&"</a>"
			end if	

			'ͳ��ָ�����ɵ�����������Ǯ�� �� ���������������ţ�
			set rs1=conn.execute("select sum(article),sum(userwealth),count(*) from [user] where usergroup='"&rs("GroupName")&"'")
			if rs1(0)<>"" then TotalArticle=rs1(0)
			if rs1(1)<>"" then TotalWealth=rs1(1)
			if rs1(2)<>"" then TotalUsernum=rs1(2)
%>
<tr>
	<td valign=middle align=center height=21 class=forumrow><a href="z_menpaiadd.asp?menu=1&groupname=<%=rs("groupname")%>" title="���ɼ��:<%=menpai_setting(7)%>"><%=rs("groupname")%></a></td>
	<td valign=middle align=center height=21 class=forumrow><%=Zangmen%></td>
    <td valign=middle align=center height=21 class=forumrow><%=rs("addtime")%></td>
    <td valign=middle align=center height=21 class=forumrow><%=MenpaiAttri(menpai_setting(0))%></td>
    <td valign=middle align=center height=21 class=forumrow><%=TotalUsernum%></td>
    <td valign=middle align=center height=21 class=forumrow><%=TotalWealth%></td>
    <td valign=middle align=center height=21 class=forumrow><%=TotalArticle%></td>
    <td valign=middle align=center height=21 class=forumrow><a href="z_admin_menpai.asp?id=<%=rs("id")%>&action=edit">�޸�</a> | <a href="z_admin_menpai.asp?id=<%=rs("id")%>&action=del" onclick="{if(confirm('��ȷ��Ҫɾ�� <%=rs("groupname")%> ��?')){return true;}return false;}">ɾ��</a></td>
</tr>
<%			rs.movenext
		loop
		rs.close
	end if
%>
	<tr><td height=21 class=forumRowHighlight colspan=8></td></tr>
</table>
<%	  
'������ʾ  ����

'���������ɲ���  ���뿪ʼ
	dim new_menpai_setting
	new_menpai_setting=split(conn.execute("select new_menpai_setting from [menpai]")(0),",")

	'set rs=conn.execute("select userwealth,usercp,userep,userpower,userclass from [user] where username='"&membername&"'")	
 	
	dim wumenwupai			'���ɲ������κ����ɵ��ӵ�����
	set rs=conn.execute("select GroupName from [GroupName] where state=1")  
	if rs.eof and rs.bof then 
		wumenwupai="��������"
	else
		wumenwupai=rs(0)
	end if
	rs.close
	
	dim mymenpai,CanCreateMenpai
	CanCreateMenpai=true
	set rs=conn.execute("select GroupName from [GroupName] where Zangmen='"&membername&"'")
	if rs.eof and rs.bof then
		mymenpai=conn.execute("select UserGroup from [user] where UserID="&UserID)(0)
		if mymenpai=wumenwupai or mymenpai="��������" then
			mymenpai="<font color=navy>"&mymenpai&"</font>"
		else
			mymenpai="<font color=navy>"&mymenpai&"</font><font color=red>����</font>"
		end if	
	else
		mymenpai="<font color=navy>"&rs(0)&"</font><font color=red>����</font>"
	end if    		
%>
<br>
<table width="95%" border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
<tr><th colspan=9 height=20>���뽨�����ɵ�����</th></tr>
<tr>
	<td align=center height=21 class=forumRowHighlight>��Ŀ</td>
    <td align=center height=21 class=forumRowHighlight>��Ǯ</td>
	<td align=center height=21 class=forumRowHighlight>����</td>
	<td align=center height=21 class=forumRowHighlight>����</td>
	<td align=center height=21 class=forumRowHighlight>����</td>
	<td align=center height=21 class=forumRowHighlight>����</td>
	<td align=center height=21 class=forumRowHighlight>�û��ȼ�</td> 
	<td align=center height=21 class=forumRowHighlight>���</td> 
	<td align=center height=21 class=forumRowHighlight>����</td>
</tr>
<tr>
	<td align=center height=21 class=forumrow>����Ҫ��</td>
    <td align=center height=21 class=forumrow><%=new_menpai_setting(0)%>+<font color=red><%=new_menpai_setting(6)%></font></td>
	<td align=center height=21 class=forumrow><%=new_menpai_setting(1)%></td>
	<td align=center height=21 class=forumrow><%=new_menpai_setting(2)%></td>
	<td align=center height=21 class=forumrow><%=new_menpai_setting(3)%></td>
	<td align=center height=21 class=forumrow><%=new_menpai_setting(4)%></td>
	<td align=center height=21 class=forumrow><%=GetUserTitle(new_menpai_setting(5))%>(<%=new_menpai_setting(5)%>��)</td>
	<td align=center height=21 class=forumrow><%=iif(cint(new_menpai_setting(7))=1,"��Ҫ","����Ҫ")%></td>
	<td align=center height=21 class=forumrow><%=wumenwupai%></td>
</tr>
<tr><td align=center height=21 class=forumRowHighlight colspan=9><a href=z_admin_menpai.asp?action=editapply>�޸�</a></td></tr>
</table>
<%
'���������ɲ���  ������� 

'�ȴ���˵����� ���뿪ʼ
%>	
<br>
<table width="95%" border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
  	<tr><th height=24 colspan="5">�ȴ���˵�����</th></tr> 
  	<tr>
		<td align=center height=21 class=forumRowHighlight>��������</th>
		<td align=center height=21 class=forumRowHighlight>��ʼ��</th> 
		<td align=center height=21 class=forumRowHighlight>��������</th> 
		<td align=center height=21 class=forumRowHighlight>����</th> 
		<td align=center height=21 class=forumRowHighlight>����</th> 
  	</tr>
<%
	sql="select * from [GroupName] where state=3"
	set rs=conn.execute(sql)
	if rs.eof and rs.bof  then
		response.write "<tr><td height=21 class=forumrow colspan=8><font color=navy>��ʱû�еȴ���˵�����</font></td></tr>"
	else 
		do while not rs.eof
			menpai_setting=split(rs("menpai_setting"),"|")
			Zangmen=trim(rs("Zangmen"))
%>
<tr>
	<td valign=middle align=center height=21 class=forumrow><a href="z_menpaiadd.asp?menu=1&groupname=<%=rs("groupname")%>" title="���ɼ��:<%=menpai_setting(7)%>"><%=rs("groupname")%></a></td>
	<td valign=middle align=center height=21 class=forumrow><a href=dispuser.asp?name=<%=htmlencode(Zangmen)%> title="�鿴�������˵�����" target=_blank><%=Zangmen%></a></td>
    <td valign=middle align=center height=21 class=forumrow><%=rs("addtime")%></td>
    <td valign=middle align=center height=21 class=forumrow><%=MenpaiAttri(menpai_setting(0))%></td>
    <td valign=middle align=center height=21 class=forumrow><a href="z_admin_menpai.asp?id=<%=rs("id")%>&Zangmen=<%=Zangmen%>&action=pass">��׼</a> | <a href="z_admin_menpai.asp?id=<%=rs("id")%>&action=edit">�༭</a> | <a href="z_admin_menpai.asp?id=<%=rs("id")%>&Zangmen=<%=Zangmen%>&action=del" onclick="{if(confirm('��ȷ��ɾ�� <%=rs("groupname")%> ��?')){return true;}return false;}">ɾ��</a></td>
</tr>
<%			rs.movenext
		loop
		rs.close
	end if
%>
	<tr><td height=21 class=forumRowHighlight colspan=5></td></tr>
</table>
<%
'�ȴ���˵�����  �������

'��ע��������  ���뿪ʼ
%>	
<br>
<table width="95%" border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
  	<tr><th height=24 colspan="4">��ע��������</th></tr> 
  	<tr>
		<td align=center height=21 class=forumRowHighlight>��������</th>
		<td align=center height=21 class=forumRowHighlight>��������</th> 
		<td align=center height=21 class=forumRowHighlight>����</th> 
		<td align=center height=21 class=forumRowHighlight>����</th> 
  	</tr>
<%
	sql="select * from [GroupName] where state=2"
	set rs=conn.execute(sql)
	if rs.eof and rs.bof  then
		response.write "<tr><td height=21 class=forumrow colspan=8><font color=navy>��ʱû����ע��������</font></td></tr>"
	else 
		do while not rs.eof
			menpai_setting=split(rs("menpai_setting"),"|")
%>
<tr>
	<td valign=middle align=center height=21 class=forumrow><a href="z_menpaiadd.asp?menu=1&groupname=<%=rs("groupname")%>" title="���ɼ��:<%=menpai_setting(7)%>"><%=rs("groupname")%></a></td>
    <td valign=middle align=center height=21 class=forumrow><%=rs("addtime")%></td>
    <td valign=middle align=center height=21 class=forumrow><%=MenpaiAttri(menpai_setting(0))%></td>
    <td valign=middle align=center height=21 class=forumrow><a href="z_admin_menpai.asp?id=<%=rs("id")%>&action=repass">�ؿ�</a> | <a href="z_admin_menpai.asp?id=<%=rs("id")%>&action=edit">�༭</a> | <a href="z_admin_menpai.asp?id=<%=rs("id")%>&action=del" onclick="{if(confirm('��ȷ���˳� <%=rs("groupname")%> ��?')){return true;}return false;}">ɾ��</a></td>
</tr>
<%			rs.movenext
		loop
		rs.close
	end if
%>
	<tr><td height=21 class=forumRowHighlight colspan=5></td></tr>
</table>
<%
'��ע��������  �������
end sub

'----------------------------���������------------------------------
sub newmenpai()       
%>                                                                                             
<table cellspacing=1 align=center class=tableborder width="95%">                                                                                                                                                                              
	<form name="form" method="post" action="z_admin_menpai.asp?action=applynewmenpai">                                                                                                                                                                                            
  	<tr>                                                                                                                                                                                           
    	<th valign=middle align=center height=24 colspan="4">���������</th>                                                                                                                                                                                          
  	</tr> 
	<tr><td height=8 class=forumRowHighlight colspan="4"></td></tr>                                                                                                                                                                                        
  	<tr> 
		<td valign=middle align=left class=forumRowHighlight height="22"><b>��������:</b><br> ������������֮�����Ų����޸�</td> 
		<td valign=middle align=left class=forumrow height="22" colspan="3">&nbsp;<INPUT name=groupname size="20" maxlength="20"> <font color=red>*</font> (��Ҫ����12��Ӣ����ĸ��6������)</td>
  	</tr>
  	<tr> 
		<td valign=middle align=left class=forumRowHighlight height="22"><b>����:</b><br> �����ʱ�����������ţ�����������</td> 
		<td valign=middle align=left class=forumrow height="22" colspan="3">&nbsp;<INPUT name=Zangmen size="20" maxlength="20" value=<%=membername%>></td>
  	</tr>
	<tr><td height=8 class=forumRowHighlight colspan="4"></td></tr>
    <tr>
		<td valign=middle align=left class=forumRowHighlight height="22"><b>��������˵��:</b></td>                                                                                                                                                                                         
		<td valign=middle align=left class=forumrow height="22" colspan="3">
			<li>�����������������������ɵ��˵�һ����������
			<li>��ע���û��ͱ������߳����ɵ��û��Զ�����������
		 	<li>ֻ�ܴ���һ�������ɣ����Ե�������һ������Ϊ������֮����ǰ���������Զ�תΪ��������
		</td> 
  	</tr>		
  	<tr> 
		<td valign=middle align=left class=forumRowHighlight height="22"><b>��������:</b><br> ������������֮�����Ų����޸�</td>
		<td valign=middle align=left class=forumrow height="22" colspan="3">
			<input type=radio name=menpaiattri value=0 checked>��������&nbsp;
			<input type=radio name=menpaiattri value=1>����&nbsp;
			<input type=radio name=menpaiattri value=2>ħ��&nbsp;
			<input type=radio name=menpaiattri value=3>������<br>
		</td>
  	</tr>
	<tr><td height=8 class=forumRowHighlight colspan="4"></td></tr>
    <tr>
		<td valign=middle align=left class=forumRowHighlight height="22"><b>����ֵ��˵��:</b></td>                                                                                                                                                                                         
		<td valign=middle align=left class=forumrow height="22" colspan="3">
		 	<li>��Ǯ,����,����,����,������,�û��ȼ� �����ܳ��������ѵĵ�ǰֵ
			<li>��Ǯ,����,����,����,������,�û��ȼ� �Ǽ��뱾�ɵı����������Сֵ
			<li>��Ǯ,����,����,����,������,�û��ȼ� Ҳ�Ǽ�����˳���������/�۳����ֵʱ�Ļ���
			<li>�������ɲ�����������
			<li>�����������ſ��������ɹ������޸�			
		</td> 
  	</tr> 
    <tr>                                                                                                                                                                                         
		<td valign=middle align=left class=forumRowHighlight height="22" width="25%"><b>��Ǯ:</b></td>                                                                                                                                                                                        
		<td valign=middle align=left class=forumrow height="22">&nbsp;<INPUT name=userwealth value=""></td>                                                                                                                                                                                        
		<td valign=middle align=right class=forumRowHighlight height="22">���Ľ�Ǯ:&nbsp;</td>
		<td valign=middle align=left class=forumrow height="22" width="20%">&nbsp;<%=mymoney%></td>
  	</tr>                                                                                                                                                                                        
  	<tr>                                                                                                                                                                                        
		<td valign=middle align=left class=forumRowHighlight height="22"><b>����:</b></td>                                                                                                                                                                                        
		<td valign=middle align=left class=forumrow height="22">&nbsp;<INPUT name=userep value=""></td>
		<td valign=middle align=right class=forumRowHighlight height="22">���ľ���:&nbsp;</td>                                                                                                                                                                                        
		<td valign=middle align=left class=forumrow height="22">&nbsp;<%=myusercp%></td>                                                                                                                                                                                        
  	</tr>                                                                                                                                                                                        
  	<tr>                                                                                                                                                                                        
		<td valign=middle align=left class=forumRowHighlight height="22"><b>����:</b></td>                                                                                                                                                                                        
		<td valign=middle align=left class=forumrow height="22">&nbsp;<INPUT name=usercp value=""></td> 
		<td valign=middle align=right class=forumRowHighlight height="22">��������:&nbsp;</td>                                                                                                                                                                                       
		<td valign=middle align=left class=forumrow height="22">&nbsp;<%=myuserep%></td>                                                                                                                                                                                        
  	</tr>                                                                                                                                                                                        
  	<tr>                                                                                                                                                                                        
		<td valign=middle align=left class=forumRowHighlight height="22"><b>����:</b></td>                                                                                                                                                                                        
		<td valign=middle align=left class=forumrow height="22">&nbsp;<INPUT name=userpower value=""></td> 
		<td valign=middle align=right class=forumRowHighlight height="22">��������:&nbsp;</td>                                                                                                                                                                                       
		<td valign=middle align=left class=forumrow height="22">&nbsp;<%=mypower%></td>                                                                                                                                                                                        
  	</tr>                                                                                                                                                                                        
  	<tr>                                                                                                                                                                                        
		<td valign=middle align=left class=forumRowHighlight height="22"><b>������:</b><br>���뱾�ɱ���ﵽ�ķ�����</td>                                                                                                                                                                                        
		<td valign=middle align=left class=forumrow height="22">&nbsp;<INPUT name=userarticle value=""></td> 
		<td valign=middle align=right class=forumRowHighlight height="22">��������:&nbsp;</td>                                                                                                                                                                                       
		<td valign=middle align=left class=forumrow height="22">&nbsp;<%=myarticle%></td>                                                                                                                                                                                        
  	</tr>                                                                                                                                                                                        
  	<tr>                                                                                                                                                                                        
		<td valign=middle align=left class=forumRowHighlight height="22"><b>�û��ȼ�:</b><br>���뱾�ɱ���߱��ĵȼ�(1��<%=maxtitleid%>)</td>                                                                                                                                                                                        
		<td valign=middle align=left class=forumrow height="22">&nbsp;<INPUT name=userclass value=""></td> 
		<td valign=middle align=right class=forumRowHighlight height="22">���ĵȼ�:&nbsp;</td>                                                                                                                                                                                     
		<td valign=middle align=left class=forumrow height="22">&nbsp;<%=myclass%>��<%=GetUserClass(myclass)%>����</td>                                                                                                                                                                                      
  	</tr>
	<tr>
		<td valign=middle align=center class=forumrow height=22 colspan="4"><font color="#CC9933">
		�������������<font color=red>��������</font>��<font color=red>������</font>��������6��(��Ǯ,����,����,����,������,�û��ȼ�)������д</font>
		</td>
	</tr> 
	<tr><td height=8 class=forumRowHighlight colspan="4"></td></tr>                                                                                                                                                                                       
  	<tr>                                                                                                                                                                                        
		<td valign=middle align=left class=forumRowHighlight height="22"><b>���ɼ��:</b><br>�������������򵥵Ľ��ܣ����������������л���ʾ����</td>                                                                                                                                                                                        
		<td valign=middle align=left class=forumrow colspan="3" height="22">&nbsp;<textarea name=menpaireadme rows="4" cols="60"></textarea>		(��Ҫ����100��Ӣ�Ļ�50������)</td>                                                                                                                                                                                      
  	</tr>  
	<tr><td height=8 class=forumRowHighlight colspan="4"></td></tr>                                                                                                                                                                                       
  	<tr>                                                                                                                                                                                       
    	<td valign=middle align=center class=forumrow colspan="4" height="22"><INPUT type=submit value=��� name=submit></td>                                                                                                                                                                                       
  	</tr>                                                                                                                                                                              
  	</form>                                                                                                                                                                                    
</table>                                                                                                                                                                                           
<%                                             
end sub                                                                                                                                                                                                                            

'------------------------------------������ɳɹ�--------------------------------
sub applynewmenpai() 
	
	dim MenpaiAttri,userwealth,usercp,userep,userpower,userclass,userartitle,Zangmen
	GroupName=checkstr(trim(Request.form("groupname")))
	MenpaiAttri=Request.form("menpaiattri")
	Zangmen=trim(Request.form("Zangmen"))
	
	if GroupName="" then
		Errmsg=Errmsg+"<br>"+"<li>��������������"
		founderr=true
	elseif strLength(GroupName)>12 then
		Errmsg=Errmsg+"<br>"+"<li>�������Ʋ��ܳ���12��Ӣ����ĸ(6������)"
		founderr=true		
	end if		
	if cint(MenpaiAttri)<>0 and cint(MenpaiAttri)<>3 then
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
			Errmsg=Errmsg+"<br>"+"<li>���������ɱ�����д�ȼ�����������Ĳ�����Ч����ֵ"
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
	
	if Zangmen<>"" then
		set rs=conn.execute("select usergroup from [user] where username='"&Zangmen&"'")
		if rs.eof and rs.bof then
			Errmsg=Errmsg+"<br>"+"<li>��ָ��Ϊ���ŵ��û�������"
			founderr=true
		end if
	end if
				
	if founderr then exit sub
                           
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
		rs("Zangmen")=Zangmen
		rs("Addtime")=date()
		rs("Menpai_setting")=MenpaiAttri & "|" & userwealth & "|" & userep & "|" & usercp & "|" & userpower & "|" & userartitle & "|" & userclass & "|" & menpaireadme
		if cint(MenpaiAttri)=3 then
			rs("State")=1
			conn.execute("update [GroupName] set State=0 where State=1")
		else
			rs("State")=0
		end if	
		rs.update  
		if Zangmen<>"" then conn.execute("update [user] set usergroup='"&GroupName&"' where username='"&Zangmen&"'")		
	end if
%>
<br><p><b>������ɳɹ�</b>&nbsp;&nbsp;<a href=z_admin_menpai.asp>����</a></p>
<%	 
end sub 
 
sub edit() 
	GroupID=request("ID")
	if GroupID="" and (not isnumeric(GroupID)) then
		Errmsg=Errmsg+"<br>"+"<li>��������ɲ���"
		founderr=true
		exit sub		
	end if
	dim Zangmen,Addtime,Menpai_setting
	sql="select * from [GroupName] where id="&GroupID&""
	set rs=conn.execute(sql)
	if not(rs.eof and rs.bof) then
		GroupName=rs("GroupName")
		Zangmen=trim(rs("Zangmen"))
		Addtime=rs("Addtime")
		Menpai_setting=split(rs("Menpai_setting"),"|")
	else
		Errmsg=Errmsg+"<br>"+"<li>��������ɲ���,ָ�����ɲ�����"
		founderr=true
		exit sub
	end if	
	rs.close
%>   
<table cellspacing=1 align=center class=tableborder width="95%">                                                                                                                                                                              
  	<tr>                                                                                                                                                                                           
    	<td valign=middle align=center height=24 class=forumRowHighlight colspan="4">���ɹ���</td>                                                                                                                                                                                          
  	</tr> 
	<form name="form" method="post" action="z_admin_menpai.asp?action=savedit">                                                                                                                                                                                            
  	<tr>                                                                                                                                                                                           
    	<th valign=middle align=left height=24 colspan="4">&nbsp;��&nbsp; ���ɻ�������</th>                                                                                                                                                                                          
  	</tr> 
	<tr><td height=8 class=forumRowHighlight colspan="4"></td></tr>                                                                                                                                                                                        
  	<tr> 
		<td valign=middle align=left class=forumRowHighlight height="22"><b>��������:</b><br> �����������Ų����޸�</td> 
		<td valign=middle align=left class=forumrow height="22" colspan="3">&nbsp;<INPUT name=GroupName size="20" maxlength="20" value=<%=GroupName%>> <font color=red>*</font> (��Ҫ����12��Ӣ����ĸ��6������)<INPUT name=GroupID size="20" maxlength="20" type=hidden value=<%=GroupID%>></td>
  	</tr>
  	<tr> 
		<td valign=middle align=left class=forumRowHighlight height="22"><b>����:</b><br> ������������κ��������ţ���������</td> 
		<td valign=middle align=left class=forumrow height="22" colspan="3">&nbsp;<INPUT name=Zangmen size="20" maxlength="20" value=<%=Zangmen%>></td>
  	</tr>
	<tr><td height=8 class=forumRowHighlight colspan="4"></td></tr>
    <tr>
		<td valign=middle align=left class=forumRowHighlight height="22"><b>��������˵��:</b></td>                                                                                                                                                                                         
		<td valign=middle align=left class=forumrow height="22" colspan="3">
			<li>�����������������������ɵ��˵�һ����������
			<li>��ע���û��ͱ������߳����ɵ��û��Զ�����������		
		 	<li>ֻ�ܴ���һ�������ɣ����Ե�������һ������Ϊ������֮����ǰ���������Զ�תΪ��������
			<li>ѡ���������ɻ������ɻ������ǰ���õĽ�Ǯ,����,����,����,������,�û��ȼ���ֵ(��Ϊ0)
		</td> 
  	</tr>		
  	<tr> 
		<td valign=middle align=left class=forumRowHighlight height="22"><b>��������:</b><br> ������������֮�����Ų����޸�</td>
		<td valign=middle align=left class=forumrow height="22" colspan="3">
			<input type=radio name=MenpaiAttri value=0 <%if cint(menpai_setting(0))=0 then%>checked<%end if%>>��������&nbsp;
			<input type=radio name=MenpaiAttri value=1 <%if cint(menpai_setting(0))=1 then%>checked<%end if%>>����&nbsp;
			<input type=radio name=MenpaiAttri value=2 <%if cint(menpai_setting(0))=2 then%>checked<%end if%>>ħ��&nbsp;
			<input type=radio name=MenpaiAttri value=3 <%if cint(menpai_setting(0))=3 then%>checked<%end if%>>������
		</td>
  	</tr>
	<tr><td height=8 class=forumRowHighlight colspan="4"></td></tr>
  	<tr> 
		<td valign=middle align=left class=forumRowHighlight height="22"><b>����ʱ��:</b><br> ���ɴ�����ʱ�䣬��ʽ 2002-9-21 </td> 
		<td valign=middle align=left class=forumrow height="22" colspan="3">&nbsp;<INPUT name=Addtime size="20" maxlength="20" value=<%=Addtime%>></td>
  	</tr>	
	<tr><td height=8 class=forumRowHighlight colspan="4"></td></tr>
    <tr>
		<td valign=middle align=left class=forumRowHighlight height="22"><b>����ֵ˵��:</b></td>                                                                                                                                                                                         
		<td valign=middle align=left class=forumrow height="22" colspan="3">
		 	<li>��Ǯ,����,����,����,������,�û��ȼ� �����ܳ��������ѵĵ�ǰֵ
			<li>��Ǯ,����,����,����,������,�û��ȼ� �Ǽ��뱾�ɵı����������Сֵ
			<li>��Ǯ,����,����,����,������,�û��ȼ� Ҳ�Ǽ�����˳���������/�۳����ֵʱ�Ļ���
			<li>�������ɺ������� ������������
		</td> 
  	</tr> 
    <tr>                                                                                                                                                                                         
		<td valign=middle align=left class=forumRowHighlight height="22" width="30%"><b>��Ǯ:</b></td>                                                                                                                                                                                        
		<td valign=middle align=left class=forumrow height="22">&nbsp;<INPUT name=userwealth value=<%=menpai_setting(1)%>></td>                                                                                                                                                                                        
		<td valign=middle align=right class=forumRowHighlight height="22">���Ľ�Ǯ:&nbsp;</td>
		<td valign=middle align=left class=forumrow height="22" width="20%">&nbsp;<%=mymoney%></td>
  	</tr>                                                                                                                                                                                        
  	<tr>                                                                                                                                                                                        
		<td valign=middle align=left class=forumRowHighlight height="22"><b>����:</b></td>                                                                                                                                                                                        
		<td valign=middle align=left class=forumrow height="22">&nbsp;<INPUT name=userep value=<%=menpai_setting(2)%>></td>
		<td valign=middle align=right class=forumRowHighlight height="22">���ľ���:&nbsp;</td>                                                                                                                                                                                        
		<td valign=middle align=left class=forumrow height="22">&nbsp;<%=myusercp%></td>                                                                                                                                                                                        
  	</tr>                                                                                                                                                                                        
  	<tr>                                                                                                                                                                                        
		<td valign=middle align=left class=forumRowHighlight height="22"><b>����:</b></td>                                                                                                                                                                                        
		<td valign=middle align=left class=forumrow height="22">&nbsp;<INPUT name=usercp value=<%=menpai_setting(3)%>></td> 
		<td valign=middle align=right class=forumRowHighlight height="22">��������:&nbsp;</td>                                                                                                                                                                                       
		<td valign=middle align=left class=forumrow height="22">&nbsp;<%=myuserep%></td>                                                                                                                                                                                        
  	</tr>                                                                                                                                                                                        
  	<tr>                                                                                                                                                                                        
		<td valign=middle align=left class=forumRowHighlight height="22"><b>����:</b></td>                                                                                                                                                                                        
		<td valign=middle align=left class=forumrow height="22">&nbsp;<INPUT name=userpower value=<%=menpai_setting(4)%>></td> 
		<td valign=middle align=right class=forumRowHighlight height="22">��������:&nbsp;</td>                                                                                                                                                                                       
		<td valign=middle align=left class=forumrow height="22">&nbsp;<%=mypower%></td>                                                                                                                                                                                        
  	</tr>                                                                                                                                                                                        
  	<tr>                                                                                                                                                                                        
		<td valign=middle align=left class=forumRowHighlight height="22"><b>������:</b><br>���뱾�ɱ���ﵽ�ķ�����</td>                                                                                                                                                                                        
		<td valign=middle align=left class=forumrow height="22">&nbsp;<INPUT name=userarticle value=<%=menpai_setting(5)%>></td> 
		<td valign=middle align=right class=forumRowHighlight height="22">��������:&nbsp;</td>                                                                                                                                                                                       
		<td valign=middle align=left class=forumrow height="22">&nbsp;<%=myarticle%></td>                                                                                                                                                                                        
  	</tr>                                                                                                                                                                                        
  	<tr>                                                                                                                                                                                        
		<td valign=middle align=left class=forumRowHighlight height="22"><b>�û��ȼ�:</b><br>���뱾�ɱ���߱��ĵȼ�(1��<%=maxtitleid%>)</td>                                                                                                                                                                                        
		<td valign=middle align=left class=forumrow height="22">&nbsp;<INPUT name=userclass value=<%=menpai_setting(6)%>>&nbsp;<%=GetUserTitle(Menpai_setting(6))%>(<%=Menpai_setting(6)%>��)</td> 
		<td valign=middle align=right class=forumRowHighlight height="22">���ĵȼ�:&nbsp;</td>                                                                                                                                                                                     
		<td valign=middle align=left class=forumrow height="22">&nbsp;<%=myclass%>��<%=GetUserClass(myclass)%>����</td>                                                                                                                                                                                      
  	</tr>
	<tr>
		<td valign=middle align=center class=forumrow height=22 colspan="4"><font color="#CC9933">
		�������������<font color=red>��������</font>��<font color=red>������</font>��������6��(��Ǯ,����,����,����,������,�û��ȼ�)������д</font>
		</td>
	</tr> 
	<tr><td height=8 class=forumRowHighlight colspan="4"></td></tr>                                                                                                                                                                                       
  	<tr>                                                                                                                                                                                        
		<td valign=middle align=left class=forumRowHighlight height="22"><b>���ɼ��:</b><br>�������������򵥵Ľ��ܣ����������������л���ʾ����</td>                                                                                                                                                                                        
		<td valign=middle align=left class=forumrow colspan="3" height="22">&nbsp;<textarea name=menpaireadme rows="4" cols="60"><%=menpai_setting(7)%></textarea>		(��Ҫ����100��Ӣ�Ļ�50������)</td>                                                                                                                                                                                      
  	</tr>  
	<tr><td height=8 class=forumRowHighlight colspan="4"></td></tr>                                                                                                                                                                                       
  	<tr>                                                                                                                                                                                       
    	<td valign=middle align=center class=forumrow colspan="4" height="22"><INPUT type=submit value=���� name=submit></td>                                                                                                                                                                                       
  	</tr>                                                                                                                                                                              
  	</form>     
</table>    
                                                                                                                                                                           
<%                                                                                                                                                                                                                             
end sub

sub savedit()
	                                                                                                                                                   
	dim MenpaiAttri,userwealth,usercp,userep,userpower,userclass,userartitle,Addtime,Zangmen
	GroupID=Request.form("GroupID")
	GroupName=checkstr(trim(Request.form("GroupName")))
	MenpaiAttri=Request.form("MenpaiAttri")
	Zangmen=trim(Request.form("Zangmen"))
	Addtime=Request.form("Addtime")
	
	if GroupName="" then
		Errmsg=Errmsg+"<br>"+"<li>��������������"
		founderr=true
	elseif strLength(GroupName)>12 then
		Errmsg=Errmsg+"<br>"+"<li>�������Ʋ��ܳ���12��Ӣ����ĸ(6������)"
		founderr=true		
	end if		
	if cint(MenpaiAttri)<>0 and cint(MenpaiAttri)<>3 then
		if not isInteger(Request.form("userwealth")) then
			Errmsg=Errmsg+"<br>"+"<li>���������ɱ�����д��Ǯֵ����������Ĳ�����Ч����ֵ"
			founderr=true
		else
			userwealth=int(Request.form("userwealth"))
		end if	
		if not isInteger(Request.form("userep")) then
			Errmsg=Errmsg+"<br>"+"<li>���������ɱ�����д����ֵ����������Ĳ�����Ч����ֵ"
			founderr=true
		else 
			userep=int(Request.form("userep"))	
		end if
		if not isInteger(Request.form("usercp")) then
			Errmsg=Errmsg+"<br>"+"<li>���������ɱ�����д����ֵ����������Ĳ�����Ч����ֵ"
			founderr=true
		else 
			usercp=int(Request.form("usercp"))	
		end if
		if not isInteger(Request.form("userpower")) then
			Errmsg=Errmsg+"<br>"+"<li>���������ɱ�����д����ֵ����������Ĳ�����Ч����ֵ"
			founderr=true
		else 
			userpower=int(Request.form("userpower"))	
		end if
		if not isInteger(Request.form("userarticle")) then
			Errmsg=Errmsg+"<br>"+"<li>���������ɱ�����д����������������Ĳ�����Ч����ֵ"
			founderr=true
		else 
			userartitle=int(Request.form("userarticle"))	
		end if
		if not isInteger(Request.form("userclass")) then
			Errmsg=Errmsg+"<br>"+"<li>���������ɱ�����д����ֵ����������Ĳ�����Ч����ֵ"
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
	
	if Zangmen<>"" then
		set rs=conn.execute("select usergroup from [user] where username='"&Zangmen&"'")
		if rs.eof and rs.bof then
			Errmsg=Errmsg+"<br>"+"<li>��ָ��Ϊ���ŵ��û�������"
			founderr=true
		end if
	end if				 	
	
	dim menpaireadme
	menpaireadme=Request.form("menpaireadme")	
	if strLength(menpaireadme)>100 then
		Errmsg=Errmsg+"<br>"+"<li>���ɼ�鲻�ܳ���100��Ӣ�Ļ�50������"
		founderr=true		
	end if
			
	if founderr then exit sub
	
	set rs= server.createobject ("adodb.recordset")                                                                                                                                                                                                                                   
	sql="select * from [GroupName] where ID="&GroupID&""                                                                                                                                                                                                                                   
	rs.Open sql,conn,1,3  
	 
	if rs.bof and rs.eof then                                                                          
		Errmsg=Errmsg+"<br>"+"<li>�����ɲ����ڰ�"
		founderr=true
		rs.close
		set rs=nothing
		exit sub
	else                                  
		rs("Groupname")=GroupName
		rs("Zangmen")=Zangmen
		if cint(MenpaiAttri)=3 then
			rs("State")=1
			conn.execute("update [GroupName] set State=0 where State=1")
		else
			rs("State")=0
		end if		
		rs("Menpai_setting")=MenpaiAttri & "|" & userwealth & "|" & userep & "|" & usercp & "|" & userpower & "|" & userartitle & "|" & userclass & "|" & menpaireadme
		if isdate(Addtime) then rs("Addtime")=Addtime
		rs.update  
		if Zangmen<>"" then conn.execute("update [user] set usergroup='"&GroupName&"' where username='"&Zangmen&"'")
	end if
%>
<br><br><p><b>�޸����ɳɹ�</b>&nbsp;&nbsp;<a href=z_admin_menpai.asp>����</a></p>
<%
end sub

sub editapply() 
	dim newmenpai_setting
	newmenpai_setting=split(conn.execute("select new_menpai_setting from [menpai]")(0),",")
%>   
<br>
<table cellspacing=1 align=center class=tableborder width="95%">                                                                                                                                                                              
  	<tr>                                                                                                                                                                                           
    	<th valign=middle height=24 colspan="4">����������������</td>                                                                                                                                                                                          
  	</tr> 
	<tr><td height=8 class=forumRowHighlight colspan="4"></td></tr>        
	<form name="form" method="post" action="z_admin_menpai.asp?action=saveditapply">                                                                                                                                                                                 
    <tr>
		<td valign=middle align=left class=forumRowHighlight height="22"><b>˵��:</b></td>                                                                                                                                                                                         
		<td valign=middle align=left class=forumrow height="22" colspan="3">
			<li>ǰ̨�û������Լ��������ɱ������������
			<li>������ò�Ҫ���õ�̫�ͣ�����ܶ��˶����Ŵ���Ҫ��
			<li>��������ֵ�������ο��ģ������ֵ������������ֵ������
			<li>����������,��Ҫ���븺������ĸ		
		</td> 
  	</tr> 
	<tr><td height=8 class=forumRowHighlight colspan="4"></td></tr>
    <tr>                                                                                                                                                                                         
		<td valign=middle align=left class=forumRowHighlight height="22" width="30%"><b>��Ǯ:</b></td>                                                                                                                                                                                        
		<td valign=middle align=left class=forumrow height="22">&nbsp;<INPUT name=userwealth value=<%=newmenpai_setting(0)%>></td>                                                                                                                                                                                        
		<td valign=middle align=right class=forumRowHighlight height="22">���Ľ�Ǯ:&nbsp;</td>
		<td valign=middle align=left class=forumrow height="22" width="20%">&nbsp;<%=mymoney%></td>
  	</tr>                                                                                                                                                                                        
  	<tr>                                                                                                                                                                                        
		<td valign=middle align=left class=forumRowHighlight height="22"><b>����:</b></td>                                                                                                                                                                                        
		<td valign=middle align=left class=forumrow height="22">&nbsp;<INPUT name=userep value=<%=newmenpai_setting(2)%>></td>
		<td valign=middle align=right class=forumRowHighlight height="22">���ľ���:&nbsp;</td>                                                                                                                                                                                        
		<td valign=middle align=left class=forumrow height="22">&nbsp;<%=myusercp%></td>                                                                                                                                                                                        
  	</tr>                                                                                                                                                                                        
  	<tr>                                                                                                                                                                                        
		<td valign=middle align=left class=forumRowHighlight height="22"><b>����:</b></td>                                                                                                                                                                                        
		<td valign=middle align=left class=forumrow height="22">&nbsp;<INPUT name=usercp value=<%=newmenpai_setting(1)%>></td> 
		<td valign=middle align=right class=forumRowHighlight height="22">��������:&nbsp;</td>                                                                                                                                                                                       
		<td valign=middle align=left class=forumrow height="22">&nbsp;<%=myuserep%></td>                                                                                                                                                                                        
  	</tr>                                                                                                                                                                                        
  	<tr>                                                                                                                                                                                        
		<td valign=middle align=left class=forumRowHighlight height="22"><b>����:</b></td>                                                                                                                                                                                        
		<td valign=middle align=left class=forumrow height="22">&nbsp;<INPUT name=userpower value=<%=newmenpai_setting(3)%>></td> 
		<td valign=middle align=right class=forumRowHighlight height="22">��������:&nbsp;</td>                                                                                                                                                                                       
		<td valign=middle align=left class=forumrow height="22">&nbsp;<%=mypower%></td>                                                                                                                                                                                        
  	</tr>                                                                                                                                                                                        
  	<tr>                                                                                                                                                                                        
		<td valign=middle align=left class=forumRowHighlight height="22"><b>������:</b><br>���뱾�ɱ���ﵽ�ķ�����</td>                                                                                                                                                                                        
		<td valign=middle align=left class=forumrow height="22">&nbsp;<INPUT name=userarticle value=<%=newmenpai_setting(4)%>></td> 
		<td valign=middle align=right class=forumRowHighlight height="22">��������:&nbsp;</td>                                                                                                                                                                                       
		<td valign=middle align=left class=forumrow height="22">&nbsp;<%=myarticle%></td>                                                                                                                                                                                        
  	</tr> 
  	<tr>                                                                                                                                                                                        
		<td valign=middle align=left class=forumRowHighlight height="22"><b>�û��ȼ�:</b><br>���뱾�ɱ���߱��ĵȼ�(1��<%=maxtitleid%>)</td>                                                                                                                                                                                        
		<td valign=middle align=left class=forumrow height="22">&nbsp;<INPUT name=userclass value=<%=newmenpai_setting(5)%>>&nbsp;<%=GetUserTitle(newmenpai_setting(5))%>(<%=newmenpai_setting(5)%>��)</td> 
		<td valign=middle align=right class=forumRowHighlight height="22">���ĵȼ�:&nbsp;</td>                                                                                                                                                                                     
		<td valign=middle align=left class=forumrow height="22">&nbsp;<%=myclass%>��<%=GetUserClass(myclass)%>����</td>                                                                                                                                                                                      
  	</tr>
	<tr><td height=8 class=forumRowHighlight colspan="4"></td></tr>  
  	<tr>                                                                                                                                                                                        
		<td valign=middle align=left class=forumRowHighlight height="22"><b>�����շ�:</b><br>��������ʱ��ȡ�ķ���</td>                                                                                                                                                                                        
		<td valign=middle align=left class=forumrow height="22" colspan="3">&nbsp;<INPUT name=creatmoney value=<%=newmenpai_setting(6)%>>
		&nbsp;<font color=navy>ע�⣺ǰ̨���������еĽ�Ǯ�� = ǰ��Ľ�Ǯ + <font color=red>�����շ�</font></font>
		</td> 
  	</tr>
	<tr><td height=8 class=forumRowHighlight colspan="4"></td></tr>  
  	<tr>                                                                                                                                                                                         
		<td valign=middle align=left class=forumRowHighlight height="22"><b>���ѡ��:</b><br>�û���������֮���Ƿ���Ҫͨ������Ա���֮�������ʽ����</td>                                                                                                                                                                                         
		<td valign=middle align=left class=forumrow height="22" colspan="3">&nbsp;
			<input type=radio name=shenhe value=0 <%if cint(newmenpai_setting(7))=0 then%>checked<%end if%>>ֱ�ӿ�ͨ&nbsp;
			<input type=radio name=shenhe value=1 <%if cint(newmenpai_setting(7))=1 then%>checked<%end if%>>��Ҫͨ����˺���ܿ�ͨ&nbsp;
		</td> 
  	</tr>
	<tr><td height=8 class=forumRowHighlight colspan="4"></td></tr>  		                                                                                                                                                                                     
  	<tr>                                                                                                                                                                                       
    	<td valign=middle align=center class=forumrow colspan="4" height="22"><INPUT type=submit value=���� name=submit></td>                                                                                                                                                                                       
  	</tr>                                                                                                                                                                              
  	</form>     
</table>    
<%                                                                                                                                                                                                                             
end sub

sub saveditapply()
	dim maxtitleid,creatmoney
	dim MenpaiAttri,userwealth,usercp,userep,userpower,userclass,userartitle,Addtime,Zangmen

	if not isInteger(Request.form("userwealth")) then
		Errmsg=Errmsg+"<br>"+"<li>������Ľ�Ǯֵ������Ч����ֵ"
		founderr=true
	else
		userwealth=int(Request.form("userwealth"))
	end if	
	if not isInteger(Request.form("userep")) then
		Errmsg=Errmsg+"<br>"+"<li>������ľ���ֵ������Ч����ֵ"
		founderr=true
	else 
		userep=int(Request.form("userep"))	
	end if
	if not isInteger(Request.form("usercp")) then
		Errmsg=Errmsg+"<br>"+"<li>�����������ֵ������Ч����ֵ"
		founderr=true
	else 
		usercp=int(Request.form("usercp"))	
	end if
	if not isInteger(Request.form("userpower")) then
		Errmsg=Errmsg+"<br>"+"<li>�����������ֵ������Ч����ֵ"
		founderr=true
	else 
		userpower=int(Request.form("userpower"))	
	end if
	if not isInteger(Request.form("userarticle")) then
		Errmsg=Errmsg+"<br>"+"<li>�����������ֵ������Ч����ֵ"
		founderr=true
	else 
		userartitle=int(Request.form("userarticle"))	
	end if
	maxtitleid=conn.execute("select top 1 UserTitleID from [UserTitle] order by UserTitleID desc")(0)
	if not isInteger(Request.form("userclass")) then
		Errmsg=Errmsg+"<br>"+"<li>������ĵȼ�ֵ������Ч����ֵ"
		founderr=true
	elseif int(Request.form("userclass"))<1 or int(Request.form("userclass"))>int(maxtitleid) then
		Errmsg=Errmsg+"<br>"+"<li>�ȼ�ֵ����С��1�����"&maxtitleid
		founderr=true						
	else 
		userclass=int(Request.form("userclass"))	
	end if
	if not isInteger(Request.form("creatmoney")) then
		Errmsg=Errmsg+"<br>"+"<li>������Ĵ��ɷ��ò�����Ч����ֵ"
		founderr=true
	else
		creatmoney=int(Request.form("creatmoney"))
	end if	
	
	if founderr then exit sub
	conn.execute("update menpai set new_menpai_setting='" & userwealth & "," & usercp & "," & userep & "," & userpower & "," & userartitle & "," & userclass & "," & creatmoney & "," & Request.form("shenhe") & "'")
%>
<br><br><p>&nbsp;<b>�޸Ĵ��������ɹ�</b>&nbsp;&nbsp;<a href=z_admin_menpai.asp>����</a></p>
<%	
end sub
'------------------ɾ������----------------------
sub del()
	conn.execute("delete from GroupName where id="&request("id"))
	if request("Zangmen")<>"" then call sendmessage(request("Zangmen"),membername,"�������֪ͨ","���Ǳ�Ǹ�����´���������û��ͨ�����")
%>
<br><br><p><b>ɾ�����ɳɹ�</b>&nbsp;&nbsp;<a href=z_admin_menpai.asp>����</a></p>
<%
end sub
'------------------��׼���ɳ���------------------
sub pass()
	conn.execute("update GroupName set state=0 where id="&request("id"))
	if request("Zangmen")<>"" then	call sendmessage(request("Zangmen"),membername,"���ɿ�֪ͨͨ","������ˣ��������������Ѿ����Կ�ͨ�ˣ�ϣ�����ܹ�������������ɣ�лл!")
%>
<br><br><p><b>����׼�˸����ɳ���</b>&nbsp;&nbsp;<a href=z_admin_menpai.asp>����</a></p>
<%	
end sub 
'------------------���¿�ͨ�Ѿ�ע��������------------------
sub repass()
	conn.execute("update GroupName set state=0,zangmen='' where id="&request("id"))
%>
<br><br><p><b>�����ؿ��ɹ����������ɹ������޸����ɵĻ������Ϻ�����</b>&nbsp;&nbsp;<a href=z_admin_menpai.asp>����</a></p>
<%	
end sub 
'--------------------------------���Ͷ���Ϣ--------------------------------
sub sendmessage(incept,sender,title,message)
	sql="insert into message (incept,sender,title,content,sendtime,flag,issend) values ('"&incept&"','"&sender&"','"&title&"','"&message&"',Now(),0,1)"                                          
	conn.execute(sql)                                          
end sub 
%>