<!--#include file="conn.asp"-->
<!-- #include file="inc/const.asp" -->
<%dim chenbackasp,chentitle
	chenbackasp="z_menpai.asp"
	chentitle="��̳����"
	stats="��̳���� ���ɱ�ƴ"
	
	call nav()
	call head_var(2,0,"","")
	if not founduser then
		Errmsg=Errmsg+"<br>"+"<li>��û�������ɱ�ƴ��Ȩ�ޣ���<a href=login.asp>��¼</a>����ͬ����Ա��ϵ��"
		call dvbbs_error()
	else
		call main()	
	end if
	call footer()
'=====================================================
sub main()
	dim rs1
	dim Article,Wealth,Usernum,Classnum
	dim TotalArticle,TotalWealth,TotalUsernum,TotalClassnum
	dim ArticleRate,WealthRate,UsernumRate,ClassnumRate
			
	sql="select * from [GroupName] where state<>2 and state<>3"        '��ע���͵ȴ���˵����ɲ������ƴ
	set rs=conn.execute(sql)
	if rs.eof and rs.bof then
		response.write "<table cellspacing=1 align=center class=tableborder1 width="&Forum_body(12)&">"
		response.Write "<tr><td valign=middle align=left class=tablebody1 height=20 colspan=5>��ʱû����������</td></tr>"
		response.Write "</table>"
	else
		TotalArticle=0:		TotalWealth=0:		TotalUsernum=0:		TotalClassnum=0
		
		set rs1=conn.execute("select sum(article),sum(userwealth),count(*) from [user]")
		if rs1(0)<>"" then TotalArticle=rs1(0)
		if rs1(1)<>"" then TotalWealth=rs1(1)
		if rs1(2)<>"" then TotalUsernum=rs1(2)
		
		set rs1=conn.execute("select count(*) from [user] where usergroupid>=1 and usergroupid<=3")
		if rs1(0)<>"" then TotalClassnum=rs1(0)
		
		response.Write "<br><table cellspacing=1 align=center class=tableborder1 width="&Forum_body(12)&">"
			response.write "<tr>"
				response.write "<td valign=middle class=tablebody2 height=25>&nbsp;��̳ͳ�� --> �����Ŷ�:"&TotalClassnum&"��&nbsp;&nbsp;�û�����:"&TotalUsernum&"��&nbsp;&nbsp;��������:"&TotalArticle&"��&nbsp;&nbsp;�ֽ�:"&TotalWealth&"Ԫ</td>"
			response.write "</tr>"
		response.Write "</table><br>"
			
		response.write "<table cellspacing=1 align=center class=tableborder1 width="&Forum_body(12)&">"
		response.write "<tr>" 
			response.write "<td valign=middle align=center class=tabletitle2 height=22>��������</td>" 
			response.write "<td valign=middle align=center class=tabletitle2 height=22>��������</td>"
			response.write "<td valign=middle align=center class=tabletitle2 height=22>������</td>"
			response.write "<td valign=middle align=center class=tabletitle2 height=22>�ֽ���</td>"
			response.write "<td valign=middle align=center class=tabletitle2 height=22>������</td>"
		response.write "</tr>"
							
		do while not rs.eof
			Article=0:		Wealth=0:		Usernum=0:		Classnum=0
			
			set rs1=conn.execute("select sum(article),sum(userwealth),count(*) from [user] where usergroup='"&rs("GroupName")&"'")
			if rs1(0)<>"" then Article=rs1(0)
			if rs1(1)<>"" then Wealth=rs1(1)
			if rs1(2)<>"" then Usernum=rs1(2)
			
			set rs1=conn.execute("select count(*) from [user] where usergroup='"&rs("GroupName")&"' and (usergroupid>=1 and usergroupid<=3)")
			if rs1(0)<>"" then Classnum=rs1(0)
			
			ClassnumRate=Classnum*120/TotalClassnum+5
			UsernumRate=Usernum*150/TotalUsernum+5
			WealthRate=Wealth*200/TotalWealth+5
			ArticleRate=Article*200/TotalArticle+5
			
			response.write "<tr>"
				response.write "<td valign=middle align=center class=tablebody2 height=25 width=""15%"">"&rs("GroupName")&"</td>"
				response.write "<td valign=middle class=tablebody1 height=25 width=""15%""><img src=pic/bar1.gif width="""&ClassnumRate&""" height=8 alt="&Classnum&"></td>"
				response.write "<td valign=middle class=tablebody1 height=25 width=""20%""><img src=pic/bar2.gif width="""&UsernumRate&""" height=8 alt="&Usernum&"></td>"
				response.write "<td valign=middle class=tablebody1 height=25 width=""25%""><img src=pic/bar3.gif width="""&WealthRate&""" height=8 alt="&wealth&"></td>" 
				response.write "<td valign=middle class=tablebody1 height=25 width=""25%""><img src=pic/bar4.gif width="""&ArticleRate&""" height=8 alt="&Article&"></td>"
			response.write "</tr>"
			rs.movenext
		loop
		response.Write "</table>"
	end if
	
	response.Write "<br><table cellspacing=1 align=center class=tableborder1 width="&Forum_body(12)&">"
		response.write "<tr>"
			response.write "<td valign=middle align=center class=tablebody2 height=25 width=""15%"">�ο���<�ܺ�></td>"
			response.write "<td valign=middle class=tablebody2 height=25 width=""15%""><img src=pic/bar1.gif width=125 height=8 alt="&TotalClassnum&"></td>"
			response.write "<td valign=middle class=tablebody2 height=25 width=""20%""><img src=pic/bar2.gif width=155 height=8 alt="&TotalUsernum&"></td>"
			response.write "<td valign=middle class=tablebody2 height=25 width=""25%""><img src=pic/bar3.gif width=205 height=8 alt="&Totalwealth&"></td>" 
			response.write "<td valign=middle class=tablebody2 height=25 width=""25%""><img src=pic/bar4.gif width=205 height=8 alt="&TotalArticle&"></td>"
		response.write "</tr>"
	response.Write "</table>"	
end sub
%>