<!--#include file="conn.asp"-->
<!-- #include file="inc/const.asp" -->
<%
'=========================================================
' File: toplist.asp
' Version:5.0
' Date: 2002-9-7
' Script Written by satan
'=========================================================
' Copyright (C) 2001,2002 AspSky.Net. All rights reserved.
' Web: http://www.aspsky.net,http://www.dvbbs.net
' Email: info@aspsky.net,eway@aspsky.net
'=========================================================

dim orders,ordername
dim currentpage,page_count,Pcount
dim totalrec,endpage
dim bbsnum,usernum
dim select1,select2,select3,select4,select5,select6,select7,select8,select9,select10,select11

if Cint(GroupSetting(1))=0 then
	Errmsg=Errmsg+"<br>"+"<li>��û���������̳��Ա���ϵ�Ȩ�ޣ���<a href=login.asp>��¼</a>����ͬ����Ա��ϵ��"
	founderr=true
end if

if founderr then
	call nav()
	call head_var(2,0,"","")
	call dvbbs_error()
else
	currentPage=request("page")
	if not isInteger(request("orders")) or request("orders")="" then
		orders=1
	else
		orders=request("orders")
	end if
	select case orders
	case 1
		orders=1
		ordername="��������Top" & Forum_Setting(68)
		select1="selected"
	case 2
		orders=2
		ordername="�µ������б�"
		select2="selected"
	case 3
		orders=3
		ordername=Forum_Setting(68) & "����"
		select3="selected"
	case 4
		orders=4
		ordername=Forum_Setting(68) & "�����"
		select4="selected"
	case 5
		orders=5
		ordername=Forum_Setting(68) & "����ʿ"
		select5="selected"
	case 6
		orders=6
		ordername=Forum_Setting(68) & "�����"
		select6="selected"
	case 7
		orders=7
		ordername="���������б�"
		select7="selected"
	case 8
		orders=8
		ordername="�����Ŷ�"
		select8="selected"
	case 9
		orders=9
		ordername="VIP�û��б�"
		select9="selected"
	case 10
		orders=10
		ordername="�ȴ�VIP������Ա"
		select10="selected"
	case 11
		orders=11
		ordername="������VIP�ʸ��Ա"
		select11="selected"
	case else
		orders=1
		ordername="��������Top" & Forum_Setting(68)
		select1="selected"
	end select
	
	stats=ordername
	call nav()
	call head_var(2,0,"","")
	call main()
end if
call footer()

sub main()

	set rs=conn.execute("select top 1 BbsNum,UserNum from config where active=1")
	bbsnum=rs(0)
	usernum=rs(1)
%>
<table cellpadding=6 cellspacing=1 align=center class=tableborder1>
<form method=POST action=toplist.asp>
<tr>
<td colspan=5 valign=top width=350 class=tabletitle2>&nbsp;>> <B><%=ordername%></B> <<<BR><BR>
&nbsp;��ע���û����� <%=usernum%> �� &nbsp; ���������� <%=bbsnum%> ƪ</td>
<td colspan=6 align=right class=tabletitle2>
<select name=orders onchange='javascript:submit()'>
<option value=1 <%=select1%>>��������Top<%=Forum_Setting(68)%></option>
<option value=2 <%=select2%>>����ע���û�</option>
<option value=3 <%=select3%>><%=Forum_Setting(68)%>����</option>
<option value=4 <%=select4%>><%=Forum_Setting(68)%>�����</option>
<option value=5 <%=select5%>><%=Forum_Setting(68)%>����ʿ</option>
<option value=6 <%=select6%>><%=Forum_Setting(68)%>�����</option>
<option value=7 <%=select7%>>���������б�</option>
<option value=8 <%=select8%>>�����Ŷ�</option>
<option value=9 <%=select9%>>��ʽVIP�б�</option>
<option value=10 <%=select10%>>�ȴ�VIP������Ա</option>
<option value=11 <%=select11%>>������VIP�ʸ��Ա</option>
</select>
</td></tr></form>
<tr> 
<th>�û���</th>
<th>Email</th>
<th>OICQ</th>
<th>��ҳ</th>
<th>����Ϣ</th>
<th>ע��ʱ��</th>
<th>�ȼ�״̬</th>
<th>��������</th>
<th>�Ʋ�</th>
<th>����</th>
<th>����</th></tr>
<%
	if currentpage="" or not isInteger(currentpage) then
		currentpage=1
	else
		currentpage=clng(currentpage)
		if err then
			currentpage=1
			err.clear
		end if
	end if
	set rs=server.createobject("adodb.recordset")
	select case orders
	case 1
		sql="select top "&Forum_Setting(68)&" username,useremail,userclass,oicq,homepage,article,addDate,userwealth as wealth,userid,userpower,userscore from [user] order by article desc"
	case 2
		sql="select top "&Forum_Setting(68)&" username,useremail,userclass,oicq,homepage,article,addDate,userwealth as wealth,userid,userpower,userscore from [user] order by AddDate desc"
	case 3
		sql="select top "&Forum_Setting(68)&" username,useremail,userclass,oicq,homepage,article,addDate,userwealth as wealth,userid,userpower,userscore from [user] order by userwealth desc"
	case 4
		sql="select top "&Forum_Setting(68)&" username,useremail,userclass,oicq,homepage,article,addDate,userwealth as wealth,userid,userpower,userscore from [user] order by userscore desc"
	case 5
		sql="select top "&Forum_Setting(68)&" username,useremail,userclass,oicq,homepage,article,addDate,userwealth as wealth,userid,userpower,userscore from [user] order by userpower desc"
	case 6
		sql="select top "&Forum_Setting(68)&" username,useremail,userclass,oicq,homepage,article,addDate,userwealth as wealth,userid,userpower,userscore from [user] order by userpower asc"
	case 7
		sql="select username,useremail,userclass,oicq,homepage,article,addDate,userwealth as wealth,userid,userpower,userscore from [user] order by userid desc"
	case 8
		sql="select username,useremail,userclass,oicq,homepage,article,addDate,userwealth as wealth,userid,userpower,userscore from [user] where usergroupid<=3 order by usergroupid,article desc"
	case 9
		sql="select username,useremail,userclass,oicq,homepage,article,addDate,userwealth as wealth,userid,userpower,userscore from [user] where vip=1 order by usergroupid,article desc"
	case 10
		sql="select username,useremail,userclass,oicq,homepage,article,addDate,userwealth as wealth,userid,userpower,userscore from [user] where vip=4 order by usergroupid,article desc"
	case 11
		sql="select username,useremail,userclass,oicq,homepage,article,addDate,userwealth as wealth,userid,userpower,userscore from [user] where vip=3 order by usergroupid,article desc"
	case else
		sql="select top "&Forum_Setting(68)&" username,useremail,userclass,oicq,homepage,article,addDate,userwealth as wealth,userid,userpower,userscore from [user] order by article desc"
	end select
	rs.open sql,conn,1,1
	if rs.eof and rs.bof then
		response.write "<tr><td colspan=11 class=tablebody1>������û���κ��û����ݡ�</font></td></tr>"
	else
	if orders=7 then
	totalrec=userNum
	else
	totalrec=rs.recordcount
	end if
	if totalrec mod Forum_Setting(11)=0 then
   		Pcount= totalrec \ Forum_Setting(11)
	else
   		Pcount= totalrec \ Forum_Setting(11)+1
	end if
	if currentpage > Pcount then currentpage = Pcount
 	if currentpage<1 then currentpage=1
	rs.PageSize = Cint(Forum_Setting(11))
	rs.AbsolutePage=currentpage
	page_count=0
	do while not rs.eof and page_count < Clng(Forum_Setting(11))
%>
<tr bgcolor=<%=Forum_body(4)%> height=20>
<td class=tablebody1>&nbsp;<a href=dispuser.asp?id=<%=rs("userid")%> target=_blank><%=rs("username")%></a></td>
<td align=center class=tablebody2><a href=mailto:<%=rs("useremail")%>><img border=0 src=<%=Forum_info(7)&Forum_TopicPic(10)%>></a></td>
<td align=center class=tablebody1><%=iimg(rs("oicq"),"û��","<a href=http://search.tencent.com/cgi-bin/friend/user_show_info?ln="&rs("oicq")&" target=_blank><img src="&Forum_info(7)&Forum_TopicPic(11)&" alt='�鿴 OICQ:"&rs("oicq")&" ������' border=0 ></a>")%></td>
<td align=center class=tablebody2><%=iimg(rs("homepage"),"û��","<a href="&rs("homepage")&" target=_blank><img border=0 src="&Forum_info(7)&Forum_TopicPic(14)&"></a>")%></td>
<td align=center class=tablebody1><a href=javascript:openScript('messanger.asp?action=new&touser=<%=htmlencode(rs("username"))%>',500,400)><img src=<%=Forum_info(7)&Forum_TopicPic(7)%> border=0></a></td>
<td align=center class=tablebody2><%=rs("addDate")%></td>
<td align=center class=tablebody1> <%=rs("userclass")%></td>
<td align=center class=tablebody2><%=rs("article")%></td>
<td align=center class=tablebody1><%=rs("wealth")%></td>
<td align=center class=tablebody1><%=rs("userpower")%></td>
<td align=center class=tablebody2><%=rs("userscore")%></td>
</tr>
<%
		page_count=page_count+1
		if orders>=1 and orders<=6 and page_count>=clng(forum_setting(68)) then exit do
		rs.movenext
	loop
	end if
%>
</table>
<table border=0 cellpadding=0 cellspacing=3 width="<%=Forum_body(12)%>" align=center>
<tr><td valign=middle nowrap>
<font color=<%=Forum_body(13)%>>ҳ�Σ�<b><%=currentpage%></b>/<b><%=Pcount%></b>ҳ
&nbsp;
ÿҳ<b><%=Forum_Setting(11)%></b> ����<b><%=totalrec%></b></font></td>
<td valign=middle nowrap align=right>��ҳ��<%call DispPageNum(currentpage,PCount,"""?page=","&orders="&orders&"""")%></td></tr></table>
<%
	rs.close
	set rs=nothing
	call activeonline()
end sub
%>
