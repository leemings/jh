<!--#include file="conn.asp"-->
<!-- #include file="inc/const.asp" -->
<!-- #include file="inc/ubbcode.asp" -->
<%
dim currentpage,page_count,Pcount
dim totalrec,endpage
dim stype,pSearch,nSearch,keyword
dim searchboard,ordername
dim searchdate,searchday
dim stable
dim scoreuser,searchscore,searchscoreuser
stats="�������"
stype=request("stype")
pSearch=request("pSearch")
nSearch=request("nSearch")
keyword=trim(checkStr(request("keyword")))
stable=checkstr(request("stable"))
scoreuser=trim(checkStr(request("scoreuser")))
if stable="" then stable=Nowusebbs
if request("page")="" or not isInteger(request("page"))  then
	currentPage=1
else
	currentPage=cint(request("page"))
end if
if stype<3 then
	if keyword="" and scoreuser="" and request("SearchScore")="ALL" then
		Errmsg=Errmsg+"<br>"+"<li>���������ѯ�ؼ��֡�"
		founderr=true
	end if
	'����������������
	if request("SearchDate")="ALL" then
		searchday=" "
	else
		if not isInteger(request("SearchDate"))  then
			Errmsg=Errmsg+"<br>"+"<li>���������������������"
			founderr=true
		else
			searchday=" datediff('d',b.dateandtime,Now()) < "&request("SearchDate")&" and "
		end if
	end if
	if scoreuser="" then searchscoreuser=" "
	select case request("SearchScore")
	case "ALL"
		SearchScore=" "
	case "NONE"
		SearchScore=" isnull(Score) and "
	case "A5"
		SearchScore=" not isnull(Score) and Score>=5 and "
	case "A4"
		SearchScore=" not isnull(Score) and Score>=4 and "
	case "A3"
		SearchScore=" not isnull(Score) and Score>=3 and "
	case "A2"
		SearchScore=" not isnull(Score) and Score>=2 and "
	case "A1"
		SearchScore=" not isnull(Score) and Score>=1 and "
	case "Z0"
		SearchScore=" not isnull(Score) and Score=0 and "
	case "B1"
		SearchScore=" not isnull(Score) and Score<=-1 and "
	case "B2"
		SearchScore=" not isnull(Score) and Score<=-2 and "
	case "B3"
		SearchScore=" not isnull(Score) and Score<=-3 and "
	case "B4"
		SearchScore=" not isnull(Score) and Score<=-4 and "
	case "B5"
		SearchScore=" not isnull(Score) and Score<=-5 and "
	end select
end if
call nav()
if cint(boardid)<1 then
	call head_var(0,0,"��̳����","query.asp?boardid=0")
	searchboard=""
else
	call head_var(0,0,boardtype,"query.asp?boardid="&boardid)
	searchboard=" b.boardid="&boardid&" and "
end if

if Cint(GroupSetting(14))=0 then
	Errmsg=Errmsg+"<br>"+"<li>��û���ڱ���̳������Ȩ�ޣ���<a href=login.asp>��¼</a>����ͬ����Ա��ϵ��"
	founderr=true
end if
if founderr then
	call dvbbs_error()
else
	call search()
	if founderr then call dvbbs_error()
end if
call footer()

sub search()
	sql=""
	set rs=server.createobject("adodb.recordset")
	select case stype
	case 1
		select case nSearch
		'����������������
		case 1
			if keyword="" then
				sql="select top 1000 b.locktopic,b.boardid,b.rootid,b.announceid,b.body,b.Expression,b.topic,b.username,b.dateandtime,b.postuserid,b.ParentID,b.isbest,f.boardtype from "&stable&" b inner join board f on b.boardid=f.boardid where "&searchboard&" "&searchday&" "&searchscoreuser&" "&searchscore&" b.parentid=0 and b.locktopic<2 ORDER BY b.announceID desc"
			else
				sql="select top 1000 b.locktopic,b.boardid,b.rootid,b.announceid,b.body,b.Expression,b.topic,b.username,b.dateandtime,b.postuserid,b.ParentID,b.isbest,f.boardtype from "&stable&" b inner join board f on b.boardid=f.boardid where b.username='"&keyword&"' and "&searchboard&" "&searchday&" "&searchscoreuser&" "&searchscore&" b.parentid=0 and b.locktopic<2 ORDER BY b.announceID desc"
			end if
			ordername="����������������"
		'�����ظ���������
		case 2
			if keyword="" then
				sql="select top 1000 b.locktopic,b.boardid,b.rootid,b.announceid,b.body,b.Expression,b.topic,b.username,b.dateandtime,b.postuserid,b.ParentID,b.isbest,f.boardtype from "&stable&" b inner join board f on b.boardid=f.boardid where "&searchboard&" "&searchday&" "&searchscoreuser&" "&searchscore&" b.parentid>0 and b.locktopic<2 ORDER BY b.announceID desc"
			else
				sql="select top 1000 b.locktopic,b.boardid,b.rootid,b.announceid,b.body,b.Expression,b.topic,b.username,b.dateandtime,b.postuserid,b.ParentID,b.isbest,f.boardtype from "&stable&" b inner join board f on b.boardid=f.boardid where b.username='"&keyword&"' and "&searchboard&" "&searchday&" "&searchscoreuser&" "&searchscore&" b.parentid>0 and b.locktopic<2 ORDER BY b.announceID desc"
			end if
			ordername="�����ظ���������"
		'���߶�����
		case 3
			if keyword="" then
				sql="select top 1000 b.locktopic,b.boardid,b.rootid,b.announceid,b.body,b.Expression,b.topic,b.username,b.dateandtime,b.postuserid,b.ParentID,b.isbest,f.boardtype from "&stable&" b inner join board f on b.boardid=f.boardid where "&searchboard&" "&searchday&" "&searchscoreuser&" "&searchscore&" b.locktopic<2 ORDER BY b.announceID desc"
			else
				sql="select top 1000 b.locktopic,b.boardid,b.rootid,b.announceid,b.body,b.Expression,b.topic,b.username,b.dateandtime,b.postuserid,b.ParentID,b.isbest,f.boardtype from "&stable&" b inner join board f on b.boardid=f.boardid where "&searchboard&" "&searchday&" "&searchscoreuser&" "&searchscore&" b.username='"&keyword&"' and b.locktopic<2 ORDER BY b.announceID desc"
			end if
			ordername="��������ͻظ���������"
		end select
	case 2
		select case pSearch
		'��������ؼ���
		case 1
			if keyword="" then
				sql="select top 1000 b.locktopic,b.boardid,b.rootid,b.announceid,b.body,b.Expression,b.topic,b.username,b.dateandtime,b.postuserid,b.ParentID,b.isbest,f.boardtype from "&stable&" b inner join board f on b.boardid=f.boardid where "&searchboard&" "&searchday&" "&searchscoreuser&" "&searchscore&" b.locktopic<2 ORDER BY b.announceID desc"
			else
				sql="select top 1000 b.locktopic,b.boardid,b.rootid,b.announceid,b.body,b.Expression,b.topic,b.username,b.dateandtime,b.postuserid,b.ParentID,b.isbest,f.boardtype from "&stable&" b inner join board f on b.boardid=f.boardid where "&searchboard&" "&searchday&" "&searchscoreuser&" "&searchscore&" (" & translate(keyword,"b.topic") & ") and b.locktopic<2 ORDER BY b.announceID desc"
			end if
			'�������ݹؼ���
			ordername="��������"
		case 2
			if keyword="" then
				sql="select top 1000  b.locktopic,b.boardid,b.rootid,b.announceid,b.body,b.Expression,b.topic,b.username,b.dateandtime,b.postuserid,b.ParentID,b.isbest,f.boardtype from "&stable&" b inner join board f on b.boardid=f.boardid where  "&searchboard&" "&searchday&" "&searchscoreuser&" "&searchscore&" b.locktopic<2 ORDER BY b.announceID desc"
			else
				sql="select top 1000  b.locktopic,b.boardid,b.rootid,b.announceid,b.body,b.Expression,b.topic,b.username,b.dateandtime,b.postuserid,b.ParentID,b.isbest,f.boardtype from "&stable&" b inner join board f on b.boardid=f.boardid where  "&searchboard&" "&searchday&" "&searchscoreuser&" "&searchscore&" (" & translate(keyword,"b.body") & ") and b.locktopic<2 ORDER BY b.announceID desc"
			end if
			'���߶�����
			ordername="��������"
		case 3
			if keyword="" then
				sql="select top 1000  b.locktopic,b.boardid,b.rootid,b.announceid,b.body,b.Expression,b.topic,b.username,b.dateandtime,b.postuserid,b.ParentID,b.isbest,f.boardtype from "&stable&" b inner join board f on b.boardid=f.boardid where  "&searchboard&" "&searchday&" "&searchscoreuser&" "&searchscore&" b.locktopic<2 ORDER BY b.announceID desc"
			else
				sql="select top 1000  b.locktopic,b.boardid,b.rootid,b.announceid,b.body,b.Expression,b.topic,b.username,b.dateandtime,b.postuserid,b.ParentID,b.isbest,f.boardtype from "&stable&" b inner join board f on b.boardid=f.boardid where  "&searchboard&" "&searchday&" "&searchscoreuser&" "&searchscore&" (" & translate(keyword,"b.topic") & " or " & translate(keyword,"b.body") & ") and b.locktopic<2 ORDER BY b.announceID desc"
			end if
			ordername="�������������"
		end select
	case 3
		if searchboard="" then
			sql="select top 100 b.locktopic,b.boardid,b.rootid,b.announceid,b.body,b.Expression,b.topic,b.username,b.dateandtime,b.postuserid,b.ParentID,b.isbest,f.boardtype from "&stable&" b inner join board f on b.boardid=f.boardid where  "&searchboard&" b.locktopic<2 ORDER BY b.announceID desc"
			ordername="����100��"
		else
			sql="select top 30 b.locktopic,b.boardid,b.rootid,b.announceid,b.body,b.Expression,b.topic,b.username,b.dateandtime,b.postuserid,b.ParentID,b.isbest,f.boardtype from "&stable&" b inner join board f on b.boardid=f.boardid where  "&searchboard&" b.locktopic<2 ORDER BY b.announceID desc"
			ordername="����30��"
		end if
	end select
	
	if sql="" then
		Errmsg=Errmsg+"<br>"+"<li>��ָ����ѯ������"
		founderr=true
		exit sub
	end if
	rs.open sql,conn,1,1

	if rs.eof and rs.bof then
		Errmsg=Errmsg+"<br>"+"<li>û���ҵ���Ҫ��ѯ�����ݡ�"
		founderr=true
		exit sub
	else
		rs.PageSize = Forum_Setting(11)
		rs.AbsolutePage=currentpage
		page_count=0
		totalrec=rs.recordcount
		call searchinfo()
		call listPages3()
	end if

	rs.close
	set rs=nothing
	call activeonline()
end sub

sub searchinfo()
	%><table cellpadding=0 cellspacing=0 border=0 width="<%=Forum_body(12)%>" align=center>
		<tr>
			<td>��ѯ<%if request("SearchDate")="ALL" then%>��������<%else%><%=request("SearchDate")%>����<%end if%>�����ӣ�<%=ordername%><%if totalrec>=1000 then%>���ֻ�ܲ�ѯ��<font color=<%=Forum_body(8)%>>1000</font>����¼������ľͲ���ʾ��<%else%>����ѯ��<font color=<%=Forum_body(8)%>><%=totalrec%></font>�����<%end if%></td>
		</tr>
	</table>
	<TABLE cellPadding=3 cellSpacing=1 class=tableborder1 align=center>
		<TR valign=middle>
			<Th height=25 width=32>״̬</Th>
			<Th width=*>�� ��</Th>
			<Th width=80>��̳</Th>
			<Th width=80>�� ��</Th>
			<Th width=70>������</Th>
			<Th width=80>�ظ���</Th>
		</TR><%
		while (not rs.eof) and (not page_count = rs.PageSize)
			%><TR> 
				<TD align=center class=tablebody2 width=32><%
					dim RsTopic,lock
					lock=0
					if rs(10)=0 then
						set RsTopic=conn.execute("select * from topic where topicid="&rs(2))
						response.write TopicIcon(rs("BoardID"),rs("rootid"),rs("DateAndTime"),rstopic("istop"),rstopic("isVote"),rstopic("LockTopic"),iif(RsTopic("child")>=Cint(forum_setting(44)),1,0))
						Set rsTopic=nothing
					else
						response.write TopicIcon(rs("BoardID"),rs("rootid"),rs("DateAndTime"),0,0,0,0)
					end if
				%></TD>
				<TD onmouseover=this.className='TableBody2' onmouseout=this.className='TableBody1'  class=tablebody1 width=*>&nbsp;<a href='dispbbs.asp?boardID=<%=rs(1)%>&replyID=<%=rs(3)%>&ID=<%=rs(2)%>&skin=1' target=_blank><img src='<%=Forum_info(8)&rs(5)%>' border=0 alt="���´������������"></a>&nbsp;<a href='dispbbs.asp?boardID=<%=rs(1)%>&replyID=<%=rs(3)%>&ID=<%=rs(2)%>&skin=1'><%
					if renzhen(rs(1),membername)=false then
						%><font color=gray>����֤��������̳���ӣ�ֻ����֤�û���������ܲ鿴��</font><%
					elseif VipBoard(rs(1),membername)=false then
						%><font color=gray>��VIP��������̳���ӣ�ֻ��VIP�������Ա���ܲ鿴��</font><%
					else
						if rs(6)="" then
							if rs("isbest")=1 and Cint(GroupSetting(41))=0 then
								%><font color=gray>������������ֹԤ����</font><%
							elseif rs("isbest")=2 then
								%><font color=gray>���������Σ���ֹԤ����</font><%
							else
								%><%=left(htmlencode(replace(reUBBCode(rs(4),false),chr(10)," ")),26)%>...<%
							end if
						else
							%><%=htmlencode(reUBBCode(rs(6),false))%><%
						end if
					end if
				%></a></TD> 
    		<TD align=center class=tablebody2  width=80><a href="list.asp?boardid=<%=htmlencode(rs("boardid"))%>"><%=htmlencode(rs("boardtype"))%></a></TD>
    		<TD align=center class=tablebody1  width=80><a href="dispuser.asp?id=<%=htmlencode(rs("postuserid"))%>"><%=htmlencode(rs(7))%></a></TD>
    		<TD align=center class=tablebody2 width=70><%=month(rs("dateandtime"))&"-"&day(rs("dateandtime"))%>&nbsp;<%=FormatDateTime(rs("dateandtime"),4)%></TD>
    		<TD align=center class=tablebody1 width=80>------</TD>
			</TR><%
	  	page_count = page_count + 1
			rs.movenext
		wend
	%></table><%
end sub

	sub listPages3()

	Pcount=rs.PageCount
	response.write "<table border=0 cellpadding=0 cellspacing=3 width="""&Forum_body(12)&""" align=center>"&_
			"<tr><td valign=middle nowrap>"&_
			"<font color="&Forum_body(13)&">ҳ�Σ�<b>"&currentpage&"</b>/<b>"&Pcount&"</b>ҳ"&_
			"ÿҳ<b>"&Forum_Setting(11)&"</b> ������<b>"&totalrec&"</b></font></td>"&_
			"<td valign=middle nowrap><font color="&Forum_body(13)&"><div align=right><p>��ҳ�� "
	call DispPageNum(currentpage,PCount,"""?page=","&stype="&stype&"&pSearch="&pSearch&"&nSearch="&nSearch&"&keyword="&keyword&"&SearchDate="&request("SearchDate")&"&boardid="&boardid&"&stable="&stable&"&scoreuser="&scoreuser&"&searchscore="&request("SearchScore")&"""")
	response.write "</p></div></font></td></tr></table>"

	end sub

public function translate(sourceStr,fieldStr)
rem �����߼����ʽ��ת������
  dim  sourceList
  dim resultStr
  dim i,j
  if instr(sourceStr," ")>0 then 
     dim isOperator
     isOperator = true
     sourceList=split(sourceStr)
     '--------------------------------------------------------
     rem Response.Write "num:" & cstr(ubound(sourceList)) & "<br>"
     for i = 0 to ubound(sourceList)
        rem Response.Write i 
    Select Case ucase(sourceList(i))
    Case "AND","&","��","��"
        resultStr=resultStr & " and "
        isOperator = true
    Case "OR","|","��"
        resultStr=resultStr & " or "
        isOperator = true
    Case "NOT","!","��","��","��"
        resultStr=resultStr & " not "
        isOperator = true
    Case "(","��","��"
        resultStr=resultStr & " ( "
        isOperator = true
    Case ")","��","��"
        resultStr=resultStr & " ) "
        isOperator = true
    Case Else
        if sourceList(i)<>"" then
            if not isOperator then resultStr=resultStr & " and "
            if inStr(sourceList(i),"%") > 0 then
                resultStr=resultStr&" "&fieldStr& " like '" & replace(sourceList(i),"'","''") & "' "
            else
                resultStr=resultStr&" "&fieldStr& " like '%" & replace(sourceList(i),"'","''") & "%' "
            end if
                isOperator=false
        End if    
    End Select
        rem Response.write resultStr+"<br>"
     next 
     translate=resultStr
  else '������
     if inStr(sourcestr,"%") > 0 then
         translate=" " & fieldStr & " like '" & replace(sourceStr,"'","''") &"' "
     else
    translate=" " & fieldStr & " like '%" & replace(sourceStr,"'","''") &"%' "
     End if
     rem ǰ�����һ���ո������sqlʱ���˼ӣ�������
  end if  
end function
%>