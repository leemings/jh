<!--#include file="conn.asp"-->
<!-- #include file="inc/const.asp" -->
<%
	dim totalrec
	dim n
	dim orders,ordername
	dim currentpage,page_count,Pcount
	currentPage=request("page")
	if not founduser then
		Errmsg="����û�е�¼����<a href=login.asp>��¼</a>�����"
		founderr=true
	end if
	orders=request("s")
	if orders="" or not isInteger(orders) then
		orders=1
	else
		orders=clng(orders)
	end if
	if orders=1 then
		ordername="�Ҳ��������"
		sql="select top 200 * from topic where topicid in (select top 200 rootid from "&NowUseBBS&" where postuserid="&userid&" and rootid<>0 order by announceid desc) and locktopic<>2 order by topicid desc"
	elseif orders=2 then
		sql="select top 200 * from topic where postuserid="&userid&" and child>0 and locktopic<>2 ORDER BY topicid desc"
		ordername="�ѱ��ظ����ҵķ���"
	else
		ordername="�ҷ��������"
		sql="select top 200 * from topic where postuserid="&userid&" and locktopic<>2 order by topicid desc"
	end if	
	if currentpage="" or not isInteger(currentpage) then
		currentpage=1
	else
		currentpage=clng(currentpage)
	end if
	stats=ordername
if founderr then
	call nav()
	call head_var(2,0,"","")
else
	call nav()
	call head_var(2,0,"","")
	call AnnounceList1()
	call listPages3()
end if
call activeonline()
call footer()

sub AnnounceList1()
	set rs=server.createobject("adodb.recordset")
	rs.open sql,conn,1,1
	if rs.bof and rs.eof then
		call showEmptyBoard1()
	else
		rs.PageSize = Forum_Setting(11)
		rs.AbsolutePage=currentpage
		page_count=0
		totalrec=rs.recordcount
		call showpagelist1()
	end if
	end sub

	REM ��ʾ�����б�	
	sub showPageList1()
	i=0
	response.write "<p align=center>��ǰ��̳���ݺܿ��ܽ����˷ֱ�������<BR>��ǰ����ʾ���ڵ�ǰ���ݱ�����ǰ200�����������Ҫ��ѯ���������뵽����ҳ��</p>"%>
	<script>
	function loadThreadFollow(t_id,b_id){
		var targetImg =eval("document.all.followImg" + t_id);
		var targetDiv =eval("document.all.follow" + t_id);
	
		if ("object"==typeof(targetImg)){
			if (targetDiv.style.display!='block'){
				targetDiv.style.display="block";
				targetImg.src="<%=Forum_info(7)&Forum_boardpic(16)%>";
				if (targetImg.loaded=="no"){
					document.frames["hiddenframe"].location.replace("loadtree1.asp?boardid="+b_id+"&rootid="+t_id);
				}
			}else{
				targetDiv.style.display="none";
				targetImg.src="<%=Forum_info(7)&Forum_boardpic(15)%>";
			}
		}
	}
	</script>
	<iframe width=0 height=0 src="" id="hiddenframe"></iframe>
	<%response.write "<TABLE cellPadding=1 cellSpacing=1 class=tableborder1 align=center>"
	response.write "<TR align=middle>"
	response.write "<Th height=27 width=32 id=tabletitlelink>״̬</th>"
	response.write "<Th width=*>�� ��  (��<img src="&Forum_info(7)&"plus.gif>����չ�������б�)</Th>"
	response.write "<Th width=80>�� ��</Th>"
	response.write "<Th width=64>�ظ�/����</Th>"
	response.write "<Th width=70>������</Th>"
	response.write "<Th width=80>�ظ���</Th>"
	response.write "</TR>"
	dim Ers, Eusername, Edateandtime,body
	dim vrs,votenum,pnum,p,iu,votenum_1
	while (not rs.eof) and (not page_count = rs.PageSize)
		boardid=rs("boardid")
		response.write "<TR align=middle>"
		response.write "<TD class=tablebody2 width=32 height=27>"
		response.write TopicIcon(rs("BoardID"),rs("topicid"),rs("lastposttime"),rs("istop"),rs("isVote"),rs("LockTopic"),iif(Rs("child")>=Cint(forum_setting(44)),1,0))
		response.write "</TD>"
		response.write "<TD onmouseover=this.className='TableBody2' onmouseout=this.className='TableBody1' align=left  class=tablebody1 width=*>"
		if Rs("child")=0 then
			response.write "<img src="""& Forum_info(7)&Forum_boardpic(16)&""" id=""followImg"& Rs("topicid") &""">"
		else
			response.write "<img loaded=""no"" src="""& Forum_info(7)&Forum_boardpic(15)&""" id=""followImg"& Rs("topicid") &""" style=""cursor:hand;"" onclick=""loadThreadFollow("& Rs("topicid") &","& Rs("boardid") &")"" title=չ�������б�>"
		end if
		'response.write "&nbsp;<img src='"& Forum_info(8) & rs("Expression") &"'>&nbsp;"
		response.write "<a href=""dispbbs.asp?boardID="& rs("boardid") &"&ID="& rs("topicid") &""">"
		if len(rs("title"))>26 then
			response.write ""&left(htmlencode(rs("title")),26)&"..."
		else
			response.write htmlencode(rs("title"))
		end if
		response.write "</a>"
		if (rs("child")+1)>cint(Forum_Setting(11)) then
			response.write "&nbsp;&nbsp;[<img src="&Forum_info(7)&Forum_statePic(6)&"><b>"
			Pnum=(Cint(rs("child")+1)/cint(Forum_Setting(11)))+1
			for p=1 to Pnum
			response.write " <a href='dispbbs.asp?boardID="& boardID &"&ID="& rs("topicid") &"&star="&P&"'><FONT color="&Forum_body(8)&"><b>"&p&"</b></font></a> "
			next
			response.write "</b>]"
		end if
		if rs("isbest")=1 then
			response.write "&nbsp;&nbsp;<img src="""&Forum_info(7)&Forum_statePic(4)&""">"
		end if
		response.write "</TD>"
		response.write "<TD class=tablebody2 width=80><a href=dispuser.asp?name="& rs("postusername") &">"& rs("postusername") &"</a></TD>"
		response.write "<TD class=tablebody1 width=64>"
		response.write ""& rs("child") &"/"& rs("hits") &"</TD>"
		response.write "<TD align=center class=tablebody2 width=70>"&month(rs("dateandtime"))&"-"&day(rs("dateandtime"))&"&nbsp;"&FormatDateTime(rs("dateandtime"),4)&"</TD>"
		response.write "<TD align=center class=tablebody1 width=80>"
		if rs("child")> 0 then
			response.write split(rs("lastpost"),"$")(0)
		else
			response.write "------"
		end if
		response.write "</TD></TR>"
		response.write "<tr style=""display:none"" id=""follow"& Rs("topicid") &"""><td colspan=6 id=""followTd"& Rs("topicid") &""" style=""padding:0px""><div style=""width:240px;margin-left:18px;border:1px solid black;background-color:lightyellow;color:black;padding:2px"" onclick=""loadThreadFollow("& Rs("topicid") &","&Rs("boardid")&")"">���ڶ�ȡ���ڱ�����ĸ��������Ժ��</div></td></tr>"
	  page_count = page_count + 1
	rs.movenext
	wend
	response.write "</table>"
	end sub

	sub listPages3()
	dim endpage
	Pcount=rs.PageCount
	response.write "<table border=0 cellpadding=0 cellspacing=3 width="""&Forum_body(12)&""" align=center>"&_
			"<tr><td valign=middle nowrap>"&_
			"<font color="&Forum_body(13)&">ҳ�Σ�<b>"&currentpage&"</b>/<b>"&Pcount&"</b>ҳ"&_
			"ÿҳ<strong>"&Forum_Setting(11)&"</b> ����<b>"&totalrec&"</b></td>"&_
			"<td valign=middle nowrap><div align=right><p>��ҳ�� "
	call DispPageNum(currentpage,PCount,"""?page=","&s="&orders&"""")
	response.write "</p></div></font></td></tr></table>"
	rs.close
	set rs=nothing	
	end sub 

	sub showEmptyBoard1()
%>
<TABLE cellPadding=4 cellSpacing=1 class=tableborder1 align=center>
<TBODY>
<TR align=middle>
<Th height=25>״̬</Th>
<Th>�� ��  (�������Ϊ���´����)</Th>
<Th>�� ��</Th> 
<Th>�ظ�/����</Th> 
<Th>���»ظ�</Th>
</TR> 
<tr><td colSpan=5 width=100% class=tablebody1>�������ݣ���ӭ��������</td></tr>
</TBODY></TABLE>
<%
	end sub
%>