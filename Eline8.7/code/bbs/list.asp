<!--#include file="conn.asp"-->
<!-- #include file="inc/const.asp" -->
<!-- #include file="inc/char_board.asp" -->
<%
response.buffer=true
'=========================================================
' File: list.asp
' Version:5.0
' Date: 2002-9-20
' Script Written by satan
'=========================================================
' Copyright (C) 2001,2002 AspSky.Net. All rights reserved.
' Web: http://www.aspsky.net,http://www.dvbbs.net
' Email: info@aspsky.net,eway@aspsky.net
'=========================================================

Dim currentPage
Dim Fast_action
Dim reBoard_Setting,BoardCount,k,ColSpanNum
Dim Board_Count
Dim TdTableWidth
Dim master_3,master_4
dim onlinenum,guestnum
Dim LastPost,Lastuser,LastID
dim LastTime,LastUserid,LastRootid,body
dim Pnum,p,replynum
dim totalrec,ii,page_count
dim n,pi
dim call_info
dim RsTopic
dim cmd,AllTopNum,RsTop
dim yorder,ydateandtime,yordertype
AllTopNum=0
Fast_action=false

if boardmaster or master or superboardmaster then
	Fast_action=true
else
	Fast_action=false
end if

if BoardID="" or (not isInteger(BoardID)) or BoardID="0" then
	Errmsg=Errmsg+"<br>"+"<li>����İ����������ȷ�����Ǵ���Ч�����ӽ��롣"
	founderr=true
else
	BoardID=clng(BoardID)
end if
if cint(Board_Setting(2))=1 then
	if not founduser then
		Errmsg=Errmsg+"<br>"+"<li>����̳Ϊ��֤��̳����<a href=login.asp>��¼</a>��ȷ�������û����Ѿ��õ�����Ա����֤����롣"
		founderr=true
	else
		if chkboardlogin(boardid,membername)=false then
		Errmsg=Errmsg+"<br>"+"<li>����̳Ϊ��֤��̳����ȷ�������û����Ѿ��õ�����Ա����֤����롣"
		founderr=true
		end if
	end if
end if
if cint(Board_Setting(1))=1 then
	if Cint(GroupSetting(37))=0 then
		Errmsg=ErrMsg+"<Br>"+"<li>��û��Ȩ�޽���������̳��"
		founderr=true
	end if
end if

'###################��������޸Ŀ�ʼ(asilas����)##################
if cint(Board_Setting(51))=1 then
	if not founduser then
		Errmsg=Errmsg+"<br>"+"<li>����̳Ϊ������̳����<a href=login.asp>��¼</a>��ȷ�������û����Ѿ��õ�����Ա����֤����롣"
		founderr=true
	else
		if chkviplogin(membername)=false then
		Errmsg=Errmsg+"<br>"+"<li>����̳����Ϊ<font color=red>VIP��Աר�ð���</font>����ȷ�����������Ƿ���ϡ�"
		founderr=true
		end if
	end if
end if
if cint(Board_Setting(52))<>0 then
	if not founduser then
		Errmsg=Errmsg+"<br>"+"<li>����̳Ϊ������̳����<a href=login.asp>��¼</a>��ȷ�������û����Ѿ��õ�����Ա����֤����롣"
		founderr=true
	else
		dim sexshow
		if cint(Board_Setting(52))=1 then
		sexshow="Ů��"
		elseif cint(Board_Setting(52))=2 then
		sexshow="����"
		end if
		if chksexlogin(cint(Board_Setting(52)),membername)=false then
		Errmsg=Errmsg+"<br>"+"<li>����̳����Ϊ<font color=red>"&sexshow&"��̳����</font>����ȷ�������Ա��Ƿ���ϡ�"
		founderr=true
		end if
	end if
end if
'####################��������޸Ľ���(asilas����)#################

if founderr then
	call nav()
	call head_var(2,0,"","")
	call dvbbs_error()
else
	call nav()
	if Cint(Board_Setting(43))=0 then
	stats="�����б�"
	else
	stats="��̳�б�"
	end if
	call head_var(1,BoardDepth,0,0)
	call main()
	if founderr then call dvbbs_error()
end if
call footer()

sub main()
currentPage=request("page")
if currentpage="" or not isInteger(currentpage) then
	currentpage=1
else
	currentpage=clng(currentpage)
end if
stats=boardtype & "�����б�"
call activeonline()
onlinenum=online(boardid,false)
guestnum=guest(boardid,false)
master_2="<select name=masterlist onchange=""if(this.selectedIndex>0) {window.open('dispuser.asp?name='+this.options[this.selectedIndex].value,'_blank');this.selectedIndex=0;}"" title=����鿴��������><option value="""">�������</option>"
if trim(boardmasterlist)<>"" then
	master_1=split(boardmasterlist, "|")
	for i = 0 to ubound(master_1)
	master_2=master_2&"<option value="""&master_1(i)&""">"&master_1(i)&"</option>"
	next
else
	master_2="�ް���"
end if
master_2=master_2&"</select>"
%>
<TABLE cellSpacing=0 cellPadding=0 width=<%=Forum_body(12)%> border=0 align=center>
<tr>
	<td align=center width=100% valign=middle><!--#include file="z_announce.asp"--></td>
</tr>
</table>
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
<%
k=0
Dim cBoardID
dim simdisp
simdisp=false
Response.Cookies("BoardList").expires= date+7
cBoardID= Request("cBoardid")
if isnumeric(cBoardID) then
	if request("CatLog")="N" then
	Response.Cookies("BoardList")(cBoardID & "BoardID")= "NotShow"
	else
	Response.Cookies("BoardList")(cBoardID & "BoardID")= "Show"
	end if
end if
if request.cookies("BoardList")(boardid & "BoardID")<>"NNotShow" then
if BoardParentID=0 then
sql="select * from board where (depth<="&Boarddepth&"+"&forum_setting(5)&" and (ParentStr='"&BoardID&"' or ParentStr like '"&BoardID&",%')) or boardid="&boardid&" order by orders"
else
sql="select * from board where (depth<="&Boarddepth&"+"&forum_setting(5)&" and (ParentStr='"&BoardParentStr&","&boardid&"' or ParentStr like '"&BoardParentStr&","&boardid&",%')) or boardid="&boardid&" order by orders"
end if
'response.write sql
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,1
if not (rs.eof and rs.bof) then BoardCount=rs.recordcount
if BoardCount>1 then
do while not rs.eof
k=k+1
reBoard_Setting=split(rs("Board_setting"),",")
if rs("BoardMaster")<>"" and not isnull(rs("BoardMaster")) then
	master_3=split(rs("BoardMaster"), "|")
	for i = 0 to ubound(master_3)
	if i>6 then
		master_4=master_4
	else
		master_4=master_4 & "<a href=dispuser.asp?name="&master_3(i)&" target=_blank>"&master_3(i)&"</a>&nbsp;"
	end if
	next
	if i>7 then master_4=master_4 & "<font color=gray>More...</font>"
else
	master_4="����"
end if
if rs("ParentID")=0 or rs("boardid")=boardid then
	if request.cookies("BoardList")(rs("boardid") & "BoardID")="NotShow" then
		ColSpanNum=reBoard_Setting(41)
		simdisp=true
	elseif request.cookies("BoardList")(rs("boardid") & "BoardID")="Show" then
		simdisp=false
		ColSpanNum=2
	else
		if Cint(reBoard_Setting(39))=1 then
			ColSpanNum=reBoard_Setting(41)
			simdisp=true
		else
			simdisp=false
			ColSpanNum=2
		end if
	end if
	response.write "<table cellspacing=1 cellpadding=0 align=center class=tableBorder1>"
	response.write "<TR><Th colSpan="&ColSpanNum&" height=25 align=left id=TableTitleLink>&nbsp;"
	if simdisp then
		response.write "<a href=""?BoardID="&rs("boardid")&"&cBoardid="&rs("boardid")&"&Catlog=Y"" title=""չ����̳�б�""><img src="""&Forum_info(7)&Forum_boardpic(15)&""" border=0></a><a href=list.asp?boardid="&rs("boardid")&" title=���뱾������̳>"&rs("boardtype")&"</a>&nbsp;<a href=""?BoardID="&rs("boardid")&"&Catlog=NN"">[��ȫ�ر�]</a>"
	else
		response.write "<a href=""?BoardID="&rs("boardid")&"&cBoardid="&rs("boardid")&"&Catlog=N"" title=""�ر���̳�б�""><img src="""&Forum_info(7)&Forum_boardpic(16)&""" border=0></a><a href=list.asp?boardid="&rs("boardid")&" title=���뱾������̳>"&rs("boardtype")&"</a>&nbsp;<a href=""?BoardID="&rs("boardid")&"&Catlog=NN"">[��ȫ�ر�]</a>"
	end if
	response.write "</Th></TR>"
	Board_Count=0
else
	if simdisp then
		call main_1_boardlist_2()
	else
		call main_1_boardlist_1()
	end if
end if
master_4=""
rs.movenext
if k<BoardCount then
	if rs("parentid")=0 and Board_count<ColSpanNum and simdisp then
		for i=Board_Count+1 to ColSpanNum
			response.write "<td class=tablebody1 width="""&TdTableWidth&"%"">&nbsp;</td>"
		next
		response.write "</tr>"
	end if
	if rs("parentid")=0 then
		response.write "</table><BR>"
		simdisp=false
		ColSpanNum=2
	end if
else
	if Board_count<ColSpanNum and simdisp then
		for i=Board_Count+1 to ColSpanNum
			response.write "<td class=tablebody1 width="""&TdTableWidth&"%"">&nbsp;</td>"
		next
		response.write "</tr>"
	end if
end if
loop
response.write "</table>"
end if
rs.close
set rs=nothing
response.write "<BR>"
end if
if request("listtime")="" then
	if cint(Board_Setting(37))=2 then
		ydateandtime="and datediff('d',lastposttime,now)<=5"
	elseif cint(Board_Setting(37))=3 then
		ydateandtime="and datediff('d',lastposttime,now)<=15"
	elseif cint(Board_Setting(37))=4 then
		ydateandtime="and datediff('d',lastposttime,now)<=30"
	elseif cint(Board_Setting(37))=5 then
		ydateandtime="and datediff('d',lastposttime,now)<=60"
	elseif cint(Board_Setting(37))=6 then
		ydateandtime="and datediff('d',lastposttime,now)<=120"
	elseif cint(Board_Setting(37))=7 then
		ydateandtime="and datediff('d',lastposttime,now)<=180"
	else
		ydateandtime=""
	end if
else
	select case request("listtime")
	case 2
		ydateandtime="and datediff('d',lastposttime,now)<=5"
	case 3
		ydateandtime="and datediff('d',lastposttime,now)<=15"
	case 4
		ydateandtime="and datediff('d',lastposttime,now)<=30"
	case 5
		ydateandtime="and datediff('d',lastposttime,now)<=60"
	case 6
		ydateandtime="and datediff('d',lastposttime,now)<=120"
	case 7
		ydateandtime="and datediff('d',lastposttime,now)<=180"
	case else
		ydateandtime=""
	end select
end if
if request("orders")="" then
	select case cint(Board_Setting(38))
	case 2
		yorder="dateandtime"
	case 3
		yorder="title"
	case 4
		yorder="postusername"
	case 5
		yorder="child"
	case 6
		yorder="hits"
	case else
		yorder="lastposttime"
	end select
else
	select case cint(request("orders"))
	case 2
		yorder="dateandtime"
	case 3
		yorder="title"
	case 4
		yorder="postusername"
	case 5
		yorder="child"
	case 6
		yorder="hits"
	case 7
		yorder="lastpost"
	case else
		yorder="lastposttime"
	end select
end if	
if request("ordertype")="1" then
	yordertype=" asc"
else
	yordertype=" desc"
end if
if currentpage=1 then
myCache.name="AllTopNum"
if myCache.valid then
	if myCache.value>0 then
	set RsTop=server.createobject("adodb.recordset")
	sql="select TopicID,boardid,title,postusername,postuserid,dateandtime,child,hits,votetotal,lastpost,lastposttime,istop,isvote,isbest,locktopic,Expression,ChenTopicColor from topic where istop=2 and locktopic<2"&ydateandtime&" ORDER BY "&yorder&yordertype
	RsTop.open sql,conn,1,1
	if RsTop.eof and RsTop.bof then
		AllTopNum=0
	else
		AllTopNum=RsTop.recordcount
	end if
	myCache.add AllTopNum,dateadd("n",60,now)
	end if 'end by value
else
	set RsTop=server.createobject("adodb.recordset")
	sql="select TopicID,boardid,title,postusername,postuserid,dateandtime,child,hits,votetotal,lastpost,lastposttime,istop,isvote,isbest,locktopic,Expression,ChenTopicColor from topic where istop=2 and locktopic<2"&ydateandtime&"  ORDER BY "&yorder&yordertype
	RsTop.open sql,conn,1,1
	if RsTop.eof and RsTop.bof then
		AllTopNum=0
	else
		AllTopNum=RsTop.recordcount
	end if
	myCache.add AllTopNum,dateadd("n",60,now)
end if 'end by valid
end if 'end by page

set RsTopic=server.createobject("adodb.recordset")
if IsAudit=1 then
if not founduser then userid=0
sql="select TopicID,boardid,title,postusername,postuserid,dateandtime,child,hits,votetotal,lastpost,lastposttime,istop,isvote,isbest,locktopic,Expression,ChenTopicColor from topic where boardID="&boardid&" and istop<2 and (locktopic<2 or (locktopic=3 and PostUserID="&userid&"))"&ydateandtime&" ORDER BY istop desc,"&yorder&yordertype
else
sql="select TopicID,boardid,title,postusername,postuserid,dateandtime,child,hits,votetotal,lastpost,lastposttime,istop,isvote,isbest,locktopic,Expression,ChenTopicColor from topic where boardid="&boardid&" and istop<2 and locktopic<2"&ydateandtime&" order by istop desc,"&yorder&yordertype
end if

RsTopic.open sql,conn,1,1
totalrec=RsTopic.RecordCount
if RsTopic.eof and RsTopic.bof and Cint(Board_Setting(43))=0 and AllTopNum=0 then
	call BoardEmpty()
elseif not (RsTopic.eof and RsTopic.bof) or (Cint(Board_Setting(43))=0 and AllTopNum>0) then
	call listinfo()
end if
end sub

sub main_1_boardlist_1()
if (Cint(reBoard_Setting(1))=1 and Cint(GroupSetting(37))=1) or Cint(reBoard_Setting(1))=0 then
response.write "<TR><TD align=middle width=100% class=tablebody1><table width=""100%"" cellspacing=0 cellpadding=0>"
LastPostInfo=split(rs("LastPost"),"$")
if not isdate(LastPostInfo(2)) then LastPostInfo(2)=Now()
'=========================BoardInfo============================
response.write "<TR><TD align=middle width=46 class=tablebody1>"
dim LastViewBoard
if request.cookies("LastView")("Board_"&rs(0))="" then
	if isnull(LastViewTime) then
		LastViewBoard=CDate("2003-1-1")
	else
		LastViewBoard=LastViewTime
	end if
else
	if isnull(LastViewTime) then
		LastViewBoard=CDate(request.cookies("LastView")("Board_"&rs(0)))
	elseif LastViewTime<CDate(request.cookies("LastView")("Board_"&rs(0))) then
		LastViewBoard=CDate(request.cookies("LastView")("Board_"&rs(0)))
	else
		LastViewBoard=LastViewTime
	end if
end if
if reBoard_Setting(0)=1 then
	if LastViewBoard<CDate(LastPostInfo(2)) then
	response.write "<img src="&Forum_info(7)&Forum_pic(15)&" width=32 height=32 alt=��������>"
	else
	response.write "<img src="&Forum_info(7)&Forum_pic(14)&" width=32 height=32 alt=��������>"
	end if
else
	if LastViewBoard<CDate(LastPostInfo(2)) then
	response.write "<img src="&Forum_info(7)&Forum_pic(7)&" width=32 height=32 alt=��������>"
	else
	response.write "<img src="&Forum_info(7)&Forum_pic(6)&" width=32 height=32 alt=��������>"
	end if
end if

response.write "</TD>"&_
	"<TD width=1 bgcolor="&Forum_body(27)&">"&_
	"<TD vAlign=top width=* class=tablebody1>"&_
	"<TABLE cellSpacing=0 cellPadding=2 width=100% border=0>"&_
	"<tr><td class=tablebody1 width=*>"
response.write "<table width=100% border=0 cellSpacing=0 cellPadding=2>"
response.write "<tr>"
response.write "<td width=150>"
'============= ͬѧ¼ ��ʼ ==================
if not isnull(rs("txlpd")) and rs("txlpd")<>"" then
	if cint(right(rs("txlpd"),1))=1 then
		response.write "<a href='z_school_class.asp?boardid="&rs(0)&"'><font color="&Forum_body(14)&">"&rs("boardtype")&"</font></a>"
	else
		response.write "<a href='list.asp?boardid="&rs(0)&"'><font color="&Forum_body(14)&">"&rs("boardtype")&"</font></a>"
	end if
else
	response.write "<a href='list.asp?boardid="&rs(0)&"'><font color="&Forum_body(14)&">"&rs("boardtype")&"</font></a>"
end if
'============= ͬѧ¼ ���� ==================
if rs("child")>0 then
response.write "(<font title=""��"&rs("child")&"��������̳"">"&rs("child")&"</font>)"
end if
response.write "</td>"
response.write "<td width=* align=left>"
response.write "<a href=announce.asp?boardid="&rs("boardid")&"><img src=pic/fabiao.gif border=0 width=12 height=13 alt=""��"&rs("boardtype")&"����������""></a> "
response.write "<a href=vote.asp?boardid="&rs("boardid")&"><img src=pic/toupiao.gif border=0 width=12 height=13 alt=""��"&rs("boardtype")&"������ͶƱ""></a> "
response.write "<a href=SmallPaper.asp?boardid="&rs("boardid")&"><img src=pic/xiaozibao.gif border=0 width=12 height=13 alt=""��"&rs("boardtype")&"����С�ֱ�""></a> "
response.write "<a href=queryresult.asp?boardid="&rs("boardid")&"&stype=3><img src=pic/findposter.gif border=0 width=12 height=13 alt=""�鿴"&rs("boardtype")&"������""></a> "
response.write "<a href=hotlist.asp?boardid="&rs("boardid")&"&stype=2&searchdate=30><img src=pic/retie.gif border=0 width=12 height=13 alt=""�鿴"&rs("boardtype")&"�����Ż���""></a> "
response.write "<a href=elist.asp?boardid="&rs("boardid")&"><img src=pic/jinhua.gif border=0 width=12 height=13 alt=""�鿴"&rs("boardtype")&"�ľ�������""></a> "
response.write "<a href=bbseven.asp?boardid="&rs("boardid")&"><img src=pic/chazhao.gif border=0 width=12 height=13 alt=""�鿴"&rs("boardtype")&"���¼�""></a>&nbsp;<font color="&forum_body(27)&" title='�����Ա:"&online(rs("boardid"),true)&"���������:"&guest(rs("boardid"),true)&"'>(�������� <b>"&online(rs("boardid"),true)+guest(rs("boardid"),true)&"</b> ��)</font>"
response.write "</td>"
response.write "</tr></table>"
response.write "</td>"
response.write "<td width=40 rowspan=2 align=center class=tablebody1>"
if rs("indeximg")<>"" then
	'============= ͬѧ¼ ��ʼ ==================
	if not isnull(rs("txlpd")) and rs("txlpd")<>"" then
		if cint(right(rs("txlpd"),1))=1 then
			response.write "<table align=left><tr><td align=right><a href=z_school_class.asp?boardid="&rs("boardid")&"><img src="&rs("indeximg")&" align=top border=0"
			if cint(forum_setting(76))=1 then
				response.write " style=""FILTER: alpha(opacity=40)"" onMouseOut=nereidFade(this,40,10,10) onMouseOver=nereidFade(this,100,0,10)"
			end if
			response.write "></a></td><td width=20></td></tr></table>"
		else
			response.write "<table align=left><tr><td align=right><a href=list.asp?boardid="&rs("boardid")&"><img src="&rs("indeximg")&" align=top border=0"
			if cint(forum_setting(76))=1 then
				response.write " style=""FILTER: alpha(opacity=40)"" onMouseOut=nereidFade(this,40,10,10) onMouseOver=nereidFade(this,100,0,10)"
			end if
			response.write "></a></td><td width=20></td></tr></table>"
		end if
	else
		response.write "<table align=left><tr><td align=right><a href=list.asp?boardid="&rs("boardid")&"><img src="&rs("indeximg")&" align=top border=0"
		if cint(forum_setting(76))=1 then
			response.write " style=""FILTER: alpha(opacity=40)"" onMouseOut=nereidFade(this,40,10,10) onMouseOver=nereidFade(this,100,0,10)"
		end if
		response.write "></a></td><td width=20></td></tr></table>"
	end if
	'============= ͬѧ¼ ���� ==================
end if
response.write "</td><td width=200 rowspan=2 class=tablebody1>"
if Cint(reBoard_Setting(2))=1 then
response.write "��֤��̳������֤�û��������"
else
if isnull(LastPostInfo(7)) or LastPostInfo(7)="" then
	response.write "���⣺û���κ�����<BR>���ߣ���<BR>���ڣ�" & FormatDateTime(now(),1) & "��" & FormatDateTime(now(),4)&"&nbsp;<IMG border=0 src="""&Forum_info(7)&Forum_statePic(5)&""" alt=""û���κ�����"">"
else
	response.write "���⣺<a href=""dispbbs.asp?Boardid="&LastPostInfo(7)&"&ID="&LastPostInfo(6)&"&replyID="&LastPostInfo(1)&"&skin=1"">"&htmlencode(left(LastPostInfo(3),10))&"..</a><BR>���ߣ�<a href=""dispuser.asp?id="&htmlencode(LastPostInfo(5))&""" target=_blank>"&htmlencode(LastPostInfo(0))&"</a><BR>���ڣ�" & FormatDateTime(LastPostInfo(2),1) & "��" & FormatDateTime(LastPostInfo(2),4)&"&nbsp;<a href=""dispbbs.asp?Boardid="&LastPostInfo(7)&"&ID="&LastPostInfo(6)&"&replyID="&LastPostInfo(1)&"&skin=1""><IMG border=0 src="""&Forum_info(7)&Forum_statePic(5)&""" title=""ת����"&htmlencode(LastPostInfo(3))&"""></a>"
end if
end if
response.write "</TD></TR><TR><TD width=*><img src="&Forum_info(7)&Forum_pic(8)&" align=middle> "
response.write rs("readme")
response.write"</TD>"

response.write "</TR><TR><TD class=tablebody2 height=20 width=*>������"&master_4&""

response.write "</TD><td width=40 align=center class=tablebody2>&nbsp;</td><TD vAlign=middle class=tablebody2 width=200><table width=100% border=0><tr>"

response.write "<td width=25% vAlign=middle><img src="&Forum_info(7)&Forum_pic(9)&" alt=������ align=absmiddle>&nbsp;<font color="&Forum_body(8)&">"&rs("TodayNum")&"</font></td><td width=30% vAlign=middle><img src="&Forum_info(7)&Forum_pic(11)&" alt=���� border=0  align=absmiddle>&nbsp;"&rs("LastTopicNum")&"</td><td width=45% vAlign=middle><img src="&Forum_info(7)&Forum_pic(10)&" alt=���� border=0 align=absmiddle>&nbsp;"&rs("LastBBSNum")&"</td></tr>"
response.write "</table></TD></TR></TBODY></TABLE></td></tr></table></td></tr>"
'============================End===============================
end if
end sub

sub main_1_boardlist_2()
TdTableWidth=100/ColSpanNum
if Board_Count=ColSpanNum or Board_Count=0 then response.write "<tr>"
response.write "<td class=tablebody1 width="""&TdTableWidth&"%"">"
	if (Cint(reBoard_Setting(1))=1 and Cint(GroupSetting(37))=1) or Cint(reBoard_Setting(1))=0 then
	response.write "<TABLE cellSpacing=2 cellPadding=2 width=100% border=0><tr><td width=""100%"" title="""&htmlencode(rs("readme"))&"<br>������"&master_4&""" colspan=2><a href=list.asp?boardid="&rs("boardid")&"><font color="&Forum_body(14)&">"&rs("boardtype")&"</font></a>"
	if rs("child")>0 then
	response.write "(<font title=""��"&rs("child")&"��������̳"">"&rs("child")&"</font>)"
	end if
	response.write "</td></tr><tr><td width=""50%"">���գ�<font color="&Forum_body(8)&">"&rs("TodayNum")&"</font></td><td width=""50%"">������"&rs("LastBbsNum")&"</td></tr></table>"
	end if
response.write "</td>"
if Board_Count=ColSpanNum-1 then response.write "</tr>"
if Cint(Board_Count)>Cint(ColSpanNum)-1 then
	Board_Count=1
else
	Board_Count=Board_Count+1
end if
end sub

sub listinfo()
response.write "<TABLE cellpadding=3 cellspacing=1 class=tableborder1 align=center>"&_
	"<TR>"&_
	"<Th height=25 width=100% align=left id=tabletitlelink style=""font-weight:normal"">������<b>"&Forum_OnlineNum&"</b>�ˣ�����"& boardtype &"�Ϲ��� <b>"& onlinenum &"</b> λ��Ա�� <b>"& guestnum &"</b> λ����.�������� <b><font color="&Forum_body(8)&">"& todaynum &"</font></b> "
if request("action")="show" then
	response.write "[<a href=list.asp?action=off&boardID="& boardid&"&amp;page=1&skin="& skin &">�ر���ϸ�б�</a>]"
else
	if cint(Forum_Setting(14))=1 and request("action")<>"off" then
	response.write "[<a href=list.asp?action=off&boardID="& boardid&"&amp;page=1&skin="& skin &">�ر���ϸ�б�</a>]"
	else
	response.write "[<a href=list.asp?action=show&boardID="& boardid&"&amp;page=1&skin="& skin &">��ʾ��ϸ�б�</a>]"
	end if
end if

response.write "</Th></TR>"

if request("action")="off" then
	call onlineuser(0,0,boardid)
elseif request("action")="show" then
	call onlineuser(1,1,boardid)
else
	call onlineuser(Forum_Setting(14),Forum_Setting(15),boardid)
end if

response.write "</td></tr></TABLE><BR>"

response.write "<table cellpadding=0 cellspacing=0 border=0 width="&Forum_body(12)&" align=center valign=middle><tr>"&_
	"<td align=center width=2> </td>"&_
	"<td align=left>"
if (Cint(Board_Setting(43))=0 and Cint(Board_Setting(0))=0) or (Cint(Board_Setting(43))=0 and Cint(Board_Setting(0))=1 and (master or superboardmaster or boardmaster)) then
	select case cint(forum_setting(73))
	case 1
		if myUserPPD*cint(forum_setting(74))>myTodayTopic then
			response.write "<a href=""announce.asp?BoardID="&BoardID&"""><img src="""&Forum_info(7)&Forum_boardpic(1)&""" alt=""������"" border=0></a>&nbsp;&nbsp;"
			response.write "<a href=""vote.asp?BoardID="&BoardID&"""><img src="""&Forum_info(7)&Forum_boardpic(2)&""" alt=""������ͶƱ"" border=0></a>&nbsp;&nbsp;"
		end if
	case 3
		if myUserPPD*cint(forum_setting(74))>myTodayTopic+myTodayReply then
			response.write "<a href=""announce.asp?BoardID="&BoardID&"""><img src="""&Forum_info(7)&Forum_boardpic(1)&""" alt=""������"" border=0></a>&nbsp;&nbsp;"
			response.write "<a href=""vote.asp?BoardID="&BoardID&"""><img src="""&Forum_info(7)&Forum_boardpic(2)&""" alt=""������ͶƱ"" border=0></a>&nbsp;&nbsp;"
		end if
	case else
		response.write "<a href=""announce.asp?BoardID="&BoardID&"""><img src="""&Forum_info(7)&Forum_boardpic(1)&""" alt=""������"" border=0></a>&nbsp;&nbsp;"
		response.write "<a href=""vote.asp?BoardID="&BoardID&"""><img src="""&Forum_info(7)&Forum_boardpic(2)&""" alt=""������ͶƱ"" border=0></a>&nbsp;&nbsp;"
	end select
	response.write "<a href=smallpaper.asp?boardid="&boardid&"><img src="&Forum_info(7)&Forum_boardpic(3)&" border=0 alt=����С�ֱ�></a>"
elseif Cint(Board_Setting(0))=1 then
response.write "<B>����̳����������������з���/�����Ȳ���</B>"
end if
response.write "</td>"&_
	"<td align=right><img src="&Forum_info(7)&Forum_pic(13)&" align=absmiddle>  "& master_2
response.write "<select onchange='document.location=""list.asp?boardid="&request("boardid")&"&listtime=""+this.options[this.selectedIndex].value;'>"
response.write "<option value=""1"""
if request("listtime")<>"2" and request("listtime")<>"3" and request("listtime")<>"4" and request("listtime")<>"5" and request("listtime")<>"6" and request("listtime")<>"7" then	response.write " selected"
response.write ">ȫ����ʾ����</option>"
response.write "<option value=""2"""
if request("listtime")="2" then response.write " selected"
response.write ">����������</option>"
response.write "<option value=""3"""
if request("listtime")="3" then response.write " selected"
response.write ">����������</option>"
response.write "<option value=""4"""
if request("listtime")="4" then response.write " selected"
response.write ">һ��������</option>"
response.write "<option value=""5"""
if request("listtime")="5" then response.write " selected"
response.write ">����������</option>"
response.write "<option value=""6"""
if request("listtime")="6" then response.write " selected"
response.write ">����������</option>"
response.write "<option value=""7"""
if request("listtime")="7" then response.write " selected"
response.write ">����������</option>"
response.write "</select>"
response.write "</td></tr></table>"

response.write "<TABLE cellPadding=1 cellSpacing=1 class=tableborder1 align=center>"&_
	"<tr><td class=tablebody1 colspan=6 height=20>"&_
	"<table width=100% >"&_
	"<tr>"&_
	"<td valign=middle height=20 width=50> <a href=AllPaper.asp?boardid="&boardid&" title=����鿴����̳����С�ֱ�><b>�㲥��</b></a> </td><td width=*> "&_
	"<marquee scrolldelay=150 scrollamount=4 onmouseout=""if (document.all!=null){this.start()}"" onmouseover=""if (document.all!=null){this.stop()}"">"
set rs=conn.execute("SELECT TOP 5 s_id,s_username,s_title FROM smallpaper where datediff('d',s_addtime,now())<=1 and s_boardid="&boardid&" ORDER BY s_addtime desc")
do while not rs.eof
	response.write "����<font color="&Forum_pic(5)&">"&htmlencode(rs(1))&"</font>˵��<a href=javascript:openScript('viewpaper.asp?id="&rs(0)&"&boardid="&boardid&"',500,400)>"&htmlencode(rs(2))&"</a>"
rs.movenext
loop
set rs=nothing

response.write "</marquee></td><td align=right width=280><a href=elist.asp?boardid="& boardid &" title=�鿴���澫��><font color="&Forum_body(8)&"><B>����</B></font></a> | <a href=queryresult.asp?boardid="& boardid &"&searchscore=A1&searchdate=ALL&stype=2&nsearch=&psearch=1 title=�鿴����÷���><font color=blue><B>����</B></font></a> | <a href=cookies.asp?action=hasview&boardid="&boardid&" title=����������������Ϊ�Ѷ�>�Ѷ�</a> | <a href=boardstat.asp?reaction=online&boardid="&boardid&" title=�鿴����������ϸ���>����</a> | <a href=bbseven.asp?boardid="&boardid&" title=�鿴�����¼�>�¼�</a> | <a href=BoardPermission.asp?boardid="&boardid&" title=�鿴�����û���Ȩ��>Ȩ��</a> | <a href=admin_boardset.asp?boardid="&boardid&">����</a>"
if isaudit=1 then
response.write " | <a href=admin_topiclist.asp?boardid="&boardid&">���</a>"
end if
response.write "</td></tr></table>"&_
	"</td></tr>"&_
	"<form action=admin_batch.asp method=post name=batch>"&_
	"<TR align=middle>"
response.write "<Th height=25 width=32 id=tabletitlelink>"
response.write "<a href=list.asp?boardid="&boardid&"&page="&request("page")
if request("action")<>"batch" then
	response.write "&action=batch"
end if
response.write " alt=""�����������������ʽ"">״̬</a>"
response.write "</th>"
response.write "<Th width=* id=tabletitlelink>"
response.write "<a href=""list.asp?boardid="&boardid
if currentPage>1 then
	response.write "&page="&currentPage
end if
if request("action")<>"" then
	response.write "&action="&request("action")
end if
response.write "&orders=3"
if yorder="title" then
	if yordertype=" desc" then
		response.write "&ordertype=1"
	end if
end if
response.write """ alt=""�����������������"">"
response.write "����"
response.write "</a>"
if yorder="title" then
	response.write "<font face=webdings color="&forum_body(8)&">"
	if yordertype=" desc" then
		response.write "6"
	else
		response.write "5"
	end if
	response.write "</font>"
end if
response.write "&nbsp;&nbsp;(��<img src="&Forum_info(7)&Forum_boardpic(15)&" align=absmiddle>����չ�������б�)"
response.write "&nbsp;[<a href=""list.asp?boardid="&boardid
if currentPage>1 then
	response.write "&page="&currentPage
end if
if request("action")<>"" then
	response.write "&action="&request("action")
end if
response.write """>�ָ�Ĭ������״̬</a>]"
response.write "</Th>"
response.write "<Th width=80 id=tabletitlelink>"
response.write "<a href=""list.asp?boardid="&boardid
if currentPage>1 then
	response.write "&page="&currentPage
end if
if request("action")<>"" then
	response.write "&action="&request("action")
end if
response.write "&orders=4"
if yorder="postusername" then
	if yordertype=" desc" then
		response.write "&ordertype=1"
	end if
end if
response.write """ alt=""�����������������"">"
response.write "����"
response.write "</a>"
if yorder="postusername" then
	response.write "<font face=webdings color="&forum_body(8)&">"
	if yordertype=" desc" then
		response.write "6"
	else
		response.write "5"
	end if
	response.write "</font>"
end if
response.write "</Th>"
response.write "<Th width=80 id=tabletitlelink>"
response.write "<a href=""list.asp?boardid="&boardid
if currentPage>1 then
	response.write "&page="&currentPage
end if
if request("action")<>"" then
	response.write "&action="&request("action")
end if
response.write "&orders=5"
if yorder="child" then
	if yordertype=" desc" then
		response.write "&ordertype=1"
	end if
end if
response.write """ alt=""��������ӻظ�������"">"
response.write "�ظ�"
response.write "</a>"
if yorder="child" then
	response.write "<font face=webdings color="&forum_body(8)&">"
	if yordertype=" desc" then
		response.write "6"
	else
		response.write "5"
	end if
	response.write "</font>"
end if
response.write "/"
response.write "<a href=""list.asp?boardid="&boardid
if currentPage>1 then
	response.write "&page="&currentPage
end if
if request("action")<>"" then
	response.write "&action="&request("action")
end if
response.write "&orders=6"
if yorder="hits" then
	if yordertype=" desc" then
		response.write "&ordertype=1"
	end if
end if
response.write """ alt=""��������ӵ��������"">"
response.write "����"
response.write "</a>"
if yorder="hits" then
	response.write "<font face=webdings color="&forum_body(8)&">"
	if yordertype=" desc" then
		response.write "6"
	else
		response.write "5"
	end if
	response.write "</font>"
end if
response.write "</Th>"
response.write "<Th width=70 id=tabletitlelink>"
response.write "<a href=""list.asp?boardid="&boardid
if currentPage>1 then
	response.write "&page="&currentPage
end if
if request("action")<>"" then
	response.write "&action="&request("action")
end if
response.write "&orders=1"
if yorder="lastposttime" then
	if yordertype=" desc" then
		response.write "&ordertype=1"
	end if
end if
response.write """ alt=""���������������ʱ������"">"
response.write "������"
response.write "</a>"
if yorder="lastposttime" then
	response.write "<font face=webdings color="&forum_body(8)&">"
	if yordertype=" desc" then
		response.write "6"
	else
		response.write "5"
	end if
	response.write "</font>"
end if
response.write "</Th>"
response.write "<Th width=80 id=tabletitlelink>"
response.write "<a href=""list.asp?boardid="&boardid
if currentPage>1 then
	response.write "&page="&currentPage
end if
if request("action")<>"" then
	response.write "&action="&request("action")
end if
response.write "&orders=7"
if yorder="lastpost" then
	if yordertype=" desc" then
		response.write "&ordertype=1"
	end if
end if
response.write """ alt=""������������ظ�������"">"
response.write "�ظ���"
response.write "</a>"
if yorder="lastpost" then
	response.write "<font face=webdings color="&forum_body(8)&">"
	if yordertype=" desc" then
		response.write "6"
	else
		response.write "5"
	end if
	response.write "</font>"
end if
response.write "</Th>"
response.write "</TR>"
if RsTopic.bof and RsTopic.eof and AllTopNum=0 then
	response.write "<tr><td colSpan=6 width=100% class=tablebody1 height=25>����̳�������ݣ���ӭ��������</td></tr>"
else
'=========================All Top=========================
if currentpage=1 and alltopnum>0 then
	do while not RsTop.eof
response.write "<TR align=middle><TD class=tablebody2 width=32 height=27>"
response.write TopicIcon(rstop("BoardID"),rstop("topicid"),rstop("lastposttime"),2,rstop("isVote"),rstop("LockTopic"),iif(RsTop("child")>=Cint(forum_setting(44)),1,0))
response.write "</TD><TD align=left onmouseover=this.className='TableBody2' onmouseout=this.className='TableBody1' class=tablebody1 width=*>"
if RsTop(6)=0 then
	response.write "<img src="""& Forum_info(7)&Forum_boardpic(16)&""" id=""followImg"& RsTop("topicid") &""">"
else
	response.write "<img loaded=""no"" src="""& Forum_info(7)&Forum_boardpic(15)&""" id=""followImg"& RsTop("topicid") &""" style=""cursor:hand;"" onclick=""loadThreadFollow("& RsTop("topicid") &","& RsTop("boardid") &")"" title=չ�������б�>"
end if

'Response.Write "<img src="&Forum_info(8)&rstop("Expression")&" border=0 alt=��������>&nbsp;"

if not isnull(RsTop("lastpost")) then
	LastPost=split(RsTop("lastpost"),"$")
	if ubound(LastPost)=7 then
	Lastuser=htmlencode(LastPost(0))
	LastID=LastPost(1)
	LastTime=LastPost(2)
	if not isdate(LastTime) then LastTime=RsTop("dateandtime")
	body=htmlencode(LastPost(3))
	if not isnull(LastPost(4)) and LastPost(4)<>"" then response.write "<img src=""images/files/"&LastPost(4)&".gif"" >"
	LastUserid=LastPost(5)
	LastRootid=LastPost(6)
	else
	Lastuser="------"
	LastID=RsTop("topicid")
	LastTime=RsTop("dateandtime")
	body="..."
	LastUserid=RsTop("postuserid")
	LastRootid=RsTop("topicid")
	end if
else
	Lastuser="------"
	LastID=RsTop("topicid")
	LastTime=RsTop("dateandtime")
	body="..."
	LastUserid=RsTop("postuserid")
	LastRootid=RsTop("topicid")
end if

if LastTime=RsTop("dateandtime") then LastUser=""

if isnull(RsTop("ChenTopicColor")) or RsTop("ChenTopicColor")=""  then 
	response.write "<a href=""dispbbs.asp?boardID="& RsTop("boardid") &"&ID="& RsTop("topicid") &""" title=""<img src=pic/zhuti.gif> ��"& htmlencode(RsTop("title")) &"��<br><img src=pic/zuozhe.gif> ���ߣ�"& htmlencode(RsTop("postusername")) &"<br><img src=pic/lasttime.gif> �����ڣ�"& RsTop("dateandtime") &"<br><img src=pic/lastpost.gif> ��������"& body &"..."">"
else
	response.write "<a href=""dispbbs.asp?boardID="& RsTop("boardid") &"&ID="& RsTop("topicid") &"""><font color="""&RsTop("ChenTopicColor")&""" title=""<img src=pic/zhuti.gif> ��"& htmlencode(RsTop("title")) &"��<br><img src=pic/zuozhe.gif> ���ߣ�"& htmlencode(RsTop("postusername")) &"<br><img src=pic/lasttime.gif> �����ڣ�"& RsTop("dateandtime") &"<br><img src=pic/lastpost.gif> ��������"& body &"..."">"
end if 

if len(RsTop("title"))>26 then
	response.write left(htmlencode(RsTop("title")),26)&"..."
else
	response.write htmlencode(RsTop("title"))
end if

if not (isnull(RsTop("ChenTopicColor"))  or RsTop("ChenTopicColor")="")  then 
	response.write "</font>"
end if 

response.write "</a>"
replynum=RsTop("child")+1
if replynum>Cint(Board_Setting(27)) then
	response.write "&nbsp;&nbsp;[<img src="&Forum_info(7)&Forum_statePic(6)&"><b>"
  	if replynum mod Cint(Board_Setting(27))=0 then
     		Pnum= replynum \ Cint(Board_Setting(27))
  	else
     		Pnum= replynum \ Cint(Board_Setting(27))+1
  	end if
	for p=1 to Pnum
	response.write " <a href='dispbbs.asp?boardID="& RsTop("boardID") &"&ID="& RsTop("topicid") &"&star="&P&"'><font color="&Forum_body(8)&">"&p&"</font></a> "
	if p+1>7 and p<Pnum then
	response.write "... <a href='dispbbs.asp?boardID="& RsTop("boardID") &"&ID="& RsTop("topicid") &"&star="&Pnum&"'><font color="&Forum_body(8)&">"&Pnum&"</font></a>"
	exit for
	end if
	next
	response.write "</b>]"
else
	pnum=1
end if

if datediff("h",LastTime,now)<=1 then
   response.write("<img src=pic/topnew.gif alt=1Сʱ�и������ӱ�־ >")
 elseif datediff("h",LastTime,now)<=2 then 
   response.write("<img src=pic/topnew1.gif alt=2Сʱ�и������ӱ�־ >") 
 elseif datediff("h",LastTime,now)<=4 then 
   response.write("<img src=pic/topnew2.gif alt=4Сʱ�и������ӱ�־ >") 
end if 
if rstop("isbest")=1 then
	response.write "&nbsp;&nbsp;<img src="""&Forum_info(7)&Forum_statePic(4)&""">"
end if
response.write "</TD>"
response.write "<TD class=tablebody2 width=80><a href=""dispuser.asp?id="& RsTop("postuserid") &""" target=_blank>"& htmlencode(RsTop("postusername")) &"</a></TD>"
response.write "<TD class=tablebody1 width=80>"

if RsTop("isvote")=1 then
	response.write "<FONT color="&Forum_body(8)&"><b>"&RsTop("votetotal")&"</b></font>  Ʊ"
else
	response.write RsTop("child") &"/"& RsTop("hits")
end if

response.write "</TD>"
response.write "<TD align=center class=tablebody2 width=70><a href=""dispbbs.asp?boardid="& RsTop("boardID") &"&id="&LastRootID&"&star="&pnum&"#"& LastID&""">"
response.write month(LastTime)&"-"&day(LastTime)&"&nbsp;"&FormatDateTime(LastTime,4)&"</a></TD>"
response.write "<TD align=center class=tablebody1 width=80>"

if LastUser="" then
	response.write "------"
else
	response.write "<a href=dispuser.asp?id="&LastUserID&" target=_blank>"&LastUser&"</a>"
end if

response.write "</TD></TR>"
response.write "<tr style=""display:none"" id=""follow"& RsTop("topicid") &"""><td colspan=6 id=""followTd"& RsTop("topicid") &""" style=""padding:0px""><div style=""width:240px;margin-left:18px;border:1px solid black;background-color:lightyellow;color:black;padding:2px"" onclick=""loadThreadFollow("& RsTop("topicid") &","&RsTop("boardid")&")"">���ڶ�ȡ���ڱ�����ĸ��������Ժ��</div></td></tr>"
	Lastuser=""
	LastID=""
	LastTime=""
	body=""
	RsTop.movenext
	loop
	set RsTop=nothing
end if
'============================end==========================
if totalrec mod Cint(Board_Setting(26))=0 then
	n= totalrec \ Cint(Board_Setting(26))
else
	n= totalrec \ Cint(Board_Setting(26))+1
end if
if not (RsTopic.bof and RsTopic.eof) then
	if currentpage > n then currentpage = n
	if currentpage<1 then currentpage=1
	rstopic.PageSize = Cint(Board_Setting(26))
	rstopic.AbsolutePage=currentpage
	dim endlistnum
	if currentpage=n then
		endlistnum=totalrec-(currentpage-1) * Cint(Board_Setting(26))
	else
		endlistnum=Cint(Board_Setting(26))
	end if
	'page_count=0
	'do while not RsTopic.eof and page_count<Cint(Board_Setting(26))
	for i=1 to endlistnum
	'page_count=page_count+1
'========================TopicInfo=========================
response.write "<TR align=middle><TD class=tablebody2 width=32 height=27>"
if request("action")="batch" and Cint(GroupSetting(45))=1 then
	response.write "<input type=checkbox name=Announceid value="&RsTopic("topicid")&">"
else
	response.write TopicIcon(rstopic("BoardID"),rstopic("topicid"),rstopic("lastposttime"),RsTopic("istop"),rstopic("isVote"),rstopic("LockTopic"),iif(RsTopic("child")>=Cint(forum_setting(44)),1,0))
end if

response.write "</TD><TD align=left onmouseover=this.className='TableBody2' onmouseout=this.className='TableBody1' class=tablebody1 width=*>"

if RsTopic(6)=0 then
	response.write "<img src="""& Forum_info(7)&Forum_boardpic(16)&""" id=""followImg"& RsTopic("topicid") &""">"
else
	response.write "<img loaded=""no"" src="""& Forum_info(7)&Forum_boardpic(15)&""" id=""followImg"& RsTopic("topicid") &""" style=""cursor:hand;"" onclick=""loadThreadFollow("& RsTopic("topicid") &","& RsTopic("boardid") &")"" title=չ�������б�>"
end if

'Response.Write "<img src="&Forum_info(8)&rstopic("Expression")&" border=0 alt=��������>&nbsp;"

if not isnull(RsTopic("lastpost")) then
	LastPost=split(RsTopic("lastpost"),"$")
	if ubound(LastPost)=7 then
	Lastuser=htmlencode(LastPost(0))
	LastID=LastPost(1)
	LastTime=LastPost(2)
	if not isdate(LastTime) then LastTime=RsTopic("dateandtime")
	body=htmlencode(LastPost(3))
	if not isnull(LastPost(4)) and LastPost(4)<>"" then response.write "<img src=""images/files/"&LastPost(4)&".gif"" >"
	LastUserid=LastPost(5)
	LastRootid=LastPost(6)
	else
	Lastuser="------"
	LastID=RsTopic("topicid")
	LastTime=RsTopic("dateandtime")
	body="..."
	LastUserid=RsTopic("postuserid")
	LastRootid=RsTopic("topicid")
	end if
else
	Lastuser="------"
	LastID=RsTopic("topicid")
	LastTime=RsTopic("dateandtime")
	body="..."
	LastUserid=RsTopic("postuserid")
	LastRootid=RsTopic("topicid")
end if
if LastTime=RsTopic("dateandtime") then LastUser=""

if isnull(RsTopic("ChenTopicColor")) or RsTopic("ChenTopicColor")=""  then 
	response.write "<a href=""dispbbs.asp?boardID="& RsTopic("boardid") &"&ID="& RsTopic("topicid") &""" title=""<img src=pic/zhuti.gif> ��"& htmlencode(RsTopic("title")) &"��<br><img src=pic/zuozhe.gif> ���ߣ�"& htmlencode(RsTopic("postusername")) &"<br><img src=pic/lasttime.gif> �����ڣ�"& RsTopic("dateandtime") &"<br><img src=pic/lastpost.gif> ��������"& body &"..."">"
else
	response.write "<a href=""dispbbs.asp?boardID="& RsTopic("boardid") &"&ID="& RsTopic("topicid") &"""><font color="""&RsTopic("ChenTopicColor")&""" title=""<img src=pic/zhuti.gif> ��"& htmlencode(RsTopic("title")) &"��<br><img src=pic/zuozhe.gif> ���ߣ�"& htmlencode(RsTopic("postusername")) &"<br><img src=pic/lasttime.gif> �����ڣ�"& RsTopic("dateandtime") &"<br><img src=pic/lastpost.gif> ��������"& body &"..."">"
end if 

if len(RsTopic("title"))>26 then
	response.write left(htmlencode(RsTopic("title")),26)&"..."
else
	response.write htmlencode(RsTopic("title"))
end if

if not (isnull(RsTopic("ChenTopicColor"))  or RsTopic("ChenTopicColor")="")  then 
	response.write "</font>"
end if 

response.write "</a>"
replynum=RsTopic("child")+1
if replynum>Cint(Board_Setting(27)) then
	response.write "&nbsp;&nbsp;[<img src="&Forum_info(7)&Forum_statePic(6)&"><b>"
  	if replynum mod Cint(Board_Setting(27))=0 then
     		Pnum= replynum \ Cint(Board_Setting(27))
  	else
     		Pnum= replynum \ Cint(Board_Setting(27))+1
  	end if
	for p=1 to Pnum
	response.write " <a href='dispbbs.asp?boardID="& RsTopic("boardID") &"&ID="& RsTopic("topicid") &"&star="&P&"'><font color="&Forum_body(8)&">"&p&"</font></a> "
	if p+1>7 and p<Pnum then
	response.write "... <a href='dispbbs.asp?boardID="& RsTopic("boardID") &"&ID="& RsTopic("topicid") &"&star="&Pnum&"'><font color="&Forum_body(8)&">"&Pnum&"</font></a>"
	exit for
	end if
	next
	response.write "</b>]"
else
	pnum=1
end if
if datediff("h",LastTime,now)<=1 then
   response.write("<img src=pic/topnew.gif alt=1Сʱ�и������ӱ�־ >")
 elseif datediff("h",LastTime,now)<=2 then 
   response.write("<img src=pic/topnew1.gif alt=2Сʱ�и������ӱ�־ >") 
 elseif datediff("h",LastTime,now)<=4 then 
   response.write("<img src=pic/topnew2.gif alt=4Сʱ�и������ӱ�־ >") 
end if 
if rstopic("isbest")=1 then
	response.write "&nbsp;&nbsp;<img src="""&Forum_info(7)&Forum_statePic(4)&""">"
end if
response.write "</TD>"
response.write "<TD class=tablebody2 width=80><a href=""dispuser.asp?id="& RsTopic("postuserid") &""" target=_blank>"& htmlencode(RsTopic("postusername")) &"</a></TD>"
response.write "<TD class=tablebody1 width=80>"

if RsTopic("isvote")=1 then
	response.write "<FONT color="&Forum_body(8)&"><b>"&RsTopic("votetotal")&"</b></font>  Ʊ"
else
	response.write RsTopic("child") &"/"& RsTopic("hits")
end if

response.write "</TD>"
response.write "<TD align=center class=tablebody2 width=70><a href=""dispbbs.asp?boardid="& RsTopic("boardID") &"&id="&LastRootID&"&star="&pnum&"#"& LastID&""">"
response.write month(LastTime)&"-"&day(LastTime)&"&nbsp;"&FormatDateTime(LastTime,4)&"</a></TD>"
response.write "<TD align=center class=tablebody1 width=80>"

if LastUser="" then
	response.write "------"
else
	response.write "<a href=dispuser.asp?id="&LastUserID&" target=_blank>"&LastUser&"</a>"
end if

response.write "</TD></TR>"
response.write "<tr style=""display:none"" id=""follow"& RsTopic("topicid") &"""><td colspan=6 id=""followTd"& RsTopic("topicid") &""" style=""padding:0px""><div style=""width:240px;margin-left:18px;border:1px solid black;background-color:lightyellow;color:black;padding:2px"" onclick=""loadThreadFollow("& RsTopic("topicid") &","&RsTopic("boardid")&")"">���ڶ�ȡ���ڱ�����ĸ��������Ժ��</div></td></tr>"
'===========================End============================
	Lastuser=""
	LastID=""
	LastTime=""
	body=""
	RsTopic.movenext
	'loop
	next
	end if
	set RsTopic=nothing
	if request("action")="batch" then

	response.write "<tr><td height=30 width=100% class=tablebody1 colspan=6>��<select name=newboard size=1>"
	response.write "<option selected>�ƶ�������ѡ��</option>"

	set rs=conn.execute("select boardid,boardtype,depth from board order by rootid,orders")
	do while not rs.EOF
	response.write "<option value="""&rs(0)&""" "
	response.write ">"
	select case rs(2)
	case 0
	response.write "��"
	case 1
	response.write "&nbsp;&nbsp;��"
	end select
	if rs(2)>1 then
	for i=2 to rs(2)
		response.write "&nbsp;&nbsp;��"
	next
	response.write "&nbsp;&nbsp;��"
	end if
	response.write rs(1)&"</option>"
	rs.MoveNext
	loop
	set rs=nothing

response.write "<input type=hidden value="&boardid&" name=boardid><input type=checkbox name=chkall value=on onclick=""CheckAll(this.form)"">ȫѡ/ȡ�� <input type=radio name=action value=dele>����ɾ�� <input type=radio name=action value=move>�����ƶ� <input type=radio name=action value=isbest>�������� <input type=radio name=action value=lock>�������� <input type=radio name=action value=istop>�����̶� <input type=submit name=Submit value=ִ��  onclick=""{if(confirm('��ȷ��ִ�д˲�����?')){this.document.batch.submit();return true;}return false;}""></td></tr>"
	end if
	response.write "</form>"
end if
set rs=nothing
if totalrec mod Board_Setting(26)=0 then
	n= totalrec \ Board_Setting(26)
else
	n= totalrec \ Board_Setting(26)+1
end if
response.write "</table>"

response.write "<table border=0 cellpadding=0 cellspacing=3 width="&Forum_body(12)&" align=center >"&_
 "</form>"&_
 "<form method=post action=list.asp?boardid="&boardid&"&action="&request("action")&">"&_
 "<tr>"&_
 "<td valign=middle nowrap>ҳ�Σ�<b>"& currentPage &"</b>/<b>"& n &"</b>ҳ ÿҳ<b>"& Board_Setting(26) &"</b> ������<b>"& totalrec &"</b></td>"&_
 "<td valign=middle align=right >"
call DispPageNum(currentpage,n,"'?boardid="&boardid&"&page=","&action="&request("action")&"&listtime="&request("listtime")&"&orders="&request("orders")&"&ordertype="&request("ordertype")&"'")
response.write "ת��:<input type=text name=Page size=3 maxlength=10  value="& currentpage &"><input type=submit value=Go name=submit>&nbsp;</td></tr></form></table>"

response.write "<table border=0 cellpadding=0 cellspacing=3 width="&Forum_body(12)&" align=center>"&_
	"<tr>"&_
	"<FORM METHOD=POST ACTION=queryresult.asp?boardid="&boardid&"&sType=2&SearchDate=30&pSearch=1>"&_
	"<td width=50% valign=middle nowrap height=40>"&_
	"����������<input type=text name=keyword>&nbsp;<input type=submit name=submit value=����>"&_
	"</td>"&_
	"</FORM>"&_
	"<td valign=middle nowrap width=50% > <div align=right>"&_
	"<select onchange=""if(this.options[this.selectedIndex].value!=''){location=this.options[this.selectedIndex].value;}"">"&_
	"<option selected>��ת��̳��...</option>"
myCache.name="BoardJumpList"
if myCache.valid then
response.write myCache.value
else
Dim BoardJumpList
set rs=conn.execute("select boardid,boardtype,depth from board order by rootid,orders")
do while not rs.EOF
BoardJumpList = BoardJumpList & "<option value=""list.asp?boardid="&rs(0)&""" "
BoardJumpList = BoardJumpList & ">"
select case rs(2)
case 0
BoardJumpList = BoardJumpList & "��"
case 1
BoardJumpList = BoardJumpList & "&nbsp;&nbsp;��"
end select
if rs(2)>1 then
for i=2 to rs(2)
	BoardJumpList = BoardJumpList & "&nbsp;&nbsp;��"
next
BoardJumpList = BoardJumpList & "&nbsp;&nbsp;��"
end if
BoardJumpList = BoardJumpList & rs(1)&"</option>"
rs.MoveNext
loop
set rs=nothing
myCache.add BoardJumpList,dateadd("n",60,now)
response.write BoardJumpList
end if
response.write "</select><div></td></tr></table><BR>"

response.write "<table cellspacing=1 cellpadding=3 width=100% class=tableborder1 align=center><tr>"&_
	"<th width=80% align=left>��-=> "& Forum_info(0) &"ͼ��</th>"&_
	"<th noWrap width=20% align=right>����ʱ���Ϊ - "&Forum_info(9)&" &nbsp;</th>"&_
	"</tr>"&_
	"<tr><td colspan=2 class=tablebody1>"&_
	"<table cellspacing=4 cellpadding=0 width=90% border=0 align=center>"&_
	"<tr>"&_
	"<td><img src=pic/topicicon/00000.gif> ������</td>"&_
	"<td><img src=pic/topicicon/10000.gif> ������</td>"&_
	"<td><img src=pic/topicicon/00001.gif> ��������</td>"&_
	"<td><img src=pic/topicicon/00010.gif> ��������</td>"&_
	"<td><img src=pic/topicicon/00100.gif> ͶƱ</td>"&_
	"<td><img src=pic/topicicon/01000.gif> �̶�����</td>"&_
	"<td><img src=pic/topicicon/02000.gif> �̶ܹ�����</td>"&_
	"</tr>"&_
	"</table>"&_
	"</td></tr></table>"
end sub

sub BoardEmpty()
response.write "<table cellpadding=0 cellspacing=0 border=0 width="&Forum_body(12)&" align=center valign=middle><tr>"&_
	"<td align=center width=2> </td>"&_
	"<td align=left> <a href=""announce.asp?boardid="& boardid &"""><img src="&Forum_info(7)&Forum_boardpic(1)&" border=0 alt=������></a>"&_
	"&nbsp;&nbsp;<a href=vote.asp?boardid="&boardid&"><img src="&Forum_info(7)&Forum_boardpic(2)&" border=0 alt=������ͶƱ></a>"&_
	"&nbsp;&nbsp;<a href=smallpaper.asp?boardid="&boardid&"><img src="&Forum_info(7)&Forum_boardpic(3)&" border=0 alt=����С�ֱ�></a></td>"&_
	"<td align=right><img src="&Forum_info(7)&Forum_pic(13)&" align=absmiddle>  "& master_2 &"</td></tr></table>"&_
	"<TABLE cellPadding=1 cellSpacing=1 class=tableborder1 align=center>"&_
	"<TR align=middle>"&_
	"<Th height=25 width=32 id=tabletitlelink><a href=list.asp?boardid="&boardid&"&page="&request("page")&"&action=batch alt=""�����������������ʽ"">״̬</a></th>"&_
	"<Th width=*>�� ��  (��<img src="&Forum_info(7)&Forum_boardpic(15)&">����չ�������б�)</Th>"&_
	"<Th width=80>�� ��</Th>"&_
	"<Th width=80>�ظ�/����</Th>"&_
	"<Th width=70>������</Th>"&_
	"<Th width=80>�ظ���</Th>"&_
	"</TR>"&_
	"<tr><td colSpan=6 width=100% class=tablebody1 height=25>����̳�������ݣ���ӭ��������</td></tr></table><BR>"
end sub

'if instr(scriptname,"index.asp")>0 or instr(scriptname,"list.asp")>0 then
'if Forum_ads(2)=1 then
'call admove()
'end if
'if Forum_ads(13)=1 then
'call fixup()
'end if
'end if
%>
<!--#include file="online_l.asp"-->
<!--#include file="inc/ad_fixup.asp"-->