<!--#include file="conn.asp"-->
<!-- #include file="inc/const.asp" -->
<%
Server.ScriptTimeOut=999999
dim topicid
dim tablename
dim trs,UpdateBoardID
stats="���ӹ���"
topicid=replace(request("topicid"),"'","")
topicid=replace(topicid,";","")
topicid=replace(topicid,"--","")
topicid=replace(topicid,"(","")
if request("action")<>"��ջ���վ" then
	if topicid="" or isnull(topicid) then
		Errmsg=Errmsg+"<li>"+"��ѡ��������Ӻ���в�����"
		Founderr=true
	end if
end if
if request("tablename")="topic" then
	tablename="topic"
elseif instr(request("tablename"),"bbs")>0 then
	tablename=replace(request("tablename"),"'","")
else
	Errmsg=Errmsg+"<li>"+"�����ϵͳ������"
	Founderr=true
end if
if not master or session("flag")="" then
	Errmsg=Errmsg+"<li>"+"������ϵͳ����Ա��������û�е�¼��"
	Founderr=true
end if

call nav()
call head_var(0,0,"����վ","recycle.asp")
if founderr=true then
	call dvbbs_error()
else
	if request("action")="ɾ��" then
		call delete()
	elseif request("action")="��ԭ" then
		call redel()
	elseif request("action")="��ջ���վ" then
		call Alldel()
	else
		Errmsg=Errmsg+"<li>"+"��ָ�����������"
		Founderr=true
		call dvbbs_error()
	end if
end if
call footer()

'ɾ������վ����
sub delete()
if instr(tablename,"bbs")>0 then
	conn.execute("delete from "&tablename&" where locktopic=2 and Announceid in ("&TopicID&")")
elseif tablename="topic" then
	for i=0 to ubound(allposttable)
	conn.execute("delete from "&allposttable(i)&" where locktopic=2 and rootid in ("&TopicID&")")
	next
	conn.execute("delete from topic where locktopic=2 and topicid in ("&TopicID&")")
end if

sucmsg="<li>���Ӳ����ɹ���<li>���Ĳ�����Ϣ�Ѿ���¼�ڰ���"
call dvbbs_suc()
end sub

'��ԭ����վ����
sub redel()
dim tempnum,todaynum
if instr(tablename,"bbs")>0 then
	sql="update "&tablename&" set locktopic=0 where Announceid in ("&TopicID&")"
	conn.execute(sql)
	'����ûظ�����Ӧ������ɾ������ͬʱ��ԭ��������
	set rs=conn.execute("select topicid,posttable,boardid from topic where locktopic=0 and topicid in (select distinct rootid from "&tablename&" where Announceid in ("&TopicID&"))")
	do while not rs.eof
		conn.execute("update "&rs(1)&" set locktopic=0 where parentid=0 and rootid="&rs(0))
		set trs=conn.execute("select count(*) from "&rs(1)&" where locktopic=0 and not parentid=0 and rootid="&rs(0))
		conn.execute("update topic set child="&trs(0)&",locktopic=0 where topicid="&rs(0))
		conn.execute("update board set LastTopicNum=LastTopicNum+1 where boardid="&rs(2))
	rs.movenext
	loop
	set rs=conn.execute("select PostUserID,BoardID,DateAndtime,ParentID from "&tablename&" where Announceid in ("&TopicID&")")
	do while not rs.eof
	sql="update [user] set article=article+1,userWealth=userWealth+"&Forum_user(3)&",userEP=userEP+"&Forum_user(8)&",userdel=userdel-1 where userid="&rs(0)
	conn.execute(sql)
	set trs=conn.execute("select ParentStr,LastPost from board where boardid="&rs(1))
	if datediff("d",rs(2),split(trs(1),"$")(2))=0 then
	todaynum=1
	else
	todaynum=0
	end if
	call AllboardNumAdd(todayNum,1,0)
	UpdateBoardID=trs(0) & "," & rs(1)
	conn.execute("update board set todaynum=todaynum+"&todaynum&",LastbbsNum=LastbbsNum+1 where boardid in ("&UpdateBoardID&")")
	LastCount(rs(1))
	rs.movenext
	loop
	set rs=nothing
elseif tablename="topic" then
	dim TotalUseTable,LastPost_a
	i=0
	todaynum=0
	sql="update topic set locktopic=0 where topicid in ("&TopicID&")"
	conn.execute(sql)
	set rs=conn.execute("select topicid,posttable,boardid from topic where topicid in ("&topicid&")")
	do while not rs.eof
		set trs=conn.execute("select ParentStr,LastPost from board where boardid="&rs(2))
		UpdateBoardID=trs(0) & "," & rs(2)
		LastPost_a=split(trs(1),"$")(2)
		conn.execute("update "&rs(1)&" set locktopic=0 where rootid="&rs(0))
		set trs=conn.execute("select postuserid,dateandtime from "&rs(1)&" where rootid="&rs(0))
		do while not trs.eof
		i=i+1
		sql="update [user] set article=article+1,userWealth=userWealth+"&Forum_user(3)&",userEP=userEP+"&Forum_user(8)&",userdel=userdel-1 where userid="&trs(0)
		conn.execute(sql)
		if datediff("d",trs(1),now())=0 then
			todaynum=todaynum+1
		else
			todaynum=todaynum
		end if
		trs.movenext
		loop
		call AllboardNumAdd(todayNum,i,1)
		conn.execute("update board set todaynum=todaynum+"&todaynum&",LastbbsNum=LastbbsNum+"&i&",LastTopicNum=LastTopicNum+1 where boardid in ("&UpdateBoardID&")")
		LastCount(rs(2))
		i=0
		todaynum=0
	rs.movenext
	loop
	set rs=nothing
	set trs=nothing
end if
sucmsg="<li>���Ӳ����ɹ���<li>���Ĳ�����Ϣ�Ѿ���¼�ڰ���"
call dvbbs_suc()
end sub

'ȫ��ɾ������վ����
sub AllDel()
for i=0 to ubound(allposttable)
	conn.execute("delete from "&allposttable(i)&" where locktopic=2")
next
conn.execute("delete from topic where locktopic=2")
sucmsg="<li>���Ӳ����ɹ���<li>���Ĳ�����Ϣ�Ѿ���¼�ڰ���"
call dvbbs_suc()
end sub

function LastCount(boardid)
Dim LastTopic,body,LastRootid,LastPostTime,LastPostUser
Dim LastPost,uploadpic_n,Lastpostuserid,Lastid
set trs=conn.execute("select top 1 T.title,b.Announceid,b.dateandtime,b.username,b.postuserid,b.rootid from "&NowUseBBS&" b inner join Topic T on b.rootid=T.TopicID where b.boardid="&boardid&" and  b.locktopic<2 order by b.announceid desc")
if not(trs.eof and trs.bof) then
	Lasttopic=replace(left(trs(0),15),"$","")
	LastRootid=trs(1)
	LastPostTime=trs(2)
	LastPostUser=trs(3)
	LastPostUserid=trs(4)
	Lastid=trs(5)
else
	LastTopic="��"
	LastRootid=0
	LastPostTime=now()
	LastPostUser="��"
	LastPostUserid=0
	Lastid=0
end if
set trs=nothing

LastPost=LastPostUser & "$" & LastRootid & "$" & LastPostTime & "$" & LastTopic & "$" & uploadpic_n & "$" & LastPostUserID & "$" & LastID & "$" & BoardID
Dim SplitUpBoardID,SplitLastPost
SplitUpBoardID=split(UpdateBoardID,",")
For i=0 to ubound(SplitUpBoardID)
	set trs=conn.execute("select LastPost from board where boardid="&SplitUpBoardID(i))
	if not (trs.eof and trs.bof) then
	SplitLastPost=split(trs(0),"$")
	if ubound(SplitLastPost)=7 and clng(LastRootID)<>clng(SplitLastPost(1)) then
		conn.execute("update board set LastPost='"&LastPost&"' where boardid="&SplitUpBoardID(i))
	end if
	end if
Next
set trs=nothing
'sql="update board set LastPost='"&LastPost&"' where boardid in ("&UpdateBoardID&")"
'conn.execute(sql)
end function


'����ԭʱ,������̳����������
function AllboardNumAdd(todayNum,postNum,topicNum)
sql="update config set TodayNum=todayNum+"&todaynum&",BbsNum=bbsNum+"&postNum&",TopicNum=topicNum+"&TopicNum
conn.execute(sql)
end function


%>