<!--#include file="conn.asp"-->
<!-- #include file="inc/const.asp" -->
<%
stats="���ӹ���"
Dim replyid
Dim id
Dim Lasttopic,Lastpost
Dim lastrootid,lastpostuser
Dim ip,url
Dim title,content
Dim TotalUseTable
Dim Topic,TopicUsername,TopicUserID
Dim canlocktopic,candeltopic,canmovetopic,cantoptopic,canbesttopic,canawardtopic
dim canpbtopic
Dim Cantoptopic_a
Dim UpdateBoardID,UpdateBoardID_1
Dim MsgContent

canlocktopic=false
candeltopic=false
canmovetopic=false
cantoptopic=false
canbesttopic=false
canawardtopic=false
canpbtopic=false
cantoptopic_a=false
'����̳���ϼ���̳ID
UpdateBoardID=BoardParentStr & "," & BoardID
ip=replace(Request.ServerVariables("REMOTE_ADDR"),"'","")
if not founduser then
	founderr=true
	Errmsg=Errmsg+"<br>"+"<li>���¼����в�����"
end if
if request("boardid")="" then
	founderr=true
	Errmsg=Errmsg+"<br>"+"<li>��ָ����̳���档"
elseif not isInteger(request("boardid")) then
	founderr=true
	Errmsg=Errmsg+"<br>"+"<li>�Ƿ��İ��������"
else
	boardid=clng(boardid)
end if
if request("id")="" then
	founderr=true
	Errmsg=Errmsg+"<br>"+"<li>��ָ��������ӡ�"
elseif not isInteger(request("id")) then
	founderr=true
	Errmsg=Errmsg+"<br>"+"<li>�Ƿ������Ӳ�����"
else
	id=request("id")
end if
if isInteger(request("replyid")) then
	replyid=request("replyid")
end if

Dim doWealth,douserEP,douserCP
Dim doWealthL,douserEPL,douserCPL
Dim doWealthH,douserEPH,douserCPH
Dim doWealthMsg,douserEPMsg,douserCPMsg,allMsg
doWealthL=0:douserEPL=0:douserCPL=0
doWealthH=0:douserEPH=0:douserCPH=0
select case request("action")
case "delet"
	doWealthL=-Forum_user(3):douserEPL=-Forum_user(8):douserCPL=-Forum_user(13)
	doWealthH=0:douserEPH=0:douserCPH=0
case "ɾ������"
	doWealthL=-Forum_user(3):douserEPL=-Forum_user(8):douserCPL=-Forum_user(13)
	doWealthH=0:douserEPH=0:douserCPH=0
case "dele"
	doWealthL=-Forum_user(3):douserEPL=-Forum_user(8):douserCPL=-Forum_user(13)
	doWealthH=0:douserEPH=0:douserCPH=0
case "ɾ������"
	doWealthL=-Forum_user(3):douserEPL=-Forum_user(8):douserCPL=-Forum_user(13)
	doWealthH=0:douserEPH=0:douserCPH=0
case "isbest"
	doWealthL=0:douserEPL=0:douserCPL=0
	doWealthH=forum_user(15):douserEPH=Forum_user(17):douserCPH=Forum_user(16)
case "����"
	doWealthL=0:douserEPL=0:douserCPL=0
	doWealthH=forum_user(15):douserEPH=Forum_user(17):douserCPH=Forum_user(16)
case "nobest"
	doWealthL=-forum_user(15):douserEPL=-Forum_user(17):douserCPL=-Forum_user(16)
	doWealthH=0:douserEPH=0:douserCPH=0
case "�������"
	doWealthL=-forum_user(15):douserEPL=-Forum_user(17):douserCPL=-Forum_user(16)
	doWealthH=0:douserEPH=0:douserCPH=0
case "award"
	doWealthL=0:douserEPL=0:douserCPL=0
	doWealthH=50:douserEPH=50:douserCPH=50
case "����"
	doWealthL=0:douserEPL=0:douserCPL=0
	doWealthH=50:douserEPH=50:douserCPH=50
case "punish"
	doWealthL=-50:douserEPL=-50:douserCPL=-50
	doWealthH=0:douserEPH=0:douserCPH=0
case "�ͷ�"
	doWealthL=-50:douserEPL=-50:douserCPL=-50
	doWealthH=0:douserEPH=0:douserCPH=0
case "ispb"
	doWealthL=-Forum_user(15):douserEPL=-Forum_user(17):douserCPL=-Forum_user(16)
	doWealthH=0:douserEPH=0:douserCPH=0
case "��������"
	doWealthL=-Forum_user(15):douserEPL=-Forum_user(17):douserCPL=-Forum_user(16)
	doWealthH=0:douserEPH=0:douserCPH=0
case "nopb"
	doWealthL=0:douserEPL=0:douserCPL=0
	doWealthH=Forum_user(15):douserEPH=Forum_user(17):douserCPH=Forum_user(16)
case "�����������"
	doWealthL=0:douserEPL=0:douserCPL=0
	doWealthH=Forum_user(15):douserEPH=Forum_user(17):douserCPH=Forum_user(16)
end select
if not isnumeric(request("doWealth")) or request("doWealth")="0" or request("doWealth")="" then
	doWealth=0
	doWealthMsg=""
else
	if doWealth>doWealthH or doWealth<doWealthL then
		doWealth=0
		doWealthMsg=""
	else
		doWealth=request("doWealth")
		doWealthMsg="��Ǯ" & request("doWealth") & "��"
	end if
end if
if not isnumeric(request("douserEP")) or request("douserEP")="0" or request("douserEP")="" then
	douserEP=0
	douserEPMsg=""
else
	if douserEP>douserEPH or douserEP<douserEPL then
		douserEP=0
		douserEPMsg=""
	else
		douserEP=request("douserEP")
		douserEPMsg="����" & request("douserEP") & "��"
	end if
end if
if not isnumeric(request("douserCP")) or request("douserCP")="0" or request("douserCP")="" then
	douserCP=0
	douserCPMsg=""
else
	if douserCP>douserCPH or douserCP<douserCPL then
		douserCP=0
		douserCPMsg=""
	else
		douserCP=request("douserCP")
		douserCPMsg="����" & request("douserCP")
	end if
end if
if doWealthMsg="" and douserEPMsg="" and douserCPMsg="" then
	allmsg="û�ж��û����з�ֵ����"
else
	allmsg="�û�������" & doWealthMsg & douserEPMsg & douserCPMsg
end if
if not founderr then
	set rs=conn.execute("select title,postusername,postuserid,PostTable from topic where topicid="&id)
	if rs.eof and rs.bof then
		Errmsg=Errmsg+"<br><li>û���ҵ�������ӣ�"
		founderr=true
	else
		Topic=CheckStr(rs(0))
		Topicusername=CheckStr(rs(1))
		Topicuserid=rs(2)
		TotalUseTable=rs(3)
	end if
	'�ж��û��Ƿ�������/�������Ȩ��
	if (master or superboardmaster or boardmaster) and Cint(GroupSetting(20))=1 then
		canlocktopic=true
	else
		canlocktopic=false
	end if
	if Cint(GroupSetting(20))=1 and UserGroupID>3 then
		canlocktopic=true
	end if
	if (Cint(GroupSetting(13))=1 and TopicUsername=membername) then
		canlocktopic=true
	end if
	if FoundUserPer and Cint(GroupSetting(13))=1 and TopicUsername=membername then
		canlocktopic=true
	elseif FoundUserPer and Cint(GroupSetting(13))=0 and TopicUsername=membername then
		canlocktopic=false
	end if
	if FoundUserPer and Cint(GroupSetting(20))=1 and TopicUsername<>membername then
		canlocktopic=true
	elseif FoundUserPer and Cint(GroupSetting(20))=0 and TopicUsername<>membername then
		canlocktopic=false
	end if
	'�ж��û��Ƿ����ƶ�����Ȩ��
	if (master or superboardmaster or boardmaster) and Cint(GroupSetting(19))=1 then
		canmovetopic=true
	else
		canmovetopic=false
	end if
	if Cint(GroupSetting(19))=1 and UserGroupID>3 then
		canmovetopic=true
	end if
	if (Cint(GroupSetting(12))=1 and TopicUsername=membername) then
		canmovetopic=true
	end if
	if FoundUserPer and Cint(GroupSetting(12))=1 and TopicUsername=membername then
		canmovetopic=true
	elseif FoundUserPer and Cint(GroupSetting(12))=0 and TopicUsername=membername then
		canmovetopic=false
	end if
	if FoundUserPer and Cint(GroupSetting(19))=1 and TopicUsername<>membername then
		canmovetopic=true
	elseif FoundUserPer and Cint(GroupSetting(19))=0 and TopicUsername<>membername then
		canmovetopic=false
	end if
	'�ж��û��Ƿ���ɾ������Ȩ��
	if (master or superboardmaster or boardmaster) and Cint(GroupSetting(18))=1 then
		candeltopic=true
	else
		candeltopic=false
	end if
	if Cint(GroupSetting(18))=1 and UserGroupID>3 then
		candeltopic=true
	end if
	if (Cint(GroupSetting(11))=1 and TopicUsername=membername) then
		candeltopic=true
	end if
	if FoundUserPer and Cint(GroupSetting(11))=1 and TopicUsername=membername then
		candeltopic=true
	elseif FoundUserPer and Cint(GroupSetting(11))=0 and TopicUsername=membername then
		candeltopic=false
	end if
	if FoundUserPer and Cint(GroupSetting(18))=1 and TopicUsername<>membername then
		candeltopic=true
	elseif FoundUserPer and Cint(GroupSetting(18))=0 and TopicUsername<>membername then
		candeltopic=false
	end if
	'�ж��û��Ƿ��й̶�/����̶�����Ȩ��
	if (master or superboardmaster or boardmaster) and Cint(GroupSetting(21))=1 then
		cantoptopic=true
	else
		cantoptopic=false
	end if
	if Cint(GroupSetting(21))=1 and UserGroupID>3 then
		cantoptopic=true
	end if
	if FoundUserPer and Cint(GroupSetting(21))=1 then
		cantoptopic=true
	elseif FoundUserPer and Cint(GroupSetting(21))=0 then
		cantoptopic=false
	end if
	'�ж��û��Ƿ����̶ܹ�����Ȩ��
	if (master or superboardmaster or boardmaster) and Cint(GroupSetting(38))=1 then
		cantoptopic_a=true
	else
		cantoptopic_a=false
	end if
	if Cint(GroupSetting(38))=1 and UserGroupID>3 then
		cantoptopic_a=true
	end if
	if FoundUserPer and Cint(GroupSetting(38))=1 then
		cantoptopic_a=true
	elseif FoundUserPer and Cint(GroupSetting(38))=0 then
		cantoptopic_a=false
	end if
	'�ж��û��Ƿ��н���/�ͷ�����Ȩ��
	if (master or superboardmaster or boardmaster) and Cint(GroupSetting(22))=1 then
		canawardtopic=true
	else
		canawardtopic=false
	end if
	if Cint(GroupSetting(22))=1 and UserGroupID>3 then
		canawardtopic=true
	end if
	if FoundUserPer and Cint(GroupSetting(22))=1 then
		canawardtopic=true
	elseif FoundUserPer and Cint(GroupSetting(22))=0 then
		canawardtopic=false
	end if
	'�ж��û��Ƿ��м���/�����������Ȩ��
	if (master or superboardmaster or boardmaster) and Cint(GroupSetting(24))=1 then
		canbesttopic=true
		canpbtopic=true
	else
		canbesttopic=false
		canpbtopic=false
	end if
	if Cint(GroupSetting(24))=1 and UserGroupID>3 then
		canbesttopic=true
		canpbtopic=true
	end if
	if FoundUserPer and Cint(GroupSetting(24))=1 then
		canbesttopic=true
		canpbtopic=true
	elseif FoundUserPer and Cint(GroupSetting(24))=0 then
		canbesttopic=false
		canpbtopic=false
	end if
end if
if founderr then
	call nav()
	call head_var(2,0,"","")
	call dvbbs_error()
else
	call nav()
	call head_var(1,BoardDepth,0,0)
	select case request("action")
	case "lock"
		if not canlocktopic then
			Errmsg=Errmsg+"<br><li>��û��ִ�д˲�����Ȩ�ޡ�"
			founderr=true
		else
			call lock()
		end if
	case "unlock"
		if not canlocktopic then
			Errmsg=Errmsg+"<br><li>��û��ִ�д˲�����Ȩ�ޡ�"
			founderr=true
		else
			call unlock()
		end if
	case "uptopic"
		if not canlocktopic then
			Errmsg=Errmsg+"<br><li>��û��ִ�д˲�����Ȩ�ޡ�"
			founderr=true
		else
			call uptopic()
		end if
	case "downtopic"
		if not canlocktopic then
			Errmsg=Errmsg+"<br><li>��û��ִ�д˲�����Ȩ�ޡ�"
			founderr=true
		else
			call downtopic()
		end if
	case "delet"
		if not candeltopic then
			Errmsg=Errmsg+"<br><li>��û��ִ�д˲�����Ȩ�ޡ�"
			founderr=true
		else
			call delete()
		end if
	case "move"
		if not canmovetopic then
			Errmsg=Errmsg+"<br><li>��û��ִ�д˲�����Ȩ�ޡ�"
			founderr=true
		else
			call Tmove()
		end if
	case "copy"
		call copy()
	case "istop"
		if not cantoptopic then
			Errmsg=Errmsg+"<br><li>��û��ִ�д˲�����Ȩ�ޡ�"
			founderr=true
		else
			call istop()
		end if
	case "notop"
		if not cantoptopic then
			Errmsg=Errmsg+"<br><li>��û��ִ�д˲�����Ȩ�ޡ�"
			founderr=true
		else
			call notop()
		end if
	case "istop_a"
		if not cantoptopic_a then
			Errmsg=Errmsg+"<br><li>��û��ִ�д˲�����Ȩ�ޡ�"
			founderr=true
		else
			call istop_a()
		end if
	case "notop_a"
		if not cantoptopic_a then
			Errmsg=Errmsg+"<br><li>��û��ִ�д˲�����Ȩ�ޡ�"
			founderr=true
		else
			call notop_a()
		end if
	case "dele"
		call dele()
	case "isbest"
		if not canbesttopic then
			Errmsg=Errmsg+"<br><li>��û��ִ�д˲�����Ȩ�ޡ�"
			founderr=true
		else
			call isbest()
		end if
	case "nobest"
		if not canbesttopic then
			Errmsg=Errmsg+"<br><li>��û��ִ�д˲�����Ȩ�ޡ�"
			founderr=true
		else
			call nobest()
		end if
	case "award"
		if not canawardtopic then
			Errmsg=Errmsg+"<br><li>��û��ִ�д˲�����Ȩ�ޡ�"
			founderr=true
		else
			call award()
		end if
	case "punish"
		if not canawardtopic then
			Errmsg=Errmsg+"<br><li>��û��ִ�д˲�����Ȩ�ޡ�"
			founderr=true
		else
			call punish()
		end if
	case "ispb"
		if not canpbtopic then
			Errmsg=Errmsg+"<br><li>��û��ִ�д˲�����Ȩ�ޡ�"
			founderr=true
		else
			call ispb()
		end if
	case "nopb"
		if not canpbtopic then
			Errmsg=Errmsg+"<br><li>��û��ִ�д˲�����Ȩ�ޡ�"
			founderr=true
		else
			call nopb()
		end if
	case else
		call main()
	end select
	if founderr then call dvbbs_error()
end if
call activeonline()
call footer()
	
sub main()
Dim doWealth,douserEP,douserCP
Dim seldisable,reaction
Dim postusername
select case request("action")
case "����"
	doWealth=0
	douserEP=0
	douserCP=0
	if canawardtopic then
	seldisable=""
	else
	seldisable="disabled"
	end if
	reaction="lock"
	if not canlocktopic then
		Errmsg=Errmsg+"<br><li>��û��ִ�д˲�����Ȩ�ޡ�"
		founderr=true
		exit sub
	end if
case "����"
	doWealth=0
	douserEP=0
	douserCP=0
	if canawardtopic then
	seldisable=""
	else
	seldisable="disabled"
	end if
	reaction="unlock"
	if not canlocktopic then
		Errmsg=Errmsg+"<br><li>��û��ִ�д˲�����Ȩ�ޡ�"
		founderr=true
		exit sub
	end if
case "����"
	doWealth=0
	douserEP=0
	douserCP=0
	if canawardtopic then
	seldisable=""
	else
	seldisable="disabled"
	end if
	reaction="uptopic"
	if not canlocktopic then
		Errmsg=Errmsg+"<br><li>��û��ִ�д˲�����Ȩ�ޡ�"
		founderr=true
		exit sub
	end if
case "�³�"
	doWealth=0
	douserEP=0
	douserCP=0
	if canawardtopic then
	seldisable=""
	else
	seldisable="disabled"
	end if
	reaction="downtopic"
	if not canlocktopic then
		Errmsg=Errmsg+"<br><li>��û��ִ�д˲�����Ȩ�ޡ�"
		founderr=true
		exit sub
	end if
case "ɾ������"
	doWealth=-Forum_user(3)
	douserEP=-Forum_user(8)
	douserCP=-Forum_user(13)
	if canawardtopic then
	seldisable=""
	else
	seldisable="disabled"
	end if
	reaction="delet"
	if not candeltopic then
		Errmsg=Errmsg+"<br><li>��û��ִ�д˲�����Ȩ�ޡ�"
		founderr=true
		exit sub
	end if
case "ɾ������"
	doWealth=-Forum_user(3)
	douserEP=-Forum_user(8)
	douserCP=-Forum_user(13)
	if canawardtopic then
	seldisable=""
	else
	seldisable="disabled"
	end if
	reaction="dele"
	set rs=conn.execute("select topic,username,postuserid from "&TotalUseTable&" where Announceid="&replyid)
	if rs.eof and rs.bof then
		Errmsg=Errmsg+"<br><li>û���ҵ���ظ�����"
		founderr=true
		exit sub
	end if
	Topic=CheckStr(rs(0))
	TopicUsername=rs(1)
	TopicUserID=rs(2)
	'�ж��û��Ƿ���ɾ������Ȩ��
	if (Cint(GroupSetting(11))=1 and TopicUsername=membername) then
		candeltopic=true
	end if
	if FoundUserPer and Cint(GroupSetting(11))=1 and TopicUsername=membername then
		candeltopic=true
	elseif FoundUserPer and Cint(GroupSetting(11))=0 and TopicUsername=membername then
		candeltopic=false
	end if
	if FoundUserPer and Cint(GroupSetting(18))=1 and TopicUsername<>membername then
		candeltopic=true
	elseif FoundUserPer and Cint(GroupSetting(18))=0 and TopicUsername<>membername then
		candeltopic=false
	end if
	if not candeltopic then
		Errmsg=Errmsg+"<br><li>��û��ִ�д˲�����Ȩ�ޡ�"
		founderr=true
		exit sub
	end if
case "����"
	doWealth=Forum_user(15)
	douserEP=Forum_user(17)
	douserCP=Forum_user(16)
	if canawardtopic then
	seldisable=""
	else
	seldisable="disabled"
	end if
	reaction="isbest"
	set rs=conn.execute("select topic,username,postuserid from "&TotalUseTable&" where Announceid="&replyid)
	if rs.eof and rs.bof then
		Errmsg=Errmsg+"<br><li>û���ҵ���ظ�����"
		founderr=true
		exit sub
	end if
	Topic=CheckStr(rs(0))
	TopicUsername=rs(1)
	TopicUserID=rs(2)
	if not canbesttopic then
		Errmsg=Errmsg+"<br><li>��û��ִ�д˲�����Ȩ�ޡ�"
		founderr=true
		exit sub
	end if
case "�������"
	doWealth=-Forum_user(15)
	douserEP=-Forum_user(17)
	douserCP=-Forum_user(16)
	if canawardtopic then
	seldisable=""
	else
	seldisable="disabled"
	end if
	reaction="nobest"
	set rs=conn.execute("select topic,username,postuserid from "&TotalUseTable&" where Announceid="&replyid)
	if rs.eof and rs.bof then
		Errmsg=Errmsg+"<br><li>û���ҵ���ظ�����"
		founderr=true
		exit sub
	end if
	Topic=CheckStr(rs(0))
	TopicUsername=rs(1)
	TopicUserID=rs(2)
	if not canbesttopic then
		Errmsg=Errmsg+"<br><li>��û��ִ�д˲�����Ȩ�ޡ�"
		founderr=true
		exit sub
	end if
case "����"
	doWealth=0
	douserEP=0
	douserCP=0
	seldisable="disabled"
	reaction="copy"
	set rs=conn.execute("select topic,username,postuserid from "&TotalUseTable&" where Announceid="&replyid)
	if rs.eof and rs.bof then
		Errmsg=Errmsg+"<br><li>û���ҵ���ظ�����"
		founderr=true
		exit sub
	end if
	Topic=CheckStr(rs(0))
	TopicUsername=rs(1)
	TopicUserID=rs(2)
	'�ж��û��Ƿ����ƶ�����Ȩ��
	if (Cint(GroupSetting(12))=1 and TopicUsername=membername) then
		canmovetopic=true
	end if
	if FoundUserPer and Cint(GroupSetting(12))=1 and TopicUsername=membername then
		canmovetopic=true
	elseif FoundUserPer and Cint(GroupSetting(12))=0 and TopicUsername=membername then
		canmovetopic=false
	end if
	if FoundUserPer and Cint(GroupSetting(19))=1 and TopicUsername<>membername then
		canmovetopic=true
	elseif FoundUserPer and Cint(GroupSetting(19))=0 and TopicUsername<>membername then
		canmovetopic=false
	end if
	if not canmovetopic then
		Errmsg=Errmsg+"<br><li>��û��ִ�д˲�����Ȩ�ޡ�"
		founderr=true
		exit sub
	end if
case "�̶�"
	doWealth=0
	douserEP=0
	douserCP=0
	if canawardtopic then
	seldisable=""
	else
	seldisable="disabled"
	end if
	reaction="istop"
	if not cantoptopic then
		Errmsg=Errmsg+"<br><li>��û��ִ�д˲�����Ȩ�ޡ�"
		founderr=true
		exit sub
	end if
case "���"
	doWealth=0
	douserEP=0
	douserCP=0
	if canawardtopic then
	seldisable=""
	else
	seldisable="disabled"
	end if
	reaction="notop"
	if not cantoptopic then
		Errmsg=Errmsg+"<br><li>��û��ִ�д˲�����Ȩ�ޡ�"
		founderr=true
		exit sub
	end if
case "�̶ܹ�"
	doWealth=0
	douserEP=0
	douserCP=0
	if canawardtopic then
	seldisable=""
	else
	seldisable="disabled"
	end if
	reaction="istop_a"
	if not cantoptopic_a then
		Errmsg=Errmsg+"<br><li>��û��ִ�д˲�����Ȩ�ޡ�"
		founderr=true
		exit sub
	end if
case "����̶ܹ�"
	doWealth=0
	douserEP=0
	douserCP=0
	if canawardtopic then
	seldisable=""
	else
	seldisable="disabled"
	end if
	reaction="notop_a"
	if not cantoptopic_a then
		Errmsg=Errmsg+"<br><li>��û��ִ�д˲�����Ȩ�ޡ�"
		founderr=true
		exit sub
	end if
case "�ƶ�"
	doWealth=0
	douserEP=0
	douserCP=0
	seldisable="disabled"
	reaction="move"
	if not canmovetopic then
		Errmsg=Errmsg+"<br><li>��û��ִ�д˲�����Ȩ�ޡ�"
		founderr=true
		exit sub
	end if
case "����"
	doWealth=5
	douserEP=1
	douserCP=2
	seldisable=""
	reaction="award"
	if not canawardtopic then
		Errmsg=Errmsg+"<br><li>��û��ִ�д˲�����Ȩ�ޡ�"
		founderr=true
		exit sub
	end if
case "�ͷ�"
	doWealth=-5
	douserEP=-1
	douserCP=-2
	seldisable=""
	reaction="punish"
	if not canawardtopic then
		Errmsg=Errmsg+"<br><li>��û��ִ�д˲�����Ȩ�ޡ�"
		founderr=true
		exit sub
	end if
case "��������"
	doWealth=-Forum_user(15)
	douserEP=-Forum_user(17)
	douserCP=-Forum_user(16)
	seldisable=""
	reaction="ispb"
	set rs=conn.execute("select topic,username,postuserid from "&TotalUseTable&" where Announceid="&replyid)
	if rs.eof and rs.bof then
		Errmsg=Errmsg+"<br><li>û���ҵ���ظ�����"
		founderr=true
		exit sub
	end if
	Topic=CheckStr(rs(0))
	TopicUsername=rs(1)
	TopicUserID=rs(2)
	if not canpbtopic then
		Errmsg=Errmsg+"<br><li>��û��ִ�д˲�����Ȩ�ޡ�"
		founderr=true
		exit sub
	end if
case "�����������"
	doWealth=Forum_user(15)
	douserEP=Forum_user(17)
	douserCP=Forum_user(16)
	seldisable=""
	reaction="nopb"
	set rs=conn.execute("select topic,username,postuserid from "&TotalUseTable&" where Announceid="&replyid)
	if rs.eof and rs.bof then
		Errmsg=Errmsg+"<br><li>û���ҵ���ظ�����"
		founderr=true
		exit sub
	end if
	Topic=CheckStr(rs(0))
	TopicUsername=rs(1)
	TopicUserID=rs(2)
	if not canpbtopic then
		Errmsg=Errmsg+"<br><li>��û��ִ�д˲�����Ȩ�ޡ�"
		founderr=true
		exit sub
	end if
case else
	Errmsg=Errmsg+"<br><li>��ָ����ز�����"
	founderr=true
	exit sub
end select
%>
<FORM METHOD=POST ACTION="admin_postings.asp?action=<%=reaction%>" name="frmAnnounce">
<table cellspacing="1" cellpadding="3" align="center" class=tableborder1>
  <tr> 
    <th height=24>��̳���ӹ������ģ�����Ҫ���еĲ�����<%=request("action")%></th>
  </tr>   
  <tr> 
    <td class=tablebody1 height=24><b>
      ��������</b>��  
	  <select name="title" size=1>
<option value="">�Զ���</option>
<option value="��ˮ">��ˮ</option>
<option value="���">���</option>
<option value="����">����</option>
<option value="��ʱ">��ʱ</option>
<option value="������">������</option>
<option value="���ݲ���">���ݲ���</option>
<option value="�ظ�����">�ظ�����</option>
<option value="����">����</option>
	  </select>
	  <input type="text" name="content" size=60>  *</td>
  </tr>   
  <tr> 
    <td class=tablebody1 height=24><b>
      �û�����</b>��  ��Ǯ
	<select name="doWealth" size=1 <%=seldisable%>>

<%for i=doWealthL to doWealthH%>
<option value="<%=i%>" <%if cint(doWealth)=cint(i) then%>selected<%end if%>><%=i%></option>
<%next%>
	</select>&nbsp;����
	<select name="douserCP" size=1 <%=seldisable%>>

<%for i=douserCPL to douserCPH%>
<option value="<%=i%>" <%if cint(dousercp)=cint(i) then%>selected<%end if%>><%=i%></option>
<%next%>
	</select>&nbsp;����
	<select name="douserEP" size=1 <%=seldisable%>>

<%for i=douserEPL to douserEPH%>
<option value="<%=i%>" <%if cint(douserep)=cint(i) then%>selected<%end if%>><%=i%></option>
<%next%>
	</select>&nbsp;<input type="checkbox" name="checkbox" value="checkbox" onclick="nocent()" <%=seldisable%>>�����û����з�ֵ����
  </td>
  </tr> 
<input type=hidden value="<%=id%>" name="id">
<input type=hidden value="<%=replyid%>" name="replyid">
<input type=hidden value="<%=boardid%>" name="boardid">
  <tr> 
    <td height=24 class=tablebody1>
<B>����֪ͨ</B>��<input type=text size=70 name="msg">&nbsp;<input type="checkbox" name="ismsg" value="1">ʹ��</td>
  </tr>   
  <tr> 
    <td height=24 class=tablebody1>
<B>����˵��</B>��<BR>
��������ʹ�ð����Ĺ���ְ�ܣ��������в���������¼<BR>
�����ѡ���˶���֪ͨ����Ὣɾ��ԭ���͸��û�����Ҳ�����ڶ����и��ϼ��˵��</td>
  </tr>   
  <tr> 
    <td height=24 class=tablebody2 align=center>
      <input type="submit" name=submit value="ȷ�ϲ���"></td>
  </tr>   
</table>
</FORM>
<script>
function nocent()
{
if(document.frmAnnounce.doWealth.disabled==true){
document.frmAnnounce.doWealth.disabled=false;
document.frmAnnounce.douserCP.disabled=false;
document.frmAnnounce.douserEP.disabled=false;
}else{
document.frmAnnounce.doWealth.disabled=true;
document.frmAnnounce.douserCP.disabled=true;
document.frmAnnounce.douserEP.disabled=true;
}
}
</script>
<%
	set rs=nothing
end sub

sub ispb()
dim datetimestr
set rs=conn.execute("select * from "&TotalUseTable&" where Announceid="&replyid)
if rs.eof and rs.bof then
	Errmsg=Errmsg+"<br><li>û���ҵ�������ӣ�"
	founderr=true
	exit sub
end if
topic=rs("topic")
topicusername=rs("username")
topicuserid=rs("postuserid")
if topic="" then
	topic=left(replace(replace(rs("body"),"''''",""),chr(10),","),26)
else
	topic=replace(topic,"''''","")
end if
datetimestr=replace(replace(rs("dateandtime"),"����",""),"����","")
title=CheckStr(htmlencode(request.form("title")))
content=CheckStr(htmlencode(request.form("content")))
content="ԭ��" & title & content
if htmlencode(request.form("title"))="" and htmlencode(request.form("content"))="" then
	Errmsg=Errmsg+"<br><li>��д������ԭ��"
	founderr=true
	exit sub
end if
sql="update "&TotalUseTable&" set isbest=2 where boardid="&boardid&" and announceid="&cstr(replyid)
conn.Execute(sql)
sql="insert into log (l_announceid,l_boardid,l_touser,l_username,l_content,l_ip) values ("&id&","&boardid&",'"&topicusername&"','"&membername&"','�������Ρ�"&topic&"����"&content& "��"&allmsg&"','"&ip&"')"
conn.execute(sql)
conn.execute("update [user] set userWealth=userWealth+"&doWealth&",userCP=userCP+"&douserCP&",userEP=userEP+"&douserEP&",userIsBest=userisBest+1 where userid="&topicuserid)
sucmsg="�������Ρ�"&topic&"����"&content& "��"&allmsg&""
call dvbbs_suc()
set rs=nothing
end sub

sub nopb()
set rs=conn.execute("select topic,username,postuserid from "&TotalUseTable&" where Announceid="&replyid)
if rs.eof and rs.bof then
	Errmsg=Errmsg+"<BR><li>û���ҵ�������ӣ�"
	founderr=true
	exit sub
else
	Topic=CheckStr(rs(0))
	topicusername=rs(1)
	topicuserid=rs(2)
	if topic="" then
		topic="������Ϊ�ظ�����"
	end if
end if
set rs=nothing
title=CheckStr(htmlencode(request.form("title")))
content=CheckStr(htmlencode(request.form("content")))
content="ԭ��" & title & content
if htmlencode(request.form("title"))="" and htmlencode(request.form("content"))="" then
	Errmsg=Errmsg+"<BR><li>��д������ԭ��"
	founderr=true
	exit sub
end if
sql="update "&TotalUseTable&" set isbest=0 where boardid="&boardid&" and announceid="&replyid
conn.Execute(sql)
sql="insert into log (l_announceid,l_boardid,l_touser,l_username,l_content,l_ip) values ("&id&","&boardid&",'"&topicusername&"','"&membername&"','����������Ρ�"&topic&"����"&content& "��"&allmsg&"','"&ip&"')"
conn.execute(sql)
conn.execute("update [user] set userWealth=userWealth+"&doWealth&",userCP=userCP+"&douserCP&",userEP=userEP+"&douserEP&",userIsBest=userisBest-1 where userid="&topicuserid)
sucmsg="����������Ρ�"&topic&"����"&content& "��"&allmsg&""
call dvbbs_suc()
end sub

sub award()
title=CheckStr(htmlencode(request.form("title")))
content=CheckStr(htmlencode(request.form("content")))
content="ԭ��" & title & content
if TopicUserName=membername then
	Errmsg=Errmsg+"<br><li>�����͹���Ա���ܶ��Լ�������ʵ�н���������"
	founderr=true
	exit sub
end if
if htmlencode(request.form("title"))="" and htmlencode(request.form("content"))="" then
	Errmsg=Errmsg+"<br><li>��д������ԭ��"
	founderr=true
	exit sub
end if
sql="insert into log (l_announceid,l_boardid,l_touser,l_username,l_content,l_ip) values ("&id&","&boardid&",'"&TopicUsername&"','"&membername&"','�����û���"&topic&"����"&content& "��"&allmsg&"','"&ip&"')"
conn.execute(sql)
if allmsg<>"" then
	conn.execute("update [user] set userWealth=userWealth+"&doWealth&",userCP=userCP+"&douserCP&",userEP=userEP+"&douserEP&" where userid="&TopicUserID)
end if
sucmsg="�����û���"&topic&"����"&content& "��"&allmsg&""
if request("ismsg")="1" then
	msgcontent="����������ӡ�[url=dispbbs.asp?boardid="&boardid&"&ID="&ID&"]"&topic&"[/url]����"&replace(Content,"ԭ��","")&"�õ�"&replace(allmsg,"�û�������","")&"�Ľ���"
	if request("msg")<>"" then
	msgContent=msgContent & chr(10) & "����Ϊ�����߸����ĸ��ԣ�" & request("msg")
	end if
	conn.Execute("insert into message(incept,sender,title,content,sendtime,flag,issend) values('"&TopicUsername&"','"&membername&"','ϵͳ��Ϣ','"&checkSTR(msgContent)&"',Now(),0,1)")
end if
call dvbbs_suc()
end sub

sub punish()
title=CheckStr(htmlencode(request.form("title")))
content=CheckStr(htmlencode(request.form("content")))
content="ԭ��" & title & content
if htmlencode(request.form("title"))="" and htmlencode(request.form("content"))="" then
	Errmsg=Errmsg+"<br><li>��д������ԭ��"
	founderr=true
	exit sub
end if
sql="insert into log (l_announceid,l_boardid,l_touser,l_username,l_content,l_ip) values ("&id&","&boardid&",'"&TopicUsername&"','"&membername&"','�ͷ��û���"&topic&"����"&content& "��"&allmsg&"','"&ip&"')"
conn.execute(sql)
if allmsg<>"" then
	conn.execute("update [user] set userWealth=userWealth+"&doWealth&",userCP=userCP+"&douserCP&",userEP=userEP+"&douserEP&" where userid="&TopicUserID)
end if
sucmsg="�ͷ��û���"&topic&"����"&content& "��"&allmsg&""
if request("ismsg")="1" then
	msgcontent="����������ӡ�[url=dispbbs.asp?boardid="&boardid&"&ID="&ID&"]"&topic&"[/url]����"&replace(Content,"ԭ��","")&"�õ�"&replace(allmsg,"�û�������","")&"�ĳͷ�"
	if request("msg")<>"" then
	msgContent=msgContent & chr(10) & "����Ϊ�����߸����ĸ��ԣ�" & request("msg")
	end if
	conn.Execute("insert into message(incept,sender,title,content,sendtime,flag,issend) values('"&TopicUsername&"','"&membername&"','ϵͳ��Ϣ','"&checkSTR(msgContent)&"',Now(),0,1)")
end if
call dvbbs_suc()
end sub

sub lock()
title=CheckStr(htmlencode(request.form("title")))
content=CheckStr(htmlencode(request.form("content")))
content="ԭ��" & title & content
if htmlencode(request.form("title"))="" and htmlencode(request.form("content"))="" then
	Errmsg=Errmsg+"<br><li>��д������ԭ��"
	founderr=true
	exit sub
end if
sql="update topic set locktopic=1 where boardid="&boardid&" and topicid="&id
conn.Execute(sql)
sql="insert into log (l_announceid,l_boardid,l_touser,l_username,l_content,l_ip) values ("&id&","&boardid&",'"&TopicUsername&"','"&membername&"','�������ӡ�"&topic&"����"&content& "��"&allmsg&"','"&ip&"')"
conn.execute(sql)
if allmsg<>"" then
	conn.execute("update [user] set userWealth=userWealth+"&doWealth&",userCP=userCP+"&douserCP&",userEP=userEP+"&douserEP&" where userid="&TopicUserID)
end if
sucmsg="�������ӡ�"&topic&"����"&content& "��"&allmsg&""
if request("ismsg")="1" then
	msgcontent="����������ӡ�[url=dispbbs.asp?boardid="&boardid&"&ID="&ID&"]"&topic&"[/url]����"&replace(Content,"ԭ��","")&"���������ҽ�����"&replace(allmsg,"�û�������","")&"�Ĳ���"
	if request("msg")<>"" then
	msgContent=msgContent & chr(10) & "����Ϊ�����߸����ĸ��ԣ�" & request("msg")
	end if
	conn.Execute("insert into message(incept,sender,title,content,sendtime,flag,issend) values('"&TopicUsername&"','"&membername&"','ϵͳ��Ϣ','"&checkSTR(msgContent)&"',Now(),0,1)")
end if
call dvbbs_suc()
end sub

sub unlock()
title=CheckStr(htmlencode(request.form("title")))
content=CheckStr(htmlencode(request.form("content")))
content="ԭ��" & title & content
if htmlencode(request.form("title"))="" and htmlencode(request.form("content"))="" then
	Errmsg=Errmsg+"<br><li>��д������ԭ��"
	founderr=true
	exit sub
end if
sql="update topic set locktopic=0 where boardid="&boardid&" and topicid="&id
conn.Execute(sql)

sql="insert into log (l_announceid,l_boardid,l_touser,l_username,l_content,l_ip) values ("&id&","&boardid&",'"&TopicUsername&"','"&membername&"','���������"&topic&"����"&content& "��"&allmsg&"','"&ip&"')"
conn.execute(sql)
if allmsg<>"" then
	conn.execute("update [user] set userWealth=userWealth+"&doWealth&",userCP=userCP+"&douserCP&",userEP=userEP+"&douserEP&" where userid="&TopicUserID)
end if
sucmsg="���������"&topic&"����"&content& "��"&allmsg&""
if request("ismsg")="1" then
	msgcontent="����������ӡ�[url=dispbbs.asp?boardid="&boardid&"&ID="&ID&"]"&topic&"[/url]����"&replace(Content,"ԭ��","")&"������������ҽ�����"&replace(allmsg,"�û�������","")&"�Ĳ���"
	if request("msg")<>"" then
	msgContent=msgContent & chr(10) & "����Ϊ�����߸����ĸ��ԣ�" & request("msg")
	end if
	conn.Execute("insert into message(incept,sender,title,content,sendtime,flag,issend) values('"&TopicUsername&"','"&membername&"','ϵͳ��Ϣ','"&checkSTR(msgContent)&"',Now(),0,1)")
end if
call dvbbs_suc()
end sub

sub uptopic()
title=CheckStr(htmlencode(request.form("title")))
content=CheckStr(htmlencode(request.form("content")))
content="ԭ��" & title & content
if htmlencode(request.form("title"))="" and htmlencode(request.form("content"))="" then
	Errmsg=Errmsg+"<br><li>��д������ԭ��"
	founderr=true
	exit sub
end if
sql="update topic set LastPostTime=Now() where boardid="&boardid&" and topicid="&id
conn.Execute(sql)

sql="insert into log (l_announceid,l_boardid,l_touser,l_username,l_content,l_ip) values ("&id&","&boardid&",'"&TopicUsername&"','"&membername&"','�������⡶"&topic&"����"&content& "��"&allmsg&"','"&ip&"')"
conn.execute(sql)
if allmsg<>"" then
	conn.execute("update [user] set userWealth=userWealth+"&doWealth&",userCP=userCP+"&douserCP&",userEP=userEP+"&douserEP&" where userid="&TopicUserID)
end if
sucmsg="�������⡶"&topic&"����"&content& "��"&allmsg&""
if request("ismsg")="1" then
	msgcontent="����������ӡ�[url=dispbbs.asp?boardid="&boardid&"&ID="&ID&"]"&topic&"[/url]����"&replace(Content,"ԭ��","")&"�����������ҽ�����"&replace(allmsg,"�û�������","")&"�Ĳ���"
	if request("msg")<>"" then
	msgContent=msgContent & chr(10) & "����Ϊ�����߸����ĸ��ԣ�" & request("msg")
	end if
	conn.Execute("insert into message(incept,sender,title,content,sendtime,flag,issend) values('"&TopicUsername&"','"&membername&"','ϵͳ��Ϣ','"&checkSTR(msgContent)&"',Now(),0,1)")
end if
call dvbbs_suc()
end sub

sub downtopic()
title=CheckStr(htmlencode(request.form("title")))
content=CheckStr(htmlencode(request.form("content")))
content="ԭ��" & title & content
if htmlencode(request.form("title"))="" and htmlencode(request.form("content"))="" then
	Errmsg=Errmsg+"<br><li>��д������ԭ��"
	founderr=true
	exit sub
end if
sql="update topic set LastPostTime=Now()-7 where boardid="&boardid&" and topicid="&id
conn.Execute(sql)

sql="insert into log (l_announceid,l_boardid,l_touser,l_username,l_content,l_ip) values ("&id&","&boardid&",'"&TopicUsername&"','"&membername&"','�³����⡶"&topic&"����"&content& "��"&allmsg&"','"&ip&"')"
conn.execute(sql)
if allmsg<>"" then
	conn.execute("update [user] set userWealth=userWealth+"&doWealth&",userCP=userCP+"&douserCP&",userEP=userEP+"&douserEP&" where userid="&TopicUserID)
end if
sucmsg="�³����⡶"&topic&"����"&content& "��"&allmsg&""
if request("ismsg")="1" then
	msgcontent="����������ӡ�[url=dispbbs.asp?boardid="&boardid&"&ID="&ID&"]"&topic&"[/url]����"&replace(Content,"ԭ��","")&"�����³����ҽ�����"&replace(allmsg,"�û�������","")&"�Ĳ���"
	if request("msg")<>"" then
	msgContent=msgContent & chr(10) & "����Ϊ�����߸����ĸ��ԣ�" & request("msg")
	end if
	conn.Execute("insert into message(incept,sender,title,content,sendtime,flag,issend) values('"&TopicUsername&"','"&membername&"','ϵͳ��Ϣ','"&checkSTR(msgContent)&"',Now(),0,1)")
end if
call dvbbs_suc()
end sub

sub istop()
title=CheckStr(htmlencode(request.form("title")))
content=CheckStr(htmlencode(request.form("content")))
content="ԭ��" & title & content
if htmlencode(request.form("title"))="" and htmlencode(request.form("content"))="" then
	Errmsg=Errmsg+"<br><li>��д������ԭ��"
	founderr=true
	exit sub
end if
	
sql="update topic set istop=1 where boardid="&boardid&" and topicid="&id
conn.Execute(sql)

sql="insert into log (l_announceid,l_boardid,l_touser,l_username,l_content,l_ip) values ("&id&","&boardid&",'"&TopicUsername&"','"&membername&"','�̶����ӡ�"&topic&"����"&content& "��"&allmsg&"','"&ip&"')"
conn.execute(sql)
if allmsg<>"" then
	conn.execute("update [user] set userWealth=userWealth+"&doWealth&",userCP=userCP+"&douserCP&",userEP=userEP+"&douserEP&" where userid="&TopicUserID)
end if
sucmsg="�̶����ӡ�"&topic&"����"&content& "��"&allmsg&""
if request("ismsg")="1" then
	msgcontent="����������ӡ�[url=dispbbs.asp?boardid="&boardid&"&ID="&ID&"]"&topic&"[/url]����"&replace(Content,"ԭ��","")&"���̶����ҽ�����"&replace(allmsg,"�û�������","")&"�Ĳ���"
	if request("msg")<>"" then
	msgContent=msgContent & chr(10) & "����Ϊ�����߸����ĸ��ԣ�" & request("msg")
	end if
	conn.Execute("insert into message(incept,sender,title,content,sendtime,flag,issend) values('"&TopicUsername&"','"&membername&"','ϵͳ��Ϣ','"&checkSTR(msgContent)&"',Now(),0,1)")
end if
call dvbbs_suc()
end sub

sub notop()
title=CheckStr(htmlencode(request.form("title")))
content=CheckStr(htmlencode(request.form("content")))
content="ԭ��" & title & content
if htmlencode(request.form("title"))="" and htmlencode(request.form("content"))="" then
	Errmsg=Errmsg+"<br><li>��д������ԭ��"
	founderr=true
	exit sub
end if
sql="update topic set istop=0 where boardid="&boardid&" and topicid="&id
conn.Execute(sql)

sql="insert into log (l_announceid,l_boardid,l_touser,l_username,l_content,l_ip) values ("&id&","&boardid&",'"&topicusername&"','"&membername&"','����̶���"&topic&"����"&content& "��"&allmsg&"','"&ip&"')"
conn.execute(sql)
if allmsg<>"" then
	conn.execute("update [user] set userWealth=userWealth+"&doWealth&",userCP=userCP+"&douserCP&",userEP=userEP+"&douserEP&" where userid="&topicuserid)
end if
sucmsg="����̶���"&topic&"����"&content& "��"&allmsg&""
if request("ismsg")="1" then
	msgcontent="����������ӡ�[url=dispbbs.asp?boardid="&boardid&"&ID="&ID&"]"&topic&"[/url]����"&replace(Content,"ԭ��","")&"������̶����ҽ�����"&replace(allmsg,"�û�������","")&"�Ĳ���"
	if request("msg")<>"" then
	msgContent=msgContent & chr(10) & "����Ϊ�����߸����ĸ��ԣ�" & request("msg")
	end if
	conn.Execute("insert into message(incept,sender,title,content,sendtime,flag,issend) values('"&TopicUsername&"','"&membername&"','ϵͳ��Ϣ','"&checkSTR(msgContent)&"',Now(),0,1)")
end if
call dvbbs_suc()
end sub

sub istop_a()
title=CheckStr(htmlencode(request.form("title")))
content=CheckStr(htmlencode(request.form("content")))
content="ԭ��" & title & content
if htmlencode(request.form("title"))="" and htmlencode(request.form("content"))="" then
	Errmsg=Errmsg+"<br><li>��д������ԭ��"
	founderr=true
	exit sub
end if
sql="update topic set istop=2 where boardid="&boardid&" and topicid="&id
conn.Execute(sql)

sql="insert into log (l_announceid,l_boardid,l_touser,l_username,l_content,l_ip) values ("&id&","&boardid&",'"&TopicUsername&"','"&membername&"','�̶ܹ����ӡ�"&topic&"����"&content& "��"&allmsg&"','"&ip&"')"
conn.execute(sql)
if allmsg<>"" then
	conn.execute("update [user] set userWealth=userWealth+"&doWealth&",userCP=userCP+"&douserCP&",userEP=userEP+"&douserEP&" where userid="&TopicUserID)
end if
sucmsg="�̶ܹ����ӡ�"&topic&"����"&content& "��"&allmsg&""
if request("ismsg")="1" then
	msgcontent="����������ӡ�[url=dispbbs.asp?boardid="&boardid&"&ID="&ID&"]"&topic&"[/url]����"&replace(Content,"ԭ��","")&"���̶ܹ����ҽ�����"&replace(allmsg,"�û�������","")&"�Ĳ���"
	if request("msg")<>"" then
	msgContent=msgContent & chr(10) & "����Ϊ�����߸����ĸ��ԣ�" & request("msg")
	end if
	conn.Execute("insert into message(incept,sender,title,content,sendtime,flag,issend) values('"&TopicUsername&"','"&membername&"','ϵͳ��Ϣ','"&checkSTR(msgContent)&"',Now(),0,1)")
end if
'��������Ϣ������cache
call cache_top()
call dvbbs_suc()
end sub

sub notop_a()
title=CheckStr(htmlencode(request.form("title")))
content=CheckStr(htmlencode(request.form("content")))
content="ԭ��" & title & content
if htmlencode(request.form("title"))="" and htmlencode(request.form("content"))="" then
	Errmsg=Errmsg+"<br><li>��д������ԭ��"
	founderr=true
	exit sub
end if
sql="update topic set istop=0 where boardid="&boardid&" and topicid="&id
conn.Execute(sql)

sql="insert into log (l_announceid,l_boardid,l_touser,l_username,l_content,l_ip) values ("&id&","&boardid&",'"&topicusername&"','"&membername&"','����̶ܹ���"&topic&"����"&content& "��"&allmsg&"','"&ip&"')"
conn.execute(sql)
if allmsg<>"" then
	conn.execute("update [user] set userWealth=userWealth+"&doWealth&",userCP=userCP+"&douserCP&",userEP=userEP+"&douserEP&" where userid="&topicuserid)
end if
sucmsg="����̶ܹ���"&topic&"����"&content& "��"&allmsg&""
if request("ismsg")="1" then
	msgcontent="����������ӡ�[url=dispbbs.asp?boardid="&boardid&"&ID="&ID&"]"&topic&"[/url]����"&replace(Content,"ԭ��","")&"������̶ܹ����ҽ�����"&replace(allmsg,"�û�������","")&"�Ĳ���"
	if request("msg")<>"" then
	msgContent=msgContent & chr(10) & "����Ϊ�����߸����ĸ��ԣ�" & request("msg")
	end if
	conn.Execute("insert into message(incept,sender,title,content,sendtime,flag,issend) values('"&TopicUsername&"','"&membername&"','ϵͳ��Ϣ','"&checkSTR(msgContent)&"',Now(),0,1)")
end if
'��������Ϣ������cache
call cache_top()
call dvbbs_suc()
end sub

sub isbest()
Dim datetimestr
set rs=conn.execute("select * from "&TotalUseTable&" where Announceid="&replyid)
if rs.eof and rs.bof then
	Errmsg=Errmsg+"<br><li>û���ҵ�������ӣ�"
	founderr=true
	exit sub
end if
topic=rs("topic")
topicusername=rs("username")
topicuserid=rs("postuserid")
if topic="" then
	topic=left(replace(replace(rs("body"),"'",""),chr(10),","),26)
else
	topic=replace(topic,"'","")
end if
datetimestr=replace(replace(rs("dateandtime"),"����",""),"����","")

title=CheckStr(htmlencode(request.form("title")))
content=CheckStr(htmlencode(request.form("content")))
content="ԭ��" & title & content
if htmlencode(request.form("title"))="" and htmlencode(request.form("content"))="" then
	Errmsg=Errmsg+"<br><li>��д������ԭ��"
	founderr=true
	exit sub
end if
sql="update "&TotalUseTable&" set isbest=1 where boardid="&boardid&" and announceid="&cstr(replyid)
conn.Execute(sql)
conn.execute("update topic set isbest=1 where boardid="&boardid&" and topicid="&id)

sql="insert into bestTopic (title,boardid,Announceid,rootid,postusername,postuserid,dateandtime,expression) values ('"&topic&"',"&rs("boardid")&","&rs("Announceid")&","&rs("rootid")&",'"&topicusername&"',"&rs("postuserid")&",'"&datetimestr&"','"&rs("expression")&"')"
conn.execute(sql)

sql="insert into log (l_announceid,l_boardid,l_touser,l_username,l_content,l_ip) values ("&id&","&boardid&",'"&topicusername&"','"&membername&"','���뾫����"&topic&"����"&content& "��"&allmsg&"','"&ip&"')"
conn.execute(sql)
conn.execute("update [user] set userWealth=userWealth+"&doWealth&",userCP=userCP+"&douserCP&",userEP=userEP+"&douserEP&",userIsBest=userisBest+1 where userid="&topicuserid)
sucmsg="���뾫����"&topic&"����"&content& "��"&allmsg&""
if request("ismsg")="1" then
	msgcontent="����������ӡ�[url=dispbbs.asp?boardid="&boardid&"&ID="&ID&"&replyID="&replyID&"&skin=1]"&topic&"[/url]����"&replace(Content,"ԭ��","")&"�����뾫�����ҽ�����"&replace(allmsg,"�û�������","")&"�Ĳ���"
	if request("msg")<>"" then
	msgContent=msgContent & chr(10) & "����Ϊ�����߸����ĸ��ԣ�" & request("msg")
	end if
	conn.Execute("insert into message(incept,sender,title,content,sendtime,flag,issend) values('"&TopicUsername&"','"&membername&"','ϵͳ��Ϣ','"&checkSTR(msgContent)&"',Now(),0,1)")
end if
call dvbbs_suc()
set rs=nothing
end sub

sub nobest()
set rs=conn.execute("select topic,username,postuserid from "&TotalUseTable&" where Announceid="&replyid)
if rs.eof and rs.bof then
	Errmsg=Errmsg+"<br><li>û���ҵ�������ӣ�"
	founderr=true
	exit sub
else
	topic=rs(0)
	topicusername=rs(1)
	topicuserid=rs(2)
	if topic="" then
		topic="������Ϊ�ظ�����"
	end if
end if
set rs=nothing
title=CheckStr(htmlencode(request.form("title")))
content=CheckStr(htmlencode(request.form("content")))
content="ԭ��" & title & content
if htmlencode(request.form("title"))="" and htmlencode(request.form("content"))="" then
	Errmsg=Errmsg+"<br><li>��д������ԭ��"
	founderr=true
	exit sub
end if

sql="update "&TotalUseTable&" set isbest=0 where boardid="&boardid&" and announceid="&replyid
conn.Execute(sql)
conn.execute("update topic set isbest=0 where boardid="&boardid&" and topicid="&id)

conn.execute("delete from besttopic where Announceid="&replyID)

sql="insert into log (l_announceid,l_boardid,l_touser,l_username,l_content,l_ip) values ("&id&","&boardid&",'"&topicusername&"','"&membername&"','���������"&topic&"����"&content& "��"&allmsg&"','"&ip&"')"
conn.execute(sql)

conn.execute("update [user] set userWealth=userWealth+"&doWealth&",userCP=userCP+"&douserCP&",userEP=userEP+"&douserEP&",userIsBest=userisBest-1 where userid="&topicuserid)
sucmsg="���������"&topic&"����"&content& "��"&allmsg&""
if request("ismsg")="1" then
	msgcontent="����������ӡ�[url=dispbbs.asp?boardid="&boardid&"&ID="&ID&"&replyID="&replyID&"&skin=1]"&topic&"[/url]����"&replace(Content,"ԭ��","")&"������������ҽ�����"&replace(allmsg,"�û�������","")&"�Ĳ���"
	if request("msg")<>"" then
	msgContent=msgContent & chr(10) & "����Ϊ�����߸����ĸ��ԣ�" & request("msg")
	end if
	conn.Execute("insert into message(incept,sender,title,content,sendtime,flag,issend) values('"&TopicUsername&"','"&membername&"','ϵͳ��Ϣ','"&checkSTR(msgContent)&"',Now(),0,1)")
end if
call dvbbs_suc()
end sub

sub dele()
dim todaynum
set rs=conn.execute("select topic,username,postuserid,DateAndTime from "&TotalUseTable&" where Announceid="&replyid)
if rs.eof and rs.bof then
	Errmsg=Errmsg+"<br><li>û���ҵ�������ӣ�"
	founderr=true
	exit sub
else
	topic=rs(0)
	topicusername=rs(1)
	topicuserid=rs(2)
	if topic="" then
		topic="������Ϊ�ظ�����"
	end if

	if datediff("d",rs(3),now())=0 then
	todaynum=1
	else
	todaynum=0
	end if
end if
set rs=nothing

'�ж��û��Ƿ���ɾ������Ȩ��
if (Cint(GroupSetting(11))=1 and TopicUsername=membername) then
	candeltopic=true
end if
if FoundUserPer and Cint(GroupSetting(11))=1 and TopicUsername=membername then
	candeltopic=true
elseif FoundUserPer and Cint(GroupSetting(11))=0 and TopicUsername=membername then
	candeltopic=false
end if
if FoundUserPer and Cint(GroupSetting(18))=1 and TopicUsername<>membername then
	candeltopic=true
elseif FoundUserPer and Cint(GroupSetting(18))=0 and TopicUsername<>membername then
	candeltopic=false
end if
if not candeltopic then
	Errmsg=Errmsg+"<br><li>��û��ִ�д˲�����Ȩ�ޡ�"
	founderr=true
	exit sub
end if

title=CheckStr(htmlencode(request.form("title")))
content=CheckStr(htmlencode(request.form("content")))
content="ԭ��" & title & content
if htmlencode(request.form("title"))="" and htmlencode(request.form("content"))="" then
	Errmsg=Errmsg+"<br><li>��д������ԭ��"
	founderr=true
	exit sub
end if

Dim LastPostime
sql="update "&TotalUseTable&" set locktopic=2 where boardid="&boardid&" and announceid="&replyid
conn.Execute(sql)
sql="select Max(dateandtime) from "&TotalUseTable&" where boardid="&boardid&" and rootid="&id&" and  locktopic<2"
set rs=conn.Execute(sql)
LastPostime=rs(0)

call isLastPost()

call LastCount(boardid)
call BoardNumSub(boardid,0,1,todaynum)

call AllboardNumSub(todaynum,1,0)

sql="update topic set child=child-1,LastPostTime='"&LastPostime&"' where boardid="&boardid&" and topicid="&id
conn.execute(sql)
	
sql="insert into log (l_announceid,l_boardid,l_touser,l_username,l_content,l_ip) values ("&id&","&boardid&",'"&topicusername&"','"&membername&"','ɾ�����ӡ�"&topic&"����"&content& "��"&allmsg&"','"&ip&"')"
conn.execute(sql)
if allmsg<>"" then
	conn.execute("update [user] set userWealth=userWealth+"&doWealth&",userCP=userCP+"&douserCP&",userEP=userEP+"&douserEP&",article=article-1,userDel=userDel-1 where userid="&topicuserid)
end if
sucmsg="ɾ�����ӡ�"&topic&"����"&content& "��"&allmsg&""
if request("ismsg")="1" then
	msgcontent="����������ӡ�[url=dispbbs.asp?boardid="&boardid&"&ID="&ID&"&replyID="&replyID&"&skin=1]"&topic&"[/url]����"&replace(Content,"ԭ��","")&"��ɾ�����ӣ��ҽ�����"&replace(allmsg,"�û�������","")&"�Ĳ���"
	if request("msg")<>"" then
	msgContent=msgContent & chr(10) & "����Ϊ�����߸����ĸ��ԣ�" & request("msg")
	end if
	conn.Execute("insert into message(incept,sender,title,content,sendtime,flag,issend) values('"&TopicUsername&"','"&membername&"','ϵͳ��Ϣ','"&checkSTR(msgContent)&"',Now(),0,1)")
end if
call dvbbs_suc()
end sub

sub delete()
Dim voteid,isvote
set rs=conn.execute("select title,postusername,postuserid,PollID,isvote from topic where topicid="&id)
if rs.eof and rs.bof then
	Errmsg=Errmsg+"<br><li>û���ҵ�������ӣ�"
	founderr=true
	exit sub
else
	topic=rs(0)
	topicusername=rs(1)
	topicuserid=rs(2)
	voteid=rs(3)
	isvote=rs(4)
	if topic="" then
		topic="������Ϊ�ظ�����"
	end if
end if
set rs=nothing

title=CheckStr(htmlencode(request.form("title")))
content=CheckStr(htmlencode(request.form("content")))
content="ԭ��" & title & content
if htmlencode(request.form("title"))="" and htmlencode(request.form("content"))="" then
	Errmsg=Errmsg+"<br><li>��д������ԭ��"
	founderr=true
	exit sub
end if
	
Dim todaynum,postnum
sql="select count(*) from "&TotalUseTable&" where rootid="&id
set rs=conn.execute(sql)
postNum=rs(0)
sql="select count(*) from "&TotalUseTable&" where rootid="&id&" and dateandtime>#"&date()&"#"
set rs=conn.execute(sql)
todayNum=rs(0)
		
if allmsg<>"" then
	sql="select postuserid from "&TotalUseTable&" where rootid="&id
	set rs=conn.execute(sql)
	do while not rs.eof
	conn.execute("update [user] set userWealth=userWealth+"&doWealth&",userCP=userCP+"&douserCP&",userEP=userEP+"&douserEP&",article=article-1,userDel=userDel-1 where userid="&rs(0))
	rs.movenext
	loop
end if
set rs=nothing
	
sql="update "&TotalUseTable&" set locktopic=2 where rootid="&id
conn.Execute(sql)
if isvote=1 then
	conn.execute("update topic set locktopic=2,isvote=0,VoteTotal=0 where topicid="&id)
	conn.execute("delete from vote where voteid="&voteid)
	conn.execute("delete from voteuser where voteid="&voteid)
else
	conn.execute("update topic set locktopic=2 where topicid="&id)
end if
	
call LastCount(boardid)
call BoardNumSub(boardid,1,postNum,todayNum)
call AllboardNumSub(todayNum,postNum,1)

sql="insert into log (l_announceid,l_boardid,l_touser,l_username,l_content,l_ip) values ("&id&","&boardid&",'"&topicusername&"','"&membername&"','ɾ�����⡶"&topic&"����"&content& "��"&allmsg&"','"&ip&"')"
conn.execute(sql)
sucmsg="ɾ�����⡶"&topic&"����"&content& "��"&allmsg&""
if request("ismsg")="1" then
	msgcontent="����������ӡ�[url=dispbbs.asp?boardid="&boardid&"&ID="&ID&"]"&topic&"[/url]����"&replace(Content,"ԭ��","")&"����ɾ�����ҽ�����"&replace(allmsg,"�û�������","")&"�Ĳ���"
	if request("msg")<>"" then
	msgContent=msgContent & chr(10) & "����Ϊ�����߸����ĸ��ԣ�" & request("msg")
	end if
	conn.Execute("insert into message(incept,sender,title,content,sendtime,flag,issend) values('"&TopicUsername&"','"&membername&"','ϵͳ��Ϣ','"&checkSTR(msgContent)&"',Now(),0,1)")
end if
call dvbbs_suc()
end sub

sub Tmove()
Dim reBoard_Setting
title=CheckStr(htmlencode(request.form("title")))
content=CheckStr(htmlencode(request.form("content")))
content="ԭ��" & title & content
if htmlencode(request.form("title"))="" and htmlencode(request.form("content"))="" then
	Errmsg=Errmsg+"<br><li>��д������ԭ��"
	founderr=true
	exit sub
end if
Dim newboardid,newParentID,nrs
Dim newtopic
set rs=server.createobject("adodb.recordset")
if request("checked")="yes" then
	if request("boardid")=request("newboardid") then
		founderr=true
		Errmsg=Errmsg+"<br>"+"<li>��������ͬ�����ڽ���ת�Ʋ�����"
		exit sub
	elseif not isInteger(request("newboardid")) then
		founderr=true
		Errmsg=Errmsg+"<br>"+"<li>�Ƿ��İ��������"
		exit sub
	else
		newboardid=request("newboardid")
	end if

	'Ŀ����̳�����ϼ���̳ID
	set rs=conn.execute("select ParentStr,Board_Setting from board where boardid="&newboardid)
	UpdateBoardID_1=rs(0) & "," & newboardid
	reBoard_Setting=split(rs(1),",")
	if Cint(reBoard_Setting(43))=1 then
		founderr=true
		Errmsg=Errmsg+"<br>"+"<li>����̳��Ϊ������̳������ת�ơ�"
		exit sub
	end if
	sql="select * from topic where boardid="&boardid&" and topicid="&id
	set rs=conn.execute(sql)
	if rs.eof and rs.bof then
		founderr=true
		Errmsg=Errmsg+"<br>"+"<li>��ѡ������Ӳ������ڡ�"
		exit sub
	else
		if request.form("isdispmove")="yes" then
			newtopic=CheckStr(request.form("topic")) & "-->" & membername & "ת��"
		else
			newtopic=CheckStr(request.form("topic"))
		end if
		if request("leavemessage")="yes" then
			sql="insert into topic (Title,Boardid,PostUsername,PostUserid,DateAndTime,Expression,LastPost,LastPostTime,child,hits,isvote,isbest,votetotal,PostTable) values ('"&newtopic&"',"&newboardid&",'"&rs("postusername")&"',"&rs("postuserid")&",'"&rs("dateandtime")&"','"&rs("Expression")&"','"&rs("LastPost")&"','"&rs("LastPosttime")&"',"&rs("child")&","&rs("hits")&","&rs("isvote")&","&rs("isbest")&","&rs("votetotal")&",'"&NowUseBBS&"')"
			conn.execute(sql)
		end if
	end if
	
	if request("leavemessage")="yes" then
		conn.execute("update topic set locktopic=1 where topicid="&id)
		set rs=conn.execute("select top 1 topicid from topic order by topicid desc")
		newparentid=rs(0)
		sql="select * from "&TotalUseTable&" where rootid="&id&" and locktopic<2 order by Announceid"
		set rs=conn.execute(sql)
		do while not rs.eof
			Sql="insert into "&NowUseBBS&"(Boardid,ParentID,username,topic,body,DateAndTime,length,rootid,layer,orders,ip,Expression,locktopic,signflag,emailflag,isbest,postuserid) values "&_
				"("&_
				newboardid&","&rs("parentid")&",'"&_
				rs("username")&"','"&_
				checkstr(rs("topic"))&"','"&_
				checkstr(rs("body"))&"','"&_
				rs("DateAndTime")&"','"&_
				rs("length")&"',"&newParentID&","&rs("layer")&","&rs("orders")&",'"&rs("ip")&"','"&_
				rs("Expression")&"',"&rs("locktopic")&","&rs("signflag")&","&rs("emailflag")&","&rs("isbest")&","&rs("postuserid")&")"
				'response.write sql
				conn.execute(sql)
		rs.movenext
		loop
		
	elseif request("leavemessage")="no" then
		if request.form("isdispmove")="yes" then
			newtopic=CheckStr(request.form("topic")) & "-->" & membername & "ת��"
		else
			newtopic=CheckStr(request.form("topic"))
		end if
		conn.execute("update topic set title='"&newtopic&"',boardid="&newboardid&" where topicid="&id)
		sql="update "&TotalUseTable&" set topic='"&newtopic&"' where announceid="&replyid
		conn.execute(sql)
		sql="update "&TotalUseTable&" set boardid="&newboardid&" where rootid="&id
		conn.Execute(sql)							
	else
		founderr=true
		Errmsg=Errmsg+"<br>"+"<li>��ѡ����Ӧ������"
		exit sub
	end if
	sql="update besttopic set boardid="&newboardid&" where rootid="&id
	conn.execute(sql)
	Dim postNum,todayNum
	'��������ӵĻظ�����������ͳ�ƶ�Ӧ����������
	sql="select count(*) from "&TotalUseTable&" where rootid="&id
	set rs=conn.execute(sql)
	postNum=rs(0)
	'����������н��ջظ�������
	sql="select count(*) from "&TotalUseTable&" where rootid="&id&" and dateandtime>=#"&date()&"#"
	set rs=conn.execute(sql)
	todayNum=rs(0)
	set rs=nothing
	'������̳��������
	call LastCount(boardid)
	call BoardNumSub(boardid,1,postNum,todayNum)
	UpdateBoardID=UpdateBoardID_1
	call LastCount(newboardid)
	call BoardNumAdd(newboardid,1,postNum,todayNum)
	'������̳���ݽ���
	sql="insert into log (l_announceid,l_boardid,l_touser,l_username,l_content,l_ip) values ("&id&","&boardid&",'"&topicusername&"','"&membername&"','ת�����⡶"&topic&"����"&content& "��"&allmsg&"','"&ip&"')"
	conn.execute(sql)
	sucmsg="ת�����⡶"&topic&"����"&content& "��"&allmsg&""

	if request("ismsg")="1" then
	msgcontent="����������ӡ�<a href=dispbbs.asp?boardid="&newboardid&"&ID="&ID&" target=_blank>"&topic&"</a>����"&replace(Content,"ԭ��","")&"����ת�ƣ��ҽ�����"&replace(allmsg,"�û�������","")&"�Ĳ���"
	if request("msg")<>"" then
	msgContent=msgContent & chr(10) & "����Ϊ�����߸����ĸ��ԣ�" & request("msg")
	end if
	conn.Execute("insert into message(incept,sender,title,content,sendtime,flag,issend) values('"&TopicUsername&"','"&membername&"','ϵͳ��Ϣ','"&checkSTR(msgContent)&"',Now(),0,1)")
	end if
	call dvbbs_suc()
else
%>
<table cellpadding=3 cellspacing=1 align=center class=tableborder1>
                        <tr>
                        <th valign=middle colspan=2>
                        <form action="admin_postings.asp" method="post">
                        <input type=hidden name="action" value="move">
                        <input type=hidden name="checked" value="yes">
                        <input type=hidden name="boardid" value="<%=request("boardid")%>">
                        <input type=hidden name="replyid" value="<%=request("replyid")%>">
                        <input type=hidden name="id" value="<%=request("id")%>">
						<input type=hidden value="<%=CheckStr(htmlencode(request.form("title")))%>" name="title">
						<input type=hidden value="<%=CheckStr(htmlencode(request.form("content")))%>" name="content">
						<input type=hidden value="<%=doWealth%>" name="doWealth">
						<input type=hidden value="<%=dousercp%>" name="dousercp">
						<input type=hidden value="<%=douserep%>" name="douserep">
<input type=hidden value="<%=request.form("msg")%>" name="msg">
<input type=hidden value="<%=request.form("ismsg")%>" name="ismsg">
                        �ƶ�����</th></tr>

                                    <tr>
                                    <td class=tablebody1 valign=middle>
                                    <b>�ƶ�ѡ��</td>
                                    <td class=tablebody1 valign=middle>
                                    <input name="leavemessage" type="radio" value="yes"> �ƶ�������һ���Ѿ�������������ԭ��̳<br>
									<input name="leavemessage" type="radio" value="no" checked> �ƶ������������ԭ��̳��ɾ��<br>
									<input name="isdispmove" type="checkbox" value="yes"> �ƶ���������Ƿ���ʾ����������
                                    </td>
                                    </tr>
                                    <tr>
                                    <td class=tablebody1 valign=middle>
                                    <b>��������</td>
                                    <td class=tablebody1 valign=middle>
   <input name="topic" value="<%=topic%>" size=50>
                                    </td>
                                    </tr>
                            <tr>
                        <td class=tablebody1 valign=top><b>�ƶ�����</b></td>
                        <td class=tablebody1 valign=top>
<%
response.write "<select name=newboardid size=1><option selected>�ƶ�������ѡ��</option>"
set rs=conn.execute("select boardid,boardtype,depth,Board_Setting from board order by rootid,orders")
do while not rs.EOF
reBoard_Setting=split(rs(3),",")
response.write "<option value="""&rs(0)&""" "
response.write ">"

select case rs(2)
case 0
response.write "��"
case 1
response.write "&nbsp;&nbsp;��"
case 2
response.write "&nbsp;&nbsp;��&nbsp;&nbsp;��"
case 3
response.write "&nbsp;&nbsp;��&nbsp;&nbsp;��&nbsp;&nbsp;��"
case 4
response.write "&nbsp;&nbsp;��&nbsp;&nbsp;��&nbsp;&nbsp;��&nbsp;&nbsp;��"
case 5
response.write "&nbsp;&nbsp;��&nbsp;&nbsp;��&nbsp;&nbsp;��&nbsp;&nbsp;��&nbsp;&nbsp;��"
case 6
response.write "&nbsp;&nbsp;��&nbsp;&nbsp;��&nbsp;&nbsp;��&nbsp;&nbsp;��&nbsp;&nbsp;��&nbsp;&nbsp;��"
case 7
response.write "&nbsp;&nbsp;��&nbsp;&nbsp;��&nbsp;&nbsp;��&nbsp;&nbsp;��&nbsp;&nbsp;��&nbsp;&nbsp;��&nbsp;&nbsp;��"
case 8
response.write "&nbsp;&nbsp;��&nbsp;&nbsp;��&nbsp;&nbsp;��&nbsp;&nbsp;��&nbsp;&nbsp;��&nbsp;&nbsp;��&nbsp;&nbsp;��&nbsp;&nbsp;��"
case 9
response.write "&nbsp;&nbsp;��&nbsp;&nbsp;��&nbsp;&nbsp;��&nbsp;&nbsp;��&nbsp;&nbsp;��&nbsp;&nbsp;��&nbsp;&nbsp;��&nbsp;&nbsp;��&nbsp;&nbsp;��"
end select
response.write rs(1)
if Cint(reBoard_Setting(43))=1 then
response.write "(����ת��)"
end if
response.write "</option>"
rs.MoveNext
loop
set rs=nothing
%>        
          </select>&nbsp;ע�⣺��������ͬ��̳�ڽ����ƶ�������ע�������ƶ�����̳Ϊ������̳
			</td>
                        </tr>
                    <tr>
                <td class=tablebody2 valign=middle colspan=2 align=center><input type=submit name="submit" value="�� ��"></td></tr></form></table>
            </table>
            </td></tr>
            </table>
<%
end if
end sub

sub copy()
Dim reBoard_Setting
set rs=conn.execute("select topic,username,postuserid from "&TotalUseTable&" where Announceid="&replyid)
if rs.eof and rs.bof then
	Errmsg=Errmsg+"<br><li>û���ҵ�������ӣ�"
	founderr=true
	exit sub
else
	Topic=CheckStr(rs(0))
	topicusername=rs(1)
	topicuserid=rs(2)
	if topic="" then
		topic="������Ϊ�ظ�����"
	end if
end if
set rs=nothing

'�ж��û��Ƿ����ƶ�����Ȩ��
if (Cint(GroupSetting(12))=1 and TopicUsername=membername) then
	canmovetopic=true
end if
if FoundUserPer and Cint(GroupSetting(12))=1 and TopicUsername=membername then
	canmovetopic=true
elseif FoundUserPer and Cint(GroupSetting(12))=0 and TopicUsername=membername then
	canmovetopic=false
end if
if FoundUserPer and Cint(GroupSetting(19))=1 and TopicUsername<>membername then
	canmovetopic=true
elseif FoundUserPer and Cint(GroupSetting(19))=0 and TopicUsername<>membername then
	canmovetopic=false
end if
if not canmovetopic then
	Errmsg=Errmsg+"<br><li>��û��ִ�д˲�����Ȩ�ޡ�"
	founderr=true
	exit sub
end if
title=CheckStr(htmlencode(request.form("title")))
content=CheckStr(htmlencode(request.form("content")))
content="ԭ��" & title & content
if htmlencode(request.form("title"))="" and htmlencode(request.form("content"))="" then
	Errmsg=Errmsg+"<br><li>��д������ԭ��"
	founderr=true
	exit sub
end if
	
if request("checked")="yes" then
	Dim newboardid
	Dim todaynum,postnum
	sql="select count(*) from "&TotalUseTable&" where rootid="&id
	set rs=conn.execute(sql)
	postNum=rs(0)
	sql="select count(*) from "&TotalUseTable&" where rootid="&id&" and dateandtime>#"&date()&"#"
	set rs=conn.execute(sql)
	todayNum=rs(0)
	if request("boardid")=request("newboardid") then
		founderr=true
		Errmsg=Errmsg+"<br>"+"<li>��������ͬ�����ڽ���ת�Ʋ�����"
		exit sub
	elseif not isInteger(request("newboardid")) then
		founderr=true
		Errmsg=Errmsg+"<br>"+"<li>�Ƿ��İ��������"
		exit sub
	else
		newboardid=request("newboardid")
	end if

	'Ŀ����̳�����ϼ���̳ID
	set rs=conn.execute("select ParentStr,Board_Setting from board where boardid="&newboardid)
	UpdateBoardID=rs(0) & "," & newboardid
	reBoard_Setting=split(rs(1),",")
	if Cint(reBoard_Setting(43))=1 then
		founderr=true
		Errmsg=Errmsg+"<br>"+"<li>����̳��Ϊ������̳������ת�ơ�"
		exit sub
	end if
	sql="select boardid from "&TotalUseTable&" where announceid="&replyid&" and boardid="&cstr(boardid)
	set rs=conn.execute(sql)
	if rs.eof and rs.bof then
		founderr=true
		Errmsg=Errmsg+"<br>"+"<li>��ѡ������Ӳ������ڡ�"
		exit sub
	end if

	Dim newtopic,trs
	set rs=server.createobject("adodb.recordset")
	sql="select * from "&TotalUseTable&" where announceid="&replyid
	rs.open sql,conn,1,1
	if request.form("isdispmove")="yes" then
		newtopic=CheckStr(request.form("topic")) & "-->" & membername & "���"
	else
		newtopic=CheckStr(request.form("topic"))
	end if
	sql="insert into topic (Title,Boardid,PostUsername,PostUserid,DateAndTime,Expression,LastPost,LastPostTime,child,hits,isvote,isbest,votetotal,PostTable) values ('"&newtopic&"',"&newboardid&",'"&rs("username")&"',"&rs("postuserid")&",Now(),'"&rs("Expression")&"','$$$$$$',Now(),0,0,0,0,0,'"&NowUseBBS&"')"
	conn.execute(sql)
	set trs=conn.execute("select top 1 topicid from topic order by topicid desc")
	Sql="insert into "&NowUseBBS&"(Boardid,ParentID,username,topic,body,DateAndTime,length,rootid,layer,orders,ip,Expression,locktopic,signflag,emailflag,isbest,postuserid) values "&_
				"("&_
				newboardid&",0,'"&_
				rs("username")&"','"&_
				newtopic&"','"&_
				rs("body")&"','"&_
				rs("DateAndTime")&"','"&_
				rs("length")&"',"&trs(0)&",1,0,'"&rs("ip")&"','"&_
				rs("Expression")&"',"&rs("locktopic")&","&rs("signflag")&","&rs("emailflag")&","&rs("isbest")&","&rs("postuserid")&")"
	conn.execute(sql)
	rs.close
	set rs=nothing
	
	'������̳��������
	call LastCount(Newboardid)
	call BoardNumAdd(newboardid,1,postNum,todayNum)
	call AllboardNumAdd(todayNum,postNum,1)
	
	sql="insert into log (l_announceid,l_boardid,l_touser,l_username,l_content,l_ip) values ("&id&","&boardid&",'"&topicusername&"','"&membername&"','�������ӡ�"&topic&"����"&content& "��"&allmsg&"','"&ip&"')"
	conn.execute(sql)
	'������̳���ݽ���
	sucmsg="�������ӡ�"&topic&"����"&content& "��"&allmsg&""

	if request("ismsg")="1" then
	msgcontent="����������ӡ�[url=dispbbs.asp?boardid="&boardid&"&ID="&ID&"&replyID="&replyID&"&skin=1]"&topic&"[/url]����"&replace(Content,"ԭ��","")&"�������Ƶ��������棬�ҽ�����"&replace(allmsg,"�û�������","")&"�Ĳ���"
	if request("msg")<>"" then
	msgContent=msgContent & chr(10) & "����Ϊ�����߸����ĸ��ԣ�" & request("msg")
	end if
	conn.Execute("insert into message(incept,sender,title,content,sendtime,flag,issend) values('"&TopicUsername&"','"&membername&"','ϵͳ��Ϣ','"&checkSTR(msgContent)&"',Now(),0,1)")
	end if
	call dvbbs_suc()
else
%>
<table cellpadding=3 cellspacing=1 align=center class=tableborder1 >
                        <tr>
                        <th valign=middle colspan=2>
                        <form action="admin_postings.asp" method="post">
                        <input type=hidden name="action" value="copy">
                        <input type=hidden name="checked" value="yes">
                        <input type=hidden name="boardid" value="<%=request("boardid")%>">
                        <input type=hidden name="replyid" value="<%=request("replyid")%>">
                        <input type=hidden name="id" value="<%=request("id")%>">
<input type=hidden value="<%=CheckStr(htmlencode(request.form("title")))%>" name="title">
<input type=hidden value="<%=CheckStr(htmlencode(request.form("content")))%>" name="content">
<input type=hidden value="<%=doWealth%>" name="doWealth">
<input type=hidden value="<%=dousercp%>" name="dousercp">
<input type=hidden value="<%=douserep%>" name="douserep">
<input type=hidden value="<%=request.form("msg")%>" name="msg">
<input type=hidden value="<%=request.form("ismsg")%>" name="ismsg">
                        ��������</th></tr>

                                    <tr>
                                    <td class=tablebody1 valign=middle>
                                    <b>����˵��</td>
                                    <td class=tablebody1 valign=middle>
�����ӽ����Ƶ������̳����Ϊ�µ����ӣ��������������������ָ�������ӵ����⣬ԭ���ӽ�������ԭ����̳
                                    </td>
                                    </tr>
                                    <tr>
                                    <td class=tablebody1 valign=middle>
                                    <b>��������</td>
                                    <td class=tablebody1 valign=middle>
   <input name="topic" value="<%=topic%>" size=50>&nbsp;<input name="isdispmove" type="checkbox" value="yes"> �ƶ���������Ƿ���ʾ����������
                                    </td>
                                    </tr>
                            <tr>
                        <td class=tablebody1 valign=top><b>�ƶ�����</b></td>
                        <td class=tablebody1 valign=top>
<%
response.write "<select name=newboardid size=1><option selected>�ƶ�������ѡ��</option>"
set rs=conn.execute("select boardid,boardtype,depth,Board_Setting from board order by rootid,orders")
do while not rs.EOF
reBoard_Setting=split(rs(3),",")
response.write "<option value="""&rs(0)&""" "
response.write ">"

select case rs(2)
case 0
response.write "��"
case 1
response.write "&nbsp;&nbsp;��"
case 2
response.write "&nbsp;&nbsp;��&nbsp;&nbsp;��"
case 3
response.write "&nbsp;&nbsp;��&nbsp;&nbsp;��&nbsp;&nbsp;��"
case 4
response.write "&nbsp;&nbsp;��&nbsp;&nbsp;��&nbsp;&nbsp;��&nbsp;&nbsp;��"
case 5
response.write "&nbsp;&nbsp;��&nbsp;&nbsp;��&nbsp;&nbsp;��&nbsp;&nbsp;��&nbsp;&nbsp;��"
case 6
response.write "&nbsp;&nbsp;��&nbsp;&nbsp;��&nbsp;&nbsp;��&nbsp;&nbsp;��&nbsp;&nbsp;��&nbsp;&nbsp;��"
case 7
response.write "&nbsp;&nbsp;��&nbsp;&nbsp;��&nbsp;&nbsp;��&nbsp;&nbsp;��&nbsp;&nbsp;��&nbsp;&nbsp;��&nbsp;&nbsp;��"
case 8
response.write "&nbsp;&nbsp;��&nbsp;&nbsp;��&nbsp;&nbsp;��&nbsp;&nbsp;��&nbsp;&nbsp;��&nbsp;&nbsp;��&nbsp;&nbsp;��&nbsp;&nbsp;��"
case 9
response.write "&nbsp;&nbsp;��&nbsp;&nbsp;��&nbsp;&nbsp;��&nbsp;&nbsp;��&nbsp;&nbsp;��&nbsp;&nbsp;��&nbsp;&nbsp;��&nbsp;&nbsp;��&nbsp;&nbsp;��"
end select
response.write rs(1)
if Cint(reBoard_Setting(43))=1 then
response.write "(����ת��)"
end if
response.write "</option>"
rs.MoveNext
loop
set rs=nothing
%>        
          </select>&nbsp;ע�⣺��ͬ��̳�䲻�ܽ��и��Ʋ�����ע�������ƶ�����̳Ϊ������̳
			</td>
                        </tr>
                    <tr>
                <td class=tablebody2 valign=middle colspan=2 align=center><input type=submit name="submit" value="�� ��"></td></tr></form></table></td></tr></table>
            </table>
            </td></tr>
            </table>
<%
	end if
	end sub


'����ָ����̳��Ϣ
function LastCount(boardid)
Dim LastTopic,body,LastRootid,LastPostTime,LastPostUser
Dim LastPost,uploadpic_n,Lastpostuserid,Lastid

set rs=conn.execute("select top 1 T.title,b.Announceid,b.dateandtime,b.username,b.postuserid,b.rootid from "&NowUseBBS&" b inner join Topic T on b.rootid=T.TopicID where b.boardid="&boardid&" and  b.locktopic<2 order by b.announceid desc")
if not(rs.eof and rs.bof) then
	Lasttopic=replace(left(rs(0),15),"$","")
	LastRootid=rs(1)
	LastPostTime=rs(2)
	LastPostUser=rs(3)
	LastPostUserid=rs(4)
	Lastid=rs(5)
else
	LastTopic="��"
	LastRootid=0
	LastPostTime=now()
	LastPostUser="��"
	LastPostUserid=0
	Lastid=0
end if
set rs=nothing

LastPost=LastPostUser & "$" & LastRootid & "$" & LastPostTime & "$" & LastTopic & "$" & uploadpic_n & "$" & LastPostUserID & "$" & LastID & "$" & BoardID
Dim SplitUpBoardID,SplitLastPost
SplitUpBoardID=split(UpdateBoardID,",")
For i=0 to ubound(SplitUpBoardID)
	set rs=conn.execute("select LastPost from board where boardid="&SplitUpBoardID(i))
	if not (rs.eof and rs.bof) then
		SplitLastPost=split(rs(0),"$")
		if isnumeric(LastRootID) and isnumeric(SplitLastPost(1)) then
			if ubound(SplitLastPost)=7 and clng(LastRootID)<>clng(SplitLastPost(1)) then
				conn.execute("update board set LastPost='"&replace(LastPost,"'","''")&"' where boardid="&SplitUpBoardID(i))
			end if
		end if
	end if
Next
set rs=nothing
end function

'���淢��������
sub BoardNumAdd(boardid,topicNum,postNum,todayNum)
sql="update board set lastbbsnum=lastbbsnum+"&postNum&",lasttopicNum=lasttopicNum+"&topicNum&",todayNum=todayNum+"&todayNum&" where boardid in ("&UpdateBoardID&")"
conn.execute(sql)
end sub
	
'���淢��������
sub BoardNumSub(boardid,topicNum,postNum,todayNum)
sql="update board set lastbbsnum=lastbbsnum-"&postNum&",lasttopicNum=lasttopicNum-"&topicNum&",todayNum=todayNum-"&todayNum&" where boardid in ("&UpdateBoardID&")"
conn.execute(sql)
end sub
	
'������̳����������
function AllboardNumAdd(todayNum,postNum,topicNum)
sql="update config set TodayNum=todayNum+"&todaynum&",BbsNum=bbsNum+"&postNum&",TopicNum=topicNum+"&TopicNum
conn.execute(sql)
end function

'������̳����������
function AllboardNumSub(todayNum,postNum,topicNum)
sql="update config set TodayNum=todayNum-"&todaynum&",BbsNum=bbsNum-"&postNum&",TopicNum=topicNum-"&TopicNum
conn.execute(sql)
end function

'�ж��Ƿ�Ϊ�������ظ�
function isLastPost()
Dim LastTopic,body,LastRootid,LastPostTime,LastPostUser
Dim LastPost,uploadpic_n,LastPostUserID,LastID
isLastPost=false
'ȡ�õ�ǰ�������ظ�id
sql="select LastPost from topic where topicid="&id
set rs=conn.execute(sql)
if not (rs.eof and rs.bof) then
	if not isnull(rs(0)) and rs(0)<>"" then
		if Clng(split(rs(0),"$")(1))=Clng(replyid) then isLastPost=true
	end if
end if
if isLastPost then
	sql="select top 1 topic,body,Announceid,dateandtime,username,PostUserid,rootid,boardid from "&TotalUseTable&" where rootid="&id&" and locktopic<2 order by Announceid desc"
	set rs=conn.execute(sql)
	if not(rs.eof and rs.bof) then
		body=rs(1)
		LastRootid=rs(2)
		LastPostTime=rs(3)
		LastPostUser=replace(rs(4),"$","")
		LastTopic=left(replace(body,"$",""),20)
		LastPostUserID=rs(5)
		LastID=rs(6)
		BoardID=rs(7)
	else
		LastTopic="��"
		LastRootid=0
		LastPostTime=now()
		LastPostUser="��"
		LastPostUserID=0
		LastID=0
		BoardID=0
	end if
	set rs=nothing
	LastPost=LastPostUser & "$" & LastRootid & "$" & LastPostTime & "$" & replace(left(LastTopic,20),"$","") & "$" & uploadpic_n & "$" & LastPostUserID & "$" & LastID & "$" & BoardID
	conn.execute("update topic set LastPost='"&LastPost&"' where topicid="&id)
end if
end function

sub cache_top()
Dim RsTop
Dim AllTopNum
myCache.name="AllTopNum"
set RsTop=server.createobject("adodb.recordset")
sql="select TopicID from topic where istop=2 and locktopic<2 ORDER BY lastposttime desc"
RsTop.open sql,conn,1,1
if RsTop.eof and RsTop.bof then
	AllTopNum=0
else
	AllTopNum=RsTop.recordcount
end if
myCache.add AllTopNum,dateadd("n",60,now)
set RsTop=nothing
end sub
%>