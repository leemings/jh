<!--#include file="conn.asp"-->
<!-- #include file="inc/const.asp" -->
<!-- #include file="inc/char_board.asp" -->
<!-- #include file="inc/char_login.asp" -->
<!-- #include file="inc/chkinput.asp" -->
<!-- #include file="inc/ubbcode.asp" -->
<!--#include file="md5.asp"-->
<!--#include file="jhConst.asp"-->
<%
dim announceid
dim UserName
dim userPassword
dim useremail
dim Topic
dim body
dim rootid
dim dateTimeStr
dim ip
dim Expression
dim signflag
dim mailflag
dim msgflag
dim usercookies
dim ihaveupfile,upfileinfo,upfilelen
dim ChenTopicColor
dim ArticleRandom
dim speak
dim msgNotify
dim mess,says,act,towhoway,towho,addwordcolor,saycolor,addsays,saystr

if BoardID="" or not isInteger(BoardID) or BoardID=0 then
	Errmsg=Errmsg+"<br>"+"<li>����İ����������ȷ�����Ǵ���Ч�����ӽ��롣"
	founderr=true
else
	BoardID=clng(BoardID)
end if
IP=replace(Request.ServerVariables("REMOTE_ADDR"),"'","")
Expression=Checkstr(Request.Form("Expression"))
Topic=Checkstr(request.Form("subject"))
Body=Checkstr(request.form("Content"))
UserName=Checkstr(trim(request.Form("username")))
signflag=Checkstr(trim(request.Form("signflag")))
mailflag=Checkstr(trim(request.Form("emailflag")))
msgflag=Checkstr(trim(request.Form("msgflag")))
ChenTopicColor=Checkstr(trim(request.Form("ChenTopicColor")))
speak=Checkstr(trim(request.Form("speak")))
ArticleRandom=Checkstr(trim(request.Form("ArticleRandom")))
msgNotify=Checkstr(trim(request.Form("msgNotify")))
if memberword=trim(request.Form("passwd")) then
	UserPassWord=Checkstr(trim(request.Form("passwd")))
else
	UserPassWord=md5(Checkstr(trim(request.Form("passwd"))))
end if
if request.form("upfilerename")<>"" then
	ihaveupfile=1
	upfileinfo=replace(request.form("upfilerename"),"'","")
	upfileinfo=replace(upfileinfo,";","")
	upfileinfo=replace(upfileinfo,"--","")
	upfileinfo=replace(upfileinfo,")","")
	upfilelen=len(upfileinfo)
	upfileinfo=left(upfileinfo,upfilelen-1)
else
	ihaveupfile=0
end if
if signflag="yes" then
	signflag=1
else
	signflag=0
end if
if mailflag="yes" then
	mailflag=1
else
	mailflag=0
end if
if msgflag="yes" then
	msgflag=1
else
	msgflag=0
end if
dim mailusemoney,msgusemoney
mailusemoney=0:msgusemoney=0
if msgflag=1 then
	if isnull(myvip) or myvip<>1 then
		msgusemoney=clng(forum_user(28))
	else
		msgusemoney=clng(forum_user(29))
	end if
end if
if mailflag=1 then
	if isnull(myvip) or myvip<>1 then
		mailusemoney=clng(forum_user(30))
	else
		mailusemoney=clng(forum_user(31))
	end if
end if
if msgflag=1 or mailflag=1 then
	if mymoney<msgusemoney+mailusemoney then
		Errmsg=ErrMsg+"<Br>"+"<li>�����ֽ𲻹�֧������֪ͨ���ʼ�֪ͨ�ķ��ã�"
		founderr=true
	end if
end if

dim rndnum,num1
if cint(Board_Setting(30))=1 then
	if not (isnull(session("lastpost")) or boardmaster or master or superboardmaster) then
		if DateDiff("s",session("lastpost"),Now())<cint(Board_Setting(31)) then
   		ErrMsg=ErrMsg+"<Br>"+"<li>����̳���Ʒ�������ʱ��Ϊ"&Board_Setting(31)&"�룬���Ժ��ٷ���"
   		FoundErr=True
		end if
	end if
end if
if chkpost=false then
	ErrMsg=ErrMsg+"<Br>"+"<li>���ύ�����ݲ��Ϸ����벻Ҫ���ⲿ�ύ���ԡ�"
	FoundErr=True
end if
if UserName="" or UserPassWord="" then
	username=membername
	UserPassWord=memberword
end if
if UserName="" then
	ErrMsg=ErrMsg+"<Br>"+"<li>����������������"
	FoundErr=True
end if
if Topic="" then
	FoundErr=True
	ErrMsg=ErrMsg+"<Br>"+"<li>���ӵ����ⲻӦΪ�ա�"
elseif strLength(topic)>50 then
	FoundErr=True
	ErrMsg=ErrMsg+"<Br>"+"<li>�������ⳤ�Ȳ��ܳ���50"
end if
if strLength(body)>Clng(Board_Setting(16)) then
	ErrMsg=ErrMsg+"<Br>"+"<li>�������ݲ��ô���" & CSTR(Board_Setting(16)) & "bytes"
	FoundErr=true
end if
if body="" then
	ErrMsg=ErrMsg+"<Br>"+"<li>û����д���ݡ�"
	FoundErr=true
end if
session("lastpost")=Now()
if founderr then
	call nav()
	call head_var(2,0,"","")
	call dvbbs_error()
else
	stats="�������ӳɹ�"
	call main()
	if founderr then 
		call nav()
		call head_var(1,BoardDepth,0,0)
		call dvbbs_error()
	end if

says="<bgsound src=wav/xintie.wav loop=1><img src=img/xintie.gif><font color=000099><b>������������</b></font>"&mess
act="��Ϣ"
towhoway=0
towho="���"
addwordcolor="660099"
saycolor="008888"
addsays="��"
saystr="<script>parent.sh("& chr(39) & addwordcolor & chr(39) &","& chr(39) & saycolor & chr(39) &","& chr(39) & act & chr(39) &","& chr(39) & jhname & chr(39) &","& chr(39) & addsays & chr(39) &","& chr(39) & towho & chr(39) &"," & chr(39) & says & chr(39) &"," & towhoway &  ",0);<"&"/script>"
addmsg saystr
Function Yushu(a)
 Yushu=(a and 31)
End Function
Sub AddMsg(Str)
Application.Lock()
Application("SayCount")=Application("SayCount")+1
i="SayStr"&YuShu(Application("SayCount"))
Application(i)=Str
Application.UnLock()
End Sub

end if
call footer()

sub main()
if (Cint(Board_Setting(43))=1 or Cint(Board_Setting(0))=1) and not (master or superboardmaster or boardmaster) then
	Errmsg=Errmsg+"<br>"+"<li>����̳�Ѿ�������Ա�����˲���������"
	founderr=true
	exit sub
end if
if Cint(GroupSetting(3))=0 then
	Errmsg=Errmsg+"<br>"+"<li>��û���ڱ���̳������Ȩ�ޣ���<a href=login.asp>��¼</a>����ͬ����Ա��ϵ��"
	founderr=true
	exit sub
end if
if cint(Board_Setting(2))=1 then
	if not founduser then
		Errmsg=Errmsg+"<br>"+"<li>����̳Ϊ��֤��̳����<a href=login.asp>��¼</a>��ȷ�������û����Ѿ��õ�����Ա����֤����롣"
		founderr=true
		exit sub
	else
		if chkboardlogin(boardid,membername)=false then
		Errmsg=Errmsg+"<br>"+"<li>����̳Ϊ��֤��̳����ȷ�������û����Ѿ��õ�����Ա����֤����롣"
		founderr=true
		exit sub
		end if
	end if
end if

usercookies=request.Cookies("aspsky")("usercookies")
if isnull(usercookies) or usercookies="" then usercookies=3

if chkuserlogin(username,userpassword,usercookies,2)=false then
	errmsg=errmsg+"<br>"+"<li>�����û����������ڣ���������������󣬻��������ʺ��ѱ�����Ա������"
	founderr=true
	exit sub
end if

if cint(Board_Setting(1))=1 then
	if Cint(GroupSetting(37))=0 then
		Errmsg=ErrMsg+"<Br>"+"<li>��û��Ȩ�޽���������̳��"
		founderr=true
		exit sub
	end if
end if

dim notifyusemoney
notifyusemoney=0
if not isnull(msgNotify) and msgNotify<>"" then
	set rs=conn.execute("select username from [user] where username='"&msgNotify&"'")
	if rs.bof or rs.eof then
		set rs=nothing
		Errmsg=ErrMsg+"<Br>"+"<li>��ָ����֪ͨ�ʹ��߲����ڣ�"
		founderr=true
		exit sub
	end if
	rs.close
	if isnull(myvip) or myvip<>1 then
		notifyusemoney=clng(forum_user(32))
	else
		notifyusemoney=clng(forum_user(33))
	end if
	if mymoney<notifyusemoney then
		Errmsg=ErrMsg+"<Br>"+"<li>�����ֽ𲻹�֧����������ע��ķ��ã�"
		founderr=true
		exit sub
	end if
end if
select case cint(forum_setting(73))
case 1
	if myUserPPD*cint(forum_setting(74))<=myTodayTopic then
		Errmsg=ErrMsg+"<Br>"+"<li>�����췢���������(<b>"&myTodayTopic&"</b>)�Ѿ�����������(<b>"&(myUserPPD*cint(forum_setting(74)))&"</b>)�����ܼ����������⣡"
		founderr=true
		exit sub
	end if
case 3
	if myUserPPD*cint(forum_setting(74))<=myTodayTopic+myTodayReply then
		Errmsg=ErrMsg+"<Br>"+"<li>�����췢���������(<b>"&(myTodayTopic+myTodayReply)&"</b>)�Ѿ�����������(<b>"&(myUserPPD*cint(forum_setting(74)))&"</b>)�����ܼ����������ӣ�"
		founderr=true
		exit sub
	end if
end select

rem ����������Ϣ
dim locktopic
if master or superboardmaster or boardmaster then
	isaudit=0
	locktopic=0
elseif isaudit=1 and not (master or superboardmaster or boardmaster) then
	isaudit=1
	locktopic=3
else
	isaudit=0
	locktopic=0
end if
dim LastPost,LastPost_1,uploadpic_n,Forumupload,u
dim LastPostTimes
DateTimeStr=replace(replace(CSTR(NOW()+Forum_Setting(0)/24),"����",""),"����","")
'���������
sql="insert into topic (Title,Boardid,PostUsername,PostUserid,DateAndTime,Expression,LastPost,LastPostTime,PostTable,locktopic,ChenTopicColor) values ('"&topic&"',"&boardid&",'"&username&"',"&userid&",'"&DateTimeStr&"','"&Expression&"','$$"&DateTimeStr&"$$$$','"&DateTimeStr&"','"&NowUseBbs&"',"&locktopic&",'"&ChenTopicColor&"')"
conn.execute(sql)

set rs=conn.execute("select max(topicid) from topic")
rootid=rs(0)
Forum_upload="gif,jpg,jpeg,bmp,zip,rar,html,swf,mid,midi,flash,rm,ra,asf,avi,wmv,exe,xls,dos,ftp,mp3,m3u,txt,mdb,dll,IMG"
Forumupload=split(Forum_upload,",")
for u=0 to ubound(Forumupload)
	if instr(body,"[upload="&Forumupload(u)&"]") or instr(body,"."&Forumupload(u)&"") or instr(body,"["&Forumupload(u)&"]")  then
		uploadpic_n=Forumupload(u)
		exit for
	end if
next
if instr(body,"viewfile.asp?ID=") then uploadpic_n="down"

'����ظ���
Sql="insert into "&NowUseBbs&"(Boardid,ParentID,username,topic,body,DateAndTime,length,rootid,layer,orders,ip,Expression,locktopic,signflag,emailflag,isbest,PostUserID,isupload,isaudit,speak) values ("&boardid&",0,'"&username&"','"&topic&"','"&body&"','"&DateTimeStr&"','"&strlength(body)&"',"&rootid&",1,0,'"&ip&"','"&Expression&"',"&locktopic&","&signflag&","&(mailflag+msgflag*2)&",0,"&userid&","&ihaveupfile&","&isaudit&",'"&speak&"')"
conn.execute(sql)

set rs=conn.execute("select max(Announceid) from "&NowUseBbs&"")
Announceid=rs(0)

if ihaveupfile=1 then conn.execute("update dv_upfile set F_AnnounceID='"&rootid&"|"&AnnounceID&"',F_Readme='"&Topic&"' where F_ID in ("&upfileinfo&")")

LastPost=replace(username,"$","") & "$" & Announceid & "$" & DateTimeStr & "$" & replace(cutStr(reUBBCode(body,false),20),"$","") & "$" & uploadpic_n & "$" & UserID & "$" & rootid & "$" & BoardID
LastPost=replace(LastPost,"'","")

conn.execute("update topic set LastPost='"&LastPost&"' where topicid="&rootid)

LastPost_1=replace(username,"$","") & "$" & Announceid & "$" & DateTimeStr & "$" & replace(cutStr(reUBBCode(Topic,false),20),"$","") & "$" & uploadpic_n & "$" & UserID & "$" & rootid & "$" & BoardID
LastPost_1=replace(LastPost_1,"'","")

if isaudit=0 then
Dim UpdateBoardID
UpdateBoardID=BoardParentStr & "," & BoardID
if datediff("d",LastPostTime,Now())=0 then
	sql="update board set lastbbsnum=lastbbsnum+1,lasttopicnum=lasttopicnum+1,todaynum=todaynum+1,LastPost='"&LastPost_1&"' where boardid in ("&UpdateBoardID&")"
else
	sql="update board set lastbbsnum=lastbbsnum+1,lasttopicnum=lasttopicnum+1,todaynum=1,LastPost='"&LastPost_1&"' where  boardid in ("&UpdateBoardID&")"
end if
conn.execute(sql)
conn.execute("update [user] set userwealth=userwealth-"&(msgusemoney+mailusemoney)&" where username='"&membername&"'")

Dim updateinfo
set rs=conn.execute("select LastPost,TodayNum,MaxPostNum from config where active=1")
LastPostTimes=split(rs(0),"$")
LastPostTime=LastPostTimes(2)
if not isdate(LastPostTime) then LastPostTime=Now()
if datediff("d",LastPostTime,Now())=0 then
        mess="<font color=#cc0000>"&jhname&"</font><font color=006600>�ڽ�����̳������������<img src=img/j.gif><a href=../bbs/dispbbs.asp?boardid="&boardid&"&id="&rootid&" target=_blank title=�����ɲ鿴������><b><u>��"&topic&"��</u></b></a><img src=img/jj.gif>��ϣ�����ֱ�ӵ��<a href=../bbs/dispbbs.asp?boardid="&boardid&"&id="&rootid&" target=_blank title=�����ɲ鿴������><b>����</b></a>ȥ������</font>"
	if rs(1)+1>rs(2) then updateinfo=",MaxPostNum=todaynum+1"
	sql="update config set topicnum=topicnum+1,bbsnum=bbsnum+1,todayNum=todayNum+1,LastPost='"&LastPost&"' "&updateinfo&" where active=1"
else
	sql="update config set topicnum=topicnum+1,bbsnum=bbsnum+1,yesterdaynum="&rs(1)&",todayNum=1,LastPost='"&LastPost&"' where active=1"
end if
conn.execute(sql)
end if

if lcase(trim(username))<>lcase(trim(msgNotify)) and trim(msgNotify)<>"" and Board_Setting(48)=1 then
	dim rs2,sql2,boardname,msgcontent
	boardname=conn.execute("select boardtype from board where boardid="&boardid)(0)
	msgcontent=msgNotify&"����ã�"&chr(10)&"��������[b]"&now&"[/b]��[color=red]"&boardname&"[/color]���淢����һƪ��Ϊ[color=red]��"&topic&"��[/color]�����£���ϣ�������ȥ������"&chr(10)&"[align=center][URL=dispbbs.asp?boardid="&boardid&"&id="&rootid&"][color=blue][b][u]�鿴��������[/u][/b][/color][/URL][/align]"&chr(10)&"[align=right]"&username&"[/align]"
	set rs2=server.createobject("adodb.recordset")
	sql2="select * from [message]"      
	rs2.open sql2,conn,1,3      
	rs2.addnew      
	rs2("sender")=username
	rs2("incept")=msgNotify
	rs2("title")="��ע��"
	rs2("content")=msgcontent
	rs2("flag")=0 
	rs2("issend")=1
	rs2.update
	rs2.close
	set rs2=nothing
	conn.execute("update [user] set userwealth=userwealth-"&notifyusemoney&" where username='"&membername&"'")
end if
response.cookies("LastView")("Topic_"&RootID&"_"&BoardID)=now()

call nav()
call head_var(1,BoardDepth,0,0)
dim PostRetrunName
select case Board_Setting(17)
case 1
	response.write "<meta http-equiv=refresh content=""5;URL=index.asp"">"
	PostRetrunName="��ҳ"
case 2
	response.write "<meta http-equiv=refresh content=""5;URL=list.asp?boardid="&boardid&""">"
	PostRetrunName="������������̳"
case 3
	if isaudit=1 then
	response.write "<meta http-equiv=refresh content=""5;URL=list.asp?boardid="&boardid&""">"
	PostRetrunName="�����������ӱ��뾭����Ա��˺󷽿ɼ�"
	else
	response.write "<meta http-equiv=refresh content=""5;URL=dispbbs.asp?boardid="&boardid&"&id="&rootid&""">"
	PostRetrunName="�������������"
	end if
end select
set rs=nothing

%>
<table cellpadding=3 cellspacing=1 align=center class=tableborder1>
<tr align=center><th width="100%">״̬��<%=stats%></th>
</tr><tr><td width="100%" class=tablebody1>
��ҳ�潫��5����Զ�����<%=PostRetrunName%>��<b>������ѡ�����²�����</b><br><ul>
<li><a href="index.asp">������ҳ</a></li>
<li><a href="list.asp?boardid=<%=boardid%>"><%=boardtype%></a></li>
<li><%if isaudit=1 then%><%=PostRetrunName%><%else%><a href="dispbbs.asp?boardid=<%=boardid%>&id=<%=rootid%>&star=<%=request("star")%>#<%=announceid%>"><%=PostRetrunName%></a><%end if%></li>
</ul></td></tr></table>
<%
if ArticleRandom="yes" then
	call eddulf()
end if
end sub
%>
<!--#include file="z_eddul.asp"-->