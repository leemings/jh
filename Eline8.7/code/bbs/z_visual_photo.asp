<!--#include file="conn.asp"-->
<!--#include file="inc/const.asp" -->
<!--#include file="inc/chkinput.asp"-->
<!--#include file="z_visual_const.asp"-->
<%
response.buffer=true
dim rsVisual,sqlvisual
if not founduser or membername=""  then
	errmsg=errmsg+"<br>"+"<li>您还没有<a href=login.asp>登录论坛</a>，不能使用个人形象设计功能。如果您还没有<a href=reg.asp>注册</a>，请先<a href=reg.asp>注册</a>！"
	founderr=true
end if
stats="照相馆"
call nav()
call head_var(2,0,"","")
if founderr then
	call dvbbs_error()
else
	set rsVisual=server.createobject("adodb.recordset")
	select case request("action")
	case "send"
		call savemsg()
	case "confirm"
		call confirmmsg()
	case "design"
		call designmsg()
	case "accept"
		call acceptmsg()
	case "delete"
		call deletemsg()
	case "recreate"
		call recreatemsg()
	case "photome"
		call photome()
	case "privatephoto"
		call privatephoto()
	case "publicphoto"
		call publicphoto()
	case else
  	errmsg=errmsg+"<br>"+"<li>请指定正确的参数。"
		founderr=true
	end select
	if founderr then call dvbbs_error()
	set rsVisual=nothing
end if
call activeonline()
call footer()

sub savemsg()
	dim CurVisual,CurVisualSplit,BgWidth,BgHeight,PhotoStatus,PhotoBg,PhotoBodyBack,PhotoBodyFore,PhotoFg,PhotoID
	dim PhotoPrice1,PhotoPrice2,PhotoPeriod,PhotoUserCount,isUpload
	sqlvisual="select top 1 photoprice1,photoprice2,photoperiod,photousercount from visual_config"
	rsvisual.open sqlvisual,conn,1,1
	if not (rsvisual.eof or rsvisual.bof) then
		PhotoPrice1=rsvisual(0)
		PhotoPrice2=rsvisual(1)
		PhotoPeriod=rsvisual(2)
		PhotoUserCount=rsvisual(3)
	else
		PhotoPrice1=20
		PhotoPrice2=10
		PhotoPeriod=5
		PhotoUserCount=5
	end if
	rsvisual.close
	
	stats="发送请求成功"
	if isnull(request.form("fromvisual")) or request.form("fromvisual")="" then
		CurVisual=v_MyVisual
	else
		CurVisual=checkstr(trim(request.form("fromvisual")))
	end if
	CurVisualSplit=split(CurVisual,"|")
	if CurVisual="" or ubound(CurVisualSplit)<>24 then
		errmsg=errmsg+"<br>"+"<li>提交的数据有错误(1)！"
		founderr=true
		exit sub
	end if
	if isnull(request.form("bgwidth")) or request.form("bgwidth")="" then
		errmsg=errmsg+"<br>"+"<li>提交的数据有错误(2)！"
		founderr=true
		exit sub
	else
		BgWidth=checkstr(trim(request.form("bgwidth")))
		if not isInteger(BgWidth) then
			errmsg=errmsg+"<br>"+"<li>提交的数据有错误(2)！"
			founderr=true
			exit sub
		end if
	end if
	if isnull(request.form("bgheight")) or request.form("bgheight")="" then
		errmsg=errmsg+"<br>"+"<li>提交的数据有错误(3)！"
		founderr=true
		exit sub
	else
		BgHeight=checkstr(trim(request.form("bgheight")))
		if not isInteger(BgHeight) then
			errmsg=errmsg+"<br>"+"<li>提交的数据有错误(3)！"
			founderr=true
			exit sub
		end if
	end if
	if not isCanShowVisual(CurVisualSplit) then
		errmsg=errmsg+"<br>"+"<li>形象不雅，不能参加照相！"
		founderr=true
		exit sub
	end if
	dim incept,photoname,message,sendfee
	if isnull(request.form("photoname")) or request.form("photoname")="" then
		errmsg=errmsg+"<br>"+"<li>您还没有填写合影的名称呀。"
		founderr=true
		exit sub
	elseif strlength(CheckStr(trim(request.form("photoname"))))>50 then
		errmsg=errmsg+"<br>"+"<li>合影的名称限定最多50个字符。"
		founderr=true
		exit sub
	else
		photoname=CheckStr(trim(request.form("photoname")))
	end if
	if request.form("photomyself")<>"1" then
		if isnull(request.form("photomsg")) or request.form("photomsg")="" then
			errmsg=errmsg+"<br>"+"<li>请求的内容是必须要填写的噢。"
			founderr=true
			exit sub
		elseif strlength(CheckStr(trim(request.form("photomsg"))))>Cint(GroupSetting(34)) then
			errmsg=errmsg+"<br>"+"<li>请求的内容限定最多"&GroupSetting(34)&"个字符。"
			founderr=true
			exit sub
		else
			message=CheckStr(trim(request.form("photomsg")))
		end if
		if isnull(request.form("photouser")) or request.form("photouser")="" then
			errmsg=errmsg+"<br>"+"<li>您忘记填写参与者吧。"
			founderr=true
			exit sub
		else
			incept=CheckStr(trim(request.form("photouser")))
			if instr(","&incept&",",","&membername&",")>0 then
				errmsg=errmsg+"<br>"+"<li>不能与自己合影。"
				founderr=true
				exit sub
			end if
			incept=split(incept,",")
		end if
	else
		message=""
	end if
	if isnull(request.form("photostatus")) or request.form("photostatus")="" then
		errmsg=errmsg+"<br>"+"<li>您忘记指定照片的状态了。"
		founderr=true
		exit sub
	else
		photostatus=CheckStr(trim(request.form("photostatus")))
		if not isInteger(photostatus) then
			errmsg=errmsg+"<br>"+"<li>提交的数据有错误(4)！"
			founderr=true
			exit sub
		end if
	end if
	if isnull(request.form("photobg")) then
		errmsg=errmsg+"<br>"+"<li>提交的数据有错误(5)！"
		founderr=true
		exit sub
	else
		photobg=CheckStr(trim(request.form("photobg")))
		if photobg="0" then photobg=GetBackground(CurVisual)
		if photobg="1" then
			if isnull(request.form("bgfname")) then
				errmsg=errmsg+"<br>"+"<li>提交的数据有错误(6)！"
				founderr=true
				exit sub
			else
				isUpload=true
				photobg=CheckStr(trim(request.form("bgfname")))
				if photobg="" then isUpload=false
			end if
		else
			isUpload=false
		end if
	end if
	if isnull(request.form("photobodyback")) then
		errmsg=errmsg+"<br>"+"<li>提交的数据有错误(7)！"
		founderr=true
		exit sub
	else
		photobodyback=CheckStr(trim(request.form("photobodyback")))
	end if
	if isnull(request.form("photobodyfore")) then
		errmsg=errmsg+"<br>"+"<li>提交的数据有错误(8)！"
		founderr=true
		exit sub
	else
		photobodyfore=CheckStr(trim(request.form("photobodyfore")))
	end if
	if isnull(request.form("photofg")) then
		errmsg=errmsg+"<br>"+"<li>提交的数据有错误(9)！"
		founderr=true
		exit sub
	else
		photofg=CheckStr(trim(request.form("photofg")))
	end if
	dim usercount,usersplit(),j,ishas
	redim usersplit(0)
	usersplit(0)=membername
	usercount=1
	if request.form("photomyself")<>"1" then
		for i=0 to ubound(incept)
			if usercount>photousercount-1 and not (master and request.form("forcephoto")="1") then
				errmsg=errmsg+"<br>"+"<li>最多允许"&photousercount&"个用户合影，用户“"&replace(incept(i),"'","")&"”将被自动剔除。"
			else
				ishas=false
				for j=0 to usercount-1
					if lcase(trim(incept(i)))=lcase(trim(usersplit(j))) then
						ishas=true
						exit for
					end if
				next
				if not ishas then
					sql="select username from [user] where username='"&replace(incept(i),"'","")&"'"
					set rs=server.createobject("ADODB.Recordset")
					rs.open sql,conn,1,1
					if rs.eof and rs.bof then
						errmsg=errmsg+"<br>"+"<li>论坛没有用户“"&replace(incept(i),"'","")&"”，看看你的请求对象写对了嘛？"
						rs.close
					else
						rs.close
						sql="select top 1 F_id from [Friend] where F_username='"&replace(incept(i),"'","")&"' and F_friend='"&membername&"' and F_type=1"
						rs.open sql,conn,1,1
						if not (rs.eof and rs.bof) then
							errmsg=errmsg+"<br>"+"<li>您已经被"&incept(i)&"屏蔽了，不能邀请他合影！"
						else
							redim preserve usersplit(usercount)
							usersplit(usercount)=incept(i)
							usercount=usercount+1
						end if
						rs.close
					end if
					set rs=nothing
				end if
			end if
		next
		if usercount<2 then
			founderr=true
			exit sub
		end if
	end if
	if isnull(v_myvip) or v_myvip<>1 then
		sendfee=usercount*photoprice1
	else
		sendfee=usercount*photoprice2
	end if
	if mymoney<sendfee then
		errmsg=errmsg+"<br>"+"<li>您的现金不够发起这次照相。"
		founderr=true
		exit sub
	end if
	if (master and request.form("forcephoto")="1") or request.form("photomyself")="1" then
		sql="insert into visual_photo (name,content,fromuser,status,background,backbody,forebody,foreground,width,height,child,adddate,enddate,isUpload,confirmed) values ('"&photoname&"','"&message&"','"&membername&"',"&photostatus&",'"&photobg&"','"&photobodyback&"','"&photobodyfore&"','"&photofg&"',"&bgwidth&","&bgheight&","&usercount&",now(),dateadd('d',"&photoperiod&",now()),"&isUpload&",true)"
	else
		sql="insert into visual_photo (name,content,fromuser,status,background,backbody,forebody,foreground,width,height,child,adddate,enddate,isUpload) values ('"&photoname&"','"&message&"','"&membername&"',"&photostatus&",'"&photobg&"','"&photobodyback&"','"&photobodyfore&"','"&photofg&"',"&bgwidth&","&bgheight&","&usercount&",now(),dateadd('d',"&photoperiod&",now()),"&isUpload&")"
	end if
	conn.execute(sql)
	set rs=conn.execute("select max(id) from visual_photo")
	photoid=rs(0)
	if isnull(v_myvip) or v_myvip<>1 then
		conn.execute("update [user] set userwealth=userwealth-"&photoprice1&" where username='"&membername&"'")
	else
		conn.execute("update [user] set userwealth=userwealth-"&photoprice2&" where username='"&membername&"'")
	end if
	sql="insert into visual_photouser (photo_id,username,uservisual,confirmed,confirmtime) values ("&photoid&",'"&membername&"','"&CurVisual&"',1,now())"
	conn.execute(sql)
	if request.form("photomyself")<>"1" then
		dim rsUser,CurUserVisual
		set rsUser=server.createobject("adodb.recordset")
		for i=1 to usercount-1
			if isnull(v_myvip) or v_myvip<>1 then
				conn.execute("update [user] set userwealth=userwealth-"&photoprice1&" where username='"&membername&"'")
			else
				conn.execute("update [user] set userwealth=userwealth-"&photoprice2&" where username='"&membername&"'")
			end if
			if master and request.form("forcephoto")="1" then
				rsUser.open "select visual,sex from [user] where username='"&usersplit(i)&"'",conn,1,1
				if isnull(rsUser(0)) or rsUser(0)="" then
					if rsUser(1)=1 then
						CurUserVisual=DefaultVisualBoy
					else
						CurUserVisual=DefaultVisualGirl
					end if
				else
					if ubound(split(rsUser(0),"|"))<>24 then
						if rsUser(1)=1 then
							CurUserVisual=DefaultVisualBoy
						else
							CurUserVisual=DefaultVisualGirl
						end if
					else
						CurUserVisual=rsUser(0)
					end if
				end if
				sql="insert into visual_photouser (photo_id,username,uservisual,confirmed,confirmtime) values ("&photoid&",'"&usersplit(i)&"','"&CurUserVisual&"',1,now())"
				rsUser.Close
			else
				sql="insert into visual_photouser (photo_id,username) values ("&photoid&",'"&usersplit(i)&"')"
			end if
			conn.execute(sql)
			if not (master and request.form("forcephoto")="1") then
				sql="insert into message (incept,sender,title,content,sendtime,flag,issend) values ('"&usersplit(i)&"','"&membername&"','邀请您参加《"&photoname&"》的合影','您好："&chr(10)&"　　冒昧地打扰您，可以和您一起照张像吗？"&chr(10)& "　　请到以下地址确认这个请求："&chr(10)&chr(10)&chr(10)&"[align=center][URL=z_Visual.asp?shopid=397&photoid="&photoid&"][color=blue][b][u]确认合影请求[/u][/b][/color][/URL][/align]',Now(),0,1)"
				conn.execute(sql)
			end if
			usercount=usercount+1
		next
		set rsUser=nothing
	end if
	if request.form("photomyself")<>"1" then
		sucmsg=errmsg+"<br>"+"<li><b>恭喜您，发送照相请求成功。</b>"
		call dvbbs_suc()
	else
		response.redirect "z_visual.asp?shopid=497&photoid="&photoid
	end if
end sub

sub photome()
	dim CurVisual,CurVisualSplit,BgWidth,BgHeight,PhotoStatus,PhotoBg,PhotoID
	dim PhotoPrice1,PhotoPrice2,PhotoPeriod,PhotoUserCount
	sqlvisual="select top 1 photoprice1,photoprice2,photoperiod,photousercount from visual_config"
	rsvisual.open sqlvisual,conn,1,1
	if not (rsvisual.eof or rsvisual.bof) then
		PhotoPrice1=rsvisual(0)
		PhotoPrice2=rsvisual(1)
		PhotoPeriod=rsvisual(2)
		PhotoUserCount=rsvisual(3)
	else
		PhotoPrice1=20
		PhotoPrice2=10
		PhotoPeriod=5
		PhotoUserCount=5
	end if
	rsvisual.close
	
	if isnull(request.cookies("myshow_"&userid)) or request.cookies("myshow_"&userid)="" then
		CurVisual=v_MyVisual
	else
		CurVisual=request.cookies("myshow_"&userid)
	end if
	CurVisualSplit=split(CurVisual,"|")
	if CurVisual="" or ubound(CurVisualSplit)<>24 then
		errmsg=errmsg+"<br>"+"<li>提交的数据有错误(1)！"
		founderr=true
		exit sub
	end if
	bgwidth=140
	bgHeight=226
	if not isCanShowVisual(CurVisualSplit) then
		errmsg=errmsg+"<br>"+"<li>形象不雅，不能进行个人写真！"
		founderr=true
		exit sub
	end if
	dim incept,photoname,message,sendfee
	photoname=membername&"的个人写真"
	message=""
	photostatus=DefaultStatus
	photobg=GetBackground(CurVisual)
	dim usercount,usersplit()
	redim usersplit(0)
	usersplit(0)=membername
	usercount=1
	if isnull(v_myvip) or v_myvip<>1 then
		sendfee=usercount*photoprice1
	else
		sendfee=usercount*photoprice2
	end if
	if mymoney<sendfee then
		errmsg=errmsg+"<br>"+"<li>您的现金不够完成这次个人写真。"
		founderr=true
		exit sub
	end if
	sql="insert into visual_photo (name,content,fromuser,status,background,width,height,child,adddate,enddate,confirmed) values ('"&photoname&"','"&message&"','"&membername&"',"&photostatus&",'"&photobg&"',"&bgwidth&","&bgheight&","&usercount&",now(),dateadd('d',"&photoperiod&",now()),true)"
	conn.execute(sql)
	set rs=conn.execute("select max(id) from visual_photo")
	photoid=rs(0)
	if isnull(v_myvip) or v_myvip<>1 then
		conn.execute("update [user] set userwealth=userwealth-"&photoprice1&" where username='"&membername&"'")
	else
		conn.execute("update [user] set userwealth=userwealth-"&photoprice2&" where username='"&membername&"'")
	end if
	sql="insert into visual_photouser (photo_id,username,uservisual,confirmed,confirmtime) values ("&photoid&",'"&membername&"','"&CurVisual&"',1,now())"
	conn.execute(sql)
	response.redirect "z_visual.asp?shopid=497&photoid="&photoid
end sub

sub confirmmsg()
	dim CurVisual,CurVisualSplit,PhotoID
	
	stats="确认请求成功"
	if isnull(request("photoid")) then
		errmsg=errmsg+"<br>"+"<li>提交的数据有错误！"
		founderr=true
		exit sub
	elseif not isInteger(request("photoid")) then
		errmsg=errmsg+"<br>"+"<li>提交的数据有错误！"
		founderr=true
		exit sub
	elseif conn.execute("select u.photo_id from visual_photouser u inner join visual_photo p on u.photo_id=p.id where u.photo_id="&checkstr(trim(request("photoid")))&" and u.username='"&membername&"' and not p.confirmed").eof then
		errmsg=errmsg+"<br>"+"<li>提交的数据有错误！"
		founderr=true
		exit sub
	else
		photoid=checkstr(trim(request("photoid")))
	end if
	if isnull(request.form("uservisual")) or request.form("uservisual")="" then
		CurVisual=v_MyVisual
	else
		CurVisual=checkstr(trim(request.form("uservisual")))
	end if
	CurVisualSplit=split(CurVisual,"|")
	if CurVisual="" or ubound(CurVisualSplit)<>24 then
		errmsg=errmsg+"<br>"+"<li>提交的数据有错误！"
		founderr=true
		exit sub
	end if
	if not isCanShowVisual(CurVisualSplit) then
		errmsg=errmsg+"<br>"+"<li>形象不雅或没有装备任何形象物品，不能参加合影！"
		founderr=true
		exit sub
	end if
	conn.execute("update visual_photouser set uservisual='"&curvisual&"',confirmed=true,confirmtime=now() where photo_id="&photoid&" and username='"&membername&"'")
	if conn.execute("select photo_id from visual_photouser where photo_id="&photoid&" and not confirmed").eof then
		set rs=server.createobject("adodb.recordset")
		sql="select fromuser,confirmed,name,addDate from visual_photo where id="&photoid
		rs.open sql,conn,1,3
		rs(1)=true
		rs.update
		sql="insert into message (incept,sender,title,content,sendtime,flag,issend) values ('"&_
			rs(0)&"','"&forum_info(0)&"','您发起的合影《"&rs(2)&"》已经被所有人确认了','您好："&_
			chr(10)&"　　很高兴地通知您，您于 [b]"&formatdatetime(rs(3),2)&"[/b] 发起的合影已经被所有的参与者确认了。"&_
			chr(10)&"　　请到以下地址设计这个合影："&chr(10)&chr(10)&chr(10)&"[align=center][URL=z_Visual.asp?shopid=497&photoid="&_
			photoid&"][color=blue][b][u]设计合影[/u][/b][/color][/URL][/align]',Now(),0,1)"
		conn.execute(sql)
		rs.close
		set rs=nothing
	end if
	sucmsg="<br>"+"<li><b>恭喜您，确认合影请求成功。</b>"
	call dvbbs_suc()
end sub

sub designmsg()
	dim PhotoID,UserCount,rsPhoto,rsPhotoUser
	dim PhotoNameLeft,PhotoNameTop,PhotoNameFont,PhotoNameSize,PhotoNameBold,PhotoNameItalic,PhotoNameColor,PhotoNameDirection
	dim DateLeft,DateTop,DateFont,DateSize,DateBold,DateItalic,DateColor,DateDirection
	dim NameLeft(),NameTop(),NameFont(),NameSize(),NameBold(),NameItalic(),NameColor(),NameDirection()
	dim OuterLeft(),OuterTop(),OuterWidth(),OuterHeight()
	dim InnerLeft(),InnerTop(),InnerWidth(),InnerHeight()
	dim LayerNo(),Direction(),UserVisual()

	stats="合影设计成功"
	if isnull(request("photoid")) then
		errmsg=errmsg+"<br>"+"<li>提交的数据有错误(1)！"
		founderr=true
		exit sub
	elseif not isInteger(request("photoid")) then
		errmsg=errmsg+"<br>"+"<li>提交的数据有错误(1)！"
		founderr=true
		exit sub
	elseif conn.execute("select u.photo_id from visual_photouser u inner join visual_photo p on u.photo_id=p.id where u.photo_id="&checkstr(trim(request("photoid")))&" and u.username='"&membername&"' and p.confirmed and not p.finished").eof then
		errmsg=errmsg+"<br>"+"<li>提交的数据有错误(1)！"
		founderr=true
		exit sub
	else
		photoid=checkstr(trim(request("photoid")))
	end if
	if isnull(request.form("PhotoNameLeft")) or request.form("PhotoNameLeft")="" then
		errmsg=errmsg+"<br>"+"<li>提交的数据有错误(2)！"
		founderr=true
		exit sub
	elseif not isNumeric(request("PhotoNameLeft")) then
		errmsg=errmsg+"<br>"+"<li>提交的数据有错误(2)！"
		founderr=true
		exit sub
	else
		PhotoNameLeft=CInt(checkstr(trim(request("PhotoNameLeft"))))
	end if
	if isnull(request.form("PhotoNameTop")) or request.form("PhotoNameTop")="" then
		errmsg=errmsg+"<br>"+"<li>提交的数据有错误(3)！"
		founderr=true
		exit sub
	elseif not isNumeric(request("PhotoNameTop")) then
		errmsg=errmsg+"<br>"+"<li>提交的数据有错误(3)！"
		founderr=true
		exit sub
	else
		PhotoNameTop=CInt(checkstr(trim(request("PhotoNameTop"))))
	end if
	if isnull(request.form("PhotoNameFont")) or request.form("PhotoNameFont")="" then
		errmsg=errmsg+"<br>"+"<li>提交的数据有错误(4)！"
		founderr=true
		exit sub
	else
		PhotoNameFont=checkstr(trim(request("PhotoNameFont")))
	end if
	if isnull(request.form("PhotoNameSize")) or request.form("PhotoNameSize")="" then
		errmsg=errmsg+"<br>"+"<li>提交的数据有错误(5)！"
		founderr=true
		exit sub
	else
		PhotoNameSize=checkstr(trim(request("PhotoNameSize")))
	end if
	if isnull(request.form("PhotoNameBold")) or request.form("PhotoNameBold")="" then
		errmsg=errmsg+"<br>"+"<li>提交的数据有错误(6)！"
		founderr=true
		exit sub
	else
		PhotoNameBold=(cint(checkstr(trim(request("PhotoNameBold"))))=1)
	end if
	if isnull(request.form("PhotoNameItalic")) or request.form("PhotoNameItalic")="" then
		errmsg=errmsg+"<br>"+"<li>提交的数据有错误(7)！"
		founderr=true
		exit sub
	else
		PhotoNameItalic=(cint(checkstr(trim(request("PhotoNameItalic"))))=1)
	end if
	if isnull(request.form("PhotoNameColor")) then
		errmsg=errmsg+"<br>"+"<li>提交的数据有错误(8)！"
		founderr=true
		exit sub
	else
		PhotoNameColor=checkstr(trim(request("PhotoNameColor")))
	end if
	if isnull(request.form("PhotoNameDirection")) or request.form("PhotoNameDirection")="" then
		errmsg=errmsg+"<br>"+"<li>提交的数据有错误(9)！"
		founderr=true
		exit sub
	elseif not isInteger(request("PhotoNameDirection")) then
		errmsg=errmsg+"<br>"+"<li>提交的数据有错误(9)！"
		founderr=true
		exit sub
	else
		PhotoNameDirection=CInt(checkstr(trim(request("PhotoNameDirection"))))
	end if
	if isnull(request.form("DateLeft")) or request.form("DateLeft")="" then
		errmsg=errmsg+"<br>"+"<li>提交的数据有错误(10)！"
		founderr=true
		exit sub
	elseif not isNumeric(request("DateLeft")) then
		errmsg=errmsg+"<br>"+"<li>提交的数据有错误(10)！"
		founderr=true
		exit sub
	else
		DateLeft=CInt(checkstr(trim(request("DateLeft"))))
	end if
	if isnull(request.form("DateTop")) or request.form("DateTop")="" then
		errmsg=errmsg+"<br>"+"<li>提交的数据有错误(11)！"
		founderr=true
		exit sub
	elseif not isNumeric(request("DateTop")) then
		errmsg=errmsg+"<br>"+"<li>提交的数据有错误(11)！"
		founderr=true
		exit sub
	else
		DateTop=CInt(checkstr(trim(request("DateTop"))))
	end if
	if isnull(request.form("DateFont")) or request.form("DateFont")="" then
		errmsg=errmsg+"<br>"+"<li>提交的数据有错误(12)！"
		founderr=true
		exit sub
	else
		DateFont=checkstr(trim(request("DateFont")))
	end if
	if isnull(request.form("DateSize")) or request.form("DateSize")="" then
		errmsg=errmsg+"<br>"+"<li>提交的数据有错误(13)！"
		founderr=true
		exit sub
	else
		DateSize=checkstr(trim(request("DateSize")))
	end if
	if isnull(request.form("DateBold")) or request.form("DateBold")="" then
		errmsg=errmsg+"<br>"+"<li>提交的数据有错误(14)！"
		founderr=true
		exit sub
	else
		DateBold=(cint(checkstr(trim(request("DateBold"))))=1)
	end if
	if isnull(request.form("DateItalic")) or request.form("DateItalic")="" then
		errmsg=errmsg+"<br>"+"<li>提交的数据有错误(15)！"
		founderr=true
		exit sub
	else
		DateItalic=(cint(checkstr(trim(request("DateItalic"))))=1)
	end if
	if isnull(request.form("DateColor")) then
		errmsg=errmsg+"<br>"+"<li>提交的数据有错误(16)！"
		founderr=true
		exit sub
	else
		DateColor=checkstr(trim(request("DateColor")))
	end if
	if isnull(request.form("DateDirection")) or request.form("DateDirection")="" then
		errmsg=errmsg+"<br>"+"<li>提交的数据有错误(17)！"
		founderr=true
		exit sub
	elseif not isInteger(request("DateDirection")) then
		errmsg=errmsg+"<br>"+"<li>提交的数据有错误(17)！"
		founderr=true
		exit sub
	else
		DateDirection=CInt(checkstr(trim(request("DateDirection"))))
	end if
	set rsPhoto=server.createobject("ADODB.recordset")
	sql="select * from visual_photo where id="&photoid
	rsPhoto.open sql,conn,1,3
	UserCount=rsPhoto("Child")
	redim NameLeft(UserCount-1),NameTop(UserCount-1),NameFont(UserCount-1),NameSize(UserCount-1),NameBold(UserCount-1),NameItalic(UserCount-1),NameColor(UserCount-1),NameDirection(UserCount-1)
	redim OuterLeft(UserCount+5),OuterTop(UserCount+5),OuterWidth(UserCount+5),OuterHeight(UserCount+5)
	redim InnerLeft(UserCount+5),InnerTop(UserCount+5),InnerWidth(UserCount+5),InnerHeight(UserCount+5)
	redim LayerNo(UserCount-1),Direction(UserCount+5),UserVisual(5)
	for i=0 to UserCount-1
		if isnull(request.form("LayerNo_"&i)) or request.form("LayerNo_"&i)="" then
			errmsg=errmsg+"<br>"+"<li>提交的数据有错误("&(18+i*18)&")！"
			founderr=true
			exit sub
		elseif not isInteger(request("LayerNo_"&i)) then
			errmsg=errmsg+"<br>"+"<li>提交的数据有错误("&(18+i*18)&")！"
			founderr=true
			exit sub
		else
			LayerNo(i)=CInt(checkstr(trim(request("LayerNo_"&i))))
		end if
		if isnull(request.form("OuterLeft_"&i)) or request.form("OuterLeft_"&i)="" then
			errmsg=errmsg+"<br>"+"<li>提交的数据有错误("&(19+i*18)&")！"
			founderr=true
			exit sub
		elseif not isNumeric(request("OuterLeft_"&i)) then
			errmsg=errmsg+"<br>"+"<li>提交的数据有错误("&(19+i*18)&")！"
			founderr=true
			exit sub
		else
			OuterLeft(i)=CInt(checkstr(trim(request("OuterLeft_"&i))))
		end if
		if isnull(request.form("OuterTop_"&i)) or request.form("OuterTop_"&i)="" then
			errmsg=errmsg+"<br>"+"<li>提交的数据有错误("&(20+i*18)&")！"
			founderr=true
			exit sub
		elseif not isNumeric(request("OuterTop_"&i)) then
			errmsg=errmsg+"<br>"+"<li>提交的数据有错误("&(20+i*18)&")！"
			founderr=true
			exit sub
		else
			OuterTop(i)=CInt(checkstr(trim(request("OuterTop_"&i))))
		end if
		if isnull(request.form("OuterWidth_"&i)) or request.form("OuterWidth_"&i)="" then
			errmsg=errmsg+"<br>"+"<li>提交的数据有错误("&(21+i*18)&")！"
			founderr=true
			exit sub
		elseif not isInteger(request("OuterWidth_"&i)) then
			errmsg=errmsg+"<br>"+"<li>提交的数据有错误("&(21+i*18)&")！"
			founderr=true
			exit sub
		else
			OuterWidth(i)=CInt(checkstr(trim(request("OuterWidth_"&i))))
		end if
		if isnull(request.form("OuterHeight_"&i)) or request.form("OuterHeight_"&i)="" then
			errmsg=errmsg+"<br>"+"<li>提交的数据有错误("&(22+i*18)&")！"
			founderr=true
			exit sub
		elseif not isInteger(request("OuterHeight_"&i)) then
			errmsg=errmsg+"<br>"+"<li>提交的数据有错误("&(22+i*18)&")！"
			founderr=true
			exit sub
		else
			OuterHeight(i)=CInt(checkstr(trim(request("OuterHeight_"&i))))
		end if
		if isnull(request.form("InnerLeft_"&i)) or request.form("InnerLeft_"&i)="" then
			errmsg=errmsg+"<br>"+"<li>提交的数据有错误("&(23+i*18)&")！"
			founderr=true
			exit sub
		elseif not isNumeric(request("InnerLeft_"&i)) then
			errmsg=errmsg+"<br>"+"<li>提交的数据有错误("&(23+i*18)&")！"
			founderr=true
			exit sub
		else
			InnerLeft(i)=CInt(checkstr(trim(request("InnerLeft_"&i))))
		end if
		if isnull(request.form("InnerTop_"&i)) or request.form("InnerTop_"&i)="" then
			errmsg=errmsg+"<br>"+"<li>提交的数据有错误("&(24+i*18)&")！"
			founderr=true
			exit sub
		elseif not isNumeric(request("InnerTop_"&i)) then
			errmsg=errmsg+"<br>"+"<li>提交的数据有错误("&(24+i*18)&")！"
			founderr=true
			exit sub
		else
			InnerTop(i)=CInt(checkstr(trim(request("InnerTop_"&i))))
		end if
		if isnull(request.form("InnerWidth_"&i)) or request.form("InnerWidth_"&i)="" then
			errmsg=errmsg+"<br>"+"<li>提交的数据有错误("&(25+i*18)&")！"
			founderr=true
			exit sub
		elseif not isInteger(request("InnerWidth_"&i)) then
			errmsg=errmsg+"<br>"+"<li>提交的数据有错误("&(25+i*18)&")！"
			founderr=true
			exit sub
		else
			InnerWidth(i)=CInt(checkstr(trim(request("InnerWidth_"&i))))
		end if
		if isnull(request.form("InnerHeight_"&i)) or request.form("InnerHeight_"&i)="" then
			errmsg=errmsg+"<br>"+"<li>提交的数据有错误("&(26+i*18)&")！"
			founderr=true
			exit sub
		elseif not isInteger(request("InnerHeight_"&i)) then
			errmsg=errmsg+"<br>"+"<li>提交的数据有错误("&(26+i*18)&")！"
			founderr=true
			exit sub
		else
			InnerHeight(i)=CInt(checkstr(trim(request("InnerHeight_"&i))))
		end if
		if isnull(request.form("Direction_"&i)) or request.form("Direction_"&i)="" then
			errmsg=errmsg+"<br>"+"<li>提交的数据有错误("&(27+i*18)&")！"
			founderr=true
			exit sub
		elseif not isInteger(request("Direction_"&i)) then
			errmsg=errmsg+"<br>"+"<li>提交的数据有错误("&(27+i*18)&")！"
			founderr=true
			exit sub
		else
			Direction(i)=CInt(checkstr(trim(request("Direction_"&i))))
		end if
		if isnull(request.form("NameLeft_"&i)) or request.form("NameLeft_"&i)="" then
			errmsg=errmsg+"<br>"+"<li>提交的数据有错误("&(28+i*18)&")！"
			founderr=true
			exit sub
		elseif not isNumeric(request("NameLeft_"&i)) then
			errmsg=errmsg+"<br>"+"<li>提交的数据有错误("&(28+i*18)&")！"
			founderr=true
			exit sub
		else
			NameLeft(i)=CInt(checkstr(trim(request("NameLeft_"&i))))
		end if
		if isnull(request.form("NameTop_"&i)) or request.form("NameTop_"&i)="" then
			errmsg=errmsg+"<br>"+"<li>提交的数据有错误("&(29+i*18)&")！"
			founderr=true
			exit sub
		elseif not isNumeric(request("NameTop_"&i)) then
			errmsg=errmsg+"<br>"+"<li>提交的数据有错误("&(29+i*18)&")！"
			founderr=true
			exit sub
		else
			NameTop(i)=CInt(checkstr(trim(request("NameTop_"&i))))
		end if
		if isnull(request.form("NameFont_"&i)) or request.form("NameFont_"&i)="" then
			errmsg=errmsg+"<br>"+"<li>提交的数据有错误("&(30+i*18)&")！"
			founderr=true
			exit sub
		else
			NameFont(i)=checkstr(trim(request("NameFont_"&i)))
		end if
		if isnull(request.form("NameSize_"&i)) or request.form("NameSize_"&i)="" then
			errmsg=errmsg+"<br>"+"<li>提交的数据有错误("&(31+i*18)&")！"
			founderr=true
			exit sub
		else
			NameSize(i)=checkstr(trim(request("NameSize_"&i)))
		end if
		if isnull(request.form("NameBold_"&i)) or request.form("NameBold_"&i)="" then
			errmsg=errmsg+"<br>"+"<li>提交的数据有错误("&(32+i*18)&")！"
			founderr=true
			exit sub
		else
			NameBold(i)=(cint(checkstr(trim(request("NameBold_"&i))))=1)
		end if
		if isnull(request.form("NameItalic_"&i)) or request.form("NameItalic_"&i)="" then
			errmsg=errmsg+"<br>"+"<li>提交的数据有错误("&(33+i*18)&")！"
			founderr=true
			exit sub
		else
			NameItalic(i)=(cint(checkstr(trim(request("NameItalic_"&i))))=1)
		end if
		if isnull(request.form("NameColor_"&i)) then
			errmsg=errmsg+"<br>"+"<li>提交的数据有错误("&(34+i*18)&")！"
			founderr=true
			exit sub
		else
			NameColor(i)=checkstr(trim(request("NameColor_"&i)))
		end if
		if isnull(request.form("NameDirection_"&i)) or request.form("NameDirection_"&i)="" then
			errmsg=errmsg+"<br>"+"<li>提交的数据有错误("&(35+i*18)&")！"
			founderr=true
			exit sub
		elseif not isInteger(request("NameDirection_"&i)) then
			errmsg=errmsg+"<br>"+"<li>提交的数据有错误("&(35+i*18)&")！"
			founderr=true
			exit sub
		else
			NameDirection(i)=CInt(checkstr(trim(request("NameDirection_"&i))))
		end if
	next
	for i=0 to 5
		if isnull(request.form("AccoutermentUserVisual_"&i)) then
			errmsg=errmsg+"<br>"+"<li>提交的数据有错误("&(18+UserCount*18+i*10)&")！"
			founderr=true
			exit sub
		else
			UserVisual(i)=checkstr(trim(request("AccoutermentUserVisual_"&i)))
		end if
		if isnull(request.form("AccoutermentOuterLeft_"&i)) or request.form("AccoutermentOuterLeft_"&i)="" then
			errmsg=errmsg+"<br>"+"<li>提交的数据有错误("&(19+UserCount*18+i*10)&")！"
			founderr=true
			exit sub
		elseif not isNumeric(request("AccoutermentOuterLeft_"&i)) then
			errmsg=errmsg+"<br>"+"<li>提交的数据有错误("&(19+UserCount*18+i*10)&")！"
			founderr=true
			exit sub
		else
			OuterLeft(UserCount+i)=CInt(checkstr(trim(request("AccoutermentOuterLeft_"&i))))
		end if
		if isnull(request.form("AccoutermentOuterTop_"&i)) or request.form("AccoutermentOuterTop_"&i)="" then
			errmsg=errmsg+"<br>"+"<li>提交的数据有错误("&(20+UserCount*18+i*10)&")！"
			founderr=true
			exit sub
		elseif not isNumeric(request("AccoutermentOuterTop_"&i)) then
			errmsg=errmsg+"<br>"+"<li>提交的数据有错误("&(20+UserCount*18+i*10)&")！"
			founderr=true
			exit sub
		else
			OuterTop(UserCount+i)=CInt(checkstr(trim(request("AccoutermentOuterTop_"&i))))
		end if
		if isnull(request.form("AccoutermentOuterWidth_"&i)) or request.form("AccoutermentOuterWidth_"&i)="" then
			errmsg=errmsg+"<br>"+"<li>提交的数据有错误("&(21+UserCount*18+i*10)&")！"
			founderr=true
			exit sub
		elseif not isInteger(request("AccoutermentOuterWidth_"&i)) then
			errmsg=errmsg+"<br>"+"<li>提交的数据有错误("&(21+UserCount*18+i*10)&")！"
			founderr=true
			exit sub
		else
			OuterWidth(UserCount+i)=CInt(checkstr(trim(request("AccoutermentOuterWidth_"&i))))
		end if
		if isnull(request.form("AccoutermentOuterHeight_"&i)) or request.form("AccoutermentOuterHeight_"&i)="" then
			errmsg=errmsg+"<br>"+"<li>提交的数据有错误("&(22+UserCount*18+i*10)&")！"
			founderr=true
			exit sub
		elseif not isInteger(request("AccoutermentOuterHeight_"&i)) then
			errmsg=errmsg+"<br>"+"<li>提交的数据有错误("&(22+UserCount*18+i*10)&")！"
			founderr=true
			exit sub
		else
			OuterHeight(UserCount+i)=CInt(checkstr(trim(request("AccoutermentOuterHeight_"&i))))
		end if
		if isnull(request.form("AccoutermentInnerLeft_"&i)) or request.form("AccoutermentInnerLeft_"&i)="" then
			errmsg=errmsg+"<br>"+"<li>提交的数据有错误("&(23+UserCount*18+i*10)&")！"
			founderr=true
			exit sub
		elseif not isNumeric(request("AccoutermentInnerLeft_"&i)) then
			errmsg=errmsg+"<br>"+"<li>提交的数据有错误("&(23+UserCount*18+i*10)&")！"
			founderr=true
			exit sub
		else
			InnerLeft(UserCount+i)=CInt(checkstr(trim(request("AccoutermentInnerLeft_"&i))))
		end if
		if isnull(request.form("AccoutermentInnerTop_"&i)) or request.form("AccoutermentInnerTop_"&i)="" then
			errmsg=errmsg+"<br>"+"<li>提交的数据有错误("&(24+UserCount*18+i*10)&")！"
			founderr=true
			exit sub
		elseif not isNumeric(request("AccoutermentInnerTop_"&i)) then
			errmsg=errmsg+"<br>"+"<li>提交的数据有错误("&(24+UserCount*18+i*10)&")！"
			founderr=true
			exit sub
		else
			InnerTop(UserCount+i)=CInt(checkstr(trim(request("AccoutermentInnerTop_"&i))))
		end if
		if isnull(request.form("AccoutermentInnerWidth_"&i)) or request.form("AccoutermentInnerWidth_"&i)="" then
			errmsg=errmsg+"<br>"+"<li>提交的数据有错误("&(25+UserCount*18+i*10)&")！"
			founderr=true
			exit sub
		elseif not isInteger(request("AccoutermentInnerWidth_"&i)) then
			errmsg=errmsg+"<br>"+"<li>提交的数据有错误("&(25+UserCount*18+i*10)&")！"
			founderr=true
			exit sub
		else
			InnerWidth(UserCount+i)=CInt(checkstr(trim(request("AccoutermentInnerWidth_"&i))))
		end if
		if isnull(request.form("AccoutermentInnerHeight_"&i)) or request.form("AccoutermentInnerHeight_"&i)="" then
			errmsg=errmsg+"<br>"+"<li>提交的数据有错误("&(26+UserCount*18+i*10)&")！"
			founderr=true
			exit sub
		elseif not isInteger(request("AccoutermentInnerHeight_"&i)) then
			errmsg=errmsg+"<br>"+"<li>提交的数据有错误("&(26+UserCount*18+i*10)&")！"
			founderr=true
			exit sub
		else
			InnerHeight(UserCount+i)=CInt(checkstr(trim(request("AccoutermentInnerHeight_"&i))))
		end if
		if isnull(request.form("AccoutermentDirection_"&i)) or request.form("AccoutermentDirection_"&i)="" then
			errmsg=errmsg+"<br>"+"<li>提交的数据有错误("&(27+UserCount*18+i*10)&")！"
			founderr=true
			exit sub
		elseif not isInteger(request("AccoutermentDirection_"&i)) then
			errmsg=errmsg+"<br>"+"<li>提交的数据有错误("&(27+UserCount*18+i*10)&")！"
			founderr=true
			exit sub
		else
			Direction(UserCount+i)=CInt(checkstr(trim(request("AccoutermentDirection_"&i))))
		end if
	next
	Dim Flag,hasFinished,photoname
	photoname=trim(rsphoto("name"))
	Flag=False
	hasFinished=(not isNull(rsPhoto("FinishDate")))
	if hasFinished then
		if rsPhoto("PhotoNameLeft")<>PhotoNameLeft or rsPhoto("PhotoNameTop")<>PhotoNameTop or rsPhoto("PhotoNameFont")<>PhotoNameFont or rsPhoto("PhotoNameSize")<>PhotoNameSize or rsPhoto("PhotoNameBold")<>PhotoNameBold or rsPhoto("PhotoNameItalic")<>PhotoNameItalic or rsPhoto("PhotoNameColor")<>PhotoNameColor or rsPhoto("PhotoNameDirection")<>PhotoNameDirection then Flag=True
		if rsPhoto("DateLeft")<>DateLeft or rsPhoto("DateTop")<>DateTop or rsPhoto("DateFont")<>DateFont or rsPhoto("DateSize")<>DateSize or rsPhoto("DateBold")<>DateBold or rsPhoto("DateItalic")<>DateItalic or rsPhoto("DateColor")<>DateColor or rsPhoto("DateDirection")<>DateDirection then Flag=True
	else
		Flag=True
	end if
	if Flag Then
		rsPhoto("PhotoNameLeft")=PhotoNameLeft
		rsPhoto("PhotoNameTop")=PhotoNameTop
		rsPhoto("PhotoNameFont")=PhotoNameFont
		rsPhoto("PhotoNameSize")=PhotoNameSize
		rsPhoto("PhotoNameBold")=PhotoNameBold
		rsPhoto("PhotoNameItalic")=PhotoNameItalic
		rsPhoto("PhotoNameColor")=PhotoNameColor
		rsPhoto("PhotoNameDirection")=PhotoNameDirection
		rsPhoto("DateLeft")=DateLeft
		rsPhoto("DateTop")=DateTop
		rsPhoto("DateFont")=DateFont
		rsPhoto("DateSize")=DateSize
		rsPhoto("DateBold")=DateBold
		rsPhoto("DateItalic")=DateItalic
		rsPhoto("DateColor")=DateColor
		rsPhoto("DateDirection")=DateDirection
	end if
	set rsPhotoUser=Server.CreateObject("ADODB.Recordset")
	sql="select * from Visual_Accouterment where photo_id="&photoid
	rsPhotoUser.open sql,conn,1,3
	dim Flag1
	for i=0 to 5
		if UserVisual(i)<>"" then
			if not rsPhotoUser.eof then
				Flag1=False
			else
				rsPhotoUser.Addnew
				rsPhotoUser("Photo_ID")=photoid
				Flag1=true
			end if
			if hasFinished then
				if rsPhotoUser("SeqNo")<>i or rsPhotoUser("ItemPicPath")<>UserVisual(i) or rsPhotoUser("OuterLeft")<>OuterLeft(UserCount+i) or rsPhotoUser("OuterTop")<>OuterTop(UserCount+i) or rsPhotoUser("OuterWidth")<>OuterWidth(UserCount+i) or rsPhotoUser("OuterHeight")<>OuterHeight(UserCount+i) or rsPhotoUser("InnerLeft")<>InnerLeft(UserCount+i) or rsPhotoUser("InnerTop")<>InnerTop(UserCount+i) or rsPhotoUser("InnerWidth")<>InnerWidth(UserCount+i) or rsPhotoUser("InnerHeight")<>InnerHeight(UserCount+i) or rsPhotoUser("Direction")<>Direction(UserCount+i) Then Flag1=True
			else
				Flag1=True
			end if
			if Flag1 Then
				rsPhotoUser("SeqNo")=i
				rsPhotoUser("ItemPicPath")=UserVisual(i)
				rsPhotoUser("OuterLeft")=OuterLeft(UserCount+i)
				rsPhotoUser("OuterTop")=OuterTop(UserCount+i)
				rsPhotoUser("OuterWidth")=OuterWidth(UserCount+i)
				rsPhotoUser("OuterHeight")=OuterHeight(UserCount+i)
				rsPhotoUser("InnerLeft")=InnerLeft(UserCount+i)
				rsPhotoUser("InnerTop")=InnerTop(UserCount+i)
				rsPhotoUser("InnerWidth")=InnerWidth(UserCount+i)
				rsPhotoUser("InnerHeight")=InnerHeight(UserCount+i)
				rsPhotoUser("Direction")=Direction(UserCount+i)
				rsPhotoUser.update
				Flag=True
			end if
			rsPhotoUser.Update
			rsPhotoUser.movenext
		end if
	next
	do while not rsPhotoUser.eof
		rsPhotoUser.delete
		rsPhotoUser.movenext
		Flag=True
	loop
	rsPhotoUser.close
	sql="select * from Visual_photouser where photo_id="&photoid
	rsPhotoUser.open sql,conn,1,3
	i=0
	do while not rsPhotoUser.eof and i<UserCount
		Flag1=False
		if hasFinished then
			if rsPhotoUser("LayerNo")<>LayerNo(i) or rsPhotoUser("OuterLeft")<>OuterLeft(i) or rsPhotoUser("OuterTop")<>OuterTop(i) or rsPhotoUser("OuterWidth")<>OuterWidth(i) or rsPhotoUser("OuterHeight")<>OuterHeight(i) or rsPhotoUser("InnerLeft")<>InnerLeft(i) or rsPhotoUser("InnerTop")<>InnerTop(i) or rsPhotoUser("InnerWidth")<>InnerWidth(i) or rsPhotoUser("InnerHeight")<>InnerHeight(i) or rsPhotoUser("Direction")<>Direction(i) Then Flag1=True
			if rsPhotoUser("NameLeft")<>NameLeft(i) or rsPhotoUser("NameTop")<>NameTop(i) or rsPhotoUser("NameFont")<>NameFont(i) or rsPhotoUser("NameSize")<>NameSize(i) or rsPhotoUser("NameBold")<>NameBold(i) or rsPhotoUser("NameItalic")<>NameItalic(i) or rsPhotoUser("NameColor")<>NameColor(i) or rsPhotoUser("NameDirection")<>NameDirection(i) then Flag1=True
		else
			Flag1=True
		end if
		if Flag1 Then
			rsPhotoUser("LayerNo")=LayerNo(i)
			rsPhotoUser("OuterLeft")=OuterLeft(i)
			rsPhotoUser("OuterTop")=OuterTop(i)
			rsPhotoUser("OuterWidth")=OuterWidth(i)
			rsPhotoUser("OuterHeight")=OuterHeight(i)
			rsPhotoUser("InnerLeft")=InnerLeft(i)
			rsPhotoUser("InnerTop")=InnerTop(i)
			rsPhotoUser("InnerWidth")=InnerWidth(i)
			rsPhotoUser("InnerHeight")=InnerHeight(i)
			rsPhotoUser("Direction")=Direction(i)
			rsPhotoUser("NameLeft")=NameLeft(i)
			rsPhotoUser("NameTop")=NameTop(i)
			rsPhotoUser("NameFont")=NameFont(i)
			rsPhotoUser("NameSize")=NameSize(i)
			rsPhotoUser("NameBold")=NameBold(i)
			rsPhotoUser("NameItalic")=NameItalic(i)
			rsPhotoUser("NameColor")=NameColor(i)
			rsPhotoUser("NameDirection")=NameDirection(i)
			rsPhotoUser.update
			Flag=True
		end if
		if ((not rsPhotoUser("Finished") or Flag) and lcase(trim(rsPhotoUser("UserName")))=lcase(trim(membername))) or (master and request.form("forcephoto")="1") then
			rsPhotoUser("Finished")=true
			rsPhotoUser.Update
		end if
		rsPhotoUser.movenext
		i=i+1
	loop
	rsPhotoUser.close
	if Flag Then
		rsphoto("FinishDate")=now()
		if (master and request.form("forcephoto")="1") or UserCount=1 then
			rsPhoto("Finished")=true
		end if
		rsphoto.update
	end if
	rsPhoto.close
	set rsPhoto=nothing
	if (master and request.form("forcephoto")="1") or UserCount=1 then
		if hasASPPainter() and LocalPic=1 then
			conn.execute("update Visual_Photo set URL='"&CreateVisualPhoto(PhotoID)&"' where id="&photoid)
		end if
		if UserCount=1 then 
			sucmsg=sucmsg+"<br>"+"<li><b>个人写真成功，请到您的虚拟相簿欣赏这张照片。</b>"
		else
			sucmsg=sucmsg+"<br>"+"<li><b>强制合影成功，请到您的虚拟相簿欣赏这张照片。</b>"
		end if
	else
		if Flag Then
			sql="select * from Visual_photouser where photo_id="&photoid
			rsPhotoUser.open sql,conn,1,1
			do while not rsPhotoUser.eof
				if lcase(trim(rsPhotoUser("username")))<>lcase(trim(membername)) then
					if hasFinished then
						sql="insert into message (incept,sender,title,content,sendtime,flag,issend) values ('"&rsPhotoUser("UserName")&"','"&membername&"','有关《"&photoname&"》的合影再次修改了设计','您好："&chr(10)&"　　很高兴地通知您，我已经再次修改了这张合影的设计。"&chr(10)& "　　请点击如下地址确定您是否满意这个设计："&chr(10)&chr(10)&chr(10)&"[align=center][URL=z_Visual.asp?shopid=497&photoid="&photoid&"][color=blue][b][u]对设计发表意见[/u][/b][/color][/URL][/align]',Now(),0,1)"
					else
						sql="insert into message (incept,sender,title,content,sendtime,flag,issend) values ('"&rsPhotoUser("UserName")&"','"&membername&"','有关《"&photoname&"》的合影已经设计完成','您好："&chr(10)&"　　很高兴地通知您，我已经完成了这张合影的设计。"&chr(10)& "　　请点击如下地址确定您是否满意这个设计："&chr(10)&chr(10)&chr(10)&"[align=center][URL=z_Visual.asp?shopid=497&photoid="&photoid&"][color=blue][b][u]对设计发表意见[/u][/b][/color][/URL][/align]',Now(),0,1)"
					end if
					if rsPhotoUser("Finished") or not hasFinished then
						conn.execute(sql)
					end if
				end if
				rsPhotoUser.movenext
			loop
			rsPhotoUser.close
			conn.execute("update visual_photouser set finished=false where photo_id="&photoid&" and username<>'"&membername&"'")
		end if
		sucmsg="<br>"+"<li><b>恭喜您，设计合影成功。</b>"
	end if
	set rsPhotoUser=nothing
	call dvbbs_suc()
end sub

sub acceptmsg()
	dim PhotoID,rsPhoto,rsPhotoUser,PhotoName,Flag
	if isnull(request("photoid")) then
		errmsg=errmsg+"<br>"+"<li>提交的数据有错误！"
		founderr=true
		exit sub
	elseif not isInteger(request("photoid")) then
		errmsg=errmsg+"<br>"+"<li>提交的数据有错误！"
		founderr=true
		exit sub
	elseif conn.execute("select u.photo_id from visual_photouser u inner join visual_photo p on u.photo_id=p.id where u.photo_id="&checkstr(trim(request("photoid")))&" and u.username='"&membername&"' and p.confirmed and not p.finished and not u.finished").eof then
		errmsg=errmsg+"<br>"+"<li>提交的数据有错误！"
		founderr=true
		exit sub
	elseif conn.execute("select u.photo_id from visual_photouser u inner join visual_photo p on u.photo_id=p.id and u.username=p.fromuser where u.photo_id="&checkstr(trim(request("photoid")))&" and u.finished").eof then
		errmsg=errmsg+"<br>"+"<li>提交的数据有错误！"
		founderr=true
		exit sub
	else
		photoid=checkstr(trim(request("photoid")))
	end if

	stats="同意设计成功"
	conn.execute("update visual_photouser set Finished=true where photo_id="&photoid&" and username='"&membername&"'")
	flag=false
	if conn.execute("select photo_id from visual_photouser where photo_id="&photoid&" and not finished").eof then
		flag=true
		set rsPhoto=server.createobject("adodb.recordset")
		sql="select Finished,name from visual_photo where id="&photoid
		rsPhoto.open sql,conn,1,3
		rsPhoto(0)=true
		PhotoName=rsPhoto(1)
		rsPhoto.update
		rsPhoto.close
		set rsPhoto=nothing
		set rsPhotoUser=server.createobject("ADODB.Recordset")
		sql="select username,ConfirmTime from visual_photouser where photo_id="&photoid
		rsPhotoUser.open sql,conn,1,1
		do while not rsPhotoUser.eof
			if lcase(trim(rsPhotoUser(0)))<>lcase(trim(membername)) then
				sql="insert into message (incept,sender,title,content,sendtime,flag,issend) values ('"&_
					rsPhotoUser(0)&"','"&forum_info(0)&"','您参与的合影《"&PhotoName&"》已经被所有人同意了','您好："&_
					chr(10)&"　　很高兴地通知您，您于 [b]"&formatdatetime(rsPhotoUser(1),2)&"[/b] 同意参与的合影已经被所有的参与者同意了。"&_
					chr(10)&"　　请到以下地址欣赏这个合影："&chr(10)&chr(10)&chr(10)&"[align=center][URL=z_Visual.asp?shopid=597"&_
					"][color=blue][b][u]欣赏合影[/u][/b][/color][/URL][/align]',Now(),0,1)"
				conn.execute(sql)
			end if
			rsPhotoUser.movenext
		loop
		rsPhotoUser.close
		set rsPhotoUser=nothing
	end if
	sucmsg="<br>"+"<li><b>恭喜您，同意合影成功。</b>"
	if flag then 
		sucmsg=sucmsg+"<br>"+"<li><b>所有人已经全部同意了这个合影，请到您的虚拟相簿欣赏这张照片。</b>"
		if hasASPPainter() and LocalPic=1 then
			conn.execute("update Visual_Photo set URL='"&CreateVisualPhoto(PhotoID)&"' where id="&photoid)
		end if
	end if
	call dvbbs_suc()
end sub

sub recreatemsg()
	dim PhotoID,rsPhoto,rsPhotoUser,PhotoName,Flag
	if isnull(request("photoid")) then
		errmsg=errmsg+"<br>"+"<li>提交的数据有错误！"
		founderr=true
		exit sub
	elseif not isInteger(request("photoid")) then
		errmsg=errmsg+"<br>"+"<li>提交的数据有错误！"
		founderr=true
		exit sub
	elseif not master and conn.execute("select u.photo_id from visual_photouser u inner join visual_photo p on u.photo_id=p.id where u.photo_id="&checkstr(trim(request("photoid")))&" and u.username='"&membername&"' and p.finished").eof then
		errmsg=errmsg+"<br>"+"<li>提交的数据有错误！"
		founderr=true
		exit sub
	else
		photoid=checkstr(trim(request("photoid")))
	end if

	stats="重建照片成功"
	sucmsg=sucmsg+"<br>"+"<li><b>照片图片已经重新生成完毕。</b>"
	if hasASPPainter() and LocalPic=1 then
		conn.execute("update Visual_Photo set URL='"&CreateVisualPhoto(PhotoID)&"' where id="&photoid)
	end if
	call dvbbs_suc()
end sub

function CreateVisualPhoto(PhotoID)
	dim photoname,bgwidth,bgheight,photobg,PhotoBodyBack,PhotoBodyFore,PhotoFg,j,TempValue,BodyOrder()
	dim usercount,uservisualsplit(),usernamesplit(),PhotoDate
	dim PhotoNameLeft,PhotoNameTop,PhotoNameFont,PhotoNameSize,PhotoNameBold,PhotoNameItalic,PhotoNameColor,PhotoNameDirection
	dim DateLeft,DateTop,DateFont,DateSize,DateBold,DateItalic,DateColor,DateDirection
	dim NameLeft(),NameTop(),NameFont(),NameSize(),NameColor(),NameDirection()
	dim OuterLeft(),OuterTop(),OuterWidth(),OuterHeight()
	dim InnerLeft(),InnerTop(),InnerWidth(),InnerHeight()
	dim LayerNo(),Direction(),isUpload,PhotoURL

	sqlvisual="select * from visual_photo where id="&photoid
	rsvisual.open sqlvisual,conn,1,1
	photoname=rsvisual("name")
	bgwidth=rsvisual("width")
	bgheight=rsvisual("height")
	isUpload=rsvisual("isupload")
	photobg=rsvisual("background")
	photoBodyBack=rsvisual("BackBody")
	photoBodyFore=rsvisual("ForeBody")
	photoFg=rsvisual("Foreground")
	PhotoDate=rsvisual("adddate")
	PhotoNameLeft=rsvisual("PhotoNameLeft")
	PhotoNameTop=rsvisual("PhotoNameTop")
	PhotoNameFont=rsvisual("PhotoNameFont")
	PhotoNameSize=rsvisual("PhotoNameSize")
	PhotoNameBold=rsvisual("PhotoNameBold")
	PhotoNameItalic=rsvisual("PhotoNameItalic")
	PhotoNameColor=rsvisual("PhotoNameColor")
	PhotoNameDirection=rsvisual("PhotoNameDirection")
	DateLeft=rsvisual("DateLeft")
	DateTop=rsvisual("DateTop")
	DateFont=rsvisual("DateFont")
	DateSize=rsvisual("DateSize")
	DateBold=rsvisual("DateBold")
	DateItalic=rsvisual("DateItalic")
	DateColor=rsvisual("DateColor")
	DateDirection=rsvisual("DateDirection")
	PhotoURL=rsvisual("URL")
	rsvisual.close
	sqlvisual="select * from visual_photouser where photo_id="&photoid
	rsvisual.open sqlvisual,conn,1,1
	usercount=0
	do while not rsvisual.eof
		redim preserve uservisualsplit(usercount)
		uservisualsplit(usercount)=rsvisual("uservisual")
		redim preserve usernamesplit(usercount)
		usernamesplit(usercount)=rsvisual("username")
		redim preserve LayerNo(UserCount)
		redim preserve OuterLeft(UserCount)
		redim preserve OuterTop(UserCount)
		redim preserve OuterWidth(UserCount)
		redim preserve OuterHeight(UserCount)
		redim preserve InnerLeft(UserCount)
		redim preserve InnerTop(UserCount)
		redim preserve InnerWidth(UserCount)
		redim preserve InnerHeight(UserCount)
		redim preserve Direction(UserCount)
		redim preserve NameLeft(UserCount)
		redim preserve NameTop(UserCount)
		redim preserve NameFont(UserCount)
		redim preserve NameSize(UserCount)
		redim preserve NameBold(UserCount)
		redim preserve NameItalic(UserCount)
		redim preserve NameColor(UserCount)
		redim preserve NameDirection(UserCount)
		LayerNo(UserCount)=rsvisual("LayerNo")
		OuterLeft(UserCount)=rsvisual("OuterLeft")
		OuterTop(UserCount)=rsvisual("OuterTop")
		OuterWidth(UserCount)=rsvisual("OuterWidth")
		OuterHeight(UserCount)=rsvisual("OuterHeight")
		InnerLeft(UserCount)=rsvisual("InnerLeft")
		InnerTop(UserCount)=rsvisual("InnerTop")
		InnerWidth(UserCount)=rsvisual("InnerWidth")
		InnerHeight(UserCount)=rsvisual("InnerHeight")
		Direction(UserCount)=rsvisual("Direction")
		NameLeft(UserCount)=rsvisual("NameLeft")
		NameTop(UserCount)=rsvisual("NameTop")
		NameFont(UserCount)=rsvisual("NameFont")
		NameSize(UserCount)=rsvisual("NameSize")
		NameBold(UserCount)=rsvisual("NameBold")
		NameItalic(UserCount)=rsvisual("NameItalic")
		NameColor(UserCount)=rsvisual("NameColor")
		NameDirection(UserCount)=rsvisual("NameDirection")
		usercount=usercount+1
		rsvisual.movenext
	loop
	rsvisual.close
	redim preserve UserVisualSplit(UserCount+5)
	redim preserve OuterLeft(UserCount+5)
	redim preserve OuterTop(UserCount+5)
	redim preserve OuterWidth(UserCount+5)
	redim preserve OuterHeight(UserCount+5)
	redim preserve InnerLeft(UserCount+5)
	redim preserve InnerTop(UserCount+5)
	redim preserve InnerWidth(UserCount+5)
	redim preserve InnerHeight(UserCount+5)
	redim preserve Direction(UserCount+5)
	for i=0 to 5
		UserVisualSplit(UserCount+i)=""
		OuterLeft(UserCount+i)=0
		OuterTop(UserCount+i)=0
		OuterWidth(UserCount+i)=140
		OuterHeight(UserCount+i)=226
		InnerLeft(UserCount+i)=0
		InnerTop(UserCount+i)=0
		InnerWidth(UserCount+i)=140
		InnerHeight(UserCount+i)=226
		Direction(UserCount+i)=0
	next
	sqlvisual="select * from visual_Accouterment where photo_id="&PhotoID
	rsVisual.open sqlvisual,conn,1,1
	do while not rsVisual.eof
		i=rsVisual("SeqNo")
		uservisualsplit(UserCount+i)=rsVisual("ItemPicPath")
		OuterLeft(UserCount+i)=rsVisual("OuterLeft")
		OuterTop(UserCount+i)=rsVisual("OuterTop")
		OuterWidth(UserCount+i)=rsVisual("OuterWidth")
		OuterHeight(UserCount+i)=rsVisual("OuterHeight")
		InnerLeft(UserCount+i)=rsVisual("InnerLeft")
		InnerTop(UserCount+i)=rsVisual("InnerTop")
		InnerWidth(UserCount+i)=rsVisual("InnerWidth")
		InnerHeight(UserCount+i)=rsVisual("InnerHeight")
		Direction(UserCount+i)=rsVisual("Direction")
		rsVisual.movenext
	loop
	rsVisual.close
	redim BodyOrders(UserCount-1)
	for i=0 to UserCount-1
		BodyOrders(i)=i
	next
	for i=0 to UserCount-2
		for j=i to UserCount-1
			if LayerNo(i)>LayerNo(j) then
				TempValue=BodyOrders(i)
				BodyOrders(i)=BodyOrders(j)
				BodyOrders(j)=TempValue
				TempValue=LayerNo(i)
				LayerNo(i)=LayerNo(j)
				LayerNo(j)=TempValue
			end if
		next
	next
	dim pic
	Set pic = CreateObject("ASPPainter.Pictures.1")
	pic.SetFormat 3
	if PhotoBg<>"" then
		if isUpload then
			pic.SetBkColor 255,255,255,255
			pic.Create 280,226
			pic.SetColor 255,255,255,255
			pic.SetColorAsTransparent
			pic.SetImageIndex 1
			if lcase(right(PhotoBg,3))="gif" then
				pic.SetFormat 3
			else
				pic.SetFormat 1
			end if
			pic.LoadFile Server.MapPath(PhotoPath&"/"&PhotoBg)
			pic.ResizeCopy 0,1,0,0,0,0,280,226,pic.Width,pic.Height
			pic.destroy
			pic.SetImageIndex 0
		else
			pic.LoadFile Server.MapPath(PicPath&PhotoBg)
		end if
	else
		if bgWidth=140 then
			pic.SetBkColor 255,255,255,255
			pic.Create 140,226
			pic.SetColor 255,255,255,255
			pic.SetColorAsTransparent
		else
			pic.SetBkColor 255,255,255,255
			pic.Create 280,226
			pic.SetColor 255,255,255,255
			pic.SetColorAsTransparent
		end if
	end if
	if PhotoBodyBack<>"" then
		pic.SetImageIndex 1
		pic.LoadFile Server.MapPath(PicPath&PhotoBodyBack)
		pic.Merge 0,1,0,0,0,0,pic.Width,pic.Height,100
		pic.destroy
		pic.SetImageIndex 0
	end if
	dim BodyStr,BodySplit
	dim SrcLeft,SrcTop,SrcWidth,SrcHeight
	dim DstLeft,DstTop,DstWidth,DstHeight
	dim Border_Cont_Left,Border_Cont_Top,Border_Cont_Width,Border_Cont_Height
	dim Border_Cont_Bg_Left,Border_Cont_Bg_Top,Border_Cont_Bg_Width,Border_Cont_Bg_Height
	for i=0 to 2
		if UserVisualSplit(UserCount+i)<>"" then
			pic.SetImageIndex 1
			pic.SetFormat 3
			pic.LoadFile Server.MapPath(PicPath&UserVisualSplit(UserCount+i))
			Border_Cont_Left=GetIntersectionBegin(InnerLeft(UserCount+i),InnerWidth(UserCount+i),0,OuterWidth(UserCount+i))
			Border_Cont_Top=GetIntersectionBegin(InnerTop(UserCount+i),InnerHeight(UserCount+i),0,OuterHeight(UserCount+i))
			Border_Cont_Width=GetIntersectionLength(InnerLeft(UserCount+i),InnerWidth(UserCount+i),0,OuterWidth(UserCount+i))
			Border_Cont_Height=GetIntersectionLength(InnerTop(UserCount+i),InnerHeight(UserCount+i),0,OuterHeight(UserCount+i))
			Border_Cont_Bg_Left=GetIntersectionBegin(-OuterLeft(UserCount+i),bgWidth,Border_Cont_Left,Border_Cont_Width)
			Border_Cont_Bg_Top=GetIntersectionBegin(-OuterTop(UserCount+i),bgHeight,Border_Cont_Top,Border_Cont_Height)
			Border_Cont_Bg_Width=GetIntersectionLength(-OuterLeft(UserCount+i),bgWidth,Border_Cont_Left,Border_Cont_Width)
			Border_Cont_Bg_Height=GetIntersectionLength(-OuterTop(UserCount+i),bgHeight,Border_Cont_Top,Border_Cont_Height)
			SrcLeft=Border_Cont_Bg_Left-InnerLeft(UserCount+i)
			SrcTop=Border_Cont_Bg_Top-InnerTop(UserCount+i)
			SrcWidth=GetSizedValue(Border_Cont_Bg_Width,InnerWidth(UserCount+i),140)
			SrcHeight=GetSizedValue(Border_Cont_Bg_Height,InnerHeight(UserCount+i),226)
			DstLeft=GetIntersectionBegin(OuterLeft(UserCount+i),OuterWidth(UserCount+i),0,bgWidth)
			DstTop=GetIntersectionBegin(OuterTop(UserCount+i),OuterHeight(UserCount+i),0,bgHeight)
			DstWidth=GetIntersectionLength(OuterLeft(UserCount+i),OuterWidth(UserCount+i),0,bgWidth)
			DstHeight=GetIntersectionLength(OuterTop(UserCount+i),OuterHeight(UserCount+i),0,bgHeight)
			if Direction(UserCount+i)\4>0 then
				pic.SetImageIndex 1
				pic.HorizontalFlip
			end if
			if SrcWidth=DstWidth and SrcHeight=DstHeight then
				pic.Merge 0,1,DstLeft,DstTop,SrcLeft,SrcTop,DstWidth,DstHeight,100
			else
				pic.SetImageIndex 2
				pic.SetFormat 3
				pic.SetBkColor 255,255,255,255
				pic.Create 280,454
				pic.SetColor 255,255,255,255
				pic.SetColorAsTransparent
				pic.ResizeCopy 2,1,0,0,0,0,InnerWidth(UserCount+i),InnerHeight(UserCount+i),140,226
				pic.SetImageIndex 1
				pic.Destroy
				pic.SetImageIndex 2
				pic.Merge 0,2,DstLeft,DstTop,SrcLeft,SrcTop,DstWidth,DstHeight,100
			end if
			pic.Destroy
		end if
	next
	for i=0 to usercount-1
		BodyStr=GetBody(UserVisualSplit(BodyOrders(i)),(UserCount=1))
		BodySplit=Split(BodyStr,"|")
		pic.SetImageIndex 1
		pic.SetFormat 3
		pic.LoadFile Server.MapPath(PicPath&BodySplit(0))
		for j=1 to ubound(bodysplit)
			pic.SetImageIndex 2
			pic.SetFormat 3
			pic.LoadFile Server.MapPath(PicPath&BodySplit(j))
			pic.Merge 1,2,0,0,0,0,140,226,100
			pic.Destroy
		next
		Border_Cont_Left=GetIntersectionBegin(InnerLeft(BodyOrders(i)),InnerWidth(BodyOrders(i)),0,OuterWidth(BodyOrders(i)))
		Border_Cont_Top=GetIntersectionBegin(InnerTop(BodyOrders(i)),InnerHeight(BodyOrders(i)),0,OuterHeight(BodyOrders(i)))
		Border_Cont_Width=GetIntersectionLength(InnerLeft(BodyOrders(i)),InnerWidth(BodyOrders(i)),0,OuterWidth(BodyOrders(i)))
		Border_Cont_Height=GetIntersectionLength(InnerTop(BodyOrders(i)),InnerHeight(BodyOrders(i)),0,OuterHeight(BodyOrders(i)))
		Border_Cont_Bg_Left=GetIntersectionBegin(-OuterLeft(BodyOrders(i)),bgWidth,Border_Cont_Left,Border_Cont_Width)
		Border_Cont_Bg_Top=GetIntersectionBegin(-OuterTop(BodyOrders(i)),bgHeight,Border_Cont_Top,Border_Cont_Height)
		Border_Cont_Bg_Width=GetIntersectionLength(-OuterLeft(BodyOrders(i)),bgWidth,Border_Cont_Left,Border_Cont_Width)
		Border_Cont_Bg_Height=GetIntersectionLength(-OuterTop(BodyOrders(i)),bgHeight,Border_Cont_Top,Border_Cont_Height)
		SrcLeft=Border_Cont_Bg_Left-InnerLeft(BodyOrders(i))
		SrcTop=Border_Cont_Bg_Top-InnerTop(BodyOrders(i))
		SrcWidth=GetSizedValue(Border_Cont_Bg_Width,InnerWidth(BodyOrders(i)),140)
		SrcHeight=GetSizedValue(Border_Cont_Bg_Height,InnerHeight(BodyOrders(i)),226)
		DstLeft=GetIntersectionBegin(OuterLeft(BodyOrders(i)),OuterWidth(BodyOrders(i)),0,bgWidth)
		DstTop=GetIntersectionBegin(OuterTop(BodyOrders(i)),OuterHeight(BodyOrders(i)),0,bgHeight)
		DstWidth=GetIntersectionLength(OuterLeft(BodyOrders(i)),OuterWidth(BodyOrders(i)),0,bgWidth)
		DstHeight=GetIntersectionLength(OuterTop(BodyOrders(i)),OuterHeight(BodyOrders(i)),0,bgHeight)
		if Direction(BodyOrders(i))\4>0 then
			pic.SetImageIndex 1
			pic.HorizontalFlip
		end if
		if SrcWidth=DstWidth and SrcHeight=DstHeight then
			pic.Merge 0,1,DstLeft,DstTop,SrcLeft,SrcTop,DstWidth,DstHeight,100
		else
			pic.SetImageIndex 2
			pic.SetFormat 3
			pic.SetBkColor 255,255,255,255
			pic.Create 280,454
			pic.SetColor 255,255,255,255
			pic.SetColorAsTransparent
			pic.ResizeCopy 2,1,0,0,0,0,InnerWidth(BodyOrders(i)),InnerHeight(BodyOrders(i)),140,226
			pic.SetImageIndex 1
			pic.Destroy
			pic.SetImageIndex 2
			pic.Merge 0,2,DstLeft,DstTop,SrcLeft,SrcTop,DstWidth,DstHeight,100
		end if
		pic.Destroy
	next
	pic.SetImageIndex 0
	for i=0 to usercount-1
		if NameColor(i)<>"" then
			pic.SetFontName NameFont(i)
			if pic.ExistFont <> 1 then pic.SetFontName DefaultFont
			pic.SetFontSize cint(mid(NameSize(i),1,len(NameSize(i))-2))
			pic.SetColor 0,0,0,255
			if NameBold(i) then
				pic.SetFontBold 1
			end if
			if NameItalic(i) then
				pic.SetFontItalic 1
			end if
			if NameDirection(i)=0 then
				pic.SetFontOrientation 0
				pic.TextOut NameLeft(i)+1,NameTop(i)+1,trim(UserNameSplit(i))
			elseif NameDirection(i)=1 then
				pic.SetFontOrientation 90
				pic.TextOut NameLeft(i)+pic.GetFontHeight-2,NameTop(i),trim(UserNameSplit(i))
			elseif NameDirection(i)=2 then
				pic.SetFontOrientation 180
				pic.TextOut NameLeft(i)+142,NameTop(i)+pic.GetFontHeight+1,trim(UserNameSplit(i))
			elseif NameDirection(i)=3 then
				pic.SetFontOrientation 270
				pic.TextOut NameLeft(i)+1,NameTop(i)+142,trim(UserNameSplit(i))
			end if
			pic.SetColor Hex2Dec(mid(NameColor(i),2,2)),Hex2Dec(mid(NameColor(i),4,2)),Hex2Dec(mid(NameColor(i),6,2)),255
			if NameDirection(i)=0 then
				pic.SetFontOrientation 0
				pic.TextOut NameLeft(i),NameTop(i),trim(UserNameSplit(i))
			elseif NameDirection(i)=1 then
				pic.SetFontOrientation 90
				pic.TextOut NameLeft(i)+pic.GetFontHeight-3,NameTop(i)-1,trim(UserNameSplit(i))
			elseif NameDirection(i)=2 then
				pic.SetFontOrientation 180
				pic.TextOut NameLeft(i)+141,NameTop(i)+pic.GetFontHeight,trim(UserNameSplit(i))
			elseif NameDirection(i)=3 then
				pic.SetFontOrientation 270
				pic.TextOut NameLeft(i),NameTop(i)+141,trim(UserNameSplit(i))
			end if
		end if
	next
	pic.SetImageIndex 0
	for i=3 to 5
		if UserVisualSplit(UserCount+i)<>"" then
			pic.SetImageIndex 1
			pic.SetFormat 3
			pic.LoadFile Server.MapPath(PicPath&UserVisualSplit(UserCount+i))
			Border_Cont_Left=GetIntersectionBegin(InnerLeft(UserCount+i),InnerWidth(UserCount+i),0,OuterWidth(UserCount+i))
			Border_Cont_Top=GetIntersectionBegin(InnerTop(UserCount+i),InnerHeight(UserCount+i),0,OuterHeight(UserCount+i))
			Border_Cont_Width=GetIntersectionLength(InnerLeft(UserCount+i),InnerWidth(UserCount+i),0,OuterWidth(UserCount+i))
			Border_Cont_Height=GetIntersectionLength(InnerTop(UserCount+i),InnerHeight(UserCount+i),0,OuterHeight(UserCount+i))
			Border_Cont_Bg_Left=GetIntersectionBegin(-OuterLeft(UserCount+i),bgWidth,Border_Cont_Left,Border_Cont_Width)
			Border_Cont_Bg_Top=GetIntersectionBegin(-OuterTop(UserCount+i),bgHeight,Border_Cont_Top,Border_Cont_Height)
			Border_Cont_Bg_Width=GetIntersectionLength(-OuterLeft(UserCount+i),bgWidth,Border_Cont_Left,Border_Cont_Width)
			Border_Cont_Bg_Height=GetIntersectionLength(-OuterTop(UserCount+i),bgHeight,Border_Cont_Top,Border_Cont_Height)
			SrcLeft=Border_Cont_Bg_Left-InnerLeft(UserCount+i)
			SrcTop=Border_Cont_Bg_Top-InnerTop(UserCount+i)
			SrcWidth=GetSizedValue(Border_Cont_Bg_Width,InnerWidth(UserCount+i),140)
			SrcHeight=GetSizedValue(Border_Cont_Bg_Height,InnerHeight(UserCount+i),226)
			DstLeft=GetIntersectionBegin(OuterLeft(UserCount+i),OuterWidth(UserCount+i),0,bgWidth)
			DstTop=GetIntersectionBegin(OuterTop(UserCount+i),OuterHeight(UserCount+i),0,bgHeight)
			DstWidth=GetIntersectionLength(OuterLeft(UserCount+i),OuterWidth(UserCount+i),0,bgWidth)
			DstHeight=GetIntersectionLength(OuterTop(UserCount+i),OuterHeight(UserCount+i),0,bgHeight)
			if Direction(UserCount+i)\4>0 then
				pic.SetImageIndex 1
				pic.HorizontalFlip
			end if
			if SrcWidth=DstWidth and SrcHeight=DstHeight then
				pic.Merge 0,1,DstLeft,DstTop,SrcLeft,SrcTop,DstWidth,DstHeight,100
			else
				pic.SetImageIndex 2
				pic.SetFormat 3
				pic.SetBkColor 255,255,255,255
				pic.Create 280,454
				pic.SetColor 255,255,255,255
				pic.SetColorAsTransparent
				pic.ResizeCopy 2,1,0,0,0,0,InnerWidth(UserCount+i),InnerHeight(UserCount+i),140,226
				pic.SetImageIndex 1
				pic.Destroy
				pic.SetImageIndex 2
				pic.Merge 0,2,DstLeft,DstTop,SrcLeft,SrcTop,DstWidth,DstHeight,100
			end if
			pic.Destroy
		end if
	next
	if PhotoBodyFore<>"" then
		pic.SetImageIndex 1
		pic.LoadFile Server.MapPath(PicPath&PhotoBodyFore)
		pic.Merge 0,1,0,0,0,0,pic.Width,pic.Height,100
		pic.destroy
		pic.SetImageIndex 0
	end if
	if PhotoFg<>"" then
		pic.SetImageIndex 1
		pic.LoadFile Server.MapPath(PicPath&PhotoFg)
		pic.Merge 0,1,0,0,0,0,pic.Width,pic.Height,100
		pic.destroy
		pic.SetImageIndex 0
	end if
	if PhotoNameColor<>"" then
		pic.SetFontName PhotoNameFont
		if pic.ExistFont <> 1 then pic.SetFontName DefaultFont
		pic.SetFontSize cint(mid(PhotoNameSize,1,len(PhotoNameSize)-2))
		pic.SetColor 0,0,0,255
		if PhotoNameBold then
			pic.SetFontBold 1
		else
			pic.SetFontBold 0
		end if
		if PhotoNameItalic then
			pic.SetFontItalic 1
		else
			pic.SetFontItalic 0
		end if
		if PhotoNameDirection=0 then
			pic.SetFontOrientation 0
			pic.TextOut PhotoNameLeft+1,PhotoNameTop+1,trim(PhotoName)
		elseif PhotoNameDirection=1 then
			pic.SetFontOrientation 90
			pic.TextOut PhotoNameLeft+pic.GetFontHeight-2,PhotoNameTop,trim(PhotoName)
		elseif PhotoNameDirection=2 then
			pic.SetFontOrientation 180
			pic.TextOut PhotoNameLeft+142,PhotoNameTop+pic.GetFontHeight+1,trim(PhotoName)
		elseif PhotoNameDirection=3 then
			pic.SetFontOrientation 270
			pic.TextOut PhotoNameLeft+1,PhotoNameTop+142,trim(PhotoName)
		end if
		pic.SetColor Hex2Dec(mid(PhotoNameColor,2,2)),Hex2Dec(mid(PhotoNameColor,4,2)),Hex2Dec(mid(PhotoNameColor,6,2)),255
		if PhotoNameDirection=0 then
			pic.SetFontOrientation 0
			pic.TextOut PhotoNameLeft,PhotoNameTop,trim(PhotoName)
		elseif PhotoNameDirection=1 then
			pic.SetFontOrientation 90
			pic.TextOut PhotoNameLeft+pic.GetFontHeight-3,PhotoNameTop-1,trim(PhotoName)
		elseif PhotoNameDirection=2 then
			pic.SetFontOrientation 180
			pic.TextOut PhotoNameLeft+141,PhotoNameTop+pic.GetFontHeight,trim(PhotoName)
		elseif PhotoNameDirection=3 then
			pic.SetFontOrientation 270
			pic.TextOut PhotoNameLeft,PhotoNameTop+141,trim(PhotoName)
		end if
	end if
	if DateColor<>"" then
		pic.SetFontName DateFont
		if pic.ExistFont <> 1 then pic.SetFontName DefaultFont
		pic.SetFontSize cint(mid(DateSize,1,len(DateSize)-2))
		pic.SetColor 0,0,0,255
		if DateBold then
			pic.SetFontBold 1
		else
			pic.SetFontBold 0
		end if
		if DateItalic then
			pic.SetFontItalic 1
		else
			pic.SetFontItalic 0
		end if
		if DateDirection=0 then
			pic.SetFontOrientation 0
			pic.TextOut DateLeft+1,DateTop+1,formatdatetime(PhotoDate,2)
		elseif DateDirection=1 then
			pic.SetFontOrientation 90
			pic.TextOut DateLeft+pic.GetFontHeight-2,DateTop,formatdatetime(PhotoDate,2)
		elseif DateDirection=2 then
			pic.SetFontOrientation 180
			pic.TextOut DateLeft+142,DateTop+pic.GetFontHeight+1,formatdatetime(PhotoDate,2)
		elseif DateDirection=3 then
			pic.SetFontOrientation 270
			pic.TextOut DateLeft+1,DateTop+142,formatdatetime(PhotoDate,2)
		end if
		pic.SetColor Hex2Dec(mid(DateColor,2,2)),Hex2Dec(mid(DateColor,4,2)),Hex2Dec(mid(DateColor,6,2)),255
		if DateDirection=0 then
			pic.SetFontOrientation 0
			pic.TextOut DateLeft,DateTop,formatdatetime(PhotoDate,2)
		elseif DateDirection=1 then
			pic.SetFontOrientation 90
			pic.TextOut DateLeft+pic.GetFontHeight-3,DateTop-1,formatdatetime(PhotoDate,2)
		elseif DateDirection=2 then
			pic.SetFontOrientation 180
			pic.TextOut DateLeft+141,DateTop+pic.GetFontHeight,formatdatetime(PhotoDate,2)
		elseif DateDirection=3 then
			pic.SetFontOrientation 270
			pic.TextOut DateLeft,DateTop+141,formatdatetime(PhotoDate,2)
		end if
	end if
	pic.SetImageIndex 0
	pic.SetFormat 0
	if isnull(PhotoURL) or PhotoURL="" then
		randomize
		PhotoURL=PhotoPath&"/Photo_"&PhotoId&"_"&right("000000"&int(rnd*1000000),6)&".png"
	end if
	pic.SaveToFile Server.MapPath(PhotoURL)
	CreateVisualPhoto=PhotoURL
	pic.DestroyAll
	Set pic = Nothing
end function

function GetIntersectionBegin(SrcBegin,SrcLength,DstBegin,DstLength)
	dim SrcEnd,DstEnd
	SrcEnd=SrcBegin+SrcLength-1
	DstEnd=DstBegin+DstLength-1
	if SrcEnd<DstBegin or SrcBegin>DstEnd or SrcBegin<DstBegin then
		GetIntersectionBegin=DstBegin
	else
		GetIntersectionBegin=SrcBegin
	end if
end function

function GetIntersectionLength(SrcBegin,SrcLength,DstBegin,DstLength)
	dim SrcEnd,DstEnd
	SrcEnd=SrcBegin+SrcLength-1
	DstEnd=DstBegin+DstLength-1
	if SrcEnd<DstBegin or SrcBegin>DstEnd then
		GetIntersectionLength=0
	elseif SrcBegin<DstBegin and SrcEnd>=DstBegin and SrcEnd<=DstEnd then 
		GetIntersectionLength=SrcLength+SrcBegin-DstBegin
	elseif SrcBegin>=DstBegin and SrcEnd<=DstEnd then
		GetIntersectionLength=SrcLength
	elseif SrcBegin>=DstBegin and SrcEnd>DstEnd then
		GetIntersectionLength=DstLength-(SrcBegin-DstBegin)
	else
		GetIntersectionLength=DstLength
	end if
end function

function GetSizedValue(SrcValue,SrcLength,DstLength)
	GetSizedValue=SrcValue*DstLength\SrcLength
end function

function Hex2Dec(HexStr)
	dim Result,i,ch
	Result=0
	for i=1 to len(trim(HexStr))
		ch=mid(trim(HexStr),i,1)
		if ucase(ch)>="A" and ucase(ch)<="F" then
			Result=Result*16+ASC(Ucase(ch))-55
		elseif ucase(ch)>="0" and ucase(ch)<="9" then
			Result=Result*16+ASC(Ucase(ch))-48
		end if
	next
	Hex2Dec=Result
end function

sub deletemsg()
	dim PhotoID
	
	stats="删除照片成功"
	if isnull(request("photoid")) then
		errmsg=errmsg+"<br>"+"<li>提交的数据有错误！"
		founderr=true
		exit sub
	elseif not isInteger(request("photoid")) then
		errmsg=errmsg+"<br>"+"<li>提交的数据有错误！"
		founderr=true
		exit sub
	elseif not master and conn.execute("select u.photo_id from visual_photouser u inner join visual_photo p on u.photo_id=p.id where u.photo_id="&checkstr(trim(request("photoid")))&" and u.username='"&membername&"' and p.finished").eof then
		errmsg=errmsg+"<br>"+"<li>提交的数据有错误！"
		founderr=true
		exit sub
	else
		photoid=checkstr(trim(request("photoid")))
	end if
	if not master then
		conn.execute("update visual_photouser set deleted=true where photo_id="&photoid&" and username='"&membername&"'")
	else
		conn.execute("update visual_photouser set deleted=true where photo_id="&photoid)
	end if
	if conn.execute("select photo_id from visual_photouser where photo_id="&photoid&" and not deleted").eof then
		dim objFSO
		set rs=server.createobject("adodb.recordset")
		sql="select URL,isUpload,Background from visual_photo where id="&photoid
		rs.open sql,conn,1,3
		if not isnull(rs(0)) then
			on error resume next
			Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
			if err=0 then
				if objFSO.fileExists(Server.MapPath(rs(0))) then
					objFSO.DeleteFile(Server.MapPath(rs(0)))
					sucmsg=sucmsg+"<br>"+"<li><b>文件 "&rs(0)&" 已经删除。</b>"
				end if
			end if
			set objFSO=nothing
			on error goto 0
		end if
		if rs(1) then
			on error resume next
			Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
			if err=0 then
				if objFSO.fileExists(Server.MapPath(PhotoPath&"/"&rs(2))) then
					objFSO.DeleteFile(Server.MapPath(PhotoPath&"/"&rs(2)))
					sucmsg=sucmsg+"<br>"+"<li><b>文件 "&PhotoPath&"/"&rs(2)&" 已经删除。</b>"
				end if
			end if
			set objFSO=nothing
			on error goto 0
		end if
		rs.close
		set rs=nothing
		conn.execute("delete from visual_photo where id="&photoid)
		conn.execute("delete from visual_photouser where photo_id="&photoid)
	end if
	sucmsg=sucmsg+"<br>"+"<li><b>恭喜您，删除照片成功。</b>"
	call dvbbs_suc()
end sub

sub privatephoto()
	dim PhotoID
	
	stats="隐藏照片成功"
	if isnull(request("photoid")) then
		errmsg=errmsg+"<br>"+"<li>提交的数据有错误！"
		founderr=true
		exit sub
	elseif not isInteger(request("photoid")) then
		errmsg=errmsg+"<br>"+"<li>提交的数据有错误！"
		founderr=true
		exit sub
	elseif not master and conn.execute("select u.photo_id from visual_photouser u inner join visual_photo p on u.photo_id=p.id where u.photo_id="&checkstr(trim(request("photoid")))&" and u.username='"&membername&"' and p.finished").eof then
		errmsg=errmsg+"<br>"+"<li>提交的数据有错误！"
		founderr=true
		exit sub
	else
		photoid=checkstr(trim(request("photoid")))
	end if
	if not master then
		conn.execute("update visual_photo set status=0 where id="&photoid&" and fromuser='"&membername&"' and finished")
	else
		conn.execute("update visual_photo set status=0 where id="&photoid&" and finished")
	end if
	sucmsg=sucmsg+"<br>"+"<li><b>恭喜您，隐藏照片成功。</b>"
	call dvbbs_suc()
end sub

sub publicphoto()
	dim PhotoID
	
	stats="公开照片成功"
	if isnull(request("photoid")) then
		errmsg=errmsg+"<br>"+"<li>提交的数据有错误！"
		founderr=true
		exit sub
	elseif not isInteger(request("photoid")) then
		errmsg=errmsg+"<br>"+"<li>提交的数据有错误！"
		founderr=true
		exit sub
	elseif not master and conn.execute("select u.photo_id from visual_photouser u inner join visual_photo p on u.photo_id=p.id where u.photo_id="&checkstr(trim(request("photoid")))&" and u.username='"&membername&"' and p.finished").eof then
		errmsg=errmsg+"<br>"+"<li>提交的数据有错误！"
		founderr=true
		exit sub
	else
		photoid=checkstr(trim(request("photoid")))
	end if
	if not master then
		conn.execute("update visual_photo set status=1 where id="&photoid&" and fromuser='"&membername&"' and finished")
	else
		conn.execute("update visual_photo set status=1 where id="&photoid&" and finished")
	end if
	sucmsg=sucmsg+"<br>"+"<li><b>恭喜您，公开照片成功。</b>"
	call dvbbs_suc()
end sub%>