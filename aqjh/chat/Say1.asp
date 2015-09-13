<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="npc_chat.asp"-->
<!--#include file="../const2.asp"-->
<!--#include file="../config.asp"-->
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
ysjh_name=Session("ysjh_name")
ysjh_grade=Session("ysjh_grade")
ysjh_jhdj=Session("ysjh_jhdj")
nowinroom=session("nowinroom")
ysjh_roominfo=split(Application("ysjh_room"),";")
chatinfo=split(ysjh_roominfo(nowinroom),"|")
if ysjh_name="" then Response.Redirect "../error.asp?id=440"
http = Request.ServerVariables("HTTP_REFERER") 
if InStr(http,"chat")=0 then 
	Response.Write "<script language=javascript>{alert('对不起，程序拒绝您的操作！！！\n     按确定返回！');}</script>" 
	Response.End 
end if
if Application("ysjh_yjdh")=1 then
ysjh_sid=trim(request.cookies("ysjh")("ysjh_sid")) 
if (ysjh_sid="" or  ysjh_sid<>session.sessionid) and Session("ysjh_grade")<6 then
Response.Write "<script language=javascript>{top.location.href='chaterr.asp?id=003';alert('您一台计算机上了多个帐号，被系统请出！');}</script>"
Response.End 
end if
end if
allhttp=LCase(Request.ServerVariables("ALL_HTTP"))
if Application("ysjh_disproxy")=1 then
if Instr(allhttp,"proxy")<>0 or Instr(allhttp,"http_via")<>0 or Instr(allhttp,"http_pragma")<>0 then 
	Session.Abandon
	Response.Write "<script language=JavaScript>{alert('提示：禁止使用代理！');location.href = 'javascript:history.go(-1)';}</script>"
	response.end
end if
end if
'#####房间处理#####
ysjh_roominfo=split(Application("ysjh_room"),";")
chatinfo=split(ysjh_roominfo(nowinroom),"|")
useronlinename=Application("ysjh_useronlinename"&nowinroom)
if Instr(LCase(useronlinename),LCase(" "&ysjh_name&" "))=0 then 
	Session("ysjh_inthechat")="0" 
	Response.Redirect "chaterr.asp?id=001" 
end if 
n=Year(date())
y=Month(date())
r=Day(date())
s=Hour(time())
f=Minute(time())
m=Second(time())
if len(y)=1 then y="0" & y
if len(r)=1 then r="0" & r
if len(s)=1 then s="0" & s
if len(f)=1 then f="0" & f
if len(m)=1 then m="0" & m
sj=s & ":" & f & ":" & m
if DateDiff("n",Session("ysjh_savetime"),now())>=15 then
	Response.Write "<script Language=Javascript>parent.f3.location.href='savevalue.asp';</script>"
end if
sj2=n & "-" & y & "-" & r & " " & sj
if DateDiff("s",Session("ysjh_lasttime"),sj2)<2 then
Response.Write "<script language=JavaScript>{alert('有话慢慢说，别噎着哦！');}</script>"
Response.End
end if
t="<font class=t>(" & sj & ")</font>"
'是否被点穴
Application.Lock
if Instr(LCase(application("ysjh_dianxuename")),LCase(ysjh_name))>0 then
dianxue=split(application("ysjh_dianxuename"),";")
dxdx=0
for x=0 to UBound(dianxue)-1
dxname=split(dianxue(x),"|")
if dxname(0)=ysjh_name then
	if DateDiff("s",dxname(1),sj2)<60 then
		Response.Write "<script Language=Javascript>alert('你已被哑穴，不能做任何事啊！还剩" & 60-DateDiff("s",dxname(1),sj2) & "秒自动解开!');</script>"
		erase dianxue
		erase dxname
		response.end
	else
		application("ysjh_dianxuename")=replace(application("ysjh_dianxuename"),dianxue(x)&";","")
		Response.Write "<script Language=Javascript>alert('时间到了，你的哑穴自动解开了!');</script>"
		erase dianxue
		erase dxname
		response.end
	end if
end if
erase dxname
next
erase dianxue
end if
Application.UnLock
towho=Trim(Request.Form("towho"))
says=Trim(Request.Form("sy"))
if Instr("/屏蔽$/发招$/投掷$/下毒$/用卡$",left(says,4))=0 and ysjh_grade<6 then
 if Instr(LCase(application("qzld_pingbi")),LCase(";"&ysjh_name&"|"&towho&";"))>0 then
   Response.Write "<script Language=javascript>alert('您已经把人家屏蔽了，要对人家说话请先把人家从屏蔽名单中删除！');</script>"
  response.end
 elseif Instr(LCase(application("qzld_pingbi")),LCase(";"&towho&"|"&ysjh_name&";"))>0 then
   Response.Write "<script Language=javascript>alert('你的说话对象已经把你屏蔽了，你不能再对人家说话了！');</script>"
  response.end
 end if
end if
'是否暂离
'判断自己说话
if Instr(LCase(application("ysjh_zanli")),LCase("!"&ysjh_name&"!"))>0 and ysjh_grade<9 then
	Response.Write "<script Language=Javascript>alert('您正在“暂离”状态中，请使用“我回来啦”功能解除！');</script>"
	response.end
end if
'判断对方操作
if Instr(LCase(application("ysjh_zanli")),LCase("!"&towho&"!"))>0 and left(says,4)<>"/用卡$" and ysjh_grade<9 then
	ysjh_zanli=split(application("ysjh_zanli"),";")
	for x=0 to UBound(ysjh_zanli)
		dxname=split(ysjh_zanli(x),"|")
	if dxname(0)="!"&towho&"!" then
		Response.Write "<script Language=Javascript>alert('"&towho&"说："&dxname(1)&"');</script>"
		erase dxname
		erase ysjh_zanli
		response.end
	end if
	erase dxname
	next
	erase ysjh_zanli
end if
if towho="" then towho="大家"
gw=left(towho,1)
if IsNumeric(gw)=true then
zz=split(gw,"级")
gw=1
else 
gw=0
end if  

if len(towho)>10 then towho=Left(towho,10)
if Not(towho=Application("ysjh_automanname") or towho=Application("figo") or towho="大家" or Instr(useronlinename," "&towho&" ")<>0) and gw=0 then
Response.Write "<script Language=Javascript>alert('“" & towho & "”不在聊天室中，不能对其发言。');parent.f2.document.af.towho.value='大家';parent.f2.document.af.towho.text='大家';parent.m.location.reload();</script>"
Response.end
end if
act=0
towhoway=Request.Form("towhoway")
if towhoway<>1 then towhoway=0
addwordcolor=Request.Form("addwordcolor")
saycolor=Request.Form("saycolor")
addsays=Request.Form("addsays")
'聊天室中发标题的等级72级。
if towho="大家" or towho=ysjh_name then towhoway=0
if instr("," & Application("ysjh_admin") & ",","," & towho & ",")<>0 or towho=Application("ysjh_automanname") or towho=Application("figo") then towhoway=1
if len(says)>150 then says=Left(says,150)
if (says="" or says="//") then Response.End
if InStr(says,"□")<>0 or InStr(says,"&#63736")<>0 then Response.End
says=Replace(says,"&amp;","&")
says=lcase(says)
says=Replace(says,"&amp;","&")
'私聊等级默认为7
says=lcase(says)
says=replace(says,"\","")
if ysjh_grade<9 then
says=replace(says,"&#","")
says=replace(says,"..","")
says=replace(says,"'","\'")
says=replace(says,chr(34),"\"&chr(34))
says=replace(says,chr(32),"")
says=replace(says,"　","")
says=Replace(says,"<"," ")
says=Replace(says,">"," ")
says=Replace(says,"\x3c"," ")
says=Replace(says,"\x3e"," ")
says=Replace(says,"\074"," ")
says=Replace(says,"\74"," ")
says=Replace(says,"\75"," ")
says=Replace(says,"\76"," ")
says=Replace(says,"&lt"," ")
says=Replace(says,"&gt"," ")
says=Replace(says,"\076"," ")
says=Replace(LCase(says),LCase("con"),"f2uck")
says=Replace(LCase(says),LCase("or"),"ｏｒ")
says=Replace(LCase(says),LCase("java"),"f2uck")
says=Replace(LCase(says),LCase("www"),"f2uck")
says=Replace(LCase(says),LCase("org"),"f2uck")
says=Replace(LCase(says),LCase("com"),"f2uck")
says=Replace(LCase(says),LCase("net"),"f2uck")
says=Replace(says,"QQ","f2uck")
says=Replace(says,"pp","f2uck")
says=Replace(says,"200.","f2uck")
says=Replace(says,"ＷＷＷ","f2uck")
says=Replace(says,"ｗｗｗ","f2uck")
says=Replace(says,"asp","f2uck")
says=Replace(says,"jh","f2uck")
says=Replace(says,"去了可以当","f2uck")
says=Replace(says,"去了可以作","f2uck")
says=Replace(says,"有一个好江湖","f2uck")
says=Replace(says,"加入腾讯","f2uck")
says=Replace(says,"加我qq给你好处","f2uck")
says=Replace(says,"家我qq给你好处","f2uck")
says=Replace(says,"加我q给你好处","f2uck")
says=Replace(says,"家我q给你好处","f2uck")
says=Replace(says,"江湖","f2uck")
says=Replace(says,"江胡","f2uck")
says=Replace(says,"我是站长","f2uck")
says=Replace(LCase(says),LCase("http"),"f2uck")
maren=instr(says,"f2uck")
if maren<>0 then
Response.Write "<script Language=Javascript>location.href='autokick.asp';</script>"
Response.end
end if
end if
if ysjh_jhdj<10 then
if InStr(says,"q")<>0 then Response.End
if InStr(says,"Q")<>0 then Response.End
if InStr(says,"ｑ")<>0 then Response.End
if InStr(says,"加我")<>0 then Response.End
end if
'文字秀
Zshow=request("Zshow")
if Zshow="1" then
show=ysjh_Zshow
says=Plus_ChkPicWords(says)
Function Plus_ChkPicWords(strText)
	If IsNull(StrText) Then Exit Function
        StrText=Ucase(StrText)
        StrText=replace(StrText,chr(32),"")
        if left(StrText,5)="[IMG]" then
           Response.Write "<script Language=Javascript>alert('使用贴图请先关闭文字秀功能！');parent.f2.document.af.sytemp.value='';</script>"
           Response.end
        end if
        if len(StrText)>35 then
           Response.Write "<script Language=Javascript>alert('为了江湖运行速度，文字秀不得超出50个！');parent.f2.document.af.sytemp.value='';</script>"
           Response.end
        end if
	Dim PicWords,i
	PicWords = Split(show,"|")
	For i = 0 To Ubound(PicWords)
	   StrText = Replace(trim(StrText),PicWords(i),"[img]../picwords/"&picwords(i)&".gif[/img]")
	Next
	Plus_ChkPicWords = StrText
End Function
end if
'文字秀
'是否战斗等级大于18才可以贴图
v0=says

If InStr(v0, "[flash]") <> 0 Then
if ysjh_jhdj<1 then
	Response.Write "<script language=JavaScript>{alert('提示：需要等级1级以上！');}</script>"
	Response.End
end if
if Instr(LCase(Application("ysjh_useronlinename"&nowinroom)),LCase(" "&ysjh_name&" "))=0 then 
	Session("ysjh_inthechat")="0" 
	Response.Redirect "chaterr.asp?id=001" 
end if 
        gsz = 0
        Do While InStr(v0, "[flash]") <> 0
			v0 = Replace(v0, "[flash]", "<object codebase='http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=5,0,0,0' height='250' width='250' onMouseOver='Play();' classid='clsid:D27CDB6E-AE6D-11cf-96B8-444553540000'><param name='movie' value='CW_swf/", 1, 1)
			gsz = gsz + 1
        Loop
        gsy = 0
        Do While InStr(v0, "[/flash]") <> 0
		v0 = Replace(v0, "[/flash]", "'><param name='quality' value='high'><param name='wmode' value='transparent'></object>", 1, 1)
		gsy = gsy + 1
        Loop
        If gsz > gsy Then v0 = v0 & ">"
End If
says=v0
dim re
	Set re=new RegExp
	re.IgnoreCase =true
	re.Global=True
're.Pattern="\[afaimg@*([0-9]*)\]"
'says=re.Replace(says,"&nbsp;<img src=QQ/gif/$1.gif border=0 alt='点击看我超酷表演' onload=ShowQQ('$1');  onmouseover=ShowQQ('$1'); style='cursor: hand;border:1px border-collapse: collapse; border-style: dotted;'>")
re.Pattern="\[afaimg@*([0-9]*)\]"
says=re.Replace(says,"&nbsp;<img src=QQ/gif/$1.gif border=0 alt='点击看我超酷表演' onload=ShowQQ('$1');  onmouseover=ShowQQ('$1'); style='cursor: hand;border:1px border-collapse: collapse; border-style: dotted;'>")
'says=re.Replace(says,"&nbsp;<img src=QQ/gif/$1.gif border=0 alt=点击看我超酷表演 onload=setTimeout(""ShowQQ('$1')"";,3000);  onmouseover=ShowQQ('$1'); style='cursor: hand;border:1px border-collapse: collapse; border-style: dotted;'>")
'v1=says
'If InStr(v1, "[afaimg@]") <> 0 Then

 '       gsz1 = 0
  '      Do While InStr(v1, "[afaimg@]") <> 0
			'v1 = Replace(v1, "[afaimg@]", "<img src=QQ/gif/", 1, 1)
			'gsz1 = gsz1 + 1
        'Loop
        'gsy1 = 0
        'Do While InStr(v1, "[/qq]") <> 0
		'v1 = Replace(v1, "[/qq]", ".gif border=0 alt=看我超酷表演 style='cursor: hand;border:1px border-collapse: collapse; border-style: dotted;' onmouseover=ShowQQ('001');>&nbsp;", 1, 1)
		'gsy1 = gsy1 + 1
        'Loop
        'If gsz1 > gsy1 Then v1 = v1 & ">"
'End If
'says=v1
if ysjh_jhdj>=18 and Instr(says," ")=0 and Instr(says,"[img]")<>0 then
if Instr(says,"[img]")<>0 then
if instr(says,"width")<>0 or instr(says,"height")<>0 then Response.End	
gsz=0
Do While Instr(says,"[img]")<>0
says=Replace(says,"[img]","<img src=pic/",1,1)
gsz=gsz+1
Loop
gsy=0
Do While Instr(says,"[/img]")<>0
says=Replace(says,"[/img]",">",1,1)
gsy=gsy+1
Loop
if gsz>gsy then says=says&">"
end if
end if
zj="<a href=javascript:parent.sw('[" & ysjh_name & "]'); target=f2><font color="&addwordcolor&">"& ysjh_name & "</font></a>"
br="<a href=javascript:parent.sw('[" & towho & "]'); target=f2><font color="&addwordcolor&">"& towho & "</font></a>"
if Left(says,2)="//" then
	says=Replace(says,"##",zj,1,3,1)
	says=Replace(says,"%%",br,1,3,1)
	says="<font color=" & sayscolor & ">" & zj & Trim(mid(says,3,len(says)-2)) & "</font>"
	act=1
end if
'删除了原来的表情
addsays=addsays
Session("ysjh_lasttime")=sj2
says=Replace(says,"\","\\")
says=Replace(says,"/","\/")
says=Replace(says,chr(34),"\"&chr(34))
if chatinfo(6)=0 then
'开始随机事件
randomize()
rnd1=int(rnd*400)+1
sayyq=""
'每小时15分可以发金币
if Minute(time())=1 then
if Application("ysjh_zzfafang")<>""  then
      zzfff=Application("ysjh_zzfafang")
       fffdata=split(zzfff,"|")
        ffsl=abs(int(clng(fffdata(0))))
        ffsj=fffdata(1)
       erase fffdata
    end if 
 if DateDiff("s",ffsj,now())>10 then
      sayyq="<bgsound src=wav/luo.wav loop=2><br><font color=red>【真情发放金币】每小时第一分钟发放5个金币，请需要的人来拿，时效60秒。</font><input  type=button value='领取金币！' onClick=javascript:zsdzff"&rnd1&".disabled=1;window.open('sjfunc/zz_fafang.asp','d') name=zsdzff"&rnd1&">"
      Application.Lock
      Application("ysjh_zzfafang")="5|"&now()
      application("ysjh_zzffjl")=""
      Application.UnLock
 end if
end if
if Minute(time())=15  then
if Application("ysjh_zzfafang1")<>""  then
      zzfff=Application("ysjh_zzfafang1")
       fffdata=split(zzfff,"|")
        ffsl=abs(int(clng(fffdata(0))))
        ffsj=fffdata(1)
       erase fffdata
    end if 
 if DateDiff("s",ffsj,now())>10 then
   sayyq="<bgsound src=wav/luo.wav loop=2><br><font color=red>【真情发放武功】每小时第15分钟发放武功，请需要的人来拿，时效60秒，不领取的属于自动放弃。</font><input  type=button value='领取武功！' onClick=javascript:zsdzff"&rnd1&".disabled=1;window.open('sjfunc/zdff1.asp','d') name=zsdzff"&rnd1&">"
      Application.Lock
      Application("ysjh_zzfafang1")="1|"&now()
      application("ysjh_zzffjl1")=""
      Application.UnLock
 end if
end if

if Minute(time())=30  then
if Application("ysjh_zzfafang2")<>""  then
      zzfff=Application("ysjh_zzfafang2")
       fffdata=split(zzfff,"|")
        ffsl=abs(int(clng(fffdata(0))))
        ffsj=fffdata(1)
       erase fffdata
    end if 
 if DateDiff("s",ffsj,now())>10 then
   sayyq="<bgsound src=wav/luo.wav loop=2><br><font color=red>【真情发放存点】每小时第30分钟发放存点，请需要的人来拿，时效60秒，不领取的属于自动放弃。</font><input  type=button value='领取存点！' onClick=javascript:zsdzff"&rnd1&".disabled=1;window.open('sjfunc/zdff2.asp','d') name=zsdzff"&rnd1&">"
      Application.Lock
      Application("ysjh_zzfafang2")="1|"&now()
      application("ysjh_zzffjl2")=""
      Application.UnLock
 end if
end if
if Minute(time())=45  then
if Application("ysjh_zzfafang3")<>""  then
      zzfff=Application("ysjh_zzfafang3")
       fffdata=split(zzfff,"|")
        ffsl=abs(int(clng(fffdata(0))))
        ffsj=fffdata(1)
       erase fffdata
    end if 
 if DateDiff("s",ffsj,now())>10 then
   sayyq="<bgsound src=wav/luo.wav loop=2><br><font color=red>【真情发放银子】每小时第45分钟发放银两，请需要的人来拿，时效60秒，不领取的属于自动放弃。</font><input  type=button value='点我发财！' onClick=javascript:zsdzff"&rnd3&".disabled=1;window.open('sjfunc/zdff3.asp','d') name=zsdzff"&rnd3&">"
      Application.Lock
      Application("ysjh_zzfafang3")="1|"&now()
      application("ysjh_zzffjl3")=""
      Application.UnLock
 end if
end if
'以下是设置运气
if rnd1<=16 then
	chance=array("体力","内力","攻击","防御","武功","银两","魅力","道德","体力","内力","攻击","防御","武功","银两","魅力","道德")
	chancetxtgood=array("遇少林达摩禅师开坛讲法，##体力增加","武当百岁真人张三丰提拨后进，##内力增加","见血河卫悲回重回人间，突有领悟，##攻击增加","在梦里见到嫦娥舞袖，领悟到神妙步法，##防御增加","##喝醉了酒，却意外见到侠客岛赏善罚恶二使见赐腊八粥，武功增加","##正在街上行走，观赏MM，却不料有人硬塞给银两","##正在华山后挖蚯蚓，做梦也想不到会见到传说中的剑侠风清扬，魅力因此而增加","在风波亭北拜谒岳王墓，##道德增加","遇少林达摩禅师开坛讲法，##体力增加","武当百岁真人张三丰提拨后进，##内力增加","见血河卫悲回重回人间，突有领悟，##攻击增加","在梦里见到嫦娥舞袖，领悟到神妙步法，##防御增加","##喝醉了酒，却意外见到侠客岛赏善罚恶二使见赐腊八粥，武功增加","##正在街上行走，观赏MM，却不料有人硬塞给银两","##正在华山后挖蚯蚓，做梦也想不到会见到传说中的剑侠风清扬，魅力因此而增加","在风波亭北拜谒岳王墓，##道德增加")
	chancetxtbad=array("##长时间熬战江湖,以致精力不济，走路也摔跤,体力因此而下降了","##在路上见到了一个美丽的小姑娘，顿起不良之心，不料却被阿紫偷偷地化去了内力","##闭关苦练躺尸剑法，却不得要领，攻击只此而减少","可笑##没事想学学荆无命的自残剑法，最后只落得满身伤疤，防御因此而减少","##闲来无事逛了逛翠花楼，结果可想而知，武功下降了","##交友不慎,银两被骗去了","##在练武场认识了薛蟠，大谈武道，魅力因此而减少了","##途见空空儿，苦心拜师学艺，不料功夫没有学到，道德却下降了","##长时间熬战江湖,以致精力不济，走路也摔跤,体力因此而下降了","##在路上见到了一个美丽的小姑娘，顿起不良之心，不料却被阿紫偷偷地化去了内力","##闭关苦练躺尸剑法，却不得要领，攻击只此而减少","可笑##没事想学学荆无命的自残剑法，最后只落得满身伤疤，防御因此而减少","##闲来无事逛了逛翠花楼，结果可想而知，武功下降了","##交友不慎,银两被骗去了","##在练武场认识了薛蟠，大谈武道，魅力因此而减少了","##途见空空儿，苦心拜师学艺，不料功夫没有学到，道德却下降了")
	rnd2=int(rnd*400)-100
	if rnd2>0 then
		rndtxt=chancetxtgood(rnd1)&"<font color=#FF8800><b>+"& abs(rnd2) & "</b></font>"
	else
		rnd2=rnd2-1
		rndtxt=chancetxtbad(rnd1)&"<font color=#FF8800><b>-"&abs(rnd2)&"</b></font>"
	end if	
	Set conn=Server.CreateObject("ADODB.CONNECTION")
	conn.open Application("ysjh_usermdb")
	conn.execute "update 用户 set "&chance(rnd1)&"="&chance(rnd1)&"+"&rnd2&" where 姓名='" & ysjh_name &"'"
	conn.close
	set conn=nothing
	zj="<font color="&addwordcolor&">"& ysjh_name & "</font>"
	sayyq="<bgsound src=wav/diaoxia.wav loop=1><font color=FF0000>【传闻】"&Replace(rndtxt,"##",zj,1,3,1)&"</font><font class=t>("&time&")</font>"
else
select case rnd1
case 18
banner="感谢大家支持『雨诗江湖网』.梦幻网络将会在这套江湖上更用心的做到最完美.如果有任何意见可随时提出来.本江湖的地址是 cys.hxjhw.com 别忘了叫你朋友一起来哦^_^; 一个最公平的江湖,在我们的江湖绝对不会乱调状态,为大家创造一个良好的休闲环境; 本站官府尽职尽责，如果有不满意的地方，请去论坛站务管理处投拆或发表意见！"
banners=Split(Trim(banner),";",-1)
total=UBound(banners)
randomize timer
x=int(rnd*(total+1))
sayyq="<bgsound src=wav/luo.wav loop=3><table align=center><td><table border=0 cellpadding=0 cellspacing=0 ><tr><td><img src=../chat/f2/img/rightct3.gif></td><td background=../chat/f2/img/rightct4.gif></td><td><img  src=../chat/f2/img/rightct1.gif></td></tr><tr><td background=../chat/f2/img/rightct8.gif ></td><td valign=center align=center><table align=center border=0 cellpadding=1 cellspacing=0 ><tr><td valign=center align=center><font style='font-size:9pt;color:red'>========官府公告========</font></td></tr><tr><td valign=center align=center><font style='font-size:9pt;color:green'>"&banners(x)&"</td></tr></table></td><td background=../chat/f2/img/rightct08.gif></td></tr><tr><td><img src=../chat/f2/img/rightct9.gif></td><td background=../chat/f2/img/rightct10.gif></td><td><img src=../chat/f2/img/rightct11.gif></td></tr></table></td></table>"
case 20,25
	useronlinename=Application("ysjh_useronlinename"&nowinroom)
	online=Split(Trim(useronlinename)," ",-1)
	x=UBound(online)
	Set conn=Server.CreateObject("ADODB.CONNECTION")
	conn.open Application("ysjh_usermdb")
        randomize timer
        rnd22=int(rnd*4)+1
        select case rnd22
        case 1
			randomize timer
			s=int(rnd*10000)
			for i=0 to x
				conn.Execute "update 用户 set 银两=银两+" & s & " where 姓名='" & online(i) & "'"
			next
			sayyq="<bgsound src=wav/zt.mid loop=1><font color=green>【系统发钱】</font><B><font color=#ff0000>发钱喽！发钱喽！<img src=img/251.gif><img src=img/251.gif><img src=img/251.gif> 聊天室里的每个人发到"& s &"两。本江湖地址→http://cys.hxjhw.com别忘了叫你朋友一起来哦 ^_^</font></b><br><marquee width=100% behavior=alternate scrollamount=15></marquee><img src=img/47.gif><img src=img/47.gif><img src=img/47.gif><img src=img/47.gif></marquee>"
        case 2
			for i=0 to x
				conn.Execute "update 用户 set 金币=金币+5 where 姓名='" & online(i) & "'"
			next
			sayyq="<bgsound src=wav/zt.mid loop=1><font color=green>【系统金币】</font><B><font color=#ff0000>站长梦幻网络感谢大家一直以来对『雨诗江湖网』的支持和厚爱,心中高兴，特为聊天室的每个人发了金币5个！</font></b><br><marquee width=100% behavior=alternate scrollamount=15></marquee><img src=img/jinbi.gif><img src=img/jinbi.gif><img src=img/jinbi.gif><img src=img/jinbi.gif><img src=img/jinbi.gif></marquee>"
        case 3
			mm=50
			for i=0 to x
				conn.Execute "update 用户 set allvalue=allvalue+" & mm & " where 姓名='" & online(i) & "'"
			next
			sayyq="<bgsound src=wav/zt.mid loop=1><font color=green>【系统经验】</font><font color=#ff0000>站长路过此地，不忍看到各位大小虾如此苦练，特为聊天室的每个人发放经验" & mm & "点！</font>"
        case 4
			randomize timer
			s=int(rnd*150)
			for i=0 to x
				conn.Execute "update 用户 set 泡豆点数=泡豆点数+" & s & " where 姓名='" & online(i) & "'"
			next
			sayyq="<bgsound src=wav/zt.mid loop=1><font color=green>【捡豆子】</font><font color=#ff0000>不知谁的袋子漏了，豆子撒了一地，在坐的各位每人捡到了" & s & "点！</font>"
        end select
case 5
Application("yx8_mhjh_klt1")=r
	sayyq="<font color=red>【运气】</font><font color=red>突然一只神鹿来偷袭<font color=#0000FF>" & ysjh_name & "</font>,却被<font color=#0000FF>" & ysjh_name & "</font>吸取体力"&Application("yx8_mhjh_klt1")*5000&"</font><bgsound src=../mid/Ye150.wav loop=3>"
Set conn=server.CreateObject("adodb.connection")
conn.open Application("ysjh_usermdb")
set rst = server.createobject("adodb.recordset") 
    conn.execute("update 用户 set 体力=体力+"&Application("yx8_mhjh_klt1")*5000&" where 姓名='" & ysjh_name & "'")

case 6
Application("yx8_mhjh_klt3")=r
	sayyq="<font color=red>【运气】</font><font color=red>突然一只庞大的恐怖飞兽闯进江湖，" & ysjh_name & "打伤他得到体力"&Application("yx8_mhjh_klt3")*2000&"！</font><bgsound src=../mid/Ye150.wav loop=3>"
Set conn=server.CreateObject("adodb.connection")
conn.open Application("ysjh_usermdb")
set rst = server.createobject("adodb.recordset") 
    conn.execute("update 用户 set 体力=体力+"&Application("yx8_mhjh_klt3")*2000&" where 姓名='" & ysjh_name & "'")
case 7
Application("yx8_mhjh_klt4")=r
	sayyq="<font color=red>【运气】</font><font color=red>聊天室突然一阵震荡，一只巨大的尸王蹒跚而来，<font color=#0000FF>" & ysjh_name & "</font>吸取他内力"&Application("yx8_mhjh_klt4")*1000&"！</font><bgsound src=../mid/Ye150.wav loop=3>"
Set conn=server.CreateObject("adodb.connection")
conn.open Application("ysjh_usermdb")
set rst = server.createobject("adodb.recordset") 
    conn.execute("update 用户 set 内力=内力+"&Application("yx8_mhjh_klt4")*1000&" where 姓名='" & ysjh_name & "'")
case 8
Application("yx8_mhjh_klt4")=r
	sayyq="<font color=red>【运气】</font><font color=red>一阵祥和的风儿吹来，一只修炼千年的神灯翩然而至，给<font color=#0000FF>" & ysjh_name & "</font>增加积分"&Application("yx8_mhjh_klt4")*20&"！</font><bgsound src=../mid/Ye150.wav loop=3>"
Set conn=server.CreateObject("adodb.connection")
conn.open Application("ysjh_usermdb")
set rst = server.createobject("adodb.recordset") 
    conn.execute("update 用户 set allvalue=allvalue+"&Application("yx8_mhjh_klt4")*20&" where 姓名='" & ysjh_name & "'")
case 9
Application("yx8_mhjh_klt6")=r
	sayyq="<font color=red>【运气】</font><font color=red>一股凄厉的寒风横扫江湖，寒冰卡比横行肆虐，大侠<font color=#0000FF>" & ysjh_name & "</font>打死他得到"&Application("yx8_mhjh_klt6")*100&"点防御！</font>"
	Set conn=server.CreateObject("adodb.connection")
	conn.open Application("ysjh_usermdb")
set rst = server.createobject("adodb.recordset") 
    conn.execute("update 用户 set 防御=防御+"&Application("yx8_mhjh_klt6")*100&" where 姓名='" & ysjh_name & "'")
case 10
Application("yx8_mhjh_klt6")=r
	sayyq="<font color=red>【运气】</font><font color=red>一股热浪突袭江湖，火龙闯入一群若女子当中，见人就咬，<font color=#0000FF>" & ysjh_name & "</font>打死他得到体力"&Application("yx8_mhjh_klt6")*2000&"点！</font><bgsound src=../mid/Ye150.wav loop=3>"
Set conn=server.CreateObject("adodb.connection")
conn.open Application("ysjh_usermdb")
set rst = server.createobject("adodb.recordset") 
    conn.execute("update 用户 set 体力=体力+"&Application("yx8_mhjh_klt6")*2000&" where 姓名='" & ysjh_name & "'")
case 11
Application("yx8_mhjh_klt7")=r
	sayyq="<font color=red>【运气】</font><font color=red>金奥克在深山被一股野火驱赶下山，游荡进"&Application("ysjh_chatroomname")&"，饥不择食，<font color=#0000FF>" & ysjh_name & "</font>打死他得到道德"&Application("yx8_mhjh_klt7")*500&"点！</font>"
Set conn=server.CreateObject("adodb.connection")
conn.open Application("ysjh_usermdb")
set rst = server.createobject("adodb.recordset") 
    conn.execute("update 用户 set 道德=道德+"&Application("yx8_mhjh_klt7")*500&" where 姓名='" & ysjh_name & "'")
case 12
Application("yx8_mhjh_klt8")=r
	sayyq="<font color=red>【运气】</font><font color=red>魔幻剑魂在哥特山顶苦心修炼，一股剑气伤了自己<font color=#0000FF>" & ysjh_name & "</font>救了他得到资质"&Application("yx8_mhjh_klt8")*100&"点！</font>"
Set conn=server.CreateObject("adodb.connection")
conn.open Application("ysjh_usermdb")
set rst = server.createobject("adodb.recordset") 
    conn.execute("update 用户 set 法力=法力+"&Application("yx8_mhjh_klt8")*100&" where 姓名='" & ysjh_name & "'")
case 13
Application("yx8_mhjh_klt9")=r
	sayyq="<font color=red>【运气】</font><font color=red>哇~~~好酷的一只鸭!真是林子大了，什么鸟都有，一只蹩脚鸭跑进"&Application("ysjh_chatroomname")&"，<font color=#0000FF>" & ysjh_name & "</font>偷了他白银"&Application("yx8_mhjh_klt9")*5000&"点！</font>"
Set conn=server.CreateObject("adodb.connection")
conn.open Application("ysjh_usermdb")
set rst = server.createobject("adodb.recordset") 
    conn.execute("update 用户 set 银两=银两+"&Application("yx8_mhjh_klt9")*5000&" where 姓名='" & ysjh_name & "'")
case 14
Application("yx8_mhjh_klt10")=r
	sayyq="<font color=red>【运气】</font><font color=red>闪电龙突然冲进"&Application("ysjh_chatroomname")&"，一道霹雳把<font color=#0000FF>" & ysjh_name & "</font>打得浑身黑糊糊的，得到积分"&Application("yx8_mhjh_klt10")*10&"点！</font>"
Set conn=server.CreateObject("adodb.connection")
conn.open Application("ysjh_usermdb")
set rst = server.createobject("adodb.recordset") 
    conn.execute("update 用户 set allvalue=allvalue+"&Application("yx8_mhjh_klt10")*10&" where 姓名='" & ysjh_name & "'")
case 15
Application("yx8_mhjh_klt11")=r
	sayyq="<font color=red>【运气】</font><font color=red>一只百年不遇的江湖之宝梦精灵突然出现在聊天室内，送给<font color=#0000FF>" & ysjh_name & "</font>体力"&Application("yx8_mhjh_klt11")*4000&"点！</font><bgsound src=../mid/Ye150.wav loop=3>"
Set conn=server.CreateObject("adodb.connection")
conn.open Application("ysjh_usermdb")
set rst = server.createobject("adodb.recordset") 
    conn.execute("update 用户 set 体力=体力+"&Application("yx8_mhjh_klt11")*4000&" where 姓名='" & ysjh_name & "'")
case 16
Application("yx8_mhjh_klt12")=r
	sayyq="<font color=red>【运气】</font><font color=red>一只火史莱姆蹦蹦跳跳闯进"&Application("ysjh_chatroomname")&"，向<font color=#0000FF>" & ysjh_name & "</font>吸取他内力"&Application("yx8_mhjh_klt12")*1000&"点！</font>"
Set conn=server.CreateObject("adodb.connection")
conn.open Application("ysjh_usermdb")
set rst = server.createobject("adodb.recordset") 
    conn.execute("update 用户 set 内力=内力+"&Application("yx8_mhjh_klt12")*1000&" where 姓名='" & ysjh_name & "'")
case 17
Application("yx8_mhjh_klt13")=r
	sayyq="<font color=red>【鬼魂附体】</font><font color=red>大侠<font color=#0000FF>" & ysjh_name & "</font>修炼内功时不小心走火入魔，被刺球宝宝的魂灵附着到身上，体力增加"&Application("yx8_mhjh_klt13")*4000&"点！</font>"
Set conn=server.CreateObject("adodb.connection")
conn.open Application("ysjh_usermdb")
set rst = server.createobject("adodb.recordset") 
    conn.execute("update 用户 set 体力=体力+"&Application("yx8_mhjh_klt13")*4000&" where 姓名='" & ysjh_name & "'")
case 18
Application("yx8_mhjh_klt14")=r
	sayyq="<font color=red>【运气】</font><font color=red>突然一只怪物闯进聊天室!" & ysjh_name & "打死他得到体力"&Application("yx8_mhjh_klt14")*3000&"点！</font>"
Set conn=server.CreateObject("adodb.connection")
conn.open Application("ysjh_usermdb")
set rst = server.createobject("adodb.recordset") 
    conn.execute("update 用户 set 体力=体力+"&Application("yx8_mhjh_klt14")*3000&" where 姓名='" & ysjh_name & "'")
case 19
Application("yx8_mhjh_klt15")=r
	sayyq="<font color=red>【运气】</font><font color=red>好漂亮啊~~~~一个不知道什么东西变的舞男翩跹起舞,兴奋中丢失了二个金币,被" & ysjh_name & "得到！</font>"
Set conn=server.CreateObject("adodb.connection")
conn.open Application("ysjh_usermdb")
set rst = server.createobject("adodb.recordset") 
    conn.execute("update 用户 set 金币=金币+2 where 姓名='" & ysjh_name & "'")
case 20
Application("yx8_mhjh_klt16")=r
	sayyq="<font color=red>【运气】</font><font color=red>好漂亮啊！一个不知道什么东西变的舞男翩跹起舞,兴奋中丢失了一个金币,被" & ysjh_name & "得到！</font>"
Set conn=server.CreateObject("adodb.connection")
conn.open Application("ysjh_usermdb")
set rst = server.createobject("adodb.recordset") 
    conn.execute("update 用户 set 金币=金币+1 where 姓名='" & ysjh_name & "'")
case 21
Application("yx8_mhjh_klt17")=r
	sayyq="<font color=red>【运气】</font><font color=red>突然一只<font color=#0000FF>怪物</font>闯进来," & ysjh_name & "打死他得到内力"&Application("yx8_mhjh_klt17")*500&"点</font>"
Set conn=server.CreateObject("adodb.connection")
conn.open Application("ysjh_usermdb")
set rst = server.createobject("adodb.recordset") 
    conn.execute("update 用户 set 内力=内力+"&Application("yx8_mhjh_klt17")*500&" where 姓名='" & ysjh_name & "'")
case 22
Application("yx8_mhjh_klt18")=r
	sayyq="<font color=red>【运气】</font><font color=red>突然一只<font color=#0000FF>怪物</font>闯进来," & ysjh_name & "打死他得到攻击"&Application("yx8_mhjh_klt18")*10&"点</font>"
Set conn=server.CreateObject("adodb.connection")
conn.open Application("ysjh_usermdb")
set rst = server.createobject("adodb.recordset") 
    conn.execute("update 用户 set 攻击=攻击+"&Application("yx8_mhjh_klt18")*10&" where 姓名='" & ysjh_name & "'")
case 23
Application("yx8_mhjh_klt19")=r
	sayyq="<font color=red>【运气】</font><font color=red>突然一只<font color=#0000FF>怪物</font>闯进来," & ysjh_name & "打死他得到防御"&Application("yx8_mhjh_klt19")*10&"点</font>"
Set conn=server.CreateObject("adodb.connection")
conn.open Application("ysjh_usermdb")
set rst = server.createobject("adodb.recordset") 
    conn.execute("update 用户 set 防御=防御+"&Application("yx8_mhjh_klt19")*10&" where 姓名='" & ysjh_name & "'")
case 24
Application("yx8_mhjh_klt20")=r
	sayyq="<font color=red>【运气】</font><font color=red>突然一只<font color=#0000FF>怪物</font>闯进来," & ysjh_name & "打死他得到银两"&Application("yx8_mhjh_klt20")*1200&"点</font>"
Set conn=server.CreateObject("adodb.connection")
conn.open Application("ysjh_usermdb")
set rst = server.createobject("adodb.recordset") 
    conn.execute("update 用户 set 银两=银两+"&Application("yx8_mhjh_klt20")*1200&" where 姓名='" & ysjh_name & "'")
case 25
Application("yx8_mhjh_klt21")=r
	sayyq="<font color=red>【运气】</font><font color=red>突然一只<font color=#0000FF>怪物</font>闯进来," & ysjh_name & "打死他得到积分"&Application("yx8_mhjh_klt21")*5&"点</font>"
Set conn=server.CreateObject("adodb.connection")
conn.open Application("ysjh_usermdb")
set rst = server.createobject("adodb.recordset") 
    conn.execute("update 用户 set allvalue=allvalue+"&Application("yx8_mhjh_klt21")*5&" where 姓名='" & ysjh_name & "'")
case 26
Application("yx8_mhjh_klt22")=r
	sayyq="<font color=red>【运气】</font><font color=red>突然一只<font color=#0000FF>怪物</font>闯进来," & ysjh_name & "打死他得到道德"&Application("yx8_mhjh_klt22")*500&"点</font>"
Set conn=server.CreateObject("adodb.connection")
conn.open Application("ysjh_usermdb")
set rst = server.createobject("adodb.recordset") 
    conn.execute("update 用户 set 道德=道德+"&Application("yx8_mhjh_klt22")*500&" where 姓名='" & ysjh_name & "'")
case 27
Application("yx8_mhjh_klt23")=r
	sayyq="<font color=red>【运气】</font><font color=red>突然一只<font color=#0000FF>怪物</font>闯进来," & ysjh_name & "打死他得到体力"&Application("yx8_mhjh_klt23")*1900&"点</font>"
Set conn=server.CreateObject("adodb.connection")
conn.open Application("ysjh_usermdb")
set rst = server.createobject("adodb.recordset") 
    conn.execute("update 用户 set 体力=体力+"&Application("yx8_mhjh_klt23")*1900&" where 姓名='" & ysjh_name & "'")
case 28
Application("yx8_mhjh_klt24")=r
	sayyq="<font color=red>【运气】</font><font color=red>突然一只<font color=#0000FF>怪物</font>闯进来," & ysjh_name & "打死他得到体力"&Application("yx8_mhjh_klt24")*2300&"点</font>"
Set conn=server.CreateObject("adodb.connection")
conn.open Application("ysjh_usermdb")
set rst = server.createobject("adodb.recordset") 
    conn.execute("update 用户 set 体力=体力+"&Application("yx8_mhjh_klt24")*2300&" where 姓名='" & ysjh_name & "'")
case 29
Application("yx8_mhjh_klt25")=r
	sayyq="<font color=red>【运气】</font><font color=red>突然一只<font color=#0000FF>怪物</font>闯进来," & ysjh_name & "打死他得到体力"&Application("yx8_mhjh_klt25")*2000&"点</font>"
Set conn=server.CreateObject("adodb.connection")
conn.open Application("ysjh_usermdb")
set rst = server.createobject("adodb.recordset") 
    conn.execute("update 用户 set 体力=体力+"&Application("yx8_mhjh_klt25")*2000&" where 姓名='" & ysjh_name & "'")
case 30
Application("yx8_mhjh_klt26")=r
	sayyq="<font color=red>【运气】</font><font color=red>突然一只<font color=#0000FF>怪物</font>闯进来," & ysjh_name & "打死他得到体力"&Application("yx8_mhjh_klt26")*1400&"点</font>"
Set conn=server.CreateObject("adodb.connection")
conn.open Application("ysjh_usermdb")
set rst = server.createobject("adodb.recordset") 
    conn.execute("update 用户 set 体力=体力+"&Application("yx8_mhjh_klt26")*1400&" where 姓名='" & ysjh_name & "'")
case 31
Application("yx8_mhjh_klt27")=r
	sayyq="<font color=red>【运气】</font><font color=red>突然一只<font color=#0000FF>怪物</font>闯进来," & ysjh_name & "打死他得到体力"&Application("yx8_mhjh_klt27")*1100&"点</font>"
Set conn=server.CreateObject("adodb.connection")
conn.open Application("ysjh_usermdb")
set rst = server.createobject("adodb.recordset") 
    conn.execute("update 用户 set 体力=体力+"&Application("yx8_mhjh_klt27")*1100&" where 姓名='" & ysjh_name & "'")
case 32
Application("yx8_mhjh_klt28")=r
	sayyq="<font color=red>【运气】</font><font color=red>突然一只<font color=#0000FF>怪物</font>闯进来," & ysjh_name & "打死他得到银两"&Application("yx8_mhjh_klt28")*3000&"两</font>"
Set conn=server.CreateObject("adodb.connection")
conn.open Application("ysjh_usermdb")
set rst = server.createobject("adodb.recordset") 
    conn.execute("update 用户 set 银两=银两+"&Application("yx8_mhjh_klt28")*3000&" where 姓名='" & ysjh_name & "'")
case 33,36
	s=int(rnd*20000)
	Application.Lock
	Application("ysjh_blh")=s+10000
	Application.UnLock
	abc="<a href='fafang/blh.asp?tl="&Application("ysjh_blh")&"' target='d'><img src='img/blh.gif' border='0'></a>"
	sayyq="<bgsound src=wav/Gui.WAV loop=1>【消息】平地里突然刮起一阵腥风，一只<font color=red>白老虎</font>冲进了聊天室四处伤人。官府发出悬赏令，打死白老虎者有重奖。白老虎体力：<font color=red>"&s+10000&"</font><br><marquee width=100% behavior=alternate scrollamount=15>"&abc&"</marquee>"
case 35,39
	s=int(rnd*20000)
	Application.Lock
	Application("ysjh_jx")=s+17000
	Application.UnLock
	abc="<a href='fafang/jx.asp?tl="&Application("ysjh_jx")&"' target='d'><img src='img/jx.gif' border='0'></a>"
	sayyq="<bgsound src=wav/Gui.WAV loop=1>【消息】一只百年巨蝎修练成精，跑下山来伤人。大家快打呀......巨蝎体力："&s+17000&"</font><br><marquee width=100% behavior=alternate scrollamount=15>"&abc&"</marquee>"
case 41,43
	s=int(rnd*20000)
	Application.Lock
	Application("ysjh_hx")=s+10000
	Application.UnLock
	abc="<a href='fafang/hx.asp?tl="&Application("ysjh_hx")&"' target='d'><img src='img/hx.gif' border='0'></a>"
	sayyq="<bgsound src=wav/Gui.WAV loop=1>【消息】动物园饲养员玩忽职守，忘了关熊圈，一只<font color=red>黑熊</font>跑出了熊圈，大家快抓呀。黑熊生命力：</font>"&s+10000&"</font><br><marquee width=100% behavior=alternate scrollamount=15>"&abc&"</marquee>"
case 48,55
		jstl=int(rnd*50000)+10000
		Application.Lock
		Application("ysjh_kl")=jstl
		Application.UnLock
		abc="<a href='fafang/kl.asp?tl="&Application("ysjh_kl")&"' target='d'><img src='img/gui.GIF' border='0'></a>"
		sayyq="<bgsound src=wav/Gui.WAV loop=1><font color=red>【消息】突然间阴风陈陈……“僵尸呀！”一女子尖叫，一头僵尸闯进聊天室，僵尸体力：+"&jstl&"</font><br><marquee width=100% behavior=alternate scrollamount=15>"&abc&"</marquee>"
case 60,66,70
		Application.Lock
		Application("ysjh_dalie")="老虎"
		Application.UnLock
		sayyq="<bgsound src=wav/hu.WAV loop=1><font color=red>【消息】</font><font color=red>突然一只野兽<img src=img/laohu.gif>窜入江湖中伤人，请高手们快去打猎啊。</font>"
case 90,91,98,120
		jstl=int(rnd*50000)+100
		Application.Lock
		Application("ysjh_yb")=jstl
		Application.UnLock
		abc="<a href='fafang/yb.asp?tl="&Application("ysjh_yb")&"' target='d'><img src='img/251.GIF' border='0'></a>"
		sayyq="<font color=red>【消息】今天是什么日子，突然间从天上掉下来一个大元宝，价值：+"&jstl&"两!</font><br><marquee width=100% behavior=alternate scrollamount=15>"&abc&"</marquee><bgsound src=wav/diaoxia.wav loop=1>"
case 130,135
		jstl=int(rnd*50000)+10000
		Application.Lock
		Application("ysjh_yb")=jstl
		Application.UnLock
		abc="<a href='fafang/yb.asp?tl="&Application("ysjh_yb")&"' target='d'><img src='img/251.GIF' border='0'></a>"
		sayyq="<font color=red>【消息】今天真是好日子，"&Application("ysjh_user")&"中了江湖风采头等奖，500万呀！给聊天室兄弟们一个大红包，价值：+"&jstl&"两!</font><br><marquee width=100% behavior=alternate scrollamount=15>"&abc&"</marquee><bgsound src=wav/diaoxia.wav loop=1>"
case 160,170
		jstl=int(rnd*50000)+10000
		Application.Lock
		Application("ysjh_kl")=jstl
		Application.UnLock
		abc="<a href='fafang/boy.asp?tl="&Application("ysjh_kl")&"' target='d'><img src='img/p42.gif' border='0'></a>"
		sayyq="<bgsound src=wav/Gui.WAV loop=1><font color=red>【消息】一个衰哥听说这个江湖的美女特别多，跑到江湖想泡美女，大家准备好棒子打呀。打跑了官府有奖。。衰哥体力：+"&jstl&"</font><br><marquee width=100% behavior=alternate scrollamount=15>"&abc&"</marquee>"
case 180,182
		jstl=int(rnd*50000)+100000
		Application.Lock
		Application("ysjh_qi")=jstl
		Application.UnLock
		abc="<a href='fafang/qi.asp?tl="&Application("ysjh_qi")&"' target='d'><img src='img/tank.gif' border='0'></a>"
		sayyq="<bgsound src=wav/Bombs020.wav loop=1><font color=red>【消息】一把精制手枪闯进江湖，江湖人士看了目瞪口呆，也不知道这是什么东西。都先打了在说。。精制手枪体力：+"&jstl&"</font><br><marquee width=100% behavior=alternate scrollamount=15>"&abc&"</marquee>"
case 190
		rftl=int(rnd*50000)+10000
		Application.Lock
		Application("ysjh_kl")=rftl
		Application.UnLock
		abc="<a href='fafang/rf.asp?tl="&Application("ysjh_kl")&"' target='d'><img src='img/rf.GIF' border='0'></a>"
		sayyq="<bgsound src=wav/by.wma loop=1><font color=#FF6600>【江湖任务】捉拿人贩子:叶三娘(ye sanniang)：“可恶的人贩子，把我那可怜的孩儿拐到哪里去啦？要是再找不到的话，我就见一个孩子杀一个。”人贩子体力：+"&rftl&"</font><br><marquee width=100% behavior=alternate scrollamount=15>"&abc&"</marquee>"
case 200
		piaoke=int(rnd*50000)+10000
		Application.Lock
		Application("ysjh_kl")=piaoke
		Application.UnLock
		abc="<a href='fafang/piaoke.asp?tl="&Application("ysjh_kl")&"' target='d'><img src='img/piaoke.GIF' border='0'></a>"
		sayyq="<bgsound src=wav/xiaotou.wav loop=1><font color=#FF6600>【江湖任务】抓小偷任务开始：这年头江湖也不平静啊！居然有位贼头贼脑的小偷闯入江湖，谁把他逮住必有重谢！”</font><br><marquee width=100% behavior=alternate scrollamount=15>"&abc&"</marquee>"
case 201
         r=int(rnd*30)+1
Application.Lock
		Application("ysjh_klt1")=r
Application.UnLock
	        sayyq="<font color=green>【怪物偷袭】</font><font color=red>突然一只神鹿来偷袭<font color=#0000FF>" &ysjh_name& "</font>,吸取<font color=#0000FF>" &ysjh_name& "</font>体力"&Application("ysjh_klt1")*500&"</font><br><marquee width=100% behavior=alternate scrollamount=15><a href=npc/npc1.asp?r="&Application("ysjh_klt1")&" target=optfrm><img src=img/tr003.gif  border=0></a></marquee><bgsound src=../mid/KAI.WAV                loop=3>"
               Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ysjh_usermdb")
               conn.execute("update 用户 set 体力=体力-"&Application("ysjh_klt1")*50&" where 姓名='" &ysjh_name& "'")
case 203
         r=int(rnd*30)+1
Application.Lock
        Application("ysjh_klt3")=r
Application.UnLock
	sayyq="<font color=green>【怪物偷袭】</font><font color=red>突然一只庞大的恐怖飞兽闯进江湖，咬伤" &ysjh_name& "内力"&Application("ysjh_klt3")*200&"！</font><br><marquee width=100% behavior=alternate scrollamount=15><a href=npc/npc3.asp?r="&Application("ysjh_klt3")&" target=optfrm><img src=img/mon24.gif   border=0></a></marquee><bgsound src=../mid/KAI.WAV loop=3>"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ysjh_usermdb")
    conn.execute("update 用户 set 内力=内力-"&Application("ysjh_klt3")*200&" where 姓名='" &ysjh_name& "'")
case 204
r=int(rnd*30)+1
Application.Lock
Application("ysjh_klt6")=r
Application.UnLock
	sayyq="<font color=green>【怪物偷袭】</font><font color=red>一股凄厉的寒风横扫江湖，寒冰卡比横行肆虐，大侠<font color=#0000FF>" &ysjh_name& "</font>被咬去"&Application("ysjh_klt6")*10&"点防御！</font><br><marquee width=100% behavior=alternate scrollamount=15><a href=npc/npc6.asp?r="&Application("ysjh_klt6")&" target=optfrm><img src=img/mon15.gif   border=0></a></marquee><bgsound src=../mid/KAI.WAV loop=3>"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ysjh_usermdb")
    conn.execute("update 用户 set 防御=防御-"&Application("ysjh_klt6")*10&" where 姓名='" &ysjh_name& "'")
case 205
r=int(rnd*30)+1
Application.Lock
Application("ysjh_klt5")=r
Application.UnLock
	sayyq="<font color=green>【怪物呈祥】</font><font color=red>一阵祥和的风儿吹来，一只修炼千年的神灯翩然而至，给<font color=#0000FF>" & username & "</font>增加积分"&Application("ysjh_klt5")*20&"！</font><br><marquee width=100% behavior=alternate scrollamount=15><a href=npc/npc5.asp?r="&Application("ysjh_klt5")&" target=optfrm><img src=img/mon20.gif   border=0></a></marquee><bgsound src=../mid/KAI.WAV loop=3>"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ysjh_usermdb")
    conn.execute("update 用户 set allvalue=allvalue+"&Application("ysjh_klt5")*20&" where 姓名='" & username & "'")
case 206
r=int(rnd*30)+1
Application.Lock
Application("ysjh_klt7")=r
Application.UnLock
	sayyq="<font color=green>【怪物袭击】</font><font color=red>一股热浪突袭江湖，火龙闯入一群弱女子当中，见人就咬，<font color=#0000FF>" &ysjh_name& "</font>被咬伤体力"&Application("ysjh_klt7")*200&"点！</font><br><marquee width=100% behavior=alternate scrollamount=15><a href=npc/npc7.asp?r="&Application("ysjh_klt7")&" target=optfrm><img src=img/mon23.gif border=0></a></marquee><bgsound src=../mid/KAI.WAV loop=3>"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ysjh_usermdb")
    conn.execute("update 用户 set 体力=体力-"&Application("ysjh_klt7")*200&" where 姓名='" &ysjh_name& "'")
case 207
r=int(rnd*30)+1
Application.Lock
Application("ysjh_klt8")=r
Application.UnLock
	sayyq="<font color=green>【怪物袭击】</font><font color=red>金奥克在深山被一股野火驱赶下山，游荡进雨诗江湖，饥不择食，吸收<font color=#0000FF>" &ysjh_name& "</font>道德"&Application("ysjh_klt8")*50&"点！</font><br><marquee width=100% behavior=alternate scrollamount=15><a href=npc/npc8.asp?r="&Application("ysjh_klt8")&" target=optfrm><img src=img/mon28.gif border=0></a></marquee><bgsound src=../mid/KAI.WAV loop=3>"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ysjh_usermdb")
    conn.execute("update 用户 set 道德=道德-"&Application("ysjh_klt8")*50&" where 姓名='" &ysjh_name& "'")
case 208
r=int(rnd*30)+1
Application.Lock
Application("ysjh_klt9")=r
Application.UnLock
	sayyq="<font color=green>【怪物偷袭】</font><font color=red>魔幻剑魂在哥特山顶苦心修炼，一股剑气伤了<font color=#0000FF>" &ysjh_name& "</font>魅力"&Application("ysjh_klt9")*10&"点！</font><br><marquee width=100% behavior=alternate scrollamount=15><a href=npc/npc9.asp?r="&Application("ysjh_klt9")&" target=optfrm><img src=img/mon27.gif border=0></a></marquee><bgsound src=../mid/KAI.WAV loop=3>"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ysjh_usermdb")
    conn.execute("update 用户 set 魅力=魅力-"&Application("ysjh_klt9")*10&" where 姓名='" &ysjh_name& "'")
case 209
r=int(rnd*30)+1
Application.Lock
Application("ysjh_klt10")=r
Application.UnLock
	sayyq="<font color=green>【怪物偷袭】</font><font color=red>哇~~~好酷的一只鸭!真是林子大了，什么鸟都有，一只蹩脚鸭跑进雨诗江湖，偷了<font color=#0000FF>" &ysjh_name& "</font>白银"&Application("ysjh_klt10")*500&"点！</font><br><marquee width=100% behavior=alternate scrollamount=15><a href=npc/npc10.asp?r="&Application("ysjh_klt10")&" target=optfrm><img src=img/mon3.gif border=0></a></marquee><bgsound src=../mid/KAI.WAV loop=3>"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ysjh_usermdb")
    conn.execute("update 用户 set 银两=银两-"&Application("ysjh_klt10")*500&" where 姓名='" &ysjh_name& "'")
case 210
r=int(rnd*30)+1
Application.Lock
Application("ysjh_klt11")=r
Application.UnLock
	sayyq="<font color=green>【怪物偷袭】</font><font color=red>闪电龙突然冲进魔幻江湖，一道霹雳把<font color=#0000FF>" &ysjh_name& "</font>打得浑身黑糊糊的，并夺取积分"&Application("ysjh_klt11")*10&"点！</font><br><marquee width=100% behavior=alternate scrollamount=15><a href=npc/npc11.asp?r="&Application("ysjh_klt11")&" target=optfrm><img src=img/mon4.gif border=0></a></marquee><bgsound src=../mid/KAI.WAV loop=3>"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ysjh_usermdb")
    conn.execute("update 用户 set allvalue=allvalue-"&Application("ysjh_klt11")*10&" where 姓名='" &ysjh_name& "'")
case 211
r=int(rnd*30)+1
Application.Lock
Application("ysjh_klt12")=r
Application.UnLock
	sayyq="<font color=green>【怪物呈祥】</font><font color=red>一只百年不遇的江湖之宝梦精灵突然出现在聊天室内，送给<font color=#0000FF>" &ysjh_name& "</font>体力"&Application("ysjh_klt12")*400&"点！</font><br><marquee width=100% behavior=alternate scrollamount=15><a href=npc/npc12.asp?r="&Application("ysjh_klt12")&" target=optfrm><img src=img/mon10.gif border=0></a></marquee><bgsound src=../mid/KAI.WAV loop=3>"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ysjh_usermdb")
    conn.execute("update 用户 set 体力=体力+"&Application("ysjh_klt12")*400&" where 姓名='" &ysjh_name& "'")
case 212
r=int(rnd*30)+1
Application.Lock
Application("ysjh_klt13")=r
Application.UnLock
	sayyq="<font color=green>【怪物袭击】</font><font color=red>一只火史莱姆蹦蹦跳跳闯进雨诗江湖，向<font color=#0000FF>" &ysjh_name& "</font>挑衅似的吸取内力"&Application("ysjh_klt13")*100&"点！</font><br><marquee width=100% behavior=alternate scrollamount=15><a href=npc/npc13.asp?r="&Application("ysjh_klt13")&" target=optfrm><img src=img/mon9.gif border=0></a></marquee><bgsound src=../mid/KAI.WAV loop=3>"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ysjh_usermdb")
    conn.execute("update 用户 set 内力=内力-"&Application("ysjh_klt13")*100&" where 姓名='" &ysjh_name& "'")
case 213
r=int(rnd*30)+1
Application.Lock
Application("ysjh_klt14")=r
Application.UnLock
	sayyq="<font color=green>【鬼魂附体】</font><font color=red>大侠<font color=#0000FF>" &ysjh_name& "</font>修炼内功时不小心走火入魔，被刺球宝宝的魂灵附着到身上，被咬伤"&Application("ysjh_klt14")*400&"点！</font><br><marquee width=100% behavior=alternate scrollamount=15><a href=npc/npc14.asp?r="&Application("ysjh_klt14")&" target=optfrm><img src=img/mon8.gif border=0></a></marquee><bgsound src=../mid/KAI.WAV loop=3>"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ysjh_usermdb")
    conn.execute("update 用户 set 体力=体力-"&Application("ysjh_klt14")*400&" where 姓名='" &ysjh_name& "'")
case 214
r=int(rnd*30)+1
Application.Lock
Application("ysjh_klt15")=r
Application.UnLock
	sayyq="<font color=green>【怪物袭击】</font><font color=red>突然一只<font color=#0000FF>怪物</font>闯进来,吸取" &ysjh_name& "内力"&Application("ysjh_klt15")*50&"点</font><br><marquee width=100% behavior=alternate scrollamount=15><a href=npc/npc15.asp?r="&Application("ysjh_klt15")&" target=optfrm><img src=img/Shenmo.gif border=0></a></marquee><bgsound src=../mid/KAI.WAV loop=3>"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ysjh_usermdb")
    conn.execute("update 用户 set 内力=内力-"&Application("ysjh_klt15")*50&" where 姓名='" &ysjh_name& "'")
case 215
r=int(rnd*30)+1
Application.Lock
Application("ysjh_klt16")=r
Application.UnLock
	sayyq="<font color=green>【怪物袭击】</font><font color=red>突然一只<font color=#0000FF>怪物</font>闯进来,打掉" &ysjh_name& "攻击"&Application("ysjh_klt16")*10&"点</font><br><marquee width=100% behavior=alternate scrollamount=15><a href=npc/npc16.asp?r="&Application("ysjh_klt16")&" target=optfrm><img src=img/Ying.gif border=0></a></marquee><bgsound src=../mid/KAI.WAV loop=3>"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ysjh_usermdb")
    conn.execute("update 用户 set 攻击=攻击-"&Application("ysjh_klt16")*10&" where 姓名='" &ysjh_name& "'")
case 216
r=int(rnd*30)+1
Application.Lock
Application("ysjh_klt17")=r
Application.UnLock
	sayyq="<font color=green>【怪物袭击】</font><font color=red>突然一只<font color=#0000FF>怪物</font>闯进来,打掉" &ysjh_name& "防御"&Application("ysjh_klt17")*10&"点</font><br><marquee width=100% behavior=alternate scrollamount=15><a href=npc/npc17.asp?r="&Application("ysjh_klt17")&" target=optfrm><img src=img/Ying1.gif border=0></a></marquee><bgsound src=../mid/KAI.WAV loop=3>"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ysjh_usermdb")
    conn.execute("update 用户 set 防御=防御-"&Application("ysjh_klt17")*10&" where 姓名='" &ysjh_name& "'")
case 217
r=int(rnd*30)+1
Application.Lock
Application("ysjh_klt18")=r
Application.UnLock
	sayyq="<font color=green>【怪物袭击】</font><font color=red>突然一只<font color=#0000FF>怪物</font>闯进来,偷走" &ysjh_name& "积分"&Application("ysjh_klt18")*5&"点</font><br><marquee width=100% behavior=alternate scrollamount=15><a href=npc/npc18.asp?r="&Application("ysjh_klt18")&" target=optfrm><img src=img/Ying1.gif border=0></a></marquee><bgsound src=../mid/KAI.WAV loop=3>"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ysjh_usermdb") 
    conn.execute("update 用户 set allvalue=allvalue-"&Application("ysjh_klt18")*5&" where 姓名='" &ysjh_name& "'")
case 218
r=int(rnd*30)+1
Application.Lock
Application("ysjh_klt19")=r
Application.UnLock
	sayyq="<font color=green>【怪物袭击】</font><font color=red>突然一只<font color=#0000FF>怪物</font>闯进来,降低" &ysjh_name& "道德"&Application("ysjh_klt19")*50&"点</font><br><marquee width=100% behavior=alternate scrollamount=15><a href=npc/npc19.asp?r="&Application("ysjh_klt19")&" target=optfrm><img src=img/tr003.gif border=0></a></marquee><bgsound src=../mid/KAI.WAV loop=3>"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ysjh_usermdb")
    conn.execute("update 用户 set 道德=道德-"&Application("ysjh_klt19")*50&" where 姓名='" &ysjh_name& "'")
case 230
if Application("ysjh_baowu")>0 and ysjh_grade<6 then 
   Set conn=Server.CreateObject("ADODB.CONNECTION") 
   Set rs=Server.CreateObject("ADODB.RecordSet") 
   conn.open Application("ysjh_usermdb") 
   rs.open "select id,姓名,宝物,等级,门派,身份 from 用户 where 宝物='" & Application("ysjh_baowuname") & "'",conn 
   if rs.eof or rs.bof  then 
    rs.close 
    rs.open "select 门派 from 用户 where 姓名='"&ysjh_name&"'" 
    if rs("门派")="出家" then 
     sayyq1="[江湖消息]：##在江湖道上偶然见到了江湖至宝<img src=img/z1.gif><font color=red>"& Application("ysjh_baowuname")&"</font>，但由于自己是个出家人，出家人戒贪，只好看了一眼宝物，悻悻然而去~~~！" 
     sayyq=Replace(sayyq1,"##",zj,1,3,1)  
    else 
     conn.execute  "update 用户 set 保护=false,宝物修练=0,宝物='"& Application("ysjh_baowuname") &"' where 姓名='" & ysjh_name &"'" 
     sayyq1="[江湖消息]：##现在拥有江湖至宝<img src=img/z1.gif><font color=red>"& Application("ysjh_baowuname")&"</font>" & Application("ysjh_baowusm") &"各位侠士互相争夺！" 
     sayyq=Replace(sayyq1,"##",zj,1,3,1)  
    end if 
   else 
    baouser=rs("姓名") 
    baowuid=rs("id") 
    if Instr(LCase(Application("ysjh_useronlinename"&nowinroom))," "&LCase(baouser)&" ")=0 then 
     conn.execute  "update 用户 set 宝物='无',宝物修练=0 where 姓名='"& baouser &"'" 
    else 
     sayyq1="[江湖消息]：江湖宝物<img src=img/z1.gif><font color=red>"& Application("ysjh_baowuname")&"</font>已被["&rs("门派")&"]"&rs("身份")&baouser&"夺走，各位大侠还不去抢..." 
     sayyq=Replace(sayyq1,"##",zj)  
    end if 
   end if 
   rs.close 
   set rs=nothing 
   conn.close 
   set conn=nothing 
  end if
case 233,235,238,250
'取江湖笑话
Set Conn=Server.CreateObject("ADODB.Connection")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ysjh_usermdb")
'取笑话的记录条数
sql="select count(id) from jokelist"
set rs=Conn.execute(sql)
jokecount=rs(0)
randomize
rand=Int(jokecount * Rnd +1 )
sql="select * from jokelist where id=" & rand & ""
set rs=Conn.execute(sql)
joke=rs("类型")&":"&rs("题目")&":"&rs("文章")
rs.close
set rs=nothing
conn.close
set conn=nothing
joke=replace(joke,chr(13),"")
joke=replace(joke,chr(10),"")
                Application.Lock
                Application("ysjh_joke")=joke
                Application.UnLock
                sayyq="<bgsound src=wav/haha.wav loop=1><img src=img/joke.GIF><font color=009933>〖开心一笑〗</font><img src=img/jokepic.GIF><font color=666666>" & Application("ysjh_joke") & "</font>"			'聊天数据
case 265,270
		jstl=int(rnd*2)+1
		Application.Lock
		Application("ysjh_jinbi")=jstl
		Application.UnLock
		abc="<a href='fafang/bxinshou.asp?tl="&Application("ysjh_jinbi")&"' target='d'><img src='img/xinshou.GIF' border='0'></a>"
		sayyq="<bgsound src=wav/xinshou.wav loop=1><font color=#000000>【照顾新手】</font><b><font color=red>有新朋友来江湖了，站长说：介绍朋友来、照顾新人者奖:+"&jstl&"个金币!</font></b><br><marquee width=100% behavior=alternate scrollamount=15>"&abc&"</marquee>"
case 280,289,295,305
'自动出题开始
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ysjh_usermdb")
'取题库记录数
sql="select count(id) from QuestLib "
set rs=Conn.execute(sql)
askcount=rs(0)
randomize
rand=Int(askcount * Rnd +1 )
sql="select * from QuestLib where id=" & rand & ""
set rs=Conn.execute(sql)
ask=rs("类型")&":"&rs("问题")
reply=rs("回答")
rs.close
set rs=nothing
conn.close
set conn=nothing
ask=replace(ask,chr(13),"")
ask=replace(ask,chr(10),"")
                Application.Lock
                Application("ysjh_ask")=ask
                Application("ysjh_reply")=reply
                Application("ysjh_askuser")="机器人"
                Application("ysjh_asksilver")=int(rnd*9999)+100000
                Application.UnLock
                sayyq="<bgsound src=wav/zhanfa.wav loop=1>【系统出题】<font color=balck>" & Application("ysjh_ask") & "？ </font>正确答案是什么？[提问人]:<font color=red>["&Application("ysjh_askuser")&"]</font><font color=green>奖励："&Application("ysjh_asksilver")&"两!</font>"
end select			
end if
end if
if sayyq<>"" then
	sayyq=replace(sayyq,"'","\'")
	sayyq=replace(sayyq,chr(34),"\"&chr(34))
	act="运气"
	saystr="<"&"script>parent.sh("& chr(39) & addwordcolor & chr(39) &","& chr(39) & saycolor & chr(39) &","& chr(39) & act & chr(39) &","& chr(39) & ysjh_name & chr(39) &","& chr(39) & addsays & chr(39) &","& chr(39) & towho & chr(39) &"," & chr(39) & sayyq & chr(39) &",0," & nowinroom & ");<"&"/script>"
	AddMsg SayStr
end if
act="正常"
says=replace(says,chr(39),"\'")
says=replace(says,chr(34),"\"&chr(34))
says=says&t
saystr="<"&"script>parent.sh("& chr(39) & addwordcolor & chr(39) &","& chr(39) & saycolor & chr(39) &","& chr(39) & act & chr(39) &","& chr(39) & ysjh_name & chr(39) &","& chr(39) & addsays & chr(39) &","& chr(39) & towho & chr(39) &"," & chr(39) & says & chr(39) &"," & towhoway & "," & nowinroom & ");"&npccls&"<"&"/script>"
AddMsg SayStr
For i = 1 to Application("SayCount")-Session("SayCount")
	Response.Write Application("SayStr"&YuShu((Session("SayCount")+i)))
Next
session("SayCount")=Application("SayCount")
Response.Write "<script Language=JavaScript>"
Response.Write "parent.f1.scrollWindow();"
Response.Write "</script>"
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
%>