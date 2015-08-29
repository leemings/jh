<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<!--#include file="func.asp"-->
<!--#include file="../../mywp.asp"-->
<!--#include file="chatconfig.asp"-->
<%'卡片♀wWw.51eline.com♀
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
'#####房间处理#####
sjjh_roominfo=split(Application("sjjh_room"),";")
chatinfo=split(sjjh_roominfo(nowinroom),"|")
if chatinfo(8)<>0 then
	Response.Write "<script language=JavaScript>{alert('提示：在["&chatinfo(0)&"]房间不可以使用卡片！');}</script>"
	Response.End
end if
if Instr(LCase(Application("sjjh_useronlinename"&nowinroom)),LCase(" "&sjjh_name&" "))=0 then 
	Session("sjjh_inthechat")="0" 
	Response.Redirect "chaterr.asp?id=001" 
end if 
towho=Trim(Request.Form("towho"))
towhoway=Request.Form("towhoway")
saycolor=Request.Form("saycolor")
addwordcolor=Request.Form("addwordcolor")
addsays=Request.Form("addsays")
says=Trim(Request.Form("sy"))
'对暂离开、点哑穴判断
call dianzan(towho)
if len(says)>150 then says=Left(says,150)
says=lcase(says)
says=Replace(says,"&amp;","")
'对系统禁止字符处理
if sjjh_grade<9 then
says=bdsays(says)
end if
act=1
towhoway=0
i=instr(says,"$")
fnn1=mid(says,i+1)
says=kapian(mid(says,i+1),towho)
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)

'卡片
function kapian(fn1,to1)
fn1=trim(fn1)
if InStr(fn1,"'")<>0 or InStr(fn1,"`")<>0 or InStr(fn1,"=")<>0 or InStr(fn1,"-")<>0 or InStr(fn1,",")<>0 then 
	Response.Write "<script language=JavaScript>{alert('滚吧，你想做什么？想捣乱吗？！');}</script>"
	Response.End 
end if
if sjjh_name=to1 and instr(";解除卡;陷害卡;查税卡;吃豆卡;捣乱卡;逮捕卡;踢人卡;冬眠卡;抢亲卡;情人卡;查税卡;穷神卡;衰神卡;死神卡;",fn1)<>0 then
	Response.Write "<script language=javascript>alert('【"&fn1&"】不能自己进行操作！');</script>"
	Response.End
end if
if to1="大家" and instr("逮捕卡;踢人卡;冬眠卡;抢亲卡;情人卡;查税卡;穷神卡;衰神卡;死神卡;",fn1)<>0 then
	Response.Write "<script language=javascript>alert('【"&fn1&"】不能大家进行操作！');</script>"
	Response.End
end if

Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
fn1=trim(fn1)
rs.open "select 会员金卡,w5,门派 from 用户 where  姓名='"&sjjh_name&"'",conn,2,2
mycard=abate(rs("w5"),fn1,1)
select case fn1
case "福神卡"
	conn.Execute ("update 用户 set sl='福神',slsj=now()+3 where 姓名='" & sjjh_name &"'")
	Response.Write "<script language=JavaScript>{alert('成功使用了["&fn1&"]现在神灵赋身了！');}</script>"
	kapian="<font color=green>【卡片】<font color=" & saycolor & ">##偷偷使用了["&fn1&"]，神啊，您与我同在...</font>"
case "财神卡"
	conn.Execute ("update 用户 set sl='财神',slsj=now()+3 where 姓名='" & sjjh_name &"'")
	Response.Write "<script language=JavaScript>{alert('成功使用了["&fn1&"]现在神灵赋身了！');}</script>"
	kapian="<font color=green>【卡片】<font color=" & saycolor & ">##偷偷使用了["&fn1&"]，神啊，您与我同在...</font>"
case "穷神卡"
    wnky=wnk(to1)
    if wnky<>"ok" then 
		kapian="<font color=green>【卡片】<font color=" & saycolor & ">##偷偷对%%使用了["&fn1&"]...</font>"&wnky
		exit function
    end if
	conn.Execute ("update 用户 set sl='穷神',slsj=now()+3 where 姓名='" & to1 &"'")
	Response.Write "<script language=JavaScript>{alert('成功使用了["&fn1&"]现在神灵与["&to1&"]赋身了！');}</script>"
	kapian="<font color=green>【卡片】<font color=" & saycolor & ">##偷偷对%%使用了["&fn1&"]，神啊，与你同在...</font>"
case "衰神卡"
    wnky=wnk(to1)
    if wnky<>"ok" then 
		kapian="<font color=green>【卡片】<font color=" & saycolor & ">##偷偷对%%使用了["&fn1&"]...</font>"&wnky
		exit function
    end if
	conn.Execute ("update 用户 set sl='衰神',slsj=now()+3 where 姓名='" & to1 &"'")
	Response.Write "<script language=JavaScript>{alert('成功使用了["&fn1&"]现在神灵与["&to1&"]赋身了！');}</script>"
	kapian="<font color=green>【卡片】<font color=" & saycolor & ">##偷偷对%%使用了["&fn1&"]，神啊，与你同在...</font>"
case "死神卡"
    wnky=wnk(to1)
    if wnky<>"ok" then 
		kapian="<font color=green>【卡片】<font color=" & saycolor & ">##偷偷对%%使用了["&fn1&"]...</font>"&wnky
		exit function
    end if
	conn.Execute ("update 用户 set sl='死神',slsj=now()+3 where 姓名='" & to1 &"'")
	Response.Write "<script language=JavaScript>{alert('成功使用了["&fn1&"]现在神灵与["&to1&"]赋身了！');}</script>"
	kapian="<font color=green>【卡片】<font color=" & saycolor & ">##偷偷对%%使用了["&fn1&"]，神啊，与你同在...</font>"
case "送神卡"
	conn.Execute ("update 用户 set sl='无',slsj=now() where 姓名='" & sjjh_name &"'")
	Response.Write "<script language=JavaScript>{alert('成功使用了["&fn1&"]成功解除了神灵赋身！');}</script>"
	kapian="<font color=green>【卡片】<font color=" & saycolor & ">##使用了["&fn1&"]，神啊，88！see tomorrow!</font>"
case "请神卡"
	dim sl(4)
	sl(0)="福神"
	sl(1)="财神"
	sl(2)="衰神"
	sl(3)="死神"
	sl(4)="穷神"
	randomize timer
	sss=int(rnd*4)+1
	if sss=5 then sss=4
	conn.Execute ("update 用户 set sl='"&sl(sss)&"',slsj=now()+3 where 姓名='" & sjjh_name &"'")
	Response.Write "<script language=JavaScript>{alert('成功使用了["&fn1&"]现在得到神灵["&sl(sss)&"]赋身！');}</script>"
	kapian="<font color=green>【卡片】<font color=" & saycolor & ">##偷偷使用了["&fn1&"]，神啊，您与我同在...请得大神["&sl(sss)&"]到自己的身上…</font>"
	erase sl
case "解除卡"
    wnky=wnk(to1)
    if wnky<>"ok" then 
		kapian="<font color=green>【卡片】<font color=" & saycolor & ">##偷偷对%%使用了["&fn1&"]...</font>"&wnky
		exit function
    end if
	conn.Execute ("update 用户 set 保护=false,操作时间=now() where 姓名='" & to1 &"'")
	Response.Write "<script language=JavaScript>{alert('成功解决了["&to1&"]的练功状态！');}</script>"
	kapian="<font color=green>【卡片】<font color=" & saycolor & ">##偷偷使用了解除卡片，也不知道哪一位大虾要倒霉了...</font>"
case "陷害卡"
    wnky=wnk(to1)
    if wnky<>"ok" then 
		kapian="<font color=green>【卡片】<font color=" & saycolor & ">##偷偷对%%使用了["&fn1&"]...</font>"&wnky
		exit function
    end if
	conn.Execute ("update 用户 set 内力=int(内力/3),体力=int(体力/3) where 姓名='" & to1 &"'")
	Response.Write "<script language=JavaScript>{alert('["&to1&"]的体力体力只剩原来的1/3了！');}</script>"
	kapian="<font color=green>【卡片】<font color=" & saycolor & ">##终于忍不可忍，拿出自己陷害卡，向%%的头上打去,%%只觉眼前一黑，体力、内力损失大半..</font>"
case "补血卡"
	conn.Execute ("update 用户 set 体力=体力+50000,内力=内力+50000  where 姓名='" & sjjh_name &"'")
	Response.Write "<script language=JavaScript>{alert('["&sjjh_name&"]的体力、体力都增加了5万点！');}</script>"
	kapian="<font color=green>【卡片】<font color=" & saycolor & ">##使用了补血卡，这一下可是有福了，体力、内力暴涨5万点！</font>"
case "涨钱卡"
	conn.Execute ("update 用户 set 存款=存款+88880000  where  姓名='" & sjjh_name &"'")
	Response.Write "<script language=JavaScript>{alert('["&sjjh_name&"]的存款上涨了8888万！');}</script>"
	kapian="<font color=green>【卡片】<font color=" & saycolor & ">##使用了涨钱卡，自己的小荷包都装不下了，8888万呀！</font>"
case "练功卡"
	conn.Execute ("update 用户 set 武功=武功+10000  where  姓名='" & sjjh_name &"'")
	Response.Write "<script language=JavaScript>{alert('["&sjjh_name&"]使用了练功卡，武功上涨1万！');}</script>"
	kapian="<font color=green>【卡片】<font color=" & saycolor & ">##使用了练功卡，武功可是大幅度上涨，看来江湖又要不太平了！</font>"
case "加点卡"
	conn.Execute ("update 用户 set allvalue=allvalue+1000  where 姓名='" & sjjh_name &"'")
	Response.Write "<script language=JavaScript>{alert('["&sjjh_name&"]使用了加点卡，泡点数上涨1000点！');}</script>"
	kapian="<font color=green>【卡片】<font color=" & saycolor & ">##使用了加点卡，呵。。不用泡点也加点，真是有福气！</font>"
case "吃豆卡"
    wnky=wnk(to1)
    if wnky<>"ok" then 
		kapian="<font color=green>【卡片】<font color=" & saycolor & ">##偷偷对%%使用了["&fn1&"]...</font>"&wnky
		exit function
    end if
	conn.Execute ("update 用户 set 暴豆时间=now()  where 姓名='" & to1 &"'")
	Response.Write "<script language=JavaScript>{alert('对["&to1&"]使用了吃豆卡，他不能再暴豆了！');}</script>"
	kapian="<font color=green>【卡片】<font color=" & saycolor & ">##实在对%%的行为看不过去，使用了吃豆卡，%%大叫一声，晕死过去。豆豆没有了...</font>"
case "暴豆卡"
	conn.Execute ("update 用户 set 暴豆时间=now()  where 姓名='" & sjjh_name &"'")
	Response.Write "<script language=JavaScript>{alert('对["&sjjh_name&"]使用了暴豆卡，现在暴豆成功了！');}</script>"
	kapian="<font color=green>【卡片】<font color=" & saycolor & ">##:大叫着，我暴，我暴....从口袋中拿出暴豆卡，暴豆成功！</font>"
case "捣乱卡"
    wnky=wnk(to1)
    if wnky<>"ok" then 
		kapian="<font color=green>【卡片】<font color=" & saycolor & ">##偷偷对%%使用了["&fn1&"]...</font>"&wnky
		exit function
    end if
	conn.Execute ("update 用户 set 武功=int(武功/3)  where 姓名='" & to1 &"'")
	Response.Write "<script language=JavaScript>{alert('对["&to1&"]使用了捣乱卡，他武功只剩1/3了！');}</script>"
	kapian="<font color=green>【卡片】<font color=" & saycolor & ">##:大叫着，谁也不要拦着我，我要为民除害！使用了捣乱卡，"&to1&"的武功失去大半....</font>"
case "清纯卡"
	conn.Execute ("update 用户 set 结婚次数=0  where 姓名='" & sjjh_name &"'")
	Response.Write "<script language=JavaScript>{alert('您使用了清纯卡，结婚次数清0完成！');}</script>"
	kapian="<font color=green>【卡片】<font color=" & saycolor & ">##好怀念过去的时光呀……</font>"
case "变性卡"
	rs.close
	rs.open "select 性别,配偶 FROM 用户 WHERE 姓名='" & sjjh_name &"'",conn,2,2
	sex=rs("性别")
	pl=rs("配偶")
	rs.close
    if pl="无" then 
        if sex="男" then
               sql="update 用户 set 性别='女' WHERE 姓名='" & sjjh_name & "'"
               xb="漂亮的女生"
         end if 
          if sex="女" then 
              sql="update 用户 set 性别='男' WHERE 姓名='" & sjjh_name & "'"
              xb="英俊的男孩"
          end if
          Set Rs=conn.Execute(sql)  
            bianxi=sjjh_name&"使用变性卡后,终于如愿以尝,变成了"&xb&"!" 
        else
          bianxi="使用变性卡失败!原因:"&sjjh_name&"是有家室的人呢,怎么还想变性呀!这样不是乱套了!为了惩戒像"&sjjh_name&"你这种道德败坏的人,特此没收"&sjjh_name&"的变性卡!"
        end if
	kapian="<table border=0><tr><td><img src=card/bxk.jpg></td><td><font color=green>【卡片】<font color=" & saycolor & ">##：这年头，克隆技术真先进呀，我变，我变……(不是在变成克琳钝，在作变性手术！）<br>【结果】"&bianxi&"</font></td></tr></table>"

case "离婚卡"
	rs.close
	rs.open "select 配偶 FROM 用户 WHERE 姓名='" & sjjh_name &"'",conn,2,2
	peiou=rs("配偶")
	if peiou="无" then
		rs.close
		set rs=nothing	
		conn.close
		set conn=nothing
		Response.Write "<script language=javascript>alert('【"&sjjh_name&"】你有钱呀，根本没有配偶呀！');</script>"
		Response.End
	end if
	conn.Execute ("update 用户 set 配偶='无'  where 姓名='" & sjjh_name &"'")
	rs.close
	rs.open "select 配偶 FROM 用户 WHERE 姓名='"&peiou&"'",conn,2,2
	if not(Rs.Bof OR Rs.Eof) then
		conn.Execute ("update 用户 set 配偶='无'  where 姓名='"&peiou&"'")
	end if
	rs.close
	kapian="<table border=0><tr><td><img src=card/lifen.jpg></td><td><font color=green><bgsound src=003.mid loop=1>【卡片】<font color=" & saycolor & ">##：想前想后，经过自己一番思想斗争，使用了离婚卡,终于想好了与["&peiou&"]离婚了……</font></td></tr></table>"
case "抢亲卡"
    wnky=wnk(to1)
    if wnky<>"ok" then 
		kapian="<font color=green>【卡片】<font color=" & saycolor & ">##偷偷对%%使用了["&fn1&"]...</font>"&wnky
		exit function
    end if
    if rs("门派")="出家" then
		Response.Write "<script language=javascript>alert('【"&sjjh_name&"】你是出家人，搞错了！！');</script>"
		Response.End
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
    end if
    rs.close
    rs.open "select * from 用户 where 姓名='"&to1&"'",conn,2,2
    if rs("门派")="出家" then
		Response.Write "<script language=javascript>alert('【"&sjjh_name&"】人家是出家人，搞错了！！');</script>"
		Response.End
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
    end if
	sex=rs("性别")
    if rs("配偶")<>"无" then
	    kapian="<table border=0><tr><td><img src=card/qinren.jpg></td><td><font color=green><bgsound src=003.mid loop=1>【卡片】<font color=" & saycolor & ">"&sjjh_name&"对"&to1&"使用抢亲卡失败:原因是"&to1&"是有家室的人</font></td></tr></table>"
    	rs.close
		set rs=nothing
		conn.close
	    set conn=nothing
    	exit function
	end if
    if rs("门派")="官府" then 
           kapian="不可以对管理员使用抢亲卡……"
           	rs.close
			set rs=nothing
			conn.close
	    	set conn=nothing
	        exit function
    end if
    rs.close
    rs.open "select * from 用户 where 姓名='"&sjjh_name&"'",conn,2,2
    if rs("配偶")<>"无" then
        kapian="<table border=0><tr><td><img src=card/qinren.jpg></td><td><font color=green><bgsound src=003.mid loop=1>【卡片】<font color=" & saycolor & ">"&sjjh_name&"对"&to1&"使用抢亲卡失败:原因是自己已经是有家室的人了!</font></td></tr></table>"
       	rs.close
		set rs=nothing
		conn.close
	    set conn=nothing
        exit function
    end if
    if rs("性别")=sex then
        kapian="<table border=0><tr><td><img src=card/qinren.jpg></td><td><font color=green><bgsound src=003.mid loop=1>【卡片】<font color=" & saycolor & ">"&sjjh_name&"对"&to1&"使用抢亲卡失败:原因是"&to1&"与"&sjjh_name&"是同性!</font></td></tr></table>"
       	rs.close
		set rs=nothing
		conn.close
	    set conn=nothing
        exit function
    end if
    conn.execute "update 用户 set 配偶='"&sjjh_name&"' where 姓名='"&to1&"'"
    conn.execute "update 用户 set 配偶='"&to1&"' where 姓名='"&sjjh_name&"'"
    kapian="<table border=0><tr><td><img src=card/qinren.jpg></td><td><font color=green><bgsound src=003.mid loop=1>【卡片】<font color=" & saycolor & ">"&sjjh_name&"对"&to1&"使用抢亲卡,终于如愿以尝的与"&to1&"结为夫妇!</font></td></tr></table>"
case "分手卡"
	rs.close
	rs.open "select 情人 FROM 用户 WHERE 姓名='" & sjjh_name &"'",conn,2,2
	peiou=rs("情人")
	if peiou="无" then
		rs.close
		set rs=nothing	
		conn.close
		set conn=nothing
		Response.Write "<script language=javascript>alert('【"&sjjh_name&"】你有钱呀，根本没有情人呀！');</script>"
		Response.End
	end if
	conn.Execute ("update 用户 set 情人='无'  where 姓名='" & sjjh_name &"'")
	rs.close
	rs.open "select 情人 FROM 用户 WHERE 姓名='"&peiou&"'",conn,2,2
	if not(Rs.Bof OR Rs.Eof) then
		conn.Execute ("update 用户 set 情人='无'  where 姓名='"&peiou&"'")
	end if
	rs.close
	kapian="<table border=0><tr><td><img src=card/lifen.jpg></td><td><font color=green><bgsound src=003.mid loop=1>【卡片】<font color=" & saycolor & ">##：想前想后，经过自己一番思想斗争，使用了分手卡,终于想好了与["&peiou&"]分手了……</font></td></tr></table>"
case "情人卡"
    wnky=wnk(to1)
    if wnky<>"ok" then 
		kapian="<font color=green>【卡片】<font color=" & saycolor & ">##偷偷对%%使用了["&fn1&"]...</font>"&wnky
		exit function
    end if
    if rs("门派")="出家" then
		Response.Write "<script language=javascript>alert('【"&sjjh_name&"】你是出家人，搞错了！！');</script>"
		Response.End
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
    end if
    rs.close
    rs.open "select * from 用户 where 姓名='"&to1&"'",conn,2,2
    if rs("门派")="出家" then
		Response.Write "<script language=javascript>alert('【"&sjjh_name&"】人家是出家人，搞错了！！');</script>"
		Response.End
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
    end if
	sex=rs("性别")
    if rs("情人")<>"无" then
	    kapian="<table border=0><tr><td><img src=card/qinren.jpg></td><td><font color=green><bgsound src=003.mid loop=1>【卡片】<font color=" & saycolor & ">"&sjjh_name&"对"&to1&"使用情人卡失败:原因是"&to1&"是有情人的人了</font></td></tr></table>"
    	rs.close
		set rs=nothing
		conn.close
	    set conn=nothing
    	exit function
	end if
    if rs("门派")="官府" then 
           kapian="不可以对管理员使用情人卡……"
           	rs.close
			set rs=nothing
			conn.close
	    	set conn=nothing
	        exit function
    end if
    rs.close
    rs.open "select * from 用户 where 姓名='"&sjjh_name&"'",conn,2,2
    if rs("情人")<>"无" then
        kapian="<table border=0><tr><td><img src=card/qinren.jpg></td><td><font color=green><bgsound src=003.mid loop=1>【卡片】<font color=" & saycolor & ">"&sjjh_name&"对"&to1&"使用情人卡失败:原因是自己已经是有情人的人了!</font></td></tr></table>"
       	rs.close
		set rs=nothing
		conn.close
	    set conn=nothing
        exit function
    end if
    if rs("性别")=sex then
        kapian="<table border=0><tr><td><img src=card/qinren.jpg></td><td><font color=green><bgsound src=003.mid loop=1>【卡片】<font color=" & saycolor & ">"&sjjh_name&"对"&to1&"使用情人卡失败:原因是"&to1&"与"&sjjh_name&"是同性!</font></td></tr></table>"
       	rs.close
		set rs=nothing
		conn.close
	    set conn=nothing
        exit function
    end if
    conn.execute "update 用户 set 情人='"&sjjh_name&"' where 姓名='"&to1&"'"
    conn.execute "update 用户 set 情人='"&to1&"' where 姓名='"&sjjh_name&"'"
    kapian="<table border=0><tr><td><img src=card/qinren.jpg></td><td><font color=green><bgsound src=003.mid loop=1>【卡片】<font color=" & saycolor & ">"&sjjh_name&"对"&to1&"使用情人卡,终于如愿以尝的与"&to1&"结为情人!</font></td></tr></table>"
case "逮捕卡"
    	wnky=wnk(to1)
	    if wnky<>"ok" then 
			kapian="<font color=green>【卡片】<font color=" & saycolor & ">##偷偷对%%使用了["&fn1&"]...</font>"&wnky
			exit function
	    end if
        rs.close
        rs.open "SELECT * FROM 用户 WHERE  姓名='" & to1 & "'",conn,2,2
        if rs("门派")="官府" then 
            kapian="<table border=0><tr><td><img src=card/xianhao.jpg></td><td><font color=green><bgsound src=003.mid loop=1>【卡片】<font color=" & saycolor & ">"&sjjh_name&",你不能对官府人员使用逮捕令!</td></tr></table>"
           	rs.close
			set rs=nothing
			conn.close
	    	set conn=nothing
	        exit function
		end if
        kapian="<table border=0><tr><td><img src=card/xianhao.jpg></td><td><font color=green><bgsound src=003.mid loop=1>【卡片】<font color=" & saycolor & ">"&sjjh_name&"拿出江湖特许的逮捕令,把"&to1&"给抓了!</td></tr></table>"
        mzky=mzk(to1)
        if mzky="ok" then   
           conn.execute "update 用户 set 状态='狱',登录=now()+3 where 姓名='" & to1 & "'"
            call boot(to1,to1&"被"&sjjh_name&"使用了逮捕令")  
        else
           kapian=kapian&mzky
        end if
case "踢人卡"
    wnky=wnk(to1)
    if wnky<>"ok" then 
		kapian="<font color=green>【卡片】<font color=" & saycolor & ">##偷偷对%%使用了["&fn1&"]...</font>"&wnky
		exit function
    end if
      rs.close
      rs.open "SELECT * FROM 用户 WHERE  姓名='" & to1 & "'",conn,2,2
      if rs("门派")="官府" then 
        kapian="<table border=0><tr><td><img src=card/dz04.gif></td><td><font color=green><bgsound src=003.mid loop=1>【卡片】<font color=" & saycolor & ">"&sjjh_name&",你不能对官府人员使用踢人卡!</td></tr></table>"
       	rs.close
		set rs=nothing
		conn.close
	   	set conn=nothing
        exit function
	  end if
       mtky=mtk(to1)
       if mtky="ok" then   
	      kapian="<table border=0><tr><td><img src=card/dz04.gif></td><td><font color=green><bgsound src=003.mid loop=1>【卡片】<font color=" & saycolor & ">"&sjjh_name&"使用踢人卡,飞起一脚，结果把"&to1&"踢了出去!</td></tr></table>"
    	  call boot(to1,to1&"被"&sjjh_name&"使用了踢人卡")     
    	else
			kapian=mtky
    	end if
case "冬眠卡"      
  wnky=wnk(to1)
  if wnky<>"ok" then 
	kapian="<font color=green>【卡片】<font color=" & saycolor & ">##偷偷对%%使用了["&fn1&"]...</font>"&wnky
	exit function
  end if
  rs.close
  rs.open "SELECT * FROM 用户 WHERE  姓名='" & to1 & "'",conn,2,2
  if rs("门派")="官府" then
     kapian="<table border=0><tr><td><img src=card/shuimian.jpg></td><td><font color=green><bgsound src=003.mid loop=1>【卡片】<font color=" & saycolor & ">"&sjjh_name&",你不能对官府人员使用冬眠卡!</td></tr></table>"
       	rs.close
		set rs=nothing
		conn.close
	   	set conn=nothing
        exit function
  end if
   rs.close
   qxky=qxk(to1)
   if qxky="ok" then   
	   conn.execute "update 用户 set 登录=now()+12/(4*144),状态='眠' where 姓名='" & to1 & "'"
	   kapian="<table border=0><tr><td><img src=card/shuimian.jpg></td><td><font color=green><bgsound src=003.mid loop=1>【卡片】<font color=" & saycolor & ">"&sjjh_name&"对"&to1&"使用冬眠卡!使"&to1&"睡着了!</td></tr></table>"
	   call boot(to1,to1&"被"&sjjh_name&"使用了冬眠卡")
   else
   		kapian=qxky
   end if
case "查税卡" 
  wnky=wnk(to1)
  if wnky<>"ok" then 
	kapian="<font color=green>【卡片】<font color=" & saycolor & ">##偷偷对%%使用了["&fn1&"]...</font>"&wnky
	exit function
  end if
  rs.close   
  rs.open "SELECT * FROM 用户 WHERE  姓名='" & to1 & "'",conn,2,2
  if rs("门派")="官府" then
    kapian="<table border=0><tr><td><img src=card/chashui.jpg></td><td><font color=green><bgsound src=003.mid loop=1>【卡片】<font color=" & saycolor & ">"&sjjh_name&",你不能对官府人员使用查税卡!</td></tr></table>"
   	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
    exit function
  end if
    yl=rs("银两")+rs("存款")
    if yl<=10000 then
    	kapian="<table border=0><tr><td><img src=card/chashui.jpg></td><td><font color=green><bgsound src=003.mid loop=1>【卡片】<font color=" & saycolor & ">"&sjjh_name&"对"&to1&"使用查税卡失败:原因"&to1&"身上的银子小于10000两!</td></tr></table>"
	   	rs.close
		set rs=nothing
		conn.close
		set conn=nothing
	    exit function
	end if
    mhky=mhk(to1)
    if mhky="ok" then   
      yl=int(yl*0.02)
      conn.execute "update 用户 set 银两=银两+"&yl&" where 姓名='"&sjjh_name&"'"
      if rs("银两")>=rs("存款") then
        conn.execute "update 用户 set 银两=银两-"&yl&" where 姓名='"&to1&"'"
      else
        conn.execute "update 用户 set 存款=存款-"&yl&" where 姓名='"&to1&"'"
      end if
      kapian="<table border=0><tr><td><img src=card/chashui.jpg></td><td><font color=green><bgsound src=003.mid loop=1>【卡片】<font color=" & saycolor & ">"&sjjh_name&"使用查税卡,查得"&to1&"共"&yl&"两银子,全部归"&sjjh_name&"所有!</td></tr></table>"
    else
       kapian="<table border=0><tr><td><img src=card/chashui.jpg></td><td><font color=green><bgsound src=003.mid loop=1>【卡片】<font color=" & saycolor & ">"&sjjh_name&"对"&to1&"使用查税卡失败!</td></tr></table>"
       kapian=kapian&mhky
    end if
case "升级卡"
 rs.close
 rs.open "select allvalue,会员等级,等级 from 用户 where 姓名='"&sjjh_name&"'",conn,2,2
 if rs("会员等级")>2 or rs("等级")>90 then
		Response.Write "<script language=javascript>alert('提示：90级以上或3级以上会员无需使用升级卡！');</script>"
		Response.End
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
    end if
 jhjy=rs("allvalue")
 jhdj=int(sqr(jhjy/50))
 jhadd=((jhdj+1)*(jhdj+1)-jhdj*jhdj)*50
 jhdj=jhdj+1
 jhjy=jhjy+jhadd
 conn.Execute ("update 用户 set allvalue="&jhjy&",等级="&jhdj&" where 姓名='"&sjjh_name&"'")
 Response.Write "<script language=Javascript>{alert('对["&sjjh_name&"]使用了升级卡！江湖经验值上升了"&jhadd&"点，战斗等级上升1级，现为"&jhdj&"级');}</script>"
 kapian=sjjh_name&"使用了升级卡，"&sjjh_name&"的泡点增加"&jhadd&"点，战斗等级上升1级...真是好福气呀，不泡点也升级！"
case "健身卡"
 conn.Execute ("update 用户 set 武功加=武功加+500,内力加=内力加+500,体力加=体力加+500 where 姓名='"&sjjh_name&"'")
 Response.Write "<script language=Javascript>{alert('对["&sjjh_name&"]使用了健身卡，锻练成功！武功、内力、体力上限值各涨500点！！！');}</script>"
 kapian=sjjh_name&"大叫着，我要更强，我要更强....从口袋中拿出健身卡，在"&Application("sjjh_user")&"的精心调教下，"&sjjh_name&"经过一翻艰苦训练，武功、内力、体力上限值各涨500点！！！"
case else
	Response.Write "<script language=JavaScript>{alert('系统并没有["&fn1&"]这种卡片,或不能使用你搞错了！');}</script>"
	Response.End
end select

'删除自己卡片，记录
conn.execute "update 用户 set w5='"&mycard&"' where  姓名='"&sjjh_name&"'"
conn.execute "insert into l(a,b,c,d,e) values (now(),'"& sjjh_name &"','"& to1 &"','操作','"& fn1 & "')"
set rs=nothing	
conn.close
set conn=nothing
end function


'免税卡
function mhk(to1)
  Set conn=Server.CreateObject("ADODB.CONNECTION")
  Set rs=Server.CreateObject("ADODB.RecordSet")
  conn.open Application("sjjh_usermdb")
  rs.open "select w5 from 用户 where 姓名='"&to1&"'",conn,2,2
	if iswp(rs("w5"),"免税卡")=0 then
		rs.close
	    mhk="ok"
	    exit function
	else
		tocard=abate(rs("w5"),"免税卡",1)
		conn.execute "update 用户 set w5='"&tocard&"' where  姓名='"&to1&"'"
	   'mhk="<br><font color=green>【免税卡】</font>"&to1&"身上的免税卡生效,因此不能抓他!"
	   mhk="<table border=0><tr><td><img src=card/mhk.jpg></td><td><font color=green><bgsound src=003.mid loop=1>【卡片】<font color=" & saycolor & ">"&sjjh_name&"正准备喜滋滋的从"&to1&"的口袋中拿钱,就在此时,"&to1&"掏出身上的免税卡说,慢着,我身上有免税卡,嘿嘿!</td></tr></table>"
	  end if
	  rs.close
	  conn.close
	  set rs=nothing
	  set conn=nothing
end function

'免罪卡
function mzk(to1)
  Set conn=Server.CreateObject("ADODB.CONNECTION")
  Set rs=Server.CreateObject("ADODB.RecordSet")
  conn.open Application("sjjh_usermdb")
  rs.open "select w5 from 用户 where 姓名='"&to1&"'",conn,2,2
	if iswp(rs("w5"),"免罪卡")=0 then
		rs.close
	    mzk="ok"
	    exit function
	else
		tocard=abate(rs("w5"),"免罪卡",1)
		conn.execute "update 用户 set w5='"&tocard&"' where  姓名='"&to1&"'"
	   'mzk="<br><font color=green>【免罪卡】</font>"&to1&"身上的免罪卡生效,因此不能抓他!"
	   mzk="<table border=0><tr><td><img src=card/myk.jpg></td><td><font color=green><bgsound src=003.mid loop=1>【卡片】<font color=" & saycolor & ">"&"官府的人准备用根铁索套在"&to1&"的脖子上,就在此时,"&to1&"掏出身上的免罪卡说,慢着,我身上有免罪卡,嘿嘿!</font></td></tr></table>"
	end if
	  rs.close
	  conn.close
	  set rs=nothing
	  set conn=nothing
end function

'清醒卡
function qxk(to1)
  Set conn=Server.CreateObject("ADODB.CONNECTION")
  Set rs=Server.CreateObject("ADODB.RecordSet")
  conn.open Application("sjjh_usermdb")
  rs.open "select w5 from 用户 where 姓名='"&to1&"'",conn,2,2
	if iswp(rs("w5"),"清醒卡")=0 then
		rs.close
	    qxk="ok"
	    exit function
	else
		tocard=abate(rs("w5"),"清醒卡",1)
		conn.execute "update 用户 set w5='"&tocard&"' where  姓名='"&to1&"'"
	   qxk="<table border=0><tr><td><img src=card/mhk.jpg></td><td><font color=green><bgsound src=003.mid loop=1>【卡片】<font color=" & saycolor & ">"&sjjh_name&"拿出水晶球，睡吧，睡吧，在催眠……"&to1&"睁着个大眼睛，傻嘻嘻的看着他……在说我吗？"&to1&"掏出身上的清醒卡,慢着,我身上有清醒卡,嘿嘿!</td></tr></table>"
	  end if
	  rs.close
	  conn.close
	  set rs=nothing
	  set conn=nothing
end function
'免踢卡
function mtk(to1)
  Set conn=Server.CreateObject("ADODB.CONNECTION")
  Set rs=Server.CreateObject("ADODB.RecordSet")
  conn.open Application("sjjh_usermdb")
  rs.open "select w5 from 用户 where 姓名='"&to1&"'",conn,2,2
	if iswp(rs("w5"),"免踢卡")=0 then
		rs.close
	    mtk="ok"
	    exit function
	else
		tocard=abate(rs("w5"),"免踢卡",1)
		conn.execute "update 用户 set w5='"&tocard&"' where  姓名='"&to1&"'"
	   mtk="<table border=0><tr><td><img src=card/mhk.jpg></td><td><font color=green><bgsound src=003.mid loop=1>【卡片】<font color=" & saycolor & ">"&sjjh_name&"使出国产臭脚，准备来个国际行动，却不小心踢到石头……"&to1&"在一边嘿嘿的笑，就你呀，还要踢我，再来20年……</td></tr></table>"
	  end if
	  rs.close
	  conn.close
	  set rs=nothing
	  set conn=nothing
end function
'万能卡
function wnk(to1)
  Set conn=Server.CreateObject("ADODB.CONNECTION")
  Set rs=Server.CreateObject("ADODB.RecordSet")
  conn.open Application("sjjh_usermdb")
  rs.open "select w5 from 用户 where 姓名='"&to1&"'",conn,2,2
	if iswp(rs("w5"),"万能卡")=0 then
		rs.close
	    wnk="ok"
	    exit function
	else
		tocard=abate(rs("w5"),"万能卡",1)
		conn.execute "update 用户 set w5='"&tocard&"' where  姓名='"&to1&"'"
	   wnk="<table border=0><tr><td><img src=card/mhk.jpg></td><td><font color=green><bgsound src=003.mid loop=1>【卡片】<font color=" & saycolor & ">"&to1&"在一边嘿嘿的笑，就你呀，万能卡，万能卡，一卡在手，走遍天下……</td></tr></table>"
	  end if
	  rs.close
	  conn.close
	  set rs=nothing
	  set conn=nothing
end function
%>
