<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<!--#include file="func.asp"-->
<!--#include file="../../mywp.asp"-->
<%'使出
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
%>
<!--#include file="pkif.asp"-->
<%
if Instr(LCase(Application("aqjh_useronlinename"&nowinroom)),LCase(" "&aqjh_name&" "))=0 then 
	Session("aqjh_inthechat")="0" 
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
says=Replace(says,"&amp;","&")
'对系统禁止字符处理
if aqjh_grade<9 then
says=bdsays(says)
end if
act=1
towhoway=0
i=instr(says,"$")
fnn1=mid(says,i+1)
says=jiuqingdao(mid(says,i+1),towho)
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)

'使出
function jiuqingdao(fn1,to1)
fn1=trim(fn1)
if InStr(fn1,"'")<>0 or InStr(fn1,"`")<>0 or InStr(fn1,"=")<>0 or InStr(fn1,"-")<>0 or InStr(fn1,",")<>0 then 
	Response.Write "<script language=JavaScript>{alert('滚吧，你想做什么？想捣乱吗？！');}</script>"
	Response.End 
end if
if aqjh_name=to1 and instr(";七伤拳;盗取令;天堂令;九阳神功;玄冥棒;铁笔银勾;圣火令;",fn1)<>0 then
	Response.Write "<script language=javascript>alert('【"&fn1&"】不能自己进行操作！');</script>"
	Response.End
end if
if to1="大家" and instr("七伤拳;盗取令;天堂令;九阳神功;玄冥棒;铁笔银勾;圣火令;",fn1)<>0 then
	Response.Write "<script language=javascript>alert('【"&fn1&"】不能大家进行操作！');</script>"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select 门派,等级,保护,grade,宝物,死亡时间,登录 from 用户 where 姓名='" & to1 &"'",conn,2,2
if rs("门派")="出家" and aqjh_grade<10 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：你想对出家人做什么，KKK！');}</script>"
	Response.End
end if
dlsj=DateDiff("n",rs("登录"),now())
if dlsj<3 and aqjh_grade<10 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('你不能对他进行攻击，他上线还没到3分钟就不要再连续使用法器了！');}</script>"
	Response.End
end if
sj=DateDiff("n",rs("死亡时间"),now())
if sj<15 and rs("宝物")="无" and aqjh_grade<10 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：他刚刚被别人杀死，还是先放过他吧！');}</script>"
	Response.End
end if
if rs("grade")>=6 and aqjh_grade<10 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：请不要对官府进行攻击！！');}</script>"
	Response.End
end if
if rs("等级")<=30 and rs("宝物")="无" and aqjh_grade<10 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：请不要对初入江湖新手操作！');}</script>"
	Response.End
end if
f=Minute(time())
if f<00 or f>25 then
	Response.Write "<script language=JavaScript>{alert('所有有动武倾向的都在每小时的前25分钟，现在是聊天时间！');window.close();}</script>"
	Response.End 
end if
if rs("保护")=True and aqjh_grade<10 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：对方正在练功保护请不要偷袭!');}</script>"
	Response.End
end if
rs.close
rs.open "select 法力,w10,登录,智力,门派 FROM 用户 WHERE 姓名='" & aqjh_name &"'" ,conn,2,2
fla=rs("法力")
if rs("法力")<3000 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：你的法力不够无法施展呀，至少也得3000点啊！');window.close();}</script>"
	response.end
end if
if rs("门派")="出家" and aqjh_grade<10 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：你是出家人啊，KKK！');}</script>"
	Response.End
end if
if rs("智力")<10 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：你的智力不够无法施展呀，至少也得10点啊！');window.close();}</script>"
	response.end
end if
dlsj=DateDiff("n",rs("登录"),now())
if dlsj<3 and aqjh_grade<10 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('你刚上线才不到3分钟，就不要再连续使用法器了！');}</script>"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
fn1=trim(fn1)
rs.open "select w10,门派,智力 from 用户 where  姓名='"&aqjh_name&"'",conn,3,3
mycard=abate(rs("w10"),fn1,1)
select case fn1
case "七伤拳"
        rs.close
        rs.open "SELECT * FROM 用户 WHERE  姓名='" & to1 & "'",conn
        if rs("门派")="官府" then 
            jiuqingdao="<font color=green>【七伤拳】<font color=" & saycolor & ">"&aqjh_name&",你不能对官府人员使用七伤拳!"
           	rs.close
			set rs=nothing
			conn.close
	    	set conn=nothing
	        exit function
		end if
        if rs("门派")="爱情训练营" then 
            jiuqingdao="<font color=green>【七伤拳】<font color=" & saycolor & ">"&aqjh_name&",你不能对爱情训练营的人使用!"
           	rs.close
			set rs=nothing
			conn.close
	    	set conn=nothing
	        exit function
		end if
		jiuqingdao="<bgsound src=wav/jiuzhao7.wav loop=1>【七伤拳】<font color=" & saycolor & ">"&aqjh_name&"使出七伤拳<img src='img/41.gif'>对"&to1&"攻击!内力和武功都减少一半，累得不轻呀"
        mzky=mzk(to1)
        if mzky="ok" then   
           conn.execute "update 用户 set  智力=智力-10,法力=int(法力/2),内力=int(内力/2),武功=int(武功/2),体力=int(体力/2) where 姓名='"&aqjh_name&"'"
           conn.execute "update 用户 set 状态='眠',体力=int(体力/2),登录=now()+3 where 姓名='" & to1 & "'"
            call boot(to1,to1&"被"&aqjh_name&"使用了七伤拳")  
        else
          jiuqingdao=jiuqingdao&mzky
        end if
        
case "盗取令"
        rs.close
        rs.open "SELECT * FROM 用户 WHERE  姓名='" & to1 & "'",conn
        money=int(rs("存款")/20)
        if rs("门派")="官府" then 
            jiuqingdao="<font color=green>【盗取令】<font color=" & saycolor & ">"&aqjh_name&",你不能对官府人员使用盗取令!"
           	rs.close
			set rs=nothing
			conn.close
	    	set conn=nothing
	        exit function
		end if
        if rs("门派")="爱情训练营" then 
            jiuqingdao="<font color=green>【盗取令】<font color=" & saycolor & ">"&aqjh_name&",你不能对爱情训练营的人使用!"
           	rs.close
			set rs=nothing
			conn.close
	    	set conn=nothing
	        exit function
		end if
        jiuqingdao="<bgsound src=wav/Bombs020.wav loop=1>【盗取令】<font color=" & saycolor & ">"&aqjh_name&"使出盗取令对"&to1&"偷取存款"&money&"!"
        mhky=mhk(to1)
        if mhky="ok" then   
          conn.execute "update 用户 set  智力=智力-10,银两=银两+"&money&",法力=法力-3000 where 姓名='"&aqjh_name&"'"
           conn.execute "update 用户 set 存款=存款-"&money&" where 姓名='" & to1 &"'"
        else
          jiuqingdao=jiuqingdao&mhky
        end if

case "天堂令"
        rs.close
        rs.open "SELECT * FROM 用户 WHERE  姓名='" & to1 & "'",conn
            if rs("门派")="官府" then 
            jiuqingdao="<font color=green>【天堂令】<font color=" & saycolor & ">"&aqjh_name&",你不能对官府人员使用天堂令!"
           	rs.close
			set rs=nothing
			conn.close
	    	set conn=nothing
	        exit function
		end if
        if rs("门派")="爱情训练营" then 
            jiuqingdao="<font color=green>【天堂令】<font color=" & saycolor & ">"&aqjh_name&",你不能对爱情训练营的人使用!"
           	rs.close
			set rs=nothing
			conn.close
	    	set conn=nothing
	        exit function
		end if
        jiuqingdao="<bgsound src=wav/Bombs020.wav loop=1>【天堂令】<font color=" & saycolor & ">"&aqjh_name&"使出天堂令让"&to1&"立刻睡觉!"
        mgky=mgk(to1)
        if mgky="ok" then   
          conn.execute "update 用户 set  智力=智力-10,法力=法力-3000 where 姓名='"&aqjh_name&"'"
          conn.execute "update 用户 set 状态='眠',登录=now()+1/144,事件原因='"&aqjh_name&" 眠:天堂令' where 姓名='" & to1 &"'"
          call boot(to1,"眠，操作者："&aqjh_name&","&fn1)
        else
          jiuqingdao=jiuqingdao&mgky
        end if

case "九阳神功"
        rs.close
        rs.open "SELECT * FROM 用户 WHERE  姓名='" & to1 & "'",conn
        if rs("门派")="官府" then 
            jiuqingdao="<font color=green>【九阳神功】<font color=" & saycolor & ">"&aqjh_name&",你不能对官府人员使用九阳神功!"
           	rs.close
			set rs=nothing
			conn.close
	    	set conn=nothing
	        exit function
		end if
        if rs("门派")="爱情训练营" then 
            jiuqingdao="<font color=green>【九阳神功】<font color=" & saycolor & ">"&aqjh_name&",你不能对爱情训练营的人使用!"
           	rs.close
			set rs=nothing
			conn.close
	    	set conn=nothing
	        exit function
		end if
	    jiuqingdao="<bgsound src=wav/jiuzhao7.wav loop=1>【九阳神功】<font color=" & saycolor & ">"&aqjh_name&"使出九阳神功<img src='img/41.gif'>对"&to1&"做猛烈攻击!法力和体力减少3000"
        mfky=mfk(to1)
        if mfky="ok" then   
           conn.execute "update 用户 set  智力=智力-10,法力=法力-3000,体力=体力-3000 where 姓名='"&aqjh_name&"'"
           conn.execute "update 用户 set 体力=int(体力/2),武功=int(武功/2),法力=int(法力/2) where 姓名='" & to1 & "'"
           call boot(to1,czsj)
  
        else
          jiuqingdao=jiuqingdao&mfky
        end if
        
case "圣火令"
        rs.close
        rs.open "SELECT * FROM 用户 WHERE  姓名='" & to1 & "'",conn
        if rs("门派")="官府" then 
            jiuqingdao="<font color=green>【圣火令】<font color=" & saycolor & ">"&aqjh_name&",你不能对官府人员使用圣火令!"
           	rs.close
			set rs=nothing
			conn.close
	    	set conn=nothing
	        exit function
		end if
        if rs("门派")="爱情训练营" then 
            jiuqingdao="<font color=green>【圣火令】<font color=" & saycolor & ">"&aqjh_name&",你不能对爱情训练营的人使用!"
           	rs.close
			set rs=nothing
			conn.close
	    	set conn=nothing
	        exit function
		end if
	    jiuqingdao="<bgsound src=wav/jiuzhao7.wav loop=1>【圣火令】<font color=" & saycolor & ">"&aqjh_name&"使出圣火令<img src='img/41.gif'>对"&to1&"做猛烈攻击想让"&to1&"变成僵尸王!"
        mdky=mdk(to1)
        if mdky="ok" then   
           conn.execute "update 用户 set  智力=智力-10,法力=法力-3000 where 姓名='"&aqjh_name&"'"
           conn.execute "update 用户 set 状态='僵尸王',身份1='僵尸王',登录=now()+3 where 姓名='" & to1 & "'"
           call boot(to1,czsj)
  
        else
          jiuqingdao=jiuqingdao&mdky
        end if

case "铁笔银勾"
        rs.close
        rs.open "SELECT * FROM 用户 WHERE  姓名='" & to1 & "'",conn
        if rs("门派")="官府" then 
            jiuqingdao="<font color=green>【铁笔银勾】<font color=" & saycolor & ">"&aqjh_name&",你不能对官府人员使用铁笔银勾!"
           	rs.close
			set rs=nothing
			conn.close
	    	set conn=nothing
	        exit function
		end if
        if rs("门派")="爱情训练营" then 
            jiuqingdao="<font color=green>【铁笔银勾】<font color=" & saycolor & ">"&aqjh_name&",你不能对爱情训练营的人使用!"
           	rs.close
			set rs=nothing
			conn.close
	    	set conn=nothing
	        exit function
		end if
	    jiuqingdao="<bgsound src=wav/jiuzhao7.wav loop=1>【铁笔银勾】<font color=" & saycolor & ">"&aqjh_name&"使出铁笔银勾<img src='img/41.gif'>对"&to1&"做猛烈攻击"&to1&"一笔进入天牢!不过自己确少了10000的积分，泡点也白泡了！"
        mfky=mfk(to1)
        if mfky="ok" then   
           conn.execute "update 用户 set  智力=智力-10,allvalue=allvalue-10000 where 姓名='"&aqjh_name&"'"
           conn.execute "update 用户 set 状态='牢',登录=now()+1/144,事件原因='"&aqjh_name&" 坐牢:"&fn1&"' where 姓名='" & to1 & "'"
           call boot(to1,"坐牢，操作者："&aqjh_name&","&fn1)
        else
          jiuqingdao=jiuqingdao&mfky
        end if
        

case "玄冥棒"
        rs.close
        rs.open "SELECT * FROM 用户 WHERE  姓名='" & to1 & "'",conn
        money=int(rs("内力")/5)
        if rs("门派")="官府" then 
            jiuqingdao="<font color=green>【玄冥棒】<font color=" & saycolor & ">"&aqjh_name&",你不能对官府人员使用盗取令!"
           	rs.close
			set rs=nothing
			conn.close
	    	set conn=nothing
	        exit function
		end if
        if rs("门派")="爱情训练营" then 
            jiuqingdao="<font color=green>【玄冥棒】<font color=" & saycolor & ">"&aqjh_name&",你不能对爱情训练营的人使用!"
           	rs.close
			set rs=nothing
			conn.close
	    	set conn=nothing
	        exit function
		end if
	    jiuqingdao="<bgsound src=wav/Bombs020.wav loop=1>【玄冥棒】<font color=" & saycolor & ">"&aqjh_name&"使出玄冥棒对"&to1&"一直吸取内力.慢慢的让"&to1&"油尽灯枯!"
        meky=mek(to1)
        if meky="ok" then   
          conn.execute "update 用户 set  智力=智力-10,内力=内力+"&money&",法力=法力-3000 where 姓名='"&aqjh_name&"'"
           conn.execute "update 用户 set 内力=内力-"&money&" where 姓名='" & to1 &"'"
        else
          jiuqingdao=jiuqingdao&meky
        end if
        
  
case else

	Response.Write "<script language=JavaScript>{alert('你并没有["&fn1&"]这种法器,若想再次使用,请在去寻法器！');}</script>"
	Response.End
end select

'删除自己法器，记录
conn.execute "update 用户 set w10='"&mycard&"' where  姓名='"&aqjh_name&"'"
conn.execute "insert into l(a,b,c,d,e) values (now(),'"& aqjh_name &"','"& to1 &"','操作','"& fn1 & "')"
set rs=nothing	
conn.close
set conn=nothing
end function
'佛山无影
function mzk(to1)
  Set conn=Server.CreateObject("ADODB.CONNECTION")
  Set rs=Server.CreateObject("ADODB.RecordSet")
  conn.open Application("aqjh_usermdb")
  rs.open "select w10 from 用户 where 姓名='"&to1&"'",conn,3,3
	if iswp(rs("w10"),"佛山无影")=0 then
		rs.close
	    mzk="ok"
	    exit function
	else
		tocard=abate(rs("w10"),"佛山无影",1)
		conn.execute "update 用户 set w10='"&tocard&"' where  姓名='"&to1&"'"
	   'mzk="<font color=green>【佛山无影】</font>"&to1&"身上有秘笈佛山无影,因此不能对他用七伤拳!"
	   mzk="<font color=green>"&to1&"使出【佛山无影】<font color=" & saycolor & ">"&"逃走嘿嘿想我死没那么容易哈哈!</font>"
	end if
	  rs.close
	  conn.close
	  set rs=nothing
	  set conn=nothing
end function

'凌波微步
function mhk(to1)
  Set conn=Server.CreateObject("ADODB.CONNECTION")
  Set rs=Server.CreateObject("ADODB.RecordSet")
  conn.open Application("aqjh_usermdb")
  rs.open "select w10 from 用户 where 姓名='"&to1&"'",conn,3,3
	if iswp(rs("w10"),"凌波微步")=0 then
		rs.close
	    mhk="ok"
	    exit function
	else
		tocard=abate(rs("w10"),"凌波微步",1)
		conn.execute "update 用户 set w10='"&tocard&"' where  姓名='"&to1&"'"
	   'mhk="<font color=green>【凌波微步】</font>"&to1&"身上有秘笈凌波微步,因此不能对他用盗取令!"
	   mhk="但没偷到"&to1&"使出【凌波微步】<font color=" & saycolor & ">"&"逃走嘿嘿想偷我的钱我看想都不要想真是受够你这种偷钱贼!</font>"
	end if
	  rs.close
	  conn.close
	  set rs=nothing
	  set conn=nothing
end function

'飞燕投林
function mgk(to1)
  Set conn=Server.CreateObject("ADODB.CONNECTION")
  Set rs=Server.CreateObject("ADODB.RecordSet")
  conn.open Application("aqjh_usermdb")
  rs.open "select w10 from 用户 where 姓名='"&to1&"'",conn,3,3
	if iswp(rs("w10"),"飞燕投林")=0 then
		rs.close
	    mgk="ok"
	    exit function
	else
		tocard=abate(rs("w10"),"飞燕投林",1)
		conn.execute "update 用户 set w10='"&tocard&"' where  姓名='"&to1&"'"
	   'mgk="<font color=green>【飞燕投林】</font>"&to1&"身上有秘笈天堂令,因此不能对他用天堂令!"
	   mgk="但没法使展"&to1&"使出【飞燕投林】<font color=" & saycolor & ">"&"逃走嘿嘿想让我睡觉嘿我就是不想睡.你能拿我这么样!</font>"
	end if
	  rs.close
	  conn.close
	  set rs=nothing
	  set conn=nothing
end function

'佛光化体
function mfk(to1)
  Set conn=Server.CreateObject("ADODB.CONNECTION")
  Set rs=Server.CreateObject("ADODB.RecordSet")
  conn.open Application("aqjh_usermdb")
  rs.open "select w10 from 用户 where 姓名='"&to1&"'",conn,3,3
	if iswp(rs("w10"),"佛光化体")=0 then
		rs.close
	    mfk="ok"
	    exit function
	else
		tocard=abate(rs("w10"),"佛光化体",1)
		conn.execute "update 用户 set w10='"&tocard&"' where  姓名='"&to1&"'"
	   'mfk="<font color=green>【佛光化体】</font>"&to1&"身上有秘笈佛光化体,因此不能对他用九阳神功!"
	   mfk="但无法攻击"&to1&"使出<font color=red>【佛光化体】</font><font color=" & saycolor & ">"&"逃走嘿嘿想把我赶尽杀决没那么容易.老子有佛光化体哈哈!</font>"
	end if
	  rs.close
	  conn.close
	  set rs=nothing
	  set conn=nothing
end function

'神行百变
function mek(to1)
  Set conn=Server.CreateObject("ADODB.CONNECTION")
  Set rs=Server.CreateObject("ADODB.RecordSet")
  conn.open Application("aqjh_usermdb")
  rs.open "select w10 from 用户 where 姓名='"&to1&"'",conn,3,3
	if iswp(rs("w10"),"神行百变")=0 then
		rs.close
	    mek="ok"
	    exit function
	else
		tocard=abate(rs("w10"),"神行百变",1)
		conn.execute "update 用户 set w10='"&tocard&"' where  姓名='"&to1&"'"
	   'mek="<font color=green>【神行百变】</font>"&to1&"身上有秘笈佛光化体,因此不能对他用九阳神功!"
	   mek="但无法使展玄冥棒"&to1&"使出<font color=red>【神行百变】</font><font color=" & saycolor & ">"&"逃走嘿嘿想把我内力吸光没那么容易.老子有伟小宝送的神行百变逃走功夫可是一流的哈.!</font>"
	end if
	  rs.close
	  conn.close
	  set rs=nothing
	  set conn=nothing
end function

'独步天下
function mdk(to1)
  Set conn=Server.CreateObject("ADODB.CONNECTION")
  Set rs=Server.CreateObject("ADODB.RecordSet")
  conn.open Application("aqjh_usermdb")
  rs.open "select w10 from 用户 where 姓名='"&to1&"'",conn,3,3
	if iswp(rs("w10"),"独步天下")=0 then
		rs.close
	    mdk="ok"
	    exit function
	else
		tocard=abate(rs("w10"),"独步天下",1)
		conn.execute "update 用户 set w10='"&tocard&"' where  姓名='"&to1&"'"
	   'mdk="<font color=green>【独步天下】</font>"&to1&"身上有秘笈独步天下,因此不能对他用圣火令!"
	   mdk="但无法使展圣火令"&to1&"使出<font color=red>【独步天下】</font><font color=" & saycolor & ">"&"逃走嘿嘿想把我变成僵尸王没那么容易.老子有小龙女送的独步天下逃走功夫可是一流的哈.!</font>"
	end if
	  rs.close
	  conn.close
	  set rs=nothing
	  set conn=nothing
end function

'醉佛无影
function mfk(to1)
  Set conn=Server.CreateObject("ADODB.CONNECTION")
  Set rs=Server.CreateObject("ADODB.RecordSet")
  conn.open Application("aqjh_usermdb")
  rs.open "select w10 from 用户 where 姓名='"&to1&"'",conn,3,3
	if iswp(rs("w10"),"醉佛无影")=0 then
		rs.close
	    mfk="ok"
	    exit function
	else
		tocard=abate(rs("w10"),"醉佛无影",1)
		conn.execute "update 用户 set w10='"&tocard&"' where  姓名='"&to1&"'"
	   'mfk="<font color=green>【铁笔银勾】</font>"&to1&"身上有秘笈醉佛无影,因此不能对他用铁笔银勾!"
	   mfk="但无法使展铁笔银勾"&to1&"使出<font color=red>【醉佛无影】</font><font color=" & saycolor & ">"&"逃走嘿嘿想让我坐牢没那么容易.老子有醉佛无影逃走功夫可是一流的哈.!</font>"
	end if
	  rs.close
	  conn.close
	  set rs=nothing
	  set conn=nothing
end function
%>醉佛无影】</font><font color=" & saycolor & ">"&"逃走嘿嘿想让我坐牢没那么容易.老子有醉佛无影逃走功夫可是一流的哈.!</font>"
	end if
	  rs.close
	  conn.close
	  set rs=nothing
	  set conn=nothing
end function
%>