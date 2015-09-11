<!--#include file="../sjfunc/sjfunc.asp"-->
<!--#include file="../sjfunc/func.asp"-->
<%
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
nowinroom=session("nowinroom")
aqjh_roominfo=split(Application("aqjh_room"),";")
chatinfo=split(aqjh_roominfo(nowinroom),"|")
if chatinfo(5)<>0 then
	Response.Write "<script language=JavaScript>{alert('提示：在["&chatinfo(0)&"]房间不可以发招！');}</script>"
	Response.End
end if
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
if trim(towho)="" or towho="大家" or towho=application("aqjh_automanname") or towho=aqjh_name then
	Response.Write "<script language=JavaScript>{alert('提示：你不可以对大家、机器人或自已进行此项操作！');}</script>"
	Response.End
end if
act=1
towhoway=0
i=instr(says,"$")
fnn1=mid(says,i+1)
says="<font color=green>【国家大战】<font color=" & saycolor & ">"+menattack(mid(says,i+1),towho)+"</font>"
call chatsay("发招",towhoway,towho,saycolor,addwordcolor,addsays,says)
'国家大战(比拚内力)
function menattack(fn1,to1)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
connstr=Application("aqjh_usermdb")
conn.open connstr
if Weekday(date())<>1 or (Hour(time())<18 or Hour(time())>=22) then
Response.Write "<script Language=Javascript>alert('提示：国家大战请于每周日晚上18-22点进行！比赛前一两天请先在群里（7872400）或论坛里公告，每周比赛结束后管理员会来查看那间门派战斗力剩下最多,战斗力最多者将可到本站总出的一星期泡点会员.！');parent.f2.document.af.mdsx.checked=true;parent.m.location.reload();</script>"
	Response.End
end if
	sql="select 体力,国家,武功,内力,防御,攻击,金币,等级,职位,门派 from 用户 where 姓名='" & aqjh_name &"'"
	rs.open sql,conn,1,1
       mp=rs("门派")
       guo=rs("国家")
       zhan1=rs("体力")
       gong1=rs("攻击")
       zhan1=rs("内力")
       wg=rs("武功")
       jin1=rs("金币")
       yx=rs("防御")
       zhiwei1=trim(rs("职位"))
       dengji1=rs("等级")
       if guo="无" or guo="" then
       Response.Write "<script language=JavaScript>{alert('您好像还没有加入到哪个国家吧！');}</script>"
       rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
       Response.End
       end if
   
     if mp="官府" then
       Response.Write "<script language=JavaScript>{alert('您是官府的人，怎么能参与国家之间的纷争呢！');}</script>"
       rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
       Response.End
       end if
         if zhan1<=50000 then
       Response.Write "<script language=JavaScript>{alert('您的战斗力怎么低,怎么替国家打仗啊！');}</script>"
       rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
       Response.End
       end if
       if zhiwei1="君主" then
          zhishu1=2
       elseif zhiwei1="太子" then
         zhishu1=1.8
       elseif zhiwei1="丞相"  then
         zhishu1=1.6
       elseif zhiwei1="将军"  then
          zhishu1=1.4
        elseif zhiwei1="侍卫"  then
          zhishu1=1.2
        else
           zhishu1=1
        end if
        rs.close
       sql="SELECT d,bh FROM guo WHERE  g='" & guo & "'"
       rs.open sql,conn,1,1
       if rs.eof or rs.bof then 
       Response.Write "<script language=JavaScript>{alert('对方好像不是国家中人');}</script>"
       rs.close
	set rs=nothing	
	conn.close
	set conn=nothing	
       Response.End
	end if
	   shidi1=rs("d")
	   baohu1=rs("bh")
	   if baohu1=1 then
       Response.Write "<script language=JavaScript>{alert('你们国家现在接受官府保护中,你无法向他开战');}</script>"
       rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
       Response.End
       end if
rs.close
       sql="select 体力,grade,国家,防御,金币,内力,等级,职位,门派 from 用户 where 姓名='" & to1 & "'"
       rs.open sql,conn,1,1
       topai=rs("门派")
       toguo=rs("国家")
       fang1=rs("防御")
       yin2=rs("金币")
       ti=rs("内力")
       dengji2=rs("等级")
       zhan2=rs("体力")
       zhiwei2=trim(rs("职位")) 
      
     
   if toguo="无" or toguo="" then
       Response.Write "<script language=JavaScript>{alert('对方好像没有国家！');}</script>"
       rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
       Response.End
       end if
 if topai="官府" then
       Response.Write "<script language=JavaScript>{alert('对方可是官府，你不想混了');}</script>"
       rs.close
	set rs=nothing	
	conn.close
	set conn=nothing	
       Response.End
	end if
       if InStr("|" & shidi1 & "|", "|" & toguo& "|")=0 and guo<>toguo then
       Response.Write "<script language=JavaScript>{alert('他们国家好像没有和你的国家开战呀,如果想开战向与君王请求');}</script>"
       rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
       Response.End
       end if
 
       
	 if zhiwei2="君主" then
          zhishu2=2
       elseif zhiwei2="太子" then
         zhishu2=1.8
       elseif zhiwei2="丞相"  then
         zhishu2=1.6
       elseif zhiwei2="将军"  then
          zhishu2=1.4
        elseif zhiwei2="侍卫"  then
          zhishu2=1.2
        else
           zhishu2=1
        end if
       rs.close
       sql="SELECT l,jin,d,bh FROM guo WHERE  g='" & toguo & "'"
       rs.open sql,conn,1,1
       if rs.eof or rs.bof then 
       Response.Write "<script language=JavaScript>{alert('对方好像不是国家中人');}</script>"
       rs.close
	set rs=nothing	
	conn.close
	set conn=nothing	
       Response.End
	end if
        shengming=rs("l")
        kunjin=int(rs("jin"))
        shidi2=rs("d")
         baohu2=rs("bh")
          if baohu2=1 then
       Response.Write "<script language=JavaScript>{alert('他们国家现在接受官府保护中,你无法向他开战');}</script>"
       rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
       Response.End
       end if
       if InStr("|" & shidi2 & "|", "|" & guo& "|")=0 and guo<>toguo then
       Response.Write "<script language=JavaScript>{alert('他们国家好像没有和你的国家开战呀,如果想开战向与君王请求');}</script>"
       rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
       Response.End
       end if
rs.close
'开始武功
rs.open "select * FROM y WHERE a='" & trim(fn1) & "' and b='" & mp & "'",conn
if rs.eof or rs.bof then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：你的门派:"& mp &" 并没有这样的武功["& fn1 &"]呀！');}</script>"
	Response.End
end if
nei=abs(rs("d"))
       if nei<100 then
       Response.Write "<script language=JavaScript>{alert('你这招也太弱了吧');}</script>"
	set rs=nothing	
	conn.close
	set conn=nothing	
       Response.End
	end if   
       if zhan1<nei or zhan1<=1000 then
       Response.Write "<script language=JavaScript>{alert('你的内力不足，小心虚脱，快去补补！');}</script>"
	set rs=nothing	
	conn.close
	set conn=nothing	
       Response.End
	end if 
 
	if mp=topai then
	    pang=2
	 else 
	    pang=1
	 end if         
       Randomize
	m1 = Int(20* Rnd + 1)
       Select Case m1
          case "1"
          e="[<a href=javascript:parent.sw('["&aqjh_name&"]'); target=f2>"&aqjh_name & "</a>]<bgsound src=READONLY/fazhao.mid loop=1>对" & to1 & "使用了一招" & mp & "的("& fn1 &")，恰好"&to1&"往旁边一晃，躲过了这一劫，好险哦!"
          baofa=0
          case"2"
          e="[<a href=javascript:parent.sw('["&aqjh_name&"]'); target=f2>"&aqjh_name & "</a>]<bgsound src=READONLY/fazhao.mid loop=1>乘着" & to1 & "不备，偷袭成功，一招" & mp & "的("& fn1 &")，重重的打在"&to1&"的背上，"
          baofa=2
          case"8"
          e="[<a href=javascript:parent.sw('["&aqjh_name&"]'); target=f2>"&aqjh_name & "</a>]<bgsound src=READONLY/fazhao.mid loop=1>突发神力，大喝一声，一招" & mp & "的("& fn1 &")打的"&to1&"狂吐鲜血，"
          baofa=3
          case"16"
          e="[<a href=javascript:parent.sw('["&aqjh_name&"]'); target=f2>"&aqjh_name & "</a>]<bgsound src=READONLY/fazhao.mid loop=1>偷袭"&to1&"，不料"&to1&"早有防备，反被"&to1&"的护体神功震伤，内力下降-5000。"
          baofa=0
          conn.execute"update 用户 set ,内力=内力-5000 where 姓名='" & aqjh_name & "'"
          case else
          baofa=1
          e="[<a href=javascript:parent.sw('["&aqjh_name&"]'); target=f2>"&aqjh_name & "</a>]<bgsound src=READONLY/fazhao.mid loop=1>扎起马步对"&to1&"使用了一招" & mp & "的("& fn1 &")，"
       end select
          if baofa=0 then
            kills=0 
            else
            kills1=nei*(gong1+1)*baofa+int(jin1/100)*dengji1     
            kills2=pang*zhishu1*kills1/zhishu2
            kills=int(kills2/(fang1+1))
            kills3=int(kills2/100)
          end if
           if kills<=100 then
	      randomize timer
	      kills=int(rnd*99)+1
          end if
	            	conn.execute"update guo set l=l-"& kills3 & " where g='" & toguo & "'"

	             	conn.execute"update 用户 set 体力=体力-"  & kills & " where 姓名='" & to1 & "'"
			         if shengming-kills3<=0 then
			              if guo=toguo then
			              conn.execute"update 用户 set 金币=金币+10,道德=道德+1000,魅力=魅力+1000,内力=内力*2,武功=武功*2 where 姓名='" & aqjh_name & "'"
			              conn.execute"update 用户 set 国家='无',职位='无',内力=内力*1/2,武功=武功*1/2 where 国家='" & toguo & "'"
                          conn.execute"delete * from guo  where g='" & toguo & "'"
                          call boot(to1,"国家大战，操作者："&aqjh_name&",["&guo&"]"&fn1)
                          conn.execute"insert into 人命(死者,时间,凶手,死因) values ('" & toguo & "',now(),'" &aqjh_name & "','被国家内奸出卖')"
                         

			              menattack=e&"杀伤"&to1&"战斗力为<b><font color=red>"&kills&"</font></b>,杀伤"&toguo&"国家战斗力为<b><font color=red>"&kills3&"</font></b>,并最终使该国家覆灭,"& aqjh_name &"作为卧底冲锋陷阵,道德上升1000,魅力上升1000,金币增加<b><font color=red>10</font></b>!"
                          else 
			            shidis=replace(shidi1,toguo&"|","")
			            conn.execute"update 用户 set 金币=金币+10,道德=道德+1000,魅力=魅力+1000 where 姓名='" & aqjh_name & "'"
                    	conn.execute"update 用户 set 国家='无',职位='无',内力=内力*1/2,武功=武功*1/2 where 国家='" & toguo & "'"
						conn.execute"delete * from guo  where g='" & toguo & "'"
						conn.execute"update 国家  set jin=jin+'"&kunjin&"',d='"&shidis&"' where g='" & guo & "'"
						call boot(to1,"国家大战，操作者："&aqjh_name&",["&guo&"]"&fn1)
                        conn.execute"insert into 人命(死者,时间,凶手,死因) values ('" & toguo & "',now(),'" & guo & "','国战时被敌人围剿覆灭')"
                        menattack=e&"杀伤"&to1&"战斗力为<b><font color=red>"&kills&"</font></b>,杀伤对方"&toguo&"国家战斗力为<b><font color=red>"&kills3&"</font></b>,并最终使该国家覆灭,"& aqjh_name &"为门派冲锋陷阵,道德上升1000,魅力上升1000,金币增加<b><font color=red>10</font></b>!"
						 end if
                             rs.close
                             set rs=nothing	
	                      conn.close
	                      set conn=nothing	
                             exit function
                    end if
                        conn.execute"update 用户 set 体力=体力-" & nei & ",道德=道德+2 where 姓名='" & aqjh_name & "'"   
                        menattack=e&"杀伤"&to1&"战斗力为<b><font color=red>"&kills&"</font></b>,杀伤对方"&toguo&"国家战斗力为<b><font color=red>"&kills3&"</font></b>,"& aqjh_name &"为国家冲锋陷阵,道德上升2!"
                       if zhan2-kills<1000 then
                       menattack=menattack&""&to1&"战斗力将为1000以下,已经丧失威胁性,不需要攻击了."
                       end if
                        rs.close
                        set rs=nothing	
	                 conn.close
                       set conn=nothing                              
          
end function
%>