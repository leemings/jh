<!--#include file="../sjfunc/sjfunc.asp"-->
<!--#include file="../sjfunc/func.asp"-->
<%'门派大战♀wWw.51eline.com♀
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
nowinroom=session("nowinroom")
sjjh_roominfo=split(Application("sjjh_room"),";")
chatinfo=split(sjjh_roominfo(nowinroom),"|")
if chatinfo(5)<>0 then
	Response.Write "<script language=JavaScript>{alert('提示：在["&chatinfo(0)&"]房间不可以发招！');}</script>"
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
if trim(towho)="" or towho="大家" or towho=application("sjjh_automanname") or towho=sjjh_name then
	Response.Write "<script language=JavaScript>{alert('提示：你不可以对大家、机器人或自已进行此项操作！');}</script>"
	Response.End
end if
act=1
towhoway=0
i=instr(says,"$")
fnn1=mid(says,i+1)
says="<font color=green>【门派大战】<font color=" & saycolor & ">"+menattack(mid(says,i+1),towho)+"</font>"
call chatsay("发招",towhoway,towho,saycolor,addwordcolor,addsays,says)
'门派大战(比拚内力)
function menattack(fn1,to1)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
connstr=Application("sjjh_usermdb")
conn.open connstr
if Weekday(date())<>6 or (Hour(time())<21 or Hour(time())>=22) then
Response.Write "<script Language=Javascript>alert('提示：门派大战请于每周五晚上21-22点进行！');parent.f2.document.af.mdsx.checked=true;parent.m.location.reload();</script>"
	Response.End
end if
	sql="select 体力,门派,武功,内力,防御,攻击,金币,等级,身份 from 用户 where 姓名='" & sjjh_name &"'"
	rs.open sql,conn,1,1
       mp=rs("门派")
       zhan1=rs("体力")
       gong1=rs("攻击")
       zhan1=rs("内力")
       wg=rs("武功")
       jin1=rs("金币")
       yx=rs("防御")
       shenfen1=trim(rs("身份"))
       dengji1=rs("等级")
       if mp="无" or mp="" then
       Response.Write "<script language=JavaScript>{alert('您好像不是门派中人！');}</script>"
       rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
       Response.End
       end if
   
     if mp="官府" then
       Response.Write "<script language=JavaScript>{alert('您是官府的人，怎么能跟江湖门派结怨呢！');}</script>"
       rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
       Response.End
       end if
         if zhan1<=1000 then
       Response.Write "<script language=JavaScript>{alert('您的战斗力这么低,怎么替门派出战啊！');}</script>"
       rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
       Response.End
       end if
       if shenfen1="掌门" then
          zhishu1=2
       elseif shenfen1="副掌门" then
         zhishu1=1.8
       elseif shenfen1="长老"  then
         zhishu1=1.6
       elseif shenfen1="护法"  then
          zhishu1=1.4
        elseif shenfen1="堂主"  then
          zhishu1=1.2
        else
           zhishu1=1
        end if
        rs.close
       sql="SELECT q,f FROM p WHERE  a='" & mp & "'"
       rs.open sql,conn,1,1
       if rs.eof or rs.bof then 
       Response.Write "<script language=JavaScript>{alert('对方好像不是门派中人');}</script>"
       rs.close
	set rs=nothing	
	conn.close
	set conn=nothing	
       Response.End
	end if
	   shidi1=rs("q")
	   baohu1=rs("f")
	   if baohu1=1 then
       Response.Write "<script language=JavaScript>{alert('你们门派现在接受官府保护中,你无法向他开战');}</script>"
       rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
       Response.End
       end if
rs.close
       sql="select 体力,grade,门派,防御,金币,内力,等级,身份 from 用户 where 姓名='" & to1 & "'"
       rs.open sql,conn,1,1
       topai=rs("门派")
       fang1=rs("防御")
       yin2=rs("金币")
       ti=rs("内力")
       dengji2=rs("等级")
       zhan2=rs("体力")
       shenfen2=trim(rs("身份")) 
      
     
   if topai="无" or topai="" then
       Response.Write "<script language=JavaScript>{alert('对方好像无门无派！');}</script>"
       rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
       Response.End
       end if
 if topai="官府" then
       Response.Write "<script language=JavaScript>{alert('对方可是官府，你是不是不想混了！');}</script>"
       rs.close
	set rs=nothing	
	conn.close
	set conn=nothing	
       Response.End
	end if
       if InStr("|" & shidi1 & "|", "|" & topai& "|")=0 and mp<>topai then
       Response.Write "<script language=JavaScript>{alert('他们门派好像没有和你的门派开战呀，如果想开战向掌门请求！');}</script>"
       rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
       Response.End
       end if
 
       
	 if shenfen2="掌门" then
          zhishu2=2
       elseif shenfen2="副掌门" then
         zhishu2=1.8
       elseif shenfen2="长老"  then
         zhishu2=1.6
       elseif shenfen2="护法"  then
          zhishu2=1.4
        elseif shenfen2="堂主"  then
          zhishu2=1.2
        else
           zhishu2=1
        end if
       rs.close
       sql="SELECT s,h,q,f FROM p WHERE  a='" & topai & "'"
       rs.open sql,conn,1,1
       if rs.eof or rs.bof then 
       Response.Write "<script language=JavaScript>{alert('对方好像不是门派中人！');}</script>"
       rs.close
	set rs=nothing	
	conn.close
	set conn=nothing	
       Response.End
	end if
        shengming=rs("s")
        kunjin=int(rs("h"))
        shidi2=rs("q")
         baohu2=rs("f")
          if baohu2=1 then
       Response.Write "<script language=JavaScript>{alert('他们门派现在接受官府保护中，你无法向他开战！');}</script>"
       rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
       Response.End
       end if
       if InStr("|" & shidi2 & "|", "|" & mp& "|")=0 and mp<>topai then
       Response.Write "<script language=JavaScript>{alert('他们门派好像没有和你的门派开战呀，如果想开战向掌门请求！');}</script>"
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
       Response.Write "<script language=JavaScript>{alert('你这招也太弱了吧！');}</script>"
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
          e="[<a href=javascript:parent.sw('["&sjjh_name&"]'); target=f2>"&sjjh_name & "</a>]<bgsound src=READONLY/fazhao.mid loop=1>对" & to1 & "使用了一招" & mp & "的("& fn1 &")，恰好"&to1&"往旁边一晃，躲过了这一劫，好险哦！"
          baofa=0
          case"2"
          e="[<a href=javascript:parent.sw('["&sjjh_name&"]'); target=f2>"&sjjh_name & "</a>]<bgsound src=READONLY/fazhao.mid loop=1>乘着" & to1 & "不备，偷袭成功，一招" & mp & "的("& fn1 &")，重重的打在"&to1&"的背上。"
          baofa=2
          case"8"
          e="[<a href=javascript:parent.sw('["&sjjh_name&"]'); target=f2>"&sjjh_name & "</a>]<bgsound src=READONLY/fazhao.mid loop=1>突发神力，大喝一声，一招" & mp & "的("& fn1 &")打的"&to1&"狂吐鲜血。"
          baofa=3
          case"16"
          e="[<a href=javascript:parent.sw('["&sjjh_name&"]'); target=f2>"&sjjh_name & "</a>]<bgsound src=READONLY/fazhao.mid loop=1>偷袭"&to1&"，不料"&to1&"早有防备，反被"&to1&"的护体神功震伤，内力下降-5000。"
          baofa=0
          conn.execute"update 用户 set ,内力=内力-5000 where 姓名='" & sjjh_name & "'"
          case else
          baofa=1
          e="[<a href=javascript:parent.sw('["&sjjh_name&"]'); target=f2>"&sjjh_name & "</a>]<bgsound src=READONLY/fazhao.mid loop=1>扎起马步对"&to1&"使用了一招" & mp & "的("& fn1 &")。"
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
	            	conn.execute"update p set s=s-"& kills3 & " where a='" & topai & "'"

	             	conn.execute"update 用户 set 体力=体力-"  & kills & " where 姓名='" & to1 & "'"
			         if shengming-kills3<=0 then
			              if mp=topai then
			              conn.execute"update 用户 set 金币=金币+10,道德=道德+1000,魅力=魅力+1000,内力=内力*2,武功=武功*2 where 姓名='" & sjjh_name & "'"
			              conn.execute"update 用户 set 门派='无',身份='无',内力=内力*1/2,武功=武功*1/2 where 门派='" & topai & "'"
                          conn.execute"delete * from p  where a='" & topai & "'"
                          call boot(to1,"门派大战，操作者："&sjjh_name&",["&mp&"]"&fn1)
                          conn.execute"insert into 人命(死者,时间,凶手,死因) values ('" & topai & "',now(),'" &sjjh_name & "','被门派内奸出卖')"
                         

			              menattack=e&"杀伤"&to1&"战斗力为<b><font color=red>"&kills&"</font></b>,杀伤"&topai&"门派战斗力为<b><font color=red>"&kills3&"</font></b>,并最终使该门派覆灭,"& sjjh_name &"作为卧底冲锋陷阵,道德上升1000,魅力上升1000,金币增加<b><font color=red>10</font></b>!"
                          else 
			            shidis=replace(shidi1,topai&"|","")
			            conn.execute"update 用户 set 金币=金币+10,道德=道德+1000,魅力=魅力+1000 where 姓名='" & sjjh_name & "'"
                    	conn.execute"update 用户 set 门派='无',身份='无',内力=内力*1/2,武功=武功*1/2 where 门派='" & topai & "'"
						conn.execute"delete * from p  where a='" & topai & "'"
						conn.execute"update 门派  set h=h+'"&kunjin&"',q='"&shidis&"' where a='" & mp & "'"
						call boot(to1,"门派大战，操作者："&sjjh_name&",["&mp&"]"&fn1)
                        conn.execute"insert into 人命(死者,时间,凶手,死因) values ('" & topai & "',now(),'" & mp & "','被仇敌围剿覆灭')"
                        menattack=e&"杀伤"&to1&"战斗力为<b><font color=red>"&kills&"</font></b>,杀伤对方"&topai&"门派战斗力为<b><font color=red>"&kills3&"</font></b>,并最终使该门派覆灭,"& sjjh_name &"为门派冲锋陷阵,道德上升1000,魅力上升1000,金币增加<b><font color=red>10</font></b>!"
						 end if
                             rs.close
                             set rs=nothing	
	                      conn.close
	                      set conn=nothing	
                             exit function
                    end if
                        conn.execute"update 用户 set 体力=体力-" & nei & ",道德=道德+2 where 姓名='" & sjjh_name & "'"   
                        menattack=e&"杀伤"&to1&"战斗力为<b><font color=red>"&kills&"</font></b>,杀伤对方"&topai&"门派战斗力为<b><font color=red>"&kills3&"</font></b>,"& sjjh_name &"为门派冲锋陷阵,道德上升2!"
                       if zhan2-kills<1000 then
                       menattack=menattack&""&to1&"战斗力降为1000以下,已经丧失威胁性,不需要攻击了."
                       end if
                        rs.close
                        set rs=nothing	
	                 conn.close
                       set conn=nothing                              
          
end function
%>


