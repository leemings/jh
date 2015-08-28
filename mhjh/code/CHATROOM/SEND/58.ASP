<%
function shenlong(un,st,mg)
timetmp=now()
f=Minute(time())
if f<30 and chatroomsn<>2 then
	     Response.Write "<script language=JavaScript>{alert('这个房间打仗时间是每个小时的后30分钟!要解决个人恩怨请到论剑大厅，那里是杀人的战场！！');}</script>"
         Response.End
end if
Set connb = Server.CreateObject("ADODB.Connection")
connb.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("../21ex/mj.asp")
Set rs = Server.CreateObject("ADODB.Recordset")
sql="select * from myanimal where username='"&un&"' and animalname='"&mg&"' and rest=0"
Set Rs=connb.Execute(sql)  
      if rs.bof or rs.eof then
        Response.Write "<script language=JavaScript>{alert('您没有["&mg&"]这种神灵!或您的神灵目前己被官官府囚禁了！');}</script>"
        Response.End
      end if
gong1=rs("attack")
sm=rs("sm")
lx=rs("lei")
rs.close
rst.open "select 保护,门派,积分,状态,形态,杀人,protect from 用户 where 姓名='"&un&"'",conn
       mp=rst("门派")
       baohu1=rst("保护")
       protect=rst("protect")
       jf=rst("积分")
       zt=rst("状态")
       xt=rst("形态")
       sharen=rst("杀人")
       rst.close
      if baohu1=True or protect>timetmp then
	     Response.Write "<script language=JavaScript>{alert('你现在处于保护阶段，无法攻击！');}</script>"
         Response.End
      end if
     if mp="官府" and chatroomsn<>2 then
       Response.Write "<script language=JavaScript>{alert('您是官府的，不能对江湖中人进行神灵攻击！');}</script>"
       Response.End
     end if
     if xt="鬼魂" then
       Response.Write "<script language=JavaScript>{alert('您是鬼魂哦，您的神灵只保护您作为人的状态，快找个法力高强的和尚帮你超度成人吧！');}</script>"
       Response.End
     end if
     if mp="出家"  then
    	Response.Write "<script language=JavaScript>{alert('你是出家人，无法攻击！');}</script>"
        Response.End
	end if
	if jf<50000 then
    	Response.Write "<script language=JavaScript>{alert('你手无缚鸡之力，还是先别打人了，至少要5万积分才行！');}</script>"
        Response.End
	end if
	rst.open "select 保护,门派,内力,形态,积分,状态,protect,内力 from 用户 where 姓名='"&st&"'",conn
if rst.EOF or rst.BOF then
Response.Write "<script language=JavaScript>{alert('是不是想称霸江湖想疯了，先把自己练练再说，神灵是用来保护自己的，不是打要塞的。别瞎打，这是找死！');}</script>"
Response.End
      end if
       mp2=rst("门派")
       xt2=rst("形态")
       baohu2=rst("保护")
       protect2=rst("protect")
       jf2=rst("积分")
       zt2=rst("状态")
       tl2=rst("内力")
       rst.close
      if baohu2=True or protect2>timetmp then
	     Response.Write "<script language=JavaScript>{alert('对方正在修练中，你不能进行神灵攻击！');}</script>"
         Response.End
      end if
     if mp2="官府" and chatroomsn<>2 then
       Response.Write "<script language=JavaScript>{alert('对方可是官府，你是不是不想在江湖里混了！');}</script>"
       Response.End
     end if
     if xt2="鬼魂" then
       Response.Write "<script language=JavaScript>{alert('对方是个鬼魂，你不能攻击它，否则它会时时纠缠你，那你就惨了，连梦都做不好！');}</script>"
       Response.End
     end if
     if mp2="出家"  then
    	Response.Write "<script language=JavaScript>{alert('对方是出家人，你还不肯放过人家？');}</script>"
        Response.End
	end if
	if jf2<50000 then
    	Response.Write "<script language=JavaScript>{alert('对方还是个孩子，你怎么能忍心用这种方法来攻击他呢？');}</script>"
        Response.End
	end if	
randomize timer
txt=int(rnd*9)+1
	       if gong1>Application("yx8_mhjh_Maxattack") then gong1=Application("yx8_mhjh_Maxattack")-txt
	       if gong1>tl2 then gong1=tl2
if zt2<>"死亡" then
       conn.execute "update 用户 set 内力=内力-"&gong1*3&" where 姓名='"&st&"'"
       connb.execute "update myanimal set attack=attack*0.9 where username='"&un&"' and animalname='"&mg&"'"
       shenlong="<font color=FF0000>【##攻击】</font>放出专门吸收别人内力的护身神灵"&mg&"("&lx&"),%%立刻被红光笼罩，被排泄内力"&gong1*3&",##的神灵攻击力消失十分之一！<bgsound src=../mid/Dianxue.wav loop=1>"
 if tl2<3000 then
	 if sharen>=10 then
	   conn.execute "update 用户 set 杀人者='"&un&"' where 姓名='"&st&"'"
       conn.execute "update 用户 set 道德=道德-100000,资质=资质+10,内力=内力+"&gong1*3&",杀人=杀人+1,总杀人=总杀人+1,攻击=攻击-50000 where 姓名='"&un&"'"
       conn.execute "insert into 英烈堂(时间,死者,凶手,死因) values(now(),'"&st&"','"&un&"','"&mg&"')"
     else
       conn.execute "update 用户 set 道德=道德-100000,资质=资质+10,内力=内力+"&gong1*3&",杀人=杀人+1,总杀人=总杀人+1,攻击=攻击-5000 where 姓名='"&un&"'"
       conn.execute "insert into 英烈堂(时间,死者,凶手,死因) values(now(),'"&st&"','"&un&"','"&mg&"')"
    end if
       especial="死亡"
	   especialtime=120
       connb.execute("update myanimal set outtime=True where username='"&st&"'")
      shenlong="<font color=FF0000>【##攻击】</font>##放出专门吸取别人内力的护身神灵"&mg&"("&lx&"),"&sm&"杀伤%%体力"&gong1*3&",%%学艺不精，当场被咬死！##为杀这个人自己损失了100000点道德和50000内力、攻击5000，##的神灵攻击力完全消失!"&newaberrant(st,un,especial,especialtime)
 end if
else
  conn.execute "update 用户 set 道德=道德-100000,内力=内力-100000,攻击=攻击-10000 where 姓名='"&un&"'"
  shenlong="<font color=FF0000>【##攻击】</font>##的神灵吸干了%%的内力，竞然开始吸%%的骨髓，造成自己内力减少100000、攻击损伤10000,道德减100000！"
end if
         set rs=nothing
	     connb.close
	     set connb=nothing
end function
%>