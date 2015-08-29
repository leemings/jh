<%
function gd(un,st,mg)
timetmp=now()
f=Minute(time())
if f<30 or chatroomsn<>1 then
	     Response.Write "<script language=JavaScript>{alert('攻打要塞的时间是每个小时的后30分钟!而且必须在[快乐江湖]中进行！');}</script>"
         Response.End
end if
Set connb = Server.CreateObject("ADODB.Connection")
connb.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("../21ex/gd.asp")
Set rs = Server.CreateObject("ADODB.Recordset")
sql="select 攻击,防御,速度,时间,图,状态 from mj where 名称='"&mg&"' and 所有者='"&un&"' and 类型='轩'"
Set Rs=connb.Execute(sql)  
      if rs.bof or rs.eof then
        Response.Write "<script language=JavaScript>{alert('您没有["&mg&"]这种要塞武器！');}</script>"
        Response.End
      end if
gong1=rs("攻击")
fangyu1=rs("防御")
sudu1=rs("速度")
tu=rs("图")
zztt=rs("状态")
rs.close
rst.open "select 保护,门派,积分,攻击,防御,状态,形态,杀人,精力,protect from 用户 where 姓名='"&un&"'",conn
       mp=rst("门派")
       baohu1=rst("保护")
       protect=rst("protect")
       jf=rst("积分")
       zt=rst("状态")
       g=rst("攻击")
       f=rst("防御")
       xt=rst("形态")
       nll1=rst("精力")
       sharen=rst("杀人")
       rst.close
      if baohu1=True or protect>timetmp then
	     Response.Write "<script language=JavaScript>{alert('你现在处于保护阶段，无法攻打敌人的要塞，先解除自己的保护再打！');}</script>"
         Response.End
     
     end if
     if mp="出家" then
    	Response.Write "<script language=JavaScript>{alert('出家人应该与世无争，你无法攻打要塞，眷恋红尘？那就先去还俗吧！');}</script>"
        Response.End
	end if     
     if xt="鬼魂" then
    	Response.Write "<script language=JavaScript>{alert('您是鬼魂哦，您现在是个鬼魂，无法参与门派战斗，快找个法力高强的和尚帮你超度成人吧！');}</script>"
        Response.End
	end if 
     if g<1000000 then
    	Response.Write "<script language=JavaScript>{alert('小孩子，你自己连100万攻击都没有，还想打要塞？去打工大厅新手练功那练练再说吧！');}</script>"
        Response.End
        end if
     if mp="无" then
    	Response.Write "<script language=JavaScript>{alert('小孩子，你没有门派，无法参与门派之间的纷争！');}</script>"
        Response.End
        end if
	if nll1<30000 then
    	Response.Write "<script language=JavaScript>{alert('攻打要塞是很耗费精力的，你连30000精力都没有，是先别打了，用命令里的精力多练练吧！');}</script>"
        Response.End
	end if
	rst.open "select 保护,门派,结盟,防御,状态,protect,体力 from 门派 where 门派='"&st&"'",conn
if rst.EOF or rst.BOF then
Response.Write "<script language=JavaScript>{alert('你缺心眼儿，哪有这个门派让你打,是不是想称霸江湖想疯了，先把自己练练再说，别瞎打，这是找死！');}</script>"
Response.End
      end if
       baohu2=rst("保护")
       protect2=rst("protect")
       g1=rst("结盟")
       f1=rst("防御")
       zt2=rst("状态")
       tl2=rst("体力")
       rst.close
      if baohu2=True or protect2>timetmp then
	     Response.Write "<script language=JavaScript>{alert('对方的要塞正处于魔幻之神的防御屏障保护之中，你的要塞冲击波无法进行攻击！先耐心等等机会吧！');}</script>"
         Response.End
       end if
     if mp2="官府" and chatroomsn<>1 then
       Response.Write "<script language=JavaScript>{alert('这可是官府的要塞，你是不是不想在江湖里混了！');}</script>"
       Response.End
      end if
if mp=g1 then
       Response.Write "<script language=JavaScript>{alert('你们两派已经结盟，不要起内讧为好，共同对外才能有机会称霸江湖。一意孤行的话，江湖中将到处是你的敌人！');}</script>"
       Response.End
      end if
     if f1=<0 then
    	Response.Write "<script language=JavaScript>{alert('对方连0万攻击都没有，你就饶了他吧！');}</script>"
        Response.End
        end if
     if f1=<0 then
    	Response.Write "<script language=JavaScript>{alert('对方连0万防御都没有，你就放过他吧！');}</script>"
        Response.End
        end if       	
	      randomize timer
          txt=int(rnd*9)+1
	       if gong1>Application("yx8_mhjh_Maxattack") then gong1=Application("yx8_mhjh_Maxattack")-txt
	       if gong1>tl2 then gong1=tl2

if zt2<>"死亡" and tl2>100000 then
      if zztt=True then
       conn.execute "update 门派 set 体力=体力-"&gong1*2&",防御=防御+1000 where 门派='"&st&"'"
       connb.execute "update mj set 攻击=攻击*0.9 where 所有者='"&un&"' and 名称='"&mg&"' and 类型='轩'" 
       gd="<font color=FF0000>【##攻打要塞】[img]"&tu&"[/img],##为实现一统江湖的梦想，发动了自己的要塞冲击波，汹涌的气流奔向%%的要塞，%%的要塞摇摇欲坠，杀伤%%要塞体力"&gong1*2&",##的要塞冲击波威力下降十分之一！<bgsound src=../mid/Dianxue.wav loop=1>"
      else
      if gong1+g>f1 then
        conn.execute "update 门派 set 体力=体力-"&fangyu1*1&",防御=防御-20000 where 门派='"&st&"'"
        connb.execute "update mj set 防御=防御*0.9 where 所有者='"&un&"' and 名称='"&mg&"' and 类型='轩'" 
        gd="<font color=FF0000>【##攻打要塞】</font>[img]"&tu&"[/img],##为实现一统江湖的梦想，发动了自己的要塞冲击波，汹涌的气流奔向%%的要塞，以柔制刚,使%%的要塞体力减少"&fangyu1*1&",内力反震回去,%%的要塞防御减少5000，##的要塞冲击波威力下降十分之一<bgsound src=../mid/DW.wav loop=1>"
      else
        conn.execute "update 门派 set 体力=体力-"&fangyu1*0.5&",防御=防御-20000 where 门派='"&st&"'"
        connb.execute "update mj set 防御=防御*0.9 where 所有者='"&un&"' and 名称='"&mg&"' and 类型='轩'"  
        gd="<font color=FF0000>【##攻打要塞】</font>[img]"&tu&"[/img],##为实现一统江湖的梦想，发动了自己的要塞冲击波，汹涌的气流奔向%%的要塞，想扭转乾坤,使%%的要塞体力降低"&fangyu1*0.5&",使%%防御减少5000，##的要塞冲击波威力下降十分之一<bgsound src=../mid/DW.wav loop=1>"
      end if
      end if
 if tl2<100000 then
	 if sharen>9 then
	   conn.execute "delete * from 门派 where 门派='"&st&"'"
       conn.execute "update 用户 set 道德=道德-100000,内力=内力-1000000,资质=资质+100000,银两=银两+20000000,杀人=杀人+1,总杀人=总杀人+1,攻击=攻击+50000 where 姓名='"&un&"'"
       conn.execute "insert into 英烈堂(时间,死者,凶手,死因) values(now(),'"&st&"','"&un&"','"&mg&"')"
     else
	   conn.execute "delete * from 门派 where 门派='"&st&"'"
       conn.execute "update 用户 set 道德=道德-100000,内力=内力-5000000,资质=资质+100000,银两=银两+50000000,杀人=杀人+1,总杀人=总杀人+1,防御=防御+100000 where 姓名='"&un&"'"
       conn.execute "insert into 英烈堂(时间,死者,凶手,死因) values(now(),'"&st&"','"&un&"','"&mg&"')"
    end if
       especial="死亡"
	   especialtime=120
      gd="<font color=FF0000>【##完美攻击】</font>[img]"&tu&"[/img],##孤注一掷，运用500万内力，摧动自己的要塞冲击波，万丈光芒笼罩住%%的要塞，%%的要塞顷刻间毁于一旦，化作一片废墟！##夺得了%%的金库中剩余银两50000000，把%%所有弟子的资质给搜刮来10万点，%%的防御中有10万点归##所有；为此不道德行为，##付出了10万道德。从此，江湖中又少了一个门派。##向称霸江湖迈进了一步！</font>"
 end if
else
  conn.execute "delete * from 门派 where 门派='"&st&"'"
       conn.execute "update 用户 set 道德=道德-100000,内力=内力-5000000,资质=资质+100000,银两=银两+50000000,杀人=杀人+1,总杀人=总杀人+1,防御=防御+100000 where 姓名='"&un&"'"
       conn.execute "insert into 英烈堂(时间,死者,凶手,死因) values(now(),'"&st&"','"&un&"','"&mg&"')"
  gd="<font color=FF0000>【##完美攻击】</font>[img]"&tu&"[/img],##孤注一掷，运用500万内力，摧动自己的要塞冲击波，万丈光芒笼罩住%%的要塞，%%的要塞顷刻间毁于一旦，化作一片废墟！##夺得了%%的金库中剩余银两50000000，把%%所有弟子的资质给搜刮来10万点，%%的防御中有10万点归##所有；为此不道德行为，##付出了10万道德。从此，江湖中又少了一个门派。##向称霸江湖迈进了一步！</font>"
end if
         set rs=nothing
	     connb.close
	     set connb=nothing
end function
%>
