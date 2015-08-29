<%
function xw(un,st,mg)
timetmp=now()
f=Minute(time())
if f>30 and chatroomsn<>1 then
	     Response.Write "<script language=JavaScript>{alert('随师傅学习武功需要在每个小时的前30分钟，而且必须在快乐江湖中进行！');}</script>"
         Response.End
end if
Set connb = Server.CreateObject("ADODB.Connection")
connb.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("../21ex/xw.asp")
Set rs = Server.CreateObject("ADODB.Recordset")
sql="select 攻击,防御,速度,时间,图,状态 from mj where 名称='"&mg&"' and 所有者='"&un&"' and 类型='轩'"
Set Rs=connb.Execute(sql)  
      if rs.bof or rs.eof then
        Response.Write "<script language=JavaScript>{alert('您没有["&mg&"]这招!');}</script>"
        Response.End
      end if
gong1=rs("攻击")
fangyu1=rs("防御")
sudu1=rs("速度")
tu=rs("图")
zztt=rs("状态")
rs.close
rst.open "select 保护,门派,积分,体力,师傅,攻击,防御,状态,杀人,精力,protect from 用户 where 姓名='"&un&"'",conn
       sf=rst("师傅")
       tili=rst("体力")
       mp=rst("门派")
       baohu1=rst("保护")
       protect=rst("protect")
       jf=rst("积分")
       g=rst("攻击")
       f=rst("防御")
       zt=rst("状态")
       nll1=rst("精力")
       sharen=rst("杀人")
       rst.close
      if baohu1=True or protect>timetmp then
	     Response.Write "<script language=JavaScript>{alert('你现在处于保护阶段，无法向师傅学习武功！');}</script>"
         Response.End
     
     end if
     if mp="出家" then
    	Response.Write "<script language=JavaScript>{alert('你是出家人，无法向世俗的高手学习杀人的功夫！');}</script>"
        Response.End
	end if     
     if g>100000000 then
    	Response.Write "<script language=JavaScript>{alert('你的攻击已经足够高强了，无法再向师傅学习武功！');}</script>"
        Response.End
        end if
     if f>100000000 then
    	Response.Write "<script language=JavaScript>{alert('你防御都这么高了，还想认师傅啊？你可以自己去修炼了！');}</script>"
        Response.End
        end if
     if tili<100000 then
    	Response.Write "<script language=JavaScript>{alert('师傅提示：你体力已经小于50万了，再练就把自己给练死了，赶紧去药店买点药品补充一下体力吧！');}</script>"
        Response.End
        end if
     if sf<>st then
    	Response.Write "<script language=JavaScript>{alert('他不是你师傅，你别乱学，会走火入魔的！！');}</script>"
        Response.End
        end if
	if jf>5000000 then
    	Response.Write "<script language=JavaScript>{alert('你等级这么高，不能再向师傅学习武功了，你可以自己闯荡江湖，创造一番属于自己的霸业了!');}</script>"
        Response.End
	end if
	rst.open "select 保护,门派,攻击,防御,积分,状态,protect,体力,精力 from 用户 where 姓名='"&st&"'",conn
if rst.EOF or rst.BOF then
Response.Write "<script language=JavaScript>{alert('是不是想称霸江湖想疯了，先把自己练练再说，轩辕武功是用来杀伤敌人的攻击或者防御的，别拿这个瞎打，这是浪费！');}</script>"
Response.End
      end if
       mp2=rst("门派")
       baohu2=rst("保护")
       protect2=rst("protect")
       jf2=rst("积分")
       g1=rst("攻击")
       f1=rst("防御")
       zt2=rst("状态")
       tl2=rst("体力")
       nll2=rst("精力")
       rst.close
      if baohu2=True or protect2>timetmp then
	     Response.Write "<script language=JavaScript>{alert('你师傅是个胆小鬼，嘿嘿，他这么菜你还向他学习？去告诉他把保护解开！');}</script>"
         Response.End
       end if
     if mp2="官府" and chatroomsn<>2 then
       Response.Write "<script language=JavaScript>{alert('对方可是官府，他没功夫教你练功！');}</script>"
       Response.End
      end if
     if g1=<1000000 then
    	Response.Write "<script language=JavaScript>{alert('对方连100万攻击都没有，你向他学什么呢？！');}</script>"
        Response.End
        end if
     if f1=<1000000 then
    	Response.Write "<script language=JavaScript>{alert('对方连100万防御都没有，你还是另寻高明吧！');}</script>"
        Response.End
        end if
       if mp2="出家" then
    	Response.Write "<script language=JavaScript>{alert('你师傅出家了，不想再指导世俗的弟子!');}</script>"
        Response.End
	end if
	if jf2<30000 then
    	Response.Write "<script language=JavaScript>{alert('对方连10万积分都没有，这么菜，你怎么认他做师傅，能学到什么好功夫？？');}</script>"
        Response.End
	end if
	      randomize timer
          txt=int(rnd*9)+1
	       if gong1>Application("yx8_mhjh_Maxattack") then gong1=Application("yx8_mhjh_Maxattack")-txt
	       if gong1>tl2 then gong1=tl2

if zt2<>"死亡" then
      if zztt=True then
       conn.execute "update 用户 set 体力=体力-"&gong1*0.002&",攻击=攻击+"&gong1*0.002&",积分=积分+5 where 姓名='"&un&"'"
       connb.execute "update mj set 攻击=攻击*0.95 where 所有者='"&un&"' and 名称='"&mg&"' and 类型='轩'" 
       xw="<font color=FF0000>【##修炼】[img]"&tu&"[/img],##运用自己从师傅那新学的武功"&mg&"，向师傅%%学习高深武学，只见其双掌飞舞，四周气浪翻滚，##的积分提高5点，攻击增加"&gong1*0.002&"。##的招试攻击剩下"&gong1&"！<bgsound src=../mid/Dianxue.wav loop=1>"
      else
      if nll1>nll2 then
        conn.execute "update 用户 set 体力=体力-"&fangyu1*2&",防御=防御-20000 where 姓名='"&st&"'"
        connb.execute "update mj set 防御=防御*0.95 where 所有者='"&un&"' and 名称='"&mg&"' and 类型='轩'" 
        xw="<font color=FF0000>【##入魔】</font>[img]"&tu&"[/img],##练功急于求成，精力爆涨，不小心走火入魔，杀伤%%体力"&fangyu1*3&"，%%防御降低20000，##的招试攻击降低到"&fangyu1&"。<bgsound src=../mid/DW.wav loop=1>"
      else
        conn.execute "update 用户 set 体力=体力-"&fangyu1*0.002&",攻击=攻击+"&gong1*0.002&",积分=积分+5 where 姓名='"&un&"'"
        connb.execute "update mj set 防御=防御*0.95 where 所有者='"&un&"' and 名称='"&mg&"' and 类型='轩'"  
        xw="<font color=FF0000>【##修炼】[img]"&tu&"[/img],##运用自己从师傅那新学的武功"&mg&"，向师傅%%学习高深武学，只见其双掌飞舞，四周气浪翻滚，##的积分提高5点，攻击增加"&fangyu1*0.002&"。##的招试攻击剩下"&fangyu1&"！<bgsound src=../mid/Dianxue.wav loop=1>"
      end if
      end if
 if tl2<3000 then
	 if sharen>=10 then
	   conn.execute "update 用户 set 杀人者='"&un&"' where 姓名='"&st&"'"
       conn.execute "update 用户 set 道德=道德-10000,资质=资质+100000,内力=内力-500000,杀人=杀人+1,总杀人=总杀人+1,攻击=攻击-10000 where 姓名='"&un&"'"
       conn.execute "insert into 英烈堂(时间,死者,凶手,死因) values(now(),'"&st&"','"&un&"','"&mg&"')"
     else
       conn.execute "update 用户 set 道德=道德-10000,资质=资质+10000,内力=内力-500000,杀人=杀人+1,总杀人=总杀人+1,攻击=攻击-10000 where 姓名='"&un&"'"
       conn.execute "insert into 英烈堂(时间,死者,凶手,死因) values(now(),'"&st&"','"&un&"','"&mg&"')"
    end if
       especial="死亡"
	   especialtime=120
      xw="<font color=FF0000>【##入魔】</font>[img]"&tu&"[/img],##练功急于求成，精力爆涨，不小心走火入魔，耗费了50万内力，竟然把%%一下子给打死了，虽然道德减少1万，却把师傅的资质给夺来10万点。"&newaberrant(st,un,especial,especialtime)
 end if
else
  conn.execute "update 用户 set 道德=道德-10000,内力=内力-500000,攻击=攻击-10000 where 姓名='"&un&"'"
  xw="<font color=FF0000>【##入魔】</font>##在练功时不小心打死了师傅%%，竞然开始鞭尸，自身内力减少500000、攻击损伤10000,道德减10000！！"
end if
         set rs=nothing
	     connb.close
	     set connb=nothing
end function
%>
