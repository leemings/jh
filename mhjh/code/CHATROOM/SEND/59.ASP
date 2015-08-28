<%
function xuanyuan(un,st,mg)
timetmp=now()
f=Minute(time())
if f<30 and chatroomsn<>2 then
	     Response.Write "<script language=JavaScript>{alert('这个房间打仗时间是每个小时的后30分钟!要解决个人恩怨请到论剑大厅，那里是杀人的战场！！');}</script>"
         Response.End
end if
Set connb = Server.CreateObject("ADODB.Connection")
connb.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("../21ex/mj.asp")
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
rst.open "select 保护,门派,积分,攻击,防御,状态,杀人,精力,protect from 用户 where 姓名='"&un&"'",conn
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
	     Response.Write "<script language=JavaScript>{alert('你现在处于保护阶段，无法攻击！');}</script>"
         Response.End
     
     end if
     if mp="出家" then
    	Response.Write "<script language=JavaScript>{alert('你是出家人，无法攻击！');}</script>"
        Response.End
	end if     
     if g<1000000 then
    	Response.Write "<script language=JavaScript>{alert('你自己连100万攻击都没有，还想打别人？去打工大厅新手练功那练练再说吧！');}</script>"
        Response.End
        end if
     if f=<1000000 then
    	Response.Write "<script language=JavaScript>{alert('你自己连100万防御都没有，还想惹事？去打工大厅新手练功那练练再说吧！');}</script>"
        Response.End
        end if
	if jf<30000 then
    	Response.Write "<script language=JavaScript>{alert('你手无缚鸡之力，还是先别打人了');}</script>"
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
	     Response.Write "<script language=JavaScript>{alert('对方正在修练中，你不能进行神灵攻击！');}</script>"
         Response.End
       end if
     if mp2="官府" and chatroomsn<>2 then
       Response.Write "<script language=JavaScript>{alert('对方可是官府，你是不是不想在江湖里混了！');}</script>"
       Response.End
      end if
     if g1=<1000000 then
    	Response.Write "<script language=JavaScript>{alert('对方连100万攻击都没有，你就饶了他吧！');}</script>"
        Response.End
        end if
     if f1=<1000000 then
    	Response.Write "<script language=JavaScript>{alert('对方连100万防御都没有，你就放过他吧！');}</script>"
        Response.End
        end if
       if mp2="出家" then
    	Response.Write "<script language=JavaScript>{alert('对方是出家人，你还不肯放过人家？');}</script>"
        Response.End
	end if
	if jf2<30000 then
    	Response.Write "<script language=JavaScript>{alert('对方还是个孩子，你怎么能忍心下手呢？');}</script>"
        Response.End
	end if
	      randomize timer
          txt=int(rnd*9)+1
	       if gong1>Application("yx8_mhjh_Maxattack") then gong1=Application("yx8_mhjh_Maxattack")-txt
	       if gong1>tl2 then gong1=tl2

if zt2<>"死亡" then
      if zztt=True then
       conn.execute "update 用户 set 体力=体力-"&gong1*0.05&",攻击=攻击-"&gong1*0.005&" where 姓名='"&st&"'"
       connb.execute "update mj set 攻击=攻击*0.9 where 所有者='"&un&"' and 名称='"&mg&"' and 类型='轩'" 
       xuanyuan="<font color=FF0000>【##攻击】</font>##使出轩辕武学之"&mg&",[img]"&tu&"[/img],杀伤%%体力"&gong1*0.05&"、攻击"&gong1*0.005&"，##的轩辕武功攻击力还剩下"&gong1&"！<bgsound src=../mid/DW.wav loop=1>"
      else
      if nll1>nll2 then
        conn.execute "update 用户 set 体力=体力-"&fangyu1*0.08&",防御=防御-"&fangyu1*0.005&" where 姓名='"&st&"'"
        connb.execute "update mj set 防御=防御*0.9 where 所有者='"&un&"' and 名称='"&mg&"' and 类型='轩'" 
        xuanyuan="<font color=FF0000>【##防守反击】</font>##使出轩辕武学之"&mg&",以柔制刚,[img]"&tu&"[/img],使%%体力减少"&fangyu1*0.08&",内力反震回去,%%防御减少"&fangyu1*0.005&"，##的轩辕武功攻击力还剩下"&fangyu1&"<bgsound src=../mid/DW.wav loop=1>"
      else
        conn.execute "update 用户 set 体力=体力-"&fangyu1*0.05&",防御=防御-"&fangyu1*0.002&" where 姓名='"&st&"'"
        connb.execute "update mj set 防御=防御*0.9 where 所有者='"&un&"' and 名称='"&mg&"' and 类型='轩'"  
        xuanyuan="<font color=FF0000>【##防守反击】</font>##使出轩辕武学之"&mg&",扭转乾坤,[img]"&tu&"[/img],内力发挥最大威力,因此杀伤%%体力"&fangyu1*0.05&",使%%防御减少"&fangyu1*0.002&"，##轩辕武功攻击力剩下"&fangyu1&"<bgsound src=../mid/DW.wav loop=1>"
      end if
      end if
 if tl2<3000 then
	 if sharen>=10 then
	   conn.execute "update 用户 set 杀人者='"&un&"' where 姓名='"&st&"'"
       conn.execute "update 用户 set 道德=道德-10000,资质=资质+10000,内力=内力-50000,杀人=杀人+1,总杀人=总杀人+1,攻击=攻击-10000 where 姓名='"&un&"'"
       conn.execute "insert into 英烈堂(时间,死者,凶手,死因) values(now(),'"&st&"','"&un&"','"&mg&"')"
     else
       conn.execute "update 用户 set 道德=道德-10000,资质=资质+10000,内力=内力-5000000,杀人=杀人+1,总杀人=总杀人+1,攻击=攻击-10000 where 姓名='"&un&"'"
       conn.execute "insert into 英烈堂(时间,死者,凶手,死因) values(now(),'"&st&"','"&un&"','"&mg&"')"
    end if
       especial="死亡"
	   especialtime=120
      xuanyuan="<font color=FF0000>【##攻击】</font>##使出轩辕武学之"&mg&",[img]"&tu&"[/img],%%学艺不精，当场被打死！##为杀这个人自己损失了100000点道德和500万内力、攻击丢失5万，获得资质1万。"&newaberrant(st,un,especial,especialtime)
 end if
else
  conn.execute "update 用户 set 道德=道德-10000,内力=内力-500000,攻击=攻击-10000 where 姓名='"&un&"'"
  xuanyuan="<font color=FF0000>【##攻击】</font>##打死了%%，竞然开始鞭尸，自身内力减少500000、攻击损伤10000,道德减10000！"
end if
         set rs=nothing
	     connb.close
	     set connb=nothing
end function
%>
