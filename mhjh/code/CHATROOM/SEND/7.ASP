<%
function attack(un,st,mg,gr)
timetmp=now()
f=Minute(time())
rst.Open "select ta.消耗内力,ta.基本攻击+tu.攻击 as 攻击,ta.特效,ta.攻击说明,tu.内力,tu.门派,tu.积分,tu.形态,tu.道德,tu.保护,tu.体力,tu.会员,tu.任务名称,tu.protect from 攻击 ta inner join 用户 tu ON tu.姓名=ta.姓名 where ta.姓名='"&un&"' and ta.招式='"&mg&"'",conn
if rst.EOF or rst.BOF then
attack="<font color=FF0000>【攻击】</font>##咬牙切齿地对%%说：'等我学会了"&mg&",一定会再来找你一比高下！"
elseif f<30 and chatroomsn<>2 then
attack="<font color=FF0000>【攻击】</font>##用"&mg&"攻击%%失败，原因：这个房间打仗时间是每个小时的后30分钟!要解决个人恩怨请到论剑大厅，那里是杀人的战场！"
else
basemp=rst("消耗内力")
jf=rst("积分")
att=rst("攻击")
especial=rst("特效")
fmorcal=rst("道德")
huiyuan=rst("会员")
zs=rst("任务名称")
atdeclaration=rst("攻击说明")
if isnull(especial) then especila="无"
if rst("内力")<basemp then
attack="<font color=FF0000>【攻击】</font>##本意是想用"&mg&"攻击%%，可惜心有余而力不足。"
elseif rst("体力")<200000 then
attack="<font color=FF0000>【攻击】</font>##你手无缚鸡之力，没200000体力还是先别打人了！先给自己吃点药吧"
elseif rst("形态")="鬼魂" then
attack="<font color=FF0000>【攻击】</font>##你现在是个鬼魂，无法打人，赶紧去凤凰蘖盘那进化成人吧"
elseif rst("保护")=True then
attack="<font color=FF0000>【攻击】</font>##你现在处于保护阶段，无法攻击%%！"
elseif rst("门派")="出家" then
attack="<font color=FF0000>【攻击】</font>##你已经是个出家人了，##你无法使用"&mg&"攻击！"
else
rst.Close
rst.Open "select 体力,内力,防御,道德,特技,积分,状态,形态,门派,保护,等级,protect,杀人 from 用户 where 姓名='"&st&"'",conn
if rst("门派")="官府" and chatroomsn<>2 then
attack="<font color=FF0000>【攻击】</font>##你不能攻击官府的老大!"	
elseif rst("积分")<5000 then
attack="<font color=FF0000>【攻击】</font>%%连5000积分都没有,还是个孩子，##你怎么能忍心下手呢？"
elseif st="爱情鸟" and jf>20000 then
attack="<font color=FF0000>【攻击】</font>%%是帮助新人尽快提高状态用的,##你都这么厉害了还想来凑热闹？"
elseif rst("保护")=True and zs="无" then
attack="<font color=FF0000>【攻击】</font>%%现在处于受保护阶段，##你无法使用"&mg&"攻击！"
elseif rst("门派")="出家" then
attack="<font color=FF0000>【攻击】</font>%%已经是个出家人了，##你无法使用"&mg&"攻击！"
elseif rst("防御")<200000 then
attack="<font color=FF0000>【攻击】</font>%%不太会玩，连20万防御都没有，##你就放过他吧！让%%先练练再说。"
elseif rst("形态")="鬼魂" then
attack="<font color=FF0000>【攻击】</font>%%这是个鬼魂，死的够惨的了。##你无法使用"&mg&"攻击！"
else
resatt=att-rst("防御")
stespecial=rst("特技")
condition=rst("状态")
tmorcal=rst("道德")
sharen=rst("杀人")
qq=gr*2000
Maxattack=Application("yx8_mhjh_Maxattack")*gr\50
randomize()
if resatt<0 then
resatt=clng(rnd()*100)+1
elseif resatt>Maxattack then
resatt=Maxattack-clng(rnd()*100)
end if
morcal=clng(fmorcal-tmorcal*0.01)
if morcal>=2100000000 then
morcal=2100000000
elseif morcal<=-2100000000 then
morcal=-2100000000
end if
if resatt>rst("体力") and condition<>"死亡" then
especial="死亡"
especialtime=25
randomize()
rnd1=cint(rnd()*100)
jstl=1+rnd1
jstl1=1+rnd1
jstl2=1+rnd1
jstl3=1+rnd1
jstl4=1+rnd1
jstl5=1+rnd1
		Application.Lock
		Application("yx8_mhjh_klt")=jstl
		Application("yx8_mhjh_klt1")=jstl1
		Application("yx8_mhjh_klt2")=jstl2
		Application("yx8_mhjh_klt3")=jstl3
		Application("yx8_mhjh_klt4")=jstl4
		Application("yx8_mhjh_klt5")=jstl5
		Application.UnLock	
   if st="武术家" then
     conn.execute "update 用户 set 体力=100000000,状态='正常' where 姓名='武术家'"
     conn.execute "update 用户 set 内力=内力-"&basemp*5&",杀人=杀人+1,总杀人=总杀人+1,攻击=攻击+1000,防御=防御+2000,积分=积分+200 where 姓名='"&un&"'"
     conn.execute "insert into 英烈堂(时间,死者,凶手,死因) values(now(),'"&st&"','"&un&"','"&mg&"')"
     attack="<font color=FF0000>【##攻击】</font><font color=FF0000>"&atdeclaration&"，成功的打败了%%！获得积分200点,攻击1000点,防御增加2000点！%%的很多东西散落到地上,手快的人快抢啊！</font><marquee width=100% behavior=alternate scrollamount=15><input type=button value='红包裹' onclick=""javascript:window.open('zlp.asp?jstl="&Application("yx8_mhjh_klt")&"','newwindow','width=380,height=210')"" style=""background-color:#FF0000;color:FFFFFF;border: 1 double"" onMouseOver=""this.style.color='FFFF00'"" onMouseOut=""this.style.color='FFFFFF'"" name=""button2231""><input type=button value='绿包裹' onclick=""javascript:window.open('zlp2.asp?jstl1="&Application("yx8_mhjh_klt1")&"','newwindow','width=380,height=210')"" style=""background-color:#00FF00;color:FFFFFF;border: 1 double"" onMouseOver=""this.style.color='FFFF00'"" onMouseOut=""this.style.color='FFFFFF'"" name=""button2231""><input type=button value='黄包裹' onclick=""javascript:window.open('zlp1.asp?jstl2="&Application("yx8_mhjh_klt2")&"','newwindow','width=380,height=210')"" style=""background-color:#FF9900;color:FFFFFF;border: 1 double"" onMouseOver=""this.style.color='FFFF00'"" onMouseOut=""this.style.color='FFFFFF'"" name=""button2231""><input type=button value='蓝包裹' onclick=""javascript:window.open('zlp3.asp?jstl3="&Application("yx8_mhjh_klt3")&"','newwindow','width=380,height=210')"" style=""background-color:#0000FF;color:FFFFFF;border: 1 double"" onMouseOver=""this.style.color='FFFF00'"" onMouseOut=""this.style.color='FFFFFF'"" name=""button223""><input type=button value='黑包裹' onclick=""javascript:window.open('zlp4.asp?jstl4="&Application("yx8_mhjh_klt4")&"','newwindow','width=380,height=210')"" style=""background-color:#000000;color:FFFFFF;border: 1 double"" onMouseOver=""this.style.color='FFFF00'"" onMouseOut=""this.style.color='FFFFFF'"" name=""button223""></marquee><br><font size=3><b><font color=FF0000>【江湖通报】</font>倒在地上的%%缓缓的从怀中摸出一颗火红的药丸，塞进自己的嘴里，吞了下去，突然间两眼放出红光，从地上跳了起来，瞬间恢复了正常,喊到：“我没问题,再来打！”</b></font>"
elseif st="魔法师" then
     conn.execute "update 用户 set 体力=50000000,状态='正常' where 姓名='魔法师'"
     conn.execute "update 用户 set 内力=内力-"&basemp*5&",杀人=杀人+1,总杀人=总杀人+1,攻击=攻击+1000,防御=防御+2000,积分=积分+200 where 姓名='"&un&"'"
     conn.execute "insert into 英烈堂(时间,死者,凶手,死因) values(now(),'"&st&"','"&un&"','"&mg&"')"
     attack="<font color=FF0000>【##攻击】</font><font color=FF0000>"&atdeclaration&"，成功的打败了%%！获得积分200点,攻击1000点,防御增加2000点！%%的很多东西散落到地上,手快的人快抢啊！</font><marquee width=100% behavior=alternate scrollamount=15><input type=button value='红包裹' onclick=""javascript:window.open('zlp.asp?jstl="&Application("yx8_mhjh_klt")&"','newwindow','width=380,height=210')"" style=""background-color:#FF0000;color:FFFFFF;border: 1 double"" onMouseOver=""this.style.color='FFFF00'"" onMouseOut=""this.style.color='FFFFFF'"" name=""button2231""><input type=button value='绿包裹' onclick=""javascript:window.open('zlp2.asp?jstl1="&Application("yx8_mhjh_klt1")&"','newwindow','width=380,height=210')"" style=""background-color:#00FF00;color:FFFFFF;border: 1 double"" onMouseOver=""this.style.color='FFFF00'"" onMouseOut=""this.style.color='FFFFFF'"" name=""button2231""><input type=button value='黄包裹' onclick=""javascript:window.open('zlp1.asp?jstl2="&Application("yx8_mhjh_klt2")&"','newwindow','width=380,height=210')"" style=""background-color:#FF9900;color:FFFFFF;border: 1 double"" onMouseOver=""this.style.color='FFFF00'"" onMouseOut=""this.style.color='FFFFFF'"" name=""button2231""><input type=button value='蓝包裹' onclick=""javascript:window.open('zlp3.asp?jstl3="&Application("yx8_mhjh_klt3")&"','newwindow','width=380,height=210')"" style=""background-color:#0000FF;color:FFFFFF;border: 1 double"" onMouseOver=""this.style.color='FFFF00'"" onMouseOut=""this.style.color='FFFFFF'"" name=""button223""><input type=button value='黑包裹' onclick=""javascript:window.open('zlp4.asp?jstl4="&Application("yx8_mhjh_klt4")&"','newwindow','width=380,height=210')"" style=""background-color:#000000;color:FFFFFF;border: 1 double"" onMouseOver=""this.style.color='FFFFFF'"" onMouseOut=""this.style.color='FFFFFF'"" name=""button223""><input type=button value='白包裹' onclick=""javascript:window.open('zlp5.asp?jstl5="&Application("yx8_mhjh_klt5")&"','newwindow','width=380,height=210')"" style=""background-color:#C0C0C0;color:FFFFFF;border: 1 double"" onMouseOver=""this.style.color='FFFF00'"" onMouseOut=""this.style.color='FFFFFF'"" name=""button223""></marquee><br><font size=2><b><font color=FF0000>【江湖通报】</font><img src=IMAGE/IMAGE/004.gif width='45' height='45'>突然一个小天使翩然而至,一道霞光笼罩在%%头顶,%%两眼放出红光，从地上跳了起来,喊到：“我是不死之神,咱们再来打过！”</b></font>"
elseif st="爱情鸟" then
     conn.execute "update 用户 set 体力=10000000,状态='正常' where 姓名='爱情鸟'"
     conn.execute "update 用户 set 内力=内力-"&basemp*5&",杀人=杀人+1,总杀人=总杀人+1,攻击=攻击+1000,防御=防御+2000,积分=积分+200 where 姓名='"&un&"'"
     conn.execute "insert into 英烈堂(时间,死者,凶手,死因) values(now(),'"&st&"','"&un&"','"&mg&"')"
     attack="<font color=FF0000>【##攻击】</font><font color=FF0000>"&atdeclaration&"，成功的打败了%%！获得积分200点,攻击1000点,防御增加2000点！%%的很多东西散落到地上,手快的人快抢啊！</font><marquee width=100% behavior=alternate scrollamount=15><input type=button value='红包裹' onclick=""javascript:window.open('zlp.asp?jstl="&Application("yx8_mhjh_klt")&"','newwindow','width=380,height=210')"" style=""background-color:#FF0000;color:FFFFFF;border: 1 double"" onMouseOver=""this.style.color='FFFF00'"" onMouseOut=""this.style.color='FFFFFF'"" name=""button2231""><input type=button value='绿包裹' onclick=""javascript:window.open('zlp2.asp?jstl1="&Application("yx8_mhjh_klt1")&"','newwindow','width=380,height=210')"" style=""background-color:#00FF00;color:FFFFFF;border: 1 double"" onMouseOver=""this.style.color='FFFF00'"" onMouseOut=""this.style.color='FFFFFF'"" name=""button2231""><input type=button value='黄包裹' onclick=""javascript:window.open('zlp1.asp?jstl2="&Application("yx8_mhjh_klt2")&"','newwindow','width=380,height=210')"" style=""background-color:#FF9900;color:FFFFFF;border: 1 double"" onMouseOver=""this.style.color='FFFF00'"" onMouseOut=""this.style.color='FFFFFF'"" name=""button2231""><input type=button value='蓝包裹' onclick=""javascript:window.open('zlp3.asp?jstl3="&Application("yx8_mhjh_klt3")&"','newwindow','width=380,height=210')"" style=""background-color:#0000FF;color:FFFFFF;border: 1 double"" onMouseOver=""this.style.color='FFFF00'"" onMouseOut=""this.style.color='FFFFFF'"" name=""button223""><input type=button value='黑包裹' onclick=""javascript:window.open('zlp4.asp?jstl4="&Application("yx8_mhjh_klt4")&"','newwindow','width=380,height=210')"" style=""background-color:#000000;color:FFFFFF;border: 1 double"" onMouseOver=""this.style.color='FFFFFF'"" onMouseOut=""this.style.color='FFFFFF'"" name=""button223""><input type=button value='白包裹' onclick=""javascript:window.open('zlp5.asp?jstl5="&Application("yx8_mhjh_klt5")&"','newwindow','width=380,height=210')"" style=""background-color:#C0C0C0;color:FFFFFF;border: 1 double"" onMouseOver=""this.style.color='FFFF00'"" onMouseOut=""this.style.color='FFFFFF'"" name=""button223""></marquee><br><font size=2><b><font color=FF0000>【江湖通报】</font><img src=IMAGE/IMAGE/004.gif width='45' height='45'>突然一阵悠扬的歌声隐隐约约传来,万丈光芒笼罩在%%头顶,%%缓缓恢复了生机,悠然说道：“我喜欢舍生取义,咱们再来打过！”</b></font>"
  else
     conn.execute "update 用户 set 道德="&morcal*0.5&",内力=内力-"&basemp*5&",杀人=杀人+1,总杀人=总杀人+1,攻击=攻击+100 where 姓名='"&un&"'"
     conn.execute "insert into 英烈堂(时间,死者,凶手,死因) values(now(),'"&st&"','"&un&"','"&mg&"')"
     attack="<font color=FF0000>【##攻击】</font>"&atdeclaration&"，竟把%%打死了！道德变成"&morcal*0.5&"，获得攻击100！"
  end if
elseif resatt>rst("体力") and condition="死亡" then
especial="死亡"
especialtime=25
conn.execute "update 用户 set 道德=道德-20000,内力=内力-"&basemp*5&",攻击=攻击+200 where 姓名='"&un&"'"
attack="<font color=FF0000>【##攻击】</font>##打死了%%，竞然开始鞭尸，自身内力减少"&basemp*5&",道德减2万，获得攻击200！！"
else
especialtime=resatt\100
if especialtime<10 then
especial="无"
elseif especialtime>18 then
especialtime=18
end if
ff=qq/2
if st="魔法师" then
conn.Execute "update 用户 set 体力=体力-"&qq*3&",内力=内力-"&basemp*5&",攻击=攻击+100,积分=积分+4 where 姓名='"&un&"'"
conn.Execute "update 用户 set 体力=体力-"&resatt-qq&",状态='正常' where 姓名='"&st&"'"
attack="<font color=FF0000>【##攻击】</font>"&atdeclaration&"，耗用内力"&basemp*5&"，攻击力提高100点，积分提高4点，使"&st&"生命减少"&resatt&"！<br><font color=FF0000>【魔法师反击】</font>"&st&"手中一道闪电袭向"&un&",使"&un&"生命减少"&qq*3&"！"
elseif st="武术家" then
conn.Execute "update 用户 set 体力=体力-"&qq&",内力=内力-"&basemp*5&",防御=防御+100,积分=积分+3 where 姓名='"&un&"'"
conn.Execute "update 用户 set 体力=体力-"&resatt-qq&",状态='正常' where 姓名='"&st&"'"
attack="<font color=FF0000>【##攻击】</font>"&atdeclaration&"，耗用内力"&basemp*5&"，防御力提高100点，积分提高2点，使%%生命减少"&resatt&"！<br><font color=FF0000>【武术家反击】</font>%%奇怪掌影笼罩在%%周围，掌力飘忽不定，遇强则强，遇若则若，而且攻守兼备，使##生命减少"&qq&"！"
elseif st="爱情鸟" then
conn.Execute "update 用户 set 体力=体力-"&qq&",内力=内力-"&basemp*5&",攻击=攻击+100,防御=防御+100,积分=积分+5 where 姓名='"&un&"'"
conn.Execute "update 用户 set 体力=体力-"&resatt-qq&",状态='正常' where 姓名='"&st&"'"
attack="<font color=FF0000>【##攻击】</font>"&atdeclaration&"，耗用内力"&basemp*5&"，攻击力提高100点，防御提高100点，积分提高5点，使%%生命减少"&resatt&"！<br><font color=FF0000>【爱情鸟反击】</font>"&st&"兰花指轻拂,道道厉气扫向##,使##生命减少"&qq&"！"
else
conn.Execute "update 用户 set 内力=内力-"&basemp*5&",攻击=攻击+50 where 姓名='"&un&"'"
conn.Execute "update 用户 set 体力=体力-"&resatt&",防御=防御+50,积分=积分+2 where 姓名='"&st&"'"
attack="<font color=FF0000>【##攻击】</font>"&atdeclaration&"，耗用内力"&basemp*5&"，攻击力提高50点，积分提高2点，使%%防御力提高50点,生命减少"&resatt&"！"
end if
end if
if instr(stespecial,";抗"&especial&";")<>0 then especial="无"
if especial<>"无" then	attack=attack&newaberrant(st,un,especial,especialtime)
end if
end if
end if
rst.Close
end function
%>
