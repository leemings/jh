<%
function huayuan(name1,sh1,sy1)
%> <!--#include file="data.asp"--> <%
sql="SELECT * FROM 用户 WHERE 姓名='" & name1 & "'"
Set Rs=conn.Execute(sql)
if rs.eof or rs.bof then
      huayuan=""
else
      if rs("宝物")=Application("aqjh_baowuname") then
            huayuan="" & name1 & "你已经有"& Application("aqjh_baowuname") &"了必须修炼以后才能到孤岛寻宝！" 
      else
            randomize timer
            rand=Int(79 * Rnd + 10) 
                 if rand>=10 and rand<=50 then
                        huayuan="" & name1 & "在孤岛上" & sy1 & sh1 & "，你发现了别人留下的100两银子！"
                        sql="update 用户 set 体力=体力-11,银两=银两+100,操作时间=now() where 姓名='" & name1 & "'"
                        conn.execute sql
                 elseif rand=51 then
                        huayuan="" & name1 & "在孤岛上" & sy1 & sh1 & "，你发现了别人留下的5000两银子！"
                        sql="update 用户 set 体力=体力-11,银两=银两+5000,操作时间=now() where 姓名='" & name1 & "'"
                        conn.execute sql
                 elseif rand=52 then
                        huayuan="" & name1 & "在孤岛上" & sy1 & sh1 & "，你发现了别人留下的500两银子！"
                        sql="update 用户 set 体力=体力-11,银两=银两+500,操作时间=now() where 姓名='" & name1 & "'"
                        conn.execute sql
                 elseif rand=53 then
                        huayuan="" & name1 & "在孤岛上" & sy1 & sh1 & "，你碰到了野兽，力战之下你打死了野兽，但却伤了3000内力！"
                        sql="update 用户 set 体力=体力-11,内力=内力-3000,操作时间=now() where 姓名='" & name1 & "'"
                        conn.execute sql
                 elseif rand=54 then
                        huayuan="" & name1 & "在孤岛上" & sy1 & sh1 & "，被脚下石头一拌摔了一跤，丢了100两银子！"
                        sql="update 用户 set 体力=体力-11,银两=银两-100,操作时间=now() where 姓名='" & name1 & "'"
                        conn.execute sql
                 elseif rand=55 or rand=85 then
                        huayuan="" & name1 & "在孤岛上" & sy1 & sh1 & "，你发现了一奇特的山草药，一吃之下，身中剧毒而死！"                         
                        sql="update 用户 set 状态='死',操作时间=now() where 姓名='" & name1 & "'"
                        conn.execute sql 
                        sql= "insert into l(b,a,c,e,d) values ('" & name1 & "',now(),'" & aqjh_name & "','在孤岛中吃错有毒食物而死','人命')"
                        'sql="insert into 人命(死者,时间,凶手,死因) values ('" & name1 & "',now(),'无','在孤岛中吃错有毒食物而死')"
                        conn.execute sql
                       call boot(name1)
                 elseif rand=56 then
                        huayuan="" & name1 & "在孤岛上" & sy1 & sh1 & "，你发现了古代宝藏，正想去拿之际，宝藏的机关打开，你被乱箭射伤，损失体力10000！"
                        sql="update 用户 set 体力=体力-10000,操作时间=now() where 姓名='" & name1 & "'"
                        conn.execute sql
                 elseif rand=57 or rand=86 then
                        huayuan="" & name1 & "在孤岛上" & sy1 & sh1 & "，你失足跌落悬崖，从此江湖少了一个无名大侠！"
                        sql="update 用户 set 状态='死' where 姓名='" & name1 & "'"
                        conn.execute sql 
                        sql= "insert into l(b,a,c,e,d) values ('" & name1 & "',now(),'" & aqjh_name & "','在孤岛失足跌落悬崖','人命')"
                        'sql="insert into 人命(死者,时间,凶手,死因) values ('" & name1 & "',now(),'无','在孤岛失足跌落悬崖')"
                        conn.execute sql
                      call boot(name1)
                 elseif rand=58 then
                        huayuan="" & name1 & "在孤岛上" & sy1 & sh1 & "，巧遇一绝色女子被老虎追赶，你英雄救美，魅力上升1000！"
                        sql="update 用户 set 体力=体力-100,魅力=魅力+1000,操作时间=now() where 姓名='" & name1 & "'"
                        conn.execute sql
                 elseif rand=59 then
                        huayuan="" & name1 & "在孤岛上" & sy1 & sh1 & "，巧遇一世外高人，经指点，你道德上升200！"
                        sql="update 用户 set 体力=体力-11,道德=道德+200,操作时间=now() where 姓名='" & name1 & "'"
                        conn.execute sql
                 elseif rand=60 then
                        huayuan="" & name1 & "在孤岛上" & sy1 & sh1 & "，洗了洗脸！魅力上升50"
                        sql="update 用户 set 体力=体力-10,魅力=魅力+50 where,操作时间=now() 姓名='" & name1 & "'"
                        conn.execute sql
                 elseif rand=61 then
                        huayuan="" & name1 & "在孤岛上" & sy1 & sh1 & "，巧遇一绝色女子，你把她救出孤岛，道德上升500！"
                        sql="update 用户 set 体力=体力-1,道德=道德+500,操作时间=now() where 姓名='" & name1 & "'"
                        conn.execute sql
                 elseif rand=62 then
                        huayuan="" & name1 & "在孤岛上" & sy1 & sh1 & "，发现了一个大宝藏，经过一番冒险，你成功拿到五十万宝藏及武功绝学,内力上升1000！"
                        sql="update 用户 set 内力=内力+1000,银两=银两+500000,操作时间=now() where 姓名='" & name1 & "'"
                        conn.execute sql
                 elseif rand>=63 or rand<=66 then
                        huayuan="" & name1 & "在孤岛上" & sy1 & sh1 & "，你发现了别人留下的1000两银子！"
                        sql="update 用户 set 体力=体力-11,银两=银两+1000,操作时间=now() where 姓名='" & name1 & "'"
                        conn.execute sql
                 elseif rand=67 then
                        huayuan="" & name1 & "在孤岛上" & sy1 & sh1 & "，洗了洗脸！魅力上升100"
                        sql="update 用户 set 体力=体力-10,魅力=魅力+100,操作时间=now() where 姓名='" & name1 & "'"
                        conn.execute sql
                 elseif rand=68 then
                        huayuan="" & name1 & "在孤岛上" & sy1 & sh1 & "，你发现了大宝藏，但在拿的途中被乱箭射死！"
                        sql="update 用户 set 状态='死' where 姓名='" & name1 & "'"
                        conn.execute sql 
                       sql= "insert into l(b,a,c,e,d) values ('" & name1 & "',now(),'" & aqjh_name & "','孤岛中拿宝藏时被乱箭射死','人命')" 
                      '  sql="insert into 人命(死者,时间,凶手,死因) values ('" & name1 & "',now(),'无','孤岛中拿宝藏时被乱箭射死')"
                        conn.execute sql
                     call boot(name1)
                 elseif rand=69 then
                        huayuan="" & name1 & "在孤岛上" & sy1 & sh1 & "，碰到毒蛇，幸亏武功高强，虚惊一场！"
                        sql="update 用户 set 体力=体力-11 where,操作时间=now() 姓名='" & name1 & "'"
conn.execute sql
                 elseif rand=70 or rand=87 then
huayuan="" & name1 & "在孤岛上" & sy1 & sh1 & "，遇到仇敌刚巧也在此地，在大斗30回合后，你不敌而死于仇敌剑下！"
                        sql="update 用户 set 状态='死' where 姓名='" & name1 & "'"
                        conn.execute sql 
                        sql= "insert into l(b,a,c,e,d) values ('" & name1 & "',now(),'" & aqjh_name & "','在孤岛冒险中被杀','人命')" 
                       ' sql="insert into 人命(死者,时间,凶手,死因) values ('" & name1 & "',now(),'无','在孤岛冒险中被杀')"
                        conn.execute sql
                       call boot(name1)
                 elseif rand=71 then
                        huayuan="" & name1 & "在孤岛上" & sy1 & sh1 & "，你发现了别人留下的1万两银子！"
                        sql="update 用户 set 体力=体力-11,银两=银两+10000,操作时间=now() where 姓名='" & name1 & "'"
                        conn.execute sql
                 elseif rand=72 then
                        huayuan="" & name1 & "在孤岛上" & sy1 & sh1 & "，突然在墙角发现1两银子！"
                        sql="update 用户 set 体力=体力-11,银两=银两+1,操作时间=now() where 姓名='" & name1 & "'"
                        conn.execute sql
                 elseif rand=73 then
                        huayuan="" & name1 & "在孤岛上" & sy1 & sh1 & "，你发现了一种奇特的山草药，一吃之下，内力增进2000！"
                        sql="update 用户 set 体力=体力-11,内力=内力+2000,操作时间=now() where 姓名='" & name1 & "'"
                        conn.execute sql
                 elseif rand=74 then
                        huayuan="" & name1 & "在孤岛上" & sy1 & sh1 & "，你发现了别人留下的300两银子！"
                        sql="update 用户 set 体力=体力-11,银两=银两+300,操作时间=now() where 姓名='" & name1 & "'"
                        conn.execute sql
                 elseif rand=75 then
                        huayuan="" & name1 & "在孤岛上" & sy1 & sh1 & "，你发现了别人留下的200两银子！"
                        sql="update 用户 set 体力=体力-11,银两=银两+200,操作时间=now() where 姓名='" & name1 & "'"
                        conn.execute sql
                 elseif rand=76 then
                        huayuan="" & name1 & "在孤岛上" & sy1 & sh1 & "，洗了洗脸！魅力上升100"
                        sql="update 用户 set 体力=体力-10,魅力=魅力+100,操作时间=now() where 姓名='" & name1 & "'"
                        conn.execute sql
                 elseif rand=77 then
                        huayuan="" & name1 & "在孤岛上" & sy1 & sh1 & "，失足跌下山坡！大难不死，体力损失5000，但大难不死，必有后福，发现5000两银子"
                        sql="update 用户 set 体力=体力-5000,银两=银两+5000,操作时间=now() where 姓名='" & name1 & "'"
                        conn.execute sql
                 elseif rand=78 then
                        huayuan="" & name1 & "在孤岛上" & sy1 & sh1 & "，巧遇一世外高人，经指点，你道德上升2000！"
                        sql="update 用户 set 体力=体力-11,道德=道德+2000,操作时间=now() where 姓名='" & name1 & "'"
                        conn.execute sql
                 elseif rand=79 or rand=88 then
                        huayuan="" & name1 & "在孤岛上" & sy1 & sh1 & "，你在冒险中碰到狮子，力战下还是惨死于狮子口下！"
                        sql="update 用户 set 状态='死' where 姓名='" & name1 & "'"
                        conn.execute sql 
                        sql= "insert into l(b,a,c,e,d) values ('" & name1 & "',now(),'" & aqjh_name & "','在孤岛冒险中身亡','人命')" 
                        'sql="insert into 人命(死者,时间,凶手,死因) values ('" & name1 & "',now(),'无','在孤岛冒险中身亡')"
                        conn.execute sql
                        call boot(name1)
                 elseif rand=80 then
                        huayuan="" & name1 & "在孤岛上" & sy1 & sh1 & "，你发现了上乘的内力心法，内力大进1000点！"
                        sql="update 用户 set 体力=体力-11,内力=内力+1000,操作时间=now() where 姓名='" & name1 & "'"
                        conn.execute sql
                 elseif rand=81 then
                        huayuan="" & name1 & "在孤岛上" & sy1 & sh1 & "，巧遇一绝色女子，你色心大起，对其施暴，道德下降1000！"
                        sql="update 用户 set 体力=体力-1000,道德=道德-1000,操作时间=now() where 姓名='" & name1 & "'"
                        conn.execute sql
                 elseif rand=82 or rand=89 then
                        huayuan="" & name1 & "在孤岛上" & sy1 & sh1 & "，你因找不到食物，活活饿死了，从此江湖少了一个无名大侠！"
                        sql="update 用户 set 状态='死' where 姓名='" & name1 & "'"
                        conn.execute sql 
                        
			sql= "insert into l(b,a,c,e,d) values ('" & name1 & "',now(),'" & aqjh_name & "','在孤岛冒险中身亡','人命')"
                        'sql="insert into 人命(死者,时间,凶手,死因) values ('" & name1 & "',now(),'无','在孤岛冒险中身亡')"
                        conn.execute sql
                     call boot(name1)
                 
                      
                else
                        huayuan="" & name1 & "在孤岛上" & sy1 & sh1 & "，你发现了一个藏宝库，但没钥匙，打不开门！"
                        sql="update 用户 set 体力=体力-11,操作时间=now() where 姓名='" & name1 & "'"
                        conn.execute sql
                end if
       end if
end if
rs.close
set rs=nothing
conn.close
set conn=nothing
end function
sub boot(to1)
Application.Lock
onlinelist=Application("aqjh_onlinelist")
dim newonlinelist()
useronlinename=""
onliners=0
js=1
for i=1 to UBound(onlinelist) step 6
if CStr(onlinelist(i+1))<>CStr(to1) then
onliners=onliners+1
useronlinename=useronlinename & " " & onlinelist(i+1)
Redim Preserve newonlinelist(js),newonlinelist(js+1),newonlinelist(js+2),newonlinelist(js+3),newonlinelist(js+4),newonlinelist(js+5)
newonlinelist(js)=onlinelist(i)
newonlinelist(js+1)=onlinelist(i+1)
newonlinelist(js+2)=onlinelist(i+2)
newonlinelist(js+3)=onlinelist(i+3)
newonlinelist(js+4)=onlinelist(i+4)
newonlinelist(js+5)=onlinelist(i+5)
js=js+6
end if
next
useronlinename=useronlinename&" "
Application.UnLock
end sub
%>