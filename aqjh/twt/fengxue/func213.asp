<%
function huayuan(name1,sh1,sy1)
titleline=1
%>
<%
set conn=server.createobject("adodb.connection")
conn.open Application("aqjh_usermdb")
set rs=server.CreateObject ("adodb.recordset")
sql="SELECT * FROM 用户 WHERE 姓名='" & session("aqjh_name") & "'"
Set Rs=conn.Execute(sql)
if rs.eof or rs.bof then
huayuan=""
else
if application("huayuan")=0 then
news
randomize
rand=Int(62 * Rnd + 20)
Select Case rand
case "1"
newer="" & session("aqjh_name") & "在舞剑的时候" & sy1 & sh1 & "，你发现了别人私藏的5000两银子！"
sql="update 用户 set 体力=体力-11,银两=银两+5000 where 姓名='" & session("aqjh_name") & "'"
conn.execute sql
conn.close
case "2"
newer="" & session("aqjh_name") & "在舞剑的时候" & sy1 & sh1 & "，你发现了别人私藏的500两银子！"
sql="update 用户 set 体力=体力-11,银两=银两+500 where 姓名='" & session("aqjh_name") & "'"
conn.execute sql
conn.close
case "3"
newer="" & session("aqjh_name") & "在舞剑的时候" & sy1 & sh1 & "，你遇到一个强敌，狠毒的你杀死了他，但却用去了3000内力，并遭到官府通缉，道德下降1000"
sql="update 用户 set 体力=体力-11,内力=内力-3000,道德=道德-1000 where 姓名='" & session("aqjh_name") & "'"
conn.execute sql
conn.close
case "4"
newer="" & session("aqjh_name") & "在舞剑的时候" & sy1 & sh1 & "，被居民追赶慌张逃跑，丢了1000两银子！"
sql="update 用户 set 体力=体力-11,银两=银两-1000 where 姓名='" & session("aqjh_name") & "'"
conn.execute sql
conn.close
case "5"
newer="" & session("aqjh_name") & "在舞剑的时候" & sy1 & sh1 & "，你发现了别人私藏的1000两银子！"
sql="update 用户 set 体力=体力-11,银两=银两+1000 where 姓名='" & session("aqjh_name") & "'"
conn.execute sql
conn.close
case "6"
newer="" & session("aqjh_name") & "在舞剑的时候" & sy1 & sh1 & "，你捉了几只麻雀，内力加1000体力加2000！"
sql="update 用户 set 体力=体力+2000,内力=内力+1000 where 姓名='" & session("aqjh_name") & "'"
conn.execute sql
conn.close
case "7"
newer="" & session("aqjh_name") & "在舞剑的时候" & sy1 & sh1 & "，你发现了别人家里的100两银子！"
sql="update 用户 set 体力=体力-11,银两=银两+100 where 姓名='" & session("aqjh_name") & "'"
conn.execute sql
conn.close
case "8"
newer="" & session("aqjh_name") & "在舞剑的时候" & sy1 & sh1 & "，你发现了宝藏，正想去拿之际，宝藏的机关打开，你被乱箭射伤，损失体力10000！"
sql="update 用户 set 体力=体力-10000 where 姓名='" & session("aqjh_name") & "'"
conn.execute sql
conn.close
case "9"
newer="" & session("aqjh_name") & "在舞剑的时候" & sy1 & sh1 & "，你发现了宝藏，正想去拿之际，宝藏的机关打开，你被乱箭射伤，损失体力10000！"
sql="update 用户 set 体力=体力-10000 where 姓名='" & session("aqjh_name") & "'"
conn.execute sql
conn.close
case "10"
newer="" & session("aqjh_name") & "在舞剑的时候" & sy1 & sh1 & "，你发现了别人私藏的5000两银子！"
sql="update 用户 set 体力=体力-11,银两=银两+5000 where 姓名='" & session("aqjh_name") & "'"
conn.execute sql
conn.close
case "11"
newer="" & session("aqjh_name") & "在舞剑的时候" & sy1 & sh1 & "，你发现了别人私藏的500两银子！"
sql="update 用户 set 体力=体力-11,银两=银两+500 where 姓名='" & session("aqjh_name") & "'"
conn.execute sql
conn.close
case "12"
newer="" & session("aqjh_name") & "在舞剑的时候" & sy1 & sh1 & "，你遇到一个强敌，狠毒的你杀死了他，但却用去了3000内力，并遭到官府通缉，道德下降1000"
sql="update 用户 set 体力=体力-11,内力=内力-3000,道德=道德-1000 where 姓名='" & session("aqjh_name") & "'"
conn.execute sql
conn.close
case "13"
newer="" & session("aqjh_name") & "在舞剑的时候" & sy1 & sh1 & "，被居民追赶慌张逃跑，丢了1000两银子！"
sql="update 用户 set 体力=体力-11,银两=银两-1000 where 姓名='" & session("aqjh_name") & "'"
conn.execute sql
conn.close
case "14"
newer="" & session("aqjh_name") & "在舞剑的时候" & sy1 & sh1 & "，你发现了别人私藏的1000两银子！"
sql="update 用户 set 体力=体力-11,银两=银两+1000 where 姓名='" & session("aqjh_name") & "'"
conn.execute sql
conn.close
case "15"
newer="" & session("aqjh_name") & "在舞剑的时候" & sy1 & sh1 & "，你捉了几只麻雀，内力加1000体力加2000！"
sql="update 用户 set 体力=体力+2000,内力=内力+1000 where 姓名='" & session("aqjh_name") & "'"
conn.execute sql
conn.close
case "16"
newer="" & session("aqjh_name") & "在舞剑的时候" & sy1 & sh1 & "，你发现了别人家里的100两银子！"
sql="update 用户 set 体力=体力-11,银两=银两+100 where 姓名='" & session("aqjh_name") & "'"
conn.execute sql
conn.close
case "17"
newer="" & session("aqjh_name") & "在舞剑的时候" & sy1 & sh1 & "，你发现了宝藏，正想去拿之际，宝藏的机关打开，你被乱箭射伤，损失体力10000！"
sql="update 用户 set 体力=体力-10000 where 姓名='" & session("aqjh_name") & "'"
conn.execute sql
conn.close
case "18"
newer="" & session("aqjh_name") & "在舞剑的时候" & sy1 & sh1 & "，你发现了别人私藏的100两银子！"
sql="update 用户 set 体力=体力-11,银两=银两+100 where 姓名='" & session("aqjh_name") & "'"
conn.execute sql
conn.close
case "19"
newer="" & session("aqjh_name") & "在舞剑的时候" & sy1 & sh1 & "，你发现了别人私藏的100两银子！"
sql="update 用户 set 体力=体力-11,银两=银两+100 where 姓名='" & session("aqjh_name") & "'"
conn.execute sql
conn.close
case "20"
newer="" & session("aqjh_name") & "在舞剑的时候" & sy1 & sh1 & "，一不小心掉进了山崖，从此江湖少了一个无名大侠！"
sql="update 用户 set 状态='死',攻击=攻击-1000,防御=防御-1000 where 姓名='" & session("aqjh_name") & "'"
conn.execute sql
sql="insert into l(b,a,c,e,d) values ('" & name1 & "',now(),'" & aqjh_name & "','舞剑时掉下山崖而死','人命')"
conn.execute sql
call boot(name1,"舞剑时掉下山崖而死")
case "21"
newer="" & session("aqjh_name") & "在舞剑的时候" & sy1 & sh1 & "，你发现了别人私藏的100两银子！"
sql="update 用户 set 体力=体力-11,银两=银两+100 where 姓名='" & session("aqjh_name") & "'"
conn.execute sql
conn.close
case "22"
newer="" & session("aqjh_name") & "在舞剑的时候" & sy1 & sh1 & "，巧遇一绝色女子被强盗强歼，你英雄救美，魅力上升1000！"
sql="update 用户 set 体力=体力-100,魅力=魅力+10000 where 姓名='" & session("aqjh_name") & "'"
conn.execute sql
conn.close
case "23"
newer="" & session("aqjh_name") & "在舞剑的时候" & sy1 & sh1 & "，你发现了别人私藏的100两银子！"
sql="update 用户 set 体力=体力-11,银两=银两+100 where 姓名='" & session("aqjh_name") & "'"
conn.execute sql
conn.close
case "24"
newer="" & session("aqjh_name") & "在舞剑的时候" & sy1 & sh1 & "，巧遇一世外高人，经指点，你道德上升2000！"
sql="update 用户 set 体力=体力-11,道德=道德+2000 where 姓名='" & session("aqjh_name") & "'"
conn.execute sql
conn.close
case "25"
newer="" & session("aqjh_name") & "在舞剑的时候" & sy1 & sh1 & "，洗了洗脸！魅力上升500"
sql="update 用户 set 体力=体力-10,知质=知质+500 where 姓名='" & session("aqjh_name") & "'"
conn.execute sql
conn.close
case "26"
newer="" & session("aqjh_name") & "在舞剑的时候" & sy1 & sh1 & "，你发现了别人丢在地上的10两银子！"
sql="update 用户 set 体力=体力-11,银两=银两+10 where 姓名='" & session("aqjh_name") & "'"
conn.execute sql
conn.close
case "27"
newer="" & session("aqjh_name") & "在舞剑的时候" & sy1 & sh1 & "，巧遇被囚禁的一绝色女子，你把她救出房子，道德上升5000！"
sql="update 用户 set 体力=体力-1,道德=道德+5000 where 姓名='" & session("aqjh_name") & "'"
conn.execute sql
conn.close
case "28"
newer="" & session("aqjh_name") & "在舞剑的时候" & sy1 & sh1 & "，你发现了别人私藏的100两银子！"
sql="update 用户 set 体力=体力-11,银两=银两+100 where 姓名='" & session("aqjh_name") & "'"
conn.execute sql
conn.close
case "29"
newer="" & session("aqjh_name") & "在舞剑的时候" & sy1 & sh1 & "，你发现了别人私藏的10000两银子！"
sql="update 用户 set 体力=体力-11,银两=银两+10000 where 姓名='" & session("aqjh_name") & "'"
conn.execute sql
conn.close
case "30"
newer="" & session("aqjh_name") & "在舞剑的时候" & sy1 & sh1 & "，发现了一个大宝藏地图，经过一番冒险，你成功拿到五十万宝藏及武功绝学（内力上升1000）！"
sql="update 用户 set 内力=内力+1000,银两=银两+500000 where 姓名='" & session("aqjh_name") & "'"
conn.execute sql
conn.close
case "31"
newer="" & session("aqjh_name") & "在舞剑的时候" & sy1 & sh1 & "，发现一把生锈的菜刀，拿去当了，得到100两！"
sql="update 用户 set 体力=体力-11,银两=银两+100 where 姓名='" & session("aqjh_name") & "'"
conn.execute sql
conn.close
case "32"
newer="" & session("aqjh_name") & "在舞剑的时候" & sy1 & sh1 & "，发现一种好吃的糕点，拿去卖得到500两！"
sql="update 用户 set 体力=体力+11,银两=银两+500 where 姓名='" & session("aqjh_name") & "'"
conn.execute sql
conn.close
case "33"
newer="" & session("aqjh_name") & "在舞剑的时候" & sy1 & sh1 & "，你发现了别人私藏的100两银子！"
sql="update 用户 set 体力=体力-11,银两=银两+100 where 姓名='" & session("aqjh_name") & "'"
conn.execute sql
conn.close
case "34"
newer="" & session("aqjh_name") & "在舞剑的时候" & sy1 & sh1 & "，你发现了别人私藏的100两银子！"
sql="update 用户 set 体力=体力-11,银两=银两+100 where 姓名='" & session("aqjh_name") & "'"
conn.execute sql
conn.close
case "35"
newer="" & session("aqjh_name") & "在舞剑的时候" & sy1 & sh1 & "，你发现了别人私藏的100两银子！"
sql="update 用户 set 体力=体力-11,银两=银两+100 where 姓名='" & session("aqjh_name") & "'"
conn.execute sql
conn.close
case "36"
newer="" & session("aqjh_name") & "在舞剑的时候" & sy1 & sh1 & "，你发现了别人私藏的100两银子！"
sql="update 用户 set 体力=体力-11,银两=银两+100 where 姓名='" & session("aqjh_name") & "'"
conn.execute sql
conn.close
case "37"
newer="" & session("aqjh_name") & "在舞剑的时候" & sy1 & sh1 & "，你发现了别人私藏的1000两银子！"
sql="update 用户 set 体力=体力-11,银两=银两+1000 where 姓名='" & session("aqjh_name") & "'"
conn.execute sql
conn.close
case "38"
newer="" & session("aqjh_name") & "在舞剑的时候" & sy1 & sh1 & "，你发现了别人私藏的100两银子！"
sql="update 用户 set 体力=体力-11,银两=银两+100 where 姓名='" & session("aqjh_name") & "'"
conn.execute sql
conn.close
case "39"
newer="" & session("aqjh_name") & "在舞剑的时候" & sy1 & sh1 & "，你偷走别人柜里的衣服,当得300两."
sql="update 用户 set 体力=体力-11,银两=银两+300 where 姓名='" & session("aqjh_name") & "'"
conn.execute sql
conn.close

case "40"
newer="" & session("aqjh_name") & "在舞剑的时候" & sy1 & sh1 & "，你发现了别人私藏的100两银子！"
sql="update 用户 set 体力=体力-11,银两=银两+100 where 姓名='" & session("aqjh_name") & "'"
conn.execute sql
conn.close
case "41"
newer="" & session("aqjh_name") & "在舞剑的时候" & sy1 & sh1 & "，你发现了别人私藏的100两银子！"
sql="update 用户 set 体力=体力-11,银两=银两+100 where 姓名='" & session("aqjh_name") & "'"
conn.execute sql
conn.close
case "42"
newer="" & session("aqjh_name") & "在舞剑的时候" & sy1 & sh1 & "，你发现了别人私藏的100两银子！"
sql="update 用户 set 体力=体力-11,银两=银两+100 where 姓名='" & session("aqjh_name") & "'"
conn.execute sql
conn.close
case "43"
newer="" & session("aqjh_name") & "在舞剑的时候" & sy1 & sh1 & "，砍杀了一个坏蛋！魅力上升100"
sql="update 用户 set 体力=体力-10,个性=个性+100 where 姓名='" & session("aqjh_name") & "'"
conn.execute sql
conn.close
case "44"
newer="" & session("aqjh_name") & "在舞剑的时候" & sy1 & sh1 & "，你发现了大宝藏，但在拿的途中被一个黑衣人偷袭而死！"
sql="update 用户 set 状态='死',攻击=攻击-1000,防御=防御-1000 where 姓名='" & session("aqjh_name") & "'"
conn.execute sql
sql="insert into l(b,a,c,e,d) values ('" & name1 & "',now(),'" & aqjh_name & "','舞剑时被偷袭而死','人命')"
conn.execute sql
call boot(name1,"舞剑时被偷袭而死")
conn.close
case "45"
newer="" & session("aqjh_name") & "在舞剑的时候" & sy1 & sh1 & "，你发现了别人私藏的100两银子！"
sql="update 用户 set 体力=体力-11,银两=银两+100 where 姓名='" & session("aqjh_name") & "'"
conn.execute sql
conn.close
case "46"
newer="" & session("aqjh_name") & "在舞剑的时候" & sy1 & sh1 & "，你发现了别人私藏的100两银子！"
sql="update 用户 set 体力=体力-11,银两=银两+100 where 姓名='" & session("aqjh_name") & "'"
conn.execute sql
conn.close
case "47"
newer="" & session("aqjh_name") & "在舞剑的时候" & sy1 & sh1 & "，碰到毒蛇，幸亏武功高强，虚惊一场！"
sql="update 用户 set 体力=体力-11 where 姓名='" & session("aqjh_name") & "'"
conn.execute sql
conn.close
case "48"
newer="" & session("aqjh_name") & "在舞剑的时候" & sy1 & sh1 & "，你发现了别人私藏的100两银子！"
sql="update 用户 set 体力=体力-11,银两=银两+100 where 姓名='" & session("aqjh_name") & "'"
conn.execute sql
conn.close
case "49"
newer="" & session("aqjh_name") & "在舞剑的时候" & sy1 & sh1 & "，你发现了别人私藏的100两银子！"
sql="update 用户 set 体力=体力-11,银两=银两+100 where 姓名='" & session("aqjh_name") & "'"
conn.execute sql
conn.close
case "50"
newer="" & session("aqjh_name") & "在舞剑的时候" & sy1 & sh1 & "，遇到仇敌刚巧也在此地，在大斗30回合后，你不敌而死于仇敌剑下！"
sql="update 用户 set 状态='死',攻击=攻击-1000,防御=防御-1000  where 姓名='" & session("aqjh_name") & "'"
conn.execute sql
sql="insert into into l(b,a,c,e,d) values ('" & name1 & "',now(),'" & aqjh_name & "','在舞剑中被杀','人命')"
conn.execute sql
call boot(name1,"在舞剑中被杀")
conn.close
case "51"
newer="" & session("aqjh_name") & "在舞剑的时候" & sy1 & sh1 & "，你发现了别人私藏的100两银子！"
sql="update 用户 set 体力=体力-11,银两=银两+100 where 姓名='" & session("aqjh_name") & "'"
conn.execute sql
conn.close
case "52"
newer="" & session("aqjh_name") & "在舞剑的时候" & sy1 & sh1 & "，你发现了别人私藏的1万两银子！"
sql="update 用户 set 体力=体力-11,银两=银两+10000 where 姓名='" & session("aqjh_name") & "'"
conn.execute sql
conn.close
case "53"
newer="" & session("aqjh_name") & "在舞剑的时候" & sy1 & sh1 & "，突然在墙角发现1两银子！"
sql="update 用户 set 体力=体力-11,银两=银两+1 where 姓名='" & session("aqjh_name") & "'"
conn.execute sql
conn.close
case "54"
newer="" & session("aqjh_name") & "在舞剑的时候" & sy1 & sh1 & "，白翻了一天什么也没有,看来是一家穷人.体力失去1000"
sql="update 用户 set 体力=体力-1000,银两=银两+0 where 姓名='" & session("aqjh_name") & "'"
conn.execute sql
conn.close
case "55"
newer="" & session("aqjh_name") & "在舞剑的时候" & sy1 & sh1 & "，你发现了别人私藏的100两银子！"
sql="update 用户 set 体力=体力-11,银两=银两+100 where 姓名='" & session("aqjh_name") & "'"
conn.execute sql
conn.close
case "56"
newer="" & session("aqjh_name") & "在舞剑的时候" & sy1 & sh1 & "，你发现了别人私藏的100两银子！"
sql="update 用户 set 体力=体力-11,银两=银两+100 where 姓名='" & session("aqjh_name") & "'"
conn.execute sql
conn.close
case "57"
newer="" & session("aqjh_name") & "在舞剑的时候" & sy1 & sh1 & "，你发现了别人私藏的100两银子！"
sql="update 用户 set 体力=体力-11,银两=银两+100 where 姓名='" & session("aqjh_name") & "'"
conn.execute sql
conn.close
case "58"
newer="" & session("aqjh_name") & "在舞剑的时候" & sy1 & sh1 & "，你发现了别人私藏的300两银子！"
sql="update 用户 set 体力=体力-11,银两=银两+300 where 姓名='" & session("aqjh_name") & "'"
conn.execute sql
conn.close
case "59"
newer="" & session("aqjh_name") & "在舞剑的时候" & sy1 & sh1 & "，你发现了别人私藏的200两银子！"
sql="update 用户 set 体力=体力-11,银两=银两+200 where 姓名='" & session("aqjh_name") & "'"
conn.execute sql
conn.close
case "60"
newer="" & session("aqjh_name") & "在舞剑的时候" & sy1 & sh1 & "，你发现了别人私藏的1000两银子！"
sql="update 用户 set 体力=体力-11,银两=银两+1000 where 姓名='" & session("aqjh_name") & "'"
conn.execute sql
conn.close
case "61"
newer="" & session("aqjh_name") & "在舞剑的时候" & sy1 & sh1 & "，劈死了一条大蛇！知质上升100"
sql="update 用户 set 体力=体力-10,知质=知质+100 where 姓名='" & session("aqjh_name") & "'"
conn.execute sql
conn.close
case "62"
newer="" & session("aqjh_name") & "在舞剑的时候" & sy1 & sh1 & "，你发现了别人私藏的100两银子！"
sql="update 用户 set 体力=体力-11,银两=银两+100 where 姓名='" & session("aqjh_name") & "'"
conn.execute sql
conn.close
case "63"
newer="" & session("aqjh_name") & "在舞剑的时候" & sy1 & sh1 & "，你发现了别人私藏的100两银子！"
sql="update 用户 set 体力=体力-11,银两=银两+100 where 姓名='" & session("aqjh_name") & "'"
conn.execute sql
conn.close
case "64"
newer="" & session("aqjh_name") & "在舞剑的时候" & sy1 & sh1 & "，被当地百姓发现追打，体力损失5000，但还是偷来了5000两银子"
sql="update 用户 set 体力=体力-5000,银两=银两+5000 where 姓名='" & session("aqjh_name") & "'"
conn.execute sql
conn.close
case "65"
newer="" & session("aqjh_name") & "在舞剑的时候" & sy1 & sh1 & "，你发现了别人私藏的100两银子！"
sql="update 用户 set 体力=体力-11,银两=银两+100 where 姓名='" & session("aqjh_name") & "'"
conn.execute sql
conn.close
case "66"
newer="" & session("aqjh_name") & "在舞剑的时候" & sy1 & sh1 & "，巧遇一世外高人，经指点，你道德上升20000！"
sql="update 用户 set 体力=体力-11,道德=道德+20000 where 姓名='" & session("aqjh_name") & "'"
conn.execute sql
conn.close
case "67"
newer="" & session("aqjh_name") & "在舞剑的时候" & sy1 & sh1 & "，你发现了别人私藏的100两银子！"
sql="update 用户 set 体力=体力-11,银两=银两+100 where 姓名='" & session("aqjh_name") & "'"
conn.execute sql
conn.close
case "68"
newer="" & session("aqjh_name") & "在舞剑的时候" & sy1 & sh1 & "，你发现了别人私藏的100两银子！"
sql="update 用户 set 体力=体力-11,银两=银两+100 where 姓名='" & session("aqjh_name") & "'"
conn.execute sql
conn.close
case "69"
newer="" & session("aqjh_name") & "在舞剑的时候" & sy1 & sh1 & "，被老虎咬伤，体力下降11"
sql="update 用户 set 体力=体力-11,银两=银两+0 where 姓名='" & session("aqjh_name") & "'"
conn.execute sql
conn.close
case "70"
newer="" & session("aqjh_name") & "在舞剑的时候" & sy1 & sh1 & "，你发现了别人私藏的100两银子！"
sql="update 用户 set 体力=体力-11,银两=银两+100 where 姓名='" & session("aqjh_name") & "'"
conn.execute sql
conn.close
case "71"
newer="" & session("aqjh_name") & "在舞剑的时候" & sy1 & sh1 & "，碰巧遇到官府人员,私下罚了你10000两才放你离开."
sql="update 用户 set 体力=体力-11,银两=银两-10000 where 姓名='" & session("aqjh_name") & "'"
conn.execute sql
conn.close
case "72"
newer="" & session("aqjh_name") & "在舞剑的时候" & sy1 & sh1 & "，你发现了别人私藏的100两银子！"
sql="update 用户 set 体力=体力-11,银两=银两+100 where 姓名='" & session("aqjh_name") & "'"
conn.execute sql
conn.close
case "73"
newer="" & session("aqjh_name") & "在舞剑的时候" & sy1 & sh1 & "，你一绝色女子,抓去卖到鸡窝,赚了10000两,道德下降1000"
sql="update 用户 set 体力=体力-11,银两=银两+10000,道德=道德-1000 where 姓名='" & session("aqjh_name") & "'"
conn.execute sql
conn.close
case "74"
newer="" & session("aqjh_name") & "在舞剑的时候" & sy1 & sh1 & "，巧遇一绝色女子，你色心大起，对其施暴，道德下降1000！"
sql="update 用户 set 体力=体力-1000,道德=道德-1000 where 姓名='" & session("aqjh_name") & "'"
conn.execute sql
conn.close
case "75"
newer="" & session("aqjh_name") & "在舞剑的时候" & sy1 & sh1 & "，你发现了别人私藏的100两银子！"
sql="update 用户 set 体力=体力-11,银两=银两+100 where 姓名='" & session("aqjh_name") & "'"
conn.execute sql
conn.close
case "76"
newer="" & session("aqjh_name") & "在舞剑的时候" & sy1 & sh1 & "，你发现了别人私藏的100两银子！"
sql="update 用户 set 体力=体力-11,银两=银两+100 where 姓名='" & session("aqjh_name") & "'"
conn.execute sql
conn.close
case "77"
newer="" & session("aqjh_name") & "在舞剑的时候" & sy1 & sh1 & "，你发现了别人私藏在树底下的100两银子,正想拿走,听见有一大帮人接近的脚步声,慌忙逃离了"
sql="update 用户 set 体力=体力-11,银两=银两+0 where 姓名='" & session("aqjh_name") & "'"
conn.execute sql
conn.close
case "78"
newer="" & session("aqjh_name") & "在舞剑的时候" & sy1 & sh1 & "，你发现了别人私藏的1000两银子！"
sql="update 用户 set 体力=体力-11,银两=银两+1000 where 姓名='" & session("aqjh_name") & "'"
conn.execute sql
conn.close
case "79"
newer="" & session("aqjh_name") & "在舞剑的时候" & sy1 & sh1 & "，你因找不到东西，灰溜溜得离开了。"
sql="update 用户 set 状态='正常' where 姓名='" & session("aqjh_name") & "'"
conn.execute sql
conn.close
case "79"
newer="" & session("aqjh_name") & "在舞剑的时候" & sy1 & sh1 & "，你因找不到东西，灰溜溜得离开了。"
sql="update 用户 set 状态='正常' where 姓名='" & session("aqjh_name") & "'"
conn.execute sql
conn.close
case "80"
newer="" & session("aqjh_name") & "在舞剑的时候" & sy1 & sh1 & "，你发现了别人私藏的100两银子！"
sql="update 用户 set 体力=体力-11,银两=银两+100 where 姓名='" & session("aqjh_name") & "'"
conn.execute sql
conn.close

case "82"
newer="" & session("aqjh_name") & "在舞剑的时候" & sy1 & sh1 & "，你发现了一具尸体,报官以后得到赏银1000"
sql="update 用户 set 体力=体力-11,银两=银两+1000 where 姓名='" & session("aqjh_name") & "'"
conn.execute sql
conn.close
case "1"
newer="" & session("aqjh_name") & "在舞剑的时候" & sy1 & sh1 & "，你发现了一个藏宝库，但没钥匙，打不开门！"
sql="update 用户 set 体力=体力-11 where 姓名='" & session("aqjh_name") & "'"
conn.execute sql
conn.close
End Select
end if
end if
huayuan=newer
end function
%>