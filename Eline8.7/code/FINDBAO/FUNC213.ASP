<%
function huayuan(name1,sh1,sy1)
titleline=1
set conn=server.createobject("adodb.connection") 
conn.open Application("sjjh_usermdb")
set rs=server.CreateObject ("adodb.recordset")
sql="SELECT * FROM 用户 WHERE 姓名='" & name1 & "'"
Set Rs=conn.Execute(sql)
sex=rs("性别")
if rs.eof or rs.bof then
huayuan=""
else
if rs("道德")<500 then
Response.write "你的道德低于500，不能进入"
Response.end
end if
if rs("体力")<1000 or rs("内力")<1000 then
Response.write "你的体力或者内力低于1000，不能进入"
Response.end
end if
if rs("攻击")<1000 or rs("防御")<1000 then
Response.write "你的攻击或者防御低于1000，不能进入"
Response.end
end if
if application("huayuan")=0 then
news
randomize
rand=Int(62 * Rnd + 20)

Select Case rand
case "10"
newer="" & name1 & "在孤岛上" & sy1 & sh1 & "，你发现了别人丢掉的300两银两！"
sql="update 用户 set 体力=体力-10,银两=银两+300 where 姓名='" & name1 & "'"
conn.execute sql
conn.close
case "11"
newer="" & name1 & "在孤岛上" & sy1 & sh1 & "，你发现野人参，到自由市场去买了600两银两！"
sql="update 用户 set 体力=体力-10,银两=银两+600 where 姓名='" & name1 & "'"
conn.execute sql
conn.close
case "12"
newer="" & name1 & "在孤岛上" & sy1 & sh1 & "，你碰到了野兽，力战之下你打死了野兽，但却伤了3000内力！"
sql="update 用户 set 体力=体力-10,内力=内力-3000 where 姓名='" & name1 & "'"
conn.execute sql
conn.close
case "13"
newer="" & name1 & "在孤岛上" & sy1 & sh1 & "，被脚下石头一拌摔了一跤，丢了100两银两！"
sql="update 用户 set 体力=体力-10,银两=银两-100 where 姓名='" & name1 & "'"
conn.execute sql
conn.close
case "14"
newer="" & name1 & "在孤岛上" & sy1 & sh1 & "，你发现了别人埋的的1000两银两！"
sql="update 用户 set 体力=体力-10,银两=银两+1000 where 姓名='" & name1 & "'"
conn.execute sql
conn.close
case "15"
newer="" & name1 & "在孤岛上" & sy1 & sh1 & "，你发现了别人留下的100两银两！"
sql="update 用户 set 体力=体力-10,银两=银两+100 where 姓名='" & name1 & "'"
conn.execute sql
conn.close
case "16"
newer="" & name1 & "在孤岛上" & sy1 & sh1 & "，你遇见了敌人，大战300回合以后，体力降低了2000！"
sql="update 用户 set 体力=体力-2000 where 姓名='" & name1 & "'"
conn.execute sql
conn.close
case "17"
newer="" & name1 & "在孤岛上" & sy1 & sh1 & "，你发现了古代宝藏，正想去拿之际，宝藏的机关打开，你被乱箭射伤，损失体力5000！"
sql="update 用户 set 体力=体力-5000 where 姓名='" & name1 & "'"
conn.execute sql
conn.close
case "18"
newer="" & name1 & "在孤岛上" & sy1 & sh1 & "，你发现了别人留下的100两银两！"
sql="update 用户 set 体力=体力-10,银两=银两+100 where 姓名='" & name1 & "'"
conn.execute sql
conn.close
case "19"
newer="" & name1 & "在孤岛上" & sy1 & sh1 & "，被脚下石头一拌摔了一跤，丢了100两银两！"
sql="update 用户 set 体力=体力-10,银两=银两-100 where 姓名='" & name1 & "'"
conn.execute sql
conn.close
case "20"
newer="" & name1 & "在孤岛上" & sy1 & sh1 & "，你失足跌落悬崖，内力减少50000！"
sql="update 用户 set 体力=体力-10,内力=内力-50000 where 姓名='" & name1 & "'"
conn.execute sql
conn.close
case "21"
newer="" & name1 & "在孤岛上" & sy1 & sh1 & "，你发现了别人留下的100两银两！"
sql="update 用户 set 体力=体力-10,银两=银两+100 where 姓名='" & name1 & "'"
conn.execute sql
conn.close
case "22"
if sex="男" then
newer="" & name1 & "在孤岛上" & sy1 & sh1 & "，巧遇一绝色女子被老虎追赶，你英雄救美！"
else
newer="" & name1 & "在孤岛上" & sy1 & sh1 & "，巧遇一美男子，爱上了他！"
end if
sql="update 用户 set 体力=体力-100 where 姓名='" & name1 & "'"
conn.execute sql
conn.close
case "23"
newer="" & name1 & "在孤岛上" & sy1 & sh1 & "，你发现了别人留下的100两银两！"
sql="update 用户 set 体力=体力-10,银两=银两+100 where 姓名='" & name1 & "'"
conn.execute sql
conn.close
case "24"
newer="" & name1 & "在孤岛上" & sy1 & sh1 & "，巧遇一世外高人，经指点，你的资质上升50，内力上升200！"
sql="update 用户 set 体力=体力-10,资质=资质+50,内力=内力+200 where 姓名='" & name1 & "'"
conn.execute sql
conn.close
case "25"
newer="" & name1 & "在孤岛上" & sy1 & sh1 & "，洗了洗脸！体力上升500！"
sql="update 用户 set 体力=体力+500 where 姓名='" & name1 & "'"
conn.execute sql
conn.close
case "26"
newer="" & name1 & "在孤岛上" & sy1 & sh1 & "，你发现了别人留下的100两银两！"
sql="update 用户 set 体力=体力-10,银两=银两+100 where 姓名='" & name1 & "'"
conn.execute sql
conn.close
case "27"
if sex="男" then
newer="" & name1 & "在孤岛上" & sy1 & sh1 & "，巧遇一绝色女子，你把她救出孤岛，道德上升500！"
else
newer="" & name1 & "在孤岛上" & sy1 & sh1 & "，巧遇美男子，道德上升500！"
end if
sql="update 用户 set 体力=体力-1,道德=道德+500 where 姓名='" & name1 & "'"
conn.execute sql
conn.close
case "28"
newer="" & name1 & "在孤岛上" & sy1 & sh1 & "，你发现了别人留下的100两银两！"
sql="update 用户 set 体力=体力-10,银两=银两+100 where 姓名='" & name1 & "'"
conn.execute sql
conn.close
case "29"
newer="" & name1 & "在孤岛上" & sy1 & sh1 & "，你发现了别人留下的100两银两！"
sql="update 用户 set 体力=体力-10,银两=银两+100 where 姓名='" & name1 & "'"
conn.execute sql
conn.close
case "30"
newer="" & name1 & "在孤岛上" & sy1 & sh1 & "，发现了一个大宝藏，经过一番冒险，你成功拿到五十万宝藏及武功绝学！"
sql="update 用户 set 内力=内力+1000,银两=银两+500000 where 姓名='" & name1 & "'"
conn.execute sql
conn.close
case "31"
newer="" & name1 & "在孤岛上" & sy1 & sh1 & "，你发现了别人留下的100两银两！"
sql="update 用户 set 体力=体力-10,银两=银两+100 where 姓名='" & name1 & "'"
conn.execute sql
conn.close
case "32"
newer="" & name1 & "在孤岛上" & sy1 & sh1 & "，你发现了别人留下的100两银两！"
sql="update 用户 set 体力=体力-10,银两=银两+100 where 姓名='" & name1 & "'"
conn.execute sql
conn.close
case "33"
newer="" & name1 & "在孤岛上" & sy1 & sh1 & "，你发现了别人留下的100两银两！"
sql="update 用户 set 体力=体力-10,银两=银两+100 where 姓名='" & name1 & "'"
conn.execute sql
conn.close
case "34"
newer="" & name1 & "在孤岛上" & sy1 & sh1 & "，你弄丢了100两银两！"
sql="update 用户 set 体力=体力-10,银两=银两-100 where 姓名='" & name1 & "'"
conn.execute sql
conn.close
case "35"
newer="" & name1 & "在孤岛上" & sy1 & sh1 & "，你发现了别人留下的100两银两！"
sql="update 用户 set 体力=体力-10,银两=银两+100 where 姓名='" & name1 & "'"
conn.execute sql
conn.close
case "36"
newer="" & name1 & "在孤岛上" & sy1 & sh1 & "，你扭伤了脚。花100两银两的医疗费！"
sql="update 用户 set 体力=体力-10,银两=银两-100 where 姓名='" & name1 & "'"
conn.execute sql
conn.close
case "37"
newer="" & name1 & "在孤岛上" & sy1 & sh1 & "，你发现了别人留下的1000两银两！"
sql="update 用户 set 体力=体力-10,银两=银两+1000 where 姓名='" & name1 & "'"
conn.execute sql
conn.close
case "38"
newer="" & name1 & "在孤岛上" & sy1 & sh1 & "，你发现了别人留下的100两银两！"
sql="update 用户 set 体力=体力-10,银两=银两+100 where 姓名='" & name1 & "'"
conn.execute sql
conn.close
case "39"
newer="" & name1 & "在孤岛上" & sy1 & sh1 & "，你发现了一奇特的山草药，一吃之下，身中剧毒，虽然花了1000两银两治病，但还是留下了后遗症，降低了100点资质"
sql="update 用户 set 资质=资质-100,银两=银两-10000 where 姓名='" & name1 & "'"
conn.execute sql
conn.close
case "40"
newer="" & name1 & "在孤岛上" & sy1 & sh1 & "，你发现了别人留下的100两银两！"
sql="update 用户 set 体力=体力-10,银两=银两+100 where 姓名='" & name1 & "'"
conn.execute sql
conn.close
case "41"
newer="" & name1 & "在孤岛上" & sy1 & sh1 & "，你发现了别人留下的100两银两！"
sql="update 用户 set 体力=体力-10,银两=银两+100 where 姓名='" & name1 & "'"
conn.execute sql
conn.close
case "42"
newer="" & name1 & "在孤岛上" & sy1 & sh1 & "，你发现了别人留下的100两银两！"
sql="update 用户 set 体力=体力-10,银两=银两+100 where 姓名='" & name1 & "'"
conn.execute sql
conn.close
case "43"
newer="" & name1 & "在孤岛上" & sy1 & sh1 & "，洗了洗脸！体力上升300"
sql="update 用户 set 体力=体力+300 where 姓名='" & name1 & "'"
conn.execute sql
conn.close
case "44"
if sex="男" then
newer="" & name1 & "在孤岛上" & sy1 & sh1 & "，巧遇一绝色女子，你色心大起，道德下降1000！"
else
newer="" & name1 & "在孤岛上" & sy1 & sh1 & "，巧遇一美男子，你想勾引他，道德下降1000！"
end if
sql="update 用户 set 体力=体力-10,道德=道德-1000 where 姓名='" & name1 & "'"
conn.execute sql
conn.close
case "45"
newer="" & name1 & "在孤岛上" & sy1 & sh1 & "，你丢失了100两银两！"
sql="update 用户 set 体力=体力-10,银两=银两-100 where 姓名='" & name1 & "'"
conn.execute sql
conn.close
case "46"
newer="" & name1 & "在孤岛上" & sy1 & sh1 & "，你发现了别人留下的100两银两！"
sql="update 用户 set 体力=体力-10,银两=银两+100 where 姓名='" & name1 & "'"
conn.execute sql
conn.close
case "47"
newer="" & name1 & "在孤岛上" & sy1 & sh1 & "，碰到毒蛇，幸亏武功高强，虚惊一场！"
sql="update 用户 set 体力=体力-10 where 姓名='" & name1 & "'"
conn.execute sql
conn.close
case "48"
newer="" & name1 & "在孤岛上" & sy1 & sh1 & "，你发现了别人留下的100两银两！"
sql="update 用户 set 体力=体力-10,银两=银两+100 where 姓名='" & name1 & "'"
conn.execute sql
conn.close
case "49"
newer="" & name1 & "在孤岛上" & sy1 & sh1 & "，你发现了别人留下的100两银两！"
sql="update 用户 set 体力=体力-10,银两=银两+100 where 姓名='" & name1 & "'"
conn.execute sql
conn.close
case "50"
if sex="男" then
newer="" & name1 & "在孤岛上" & sy1 & sh1 & "，巧遇一绝色女子，上前调戏，道德下降100！"
else
newer="" & name1 & "在孤岛上" & sy1 & sh1 & "，巧遇一美男子，上前勾引，可是人家不理你，道德下降100！"
end if
sql="update 用户 set 体力=体力-1000,道德=道德-100 where 姓名='" & name1 & "'"
conn.execute sql
conn.close
case "51"
newer="" & name1 & "在孤岛上" & sy1 & sh1 & "，你发现了别人留下的50两银两！"
sql="update 用户 set 体力=体力-10,银两=银两+50 where 姓名='" & name1 & "'"
conn.execute sql
conn.close
case "52"
newer="" & name1 & "在孤岛上" & sy1 & sh1 & "，你发现了别人留下的10000两银两！"
sql="update 用户 set 体力=体力-10,银两=银两+10000 where 姓名='" & name1 & "'"
conn.execute sql
conn.close
case "53"
newer="" & name1 & "在孤岛上" & sy1 & sh1 & "，突然在墙角发现10两银两！"
sql="update 用户 set 体力=体力-10,银两=银两+10 where 姓名='" & name1 & "'"
conn.execute sql
conn.close
case "54"
newer="" & name1 & "在孤岛上" & sy1 & sh1 & "，你发现了一种奇特的山草药，一吃之下，内力增进2000，内力提高1000！"
sql="update 用户 set 体力=体力-10,内力=内力+2000,内力=内力+1000 where 姓名='" & name1 & "'"
conn.execute sql
conn.close
case "55"
newer="" & name1 & "在孤岛上" & sy1 & sh1 & "，你发现了别人留下的80两银两！"
sql="update 用户 set 体力=体力-10,银两=银两+80 where 姓名='" & name1 & "'"
conn.execute sql
conn.close
case "56"
newer="" & name1 & "在孤岛上" & sy1 & sh1 & "，你发现了别人留下的120两银两！"
sql="update 用户 set 体力=体力-10,银两=银两+120 where 姓名='" & name1 & "'"
conn.execute sql
conn.close
case "57"
newer="" & name1 & "在孤岛上" & sy1 & sh1 & "，你不小心弄坏了别人的房子，赔了500两银两！"
sql="update 用户 set 体力=体力-10,银两=银两-500 where 姓名='" & name1 & "'"
conn.execute sql
conn.close
case "58"
newer="" & name1 & "在孤岛上" & sy1 & sh1 & "，你发现了别人留下的300两银两！"
sql="update 用户 set 体力=体力-10,银两=银两+300 where 姓名='" & name1 & "'"
conn.execute sql
conn.close
case "59"
newer="" & name1 & "在孤岛上" & sy1 & sh1 & "，你发现了别人留下的200两银两！"
sql="update 用户 set 体力=体力-10,银两=银两+200 where 姓名='" & name1 & "'"
conn.execute sql
conn.close
case "60"
newer="" & name1 & "在孤岛上" & sy1 & sh1 & "，你发现了别人留下的1000两银两！"
sql="update 用户 set 体力=体力-10,银两=银两+1000 where 姓名='" & name1 & "'"
conn.execute sql
conn.close
case "61"
newer="" & name1 & "在孤岛上" & sy1 & sh1 & "，洗了洗脸！体力上升300"
sql="update 用户 set 体力=体力+300 where 姓名='" & name1 & "'"
conn.execute sql
conn.close
case "62"
newer="" & name1 & "在孤岛上" & sy1 & sh1 & "，你发现了别人留下的300两银两！"
sql="update 用户 set 体力=体力-10,银两=银两+300 where 姓名='" & name1 & "'"
conn.execute sql
conn.close
case "63"
newer="" & name1 & "在孤岛上" & sy1 & sh1 & "，你发现了别人留下的100两银两！"
sql="update 用户 set 体力=体力-10,银两=银两+100 where 姓名='" & name1 & "'"
conn.execute sql
conn.close
case "64"
newer="" & name1 & "在孤岛上" & sy1 & sh1 & "，失足跌下山坡！大难不死，内力损失5000,银两+5000，但大难不死，必有后福，发现5000两银两"
sql="update 用户 set 内力=内力-5000,银两=银两+5000 where 姓名='" & name1 & "'"
conn.execute sql
conn.close
case "65"
newer="" & name1 & "在孤岛上" & sy1 & sh1 & "，你发现了别人留下的100两银两！"
sql="update 用户 set 体力=体力-10,银两=银两+100 where 姓名='" & name1 & "'"
conn.execute sql
conn.close
case "66"
newer="" & name1 & "在孤岛上" & sy1 & sh1 & "，巧遇一世外高人，经指点，你内力上升800！"
sql="update 用户 set 体力=体力-10,内力=内力+800 where 姓名='" & name1 & "'"
conn.execute sql
conn.close
case "67"
newer="" & name1 & "在孤岛上" & sy1 & sh1 & "，你发现了别人留下的100两银两！"
sql="update 用户 set 体力=体力-10,银两=银两+100 where 姓名='" & name1 & "'"
conn.execute sql
conn.close
case "68"
newer="" & name1 & "在孤岛上" & sy1 & sh1 & "，你丢失了100两银两！"
sql="update 用户 set 体力=体力-10,银两=银两+100 where 姓名='" & name1 & "'"
conn.execute sql
conn.close
case "69"
if sex="男" then
newer="" & name1 & "在孤岛上" & sy1 & sh1 & "，路遇绝色女子遇难，英雄救美，道德上升1000！体力=体力-010"
else
newer="" & name1 & "在孤岛上" & sy1 & sh1 & "，路遇一美男遇难，舍身相救，道德上升1000！体力=体力-010"
end if
sql="update 用户 set 体力=体力-1000,道德=道德+1000 where 姓名='" & name1 & "'"
conn.execute sql
conn.close
case "70"
newer="" & name1 & "在孤岛上" & sy1 & sh1 & "，你发现了别人留下的100两银两体力=体力-10！"
sql="update 用户 set 体力=体力-10,银两=银两+100 where 姓名='" & name1 & "'"
conn.execute sql
conn.close
case "71"
newer="" & name1 & "在孤岛上" & sy1 & sh1 & "，你发现了别人留下的150两银两！体力=体力-10"
sql="update 用户 set 体力=体力-10,银两=银两+150 where 姓名='" & name1 & "'"
conn.execute sql
conn.close
case "72"
newer="" & name1 & "在孤岛上" & sy1 & sh1 & "，你钱太多了，扔了自己的100两银两！体力=体力-10"
sql="update 用户 set 体力=体力-10,银两=银两-100 where 姓名='" & name1 & "'"
conn.execute sql
conn.close
case "73"
newer="" & name1 & "在孤岛上" & sy1 & sh1 & "，你发现了上乘的内力心法，内力大进1000点，体力也增加了1000点！"
sql="update 用户 set 内力=内力+1000,体力=体力+1000 where 姓名='" & name1 & "'"
conn.execute sql
conn.close
case "74"
if sex="男" then
newer="" & name1 & "在孤岛上" & sy1 & sh1 & "，巧遇一绝色女子，你色心大起，对其施暴，道德下降5000！体力=体力-10"
else
newer="" & name1 & "在孤岛上" & sy1 & sh1 & "，巧遇一美男子，你春心动，上前搭讪，道德下降5000！体力=体力-10"
end if
sql="update 用户 set 体力=体力-10,道德=道德-5000 where 姓名='" & name1 & "'"
conn.execute sql
conn.close
case "75"
newer="" & name1 & "在孤岛上" & sy1 & sh1 & "，你发现了别人留下的100两银两！体力=体力-10"
sql="update 用户 set 体力=体力-10,银两=银两+100 where 姓名='" & name1 & "'"
conn.execute sql
conn.close
case "76"
newer="" & name1 & "在孤岛上" & sy1 & sh1 & "，你发现了别人留下的100两银两！体力=体力-10"
sql="update 用户 set 体力=体力-10,银两=银两+100 where 姓名='" & name1 & "'"
conn.execute sql
conn.close
case "77"
newer="" & name1 & "在孤岛上" & sy1 & sh1 & "，你发现了别人留下的2000两银两！体力=体力-10"
sql="update 用户 set 体力=体力-10,银两=银两+2000 where 姓名='" & name1 & "'"
conn.execute sql
conn.close
case "78"
newer="" & name1 & "在孤岛上" & sy1 & sh1 & "，你发现了别人留下的1000两银两,体力=体力-10！"
sql="update 用户 set 体力=体力-10,银两=银两+1000 where 姓名='" & name1 & "'"
conn.execute sql
conn.close
case "79"
if sex="男" then
newer="" & name1 & "在孤岛上" & sy1 & sh1 & "，夜里和一绝色女子共宿，当了一夜君子，道德上升3000！体力=体力-1000"
else
newer="" & name1 & "在孤岛上" & sy1 & sh1 & "，单枪挑了曹操老賊，不愧为巾帼英雄，道德上升3000！体力=体力-1000"
end if
rs.close
sql="update 用户 set 体力=体力-1000,道德=道德+3000 where 姓名='" & name1 & "'"
conn.execute sql
conn.close

case "80"
sql="SELECT top 1 * FROM 宝物 WHERE 拿走='0' order by id desc"
Set Rs=conn.Execute(sql)
If Rs.Bof OR Rs.Eof Then
newer="" & name1 & "在孤岛上" & sy1 & sh1 & "，意外地发现了先人在孤岛留下的宝物藏身地，但是里面什么都没有！"
else
bao=rs("宝物名")
at=rs("攻击力")
fy=rs("防御力")
newer="" & name1 & "在孤岛上" & sy1 & sh1 & "，意外地发现了先人在孤岛留下的宝物藏身地，得到了稀世奇珍" & bao & "！！"
sql="insert into 物品(名称,属性,所有者,攻击,防御,价格,数量,特效,寄售,装备)values('"& bao &"','武器','"&name1&"'," & at & "," & fy & ",0,1,'无',False,False)"
conn.execute sql
sql="update 宝物 set 拿走=1,获宝者='"&name1&"', 时间=now() where 宝物名='" & bao & "'"
set Rs=conn.Execute(sql)
conn.close
end if

case "81"
newer="" & name1 & "在孤岛上" & sy1 & sh1 & "，你发现了别人留下的1000两银两！"
sql="update 用户 set 体力=体力-10,银两=银两+1000 where 姓名='" & name1 & "'"
conn.execute sql
conn.close

case "82"
newer="" & name1 & "在孤岛上" & sy1 & sh1 & "，你发现了别人留下的100两银两！"
sql="update 用户 set 体力=体力-10,银两=银两+100 where 姓名='" & name1 & "'"
conn.execute sql
conn.close
case "1"
newer="" & name1 & "在孤岛上" & sy1 & sh1 & "，你发现了一个藏宝库，但没钥匙，打不开门！"
sql="update 用户 set 体力=体力-10 where 姓名='" & name1 & "'"
conn.execute sql
conn.close
End Select
end if
end if
huayuan=newer
end function
%>