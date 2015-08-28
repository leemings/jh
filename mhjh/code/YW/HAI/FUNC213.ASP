<!--#include file="data.asp"-->
<%
function huayuan(name1,sh1,sy1)
titleline=1
set conn=server.createobject("adodb.connection") 
conn.open Application("yx8_mhjh_connstr")
set rs=server.CreateObject ("adodb.recordset")
sql="SELECT 性别,道德,体力,内力,银两,杀人,攻击,防御 FROM 用户 WHERE 姓名='" & name1 & "'"
Set Rs=conn.Execute(sql)
sex=rs("性别")
if rs.eof or rs.bof then
huayuan=""
else
if rs("道德")<0 then
Response.write "你的道德低于0，不能进入"
Response.end
end if
if rs("体力")<10000 or rs("内力")<10000 then
Response.write "你的体力或者内力低于10000，不能进入"
Response.end
end if
if rs("杀人")>9 then
Response.write "意外,通缉犯居然跑到海外孤岛来了!遭到众人围攻!你真是赔了夫人又折兵啊,嘿嘿"
Response.end
end if
if rs("银两")<0 then
Response.write "有够倒霉的，什么也没得到,却赔上了自己的所带的全部银两,你回去吧"
Response.end
end if
if rs("攻击")<10000 or rs("防御")<10000 then
Response.write "你的攻击或者防御低于10000，不能进入"
Response.end
end if
if application("huayuan")=0 then
news
randomize
rand=Int(62 * Rnd + 20)

Select Case rand
case "10"
newer="" & name1 & "在孤岛上" & sy1 & sh1 & "，发现了别人丢掉的100000两银子！"
sql="update 用户 set 体力=体力-10,银两=银两+100000 where 姓名='" & name1 & "'"
conn.execute sql
conn.close
case "11"
newer="" & name1 & "在孤岛上" & sy1 & sh1 & "，发现野人参，到自由市场去买了200000两银子！"
sql="update 用户 set 体力=体力-10,银两=银两+200000 where 姓名='" & name1 & "'"
conn.execute sql
conn.close
case "12"
newer="" & name1 & "在孤岛上" & sy1 & sh1 & "，碰到了野兽，力战之下你打死了野兽，但却伤了300000内力！"
sql="update 用户 set 体力=体力-10,内力=内力-300000 where 姓名='" & name1 & "'"
conn.execute sql
conn.close
case "13"
newer="" & name1 & "在孤岛上" & sy1 & sh1 & "，被脚下石头一拌摔了一跤，丢了100000两银子！"
sql="update 用户 set 体力=体力-10,银两=银两-100000 where 姓名='" & name1 & "'"
conn.execute sql
conn.close
case "14"
newer="" & name1 & "在孤岛上" & sy1 & sh1 & "，发现了别人埋的的100000两银子！"
sql="update 用户 set 体力=体力-10,银两=银两+100000 where 姓名='" & name1 & "'"
conn.execute sql
conn.close
case "15"
newer="" & name1 & "在孤岛上" & sy1 & sh1 & "，发现了别人留下的200000两银子！"
sql="update 用户 set 体力=体力-10,银两=银两+200000 where 姓名='" & name1 & "'"
conn.execute sql
conn.close
case "16"
newer="" & name1 & "在孤岛上" & sy1 & sh1 & "，遇见了敌人，大战300回合以后，体力降低了200000！"
sql="update 用户 set 体力=体力-200000 where 姓名='" & name1 & "'"
conn.execute sql
conn.close
case "17"
newer="" & name1 & "在孤岛上" & sy1 & sh1 & "，发现了古代宝藏，正想去拿之际，宝藏的机关打开，你被乱箭射伤，损失体力500000！"
sql="update 用户 set 体力=体力-500000 where 姓名='" & name1 & "'"
conn.execute sql
conn.close
case "18"
newer="" & name1 & "在孤岛上" & sy1 & sh1 & "，发现了别人留下的500000两银子！"
sql="update 用户 set 体力=体力-10,银两=银两+500000 where 姓名='" & name1 & "'"
conn.execute sql
conn.close
case "19"
newer="" & name1 & "在孤岛上" & sy1 & sh1 & "，被脚下石头一拌摔了一跤，丢了20000两银子！"
sql="update 用户 set 体力=体力-10,银两=银两-20000 where 姓名='" & name1 & "'"
conn.execute sql
conn.close
case "20"
newer="" & name1 & "在孤岛上" & sy1 & sh1 & "，失足跌落悬崖，内力减少500000！"
sql="update 用户 set 体力=体力-10,内力=内力-500000 where 姓名='" & name1 & "'"
conn.execute sql
conn.close
case "21"
newer="" & name1 & "在孤岛上" & sy1 & sh1 & "，发现了别人留下的400000两银子！"
sql="update 用户 set 体力=体力-10,银两=银两+400000 where 姓名='" & name1 & "'"
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
newer="" & name1 & "在孤岛上" & sy1 & sh1 & "，发现了别人留下的100000两银子！"
sql="update 用户 set 体力=体力-10,银两=银两+100000 where 姓名='" & name1 & "'"
conn.execute sql
conn.close
case "24"
newer="" & name1 & "在孤岛上" & sy1 & sh1 & "，巧遇一世外高人，经指点，你的资质上升10000，精力上升2000！"
sql="update 用户 set 体力=体力-10,资质=资质+10000,精力=精力+2000 where 姓名='" & name1 & "'"
conn.execute sql
conn.close
case "25"
newer="" & name1 & "在孤岛上" & sy1 & sh1 & "，洗了洗脸！精力上升50000！"
sql="update 用户 set 精力=精力+50000 where 姓名='" & name1 & "'"
conn.execute sql
conn.close
case "26"
newer="" & name1 & "在孤岛上" & sy1 & sh1 & "，发现了别人留下的300000两银子！"
sql="update 用户 set 体力=体力-10,银两=银两+300000 where 姓名='" & name1 & "'"
conn.execute sql
conn.close
case "27"
if sex="男" then
newer="" & name1 & "在孤岛上" & sy1 & sh1 & "，巧遇一绝色女子，你把她救出孤岛，道德上升50000！"
else
newer="" & name1 & "在孤岛上" & sy1 & sh1 & "，巧遇美男子，道德上升50000！"
end if
sql="update 用户 set 体力=体力-1,道德=道德+50000 where 姓名='" & name1 & "'"
conn.execute sql
conn.close
case "28"
newer="" & name1 & "在孤岛上" & sy1 & sh1 & "，发现了别人留下的200000两银子！"
sql="update 用户 set 体力=体力-10,银两=银两+200000 where 姓名='" & name1 & "'"
conn.execute sql
conn.close
case "29"
newer="" & name1 & "在孤岛上" & sy1 & sh1 & "，发现了别人留下的500000两银子！"
sql="update 用户 set 体力=体力-10,银两=银两+500000 where 姓名='" & name1 & "'"
conn.execute sql
conn.close
case "30"
newer="" & name1 & "在孤岛上" & sy1 & sh1 & "，发现了一个大宝藏，经过一番冒险，你成功拿到300万宝藏及武功绝学(10000防御)！"
sql="update 用户 set 防御=防御+10000,银两=银两+3000000 where 姓名='" & name1 & "'"
conn.execute sql
conn.close
case "31"
newer="" & name1 & "在孤岛上" & sy1 & sh1 & "，发现了别人留下的50000两银子！"
sql="update 用户 set 体力=体力-10,银两=银两+50000 where 姓名='" & name1 & "'"
conn.execute sql
conn.close
case "32"
newer="" & name1 & "在孤岛上" & sy1 & sh1 & "，发现了电光之铠，防御增加5000！"
sql="update 用户 set 体力=体力-10,防御=防御+5000 where 姓名='" & name1 & "'"
conn.execute sql
conn.close
case "33"
newer="" & name1 & "在孤岛上" & sy1 & sh1 & "，发现了别人留下的200000两银子！"
sql="update 用户 set 体力=体力-10,银两=银两+200000 where 姓名='" & name1 & "'"
conn.execute sql
conn.close
case "34"
newer="" & name1 & "在孤岛上" & sy1 & sh1 & "，弄丢了100000两银子！"
sql="update 用户 set 体力=体力-10,银两=银两-100000 where 姓名='" & name1 & "'"
conn.execute sql
conn.close
case "35"
newer="" & name1 & "在孤岛上" & sy1 & sh1 & "，接受了古人留下的20000精力！"
sql="update 用户 set 体力=体力-10,精力=精力+20000 where 姓名='" & name1 & "'"
conn.execute sql
conn.close
case "36"
newer="" & name1 & "在孤岛上" & sy1 & sh1 & "，扭伤了脚。花400000两银子的医疗费！"
sql="update 用户 set 体力=体力-10,银两=银两-400000 where 姓名='" & name1 & "'"
conn.execute sql
conn.close
case "37"
newer="" & name1 & "在孤岛上" & sy1 & sh1 & "，发现了别人留下的50000两银子！"
sql="update 用户 set 体力=体力-10,银两=银两+50000 where 姓名='" & name1 & "'"
conn.execute sql
conn.close
case "38"
newer="" & name1 & "在孤岛上" & sy1 & sh1 & "，发现了别人留下的40000两银子！"
sql="update 用户 set 体力=体力-10,银两=银两+40000 where 姓名='" & name1 & "'"
conn.execute sql
conn.close
case "39"
newer="" & name1 & "在孤岛上" & sy1 & sh1 & "，发现了一奇特的山草药，一吃之下，身中剧毒，虽然花了1000000两银子治病，但还是留下了后遗症，降低了10000点资质"
sql="update 用户 set 资质=资质-10000,银两=银两-1000000 where 姓名='" & name1 & "'"
conn.execute sql
conn.close
case "40"
newer="" & name1 & "在孤岛上" & sy1 & sh1 & "，发现了别人留下的200000两银子！"
sql="update 用户 set 体力=体力-10,银两=银两+200000 where 姓名='" & name1 & "'"
conn.execute sql
conn.close
case "41"
newer="" & name1 & "在孤岛上" & sy1 & sh1 & "，发现了别人留下的王者之剑，攻击增加2000！"
sql="update 用户 set 体力=体力-10,攻击=攻击+2000 where 姓名='" & name1 & "'"
conn.execute sql
conn.close
case "42"
newer="" & name1 & "在孤岛上" & sy1 & sh1 & "，发现了别人留下的400000两银子！"
sql="update 用户 set 体力=体力-10,银两=银两+400000 where 姓名='" & name1 & "'"
conn.execute sql
conn.close
case "43"
newer="" & name1 & "在孤岛上" & sy1 & sh1 & "，洗了洗脸！体力上升300000"
sql="update 用户 set 体力=体力+300000 where 姓名='" & name1 & "'"
conn.execute sql
conn.close
case "44"
if sex="男" then
newer="" & name1 & "在孤岛上" & sy1 & sh1 & "，巧遇一绝色女子，你色心大起，道德下降20000！"
else
newer="" & name1 & "在孤岛上" & sy1 & sh1 & "，巧遇一美男子，你想勾引他，道德下降20000！"
end if
sql="update 用户 set 体力=体力-10,道德=道德-20000 where 姓名='" & name1 & "'"
conn.execute sql
conn.close
case "45"
newer="" & name1 & "在孤岛上" & sy1 & sh1 & "，丢失了200000两银子！"
sql="update 用户 set 体力=体力-10,银两=银两-200000 where 姓名='" & name1 & "'"
conn.execute sql
conn.close
case "46"
newer="" & name1 & "在孤岛上" & sy1 & sh1 & "，发现了别人留下的300000两银子！"
sql="update 用户 set 体力=体力-10,银两=银两+300000 where 姓名='" & name1 & "'"
conn.execute sql
conn.close
case "47"
newer="" & name1 & "在孤岛上" & sy1 & sh1 & "，碰到毒蛇，幸亏武功高强，虚惊一场！"
sql="update 用户 set 体力=体力-10 where 姓名='" & name1 & "'"
conn.execute sql
conn.close
case "48"
newer="" & name1 & "在孤岛上" & sy1 & sh1 & "，发现了别人留下的攻击，加5000点！"
sql="update 用户 set 体力=体力-10,攻击=攻击+5000 where 姓名='" & name1 & "'"
conn.execute sql
conn.close
case "49"
newer="" & name1 & "在孤岛上" & sy1 & sh1 & "，发现了别人留下的250000两银子！"
sql="update 用户 set 体力=体力-10,银两=银两+250000 where 姓名='" & name1 & "'"
conn.execute sql
conn.close
case "50"
if sex="男" then
newer="" & name1 & "在孤岛上" & sy1 & sh1 & "，巧遇一绝色女子，上前调戏，道德下降5000！"
else
newer="" & name1 & "在孤岛上" & sy1 & sh1 & "，巧遇一美男子，上前勾引，可是人家不理你，道德下降5000！"
end if
sql="update 用户 set 体力=体力-1000,道德=道德-5000 where 姓名='" & name1 & "'"
conn.execute sql
conn.close
case "51"
newer="" & name1 & "在孤岛上" & sy1 & sh1 & "，发现了别人留下的200000两银子！"
sql="update 用户 set 体力=体力-10,银两=银两+200000 where 姓名='" & name1 & "'"
conn.execute sql
conn.close
case "52"
newer="" & name1 & "在孤岛上" & sy1 & sh1 & "，发现了别人留下的100000两银子！"
sql="update 用户 set 体力=体力-10,银两=银两+100000 where 姓名='" & name1 & "'"
conn.execute sql
conn.close
case "53"
newer="" & name1 & "在孤岛上" & sy1 & sh1 & "，突然在墙角发现180000两银子！"
sql="update 用户 set 体力=体力-10,银两=银两+180000 where 姓名='" & name1 & "'"
conn.execute sql
conn.close
case "54"
newer="" & name1 & "在孤岛上" & sy1 & sh1 & "，发现了一种奇特的山草药，一吃之下，内力增进20000，精力提高10000！"
sql="update 用户 set 体力=体力-10,内力=内力+20000,精力=精力+10000 where 姓名='" & name1 & "'"
conn.execute sql
conn.close
case "55"
newer="" & name1 & "在孤岛上" & sy1 & sh1 & "，发现了别人留下的15000防御！"
sql="update 用户 set 体力=体力-10,防御=防御+15000 where 姓名='" & name1 & "'"
conn.execute sql
conn.close
case "56"
newer="" & name1 & "在孤岛上" & sy1 & sh1 & "，发现了别人留下的120000两银子！"
sql="update 用户 set 体力=体力-10,银两=银两+120000 where 姓名='" & name1 & "'"
conn.execute sql
conn.close
case "57"
newer="" & name1 & "在孤岛上" & sy1 & sh1 & "，不小心弄坏了别人的房子，赔了500000两银子！"
sql="update 用户 set 体力=体力-10,银两=银两-500000 where 姓名='" & name1 & "'"
conn.execute sql
conn.close
case "58"
newer="" & name1 & "在孤岛上" & sy1 & sh1 & "，发现了别人留下的300000两银子！"
sql="update 用户 set 体力=体力-10,银两=银两+300000 where 姓名='" & name1 & "'"
conn.execute sql
conn.close
case "59"
newer="" & name1 & "在孤岛上" & sy1 & sh1 & "，发现了别人留下的200000两银子！"
sql="update 用户 set 体力=体力-10,银两=银两+200000 where 姓名='" & name1 & "'"
conn.execute sql
conn.close
case "60"
newer="" & name1 & "在孤岛上" & sy1 & sh1 & "，发现了别人留下的100000两银子！"
sql="update 用户 set 体力=体力-10,银两=银两+100000 where 姓名='" & name1 & "'"
conn.execute sql
conn.close
case "61"
newer="" & name1 & "在孤岛上" & sy1 & sh1 & "，洗了洗脸！体力上升100000"
sql="update 用户 set 体力=体力+100000 where 姓名='" & name1 & "'"
conn.execute sql
conn.close
case "62"
newer="" & name1 & "在孤岛上" & sy1 & sh1 & "，发现了别人留下的30000两银子！"
sql="update 用户 set 体力=体力-10,银两=银两+30000 where 姓名='" & name1 & "'"
conn.execute sql
conn.close
case "63"
newer="" & name1 & "在孤岛上" & sy1 & sh1 & "，发现了别人留下的100000两银子！"
sql="update 用户 set 体力=体力-10,银两=银两+10000 where 姓名='" & name1 & "'"
conn.execute sql
conn.close
case "64"
newer="" & name1 & "在孤岛上" & sy1 & sh1 & "，失足跌下山坡！大难不死，内力损失200000,但大难不死，必有后福，发现50000两银子"
sql="update 用户 set 内力=内力-200000,银两=银两+500000 where 姓名='" & name1 & "'"
conn.execute sql
conn.close
case "65"
newer="" & name1 & "在孤岛上" & sy1 & sh1 & "，发现了别人留下的100000两银子！"
sql="update 用户 set 体力=体力-10,银两=银两+100000 where 姓名='" & name1 & "'"
conn.execute sql
conn.close
case "66"
newer="" & name1 & "在孤岛上" & sy1 & sh1 & "，巧遇一世外高人，经指点，你精力上升8000！"
sql="update 用户 set 体力=体力-10,精力=精力+8000 where 姓名='" & name1 & "'"
conn.execute sql
conn.close
case "67"
newer="" & name1 & "在孤岛上" & sy1 & sh1 & "，发现了别人留下的30000两银子！"
sql="update 用户 set 体力=体力-10,银两=银两+30000 where 姓名='" & name1 & "'"
conn.execute sql
conn.close
case "68"
newer="" & name1 & "在孤岛上" & sy1 & sh1 & "，丢失了500000体力！"
sql="update 用户 set 体力=体力-500000,银两=银两+100 where 姓名='" & name1 & "'"
conn.execute sql
conn.close
case "69"
if sex="男" then
newer="" & name1 & "在孤岛上" & sy1 & sh1 & "，路遇绝色女子遇难，英雄救美，道德上升1000！"
else
newer="" & name1 & "在孤岛上" & sy1 & sh1 & "，路遇一美男遇难，舍身相救，道德上升1000！"
end if
sql="update 用户 set 体力=体力-1000,道德=道德+1000 where 姓名='" & name1 & "'"
conn.execute sql
conn.close
case "70"
newer="" & name1 & "在孤岛上" & sy1 & sh1 & "，发现了别人留下的80000两银子！"
sql="update 用户 set 体力=体力-10,银两=银两+80000 where 姓名='" & name1 & "'"
conn.execute sql
conn.close
case "71"
newer="" & name1 & "在孤岛上" & sy1 & sh1 & "，发现了别人留下的150000两银子！"
sql="update 用户 set 体力=体力-10,银两=银两+150000 where 姓名='" & name1 & "'"
conn.execute sql
conn.close
case "72"
newer="" & name1 & "在孤岛上" & sy1 & sh1 & "，钱太多了，扔了自己的250000两银子！"
sql="update 用户 set 体力=体力-10,银两=银两-250000 where 姓名='" & name1 & "'"
conn.execute sql
conn.close
case "73"
newer="" & name1 & "在孤岛上" & sy1 & sh1 & "，发现了上乘的内力心法，内力大进100000点，体力也增加了100000点！"
sql="update 用户 set 内力=内力+100000,体力=体力+100000 where 姓名='" & name1 & "'"
conn.execute sql
conn.close
case "74"
if sex="男" then
newer="" & name1 & "在孤岛上" & sy1 & sh1 & "，巧遇一绝色女子，你色心大起，对其施暴，道德下降5000！"
else
newer="" & name1 & "在孤岛上" & sy1 & sh1 & "，巧遇一美男子，你春心动，上前搭讪，道德下降5000！"
end if
sql="update 用户 set 体力=体力-10,道德=道德-5000 where 姓名='" & name1 & "'"
conn.execute sql
conn.close
case "75"
newer="" & name1 & "在孤岛上" & sy1 & sh1 & "，发现了别人留下的10000两银子！"
sql="update 用户 set 体力=体力-10,银两=银两+10000 where 姓名='" & name1 & "'"
conn.execute sql
conn.close
case "76"
newer="" & name1 & "在孤岛上" & sy1 & sh1 & "，发现了别人留下的20000两银子！"
sql="update 用户 set 体力=体力-10,银两=银两+20000 where 姓名='" & name1 & "'"
conn.execute sql
conn.close
case "77"
newer="" & name1 & "在孤岛上" & sy1 & sh1 & "，发现了别人留下的80000两银子！"
sql="update 用户 set 体力=体力-10,银两=银两+80000 where 姓名='" & name1 & "'"
conn.execute sql
conn.close
case "78"
newer="" & name1 & "在孤岛上" & sy1 & sh1 & "，发现了别人留下的60000两银子！"
sql="update 用户 set 体力=体力-10,银两=银两+60000 where 姓名='" & name1 & "'"
conn.execute sql
conn.close
case "79"
if sex="男" then
newer="" & name1 & "在孤岛上" & sy1 & sh1 & "，夜里和一绝色女子共宿，当了一夜君子，道德上升3000！"
else
newer="" & name1 & "在孤岛上" & sy1 & sh1 & "，单枪挑了曹操老賊，不愧为巾帼英雄，道德上升3000！"
end if
rs.close
sql="update 用户 set 体力=体力-1000,道德=道德+3000 where 姓名='" & name1 & "'"
conn.execute sql
conn.close

case "80"
sql="SELECT top 1 * FROM 宝物 WHERE 拿走='0' order by id desc"
Set Rs=connt.Execute(sql)
If Rs.Bof OR Rs.Eof Then
newer="" & name1 & "在孤岛上" & sy1 & sh1 & "，意外地发现了先人在孤岛留下的宝物藏身地，但是里面什么都没有！"
sql="update 宝物 set 拿走=0"
set Rs=connt.Execute(sql)
connt.close
else
bao=rs("宝物名")
at=rs("攻击力")
fy=rs("防御力")
newer="" & name1 & "在孤岛上" & sy1 & sh1 & "，意外地发现了先人在孤岛留下的宝物藏身地，得到了稀世奇珍" & bao & "！！"
sql="SELECT * FROM 物品 WHERE 所有者='" & name1 & "' and 名称='"& bao &"'"
Set Rs=conn.Execute(sql)
if rs.eof or rs.bof then
sql="insert into 物品(名称,属性,所有者,攻击,防御,价格,数量,特效,寄售,装备)values('"& bao &"','武器','"&name1&"'," & at & "," & fy & ",0,1,'无',False,False)"
conn.execute sql
else
sql="update 物品 set 数量=数量+1 where 名称='" & bao & "' and 所有者='" & name1 & "'"
conn.execute sql
end if
sql="update 宝物 set 拿走=1,获宝者='"&name1&"', 时间=now() where 宝物名='" & bao & "'"
set Rs=connt.Execute(sql)
connt.close
conn.close
end if

case "81"
newer="" & name1 & "在孤岛上" & sy1 & sh1 & "，发现了别人留下的160000两银子！"
sql="update 用户 set 体力=体力-10,银两=银两+160000 where 姓名='" & name1 & "'"
conn.execute sql
conn.close

case "82"
newer="" & name1 & "在孤岛上" & sy1 & sh1 & "，发现了别人留下的火焰之剑，装备上，居然把自己砍伤，幸亏有高人帮助，防御增加了2000！"
sql="update 用户 set 体力=体力-10,防御=防御+2000 where 姓名='" & name1 & "'"
conn.execute sql
conn.close
case "1"
newer="" & name1 & "在孤岛上" & sy1 & sh1 & "，发现了一个藏宝库，但没钥匙，打不开门！"
sql="update 用户 set 体力=体力-10 where 姓名='" & name1 & "'"
conn.execute sql
conn.close
End Select
end if
end if
huayuan=newer
end function
%>
