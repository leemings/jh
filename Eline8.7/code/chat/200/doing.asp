
<%
Response.buffer=true
reno=trim(cstr(application("reno")))      
charp=request("charp")  
ddwho=trim(request("ddwho"))
    if  ddwho="1" or ddwho="2" or ddwho="3" or ddwho="4"   then  
    application("dddwho")=ddwho
    end if
    if  charp="stopp" or  charp="addmoney" or charp="nomoney" or charp="myhouse" or   charp="pmoney"   then application("charp")=charp	

if charp="stopp" and ddwho<>"" then session("stopp")=""
if charp="nomoney" and ddwho<>""  then session("nomoney")=""
if charp="addmoney" and ddwho<>""  then session("addmoney")=""
if charp="pmoney" and ddwho<>""  then session("pmoney")=""
if charp="myhouse"  and ddwho<>"" then session("myhouse")=""



 if reno="1" then bname="小丹尼"
 if reno="2" then bname="宫本宝藏"
 if reno="3" then bname="钱夫人"
 if reno="4" then bname="沙隆巴斯"
  if application("dddwho")="1" then toname="小丹尼"
  if application("dddwho")="2" then toname="宫本宝藏"
  if application("dddwho") ="3" then toname="钱夫人"
  if application("dddwho")  ="4"then  toname="沙隆巴斯"
 
 if   session("players")=1 then sname="小丹尼"
 if   session("players")=2 then sname="宫本宝藏"
 if  session("players")=3  then sname="钱夫人"
 if   session("players")=4 then sname="沙隆巴斯"




saymsg=left(trim(request("saymsg")),23)
if saymsg<> "" then 
saymsg=trim(saymsg)
saymsg="<font color=#339933>"&sname&"说<b>:</b>"&saymsg&" </font>"
application.lock
  for i=10 to 2 step-1
   j=i-1
    application("msgg"&i)=application("msgg"&j)
next
application("msgg1")=saymsg
application.unlock
end if
'11
if application("reno")=session("players") then 
session("waityou")=1

 if  trim(application("dddwho"))=trim(cstr(session("players"))) then 
  charp=trim(application("charp"))
if charp="stopp" then session("zjstopp")=1  
if charp="nomoney" then session("zjnomoney")=1  
application("dddwho")=""
application("charp")=""
end if



if application("charp")<>""    and charp<>""  then   '注意条件charp
if application("charp")="stopp" then charpname="停留卡"
if application("charp")="nomoney" then charpname="免税卡"
if application("charp")="addmoney" then charpname="查税卡"
if application("charp")="myhouse" then charpname="购地卡"
if application("charp")="pmoney" then charpname="均富卡"

IF application("charp")<>"myhouse"   then 

application.lock
  for i=10 to 2 step-1
   j=i-1
    application("msgg"&i)=application("msgg"&j)
next
if bname=toname then toname="自己"
application("msgg1")=" <font color=red>"&bname&"对"&toname&"使用"&charpname&"</font>"

if application("charp")="addmoney" then 
application("msgg1")=" <font color=red>"&bname&"查"&toname&"的税</font>交税"&fix((application("money"&ddwho)*0.4))&"元给自己,人无横财不富哈~~哈~~~哈"
end if

if application("charp")="pmoney" then 
application("msgg1")=" <font color=red>"&bname&"对"&toname&"使用"&charpname&"</font>,均分所有现金"
end if

application.unlock

session("stopp")=""
end if
	END IF

myhome=1  '购地的房子号
dj=1   ' '购地的房子等级初始值
IF application("dddwho")<>"" AND application("charp")="myhouse" THEN
for i=1 to 40 
     if    mid(trim(application("house"&I)),1,1)=trim(application("dddwho")) then 
	if   mid(trim(application("house"&i)),2,1)>dj then
	     dj=mid(trim(application("house"&i)),2,1)
	     myhome=i
  	   end if
	end if
next 

myhomez=""&trim(reno)&""&trim(dj)&""
application("house"&myhome)=trim(myhomez)

application.lock
  for i=10 to 2 step-1
   j=i-1
    application("msgg"&i)=application("msgg"&j)
next
application("msgg1")=" <font color=red>"&bname&"对"&toname&"使用"&charpname&"</font>,获得豪宅"&myhome&"~"&dj&"级哈~~哈~~~哈"
application.unlock
application("dddwho")=""
application("charp")=""
session("myhouse")=""
END IF

 if    application("charp")="pmoney" then 
ddwho=trim(application("dddwho"))
allmoney=application("money"&reno)+application("money"&ddwho)
allmoney=fix(allmoney/2)
application("money"&reno)=allmoney
application("money"&ddwho)=allmoney
'response.write application("money"&reno)
'response.write "<br>"
'response.write application("money"&ddwho)

'response.end
application("charp")=""
end if




if application("charp")="addmoney"    then 
ddwho=trim(application("dddwho"))
application("money"&reno)=application("money"&reno)+fix((application("money"&ddwho)*0.4))
application("money"&ddwho)=application("money"&ddwho)-fix((application("money"&ddwho)*0.4))
application.unlock
application("dddwho")=""
application("charp")=""
end if 
if session("zjstopp")<>""   then  
	 application.lock
	   for i=10 to 2 step-1
	   j=i-1
	   application("msgg"&i)=application("msgg"&j)
	   next
application("msgg1")="停留卡生效,"&bname&"原地停留"&session("zjstopp")&"回合"
application.unlock
application("reno")=application("reno")+1
application("dddwho")=""
application("charp")=""
session("zjstopp")=session("zjstopp")+1
if session("zjstopp")>3 then session("zjstopp")=""     '  停留共五回合
response.redirect "nan.asp"

END IF  


randomize
weiz=rnd
weiz=midb(weiz,5,1)  '产生随机数
if weiz>6 then weiz=6 '设置计算
if weiz=0 then weiz=1 '设置计算
session("weiz")=weiz

if session("players")=1  then 
application("play1")=clng(application("play1"))
application("play1")=application("play1")+weiz '位置累加
if  application("play1")>=40 then   application("play1")=application("play1")-39 
'       application("play1")=17
       application("play1")=cstr(application("play1"))

end if

if session("players")=2 then 
	application("play2")=clng(application("play2"))
                  application("play2")=application("play2")+weiz '每个玩家的不同的位置

                  if  application("play2")>=40 then application("play2")=application("play2")-39 '如果最多只是40步时，可以根据地图的大小设置
'       application("play2")=7

'22
	application("play2")=cstr(application("play2"))
                 end if

if session("players")=3 then 
   application("play3")=clng(application("play3"))
    application("play3")=application("play3")+weiz '每个玩家的不同的位置
    if  application("play3")>40 then  application("play3")= application("play3")-39 '如果最多只是40步时，可以根据地图的大小设置
'application("play3")=17

   application("play3")=cstr(application("play3"))
end if


if session("players")=4 then 
   application("play4")=clng(application("play4"))
    application("play4")=application("play4")+weiz '每个玩家的不同的位置
    if  application("play4")>40 then  application("play4")= application("play4")-39 '如果最多只是40步时，可以根据地图的大小设置

   application("play4")=cstr(application("play4"))
end if


if application("play"&reno)="9" or application("play"&reno)="28" then 
randomize
nohouse=rnd
nohouse=midb(nohouse,5,1)  '产生随机数
response.write nohouse
response.write "<br>"
if nohouse >5 then 
nohouse=2
else
nohouse=1
end if
response.write nohouse
'response.end
if nohouse=1 then  
 	if session("stopp")="" then session("stopp")="stopp"
 	if session("addmoney")="" then session("addmoney")="addmoney"
 	if session("nomoney")="" then session("nomoney")="nomoney"
 	if session("myhouse")="" then session("myhouse")="myhouse"
application.lock
  for i=10 to 2 step-1
   j=i-1
    application("msgg"&i)=application("msgg"&j)
next
application("msgg1")="突发事件位置"&application("play"&reno)&" <font color=red>"&bname&"获得四张卡片</font>"
application.unlock
end if

if nohouse=2 then  
dj=1
for i=1 to 40 
     if    mid(trim(application("house"&I)),1,1)=trim(cstr(session("players"))) then 
	if   mid(trim(application("house"&i)),2,1)>dj then
	     dj=mid(trim(application("house"&i)),2,1)
	     myhome=i  
  	   end if
	end if
next 

'myhomez=""&trim(reno)&""&trim(dj)&""  '房子的值,抄前面的程序
application("house"&myhome)=""

application.lock
  for i=10 to 2 step-1
   j=i-1
    application("msgg"&i)=application("msgg"&j)
next
application("msgg1")="突发事件位置"&application("play"&reno)&" <font color=red>没收"&bname&"的房子"&myhome&"~"&dj&"级</font>"
application.unlock






end if
end if






    reno=trim(cstr(application("reno")))       '现在的玩家,因为现在是reno决定是哪一个玩家
if  application("play"&reno) ="4"  or  application("play"&reno) ="9" or   application("play"&reno) ="28"    or   application("play"&reno) ="14"  or application("play"&reno) ="20"  or application("play"&reno) =25  or application("play"&reno) ="27"   or application("play"&reno) ="34"  or application("play"&reno) ="38" then 
 application("reno")=application("reno")+1
response.redirect "nan.asp"
end if 

                   strno=trim(cstr(application("play"&reno)))   '第N个房子字符串strno 表示第N个玩家现在所在的位置=房子编号


                         if application("house"&strno)=""  then
                          application("house"&strno)=cstr(session("players")) '你是第一个进来的人
application.lock
   for i=10 to 2 step-1
    j=i-1
    application("msgg"&i)=application("msgg"&j)
    next
   application("msgg1")=" "&bname&"<font color='red'>购买</font>空地"&strno&"成功花费400元</font>" 
   application.unlock
	       'application("house"&strno)=" "&application("house"&strno)&"1  "
	       application("money"&reno)= application("money"&reno)-400
                   	     END IF

whohouse=mid(trim(application("house"&strno)),1,1)
 if whohouse="1" then toname="小丹尼"
 if whohouse="2" then toname="宫本宝藏"
 if whohouse="3" then toname="钱夫人"
 if whohouse="4" then toname="沙隆巴斯"


houseTop=mid(trim(application("house"&strno)),2,1)
         if  trim(whohouse)=trim(cstr(session("players")))     then 
IF application("house"&strno)<>"" THEN
application.lock
  for i=10 to 2 step-1
   j=i-1
    application("msgg"&i)=application("msgg"&j)
next

	   Select Case  houseTop
	       case ""       
	       application("house"&strno)=" "&application("house"&strno)&"1  "
	      case  "1"       '为1级时,升为2级
application("msgg1")=" "&bname&"将房间"&strno&"<font color='red'>升为~2级</font>花费800元</font>" 

	       application("money"&reno)= application("money"&reno)-800
	       application("house"&strno)=left(trim(application("house"&strno)),1)
                       application("house"&strno)=" "&application("house"&strno)&"2  "
	     case  "2"
application("msgg1")=" "&bname&"将房间"&strno&"<font color='red'>升为~3级</font>花费1600元</font>" 
                     application("money"&reno)= application("money"&reno)-1600
	       application("house"&strno)=left(trim(application("house"&strno)),1)
                       application("house"&strno)=" "&application("house"&strno)&"3  "
	     case  "3"     '升为4级
application("msgg1")=" "&bname&"将房间"&strno&"<font color='red'>升为~4级</font>花费3200元</font>" 
	       application("money"&reno)= application("money"&reno)-3200
	       application("house"&strno)=left(trim(application("house"&strno)),1)
                       application("house"&strno)=" "&application("house"&strno)&"4  "
               End Select 
        end if	   
application.unlock
END IF
IF application("house"&strno)<>"" THEN
for st = 1 to 4
         if  trim(whohouse)<>trim(cstr(session("players")))     then    '要付钱了

 	   ns=cstr(st)
                   	if whohouse= ns  then 	'钱付给ns号玩家
   application.lock
  for i=10 to 2 step-1
   j=i-1
    application("msgg"&i)=application("msgg"&j)
next
	IF    session("zjnomoney")  ="" THEN 
	Select Case  houseTop
	       case "1"      
	  application("money1")=application("money"&ns)+800
	  application("money"&reno)=application("money"&reno)-800
                   	application("houseno")=strno
	application("msgg1")=" "&bname&"在房间"&trim(application("houseno"))&"~1级付800元<font color='red'>给"&toname&"</font>" 

	       case "2"      
	  application("money1")=application("money"&ns)+1600
	  application("money"&reno)=application("money"&reno)-1600
                   	application("houseno")=strno
	application("msgg1")=" "&bname&"在房间"&trim(application("houseno"))&"~2级付1600元<font color='red'>给"&toname&"</font>" 

	       case "3"      
	  application("money1")=application("money"&ns)+3200
	  application("money"&reno)=application("money"&reno)-3200
                   	application("houseno")=strno
	application("msgg1")=" "&bname&"在房间"&trim(application("houseno"))&"~3级付3200元<font color='red'>给"&toname&"</font>" 

	       case "4"      
	  application("money1")=application("money"&ns)+6400
	  application("money"&reno)=application("money"&reno)-6400
                   	application("houseno")=strno
	application("msgg1")=" "&bname&"在房间"&trim(application("houseno"))&"~4级付6400元<font color='red'>给"&toname&"</font>" 
end select
	END IF
'session("zjnomoney")=4
if  session("zjnomoney")<>"" then
	application("msgg1")=" "&bname&"在房间"&trim(application("play"&reno))&"<font color='red'>免税卡生效"&session("zjnomoney")&"回</font>,不用付钱给"&toname&"" 
	   session("zjnomoney")=session("zjnomoney")+1
	if session("zjnomoney")>4 then session("zjnomoney")=""
end if
end if
application.unlock
END IF

NEXT	
	END IF
application("reno")=application("reno")+1
end if
response.redirect "nan.asp"

%>