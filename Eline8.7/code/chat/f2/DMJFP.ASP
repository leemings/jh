<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="../sjfunc/sjfunc.asp"-->
<!--#include file="../sjfunc/func.asp"-->
<%'搓麻将♀wWw.happyjh.com♀
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
nowinroom=session("nowinroom")
if Instr(LCase(Application("sjjh_useronlinename"&nowinroom)),LCase(" "&sjjh_name&" "))=0 then 
	Session("sjjh_inthechat")="0" 
	Response.Redirect "../chaterr.asp?id=001" 
end if 
towho=Trim(Request.Form("towho"))
towhoway=Request.Form("towhoway")
saycolor=Request.Form("saycolor")
addwordcolor=Request.Form("addwordcolor")
addsays=Request.Form("addsays")
says=Trim(Request.Form("sy"))
action=Request("action")
'对暂离开、点哑穴判断
call dianzan(towho)
if len(says)>150 then says=Left(says,150)
says=lcase(says)
says=Replace(says,"&amp;","&")
'对系统禁止字符处理
if sjjh_grade<9 then
  says=bdsays(says)
end if
says=Ucase(says)
act=1
towhoway=0
i=instr(says,"$")
fnn1=trim(mid(says,i+1))
select case action
	case 1
		says=dmjfp(fnn1,sjjh_name,towho)
		says="<font color=green><b>【发牌】</b><font color=" & saycolor & ">"&says&"</font>"
	case 2
		says=dmjGetmj(sjjh_name,towho)
		says="<font color=green><b>【摸牌】</b><font color=" & saycolor & ">"&says&"</font>"
	case 3
		says=dmjPeng("问",sjjh_name,towho)
		says="<font color=green><b>【问牌】</b><font color=" & saycolor & ">"&says&"</font>"
	case 4
		says=dmjPeng(fnn1,sjjh_name,towho)
		says="<font color=green><b>【吃牌】</b><font color=" & saycolor & ">"&says&"</font>"
	case 5
		says=dmjPeng(fnn1,sjjh_name,towho)
		says="<font color=green><b>【碰牌】</b><font color=" & saycolor & ">"&says&"</font>"
	case 6
		says=dmjPeng(says,sjjh_name,towho)
		says="<font color=green><b>【杠牌】</b><font color=" & saycolor & ">"&says&"</font>"
	case else
		call ErrALT("不明命令")
end select

call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)


'搓麻将
function dmjfp(fn1,from1,to1)
if to1="大家" and fn1<>"认输了" then call ErrALT("请选择说话对象，不能象大家发送这个命令!")
sjjh_name=Session("sjjh_name")
if sjjh_name="" then Response.Redirect "../../error.asp?id=440"

nickname=Session("sjjh_name")
grade=Session("sjjh_grade")
chatroomsn=session("nowinroom")

if left(fn1,2)="公证" then
	Errstr="错误的命令格式：\n正确示例：\n/发牌 公证|负\n/发牌 公证|胜"
	arr_fn1=split(fn1,"|")
	if ubound(arr_fn1)<>1 then call ErrALT(Errstr)
	fn_62=left(arr_fn1(1),1)

	if fn_62<>"负" and fn_62<>"胜" then call ErrALT(Errstr)
	if grade<6 then call ErrALT("你不是官府管理员，不能进行牌局公证！" )
	from1=to1
end if

f=Minute(time()) 

Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
db="edmj.asp"
'connstr="driver={Microsoft Access Driver (*.mdb)};dbq=" & Server.MapPath(""&db&"")
'如果你的服务器支持jet4.0，请使用下载的连接方法，提高程序性能
connstr="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(""&db&"")
conn.open connstr 

sql="select * from dmj where ufrom='"& from1 & "'"
Set Rs=conn.Execute(sql)
if rs.eof or rs.bof then
	rs.close
	conn.close
	set rs=nothing
	set conn=nothing
	call ErrALT(from1 & "没有参与[麻将]牌局！")
end if
isMy=rs("isMy")
isFp=rs("isFp")
isGet=rs("isGet")
Mymj=rs("Mymj")
dmjto=rs("uto")
xiazhu=rs("duz")
logtime=rs("logtime")
mjID=rs("mjID")
rs.close

'------------------------新格式------------------------
dim xia_class,xia_fir,xia_sql
xia_fir=left(xiazhu,1)
xiazhu=mid(xiazhu,2)

select case xia_fir
	case "1"
		xia_class="金币"
		xia_sql="金币"
	case "2"
		xia_class="银两"
		xia_sql="银两"
	case "3"
		xia_class="豆点"
		xia_sql="泡豆点数"
	case else
		call ErrALT("非法操作！")
end select
'------------------------新格式------------------------
d_sql="delete from dmj where ufrom='" & from1 & "' or uto='" & from1 &"'"
d_sql2="delete from mjInfo where id=" & mjID

if fn_62<>"" then
	if fn_62="负" then
		s_pk=dmjto
		f_pk=to1
	elseif fn_62="胜" then
		f_pk=dmjto
		s_pk=to1
	End if

	Set Rs=conn.Execute(d_sql)
	Set Rs=conn.Execute(d_sql2)
	conn.close
	connstr=Application("sjjh_usermdb")
	Set conn=Server.CreateObject("ADODB.CONNECTION")
	conn.open connstr

	'------------------------新格式------------------------
	conn.execute "update 用户 set " & xia_sql & "=" & xia_sql & "+"&xiazhu&"*3 where 姓名='"& s_pk &"'"
	conn.execute "update 用户 set " & xia_sql & "=" & xia_sql & "+"&xiazhu&" where 姓名='"& f_pk &"'"
	conn.close
	dmjfp="【官府公证】" & ErrSays(f_pk,"不遵守[麻将牌局]规定拒绝完成牌局不肯认输<br> ≮≮≮≮ 官府人员[ " & nickname & " ]现裁定他输给[ " & s_pk & " ]" & xiazhu & xia_class & sayxia & " ≯≯≯≯ ")
elseif fn1<>"认输了" then
		if to1<>dmjto then
			dmjfp= ErrSays(nickname,"你的对手不是[" & to1 & "]啊？")
			conn.close
			exit function
		end if
		if not isMy then
			dmjfp= ErrSays(nickname,"现在不是你出牌，请等待对方出牌！")
			conn.close
			set rs=nothing
			set conn=nothing
			exit function
		end if
		if isGet and (not isFp) then
			dmjfp= ErrSays(nickname,"你应当先摸牌，然后再打牌,在[洗牌]桌点击[摸 牌]使用(/摸牌)命令")
			%>
			<script language="JavaScript">parent.f2.document.forms[0].sytemp.value="/摸牌";</script>
			<%
			conn.close
			set rs=nothing
			set conn=nothing
			exit function
		end if
		if instr(Mymj,fn1 & "|")<>0 then
			Mymj=replace(Mymj,fn1 & "|","",1,1,1)
		else
			dmjfp= ErrSays(nickname,"你有没有搞错啊，你没有这张牌啊？")
			conn.close
			set rs=nothing
			set conn=nothing
			exit function
		end if
		sql="update dmj set Mymj='" & Mymj & "',oldmj='" & fn1 & "',isGet=true,isFp=false,isMy=false,logtime='" & now() & "' where ufrom='"& nickname & "'"
		Set Rs=conn.Execute(sql)
		sql="update dmj set isGet=true,isFp=false,isMy=true where ufrom='"& dmjto & "'"
		Set Rs=conn.Execute(sql)
		conn.close
		dmjfp= ErrSays(nickname,"拿出一张牌用手指头搓了搓,嘿......掷上了牌桌，原来是 " & showMj(fn1))%>
		<script language="JavaScript">parent.f3.location.reload();parent.f2.document.forms[0].sytemp.value="/摸牌";</script>
		<%
elseif fn1="认输了" then
		Set Rs=conn.Execute(d_sql)
		Set Rs=conn.Execute(d_sql2)
		conn.close
		connstr=Application("sjjh_usermdb")
		Set conn=Server.CreateObject("ADODB.CONNECTION")
		conn.open connstr 

'------------------------新格式------------------------
		conn.execute "update 用户 set " & xia_sql & "=" & xia_sql & "+"&xiazhu&"*3 where 姓名='"& dmjto &"'"
		conn.execute "update 用户 set " & xia_sql & "=" & xia_sql & "+"& xiazhu&" where 姓名='"& nickname &"'"
		conn.close
		dmjfp= ErrSays(nickname,"认输不打了，输给[" & dmjto & "]" & xiazhu  & xia_class & sayxia )
'------------------------新格式------------------------
end if
set conn=nothing
set rs=nothing
end function
'/////////////////////////////////////////////////抓牌
function dmjGetmj(from1,to1)
if to1="大家" and fn1<>"认输了" then call ErrALT("请选择说话对象，不能象大家发送这个命令!")
sjjh_name=Session("sjjh_name")
if sjjh_name="" then Response.Redirect "../../error.asp?id=440"

nickname=Session("sjjh_name")
grade=Session("sjjh_grade")
chatroomsn=session("nowinroom")

Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
db="edmj.asp"
'connstr="driver={Microsoft Access Driver (*.mdb)};dbq=" & Server.MapPath(""&db&"")
'如果你的服务器支持jet4.0，请使用下载的连接方法，提高程序性能
connstr="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(""&db&"")
conn.open connstr 

sql="select * from dmj where ufrom='"& nickname & "'"
Set Rs=conn.Execute(sql)

if rs.eof and rs.bof then
	rs.close
	conn.close
	set rs=nothing
	set conn=nothing
	call ErrAlt("您没有参加打麻将，怎么就想到摸牌了?")
else
	isMy=rs("isMy")
	isFp=rs("isFp")
	isGet=rs("isGet")
	Mymj=rs("Mymj")
	dmjto=rs("uto")
	xiazhu=rs("duz")
	mjID=rs("mjID")
	logtime=rs("logtime")
	rs.close

if not isMy then
	dmjGetmj= ErrSays(nickname,"现在好象不是你摸牌啊,等对方出牌了才是你摸牌？")
	conn.close
	set rs=nothing
	set conn=nothing
	exit function
end if
if isFp and (not isGet) then
	dmjGetmj= ErrSays(nickname,"你现在应当打牌(出牌),在[洗牌]桌点击麻将使用(/打麻将 XXX) 命令")
	conn.close
	set rs=nothing
	set conn=nothing
	exit function
end if
sql="select 麻将 from mjInfo where id="& mjID 
Set Rs=conn.Execute(sql)
Getmjs=rs("麻将")
rs.close

if len(Getmjs)>39 then
Getmjs2=right(Getmjs,3)
Mymj=Mymj & Getmjs2
Getmjs=left(Getmjs,len(Getmjs)-3)
sql="update mjInfo set 麻将='" & Getmjs & "' where id="& mjID 
Set Rs=conn.Execute(sql)
sql="update dmj set Mymj='" & Mymj & "',isGet=false,isFp=true,logtime='" & now() & "' where ufrom='"& nickname & "'"
Set Rs=conn.Execute(sql)
%>
<script language="JavaScript">parent.f3.location.href="dmj-xp.asp";alert("你摸到<%=strmj2(left(Getmjs2,2))%>")</script>
<%
dmjGetmj= ErrSays(nickname,"在牌桌摸了一张牌，小心的看了看，瞪了瞪眼，不知在想什么......" )
else
sql="delete from dmj where ufrom='" & nickname & "' or ufrom='" & dmjto &"'"
Set Rs=conn.Execute(sql)
sql="delete from mjInfo where id=" & mjID
Set Rs=conn.Execute(sql)
dmjGetmj= ErrSays(nickname,"只有13张牌了，不能再摸牌，你们打成和局，没有胜负。" )
end if
end if
End Function
'////////////////////////////////////////////////////////////////////吃牌
function dmjPeng(fn1,from1,to1)

if to1="大家" and fn1<>"认输了" then call ErrALT("请选择说话对象，不能象大家发送这个命令!")
sjjh_name=Session("sjjh_name")
if sjjh_name="" then Response.Redirect "../../error.asp?id=440"

nickname=Session("sjjh_name")
grade=Session("sjjh_grade")
chatroomsn=session("nowinroom")

Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
db="edmj.asp"
'connstr="driver={Microsoft Access Driver (*.mdb)};dbq=" & Server.MapPath(""&db&"")
'如果你的服务器支持jet4.0，请使用下载的连接方法，提高程序性能
connstr="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(""&db&"")
conn.open connstr 

sql="select oldmj from dmj where uto='"& nickname & "' and ufrom='"& to1 & "'"
Set Rs=conn.Execute(sql)
if rs.eof then
	rs.close
	conn.close
	set rs=nothing
	set conn=nothing
	call ErrAlt("你没有跟[ " & to1 & " ]打麻将!")
end if
oldmj=rs(0)
	rs.close
if fn1="问" then
	dmjPeng=nickname & ":" & Errsays(to1,"刚才打了一张 " & showMj(oldmj)) 
	conn.close
	set rs=nothing
	set conn=nothing
	Exit Function
end if
sql="select * from dmj where ufrom='"& nickname & "'"
Set Rs=conn.Execute(sql)
isMy=rs("isMy")
isFp=rs("isFp")
isGet=rs("isGet")
Mymj=rs("Mymj")
Gangmj4=rs("杠牌")
isGang=true
dmjto=rs("uto")
xiazhu=rs("duz")
logtime=rs("logtime")
if to1<>dmjto then
dmjPeng= ErrSays(nickname,"你的对手不是[" & to1 & "]啊？")
rs.close
conn.close
set conn=nothing
set rs=nothing
exit function
end if
if not isMy then
dmjPeng= ErrSays(nickname,"现在好象不是你发牌啊,等对方出牌了才是你发牌？")
rs.close
conn.close
set conn=nothing
set rs=nothing
exit function
end if
fn1=fn1 & "+"
fn1=replace(fn1,"++","+")
fn2=split(fn1,"+")

fn3=ubound(fn2)
if isFp and (not isGet) and fn3<>4 then
dmjPeng= ErrSays(nickname,"你已经摸了一张牌，不可以再吃别人的牌。")
rs.close
conn.close
set conn=nothing
set rs=nothing
exit function
end if
dim Mymj3
Mymj3=Mymj

for i=0 to fn3-1
if instr(Mymj3,fn2(i) & "|")<>0 then
Mymj3=replace(Mymj3,fn2(i) & "|","",1,1,1)
else
dmjPeng= ErrSays(nickname,"你有没有搞错啊，你有这些牌吗？")
rs.close
conn.close
set conn=nothing
set rs=nothing
exit function
end if
next
Rtp=right(oldmj,1)
Ltp=left(oldmj,1)
select case fn3
case 2
if oldmj=fn2(0) and fn2(0) =fn2(1) then
	Mymj=Mymj & oldmj & "|"
	dmjPeng= ErrSays(nickname,"出了两张 " & showMj(oldmj) & " 吃定[" & to1 & "]碰对子得到了三张牌，哈哈大笑，春风得意.....呵呵呵")
elseif Rtp=right(fn2(0),1) and right(fn2(0),1)=right(fn2(1),1) then
	if (instr(fn1,Ltp-1 & Rtp)<>0 and instr(fn1,Ltp-2 & Rtp)<>0) or (instr(fn1,Ltp+1 & Rtp)<>0 and instr(fn1,Ltp+2 & Rtp)<>0) or (instr(fn1,Ltp-1 & Rtp)<>0 and instr(fn1,Ltp+1 & Rtp)<>0) then
		dmjPeng= ErrSays(nickname,"出了两张 " & showMj(fn2(0)) & showMj(fn2(1)) & " 吃定 [" & to1 & "] " & showMj(oldmj) & "，.....得到了三张牌")
		Mymj=Mymj & oldmj & "|"
	else
		dmjPeng= ErrSays(nickname,"你出的这两张牌怎么吃别人的牌啊?")
		Mymj=""
	end if
else
	dmjPeng= ErrSays(nickname,"你出的这两张牌怎么吃别人的牌啊?")
	Mymj=""
end if

case 3
if oldmj=fn2(0) and fn2(0) =fn2(1) and fn2(1)=fn2(2) then
	Mymj=Mymj & oldmj & "|"
	Gangmj4=Gangmj4 & oldmj & "|"
	dmjPeng= ErrSays(nickname,"出了三张 " & showMj(oldmj) & " 吃定[" & to1 & "],哈哈大笑，春风得意.....[明]杠！！！")
	sql="update dmj set Mymj='" & Mymj & "',isGet=true,isFp=false,isMy=true,杠牌='" & Gangmj4 & "',logtime='" & now() & "' where ufrom='"& nickname & "'"
	Set Rs=conn.Execute(sql)
	dmjPeng=dmjPeng 
else
	dmjPeng= ErrSays(nickname,"你出的这三张牌怎 吃别人的牌啊?")
end if
Mymj=""
case 4
if instr(Gangmj4,fn2(0) & "|")<>0 then isGang=false
if fn2(0) =fn2(1) and fn2(1)=fn2(2) and fn2(2) =fn2(3) then
	if isGang then
		if isGet and (not isFp) then
			dmjPeng= ErrSays(nickname,"你应当先摸牌，然後再[暗]杠牌,在[洗牌]桌点击[该我摸牌了]使用(/摸牌)命令")
			%>
			<script language="JavaScript">parent.f2.document.forms[0].sytemp.value="/摸牌";</script>
			<%
			conn.close
			exit function
		end if

		Gangmj4=Gangmj4 & fn2(0) & "|"
		dmjPeng= ErrSays(nickname,"掷出四张......" & showMj(fn2(0)) & "，哈哈大笑，春风得意.....[暗]杠！！！")
		sql="update dmj set isGet=true,isFp=false,isMy=true,杠牌='" & Gangmj4 & "',logtime='" & now() & "' where ufrom='"& nickname & "'"
		Set Rs=conn.Execute(sql)
		conn.close
		dmjPeng=dmjPeng 
	else
		dmjPeng= ErrSays(nickname,showMj(fn2(0)) & "这张牌你已经杠了啊?还想杠？发晕了你？")
	end if
else
	dmjPeng= ErrSays(nickname,"你出的这四张牌怎么能杠啊?")
end if
Mymj=""

case else
dmjPeng= ErrSays(nickname,"你出的不是两张或三张？")
Mymj=""
end select
if Mymj<>"" then
sql="update dmj set Mymj='" & Mymj & "',isGet=false,isFp=true,isMy=true,logtime='" & now() & "' where ufrom='"& nickname & "'"
Set Rs=conn.Execute(sql)
sql="update dmj set isGet=true,isFp=false,isMy=false where ufrom='"& dmjto & "'"
Set Rs=conn.Execute(sql)
conn.close
end if
%>
<script language="JavaScript">parent.f3.location.reload();</script>
<%
set conn=nothing
set rs=nothing
end function
function showMj(Mj)
showMj="<input type=image border=0 src=""f2/mj/" & intMjp(Mj) & ".gif"" onclick=""javascript:parent.f3.location.href='f2/dmj-xp.asp';"" >"
end function
function strmj2(Mj)
dim Mj2
Mj2=replace(Mj,"1","一")
Mj2=replace(Mj2,"2","二")
Mj2=replace(Mj2,"3","三")
Mj2=replace(Mj2,"4","四")
Mj2=replace(Mj2,"5","五")
Mj2=replace(Mj2,"6","六")
Mj2=replace(Mj2,"7","七")
Mj2=replace(Mj2,"8","八")
strmj2=replace(Mj2,"9","九")
end function
function intMjp(cmj)
dim mj2
dim mjL
mj2=cmj
mjL=left(cmj,1)
if isNumeric(mjL) then 
mj2=right(cmj,1) & mjL
mj2=replace(mj2,"索","1")
mj2=replace(mj2,"筒","2")
mj2=replace(mj2,"万","3")
else
mj2=replace(mj2,"东风","10")
mj2=replace(mj2,"南风","20")
mj2=replace(mj2,"西风","30")
mj2=replace(mj2,"北风","40")
mj2=replace(mj2,"红中","41")
mj2=replace(mj2,"白板","42")
mj2=replace(mj2,"发财","43")
end if
intMjp=mj2
end function
%>
