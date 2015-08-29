<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="../sjfunc/sjfunc.asp"-->
<!--#include file="../sjfunc/func.asp"-->
<%'打扑克牌♀wWw.51eline.com♀
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
says=dpkfp(fnn1,sjjh_name,towho)
says="<font color=green><b>【发牌】</b><font color=" & saycolor & ">"&says&"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)

function dpkfp(fn1,from1,to1)
on error Resume next

if to1="大家" and fn1<>"认输了" then call ErrALT("请选择说话对象，不能象大家发送这个命令!"+fn1)

sjjh_name=Session("sjjh_name")
if sjjh_name="" then Response.Redirect "../../error.asp?id=440"

nickname=Session("sjjh_name")
grade=Session("sjjh_grade")
chatroomsn=session("nowinroom")

if left(fn1,2)="公证" then
	Errstr="错误的命令格式：\n正确示例：\n/发牌$ 公证|负\n/发牌$ 公证|胜"
	arr_fn1=split(fn1,"|")
	if ubound(arr_fn1)<>1 then call ErrALT(Errstr)
	fn_62=left(arr_fn1(1),1)

	if fn_62<>"负" and fn_62<>"胜" then call ErrALT(Errstr)
	if grade<6 then call ErrALT("你不是官府人员，不能进行牌局公证！")
	from1=to1
end if

if chatroomsn<>"0" then
	dpkfp=ErrSays(nickname,"『公证』或『认输』必须在大厅进行并接授大家的临督" )
	exit function
End if

Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
db="edpk.asp"
connstr="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(""&db&"")
conn.open connstr

sql="select * from dpk where ufrom='"& from1 & "'"
Set Rs=conn.Execute(sql)

if rs.eof or rs.bof then
	rs.close
	conn.close
	set rs=nothing
	set conn=nothing
	call ErrALT(from1 & "没有参与[扑克]牌局！")
end if

dpk=rs("pk")
dpkto=rs("uto")
xiazhu=rs("duz")
fpok=rs("fp")
oldoldpn=Cint(rs("oldpn"))
oldmaxpk=rs("maxpk")
logtime=rs("logtime")
iszai=rs("iszai")
rs.close
'dpkfp=cstr(oldoldpn)
'exit function

dim pk_geshi,old_pk_geshi
pk_geshi=1

if oldoldpn<10 then oldoldpn=90
old_pk_geshi=CINT(left(oldoldpn,1))
oldoldpn=CINT(mid(oldoldpn,2))

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
		conn.close
		set rs=nothing
		set conn=nothing
		call ErrALT("非法操作！")
end select

'------------------------新格式------------------------
d_sql="delete from dpk where ufrom='" & from1 & "' or uto='" & from1 &"'"

if fn_62<>"" then
	if fn_62="负" then
		s_pk=dpkto
		f_pk=to1
	elseif fn_62="胜" then
		f_pk=dpkto
		s_pk=to1
	End if
	Set Rs=conn.Execute(d_sql)
	conn.close
	connstr=Application("sjjh_usermdb")
	Set conn=Server.CreateObject("ADODB.CONNECTION")
	conn.open connstr 

	'------------------------新格式------------------------
	conn.execute "update 用户 set " & xia_sql & "=" & xia_sql & "+"&xiazhu&"*2 where 姓名='"& s_pk &"'"
	conn.close
	dpkfp="【官府公证】" & ErrSays(f_pk,"不遵守[扑克牌局]规定拒绝完成牌局不肯认输<br> ≮≮≮≮ 官府[ " & nickname & " ]裁定他输给[ " & s_pk & " ]" & xiazhu & xia_class & sayxia & " ≯≯≯≯ ")
elseif fn1<>"不要了" and fn1<>"认输了" then
  if to1<>dpkto then
    dpkfp= ErrSays(nickname,"你的对手不是[" & to1 & "]啊？")
    conn.close
    exit function
  end if
  if fpok then
    dpkfp= ErrSays(nickname,"现在好象不是你发牌啊,等对方出牌了才是你发牌？")
    conn.close
    exit function
  end if
'olddpk22=split(dpk,"|")
'olddpk3=ubound(olddpk22)
dim qysall
dim qysn
dim pkx
dim isArrpk(15)
dim isArrpkn(15)

fn1=fn1 & "+"
fn1=replace(fn1,"++","+")
fn2=split(fn1,"+")
fn3=ubound(fn2)
for i=0 to fn3-1
  if instr(dpk,fn2(i) & "|")<>0 then
     oldpn=oldpn+1
     pkx=mid(fn2(i),2)
     if fn2(i)="大王" or fn2(i)="小王" then pkx=fn2(i)
     pkx=unconvpk(pkx)
     isArrpkn(pkx)=isArrpkn(pkx) + 1
     isArrpk(pkx)=isArrpk(pkx) & " " & showpk(fn2(i))
     qys=left(fn2(i),1)
     qysn=qysn+1
     if qysall<>qys then qysall=qysall & qys
  end if
  dpk=replace(dpk,fn2(i) & "|","")
next
maxpk=0
dim zaid,sang,yidu,pkyizha2,pkyizha3,liandui
zaid=0
sang=0
yidu=0
pkyizha2=0
pkyizha3=0
pkyizha3=1
liandui=2

for i=15 to 1 step -1
  if isArrpkn(i)=4 then
    zaid=zaid+4
    fpnow =fpnow & " 四个(炸)" & isArrpk(i)
    if maxpk=0 then 
	maxpk=i
	pk_geshi=4
    end if
  end if
next

for i=15 to 1 step -1
  if isArrpkn(i)=3 then
    fpnow =fpnow & " 三个" & isArrpk(i)
    sang=sang+3
    if maxpk=0 then 
	maxpk=i
	pk_geshi=3
    end if
  end if
next

for i=15 to 1 step -1
  if isArrpkn(i)=2 then
     fpnow =fpnow & " 一对" & isArrpk(i)
     yidu=yidu+2
     if isArrpkn(i-1)=2 then liandui=liandui+2
     if maxpk=0 then 
	maxpk=i
	pk_geshi=2
     end if
  end if
next
for i=15 to 1 step -1
  if isArrpkn(i)=1 then
    pkyizha2=pkyizha2+1
    fpnow =fpnow & " 一张" & isArrpk(i)
    if isArrpkn(i-1)=1 then pkyizha3=pkyizha3+1
    if maxpk=0 then 
	maxpk=i
	pk_geshi=1
    end if
  end if
next
dpk22=split(dpk,"|")
dpk3=ubound(dpk22)
if len(qysall)=1 and qysn>1 then 
qysall="<b><font color=red>[清一色]</font></b>"
if maxpk=0 then 
	pk_geshi=pk_geshi+5
end if
else
qysall=""
end if

dim Errfp
Errfp=false
dim ErrStr
ErrStr=""
dim show_old_pk
show_old_pk=oldmaxpk

if oldmaxpk=11 then show_old_pk="J"
if oldmaxpk=12 then show_old_pk="Q"
if oldmaxpk=13 then show_old_pk="K"
if oldmaxpk=14 then show_old_pk="A"
if oldmaxpk=15 then show_old_pk="2"
if oldmaxpk=16 then show_old_pk="小王"
if oldmaxpk=17 then show_old_pk="大王"

        IF old_pk_geshi=2 then
		if oldoldpn<3 then
			e_Str="一对"
		else
			e_Str="连对"
		end if
	elseif old_pk_geshi=3 then
		e_Str="三个"
	elseif old_pk_geshi=4 then
		e_Str="炸弹"
	elseif old_pk_geshi=1 then
		if oldoldpn>4 then
			e_Str="一条龙最大的一张牌："
		else
			e_Str="一张" 
		end if
	end if
	if (old_pk_geshi-5>0 and oldoldpn<>0) then e_Str=e_Str & "清一色" 

if zaid<>4 and oldoldpn<>0 then
        
	if old_pk_geshi<>pk_geshi then
        	ErrStr= ErrSays(nickname,to1 & "刚才出的是[" & e_Str & "<B>" & show_old_pk & "</B>]啊")
	elseif oldpn<>oldoldpn then
		'erralt(oldpn)
		ErrStr= ErrSays(nickname,to1 & "刚才出的是[" & oldoldpn & "]张啊？")
	end if

	if ErrStr<>"" then
	dpkfp=ErrStr & "如果你没有这种牌，请在[洗牌]桌按[让 牌]" 
	conn.close
	exit function
	end if
end if


if pkyizha3=oldpn and oldpn>4 then 
qysall=qysall & "<b><font color=red>[一条龙]</font></b>"
IsyitiaoL=true
else
IsyitiaoL=false
end if
if zaid=0 and yidu=0 and sang=0 and IsyitiaoL=false then
isgood=false
else
isgood=true
end if


if (yidu>0 and pkyizha2>0) then Errfp=true 
if (zaid=0 and (yidu>2 and yidu<>liandui) ) then Errfp=true '两对以上不是连对
if (sang<>0 and yidu=0 and pkyizha2>2) then Errfp=true '有三个有一对还有一张以上零牌

if (sang<>0 and oldpn<>5 and oldpn<>3) then Errfp=true '三个大于五张 
if (zaid<>0 and oldpn>7) then Errfp=true '炸弹大于7张

if (oldpn=3 and yidu<>0) then Errfp=true '三张牌中有一对

if (pkyizha2>2 and sang<>0) then Errfp=true '如果两张不是一对
if (isgood=false and oldpn>1) then Errfp=true '没有炸没有对没有三个不是一条龙，而且不是一张牌
if Errfp=true then
dpkfp= ErrSays(nickname,"你想发的是没有规律的零牌，这么大的人打扑克也不会？看看帮助先！" )
conn.close
exit function
end if
if maxpk=14 then maxpk=16
if maxpk=15 then maxpk=17
'新的14.15必须在原来的1415后面
if maxpk=1 then maxpk=14
if maxpk=2 then maxpk=15

'response.write maxpk
'response.end
if maxpk<=Cint(oldmaxpk) then
  if (zaid>0 and iszai) or zaid=0 then
     'ErrALT(e_Str)
     dpkfp= ErrSays(nickname,to1 & "刚才出的牌[" & e_Str & "<B>" & show_old_pk & "</B>]比你大！如果你没有牌比刚才的牌大，请在[洗牌]桌按[让 牌]")
     conn.close
     exit function
   end if
end if

if dpk3>0 then
sql="update dpk set pk='" & dpk & "',fp=true,oldpn=" & pk_geshi & oldpn & ",maxpk=" & maxpk & ",logtime='" & now() & "' where ufrom='"& nickname & "'"
Set Rs=conn.Execute(sql)
sql="update dpk set fp=false,oldpn=" & pk_geshi & oldpn & ",maxpk=" & maxpk & ",iszai=" & iszai & " where ufrom='"& dpkto & "'"
Set Rs=conn.Execute(sql)
conn.close
dpkfp= ErrSays(nickname,"想了想，拿出[" & oldpn & "]张牌潇洒的甩上了牌桌，原来是......" & ismore & "<br> " & qysall & fpnow & " , 还有" & dpk3 & "张牌")
%>
<script language="JavaScript">parent.f3.location.reload();</script>
<%
elseif oldpn=18 then
sql="update dpk set fp=true,oldpn=90,maxpk=0 where ufrom='"& nickname & "'"
Set Rs=conn.Execute(sql)
sql="update dpk set fp=false,oldpn=90,maxpk=0,iszai=false where ufrom='"& dpkto & "'"
Set Rs=conn.Execute(sql)
conn.close
dpkfp= ErrSays(nickname,"你想作弊啊(不遵守游戏规则想抛光所有的牌)？现在要罚你，请[" & dpkto & "]出牌")
else

Set Rs=conn.Execute(d_sql)
conn.close
connstr=Application("sjjh_usermdb")
Set conn=Server.CreateObject("ADODB.CONNECTION")
conn.open connstr 

'------------------------新格式------------------------
conn.execute "update 用户 set " & xia_sql & "=" & xia_sql & "+"& xiazhu&"*2 where 姓名='"& nickname&"'" 
conn.close
dpkfp= ErrSays(nickname,"哈哈大笑，把所有的牌发了出来......" & ismore & "<br> " & fpnow & " , 没有牌了，从") & ErrSays(dpkto,"赢到" & xiazhu & xia_class & sayxia )
'------------------------新格式------------------------

end if

elseif fn1="不要了" then
sql="update dpk set fp=true,oldpn=90,maxpk=0 where ufrom='"& nickname & "'"
Set Rs=conn.Execute(sql)
sql="update dpk set fp=false,oldpn=90,maxpk=0,iszai=false where ufrom='"& dpkto & "'"
Set Rs=conn.Execute(sql)
dpkfp= ErrSays(nickname,"放弃打这张牌，请[" & dpkto & "]出牌")
conn.close
elseif fn1="认输了" then

Set Rs=conn.Execute(d_sql)
conn.close
connstr=Application("sjjh_usermdb")
Set conn=Server.CreateObject("ADODB.CONNECTION")
conn.open connstr 


'------------------------新格式------------------------
conn.execute "update 用户 set " & xia_sql & "=" & xia_sql & "+"&xiazhu&"*2 where 姓名='"& dpkto &"'"
conn.close
dpkfp= ErrSays(nickname,"认输了放弃打这局牌,输给[" & dpkto & "]" & xiazhu & xia_class & sayxia)
'------------------------新格式------------------------

end if
set conn=nothing
set rs=nothing
end function

function showpk(pk)
dim wpk
wpk=replace(pk,"红","<img src=f2/dpk/1.GIF border=0><font color=#FF0000 size=2>")
wpk=replace(wpk,"黑","<img src=f2/dpk/2.GIF border=0><font color=#000000 size=2>")
wpk=replace(wpk,"梅","<img src=f2/dpk/3.GIF border=0><font color=#000000 size=2>")
wpk=replace(wpk,"方","<img src=f2/dpk/4.GIF border=0><font color=#FF0000 size=2>")
wpk=replace(wpk,"小王","<img src=f2/dpk/xw.gif border=0>")
wpk=replace(wpk,"大王","<img src=f2/dpk/dw.gif border=0>")
showpk=wpk & "</font>"
end function
function unconvpk(cpk)
dim pk2
pk2=replace(cpk,"A","1")
pk2=replace(pk2,"J","11")
pk2=replace(pk2,"Q","12")
pk2=replace(pk2,"K","13")
pk2=replace(pk2,"小王","14")
pk2=replace(pk2,"大王","15")
unconvpk=pk2
end function
%>
