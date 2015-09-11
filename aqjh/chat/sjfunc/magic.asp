<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<!--#include file="func.asp"-->
<%'魔法
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if Instr(LCase(" " & application("aqjh_useronlinename"&nowinroom)),LCase(" "&aqjh_name&" "))=0 then 
	session("aqjh_inthechat")="0" 
	Response.Redirect "chaterr.asp?id=001" 
end if 

mc_shengyan=10				'圣言术需要的魔法点数
mc_dingshen=1				'定身术
mc_zhiyu=1						'治愈术
mc_haidilaoyue=5			'海底捞月(天眼开光)
mc_youhuozhiguang=30		'诱惑之光
mc_jitiliaoshang=500			'集体疗伤
mc_xinlingqishi=10				'心灵启示
mc_dizhenshu=20					'地震术
mc_beimingshengong=25			'北冥神功
mc_xilingzhou=20					'吸灵咒
mc_yuxue=1								'浴血
mc_syuxue=5							'圣浴血
mc_yuling=10								'浴灵
mc_syuling=50								'圣浴灵
mc_jingshen=50 
mc_jingshen=50 						'净身

aqjh_name=session("aqjh_name")
towho=Trim(Request.Form("towho"))
towhoway=Request.Form("towhoway")
saycolor=Request.Form("saycolor")
addwordcolor=Request.Form("addwordcolor")
addsays=Request.Form("addsays")
says=Trim(Request.Form("sy"))
call dianzan(towho)
if len(says)>150 then says=Left(says,150)
says=lcase(says)
says=Replace(says,"&amp;","&")
if aqjh_grade<9 then
says=bdsays(says)
end if
act=1
towhoway=0
i=instr(says," ")
fnn1=trim(mid(says,i+1))

set rsuser=session("bug_conn_npc").execute("select * from 用户 where 姓名='" & aqjh_name & "'")
strmagic=rsuser("魔法")
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select * from 用户 where 姓名='" & towho &"'",conn,2,2
if rs.eof then
   rs.close
   rs.open "select * from [npc] where n姓名='" & towho &"'",conn,2,2
   if rs.eof then
      wt=""
   else
      wt="npc"
   end if
else
   wt="ren"
end if
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
if wt="npc" then
	set rsn=session("bug_conn_npc").execute("select * from npc where 姓名='" & towho & "'")
else
	set rsn=session("bug_conn_npc").execute("select * from 用户 where 姓名='" & towho & "'")
end if
if rsn.eof and wt=false and towho<>"大家" then response.write "<script>alert('你不能对" & towho & "操作！');</script>" :  response.end

'response.write strmagic
if instr("," & strmagic,"," & fnn1 & "|")=0 then response.write "<script>alert('你目前还无法领悟" & fnn1 & "这种高深的武功');</script>" :  response.end

'魔法公共等级参数
	dim tmp
	dim tmp1
	tmp=instr("," & strmagic,"," & fnn1 & "|")
	tmp1=right(strmagic,len(strmagic)-tmp+1)
	dj=mid(tmp1 ,instr(tmp1,"|")+1,instr(tmp1 & ",",",")-instr(tmp1,"|")-1)
	if int(sqr(dj+1))-int(sqr(dj))=1 then sjaddwords=",<font color=red><b>" & aqjh_name & "的" & fnn1 & "(" & int(sqr(dj)) & "->" & int(sqr(dj+1)) & ")升级了！</b></font>"
	kuandu=50+int(sqr(dj))
select case fnn1
	case "治愈术"
		says="<font color=red>【魔法】<img src='../shenbing/mofa/zhiyu.gif' width="&kuandu&"><font color=red><b>(" & int(sqr(dj)) & "级)</b></font><font color=" & saycolor & ">" & magic_zhiyu(towho) & "</font>"
	case "圣言术"
		says="<font color=red>【魔法】<img src='../shenbing/mofa/shengyan.gif' width="&kuandu&"><font color=red><b>(" & int(sqr(dj)) & "级)</b></font><font color=" & saycolor & ">" & magic_shengyan(towho) & "</font>"
	case "定身术"
		says="<font color=red>【魔法】<img src='../shenbing/mofa/dingshen.gif' width="&kuandu&"><font color=red><b>(" & int(sqr(dj)) & "级)</b></font><font color=" & saycolor & ">" & magic_dingshenshu(towho) & "</font>"
	case "天眼开光"
		says="<font color=red>【魔法】<img src='../shenbing/mofa/haidi.gif' width="&kuandu&"><font color=red><b>(" & int(sqr(dj)) & "级)</b></font><font color=" & saycolor & ">" & magic_haidilaoyue() & "</font>"
	case "诱惑之光"
		says="<font color=red>【魔法】<img src='../shenbing/mofa/youhuo.gif' width="&kuandu&"><font color=red><b>(" & int(sqr(dj)) & "级)</b></font><font color=" & saycolor & ">" & magic_youhuozhiguang(towho) & "</font>"
	case "集体疗伤"
		says="<font color=red>【魔法】<img src='../shenbing/mofa/jiti.gif' width="&kuandu&"><font color=red><b>(" & int(sqr(dj)) & "级)</b></font><font color=" & saycolor & ">" & magic_jitiliaoshang() & "</font>"
	case "心灵启示"
		says="<font color=red>【魔法】<img src='../shenbing/mofa/xinling.gif' width="&kuandu&"><font color=red><b>(" & int(sqr(dj)) & "级)</b></font><font color=" & saycolor & ">" & magic_xinlingqishi(towho)  & "</font>" 
	case "地震术"
		says="<font color=red>【魔法】<img src='../shenbing/mofa/dizhen.gif' width="&kuandu&"><font color=red><b>(" & int(sqr(dj)) & "级)</b></font><font color=" & saycolor & ">" & magic_dizhenshu(towho)  & "</font>" 
	case "吸灵咒"
		says="<font color=red>【魔法】<img src='../shenbing/mofa/xiling.gif' width="&kuandu&"><font color=red><b>(" & int(sqr(dj)) & "级)</b></font><font color=" & saycolor & ">" & magic_xilingzhou(towho)  & "</font>" 
	case "北冥神功"
		says="<font color=red>【魔法】<img src='../shenbing/mofa/beiming.gif' width="&kuandu&"><font color=red><b>(" & int(sqr(dj)) & "级)</b></font><font color=" & saycolor & ">" & magic_beimingshengong(towho)  & "</font>" 
	case "浴灵"
		says="<font color=red>【魔法】<img src='../shenbing/mofa/beiming.gif' width="&kuandu&"><font color=red><b>(" & int(sqr(dj)) & "级)</b></font><font color=" & saycolor & ">" & magic_yuling(towho)  & "</font>" 
	case "圣浴灵"
		says="<font color=red>【魔法】<img src='../shenbing/mofa/beiming.gif' width="&kuandu&"><font color=red><b>(" & int(sqr(dj)) & "级)</b></font><font color=" & saycolor & ">" & magic_syuling(towho)  & "</font>" 
	case "浴血"
		says="<font color=red>【魔法】<img src='../shenbing/mofa/beiming.gif' width="&kuandu&"><font color=red><b>(" & int(sqr(dj)) & "级)</b></font><font color=" & saycolor & ">" & magic_yuxue(towho)  & "</font>" 
	case "圣浴血"
		says="<font color=red>【魔法】<img src='../shenbing/mofa/beiming.gif' width="&kuandu&"><font color=red><b>(" & int(sqr(dj)) & "级)</b></font><font color=" & saycolor & ">" & magic_syuxue(towho)  & "</font>" 
	case "净身"
		says="<font color=red>【魔法】<img src='../shenbing/mofa/beiming.gif' width="&kuandu&"><font color=red><b>(" & int(sqr(dj)) & "级)</b></font><font color=" & saycolor & ">" & magic_jingshen(towho)  & "</font>" 
end select

says=says & sjaddwords
session("bug_conn_npc").execute("update 用户 set 魔法='" & replace(strmagic,fnn1 & "|" & dj,fnn1 & "|" & cstr(dj+1)) & "' where 姓名='" & aqjh_name & "'")
if fnn1="心灵启示" then
	towhoway=1
else
	towhoway=0
end if
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
'魔法

function magic_zhiyu(towho)					'治愈术
'set rsnpc=session("bug_conn_npc").execute("select * from npc where 姓名='" & towho & "'")
	need_nl=int(10 * rsuser("等级") * (sqr(int(sqr(dj)))))
	if rsuser("内力")<need_nl then response.write "<script>alert('你的内力好像不够了~');</script>" :  response.end
	if rsuser("魔力")<zhiyu then response.write "<script>alert('你的魔力不够~');</script>" :  response.end
	if bug_ifnpc(towho) then
		if bug_read(towho,"体力")>= rsn("体力") then response.write "<script>alert('对方已经不需要再治疗了~');</script>" :  response.end
		rsn.close
		set rsn=nothing
		session("bug_conn_npc").execute("update 用户 set 内力=内力-" & need_nl & " where 姓名='" & aqjh_name & "'")
		call bug_modify(towho,"体力",10*need_nl)
	else
		if rsn("体力")>=rsn("等级")*application("aqjh_tlsx")+5260+rsn("体力加") then response.write "<script>alert('对方已经不需要再治疗了~');</script>" :  response.end
		session("bug_conn_npc").execute("update 用户 set 内力=内力-" & need_nl & " where 姓名='" & aqjh_name & "'")
		session("bug_conn_npc").execute("update 用户 set 体力=体力+" & need_nl*10 & " where 姓名='" & towho & "'")
	end if
	
	magic_zhiyu="##全身放松，开始运功替%%疗伤。%%吐出一口瘀血，脸色看起来好多了,%%体力上升了<font color=red>" & need_nl*10 & "</font>点。##消耗了<font color=red>" &  need_nl & "</font>点内力" 
end function


function magic_shengyan(towho)					'圣言术	
	if wt<>"npc" then response.write "<script>alert('圣言术只能对NPC使用~');</script>" :  response.end
	tmp3=int(bug_read(towho,"类型"))-9
	if rsuser("魔力")<mc_shengyan*tmp3 then response.write "<script>alert('你的魔力不够~');</script>" :  response.end
	if tmp3>sqr(dj) then response.write "<script>alert('对方太强大了，恐怕是不行');</script>" :  response.end
	session("bug_conn_npc").execute("update 用户 set 魔力=魔力-" & mc_shengyan*tmp3 & " where 姓名='" & aqjh_name & "'")
	randomize(timer)
	if rnd<0.8 then
		session("bug_conn_npc").execute("update 用户 set 魔力=魔力-" & mc_shengyan*tmp3 & " where 姓名='" & aqjh_name & "'")
		magic_shengyan="##喃喃念了几句咒语，突然天昏地暗，一阵狂风席卷过来，伴随着一声惨叫，%%竟然消失得无影无踪,##消耗了<font color=red>" &  mc_shengyan*tmp3 & "</font>点魔法" 
	else
		magic_shengyan="##喃喃念了几句咒语，可惜%%没什么反应，##消耗了<font color=red>" &  mc_shengyan*tmp3 & "</font>点魔法" 
	end if
end function
	
function magic_youhuozhiguang(towho)			'诱惑之光
	if wt<>"npc" then response.write "<script>alert('此魔法只能对NPC使用~');</script>" :  response.end
	if bug_read(towho,"主人")=aqjh_name then response.write "<script>alert('喂喂喂！诱惑自己的NPC干嘛？');</script>" :  response.end
	tmp3=int(bug_read(towho,"类型"))-9
	if rsuser("魔力")<mc_youhuozhiguang*tmp3 then response.write "<script>alert('你的魔力不够~');</script>" :  response.end
	if tmp3>sqr(dj) then response.write "<script>alert('对方太强大了，恐怕是不行');</script>" :  response.end
	session("bug_conn_npc").execute("update 用户 set 魔力=魔力-" & mc_youhuozhiguang*tmp3 & " where 姓名='" & aqjh_name & "'")
	
	randomize(timer)
	if rnd<0.8 then
		call bug_modify(towho,"主人",aqjh_name)
		magic_youhuozhiguang="##冲着%%喃喃念了几句咒语，突然无数条光芒射向%%,%%一下子目光僵直，居然乖乖听##的吩咐了,##消耗了<font color=red>" &  mc_youhuozhiguang*tmp3 & "</font>点魔法" 
	else
		magic_youhuozhiguang="##冲着%%喃喃念了几句咒语，可惜%%没什么反应,##消耗了<font color=red>" &  mc_youhuozhiguang*tmp3 & "</font>点魔法" 
	end if
end function

function magic_dingshenshu(towho)						'定身术
	if bug_ifnpc(towho)=false then response.write "<script>alert('此魔法只能对NPC使用~');</script>" :  response.end
	tmp3=int(bug_read(towho,"类型"))
	if rsuser("魔力")<mc_dingshen*tmp3 then response.write "<script>alert('你的魔力不够~');</script>" :  response.end
	session("bug_conn_npc").execute("update 用户 set 魔力=魔力-" & mc_dingshen*tmp3 & " where 姓名='" & aqjh_name & "'")
	call bug_modify(towho,"时间",cstr(now()+(sqr(sqr(dj))*60+10)/864000))
	magic_dingshenshu="##冲着%%大吼了一声“定！” %%乖乖不动了,要到" & cstr(now()+(sqr(sqr(dj))*60+10)/864000) & "才可以活动,##消耗了<font color=red>" &  mc_dingshen*tmp3 & "</font>点魔法" 
end function

function magic_yuling(towho)									'浴灵
	'if aqjh_name<>towho then response.write "<script>alert('此魔法只能对自己使用~');</script>" :  response.end
	if rsuser("魔力")<mc_yuling*int(sqr(dj)) then response.write "<script>alert('你的魔力不够~');</script>" :  response.end
    session("magic_gj")=3
    session("magic_baodou")=0.1+int(int(sqr(dj))^0.3*100)/100
hchyl=0.1+int(int(sqr(dj))^0.3*100)/100
	session("bug_conn_npc").execute("update 用户 set 魔力=魔力-" & mc_yuling*int(sqr(dj))& ",攻击=攻击*" & hchyl & "  where 姓名='" & aqjh_name & "'")
	magic_yuling="##用食指压住静脉，瞬间聚起一股劲道，施展[浴灵]，灌输到##体内,一瞬间##攻击力暴涨<font color=red><b>" & 0.1+int(int(sqr(dj))^0.3*100)/100 & "</b></font>倍!消耗了<font color=red>" &  mc_yuling*int(sqr(dj)) & "</font>点魔法." 
end function

function magic_syuling(towho)									'圣浴灵
	'if aqjh_name<>towho then response.write "<script>alert('此魔法只能对自己使用~');</script>" :  response.end
	if rsuser("魔力")<mc_syuling*int(sqr(dj)) then response.write "<script>alert('你的魔力不够~');</script>" :  response.end
    session("magic_gj")=30
    session("magic_baodou")=0.1+int(int(sqr(dj))^0.3*100)/100
hchsyl=0.1+int(int(sqr(dj))^0.3*100)/100
	session("bug_conn_npc").execute("update 用户 set 魔力=魔力-" & mc_syuling*int(sqr(dj)) & ",攻击=攻击*" & hchsyl & " where 姓名='" & aqjh_name & "'")
	magic_syuling="##用食指压住静脉，瞬间聚起一股劲道，施展[圣浴灵]，灌输到##体内,一瞬间##攻击力暴涨<font color=red><b>" & 0.1+int(int(sqr(dj))^0.3*100)/100 & "</b></font>倍!消耗了<font color=red>" &  mc_syuling*int(sqr(dj)) & "</font>点魔法." 
end function

function magic_yuxue(towho)									'浴血
	if aqjh_name<>towho then response.write "<script>alert('此魔法只能对自己使用~');</script>" :  response.end
	if rsuser("魔力")<mc_yuxue*int(sqr(dj)) then response.write "<script>alert('你的魔力不够~');</script>" :  response.end
    session("magic_fy")=3
    session("magic_baodou1")=0.1+int(int(sqr(dj))^0.3*100)/100
hchyx=0.1+int(int(sqr(dj))^0.3*100)/100
	session("bug_conn_npc").execute("update 用户 set 魔力=魔力-" & mc_yuxue*int(sqr(dj)) & ",防御=防御*" & hchyx & " where 姓名='" & aqjh_name & "'")
	magic_yuxue="##咬破舌尖，哇的突出一口鲜血，瞬间聚起一股气，施展[浴血]，灌输到##体内,一瞬间##防御力暴涨<font color=red><b>" & 0.1+int(int(sqr(dj))^0.3*100)/100 & "</b></font>倍!消耗了<font color=red>" &  mc_yuxue*int(sqr(dj)) & "</font>点魔法." 
end function

function magic_syuxue(towho)									'圣浴血
	'if aqjh_name<>towho then response.write "<script>alert('此魔法只能对自己使用~');</script>" :  response.end
	if rsuser("魔力")<mc_syuxue*int(sqr(dj)) then response.write "<script>alert('你的魔力不够~');</script>" :  response.end
    session("magic_fy")=10
    session("magic_baodou1")=0.1+int(int(sqr(dj))^0.3*100)/100
hchsyx=0.1+int(int(sqr(dj))^0.3*100)/100
	session("bug_conn_npc").execute("update 用户 set 魔力=魔力-" & mc_syuxue*int(sqr(dj)) & ",防御=防御*" & hchsyx & " where 姓名='" & aqjh_name & "'")
	magic_syuxue="##咬破舌尖，哇的突出一口鲜血，瞬间聚起一股气，施展[圣浴血]，灌输到##体内,一瞬间##防御力暴涨<font color=red><b>" & 0.1+int(int(sqr(dj))^0.3*100)/100 & "</b></font>倍!消耗了<font color=red>" &  mc_syuxue*int(sqr(dj)) & "</font>点魔法." 
end function

function magic_jingshen(towho)									'净身
	if aqjh_name<>towho then response.write "<script>alert('此魔法只能对自己使用~');</script>" :  response.end
	if rsuser("魔力")<mc_jingshen then response.write "<script>alert('你的魔力不够~');</script>" :  response.end
	
	baolou=""
	for i = 1 to 6
		c1="z" & i
		c2=rsn(c1)
		if c2<>"" then 	
			c3=left(c2,instr(c2,"|")-1)
			application.lock
			application("aqjh_baowp") = application("aqjh_baowp") & c3 & "|" & aqjh_name & "|" & now() & ";"
			application.unlock
			baolou =  baolou & "<input type=submit name=id value='" & c3 & "' style='border: 1px solid;font-size: 9pt;' onclick=window.open('bao.asp?id=" & c3 & "','d');>"
		end if
	next
	session("bug_conn_npc").execute("update 用户 set z1='',z2='',z3='',z4='',z5='',z6='' where 姓名='" & aqjh_name & "'")
	session("bug_conn_npc").execute("update 用户 set 魔力=魔力-" & mc_jingshen & " where 姓名='" & aqjh_name & "'")
	magic_jingshen="##慢慢运功，脸由白到红，由红到紫，由紫到黑，突然大喝一声:“开!” 身上的装备全部被震掉。" & baolou & " ##消耗了<font color=red>" &  mc_jingshen & "</font>点魔法." 
	baolou=""
end function

function magic_haidilaoyue()					'海底捞月（天眼开光）
	'call aqjh_autodel()
	baowp=application("aqjh_baowp")
	wlist=split(baowp,";")
	for each j in wlist
		if trim(j)<>"" then
			zlist=split(j,"|")
			tmps = tmps & "<input type=submit name=id value='" & zlist(0) & "' style='border: 1px solid;font-size: 9pt;' onclick=window.open('bao.asp?id=" & zlist(0) & "','d');> "
		end if
	next
	session("bug_conn_npc").execute("update 用户 set 魔力=魔力-" & mc_haidilaoyue & " where 姓名='" & aqjh_name & "'")
	magic_haidilaoyue="##微一凝神，施展「天眼开光」之术。双手合为剑指在眼前慢慢拉开，双眼竟变为金黄色，仿佛能洞穿一切！地面上的物品也全部显示了出来。" & tmps & ",##消耗了<font color=red>" &  mc_haidilaoyue & "</font>点魔法"
end function

function magic_jitiliaoshang()							'集体疗伤
	if rsuser("魔力")<mc_jitiliaoshang then response.write "<script>alert('你的魔力不够~');</script>" :  response.end
	ql_nlsx=rsuser("等级")*application("aqjh_nlsx")+5260+rsuser("内力加")
	if rsuser("内力")<ql_nlsx*0.9 then response.write "<script>alert('你的内力不足九成，恐怕一运功会有生命危险！');</script>" :  response.end
	liaoed=int(ql_nlsx*(sqr(int(sqr(dj)))/30+0.1))
	if liaoed>rsuser("内力") then liaoed=rsuser("内力")
	tmp=trim(application("aqjh_useronlinename" & session("nowinroom")))
	while instr(tmp,"  ")>0 
		tmp=replace(tmp,"  "," ")
	wend
	tmp1=split(tmp," ")
	for each h in tmp1
		if bug_ifnpc(h)  then
			if  bug_read(h,"主人")<>"无" then
				set rs1=session("bug_conn_npc").execute("select * from npc where 姓名='" & h & "'")
				call bug_modify(h,"体力",rs1("体力")-bug_read(h,"体力"))
			end if
		else
			set rs1=session("bug_conn_npc").execute("select * from 用户 where 姓名='" & h & "'")
			session("bug_conn_npc").execute("update 用户 set 体力=" & rs1("等级")*application("aqjh_tlsx")+5260+rs1("体力加") & " where 姓名='" & h & "'")
		end if
	next
	session("bug_conn_npc").execute("update 用户 set 魔力=魔力-" & mc_jitiliaoshang & ",内力=" & ql_nlsx*0.1 & " where 姓名='" & aqjh_name & "'")
	magic_jitiliaoshang="##使出集体疗伤术，所有人的体力都补满了,##消耗了<font color=red>" &  mc_jitiliaoshang & "</font>点魔法,但是由于发功过度，自己的内力只剩下<font color=red>" & liaoed & "</font>点。" 
end function

function magic_xinlingqishi(towho)					'心灵启示
	if rsuser("魔力")<mc_xinlingqishi then response.write "<script>alert('你的魔力不够~');</script>" :  response.end
	if ifnpc then
		g_tl=bug_read(towho,"体力")
		g_nl=bug_read(towho,"内力")
		g_l=int(100*g_tl/rsn("体力"))
		if g_l<0 then g_l=0
		if g_l>100 then g_l=100
		g_r=100-g_l
		tlt="<img height=6 border=0 src=img/11.gif width=" & g_l & "><img height=6 border=0 src=img/5.gif width=" & g_r & ">"
		strout= "##消耗了<font color=red>" & mc_xinlingqishi  & "</font>点魔法，看到了<b>" & towho & "</b><img src='npc/" & bug_read(towho,"图片") & "'>体力:" & tlt & "(" & g_tl & "/" & rsn("体力") & ")"
		if dj>60 then 
			g_l=int(100*g_nl/rsn("内力"))
			if g_l<0 then g_l=0
			if g_l>100 then g_l=100
			g_r=100-g_l
			nlt="<img height=6 border=0 src=img/11.gif width=" & g_l & "><img height=6 border=0 src=img/5.gif width=" & g_r & ">"
			strout = strout & " 内力:" & nlt & "(" & g_nl & "/" & rsn("内力") & ")"
		end if
		if dj>80 then strout = strout & " 银两:" & bug_read(towho,"银两")
		if dj>120 then strout = strout & " 敌人:" & bug_read(towho,"敌人")
		if dj>150 then strout = strout & " 攻击:" & bug_read(towho,"攻击") 
		if dj>160 then strout = strout & " 防御:" & bug_read(towho,"防御")
	else
		g_tl=rsn("体力")
		g_nl=rsn("内力")
		g_n=rsn("等级")*application("aqjh_tlsx")+5260+rsn("体力加")
		g_l=int(100*g_tl/g_n)
		if g_l<0 then g_l=0
		if g_l>100 then g_l=100
		g_r=100-g_l
		tlt="<img height=6 border=0 src=img/11.gif width=" & g_l & "><img height=6 border=0 src=img/5.gif width=" & g_r & ">"
		strout= "##消耗了<font color=red>" & mc_xinlingqishi  & "</font>点魔法，看到了<b>" & towho & "</b><img src='" & rsn("名单头像") & "'>体力:" & tlt & "(" & rsn("体力") & "/" & g_n  & ")"
		if dj>60 then 
			g_t=rsn("等级")*application("aqjh_nlsx")+2000+rsn("内力加")
			g_l=int(100*g_nl/g_t)
			if g_l<0 then g_l=0
			if g_l>100 then g_l=100
			g_r=100-g_l
			nlt="<img height=6 border=0 src=img/11.gif width=" & g_l & "><img height=6 border=0 src=img/5.gif width=" & g_r & ">"
			strout = strout & " 内力:" & nlt & "(" & rsn("内力") & "/" & g_t & ")"
		end if
		if dj>200 then strout = strout & " 攻击:" & rsn("攻击")
		if dj>300 then strout = strout & " 防御:" & rsn("防御")
		if dj>380 then strout = strout & " <BR>头盔:" & myzb("z1") 
		if dj>460 then strout = strout & " <BR>装饰:" & myzb("z2") 
		if dj>500 then strout = strout & " <br>盔甲:" & myzb("z3") 
		if dj>550 then strout = strout & " <br>左手:" & myzb("z4") 
		if dj>600 then strout = strout & " <br>右手:" & myzb("z5") 
		if dj>800 then strout = strout & " <br>双脚:" & myzb("z6") 
	end if
	session("bug_conn_npc").execute("update 用户 set 魔力=魔力-" & mc_xinlingqishi & " where 姓名='" & aqjh_name & "'")
	magic_xinlingqishi=strout 
end function

function magic_dizhenshu(towho)						'地震术
	if towho="大家" or towho=aqjh_name or wt="npc" then response.write "<script>alert('你不能对大家或自己或者NPC使用地震术~');</script>" :  response.end
	if rsuser("魔力")<mc_dizhenshu then response.write "<script>alert('你的魔力好像不够哦');</script>" :  response.end
	session("bug_conn_npc").execute("update 用户 set 魔力=魔力-" & mc_dizhenshu & " where 姓名='" & aqjh_name & "'")
	
		dizhenshu_tl=int(sqr(int(sqr(dj)))*3000*rsuser("等级"))
		if dizhenshu_tl>rsuser("体力")*0.1 then dizhenshu_tl=int(rsuser("体力")*0.1)
		dizhenshu_nl=int(sqr(int(sqr(dj)))*3000*rsuser("等级"))
		if dizhenshu_nl>rsuser("内力")*0.1 then dizhenshu_nl=int(rsuser("内力")*0.1)
		session("bug_conn_npc").execute("update 用户 set 内力=内力-" & dizhenshu_nl & ",体力=体力-" & dizhenshu_tl & " where 姓名='" & towho & "'")
	
	magic_dizhenshu="##闭上眼睛，默默念起咒语，%%身边居然冒出一阵阵青烟，突然轰隆隆一声，%%遭到地震的袭击！%%体力下降<font color=red>" & dizhenshu_tl & "</font>点,内力下降<font color=red>" & dizhenshu_nl & "</font>点.##耗费了<font color=red>" & mc_dizhenshu & "</font>点魔法"
end function

function magic_xilingzhou(towho)							'吸灵咒
	if towho="大家" or towho=aqjh_name or wt="npc" then response.write "<script>alert('你不能对大家或自己或者NPC使用吸灵咒~');</script>" :  response.end
	if rsuser("魔力")<mc_xilingzhou then response.write "<script>alert('你的魔力好像不够哦');</script>" :  response.end
	
		xilingzhou_tl=int(sqr(int(sqr(dj)))*1000*rsuser("等级"))
		if xilingzhou_tl>rsuser("体力")*0.1 then xilingzhou_tl=int(rsuser("体力")*0.1)
		if xilingzhou_tl>rsn("体力") then xilingzhou_tl=rsn("体力")
		session("bug_conn_npc").execute("update 用户 set 体力=体力-" & xilingzhou_tl & " where 姓名='" & towho & "'")	
		session("bug_conn_npc").execute("update 用户 set 体力=体力+" & xilingzhou_tl & ",魔力=魔力-" & mc_xilingzhou & " where 姓名='" & aqjh_name & "'")
	magic_xilingzhou="##叽里咕噜念了一阵咒语，把%%看作一支雪茄，拼命地吸，结果%%被吸取了体力<font color=red>" & xilingzhou_tl & "</font>点，##耗费了<font color=red>" & mc_xilingzhou & "</font>点魔力的同时，体力增长了<font color=red>" & xilingzhou_tl & "</font>点"
end function

function magic_beimingshengong(towho)					'北冥神功
	if towho="大家" or towho=aqjh_name or wt="npc" then response.write "<script>alert('你不能对大家或自己或者NPC使用北冥神功~');</script>" :  response.end
	if rsuser("魔力")<mc_beimingshengong then response.write "<script>alert('你的魔力好像不够哦');</script>" :  response.end
	
		beimingshengong_nl=int(sqr(int(sqr(dj)))*2000*rsuser("等级"))
		if beimingshengong_nl>rsuser("体力")*0.1 then beimingshengong_nl=int(rsuser("内力")*0.1)
		session("bug_conn_npc").execute("update 用户 set 内力=内力-" & beimingshengong_nl & " where 姓名='" & towho & "'")
		session("bug_conn_npc").execute("update 用户 set 内力=内力+" & beimingshengong_nl & ",魔力=魔力-" & mc_beimingshengong & " where 姓名='" & aqjh_name & "'")
	magic_beimingshengong="##呜哩哇啦念了一阵咒语，把%%看作一块蛋糕，大嚼特嚼，结果%%被吸取了内力<font color=red>" & beimingshengong_nl & "</font>点，##耗费了<font color=red>" & mc_beimingshengong & "</font>点魔力的同时，内力增长了<font color=red>" & beimingshengong_nl & "</font>点"	
end function

function myzb(lb)
if isnull(rsn(lb)) or trim(rsn(lb))="" then
	myzb="未装备任何物品!"
	exit function
end if
data1=split(rsn(lb),"|")
set rs=session("bug_conn_npc").execute("select * FROM b where a='" & data1(0) &"'")
If Rs.Bof OR Rs.Eof then
	myzb="未装备任何物品!"
	exit function
end if
if instr(rs("b"),"[")>0 then wid="65" else wid="40"
myzb="<img src='../hcjs/jhjs/images/" & rs("i") & "' alt='"&rs("c")&"' width=" & wid & ">["&rs("a")&"] 攻击:"&rs("f")&" 防御:"&rs("g")&" 耐久:<font color=red>"&data1(1)&"</font>/<font color=blue>"&rs("j")&"</font> 特性:"&rs("k")
rs.close
end function
%>
