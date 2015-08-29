<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="../sjfunc/sjfunc.asp"-->
<!--#include file="../sjfunc/func.asp"-->
<%'打扑克♀wWw.51eline.com♀
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
fnn1=mid(says,i+1)
says=dpkask(fnn1,sjjh_name,towho)
says="<font color=green><b>【打扑克】</b><font color=" & saycolor & ">"&says&"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)

function dpkask(fn1,from1,to1)
if sjjh_jhdj<=jhdj_dpk then
	call mess ("提示：打扑克需要["&jhdj_dpk&"]级才可以操作！",1)
end if
nickname=Session("sjjh_name")
grade=Session("sjjh_grade")
chatroomsn=session("nowinroom")

if to1="大家" or to1=from1 then call ErrALT("你不可以选择大家或自已作为对手！")

'------------------------新格式------------------------
ARR_FN=Split(fn1,"|")
dim ub_Err,xia_class,xiamoney,xia_fir,xia_sql
ub_Err="≮操作错误≯命令格式如下：\n /打扑克 下注数目|银两[或：金币，豆点] \n\n≮示例≯：\n /打扑克 1000|银两\n /打扑克 1000|金币\n /打扑克 1000|体力\n /打扑克 1000|内力 "

if ubound(ARR_FN)<>1 Then call ErrALT(ub_Err)
xia_class=ARR_FN(1)
xiamoney=ARR_FN(0)

select case xia_class
	case "金币"
		xia_fir="1"
		xia_sql="金币"
	case "银两"
		xia_fir="2"
		xia_sql="银两"
	case "豆点"
		xia_fir="3"
		xia_sql="泡豆点数"
	case else
		call ErrALT(ub_Err)
end select

If not isnumeric(xiamoney) Then call ErrALT(ub_Err)

xiamoney=abs(int(xiamoney))

if (xia_fir="1")and(xiamoney<10 or xiamoney>100) then Call ErrALT("最少押：10" & xia_class & "，最多100个金币！")
if (xia_fir="2")and(xiamoney<1000 or xiamoney>50000000) then Call ErrALT("最少押：1000" & xia_class & "，最多5000万银两！")
if (xia_fir="3")and(xiamoney<100 or xiamoney>3000) then Call ErrALT("最少押：100" & xia_class & "，最多5000豆点！")

Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
'------------------------新格式------------------------
sql="select " & xia_sql & " from 用户 where 姓名='" &from1&"'"
set rs=conn.execute(sql)
yin1=rs(0)
rs.close
if xiamoney> yin1 then
	set rs=nothing
	conn.close
	set conn=nothing
	Call ErrALT("你的身上好象没有这么多" & xia_class)
end if

sql="select " & xia_sql & " from 用户 where 姓名='" &to1&"'"
set rs=conn.execute(sql)
yin2=rs(0)
rs.close
if xiamoney>yin2 then
	set rs=nothing
	conn.close
	set conn=nothing
	Call ErrALT("对方身上好象没有这么多" & xia_class)
end if


'------------------------新格式------------------------
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
db="eDPK.ASP"

'connstr="driver={Microsoft Access Driver (*.mdb)};dbq=" & Server.MapPath(""&db&"")
'如果你的服务器支持jet4.0，请使用下载的连接方法，提高程序性能
connstr="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(""&db&"")
conn.open connstr 
sql="select * from dpk where ufrom='"& nickname & "' or uto='"& nickname &"'"
Set Rs=conn.Execute(sql)

if not (rs.eof and rs.bof) then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	response.write "<script Language='Javascript'>alert('您还有牌局没有结束，不能开局');parent.f3.location.href='dpk-xp.asp';</script>"
	response.end
end if
rs.close
sql="select * from dpk where ufrom='"& to1 & "' or uto='"& to1 &"'"
Set Rs=conn.Execute(sql)
if not (rs.eof and rs.bof) then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	response.write "<script Language='Javascript'>alert('"&to1&"还有牌局没有结束，不能开局');parent.f3.location.href='dpk-xp.asp';</script>"
	response.end
end if
rs.close
'------------------------新格式------------------------
dpkask="<bgsound src=wav/ring.WAV loop=1><b><font color=green>["&from1&"]</font></b>对<b><font color=black>["&to1&"]</font></b>说：打扑克吗？赌注为<b><font color=red>"&xiamoney & xia_class & "</font></b>，<img src='f2/dpk/1.GIF'><input type=button value='愿意' onclick=""javascript:window.open('f2/dpkok.asp?id="&myid&"&from1="&from1&"&to1="&to1&"&yn=1','a','width=380,height=210');this.disabled=1"" style=""background-color:#86A231;color:FFFFFF;border: 1 double"" onMouseOver=""this.style.color='FFFF00'"" onMouseOut=""this.style.color='FFFFFF'"" name=""button223""><img src='f2/dpk/2.GIF'><input type=button value='拒绝' onclick=""javascript:window.open('f2/dpkok.asp?id="&myid&"&from1="&from1 & "&to1="&to1&"&yn=0','a','width=380,height=210');this.disabled=1"" style=""background-color:#86A231;color:FFFFFF;border: 1 double"" onMouseOver=""this.style.color='FFFF00'"" onMouseOut=""this.style.color='FFFFFF'"" name=""button223"">"

sql="insert into dpk(ufrom,uto,duz) values ('"& nickname & "$下注','"& to1 & "$下注', " & xia_fir & xiamoney & ")"
Set Rs=conn.Execute(sql)
conn.close
set rs=nothing
set conn=nothing
end function
%>
