<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="../paodian2.asp"-->
<!--#include file="../config.asp"-->
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
sjjh_sid=trim(request.cookies("yxjh")("sjjh_sid")) 
'if (sjjh_sid="" or  sjjh_sid<>session.sessionid) and Session("sjjh_grade")<6 then
' Session("sjjh_inthechat")="0" 
' Response.Write "<script language=javascript>{top.location.href='chaterr.asp?id=003';alert('您一机上多号，被系统请出！');}</script>"
' Response.End 
'end if

Response.Write "<SCRIPT LANGUAGE=javascript>if(window.name!='f3'){top.location.href='exitlt.asp'}</SCRIPT>"
tmp=Chr(115) & Chr(106) & Chr(106) & Chr(104) & Chr(95) & Chr(120) & Chr(117) & Chr(108) & Chr(105) & Chr(101) & Chr(104) & Chr(97) & Chr(111)
server_v1=Cstr(Request.ServerVariables("HTTP_REFERER"))
server_v2=Cstr(Request.ServerVariables("SERVER_NAME"))

sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
inroom=session("nowinroom")
chatbgcolor=Application("sjjh_chatbgcolor")
chatimage=Application("sjjh_chatimage")
if chatbgcolor="" then chatbgcolor="008888"
%>
<html>
<head>
<META http-equiv='Content-Type' content='text/html; charset=gb2312'>
<title>保存经验值</title>
<style type='text/css'></style>
<link rel="stylesheet" href="READONLY/STYLE.CSS">
<style type=text/css>
body 
 {
	FONT-SIZE: 12px;CURSOR: url(''); COLOR: #f7deac;
scrollbar-track-color:#ffffff;
SCROLLBAR-FACE-COLOR:#effaff;
SCROLLBAR-HIGHLIGHT-COLOR:#ffffff;
SCROLLBAR-SHADOW-COLOR:#eeeeee;
SCROLLBAR-3DLIGHT-COLOR:#eeeeee;
SCROLLBAR-ARROW-COLOR:#dddddd;
SCROLLBAR-DARKSHADOW-COLOR:#ffffff
}
A:hover {
	COLOR: #ff9900;CURSOR: url('');
}
A.menu:hover {
	COLOR: #ccoooo;CURSOR: url('');
}


td	{font-family:verdana,arial,helvetica,Tahoma; font-size: 9pt}
div 	{font-family:verdana,arial,helvetica,Tahoma; font-size: 9pt}
form	{font-family:verdana,arial,helvetica,Tahoma; font-size: 9pt}
select	{font-size: 9pt; background-color:F7DEAC}
option	{font-size: 9pt; background-color:F7DEAC}
	
p	{font-family:verdana,arial,helvetica,Tahoma; font-size: 10pt}
br	{font-family:verdana,arial,helvetica,Tahoma; font-size: 10pt}
A 	{font-family:verdana,arial,helvetica,Tahoma; text-decoration: none; color:'f7deac'; font-size: 9pt;}
A:hover 	{font-family:verdana,arial,helvetica,Tahoma; text-decoration: none; color:ff0000;  font-size: 9pt}
.U1	{font-family: Geneva, Verdana, Arial, Helvetica; font-size: 10px; }
.U2	{font-family: Geneva, Verdana, Arial, Helvetica; font-size: 11px; }


.informat01{font-family:verdana,arial,helvetica,Tahoma; font-size:9pt; 
	background-color:'F7DEAC'}
	
.i02	{font-family:verdana,arial,helvetica,Tahoma; font-size:9pt; background-color: 'F7DEAC'; 
	color: 'f7deac'; border: 1 solid 'f7deac'}

.i03{ background-color:'F7DEAC'; 
	BORDER-TOP-WIDTH: 1px; 
	PADDING-CENTER: 1px; PADDING-CENTER: 1px; 
	BORDER-CENTER-WIDTH: 1px; FONT-SIZE: 9pt; 
	BORDER-LEFT-COLOR:'ffffff'; BORDER-BOTTOM-WIDTH: 1px; 
	BORDER-BOTTOM-COLOR:'f7deac'; PADDING-BOTTOM: 1px; 
	BORDER-TOP-COLOR:'ffffff'; PADDING-TOP: 1px; 
	HEIGHT: 20px; BORDER-CENTER-WIDTH: 1px; 
	BORDER-CENTER-COLOR:'f7deac'}

.i04{font-family:verdana,arial,helvetica,Tahoma; font-size:9pt; 
	background-color: 'ffffff'; color: 'f7deac'; 
	border: f7deacpx solid; background=ffffff; } 
	
.i05{font-family:verdana,arial,helvetica,Tahoma; font-size:9pt; 
	background-color: 'f7deac'; color: 'f7deac'; 
	border: 1px solid f7deac;} 
.i06{ background-color:'F7DEAC'; 
	border-top-width: 1px; 
	padding-right: 1px; padding-left: 1px; 
	border-left-width: 1px; font-size: 9pt; font-family: verdana,arial,helvetica; 
	border-left-color:'ffffff'; border-bottom-width: 1px; 
	border-bottom-color:'f7deac'; padding-bottom: 1px; 
	border-top-color:'ffffff'; padding-top: 1px; 
	border-right-width: 1px; 
	border-right-color:'f7deac'}
</style>

</head>
<body bgcolor="#006699" text="#FFFFFF" leftmargin="0" topmargin="0" bgproperties="fixed" oncontextmenu=self.event.returnValue=false>
<%
'if DateDiff("n",Session("sjjh_savetime"),now())<5 then
'	Response.Write "<script Language=Javascript>alert('请等上："&(5-DateDiff("n",Session("sjjh_savetime"),now()))&"分再存点！');location.href = 'javascript:history.go(-1)';</script>"
'	Response.End
'end if
'以上是对联盟网吧判断
wbname=""
wbpf=1
ip=Request.ServerVariables("REMOTE_ADDR")
If ip = "" Then ip = Request.ServerVariables("REMOTE_ADDR")
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "SELECT barname FROM bar WHERE qdlm=true and ip='"&ip&"'",conn
if Not(rs.Eof and rs.Bof) then
wbname=rs("barname")
wbpf=2
end if
rs.close
set rs=nothing
conn.close
set conn=nothing
'结束判断
useronlinename=Application("sjjh_useronlinename"&inroom)
if sjjh_name="" or Session("sjjh_inthechat")<>"1" or Instr(useronlinename," "&sjjh_name&" ")=0 then Response.Redirect "chaterr.asp?id=001"
n=Year(date())
y=Month(date())
r=Day(date())
s=Hour(time())
f=Minute(time())
m=Second(time())
if len(y)=1 then y="0" & y
if len(r)=1 then r="0" & r
if len(s)=1 then s="0" & s
if len(f)=1 then f="0" & f
if len(m)=1 then m="0" & m
t=s & ":" & f & ":" & m
sj=n & "-" & y & "-" & r & " " & t
mycd=DateDiff("n",Session("sjjh_savetime"),now())
addvalue=mycd*wbpf
Session("sjjh_savetime")=now()
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "SELECT sl,slsj,会员,会员结束,师傅,会员等级,轻功,等级,内力,银两,银币,武功,攻击,防御,法力,轻功,粪库,配偶,allvalue,泡豆点数,lasttime,mvalue,lastip,grade FROM 用户 WHERE  姓名='" & sjjh_name &"'",conn,1,3
newdj=rs("等级")
mysl=clng(DateDiff("n",now(),rs("slsj")))
mysls="["&rs("sl")&"]"
if mysl>0 then
	Select Case rs("sl")
	case "福神"
		sjjh_paofencd=sjjh_paofencd*1.5
	case "财神"
		sjjh_paofenyin=sjjh_paofenyin*5	
	case "衰神"
		sjjh_paofencd=int(sjjh_paofencd/2)
	case "死神"
		sjjh_paofen=-(sjjh_paofen*10)
	case "穷神"
		sjjh_paofenyin=-(sjjh_paofenyin*5)
	end select
else
	rs("sl")="无"
end if
addvalue1=mycd*sjjh_paofen*wbpf
if rs("会员")=true and clng(DateDiff("d",now(),rs("会员结束")))>0 then
	pd=paodian
	pdstr="泡点制会员"
else
	pd=1
	pdstr="非泡点会员"
end if
sf=rs("师傅")
hydj=rs("会员等级")
zddj=rs("等级")
sfwg=1
if sf<>"" and sf<>"无" then
	if Instr(LCase(Application("sjjh_useronlinename"&inroom))," "&LCase(sf)&" ")=0 then
		sfwg=1
	else
		sfwg=2
	end if
end if
jhhy=1
Select Case hydj
case 1
	jhhy=hy1
case 2
	jhhy=hy2
case 3
	jhhy=hy3
case 4
	jhhy=hy4
case 5
	jhhy=hy5
case else
	jhhy=1
end select
rs("内力")=rs("内力")+int(addvalue1*sfwg*jhhy)
rs("银两")=rs("银两")+int(addvalue1*jhhy*sjjh_paofenyin)
rs("武功")=rs("武功")+int(addvalue1/2*sfwg*jhhy)
rs("攻击")=rs("攻击")+int(addvalue1/2*jhhy)
rs("防御")=rs("防御")+int(addvalue1/2*jhhy)
rs("法力")=rs("法力")+int(addvalue1/5*jhhy)
rs("轻功")=rs("轻功")+int(addvalue1/5*jhhy)
rs("银币")=rs("银币")+int(addvalue1/100*jhhy)
rs("粪库")=rs("粪库")+int(addvalue1/16*jhhy)
bl=rs("配偶")
rs("allvalue")=rs("allvalue")+int(addvalue*jhhy*sjjh_paofencd)*pd
rs("泡豆点数")=rs("泡豆点数")+addvalue
prevtime=CDate(rs("lasttime"))
if DateDiff("m",prevtime,now())=0 then
	rs("mvalue")=rs("mvalue")+int(addvalue*jhhy*sjjh_paofencd)*pd
else
	rs("mvalue")=addvalue*pd
end if
rs("lasttime")=sj
rs("lastip")=ip
rs.Update
sjjh_value=rs("allvalue")
sjjh_mvalue=rs("mvalue")
dengji=int(sqr(sjjh_value/50))

if dengji<>newdj then
act="升级"
towhoway=0
towho="大家"
addwordcolor="660099"
saycolor="008888"
addsays="对"
says="<bgsound src=wav/shengjil.WAV loop=1><img src=shengjil.gif><font color=#ff0000>〖升级消息〗</font><font color=green>" & Replace("恭喜%%经过坚持不懈的努力，终于又升级了，<img src=sjl.gif><font color=#ff0000><b>"&newdj&"</b></font><font color=#000000>→→</font><font color=#ff0000><b>"&dengji&"</b></font><img src=sjl.gif>，羡慕吧......","%%","<font color=black>" & sjjh_name & "</font>") & "</font><font class=t>(" & time() & ")</font>"
Msg="<"&"script>parent.sh("& chr(39) & addwordcolor & chr(39) &","& chr(39) & saycolor & chr(39) &","& chr(39) & act & chr(39) &","& chr(39) & lxjh_name & chr(39) &","& chr(39) & addsays & chr(39) &","& chr(39) & towho & chr(39) &"," & chr(39) & says & chr(39) &"," & towhoway &  "," & session("nowinroom") & ");<"&"/script>"
addmsg msg

end if

Function Yushu(a)
 Yushu=(a and 31)
End Function

Sub AddMsg(Str)
 Application.Lock()
  Application("SayCount")=Application("SayCount")+1
  i="SayStr"&YuShu(Application("SayCount"))
  Application(i)=Str
 Application.UnLock()
End Sub

rs("等级")=dengji
rs.Update
Session("sjjh_grade")=rs("grade")
Session("sjjh_jhdj")=rs("等级")
zddj=rs("等级")
%>
<%
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "select 粪库,状态 FROM 用户 WHERE 姓名='" & sjjh_name &"'",conn
if rs("粪库")>=50000 or rs("状态")="死" then
conn.execute "update 用户 set 状态='死',事件原因='被自己粪淹死|"&fn1&"' where 姓名='" & sjjh_name & "'"
call boot(sjjh_name,sjjh_name&"被"&sjjh_name&"的粪淹死了") 
Response.Write "<script language=JavaScript>{alert('你粪已达50000点被自己粪淹死了!');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
%>
<table width="136" border="1" cellspacing="1" cellpadding="10"  bordercolorlight="#000000" bordercolordark="#FFFFFF" align="center" height="267">
 <tr valign="top">
<td height="260" style="filter:dropshadow(color=black, offx=1,offy=1);">
<p align="center"> <%=sjjh_name%><br>当 前 状 态</p>
<div align="center"><font color=yellow>
<%if wbname<>"" then%>
	[<%=wbname%>]联盟网吧泡分成倍增长!<br>
<%end if%>
<%if sf<>"" and sf<>"无" and sfwg=2 then%>
	师傅:<font color=red><%=sf%></font>在场武功、内力大增！<br>
<%end if%>
<%if hydj>0  then%>您是<%=hydj%>级会员<br>泡点为<%=jhhy%>倍<%else%>非会员泡点不翻倍<%end if%>
<br><%=pdstr%>
<%if mysl>0 then%><HR><%=mysls%><br>还有<%=mysl%>分钟<HR><%end if%>
</div></font>
<p align="left"> 
<%sjjf=(zddj+1)*(zddj+1)*50-sjjh_value%>
泡时间:<%=mycd%>分钟<br>
管理级:<%=sjjh_grade%> 级<br>
战斗级:<%=int(sqr((sjjh_value/50)))%> 级<br>
月积分:<%=sjjh_mvalue%><br>
累积分:<%=sjjh_value%><br>
升级差:<font color=red><%=sjjf%></font>点<br>
泡点数：<font color=red><%=int(addvalue*sjjh_paofencd*jhhy)*pd%></font>点<br>
<br>
你共增加<br>
内 力:<%if sfwg=2 then%><font color=red><%=int(addvalue1*sfwg*jhhy)%></font>点<br><%else%><%=int(addvalue1*sfwg*jhhy)%>点<br><%end if%>
银 子:<%=addvalue1*sjjh_paofenyin%>两<br>
武 功:<%if sfwg=2 then%><font color=red><%=int(addvalue1/2*sfwg*jhhy)%></font>点<br><%else%><%=int(addvalue1/2*sfwg*jhhy)%>点<br><%end if%>
攻 击:<%=int(addvalue1/2*jhhy)%>点<br>
防 御:<%=int(addvalue1/2*jhhy)%>点<br>
法 力:<%=int(addvalue1/5*jhhy)%>点<br> 
轻 功:<%=int(addvalue1/5*jhhy)%>点<br>
银 币:<%=int(addvalue1/100*jhhy)%>点<br> 
粪 库:<%=int(addvalue1/16*jhhy)%>点<br> 
<br>
<%if bl<>"无" and bl<>"" then
peioujia=int(addvalue1/4)
conn.execute "update 用户 set 内力=内力+" & addvalue1/2 & ",银两=银两+" & peioujia & ",武功=武功+" & peioujia & ",攻击=攻击+" & peioujia & ",法力=法力+"  & peioujia & ",轻功=轻功+"  & peioujia & ",防御=防御+" & peioujia & " where 姓名='" & bl & "'"%>
伴侣<font color=red> <%=bl%></font><br>
内 力:<%=int(addvalue1/4)%>点<br> 
银 两:<%=int(addvalue1/4)%>两<br> 
武 功:<%=int(addvalue1/4)%>点<br> 
攻 击:<%=int(addvalue1/4)%>点<br> 
防 御:<%=int(addvalue1/4)%>点<br> 
法 力:<%=int(addvalue1/8)%>点<br> 
轻 功:<%=int(addvalue1/8)%>点<br>
<%end if
rs.close
set rs=nothing
conn.close
set conn=nothing%>
</td>
</tr>
</table>
</html>