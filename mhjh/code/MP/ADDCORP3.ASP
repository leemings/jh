<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
username=session("yx8_mhjh_username")
if username="" then Response.redirect "../error.asp?id=016"
corpcond=server.HTMLEncode(Trim(Request.Form("corpcond")))
if instr(corpcond,"'")<>0 or instr(corpcond,chr(34))<>0 or instr(corpcond,"?")<>0 or instr(corpcond,"=")<>0 or instr(corpcond,"or")<>0 or instr(corpcond,"/")<>0 or instr(corpcond,"#")<>0 or instr(corpcond,"&")<>0 then Response.Redirect "../error.asp?id=056"
tj=Request.Form("corptj")
corpname=server.HTMLEncode(Trim(Request.Form("corpname")))
if corpname="" or corpname="无" or corpname="十八地狱" or corpname="机器人" or corpname="天网" or corpname="出家" or instr(corpname,"'")<>0 or instr(corpname,chr(34))<>0 or instr(corpname," ")<>0 or instr(corpname,"=")<>0 or instr(corpname,"or")<>0 or instr(corpname,"?")<>0 or instr(corpname,"&#63733;")<>0 or instr(corpname,"&#")<>0 or instr(corpname,",")<>0 or instr(corpname,"=")<>0 or instr(corpname,"\")<>0 or instr(corpname,"<")<>0 or instr(corpname,">")<>0 or instr(corpname,"script:")<>0 or instr(corpname,"file:")<>0  then Response.Redirect "../error.asp?id=056"
function chuser(u)
dim filter,xx,usernameenable,su
for i=1 to len(u)
su=mid(u,i,1)
xx=asc(su)
zhengchu = -1*xx \ 256
yushu = -1*xx mod 256
if (xx>122 or (xx>57 and xx<97) or (xx<-10241 and xx>-10247) or yushu=129 or yushu>192 or (yushu<2 and yushu>-1) or (((zhengchu>1 and zhengchu<8) or (zhengchu>79 and zhengchu<86)) and yushu<96 ) or (xx>-352 and xx<48) or (xx<-22016 and xx>-24321) or (xx<-32448)) then
chuser=true
exit function
end if
next
chuser=false
end function
if chuser(corpname) then Response.Redirect "error.asp?id=056"
corpsilver=server.HTMLEncode(Trim(Request.Form("corpsilver")))
if not isnumeric(corpsilver) then Response.Redirect "../error.asp?id=024"
silver=clng(corpsilver)
if silver<0 then
silver=0
elseif silver>1000 then
silver=1000
end if
if tj="" then Response.Redirect "../error.asp?id=102"
if instr(tj,"'")<>0 or instr(tj,chr(34))<>0 or instr(tj,"?")<>0 or instr(tj,"=")<>0 or instr(tj,"or")<>0 or instr(tj,"&#63733;")<>0 or instr(tj,"&#")<>0 or instr(tj,"&")<>0 or instr(tj,"\")<>0 or  instr(tj,"script:")<>0 or instr(tj,"file:")<>0  then Response.Redirect "../error.asp?id=056"
set conn=server.CreateObject("adodb.connection")
conn.Open Application("yx8_mhjh_connstr")
set rst=server.CreateObject("Adodb.recordset")
rst.Open "select * from 用户 where 姓名='"&username&"' and 门派<>'官府' and 身份<>'掌门' and 攻击>500000 and 防御>500000 and 体力>5000000 and 等级>=20 and 内力>1000000 and 银两>5000000 and 精力>100000",conn
if rst.EOF or rst.BOF then Response.Redirect "../error.asp?id=063"
rst.Close
rst.Open "select * from 门派 where 门派='"&corpname&"'",conn,1,2
if rst.EOF or rst.BOF then
rst.AddNew
rst("门派")=corpname
rst("工资系数")=silver
rst("简介")=corpcond
rst("加入条件")=tj
rst("chaton")=false
rst.Update
conn.Execute "update 用户 set 精力=精力-100000,银两=银两-5000000,身份='掌门',门派='"&corpname&"' where 姓名='"&username&"'"
Session("yx8_mhjh_usercorp")=corpname
else
Response.Redirect "../error.asp?id=064"
end if
rst.Close
set rst=nothing
conn.Close
set conn=nothing
dim newtalkarr(600) 
   talkarr=Application("yx8_mhjh_talkarr") 
		talkpoint=Application("yx8_mhjh_talkpoint") 
		j=1 
		for i=11 to 600 step 10 
			newtalkarr(j)=talkarr(i) 
			newtalkarr(j+1)=talkarr(i+1) 
			newtalkarr(j+2)=talkarr(i+2) 
			newtalkarr(j+3)=talkarr(i+3) 
			newtalkarr(j+4)=talkarr(i+4) 
			newtalkarr(j+5)=talkarr(i+5) 
			newtalkarr(j+6)=talkarr(i+6) 
			newtalkarr(j+7)=talkarr(i+7) 
			newtalkarr(j+8)=talkarr(i+8) 
			newtalkarr(j+9)=talkarr(i+9) 
			j=j+10 
		next 
		newtalkarr(591)=talkpoint+1 
		newtalkarr(592)=1 
		newtalkarr(593)=0 
		newtalkarr(594)=username 
		newtalkarr(595)="大家" 
		newtalkarr(596)="" 
		newtalkarr(597)="000000" 
		newtalkarr(598)="000000" 
		newtalkarr(599)="<font color=red>【开山立派】</font><font color=blue>"&username&"经过许久的艰苦修炼，终于成长为一代大侠，并成功建立了自己的门派<font color=red>"&corpname&"</font>，所谓创业难,守业更难!但愿他能广招弟子,能保护好自己的基业,称霸江湖!</font>" 
newtalkarr(600)=Session("yx8_mhjh_userchatroomsn")
Application.Lock
Application("yx8_mhjh_talkpoint")=talkpoint+1
Application("yx8_mhjh_talkarr")=newtalkarr
Application.UnLock
erase newtalkarr
erase talkarr
%>
<head><link rel='stylesheet' href='../style.css'></head><body oncontextmenu=self.event.returnValue=false  background='../chatroom/bg1.gif' topmargin=100><p align=center><b><font color="ff0000" size="4">自立门派</font></b></p><hr>
恭喜恭喜，经过长期的奋斗,你终于建立了自己的门派<font color=0000ff size=6><%=corpname%></font> 精力减少100000,银两减少500万<br><p align=center><a href=javascript:location.replace('addcorp.asp');>返回</a></p>


