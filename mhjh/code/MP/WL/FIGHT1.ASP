<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
my=session("yx8_mhjh_username")
types=Request.QueryString("typename")
id=Request.QueryString("id")
Select Case id
Case "齐"
  classes="齐国"
Case "楚" 
   classes="楚国"
Case "燕"
   classes="燕国"
Case "韩"
   classes="韩国"
Case "赵"
   classes="赵国"
Case "魏"
   classes="魏国"
Case "秦"
   classes="秦国"
End Select
Select Case types
Case "gold"
  typeses="君主"
Case "silver" 
   typeses="丞相"
Case "copper"
   typeses="元帅"
End Select
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
connstr=Application("yx8_mhjh_connstr")
conn.open connstr
sql="select * from 用户 where 姓名='" & my & "' and 身份='掌门'"
set rs=conn.execute(sql)
if rs.eof or rs.bof then
	Response.Write "<script Language=Javascript>alert('失败：只有本国的掌门才有资格来本国争夺职位！你损失很大');</script>"
	Response.Write "<script Language=Javascript>location.href ='javascript:window.close()';</script>"
	conn.close
	Response.end
end if
if rs("国家")<>id then
	Response.Write "<script Language=Javascript>alert('错误：别搞错了,你是哪国人！连自己的祖国都忘记了?罚你一些状态!!');</script>"
	Response.Write "<script Language=Javascript>location.href ='javascript:window.close()';</script>"
	conn.close
	Response.end
end if
if rs("积分")<50000 then
	Response.Write "<script Language=Javascript>alert('错误：你的级别还这么低,来这里只能白白送死,快回去练练吧');</script>"
	Response.Write "<script Language=Javascript>location.href ='javascript:window.close()';</script>"
	conn.close
	Response.end
end if
if rs("银两")<500000 then
	Response.Write "<script Language=Javascript>alert('穷鬼：你连50万银子都没有,还想统治一个国家啊?去死吧');</script>"
	Response.Write "<script Language=Javascript>location.href ='javascript:window.close()';</script>"
	conn.close
	Response.end
end if
if rs("内力")<50000 then
	Response.Write "<script Language=Javascript>alert('错误：你内力不足,这样胡来简直是找死,去死吧');</script>"
	Response.Write "<script Language=Javascript>location.href ='javascript:window.close()';</script>"
	conn.close
	Response.end
end if
if rs("门派")=" " or rs("门派")="无" then
	Response.Write "<script Language=Javascript>alert('错误：你无门无派,无法参与武林大会职位争夺!魔幻之神会惩罚你!');</script>"
	Response.Write "<script Language=Javascript>location.href ='javascript:window.close()';</script>"
	conn.close
	Response.end
end if
connstr1="DBQ="+server.mappath("fight.mdb")+";DefaultDir=;DRIVER={Microsoft Access Driver (*.mdb)};"
Set conn1=Server.CreateObject("ADODB.CONNECTION")
  conn1.open connstr1
set rs1=server.createobject("adodb.recordset")
sql1="select * from "&types&" where 国家='"&id&"'"
set rs1=conn1.execute(sql1)
xm=rs1("姓名")
if rs("职位")="君主" then
%>
<script language=vbscript>  
  MsgBox "你已经是最高级别的君主了,还想往哪爬?"  
  location.href = "meeting.asp"  
</script>
<%rs.close
rs1.close
conn.close
conn1.close
response.end
end if
if rs("职位")="丞相" and types="silver" then
%>
<script language=vbscript>  
  MsgBox "你已经取得这个职位了,无法争夺此职位！"  
  location.href = "meeting.asp"  
</script>
<%rs.close
rs1.close
conn.close
conn1.close
response.end
end if
if rs("职位")="元帅" and types="copper" then
%>
<script language=vbscript>  
  MsgBox "你已经夺得这个职位,别和自己过不去。"  
  location.href = "meeting.asp"  
</script>
<%rs.close
rs1.close
conn.close
conn1.close
response.end
end if
if rs("职位")="丞相" and types="copper" then
%>
<script language=vbscript>  
  MsgBox "你的职位比这个高,无法争夺低职位！"  
  location.href = "meeting.asp"  
</script>
<%rs.close
rs1.close
rs2.close
conn.close
conn1.close
response.end
end if
if rs("防御")>rs1("防御") and rs("攻击")+rs("内力")>rs1("攻击")+rs1("内力") then
sql1="update "&types&" set 姓名='"& my & "',职位='"&rs("职位")&"',门派='"&rs("门派")&"',身份='"&rs("身份")&"',内力="&rs("内力")&",体力="&rs("体力")&",攻击="&rs("攻击")&",防御="&rs("防御")&" where 国家='"&id&"'"
conn1.execute(sql1)
sql="update 用户 set 职位='"& typeses &"',积分=积分+2000, 银两=银两+500000 where 姓名='"&my&"'"
conn.execute(sql)
sql="update 用户 set 职位='无',积分=积分-2000, 银两=银两-500000 where 姓名='"&xm&"'"
conn.execute(sql)
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
		newtalkarr(599)="<font color=red>【消息】</font><font color=#9900ff>恭喜！！<font color=#ff0000>"& my &"</font>参加魔幻江湖的武林大会，战胜"& xm &"，成功登上了<font color=#ff0000>"& classes & typeses&"</font>的宝座！并夺得"& xm &"的积分2000点！实在是威风八面！</font>" 
newtalkarr(600)=Session("yx8_mhjh_userchatroomsn")
Application.Lock
Application("yx8_mhjh_talkpoint")=talkpoint+1
Application("yx8_mhjh_talkarr")=newtalkarr
Application.UnLock
erase newtalkarr
erase talkarr

 %>
<script language=vbscript>  
  MsgBox "恭喜你，你揭榜成功！获得积分2000点！"  
  location.href = "meeting.asp"  
</script>
 <%else
sql="update 用户 set  内力=内力-50000,体力=体力-500000,银两=银两-500000 where 姓名='"&my&"'"
conn.execute(sql)
if rs("体力")<0 then
	conn.execute "update 用户 set 状态='死亡',被杀=被杀+1 where 姓名='"&my&"'"
conn.execute "insert into 英烈堂(时间,死者,凶手,死因) values(now(),'"&my&"','战死','争夺国家职位失手')"
	Response.Write "<script Language=Javascript>alert('穷鬼：你连50万银子都没有,还想统治一个国家啊?');</script>"
	Response.Write "<script Language=Javascript>location.href ='javascript:window.close()';</script>"
	session.Abandon
	conn.close
	Response.end
end if
rs.close
rs1.close
conn.close
conn1.close
%>
<script language=vbscript>  
  MsgBox "对不起！你战斗失败！结果被打得鼻青脸肿的，损失内力50000，体力500000，丢失银子500000。"  
  location.href = "meeting.asp"  
</script>
<% end if
%>

