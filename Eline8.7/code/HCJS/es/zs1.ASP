<%@ LANGUAGE=VBScript%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
sjjh_id=Session("sjjh_id")
sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
if sjjh_name="" then Response.Redirect "../../error.asp?id=440"
http = Request.ServerVariables("HTTP_REFERER") 
if InStr(http,"hcjs/es")=0 then 
	Response.Write "<script language=javascript>{alert('对不起，程序拒绝您的操作！！！\n     按确定返回！');parent.history.go(-1);}</script>" 
	Response.End 
end if
sl=abs(int(request.querystring("sl")))
if InStr(id,"oR")<>0 or InStr(id,"Or")<>0 or InStr(id,"OR")<>0 or InStr(id,"or")<>0 or InStr(id,"=")<>0 or InStr(id,"`")<>0 or InStr(id,"'")<>0 or InStr(id," ")<>0 or InStr(id,"　")<>0 or InStr(id,"'")<>0 or InStr(id,chr(34))<>0 or InStr(id,"\")<>0 or InStr(id,",")<>0 or InStr(id,"<")<>0 or InStr(id,">")<>0 then Response.Redirect "../../error.asp?id=54"
id=trim(request.querystring("id"))
if not isnumeric(id) then Response.Redirect "../../error.asp?id=54"
if Session("sjjh_inthechat")<>"1" then 
	Response.Write "<script language=JavaScript>{alert('你不能进行操作，进行此操作必须进入聊天室！');window.close();}</script>"
	Response.End 
end if
sj=DateDiff("s",Application("sjjh_czsj"),now())
if sj<8 then
	s=8-sj
	Response.Write "<script language=JavaScript>{alert('防止刷物品，请您等上["&s&"秒钟]再操作！');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End
end if	
Application.Lock
Application("sjjh_czsj")=now()
Application.UnLock
myname=sjjh_name
name=LCase(trim(Request.form("name")))
zy=LCase(trim(Request.form("zy")))
zy=replace(zy,"'","")
zy=replace(zy,chr(34),"")
zy=Replace(zy,"<","")
zy=Replace(zy,">","")
zy=Replace(zy,"\x3c","")
zy=Replace(zy,"\x3e","")
zy=Replace(zy,"\074","")
zy=Replace(zy,"\74","")
zy=Replace(zy,"\75","")
zy=Replace(zy,"\76","")
zy=Replace(zy,"&lt","")
zy=Replace(zy,"&gt","")
zy=Replace(zy,"\076","")
badstr="射精|奸|去死|吃屎|你妈|你娘|日你|尻|操你|干死你|王八|逼|傻B|贱人|狗娘|婊子|表子|靠你|叉你|叉死|插你|插死|干你|干死|日死|鸡巴|睾丸|死去 |爬你达来蛋|撅你达来蛋|死你达来蛋|包皮|龟头|屄|赑|妣|肏|奶子|尻|屌|作爱|做爱|床上|抱抱|鸡八|处女|打炮|十八摸|你爷|你爸|我儿|操你|妈|逼|asp|com|net|www|xajh|202|61|jh|江湖|or|261|网管|掌门"
bad=split(badstr,"|")
for i=0 to UBound(bad)
zy=Replace(zy,bad(i),"**")
next
if instr(zy,"or")<>0 or instr(zy,"'")<>0 or instr(name,"or")<>0 or instr(name,"'")<>0 then Response.Redirect "../../error.asp?id=54"
if name="" or zy="" or name=myname then
	Response.Write "<script Language=Javascript>alert('提示：赠送不成功，赠人与自己姓名不能相同，姓名与赠言不能为空！');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "select 性别 from 用户 where id="&sjjh_id,conn
sex=rs("性别")
rs.close
rs.open "select 性别 from 用户 where 姓名='"&name &"'",conn
if rs.eof or rs.bof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：赠送不成功，所输入姓名:"& name &"找不到，请重新输入！');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
if sex=rs("性别") then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：赠送不成功，礼品赠送仅限于异性之间！');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
rs.close
rs.open "select * from w where id=" & id & " and c<>'卡片' and i>="&sl&" and b="&sjjh_id,conn
if rs.eof or rs.bof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：赠送不成功，你没有这样的物品！');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
wpname=rs("a")
lx=rs("c")
nl=rs("e")
tl=rs("f")
fy=rs("h")
gj=rs("g")
yin=rs("l")
esj=0
conn.execute "update w set i=i-"& sl &" where id=" & id
conn.execute "insert into s (a,b,c,d,e,f,g,h,i,j,k,l,m,n) values ('"&wpname&"',"&sjjh_id&",'赠送','"& name &"','"&lx& "',"& nl &","& tl &","& gj &","& fy &","& sl &",'"& zy &"',"& esj &",now(),"& yin &")"
rs.close
set rs=nothing
conn.close
set conn=nothing
Response.Write "<script Language=Javascript>alert('提示："& myname &"赠送"& name &"物品："& wpname & sl &"提交完成！');location.href = 'wupin.asp';</script>"
%>

