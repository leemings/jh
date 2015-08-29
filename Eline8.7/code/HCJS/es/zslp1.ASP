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

id=trim(request.querystring("id"))
if not isnumeric(id) then Response.Redirect "../../error.asp?id=54"
myname=sjjh_name
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "select * from s where id=" & id & " and j>0",conn
if rs.eof or rs.bof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：情意礼品屋并没有这种物品！');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
duixiang=rs("d")
wpname=rs("a")
lx=rs("e")
nl=rs("f")
tl=rs("g")
fy=rs("h")
gj=rs("g")
yin=rs("n")
sl=rs("j")
esmoney=rs("l")
toname=rs("b")
if duixiang<>myname then
	conn.close
	set rs=nothing
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：呵。。笑话，想作什么呀，这个也不是送给你的！');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
conn.execute "delete * from s where  id="&id
rs.close
rs.open "select id from w where a='" & wpname & "' and b="&sjjh_id
If Rs.Bof OR Rs.Eof then
	conn.execute "insert into w (a,b,c,e,f,g,h,i,l) values ('"&wpname&"',"&sjjh_id&",'"&lx& "',"& nl &","& tl &","& gj &","& fy &","& sl &","& yin &")"
else
	id=rs("id")
	conn.execute "update w set l=" & yin & ",i=i+" & sl & ",e="& nl &",f="& tl &",g="& gj &",h="& fy &" where id="&id
end if
rs.close
set rs=nothing
conn.close
set conn=nothing
Response.Write "<script Language=Javascript>alert('提示："& myname &"在情意礼品屋领取["& wpname &"]成功！');location.href = 'zslp.asp';</script>"
%>


