<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
if Instr(LCase(Application("aqjh_useronlinename"&session("nowinroom")))," "&LCase(aqjh_name)&" ")=0 then
	Response.Write "<script language=JavaScript>{alert('提示：必须进入聊天室才可以操作！');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End 
end if
id=request("id")
if not(isnumeric(id)) then
	Response.Write "<script language=JavaScript>{alert('滚！');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End
end if	
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select id,性别,配偶,银两,金币,会员金卡,保留,hw from 用户 where 姓名='"&aqjh_name&"'",conn
if rs.eof or rs.bof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('你不是江湖中人，不能租用！');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End
end if
if rs("性别")="女" and (rs("保留")<>"保留" or rs("hw")<>"") then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	response.write "<script language=javascript>{alert('你已怀孕或已有小孩，不需要再来租房，还是快回去照顾你的孩子吧…！');location.href='javascript:history.go(-1)';}</script>"
	response.end
end if
yinl=rs("银两")
jb=rs("金币")
hf=rs("会员金卡")
myid=rs("id")
poxm=rs("配偶")
rs.close
if poxm="" or poxm="无" then
	set rs=nothing
	conn.close
	set conn=nothing
	response.write "<script language=javascript>{alert('你还未结婚，本店不对未婚人氏开放！');location.href='javascript:history.go(-1)';}</script>"
	response.end
end if
rs.open "select id,性别,配偶,保留,hw from 用户 where 姓名='"&poxm&"'",conn
if rs.eof or rs.bof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('你的配偶不是江湖中人，不能租用！');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End
end if
if rs("性别")="女" and (rs("保留")<>"保留" or rs("hw")<>"") then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	response.write "<script language=javascript>{alert('你的妻子已怀孕或已有小孩，不需要再来租房，还是快回去照顾你的妻儿老小吧…！');location.href='javascript:history.go(-1)';}</script>"
	response.end
end if
poid=rs("id")
rs.close
'rs.open "select * from fw where yhid="&myid&" or yhid="&poid,conn
'if not(rs.eof or rs.bof) then
'	set rs=nothing
'	conn.close
'	set conn=nothing
'	Response.Write "<script language=JavaScript>{alert('你或你的妻子已经租上房间，房间有限，不能再租用！');location.href = 'javascript:history.go(-1)';}</script>"
'	Response.End
'end if
'rs.close
rs.open "select * from fw where id="&id,conn
if rs.eof or rs.bof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('情人小筑中并没有这个房屋！');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End
end if
yhid=rs("yhid")
if yhid<>0 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('此房间已有人租用了！');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End
end if
gmfs=rs("jgfs")
jg=rs("b")
fwm=rs("a")
rs.close
select case gmfs
	case "银两"
		if jg>yinl then
			set rs=nothing
			conn.close
			set conn=nothing
			Response.Write "<script language=JavaScript>{alert('租用此房屋一次需要："&gmfs&"："&jg&"，你的钱不够！');location.href = 'javascript:history.go(-1)';}</script>"
			Response.End
		end if
	case "金币"
		if jg>jb then
			set rs=nothing
			conn.close
			set conn=nothing
			Response.Write "<script language=JavaScript>{alert('租用此房屋一次需要："&gmfs&"："&jg&"，你的钱不够！');location.href = 'javascript:history.go(-1)';}</script>"
			Response.End
		end if
	case "会员金卡"
		if jg>hf then
			set rs=nothing
			conn.close
			set conn=nothing
			Response.Write "<script language=JavaScript>{alert('租用此房屋一次需要："&gmfs&"："&jg&"，你的钱不够！');location.href = 'javascript:history.go(-1)';}</script>"
			Response.End
		end if
end select
conn.execute "update fw set yhid="&myid&",zysj=now() where id="&id
conn.execute "update 用户 set "&gmfs&"="&gmfs&"-"&jg&" where 姓名='"&aqjh_name&"'"
set rs=nothing
conn.close
set conn=nothing
if Instr(Application("aqjh_useronlinename"&session("nowinroom"))," "&poxm&" ")<>0 then
says="<font color=red>【情人小筑】</font><font color=#8000FF>"&aqjh_name&"在快乐江湖的情人小筑为他和"&poxm&"租下了<font color=red>※"&fwm&"※</font>房间，"&aqjh_name&"已在小屋中等候，"&poxm&"快去呀，呵呵..</font>"		'聊天数据
act="消息"
towhoway=0
towho="大家"
addwordcolor="660099"
saycolor="008888"
addsays="对"
saystr="<script>parent.sh("& chr(39) & addwordcolor & chr(39) &","& chr(39) & saycolor & chr(39) &","& chr(39) & act & chr(39) &","& chr(39) & sjjh_name & chr(39) &","& chr(39) & addsays & chr(39) &","& chr(39) & towho & chr(39) &"," & chr(39) & says & chr(39) &"," & towhoway &  ","& session("nowinroom") & ");<"&"/script>"
addmsg saystr
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
end if
%>
<html>
<head>
<title>租用房屋成功</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>
<body bgcolor="#8800FF" text="#FFFF00">
<div align="center">
  <p>&nbsp;</p>
  <p>&nbsp;</p>
  <p>恭喜你成功租用了江湖情人小筑中的[<b><font color="#00FFFF"><%=fwm%></font></b>]房间，快去通知你的爱人来这里相聚吧。 
  </p>
  <p>&nbsp;</p>
  <p><a href="index.asp">返回</a> 
  </p>
</div>
</body>
</html>