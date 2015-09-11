<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../../error.asp?id=440"
name=trim(Request.form("name"))
mess=trim(Request.form("mess"))
mess=LCase(trim(Request.form("mess")))
mess=replace(mess,"'","")
mess=replace(mess,chr(34),"")
mess=Replace(mess,"<","")
mess=Replace(mess,">","")
mess=Replace(mess,"\x3c","")
mess=Replace(mess,"\x3e","")
mess=Replace(mess,"\074","")
mess=Replace(mess,"\74","")
mess=Replace(mess,"\75","")
mess=Replace(mess,"\76","")
mess=Replace(mess,"&lt","")
mess=Replace(mess,"&gt","")
mess=Replace(mess,"\076","")
badstr="射精|奸|去死|吃屎|你妈|你娘|日你|尻|操你|干死你|王八|逼|傻B|贱人|狗娘|婊子|表子|靠你|叉你|叉死|插你|插死|干你|干死|日死|鸡巴|睾丸|死去 |爬你达来蛋|撅你达来蛋|死你达来蛋|包皮|龟头|屄|赑|妣|肏|奶子|尻|屌|作爱|做爱|床上|抱抱|鸡八|处女|打炮|十八摸|你爷|你爸|我儿|操你|妈|逼|asp|com|net|www|xajh|202|61|jh|江湖|or|261|网管|掌门"
bad=split(badstr,"|")
for i=0 to UBound(bad)
mess=Replace(mess,bad(i),"**")
next
money=int(abs(Request.form("money")))
if name=aqjh_name or name="" or mess="" then 
	Response.Write "<script Language=Javascript>alert('提示：姓名不能为空，不能杀自己？！！');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
if money<10000 or money>100000000 then
	Response.Write "<script Language=Javascript>alert('提示：你有这么多的钱嘛？！！');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "SELECT * FROM h WHERE a='"&aqjh_name&"'",conn
If not(Rs.Bof OR Rs.Eof) Then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：你要杀的人不存在。。！');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
rs.close
'校验用户 魅力大于100，钱大于1000
rs.open "SELECT * FROM 用户 WHERE 姓名='"&aqjh_name&"'"&" and 银两>"&money,conn
If Rs.Bof OR Rs.Eof Then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：你有这么多的钱嘛？！！');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
rs.close
rs.open "SELECT * FROM e WHERE a='"&aqjh_name&"' and b='"&name&"'",conn
If not(Rs.Bof OR Rs.Eof) Then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：你已经申请过了不能重复操作！');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
	conn.execute "update 用户 set 银两=银两-"&money&" where 姓名='"&aqjh_name&"'"
	conn.Execute "INSERT INTO e (a,b,c,d,e,f) VALUES('"&aqjh_name&"','"&name&"','"&mess&"',now(),"&money&",'无')"
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
Response.Write "<script Language=Javascript>alert('提示：数据提交成功！');location.href='killer.asp';</script>"
response.end
%>