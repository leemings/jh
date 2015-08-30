<%@ LANGUAGE=VBScript%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
if session("sjjh_name")="" then Response.Redirect "../error.asp?id=440"
if Instr(LCase(Application("sjjh_useronlinename"&session("nowinroom")))," "&LCase(sjjh_name)&" ")=0 then
	Response.Write "<script Language=Javascript>alert('提示：想给朋友点歌请先进入江湖聊天室！');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
id=request("id")
if id="" then
response.write "参数不足"
response.end
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "SELECT 银两,金币 FROM 用户 WHERE 姓名='"& session("sjjh_name") &"'",conn,1,3
if rs("银两")<1000000 or rs("金币")<2 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	response.write "身上没1000000银两或金币不足2个，不能给朋友点歌！"
	response.end
end if
rs("银两")=rs("银两")-1000000
rs("金币")=rs("金币")-2
rs.update
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
<%Option Explicit%>
<%
dim webname,weburl,forumurl
webname="一线网络"
weburl="http://wWw.happyjh.com" 
forumurl="http://wWw.happyjh.com"
%>
<%
dim name,mail,cname,cmail,dmail,email,fmail,heyan,biaote
dim objRS,sql,td,mymail,bodys

biaote      = replace(Request.Form("biaote"),"""","“")
biaote  	= replace(biaote,"&","＆")
name        = replace(Request.Form("name"),"""","“")
name		= replace(name,"&","＆")
mail        = replace(Request.Form("mail"),"""","“")
mail		= replace(mail,"&","＆")
cname       = replace(Request.Form("cname"),"""","“")
cname		= replace(cname,"&","＆")
cmail       = replace(Request.Form("cmail"),"""","“")
cmail		= replace(cmail,"&","＆")
dmail       = replace(Request.Form("dmail"),"""","“")
dmail		= replace(dmail,"&","＆")
email       = replace(Request.Form("email"),"""","“")
email		= replace(email,"&","＆")
fmail       = replace(Request.Form("fmail"),"""","“")
fmail		= replace(fmail,"&","＆")
heyan       = replace(Request.Form("heyan"),"""","“")
heyan           = replace(heyan,"&","＆")

if Trim(name) = "" or Trim(mail) = "" or Trim(cname) = "" or Trim(cmail) = "" or Trim(dmail) = "" or Trim(email) = ""or Trim(fmail) = "" then
    Response.Write("<script language=""JavaScript"">alert(""还没填完哦!^_^"");history.go(-1);</script>")
    Response.End
	end if
	
	
	on error resume next
			set mymail=server.createobject("cdonts.newmail")
			mymail.from=cmail
			mymail.to=mail
			 mymail.subject="您的朋友 "&cname&" 送给您一首歌!"
				bodys=      ""&chr(10)
				bodys=bodys&chr(10)
				bodys=bodys&"Hi! "&name&" "&chr(10)
				bodys=bodys&chr(10)
				bodys=bodys&"您的朋友 "&cname&" 为您点播了 "&email&" "&chr(10)
				bodys=bodys&chr(10)
				bodys=bodys&"请点击下面的地址收听这首歌："&chr(10)
				bodys=bodys&chr(10)
				bodys=bodys&" "&fmail&" "&chr(10)
				bodys=bodys&chr(10)
				bodys=bodys&" "&dmail&" "&chr(10)
				bodys=bodys&chr(10)
				bodys=bodys&"一线网络愿您天天都有好心情！"&chr(10)
				bodys=bodys&chr(10)
				bodys=bodys&"------------------------------------------------"&chr(10)
				bodys=bodys&"欢迎您光临一线视听：http://wWw.happyjh.com"&chr(10)
				bodys=bodys&"------------------------------------------------"&chr(10)
				mymail.body=bodys
				mymail.send
				%>
				<title>发送完成</title>
<br><br><p align=center>发送完成<br><br><a href=# onclick="javascript:window.close()">关闭本窗口</a></p>