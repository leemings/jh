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
	Response.Write "<script Language=Javascript>alert('��ʾ��������ѵ�����Ƚ��뽭�������ң�');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
id=request("id")
if id="" then
response.write "��������"
response.end
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "SELECT ����,��� FROM �û� WHERE ����='"& session("sjjh_name") &"'",conn,1,3
if rs("����")<1000000 or rs("���")<2 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	response.write "����û1000000�������Ҳ���2�������ܸ����ѵ�裡"
	response.end
end if
rs("����")=rs("����")-1000000
rs("���")=rs("���")-2
rs.update
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
<%Option Explicit%>
<%
dim webname,weburl,forumurl
webname="һ������"
weburl="http://wWw.happyjh.com" 
forumurl="http://wWw.happyjh.com"
%>
<%
dim name,mail,cname,cmail,dmail,email,fmail,heyan,biaote
dim objRS,sql,td,mymail,bodys

biaote      = replace(Request.Form("biaote"),"""","��")
biaote  	= replace(biaote,"&","��")
name        = replace(Request.Form("name"),"""","��")
name		= replace(name,"&","��")
mail        = replace(Request.Form("mail"),"""","��")
mail		= replace(mail,"&","��")
cname       = replace(Request.Form("cname"),"""","��")
cname		= replace(cname,"&","��")
cmail       = replace(Request.Form("cmail"),"""","��")
cmail		= replace(cmail,"&","��")
dmail       = replace(Request.Form("dmail"),"""","��")
dmail		= replace(dmail,"&","��")
email       = replace(Request.Form("email"),"""","��")
email		= replace(email,"&","��")
fmail       = replace(Request.Form("fmail"),"""","��")
fmail		= replace(fmail,"&","��")
heyan       = replace(Request.Form("heyan"),"""","��")
heyan           = replace(heyan,"&","��")

if Trim(name) = "" or Trim(mail) = "" or Trim(cname) = "" or Trim(cmail) = "" or Trim(dmail) = "" or Trim(email) = ""or Trim(fmail) = "" then
    Response.Write("<script language=""JavaScript"">alert(""��û����Ŷ!^_^"");history.go(-1);</script>")
    Response.End
	end if
	
	
	on error resume next
			set mymail=server.createobject("cdonts.newmail")
			mymail.from=cmail
			mymail.to=mail
			 mymail.subject="�������� "&cname&" �͸���һ�׸�!"
				bodys=      ""&chr(10)
				bodys=bodys&chr(10)
				bodys=bodys&"Hi! "&name&" "&chr(10)
				bodys=bodys&chr(10)
				bodys=bodys&"�������� "&cname&" Ϊ���㲥�� "&email&" "&chr(10)
				bodys=bodys&chr(10)
				bodys=bodys&"��������ĵ�ַ�������׸裺"&chr(10)
				bodys=bodys&chr(10)
				bodys=bodys&" "&fmail&" "&chr(10)
				bodys=bodys&chr(10)
				bodys=bodys&" "&dmail&" "&chr(10)
				bodys=bodys&chr(10)
				bodys=bodys&"һ������Ը�����춼�к����飡"&chr(10)
				bodys=bodys&chr(10)
				bodys=bodys&"------------------------------------------------"&chr(10)
				bodys=bodys&"��ӭ������һ��������http://wWw.happyjh.com"&chr(10)
				bodys=bodys&"------------------------------------------------"&chr(10)
				mymail.body=bodys
				mymail.send
				%>
				<title>�������</title>
<br><br><p align=center>�������<br><br><a href=# onclick="javascript:window.close()">�رձ�����</a></p>