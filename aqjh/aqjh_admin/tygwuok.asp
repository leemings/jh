<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.charset="gb2312"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
if session("aqjh_adminok")<>true then Response.Redirect "../chat/readonly/bomb.htm"
if aqjh_grade<>10 or instr(Application("aqjh_admin"),aqjh_name)=0  then Response.Redirect "../error.asp?id=439"
a=request("a")
a0=trim(request.form("a"))
a1=trim(request.form("a1"))
b=trim(request.form("b"))
c=trim(request.form("c"))
d=trim(request.form("d"))
e=trim(request.form("e"))
f=trim(request.form("f"))
id=request("id")
num=trim(Request.Form("select"))
select case num
case "aa"
num="*2"
case "bb"
num="*4"
case "cc"
num="*6"
case "dd"
num="*8"
case "ee"
num="/2"
case "ff"
num="/4"
case "gg"
num="/6"
case "hh"
num="/8"
case else
%>
<script language="vbscript">
alert("������Χ!")
history.back(1)
</script>
<%
end select

Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
if a="m" then
rs.open "SELECT id from ������ where id="&id,conn
if rs.bof or rs.eof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	response.write "�Բ���û�иù��"
	response.end
end if

	nameid=int(abs(request("id")))
	conn.Execute "update ������ set  ����='" & a0 & "',����=����" & num & ", ����=����" & num & ",����=����" & num & " where id="&nameid
end if
conn.execute "insert into l(a,b,c,d,e) values (now(),'"& aqjh_name &"','"& Request.ServerVariables("REMOTE_ADDR") &"','�����¼','���������޸�')"
set rs=nothing
conn.close
set conn=nothing
Response.Write "<script Language=Javascript>alert('�����ɹ���');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
%>

</body>
</html>
