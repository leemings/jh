<%@ LANGUAGE=VBScript%>
<%Response.Expires=0
nowinroom=session("nowinroom")
useronlinename=Application("aqjh_useronlinename"&nowinroom)
nickname=session("aqjh_name")
aqjh_name=Session("aqjh_name")

if aqjh_name="" or Session("aqjh_inthechat")<>"1" or Instr(useronlinename," "&aqjh_name&" ")=0 then Response.Redirect "chaterr.asp?id=001"

id=Request.QueryString ("id")
if id="" then Response.end
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
sj1=n & "-" & y & "-" & r
sj2=s & ":" & f & ":" & m
sj3=sj1 & " " & sj2
chatbgcolor=aqjh_chatbgcolor
chatimage=aqjh_chatimage
if chatbgcolor="" then chatbgcolor="008888"
useronlinename=Application("aqjh_useronlinename"&session("nowinroom"))

show=Split(Trim(useronlinename)," ",-1)
x=UBound(show)

Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
connstr=Application("aqjh_usermdb")
conn.open connstr
%>
<html>
<head>
<script language="JavaScript">function s(list){parent.f2.document.af.sytemp.value='';parent.f2.document.af.sytemp.value=parent.f2.document.af.sytemp.value+list;parent.f2.document.af.sytemp.focus();}</script>
<style></style>
<link rel="stylesheet" href="READONLY/STYLE1.CSS">
</head>
<body oncontextmenu=self.event.returnValue=false bgcolor="<%=chatbgcolor%>" background="../<%=chatimage%>" bgproperties="fixed" leftmargin="3">
<div align="center"><br>
  <font size="3"> <span style='font-size:9pt'><font size="5" face="隶书" color="#FFFFFF">转 
  送</font></span></font><br>
<%
rs.Open ("SELECT * FROM myanimal where username='"&nickname&"'and rest=0 and  id="&id),conn
if not(rs.EOF or rs.BOF) then
animalname=rs("animalname")
%>
<table cellpadding="3" cellspacing="1" border="0" align="center" width="100%">
<form method="post" action="zengmyanimalok.asp" id=form1 name=form1 target=ps>
<tr align="center"> 
        <td colspan="2"> 
          <input type="hidden" name="id" value="<%=id%>">
          <input type="text" readonly style="text-align:center;" name="animalname" size="10" maxlength="10" value="<%=animalname%>">
<br>
<font color="#FFFF00">选择好友：</font><br>
<select name="towho">
<%
for i=0 to x
	if show(i)<>aqjh_name and show(i)<>peiou then
	%>
		<option value="<%=show(i)%>"<%if CStr(show(i))=CStr(who) then Response.Write " selected"%>><%=show(i)%></option>
	<%
	end if
next
%>
</select>
      共 <font color=red><%=i%></font> 人</font><br>
<input type="submit" name="Submit" value="赠送">
</td>
</tr>
<%
end if
rs.Close
set rs=nothing
conn.close
set conn=nothing
%> 
    </form>
  </table>
   <font color="#FFFFFF">你可以将<%=animalname%>送给你的朋友！<br>不过神兽因水土不服会降低一半的攻击力！
  </font> </div>
</body>
</html>
</body>
</html>
