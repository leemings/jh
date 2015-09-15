<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Buffer=true
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
Response.Expires=0
chatbgcolor=Session("afa_chatbgcolor")
chatimage=Application("aqjh_chatimage")
if chatbgcolor="" then chatbgcolor="008888"
nowinroom=session("nowinroom")
online=Application("aqjh_onlinelist"&nowinroom)
onlineno=ubound(online)
aqjh_roominfo=split(Application("aqjh_room"),";")
chatinfo=split(aqjh_roominfo(nowinroom),"|")
%>
<script Language="JavaScript">
if(parent.document.URL.indexOf("file:")>=0||parent.f2.document.URL.indexOf("file:")>=0){parent.location.href='chaterr.asp?id=001';}
parent.md1(<%=onlineno%>, '<%=Session("afa_chatbgcolor")%>');
var chat="<%=nowinroom%>"
<%
for o = 1 to onlineno 
Response.Write "parent.md2("&chr(34)& online(o) &chr(34)&");"& chr(13) & chr(10)
next
%>
if (chat == 5){

for(var gw=10;gw<3200;gw++){gw=gw+100;
parent.f3.document.writeln("<a href=javascript:parent.sw('["+gw+"级怪物]');parent.sws('[0]') title=\"============&#13&#10怪物凶猛 切勿乱打&#13&#10============\";><br><img src='gwimage/"+gw+".gif'  border='0' width='36' height='36'><font color=yellow><font size=3>"+gw+"级怪物</font></font></a>&nbsp;<br>");}}


parent.md3();
</script>




</body>
</html>
