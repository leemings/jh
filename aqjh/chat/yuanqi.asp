<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
http = Request.ServerVariables("HTTP_REFERER") 
if InStr(http,"chat")=0 then 
	Response.Write "<script language=javascript>{alert('对不起，程序拒绝您的操作！！！\n     按确定返回！');}</script>" 
	Response.End 
end if
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
chatbgcolor=Session("afa_chatbgcolor")
chatimage=Application("aqjh_chatimage")
if chatbgcolor="" then chatbgcolor="008888"
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
if Instr(LCase(Application("aqjh_useronlinename"&session("nowinroom")))," "&LCase(aqjh_name)&" ")=0 then
	Response.Write "<script Language=Javascript>alert('你不能进行操作，进行此操作必须进入聊天室！');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
aqjh_roominfo=split(Application("aqjh_room"),";")
chatinfo=split(aqjh_roominfo(session("nowinroom")),"|")
if chatinfo(0)="夺宝之战" then
	tj="dgjjokdb.asp"
else
	tj="hyjzok.asp"
end if
erase aqjh_roominfo
erase chatinfo
%>
<html>
<head>
<title>元气攻击</title>
</head>
<style type='text/css'>
body{font-size:9pt;}
td{font-size:9pt;}
input{font-size:9pt;}
a{font-size:9pt; color:black;text-decoration:none;}
a:hover{color:red;text-decoration:none;}
</style>
<script Language=Javascript>
//parent.show.fm1.innerHTML="&chr(34)&mcsl&chr(34)&";
//parent.show.fm2.innerHTML="&chr(34)&"0"&chr(34)&";parent.show.fm3.innerHTML="&chr(34)&"0"&chr(34)&";parent.show.shiform.fmsl.value=0
function hyjz(zsname)
{mydj=<%=aqjh_jhdj%>
var dj=0;yuanqi=0
switch(zsname){
case "至高无上":dj=20;yuanqi=4000;break;
case "惟我独尊":dj=30;yuanqi=6000;break;
case "心比天高":dj=40;yuanqi=7000;break;
case "天之骄子":dj=50;yuanqi=8000;break;
case "终极无限":dj=60;yuanqi=10000;break;
}
document.myhyzj.maxyuanqi.value=yuanqi;
document.myhyzj.yuanqi.value=yuanqi;
document.myhyzj.dj.value=dj;
if (dj>mydj){
alert("你的等级不够呀，发不了此招式！");
document.myhyzj.submit.disabled=true;
}
else{
document.myhyzj.submit.disabled=false;
}
sm.innerHTML='<font size="-1">[<font color=blue><b>'+zsname+'</b></font>]需江湖等级大于等于<font color=red>'+dj+'</font>级，元气:<font color=red>'+yuanqi+'</font>点!</font>'
}
</script>
<body oncontextmenu=self.event.returnValue=false bgcolor="<%=chatbgcolor%>" background="<%=chatimage%>" bgproperties="fixed">
<div align="center"><font color="red">元气攻击</font><br>
  <br>
  <form method=POST action='yuanqiok.asp' name='myhyzj'>
 <font color="#FF0033">攻击对象</font><br>
    <select name="to1">
      <%useronlinename=Application("aqjh_useronlinename"&session("nowinroom"))
show=Split(Trim(useronlinename),"  ",-1)
x=UBound(show)
for i=0 to x
if show(i)<>aqjh_name Then
%>
      <option value="<%=show(i)%>" selected><%=show(i)%></option>
<%end if
next%>
    </select>
    <font color="#FFFFFF"><br>
    <br>
    <select name='hyjzzs' onChange="hyjz(this.value);" style='font-size:12px'>
      <option value="至高无上" selected>至高无上</option>
      <option value="惟我独尊">惟我独尊</option>
      <option value="心比天高">心比天高</option>
      <option value="天之骄子">天之骄子</option>
      <option value="终极无限">终极无限</option>
</select>
    <br>
    <br>
    使用元气 
    <input type="text" name="yuanqi" value="400" maxlength="5" size="5">
    <input type="hidden" name="maxyuanqi" value="400" maxlength="5" size="5">
    <input type="hidden" name="dj" value="4" maxlength="5" size="5">
    <br>
</font><font color=#ffffff> 
    <input type=submit value="施展元气" name=submit>
    </font>
  </form>
  <table width="95%" border="1" cellspacing="3" cellpadding="3">
    <tr > 
      <td id="sm"><font color="#FFFF00">选择攻击对象、招式名称、元气数，然后点施展元气进行攻击!</font></td>
    </tr>
  </table>
</div>
</body>
</html>
