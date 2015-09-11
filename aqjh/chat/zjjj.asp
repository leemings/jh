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
chatbgcolor=Application("aqjh_chatbgcolor")
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
<title>终极绝技</title>
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
var dj=0;money=0;tl=0;nl=0
switch(zsname){
case "火之烈舞":dj=90;money=10000000;tl=40000;nl=40000;break;
case "风之乱舞":dj=100;money=15000000;tl=60000;nl=60000;break;
case "冰之封舞":dj=110;money=20000000;tl=70000;nl=70000;break;
case "终极杀着":dj=120;money=50000000;tl=90000;nl=90000;break;
}
document.myhyzj.maxnl.value=nl;
document.myhyzj.maxtl.value=tl;
document.myhyzj.nl.value=nl;
document.myhyzj.tl.value=tl;
document.myhyzj.money.value=money;
document.myhyzj.dj.value=dj;
if(tl<100) alert("选择错误！");
if(dj>mydj) alert("你的等级不够呀，发不了此招式！");
sm.innerHTML='<font size="-1">[<font color=blue><b>'+zsname+'</b></font>]需江湖等级大于等于<font color=red>'+dj+'</font>级，现金:<font color=red>'+money+'</font>两，体力:<font color=red>'+tl+'</font>点,内力:<font color=red>'+nl+'</font>点!</font>'
}
</script>
<body oncontextmenu=self.event.returnValue=false bgcolor="<%=chatbgcolor%>" background="<%=chatimage%>" bgproperties="fixed">
<div align="center"><font color="#FFFF00">终极绝技</font><br>
  <br>
  <form method=POST action='zjjjok.asp' name='myhyzj'>
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
      <option value="火之烈舞" selected>火之烈舞</option>
      <option value="风之乱舞">风之乱舞</option>
      <option value="冰之封舞">冰之封舞</option>
      <option value="终极杀着">终极杀着</option>
    </select>
    <br>
    <br>
    使用体力 
    <input type="text" name="tl" value="50000" maxlength="6" size="6">
    <input type="hidden" name="maxtl" value="50000" maxlength="6" size="6">
    <br>
    使用内力 
    <input type="text" name="nl" value="50000" maxlength="6" size="6">
    <input type="hidden" name="maxnl" value="50000" maxlength="6" size="6">
    <input type="hidden" name="money" value="5000000" maxlength="10" size="10">
    <input type="hidden" name="dj" value="4" maxlength="5" size="5">
    <br>
    </font><font color=#ffffff> 
    <input type=submit value="施展绝技" name=submit>
    </font>
  </form>
  <table width="95%" border="1" cellspacing="3" cellpadding="3">
    <tr > 
      <td id="sm"><font color="#FFFF00">选择攻击对象、招式名称、攻击内力、体力，然后点施展绝技进行攻击!</font></td>
    </tr>
  </table>
</div>
</body>
</html>
