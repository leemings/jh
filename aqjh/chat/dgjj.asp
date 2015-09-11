<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
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

%>
<html>
<head>
<title>独孤九剑</title>
<style type='text/css'>
body{font-size:9pt;}
td{font-size:9pt;}
input{font-size:9pt;}
a{font-size:9pt; color:black;text-decoration:none;}
a:hover{color:red;text-decoration:none;}
</style>
</head>
<script Language=Javascript>
//parent.show.fm1.innerHTML="&chr(34)&mcsl&chr(34)&";
//parent.show.fm2.innerHTML="&chr(34)&"0"&chr(34)&";parent.show.fm3.innerHTML="&chr(34)&"0"&chr(34)&";parent.show.shiform.fmsl.value=0
function dgjj(zsname)
{mydj=<%=aqjh_jhdj%>
var dj=0;money=0;tl=0;nl=0
switch(zsname){
case "总决式":dj=24;money=1205000;tl=3000;nl=2000;break;
case "破刀式":dj=26;money=2511000;tl=4000;nl=3000;break;
case "破剑式":dj=28;money=3500000;tl=5000;nl=4000;break;
case "破枪式":dj=30;money=4520000;tl=6000;nl=5000;break;
case "破鞭式":dj=32;money=5025000;tl=7000;nl=6000;break;
case "破索式":dj=34;money=6030000;tl=8000;nl=6000;break;
case "破箭式":dj=36;money=7035000;tl=9000;nl=8000;break;
case "破掌式":dj=36;money=7035000;tl=9000;nl=8000;break;
case "破气式":dj=40;money=11002000;tl=11000;nl=33500;break;
}
document.mydgjj.maxnl.value=nl;
document.mydgjj.maxtl.value=tl;
document.mydgjj.nl.value=nl;
document.mydgjj.tl.value=tl;
document.mydgjj.money.value=money;
document.mydgjj.dj.value=dj;
if(tl<100) alert("选择错误！");
if(dj>mydj) alert("你的等级不够呀，发不了此招式！");
sm.innerHTML='<font size="-1">[<font color=blue><b>'+zsname+'</b></font>]需江湖等级大于等于<font color=red>'+dj+'</font>级，现金:<font color=red>'+money+'</font>两，体力:<font color=red>'+tl+'</font>点,内力:<font color=red>'+nl+'</font>点!</font>'
}
</script>
<body oncontextmenu=self.event.returnValue=false bgcolor="<%=chatbgcolor%>" background="<%=chatimage%>" bgproperties="fixed">
<div align="center">独孤九剑<br>
  <br>
  <form method=POST action='dgjjok.asp' name='mydgjj'  target='d'>
    <font color="#FFFFFF"> </font><font color="#FF0033">攻击对象</font><br>
    <select name="to1">
      <%useronlinename=Application("aqjh_useronlinename"&session("nowinroom"))
show=Split(Trim(useronlinename),"  ",-1)
x=UBound(show)
for i=0 to x
if show(i)<>aqjh_name and show(i)<>peiou then
%>
      <option value="<%=show(i)%>" selected><%=show(i)%></option>
<%end if
next%>
    </select>
    <font color="#FFFFFF"><br>
    <br>
    <select name='dgjjzs' onChange="dgjj(this.value);" style='font-size:12px'>
      <option value="总决式" selected>总决式</option>
      <option value="破刀式">破刀式</option>
      <option value="破剑式">破剑式</option>
      <option value="破枪式">破枪式</option>
      <option value="破鞭式">破鞭式</option>
      <option value="破索式">破索式</option>
      <option value="破箭式">破箭式</option>
      <option value="破掌式">破掌式</option>
      <option value="破气式">破气式</option>
    </select>
    <br>
    <br>
    使用体力 
    <input type="text" name="tl" value="3000" maxlength="5" size=5>
    <input type="hidden" name="maxtl" value="3000" maxlength="5" size=5>
    <br>
    使用内力 
    <input type="text" name="nl" value="2000" maxlength="5" size="5">
    <input type="hidden" name="maxnl" value="2000" maxlength="5" size="5">
    <input type="hidden" name="money" value="1205000" maxlength="10" size="10">
    <input type="hidden" name="dj" value="4" maxlength="5" size="5">
    <br>
    </font><font color=#ffffff> 
    <input type=submit value="施展绝技" name=submit>
  </form>
  <table width="95%" border="1" cellspacing="3" cellpadding="3">
    <tr > 
      <td id="sm">选择攻击对象、招式名称、攻击内力、体力，然后点施展绝技进行攻击!</font></td>
    </tr>
  </table>
</div>
</body>
</html>
