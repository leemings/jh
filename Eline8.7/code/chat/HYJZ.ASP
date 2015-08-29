<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
chatbgcolor=Application("sjjh_chatbgcolor")
chatimage=Application("sjjh_chatimage")
if chatbgcolor="" then chatbgcolor="008888"
if sjjh_name="" then Response.Redirect "../error.asp?id=440"
if Instr(LCase(Application("sjjh_useronlinename"&session("nowinroom")))," "&LCase(sjjh_name)&" ")=0 then
	Response.Write "<script Language=Javascript>alert('你不能进行操作，进行此操作必须进入聊天室！');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
%>
<html>
<head>
<title>会员绝招</title>
<style>
body{
CURSOR: url('3.ani');
scrollbar-track-color:#ffffff;
SCROLLBAR-FACE-COLOR:#effaff;
SCROLLBAR-HIGHLIGHT-COLOR:#ffffff;
SCROLLBAR-SHADOW-COLOR:#eeeeee;
SCROLLBAR-3DLIGHT-COLOR:#eeeeee;
SCROLLBAR-ARROW-COLOR:#dddddd;
SCROLLBAR-DARKSHADOW-COLOR:#ffffff
}
</style>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312"></head>
<script Language=Javascript>
//parent.show.fm1.innerHTML="&chr(34)&mcsl&chr(34)&";
//parent.show.fm2.innerHTML="&chr(34)&"0"&chr(34)&";parent.show.fm3.innerHTML="&chr(34)&"0"&chr(34)&";parent.show.shiform.fmsl.value=0
function hyjz(zsname)
{mydj=<%=sjjh_jhdj%>
var dj=0;money=0;tl=0;nl=0
switch(zsname){
case "分筋错骨":dj=70;money=60000000;tl=140000;nl=140000;break;
case "化骨绵掌":dj=30;money=15000000;tl=60000;nl=60000;break;
case "明月神功":dj=35;money=20000000;tl=70000;nl=70000;break;
case "排云掌法":dj=40;money=25000000;tl=80000;nl=80000;break;
case "浮影剑法":dj=45;money=30000000;tl=90000;nl=90000;break;
case "雷霆三式":dj=50;money=35000000;tl=100000;nl=100000;break;
case "覆雨翻云":dj=55;money=40000000;tl=110000;nl=110000;break;
case "冰雪飞尘":dj=60;money=45000000;tl=120000;nl=120000;break;
case "天外飞仙":dj=70;money=50000000;tl=130000;nl=130000;break;
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
<body bgcolor="#006699" leftmargin="0" topmargin="2" bgproperties="fixed" oncontextmenu=self.event.returnValue=false>
<div align="center"><font color="#FFFFFF">会员绝招</font><br>
  <br>
  <form method=POST action='hyjzok.asp' name='myhyzj'  target='d'>
    <font color="#FFFFFF" size="2"> </font><font size="2" color="#FF0033">攻击对象</font><br>
    <select name="to1" style='BORDER-RIGHT: #000000 1px solid; BORDER-TOP: #000000 1px solid; FONT-SIZE: 12px; BACKGROUND: #666666; BORDER-LEFT: #000000 1px solid; COLOR: #f6e700; BORDER-BOTTOM: #000000 1px solid; HEIGHT: 16px;font-size:12px'>
      <%useronlinename=Application("sjjh_useronlinename"&session("nowinroom"))
show=Split(Trim(useronlinename),"  ",-1)
x=UBound(show)
for i=0 to x%>
      <option value="<%=show(i)%>" selected><%=show(i)%></option>
<%next%>
    </select>
    <font color="#FFFFFF" size="2"><br>
    <br>
    <select name='hyjzzs' onChange="hyjz(this.value);" style='BORDER-RIGHT: #000000 1px solid; BORDER-TOP: #000000 1px solid; FONT-SIZE: 12px; BACKGROUND: #666666; BORDER-LEFT: #000000 1px solid; COLOR: #f6e700; BORDER-BOTTOM: #000000 1px solid; HEIGHT: 16px;font-size:12px'>
      <option value="分筋错骨" selected>分筋错骨</option>
      <option value="化骨绵掌">化骨绵掌</option>
      <option value="明月神功">明月神功</option>
      <option value="排云掌法">排云掌法</option>
      <option value="浮影剑法">浮影剑法</option>
      <option value="雷霆三式">雷霆三式</option>
      <option value="覆雨翻云">覆雨翻云</option>
      <option value="冰雪飞尘">冰雪飞尘</option>
      <option value="天外飞仙">天外飞仙</option>
    </select>
    <br>
    <br>
    <font color="#cc3300">使用体力</font> 
    <input type="text" name="tl" value="50000" maxlength="5" size="5" style="BACKGROUND-COLOR: #f0f0f0; BORDER-BOTTOM: #cc3300 1px solid; BORDER-LEFT: #cc3300 1px solid; BORDER-RIGHT: #cc3300 1px solid; BORDER-TOP: #cc3300 1px solid; COLOR: #000000; FONT-SIZE: 9pt">
    <input type="hidden" name="maxtl" value="50000" maxlength="5" size="5">
    <br>
    <font color="#cc3300">使用内力</font> 
    <input type="text" name="nl" value="50000" maxlength="5" size="5" style="BACKGROUND-COLOR: #f0f0f0; BORDER-BOTTOM: #cc3300 1px solid; BORDER-LEFT: #cc3300 1px solid; BORDER-RIGHT: #cc3300 1px solid; BORDER-TOP: #cc3300 1px solid; COLOR: #000000; FONT-SIZE: 9pt">
    <input type="hidden" name="maxnl" value="50000" maxlength="5" size="5">
    <input type="hidden" name="money" value="50000000" maxlength="10" size="10">
    <input type="hidden" name="dj" value="4" maxlength="5" size="5">
    <br><br>
    </font><font color=#ffffff size=2> 
    <input type=submit value="施展绝技" name=submit style="BACKGROUND-COLOR: #009933; BORDER-BOTTOM: #ffffff 1px solid; BORDER-LEFT: #ffffff 1px solid; BORDER-RIGHT: #ffffff 1px solid; BORDER-TOP: #ffffff 1px solid; CURSOR: hand; COLOR: #FFFFFF; FONT-SIZE: 9pt; WIDTH:55;HEIGHT: 18px;">
    </font><font color="#FFFFFF" size="2"><br>
    </font> 
  </form>
  <table width="80%" border="1" cellspacing="0" cellpadding="2" bordercolorlight="#000000" bordercolordark="#FFFFFF">
    <tr > 
      <td id="sm"><font size="-1" color="#FFFFFF">选择攻击对象、招式名称、攻击内力、体力，然后点施展绝技进行攻击</font></td>
    </tr>
  </table><br>
  <div align="center"> <a href=javascript:history.go(-1)><font color=#00FF00 size=-1>返回上级</font></a></div>
</div>
</body>
</html>
