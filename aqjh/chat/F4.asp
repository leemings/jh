<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_roominfo=split(Application("aqjh_room"),";")
chatinfo=split(aqjh_roominfo(session("nowinroom")),"|")
chatbgcolor=Application("aqjh_chatbgcolor")
chatimage=Application("aqjh_chatimage")
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
if chatbgcolor="" then chatbgcolor="008888"%>
<HTML><HEAD><TITLE>菜单</TITLE>
<META http-equiv=Content-Type content="text/html; charset=gb2312">
<STYLE type=text/css>BODY {FONT-SIZE: 9pt}
INPUT {FONT-SIZE: 9pt}
A {FONT-SIZE: 9pt; COLOR: #ffffff; TEXT-DECORATION: none}
A:hover {COLOR: red; TEXT-DECORATION: underline}
SELECT {BORDER-TOP-WIDTH: 1px; BORDER-LEFT-WIDTH: 1px; FONT-SIZE: 9pt; BORDER-BOTTOM-WIDTH: 1px; COLOR: black; FONT-FAMILY: 宋体; BACKGROUND-COLOR: #efefef; BORDER-RIGHT-WIDTH: 1px}
TD {FONT-SIZE: 9pt; COLOR: #ffffff; FONT-FAMILY: 宋体}
.btnStyle　　 {FONT-SIZE: 9pt; CURSOR: hand}
.webstyle {FONT-SIZE: 8pt; FONT-FAMILY: Webdings}
</STYLE>
<script Language="Javascript">
function s(){parent.f2.document.af.mdsx.checked=false;parent.f2.document.af.sytemp.focus();}
if(parent.document.URL.indexOf("file:")>=0||parent.f2.document.URL.indexOf("file:")>=0){parent.location.href='chaterr.asp?id=001';}
function showname(){
parent.f2.document.af.mdsx.checked=true;parent.f2.document.af.sytemp.focus();
if(parent.m.location.href=="about:blank"){
parent.m.location.href="f3.asp";}
else
{parent.m.location.reload();}
}
function ex(){if(confirm("本程序是由『快乐江湖网』提供，感谢您的使用。欢迎再来[<%=Application("aqjh_chatroomname")%>]，see you")){parent.t.location.href='about:blank';top.location.href="exitlt.asp";return true;}}
</script>
<script language="JavaScript"> 
if(window.top==window.self){var i=1;while (i<=50){window.alert("你想作什么呀，黑我？这里是不行的，去别处玩去吧！哈！慢慢点50次！！");i=i+1;}top.location.href="../exit.asp"}
</script>
</HEAD>
<BODY bgColor=#006699 leftMargin=0 topMargin=0 oncontextmenu=window.event.returnValue=false ondragstart=window.event.returnValue=false onselectstart=event.returnValue=false>
<TABLE height=120 cellSpacing=0 cellPadding=0 width=100% align=center border=0>  <TBODY>
  <TR>
    <TD width=100% background=1.jpg height=115 rowSpan=2>
      <DIV align=left>
      <TABLE style="BORDER-COLLAPSE: collapse" borderColor=#ffffff height=125 
      cellSpacing=0 borderColorDark=#ffffff cellPadding=0 width=94% 
      borderColorLight=#ffffff border=1>
        <TBODY>
        <TR align=middle>
          <TD><a href="savevalue.asp" onClick="javascript:s()" target="f3">存点</a></TD>
          <TD><a href="#" onClick="javascript:showname();">名单</a></TD>
          <TD><a href="selectchatroom.asp" onClick="javascript:s()" target="f3">房间</A></TD><td><a href="yuanqi.asp" onClick="javascript:s()" title="泡点得到的元气可以杀人！" target="f3">元气</A></td></TR>
        <TR align=middle>
          <TD><a href="KING1.asp" onClick="javascript:s()" title="江湖功能区！" target=f3><font color=blue>功</font></a><a href="king.asp" onClick="javascript:s()" target="f3"><font color=red>能</font></a></TD>
          <TD><a href="#" onClick="window.open('../garden/hua.asp','','scrollbars=yes,resizable=yes,width=700,height=400')" title="种花参加花园大赛做种花状元">花园</a></TD>
          <TD><a href="#" onClick="window.open('../twt/mhmj/mhmj/index.asp','','scrollbars=yes,resizable=yes,width=700,height=400')" title="打怪的好地方！"><font color=red>魔界</font></font></a></TD>
          <TD><a href="new2.asp" target="_blank" title="快乐茶壶任务系统"><font color="red">任务</font></a></td>         
        <TR align=middle>
          <TD><a href='setwg3.asp' onClick="javascript:s()" target=f3>动武</a></TD>
          <TD><a href="setwg4.asp" title="轩辕" target="f3" onClick="javascript:s()">轩辕</a></TD>
          <TD><a href="wupin.asp" onClick="javascript:s()" target="f3">物品</font></a></TD>
          <TD><a href="#" onClick="window.open('jhjh/index.asp','','scrollbars=yes,resizable=yes,width=800,height=600')" title="快乐江湖进化系统"><font color="#00FFFF">进化</font></a></TD></TR>
        <TR align=middle>
          <TD><a href="zhuangtai.asp" onClick="javascript:s()" target="f3"><font color=red>状</font></a><a href="#"  onclick="javascript:window.open('zt.ASP','zhuantai','scrollbars=no,toolbar=no,menubar=no,location=no,status=no,resizable=no,width=630,height=480,top=0,left=0')" target="_self"><font color=green>态</font></font></a></TD>
          <TD><a href="stuntlist.asp" onClick="javascript:s()" target="f3">特技</a></TD>
          <TD><a href="#" onClick="window.open('../fish/index.asp','','scrollbars=yes,resizable=yes,width=700,height=400')" title="全新的钓鱼系统"><font color=blue>钓鱼</font></a></TD>
          <TD><a href="331.asp" onClick="javascript:s()" target="f3"><font color="#00FFFF">兽</font></a><a href="long/myanimal.asp" onClick="javascript:s()" target="f3"><font color="red">龙</TD>
        <TR align=middle>
          <TD><a href=dushu.asp onClick="javascript:s()" target=f3>赌局</a></TD>
          <TD><a href="mp/mp.asp" onClick="javascript:s()" target=f3>大战</a></TD>
          <TD><a href="boy/boy.ASP" onClick="javascript:s()" target=f3><font color=yellow>小孩</font></a></TD>
          <TD><a href="cw/cw.asp" onClick="javascript:s()" target=f3>宠物</a></TD></TR>
        <TR align=middle>
          <TD><a href="xlwupin.asp" onClick="javascript:s()" target="f3">拍卖</font></a></TD>
          <TD><a href="mtv/song.asp" onClick="javascript:s()" target="f3"><font color=blue>音</font></a><a href="#" onClick="window.open('../song/top.asp','','scrollbars=yes,resizable=yes,width=600,height=500')" title="快乐江湖点歌系统">乐</TD>
          <TD><a href="map/npc.asp" onClick="javascript:s()" target=f3>行走</a></TD>
          <TD><a href="#" onMouseOver="window.status='退出聊天室，并自动保存经验值。';return true" onMouseOut="window.status='';return true" onclick="javascript:ex()">离开</font></a></TD></TR>	
</TBODY></TABLE></CENTER></DIV></TD></TR>
		</TBODY></TABLE>
</BODY></HTML>
