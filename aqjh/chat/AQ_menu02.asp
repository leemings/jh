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
<HTML><HEAD><TITLE>�˵�</TITLE>
<META http-equiv=Content-Type content="text/html; charset=gb2312">
<STYLE type=text/css>BODY {FONT-SIZE: 9pt}
INPUT {FONT-SIZE: 9pt}
A {FONT-SIZE: 9pt; COLOR: #ffffff; TEXT-DECORATION: none}
A:hover {COLOR: red; TEXT-DECORATION: underline}
SELECT {BORDER-TOP-WIDTH: 1px; BORDER-LEFT-WIDTH: 1px; FONT-SIZE: 9pt; BORDER-BOTTOM-WIDTH: 1px; COLOR: black; FONT-FAMILY: ����; BACKGROUND-COLOR: #efefef; BORDER-RIGHT-WIDTH: 1px}
TD {FONT-SIZE: 9pt; COLOR: #ffffff; FONT-FAMILY: ����}
.btnStyle���� {FONT-SIZE: 9pt; CURSOR: hand}
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
function ex(){if(confirm("���������ɡ����ֽ��������ṩ����л����ʹ�á���ӭ����[<%=Application("aqjh_chatroomname")%>]��see you")){parent.t.location.href='about:blank';top.location.href="exitlt.asp";return true;}}
</script>
<script language="JavaScript"> 
//if(window.top==window.self){var i=1;while (i<=50){window.alert("������ʲôѽ�����ң������ǲ��еģ�ȥ����ȥ�ɣ�����������50�Σ���");i=i+1;}top.location.href="../exit.asp"}
</script>
</HEAD>
<BODY bgColor=#666699 leftMargin=0 topMargin=0 oncontextmenu=window.event.returnValue=false ondragstart=window.event.returnValue=false onselectstart=event.returnValue=false>
<table height="26" cellSpacing="0" cellPadding="0" width="100%" align="center" border="0">
  <tr align="middle" bgColor="#b7d4f1">
    <td onmouseover="this.style.backgroundColor='#efefef'" onmouseout="this.style.backgroundColor=''" width="361" bgColor="#E7E3F7" height="26">
      <p align="center"><img src="AQ_menu01.gif" width="69" height="24" border="0" usemap="#Map"></p>    </td>
    <td onmouseover="this.style.backgroundColor='#efefef'" onmouseout="this.style.backgroundColor=''" width="424" bgColor="#E7E3F7" height="26">
      <p align="center"><img src="AQ_menu02.gif" width="69" height="23" border="0" usemap="#Map2"></p>    </td>
  </tr>
</table>
<TABLE height=85 cellSpacing=0 cellPadding=0 width=100% align=center border=0>  <TBODY>
  <TR>
    <TD width=100% background=BG1.jpg height=57 rowSpan=2>
      <DIV align=left>
      <TABLE style="BORDER-COLLAPSE: collapse" borderColor=#ffffff height=63 
      cellSpacing=0 borderColorDark=#ffffff cellPadding=0 width=94% 
      borderColorLight=#ffffff border=1>
        <TBODY>
        <TR align=middle>
          <TD><a href="#" onClick="window.open('../garden/index.asp','','scrollbars=yes,resizable=yes,width=900,height=800')" title="�μӻ�԰�������ֻ�״Ԫ"><font color="#00FFFF">����</font></a>
          <TD><a href="#" onClick="window.open('../jhshow/jl.asp','','scrollbars=yes,resizable=yes,width=900,height=800')" title="�μ���װ���������һ��"><font color="#00FF00">����</font></a>
          <TD><a href="331.asp" onClick="javascript:s()" target="f3"><font color="#00FFFF">��</font></a><a href="long/myanimal.asp" onClick="javascript:s()" target="f3"><font color="red">��</TD>
          <TD><a href="#" onClick="window.open('../jhshow/index.asp','','scrollbars=yes,resizable=yes,width=900,height=800')" title="��������һ����"><font color="#FFFF00">��װ</font></a> 
        <TR align=middle>
          <TD><a href=dushu.asp onClick="javascript:s()" target=f3>�ľ�</a></TD>
          <TD><a href="mp/mp.asp" onClick="javascript:s()" target=f3>��ս</a></TD>
          <TD><a href="boy/boy.ASP" onClick="javascript:s()" target=f3><font color=yellow>С��</font></a></TD>
          <TD><a href="cw/cw.asp" onClick="javascript:s()" target=f3>����</a></TD></TR>
        <TR align=middle>
          <TD><a href=../help.asp target=_blank>����</a></TD>
          <TD><a href="mtv/song.asp" onClick="javascript:s()" target="f3"><font color=blue>��</font></a><a href="#" onClick="window.open('../song/top.asp','','scrollbars=yes,resizable=yes,width=600,height=500')" title="���ֽ������ϵͳ">��</TD>
          <TD><a href="face.asp" onClick="javascript:s()" target="f3">ͷ��</a></TD>
          <TD><a href='setfontsize.asp' onClick="javascript:s()" target=f3>����</a></TD></TR>
        <TR align=middle>
          <TD><a href="xlwupin.asp" onClick="javascript:s()" target="f3">����</font></a></TD>
          <TD><a href="#" onMouseOver="window.status='���� Web ICQ ��Ϣ��';return true" onMouseOut="window.status='';return true" onClick="javascript:window.open('webicq.asp','hqtwebicq','width=380,height=320')">����</a></TD> 
          <TD><a href="map/npc.asp" onClick="javascript:s()" target=f3>����</a></TD>
          <TD><a href="#" onMouseOver="window.status='�˳������ң����Զ����澭��ֵ��';return true" onMouseOut="window.status='';return true" onclick="javascript:ex()">�뿪</font></a></TD></TR></TBODY></TABLE></CENTER></DIV></TD>
</TR></TBODY></TABLE></BODY></HTML>
<map name="Map">
  <area shape="rect" coords="4,2,63,21" href="AQ_menu01.asp">
</map>
<map name="Map2">
  <area shape="rect" coords="4,1,67,21" href="AQ_menu02.asp">
</map>
</BODY></HTML>