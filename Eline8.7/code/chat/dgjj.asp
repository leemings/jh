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
	Response.Write "<script Language=Javascript>alert('�㲻�ܽ��в��������д˲���������������ң�');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if

%>
<html>
<head>
<title>���¾Ž���һ�������wWw.51eline.com��</title>
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
function dgjj(zsname)
{mydj=<%=sjjh_jhdj%>
var dj=0;money=0;tl=0;nl=0
switch(zsname){
case "�ܾ�ʽ":dj=14;money=1205000;tl=3000;nl=2000;break;
case "�Ƶ�ʽ":dj=16;money=2511000;tl=4000;nl=3000;break;
case "�ƽ�ʽ":dj=18;money=3500000;tl=5000;nl=4000;break;
case "��ǹʽ":dj=20;money=4520000;tl=6000;nl=5000;break;
case "�Ʊ�ʽ":dj=22;money=5025000;tl=7000;nl=6000;break;
case "����ʽ":dj=24;money=6030000;tl=8000;nl=6000;break;
case "�Ƽ�ʽ":dj=26;money=7035000;tl=9000;nl=8000;break;
case "����ʽ":dj=26;money=7035000;tl=9000;nl=8000;break;
case "����ʽ":dj=30;money=11002000;tl=11000;nl=33500;break;
}
document.mydgjj.maxnl.value=nl;
document.mydgjj.maxtl.value=tl;
document.mydgjj.nl.value=nl;
document.mydgjj.tl.value=tl;
document.mydgjj.money.value=money;
document.mydgjj.dj.value=dj;
if(tl<100) alert("ѡ�����");
if(dj>mydj) alert("��ĵȼ�����ѽ�������˴���ʽ��");
sm.innerHTML='<font size="-1">[<font color=blue><b>'+zsname+'</b></font>]�轭���ȼ����ڵ���<font color=red>'+dj+'</font>�����ֽ�:<font color=red>'+money+'</font>��������:<font color=red>'+tl+'</font>��,����:<font color=red>'+nl+'</font>��!</font>'
}
</script>
<body bgcolor="#006699" leftmargin="0" topmargin="2" bgproperties="fixed" oncontextmenu=self.event.returnValue=false>
<div align="center"><font color="#FFFFFF">���¾Ž�</font><br>
  <br>
  <form method=POST action='dgjjok.asp' name='mydgjj'  target='d'>
    <font color="#FFFFFF" size="2"> </font><font size="2" color="#FF0033">��������</font><br>
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
    <select name='dgjjzs' onChange="dgjj(this.value);" style='BORDER-RIGHT: #000000 1px solid; BORDER-TOP: #000000 1px solid; FONT-SIZE: 12px; BACKGROUND: #666666; BORDER-LEFT: #000000 1px solid; COLOR: #f6e700; BORDER-BOTTOM: #000000 1px solid; HEIGHT: 16px;font-size:12px'>
      <option value="�ܾ�ʽ" selected>�ܾ�ʽ</option>
      <option value="�Ƶ�ʽ">�Ƶ�ʽ</option>
      <option value="�ƽ�ʽ">�ƽ�ʽ</option>
      <option value="��ǹʽ">��ǹʽ</option>
      <option value="�Ʊ�ʽ">�Ʊ�ʽ</option>
      <option value="����ʽ">����ʽ</option>
      <option value="�Ƽ�ʽ">�Ƽ�ʽ</option>
      <option value="����ʽ">����ʽ</option>
      <option value="����ʽ">����ʽ</option>
    </select>
    <br>
    <br>
    <font color="#cc3300">ʹ������ </font>
    <input type="text" name="tl" value="3000" maxlength="5" size="5" style="BACKGROUND-COLOR: #f0f0f0; BORDER-BOTTOM: #cc3300 1px solid; BORDER-LEFT: #cc3300 1px solid; BORDER-RIGHT: #cc3300 1px solid; BORDER-TOP: #cc3300 1px solid; COLOR: #000000; FONT-SIZE: 9pt">
    <input type="hidden" name="maxtl" value="3000" maxlength="5" size="5">
    <br>
    <font color="#cc3300">ʹ������ </font>
    <input type="text" name="nl" value="2000" maxlength="5" size="5" style="BACKGROUND-COLOR: #f0f0f0; BORDER-BOTTOM: #cc3300 1px solid; BORDER-LEFT: #cc3300 1px solid; BORDER-RIGHT: #cc3300 1px solid; BORDER-TOP: #cc3300 1px solid; COLOR: #000000; FONT-SIZE: 9pt">
    <input type="hidden" name="maxnl" value="2000" maxlength="5" size="5">
    <input type="hidden" name="money" value="1205000" maxlength="10" size="10">
    <input type="hidden" name="dj" value="4" maxlength="5" size="5">
    <br>
    </font><br> 
    <input type=submit value="ʩչ����" name=submit style="BACKGROUND-COLOR: #009933; BORDER-BOTTOM: #ffffff 1px solid; BORDER-LEFT: #ffffff 1px solid; BORDER-RIGHT: #ffffff 1px solid; BORDER-TOP: #ffffff 1px solid; CURSOR: hand; COLOR: #FFFFFF; FONT-SIZE: 9pt; WIDTH:55;HEIGHT: 18px;">
    <font color="#FFFFFF" size="2"><br>
    </font> 
  </form>
  <table width="80%" border="1" cellspacing="0" cellpadding="2" bordercolorlight="#000000" bordercolordark="#FFFFFF">
    <tr > 
      <td id="sm"><font color="#FFFFFF" size="-1">ѡ�񹥻�������ʽ���ơ���������������.Ȼ���ʩչ�������й���!</font></td>
    </tr>
  </table><br>
    
  <div align="center"> <a href=javascript:history.go(-1)><font color=#00FF00 size=-1>�����ϼ�</font></a></div>
</div>
</body>
</html>
