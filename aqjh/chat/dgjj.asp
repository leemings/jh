<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
chatbgcolor=Session("afa_chatbgcolor")
chatimage=Application("aqjh_chatimage")
if chatbgcolor="" then chatbgcolor="008888"
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
if Instr(LCase(Application("aqjh_useronlinename"&session("nowinroom")))," "&LCase(aqjh_name)&" ")=0 then
	Response.Write "<script Language=Javascript>alert('�㲻�ܽ��в��������д˲���������������ң�');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if

%>
<html>
<head>
<title>���¾Ž�</title>
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
case "�ܾ�ʽ":dj=24;money=1205000;tl=3000;nl=2000;break;
case "�Ƶ�ʽ":dj=26;money=2511000;tl=4000;nl=3000;break;
case "�ƽ�ʽ":dj=28;money=3500000;tl=5000;nl=4000;break;
case "��ǹʽ":dj=30;money=4520000;tl=6000;nl=5000;break;
case "�Ʊ�ʽ":dj=32;money=5025000;tl=7000;nl=6000;break;
case "����ʽ":dj=34;money=6030000;tl=8000;nl=6000;break;
case "�Ƽ�ʽ":dj=36;money=7035000;tl=9000;nl=8000;break;
case "����ʽ":dj=36;money=7035000;tl=9000;nl=8000;break;
case "����ʽ":dj=40;money=11002000;tl=11000;nl=33500;break;
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
<body oncontextmenu=self.event.returnValue=false bgcolor="<%=chatbgcolor%>" background="<%=chatimage%>" bgproperties="fixed">
<div align="center">���¾Ž�<br>
  <br>
  <form method=POST action='dgjjok.asp' name='mydgjj'  target='d'>
    <font color="#FFFFFF"> </font><font color="#FF0033">��������</font><br>
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
    ʹ������ 
    <input type="text" name="tl" value="3000" maxlength="5" size=5>
    <input type="hidden" name="maxtl" value="3000" maxlength="5" size=5>
    <br>
    ʹ������ 
    <input type="text" name="nl" value="2000" maxlength="5" size="5">
    <input type="hidden" name="maxnl" value="2000" maxlength="5" size="5">
    <input type="hidden" name="money" value="1205000" maxlength="10" size="10">
    <input type="hidden" name="dj" value="4" maxlength="5" size="5">
    <br>
    </font><font color=#ffffff> 
    <input type=submit value="ʩչ����" name=submit>
  </form>
  <table width="95%" border="1" cellspacing="3" cellpadding="3">
    <tr > 
      <td id="sm">ѡ�񹥻�������ʽ���ơ�����������������Ȼ���ʩչ�������й���!</font></td>
    </tr>
  </table>
</div>
</body>
</html>
