<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
http = Request.ServerVariables("HTTP_REFERER") 
if InStr(http,"chat")=0 then 
	Response.Write "<script language=javascript>{alert('�Բ��𣬳���ܾ����Ĳ���������\n     ��ȷ�����أ�');}</script>" 
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
	Response.Write "<script Language=Javascript>alert('�㲻�ܽ��в��������д˲���������������ң�');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
aqjh_roominfo=split(Application("aqjh_room"),";")
chatinfo=split(aqjh_roominfo(session("nowinroom")),"|")
if chatinfo(0)="�ᱦ֮ս" then
	tj="dgjjokdb.asp"
else
	tj="hyjzok.asp"
end if
erase aqjh_roominfo
erase chatinfo
%>
<html>
<head>
<title>��Ա����</title>
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
case "�ֽ���":dj=25;money=1000000;tl=4000;nl=4000;break;
case "��������":dj=30;money=1500000;tl=6000;nl=6000;break;
case "������":dj=35;money=2000000;tl=7000;nl=7000;break;
case "�����Ʒ�":dj=40;money=2500000;tl=8000;nl=8000;break;
case "��Ӱ����":dj=45;money=3000000;tl=9000;nl=9000;break;
case "������ʽ":dj=50;money=3500000;tl=10000;nl=10000;break;
case "���귭��":dj=55;money=4000000;tl=11000;nl=11000;break;
case "��ѩ�ɳ�":dj=60;money=4500000;tl=12000;nl=12000;break;
case "�������":dj=70;money=5000000;tl=13000;nl=13000;break;
}
document.myhyzj.maxnl.value=nl;
document.myhyzj.maxtl.value=tl;
document.myhyzj.nl.value=nl;
document.myhyzj.tl.value=tl;
document.myhyzj.money.value=money;
document.myhyzj.dj.value=dj;
if(tl<100) alert("ѡ�����");
if(dj>mydj) alert("��ĵȼ�����ѽ�������˴���ʽ��");
sm.innerHTML='<font size="-1">[<font color=blue><b>'+zsname+'</b></font>]�轭���ȼ����ڵ���<font color=red>'+dj+'</font>�����ֽ�:<font color=red>'+money+'</font>��������:<font color=red>'+tl+'</font>��,����:<font color=red>'+nl+'</font>��!</font>'
}
</script>
<body oncontextmenu=self.event.returnValue=false bgcolor="<%=chatbgcolor%>" background="<%=chatimage%>" bgproperties="fixed">
<div align="center"><font color="#FFFF00">��Ա����</font><br>
  <br>
  <form method=POST action='hyjzok.asp' name='myhyzj'>
 <font color="#FF0033">��������</font><br>
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
      <option value="�ֽ���" selected>�ֽ���</option>
      <option value="��������">��������</option>
      <option value="������">������</option>
      <option value="�����Ʒ�">�����Ʒ�</option>
      <option value="��Ӱ����">��Ӱ����</option>
      <option value="������ʽ">������ʽ</option>
      <option value="���귭��">���귭��</option>
      <option value="��ѩ�ɳ�">��ѩ�ɳ�</option>
      <option value="�������">�������</option>
    </select>
    <br>
    <br>
    ʹ������ 
    <input type="text" name="tl" value="50000" maxlength="5" size="5">
    <input type="hidden" name="maxtl" value="50000" maxlength="5" size="5">
    <br>
    ʹ������ 
    <input type="text" name="nl" value="50000" maxlength="5" size="5">
    <input type="hidden" name="maxnl" value="50000" maxlength="5" size="5">
    <input type="hidden" name="money" value="50000000" maxlength="10" size="10">
    <input type="hidden" name="dj" value="4" maxlength="5" size="5">
    <br>
    </font><font color=#ffffff> 
    <input type=submit value="ʩչ����" name=submit>
    </font>
  </form>
  <table width="95%" border="1" cellspacing="3" cellpadding="3">
    <tr > 
      <td id="sm"><font color="#FFFF00">ѡ�񹥻�������ʽ���ơ�����������������Ȼ���ʩչ�������й���!</font></td>
    </tr>
  </table>
</div>
</body>
</html>
