<!--#include file="conn.asp"-->
<!-- #include file="inc/const.asp" -->
<!--#include file="z_plus_check.asp"-->

<%
call nav()
stats="�����ĳ�"
call head_var(2,0,"","")
if Cint(GroupSetting(14))=0 then
Errmsg=Errmsg+"<br>"+"<li>��û���ڱ������ĳ���Ȩ�ޣ���<a href=login.asp>��¼</a>����ͬ����Ա��ϵ��"
call dvbbs_error()
else
response.write "<table cellpadding=3 cellspacing=1 align=center class=tableborder1><tr><th valign=middle colspan=2 align=center height=25><b>�����ĳ�</b></td></tr><tr><td valign=middle class=tablebody1 height=100><CENTER>"


dim d1,d2,d3,b1,b2,b3,zj,ni,zs,zds,zdss,dmoney
set rs=server.createobject("adodb.recordset")
sql="select * from [user] where username='"&membername&"'"
rs.Open sql,Conn,1,3
select case Request("menu")
case ""
z_play()

case "dx"
dx()

case "ds"
ds()

case "sz"
sz()

case "dx1"
checkup()
dx1()

case "ds1"
checkup()
ds1()

case "sz1"
checkup()
sz1()

end select

sub checkup()
dmoney=int(request("money"))
if dmoney > rs("userwealth") then
error2("���Ǯ������")
end if
if dmoney<1 or dmoney>100 then
error2("������󣡶�ע��Χ1-100")
end if

if rs("usercp")<1 then
error2("��û��������!")
end if
end sub



sub z_play()%>
<title>��������</title><BODY><br>

  <center>
  <table width="600" border="1" height="143" cellpadding="4" cellspacing="4" bordercolor="#004080" style="border-collapse: collapse">
<tr align="center" bgcolor="#99CCFF">
  <td bgcolor="#ECF5FF" bordercolor="#004080" height="129">
  <table border="0" width="100%" cellspacing="0" cellpadding="6" style="border-collapse: collapse" bordercolor="#111111"><tr>
            <td width="50%" align="center" bgcolor="#C4D4E5"> <p><img src="duchang_img/daxiao.gif" width="303" height="237" border="0"><br>
            </td>
            <td width="50%" bgcolor="#ECF5FF"> <ul>
                <li><font color="#FF0000" class="c"><b>������</b></font> <a href="z_play.asp?menu=sz">���ڽ���&gt;&gt;</a> 
                  <br>
                  <br>
                  <br>
                <li><b><font color="#FF0000">�Ĵ�С</font></b> <a href="z_play.asp?menu=dx">���ڽ���&gt;&gt;</a>
                	<br>
                  <br>
                  <br>
                </li>
              </ul>
              <ul>
                <li><b><font color="#FF0000">�µ���</font></b> <a href="z_play.asp?menu=ds">���ڽ���&gt;&gt;</a>
                	<br>
                  <br>
                  <br>
                <li><b><font color="#FF0000">�ϻ���</font></b> <a href="z_tiger.asp">���ڽ���&gt;&gt;</a><br>
                
              </ul></td></tr></table></td></tr></table></center>

<%end sub

sub dx()%>
<br>

<form method="POST" action="z_play.asp?menu=dx1"><input type="hidden" name=menu value=dxover>

    <center>
<table borderColor="#a4b6d7" cellSpacing="0" width="400" border="1" height="54" style="border-collapse: collapse" cellpadding="6">
<tr align="center"><td width="190" colspan="2">
  <font size="3" style="font-size: 9pt"><b>
�� �� - �� �� С</b></font></td>
  <td width="183" colspan="2"><p>�����ע�� <b><font color="#CC0000">
  100</font> Ԫ</b></p>  
  </td>
</tr>
<tr align="center"><td width="134"><img src="duchang_img/run.gif" width="38" height="36"></td>
<td width="109" colspan="2"><img src="duchang_img/run.gif" width="38" height="36"></td>
  <td width="117"><img src="duchang_img/run.gif" width="38" height="36"></td>
</tr><tr><td bgcolor="#C4D4E5" width="386" colspan="4">
  <p align="center"><font style=font-size:9pt><b>
  &nbsp;&nbsp;��ѡ��</b></font></td>
</tr><tr><td align=center colspan="4" width="386" >
  <table border="0" cellspacing="0" cellpadding="0" width="386">
<tr align="center"><td width="194">
  <p align="right"><img src="duchang_img/big.gif"></td>
<td width="192">
<p align="left"><img src="duchang_img/small.gif"></td></tr><tr align="center">
<td width="194">
<p align="right"><input type="radio" name="dddimg" value="big" checked></td>
<td width="192">
<p align="left">
<input type="radio" name="dddimg" value="small"></td></tr></table></td></tr><tr>
  <td align=center colspan="4" width="386">��Ҫ��ע��
<input type="text" name="money" size="10" value="1"> &nbsp;<b> Ԫ</b></td></tr><tr>
  <td align=center bgColor=#C4D4E5 colspan="4" width="386"><input type="submit" value="��ע��������" name="B1" style="font-size: 12px">
</td>
</tr></table></center>
  </form><center>
���棺ÿ�ζ�Ǯ����������ֵ��� <font color="#CC0000">1</font> �� <br>
������һ����  <b><font color="#CC0000"><%=rs("userwealth")%></font></b> <b>Ԫ</b> ������Ϊ��ע<br>

<%end sub

sub ds()%>
<style>input {  font-size: 9pt; color: #000000; background-color: #f7f7f7; padding-top: 3px}</style>
</HEAD><BODY><br>
<form method="POST" action="z_play.asp?menu=ds1"><input type="hidden" name=menu value=dsover>
    <center>
<table border=1 cellspacing=0 cellpadding=3 width="350" bordercolor="#A4B6D7" style="border-collapse: collapse">
<tr align="center"><td width="33%"><p align="center"> 
  <font size="3" style="font-size: 9pt"><b>�� �� - �� �� ��</b></font></p>
  </td></tr>
<tr align="center"><td width="33%"><img src="duchang_img/run.gif"></td></tr><tr>
<td bgcolor="#A4B6D7" width="960"><font style=font-size:9pt><b>&nbsp;&nbsp;��ѡ��</b></font></td>
</tr><tr><td align=center><table border="0" cellspacing="0" cellpadding="0"><tr align="center">
<td width="17%"><img src="duchang_img/1.gif"></td><td width="17%"><img src="duchang_img/2.gif"></td>
<td width="16%"><img src="duchang_img/3.gif"></td><td width="17%"><img src="duchang_img/4.gif"></td>
<td width="17%"><img src="duchang_img/5.gif"></td><td width="16%"><img src="duchang_img/6.gif"></td>
</tr><tr align="center"><td width="17%"><input type="radio" name="dddimg" value="1" checked>
</td><td width="17%"><input type="radio" name="dddimg" value="2"></td><td width="16%">
<input type="radio" name="dddimg" value="3"></td><td width="17%"><input type="radio" name="dddimg" value="4">
</td><td width="17%"><input type="radio" name="dddimg" value="5"></td><td width="16%">
<input type="radio" name="dddimg" value="6"></td></tr></table></td></tr><tr><td align=center>
  ��Ҫ��ע��
<input type="text" name="money" size="10" value="1"> &nbsp;<b> Ԫ</b></td></tr><tr><td align=center bgColor=#A4B6D7><input type="submit" value="��ע��������" name="B1" style="font-size: 12px">
</td>
</tr></table>  </form>
���棺ÿ�ζ�Ǯ����������ֵ��� <font color="#CC0000">1</font> ��<br>
������һ���� <b><font color="#CC0000"><%=rs("userwealth")%></font> Ԫ</b> ������Ϊ��ע
<br>
<%end sub

sub sz()%>
<form method="POST" action="z_play.asp?menu=sz1"><input type="hidden" name=menu value=szover>
    <center>
<table border=1 cellspacing=0 cellpadding=6 width="400" style="border-collapse: collapse" bordercolor="#A4B6D7" height="115">
<tr>
  <td bgcolor="#C4D4E5" width="160" bordercolor="#A4B6D7" height="9" align="center">
  <font size="3" style="font-size: 9pt"><b>�ĳ� - ������</b></font></td>
  <td bgcolor="#C4D4E5" width="160" bordercolor="#A4B6D7" height="9" align="right">�����ע�� <b><font color="#CC0000">100</font> Ԫ</b></td>
</tr><tr bgColor=FAF0E2>
  <td align=center bgcolor="#FFFFFF" bordercolor="#A4B6D7" height="12" colspan="2">������һ���� <b><font color="#CC0000"><%=rs("userwealth")%></font>
Ԫ</b> ������Ϊ��ע</td></tr><tr>
  <td align=center bgcolor="#ECF5FF" bordercolor="#A4B6D7" height="20" colspan="2">��Ҫ��ע��
<input type="text" name="money" size="10" value="1"> <b> Ԫ</b></td></tr><tr>
  <td align=center bgColor=#C4D4E5 height="23" colspan="2"><input type="submit" value="��ע��������" name="B1" style="font-size: 12px"></td>
</tr></table>  
  </form>
���棺ÿ�ζ�Ǯ����������ֵ��� <font color="#CC0000">1</font> ��<br>

<%end sub



sub dx1()
dmoney=request("money")
dim ds
ds=request("dddimg")
Randomize
d1=int(rnd*6)+1
d2=int(rnd*6)+1
d3=int(rnd*6)+1
zs=d1+d2+d3
if zs>10 then 
	zds="��"
	zdss="big"
else
	zds="С"
	zdss="small"
end if
%>
<center><font style=font-size:9pt class="c"><b><font color="#0000CC">��Ѻ����</font><br><font color="#CC0000"></font><font color="#CC0000">
</font></b></font><img src=duchang_img/<%=ds%>.gif><font style=font-size:9pt class="c"><b><font color="#CC0000"></font></b></font>
    <center>
<table borderColor="#a4b6d7" cellSpacing="0" width="400" border="1" height="54" style="border-collapse: collapse" cellpadding="6">
<tr bgcolor="#F2D9B3"><td colspan="3" align="center" bgcolor="#C4D4E5"><font class="c"><b>���:</b></font><b><%=zs%>�㣬<%=zds%></b></td>
</tr><tr><td width="33%" align="center" bgcolor="#ECF5FF"><img src=duchang_img/<%=d1%>.gif></td>
  <td width="33%" align="center" bgcolor="#ECF5FF"><img src=duchang_img/<%=d2%>.gif></td>
<td width="34%" align="center" bgcolor="#ECF5FF"><img src=duchang_img/<%=d3%>.gif></td></tr>
<tr><td colspan="3" align="center" bgcolor="#ECF5FF"> <br>
<%
if ds=zdss then
%>
<p><font size="3">����~��Ӯ�ˣ� <b><font color="#CC0000"><%=dmoney%></font> Ԫ</b></font></p>
<table width="80%" border="0" cellspacing="0" cellpadding="5"><tr><td align="right"><b><font color="#000099">ׯ�ң�</font></b></td>
<td colspan="2">��������ѽ~����Ҫ����</td></tr><tr><td align="right"><b><font color="#CC0000">��˵��</font></b></td>
<td colspan="2"><a href="z_play.asp?menu=dx">�������ҽ�����������Ү~��</a></td></tr>
<tr><td align="right">��</td><td colspan="2"><a href=z_play.asp target=_top>���þ��գ��ǣ�����ɵ��~��</a></td>
</tr></table>
<%
dmoney=rs("userwealth")+dmoney
else

%>
<p><font size="3">�浹ù�����ˣ� <b><font color="#CC0000"><%=dmoney%></font> Ԫ</b></font></p>
<table width="80%" border="0" cellspacing="0" cellpadding="5"><tr><td align="right"><b><font color="#000099">ׯ�ң�</font></b></td>
<td colspan="2">�Ǻǻ�Ҫ���� �ж�δΪ��Ŷ~��</td></tr><tr><td align="right"><b><font color="#CC0000">��˵��</font></b></td>
<td colspan="2"><a href="z_play.asp?menu=dx">�������ޣ��ҾͲ�������ô��ù~��</a></td>
</tr><tr><td align="right">��</td><td colspan="2"><a href=z_play.asp target=_top>�����浹ù�������ˣ��´�����~��</a></td>
</tr></table>
<%
dmoney=rs("userwealth")-dmoney
end if
rs("userwealth")=dmoney
rs("usercp")=rs("usercp")-1
rs.update
%>
<hr width="250" SIZE="1">
�������У�<b><%=rs("userwealth")%></b>:Ԫ
<hr width="250" SIZE="1"></td></tr></table></center><font size="3"></font>
<%end sub



sub ds1()
dmoney=request("money")
dim ds
ds=cint(request("dddimg"))
Randomize
d1=fix(rnd*6)+1
%>
<br>
    <center>
    <table border="1" cellpadding="3" width="400" cellspacing="0" bordercolor="#A4B6D7" style="border-collapse: collapse">
<tr bgcolor="#F2D9B3"><td align="center" bgcolor="#A4B6D7"><font style=font-size:9pt class="c"><b>��Ѻ����</b></font></td>
  <td align="center" bgcolor="#A4B6D7"><font class="c"><b>���:</b></font><b><%=d1%>��</b></td>
</tr><tr><td align="center"><img src="duchang_img/<%=ds%>.gif"></td>
  <td align="center"><Img src="duchang_img/<%=d1%>.gif">
</span></a></td></tr>

<%
if ds=d1 then
dmoney=rs("userwealth")+dmoney*6
%>
<tr><td colspan="2" align="center" bgcolor="#ECF5FF" width="392"> <br><p><font size="3">����~��Ӯ�ˣ� <b><font color="#CC0000"><%=request("money")*6%></font> Ԫ</b></font></p>
<table width="80%" border="0" cellspacing="0" cellpadding="5"><tr><td align="right"><b><font color="#000099">ׯ�ң�</font></b></td>
<td colspan="2">��������ѽ~����Ҫ����</td></tr><tr><td align="right"><b><font color="#CC0000">��˵��</font></b></td>
<td colspan="2"><a href="z_play.asp?menu=ds">�������ҽ�����������Ү~��</a></td></tr>
<tr><td align="right">��</td><td colspan="2"><a href=z_play.asp target=_top>���þ��գ��ǣ�����ɵ��~��</a></td>
</tr></table>
<%else%>
<tr><td colspan="2" align="center" bgcolor="#ECF5FF" width="392"> <br><p><font size="3">�浹ù�����ˣ� <b><font color="#CC0000"><%=dmoney%> </font> Ԫ</b></font></p>
<table width="80%" border="0" cellspacing="0" cellpadding="5"><tr><td align="right"><b><font color="#000099">ׯ�ң�</font></b></td>
<td colspan="2">�Ǻǻ�Ҫ���� �ж�δΪ��Ŷ~��</td></tr><tr><td align="right"><b><font color="#CC0000">��˵��</font></b></td>
<td colspan="2"><a href=z_play.asp?menu=ds>�������ޣ��ҾͲ�������ô��ù~��</a></td>
</tr><tr><td align="right">��</td><td colspan="2"><a href=z_play.asp target=_top>�����浹ù�������ˣ��´�����~��</a></td>
</tr></table>
<%
dmoney=rs("userwealth")-dmoney
end if
rs("userwealth")=dmoney
rs("usercp")=rs("usercp")-1
rs.update
%>
<hr width="250" SIZE="1">
�������У�<b><font color="#CC0000"><%=rs("userwealth")%></font></b>:Ԫ
<hr width="250" SIZE="1"></td></tr></table></center>
<font size="3"></font>
<%end sub

sub sz1()
dim ks
dmoney=request("money")
Randomize
d1=fix(rnd*6)+1
d2=fix(rnd*6)+1
d3=fix(rnd*6)+1
b1=fix(rnd*6)+1
b2=fix(rnd*6)+1
b3=fix(rnd*6)+1
zj=d1+d2+d3
ni=b1+b2+b3
%>
<br>
<center>
<table border="1" cellspacing="0" cellpadding="6" width="400" bordercolor="#A4B6D7" style="border-collapse: collapse" height="314">
<tr><td colspan="3" bgcolor="#A4B6D7" height="12"><font style=font-size:9pt class="c"><b>
  &nbsp;�� �� - �� �� ��</b></font></td></tr><tr>
  <td colspan="3" align="center" bgcolor="#ECF5FF" height="12"><font color="#0000CC">��-ׯ�����ӣ�</font><%=zj%>
�� </td></tr><tr><td align="center" height="30"><img src=duchang_img/<%=d1%>.gif></td>
<td align="center" height="30"><img src=duchang_img/<%=d2%>.gif></td>
  <td align="center" height="30"><img src=duchang_img/<%=d3%>.gif></td>
</tr><tr><td colspan="3" align="center" bgcolor="#ECF5FF" height="12"><font color="#CC0000">��-�������ӣ�</font><%=ni%>
�� </td></tr><tr><td align="center" height="30"><img src=duchang_img/<%=b1%>.gif></td>
<td align="center" height="30"><img src=duchang_img/<%=b2%>.gif></td>
  <td align="center" height="30"><img src=duchang_img/<%=b3%>.gif></td>
</tr><tr bgcolor="#FAF0E2">
  <td colspan="3" align="center" bgcolor="#ECF5FF" height="116">
<%
if zj<ni then
%>
<table width="100%" border="0" cellspacing="0" cellpadding="5"><tr><td align="right" height="50">��</td>
<td colspan="2" height="50"><font size="3">&nbsp;����~��Ӯ�ˣ�<b><font color="#CC0000"><%=dmoney%></font> Ԫ</b></font></td></tr><tr>
<td align="right"><b><font color="#000099"> ׯ�ң�</font></b></td><td colspan="2">&nbsp;��������ѽ~����Ҫ���� </td></tr><tr><td align="right"><b><font color="#CC0000">��Ҫ��</font></b></td>
<td colspan="2">&nbsp;<a href="z_play.asp?menu=sz">�������ҽ�����������Ү~��</a></td></tr><tr><td align="right">��</td>
<td colspan="2">&nbsp;<a href=z_play.asp target=_top>���þ��գ��Ǻǣ����˿�~��</a></td></tr></table>
<%
dmoney=rs("userwealth")+dmoney
else
%>
<center>
<table width="100%" border="0" cellspacing="0" cellpadding="5" style="border-collapse: collapse" bordercolor="#111111"><tr>
<td height="50">��</td><td colspan="2" height="50"><p><font size="3">&nbsp;�浹ù�����ˣ� <b><font color="#CC0000"><%=dmoney%></font> Ԫ</b></font></p></td></tr><tr>
<td align="right"><b><font color="#000099"> ׯ�ң�</font></b></td><td colspan="2">&nbsp;�Ǻ�~��Ҫ���� �ж�δ����Ŷ~��</td></tr><tr><td align="right" rowspan="2"><b><font color="#CC0000"> ��Ҫ��</font></b></td>
<td colspan="2">&nbsp;<a href="z_play.asp?menu=sz">�������ޣ��ҾͲ�������ô��ù~��</a></td></tr><tr><td colspan="2">&nbsp;<a href=z_play.asp target=_top>�����浹ù����~�ˣ��´�����~��</a></td></tr></table>
</center>
<%
dmoney=rs("userwealth")-dmoney
end if

rs("userwealth")=dmoney
rs("usercp")=rs("usercp")-1
rs.update
%>
<hr size="1" width="250">
�������У�<font color="#CC0000"><b><%=rs("userwealth")%></b></font><b> ��Ԫ
<hr size="1" width="250"></td></tr></table></center>
<%end sub
rs.close
%>


<%
sub error2(message)
%>
<script>alert('<%=message%>');history.back();</script><script>window.close();</script>
<%
response.end
end sub
%>

<%
response.write "</CENTER></td></tr></table>"
end if
call activeonline()
call footer()
%>
