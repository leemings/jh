<!--#include file="conn.asp"-->
<!--#include file="inc/const.asp"-->
<%
'=========================================================
' File: zlyx.asp,doubletake.asp,FIVE.asp,klondike.asp,minesweeper.asp,RUBIK.asp,TETRIS.asp
' For DVbbs5.0 final or DVbbs6.0
' Date: 2003-7-20
' Script Written by ϧϦ��
' Web: http://www.51ehome.net
' QQ: 13326598
'=========================================================
stats="������Ϸ"
call nav()
call head_var(2,0,"","")
if not founduser then
   Errmsg=Errmsg+"<br>"+"<li>ֻ�л�Ա�ſ��Բμ�������������<a href=login.asp>��½</a>����ͬ����Ա��ϵ��"
   call dvbbs_error()
else%>
<table cellpadding=3 cellspacing=1 border=0 align="center" class=tableborder1>
  <TR> 
    <th id=tabletitlelink > <a href="game/five.asp" target="frm">������</a></Th>
    <th id=tabletitlelink ><a href="game/doubletake.asp" target="frm">�����</a></Th>
    <th id=tabletitlelink ><a href="game/klondike.asp" target="frm">������Ϸ</a></Th>
    <th id=tabletitlelink ><a href="game/minesweeper.asp" target="frm">��������</a></Th>
    <th id=tabletitlelink ><a href="game/rubik.asp" target="frm">ħ��</a></Th>
    <th id=tabletitlelink ><a href="game/tetris.asp" target="frm">����˹����</a></Th>
  </TR>
</TABLE><br>
<table cellpadding=3 cellspacing=1 border=0 align="center" class=tableborder1>
  <TR>
    <TD class=tablebody2 align="center" valign="center">
      <iframe name=frm style="HEIGHT: 490; WIDTH: 100%;" src="game/tetris.asp" frameborder=0></iframe>
    </TD>
  </TR>
</TABLE>
<%
end if
call footer()
%>