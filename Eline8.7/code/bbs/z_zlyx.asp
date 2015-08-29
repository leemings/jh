<!--#include file="conn.asp"-->
<!--#include file="inc/const.asp"-->
<%
'=========================================================
' File: zlyx.asp,doubletake.asp,FIVE.asp,klondike.asp,minesweeper.asp,RUBIK.asp,TETRIS.asp
' For DVbbs5.0 final or DVbbs6.0
' Date: 2003-7-20
' Script Written by 惜夕湾
' Web: http://www.51ehome.net
' QQ: 13326598
'=========================================================
stats="智力游戏"
call nav()
call head_var(2,0,"","")
if not founduser then
   Errmsg=Errmsg+"<br>"+"<li>只有会员才可以参加智力竞赛，请<a href=login.asp>登陆</a>或者同管理员联系。"
   call dvbbs_error()
else%>
<table cellpadding=3 cellspacing=1 border=0 align="center" class=tableborder1>
  <TR> 
    <th id=tabletitlelink > <a href="game/five.asp" target="frm">五子棋</a></Th>
    <th id=tabletitlelink ><a href="game/doubletake.asp" target="frm">记忆板</a></Th>
    <th id=tabletitlelink ><a href="game/klondike.asp" target="frm">接龙游戏</a></Th>
    <th id=tabletitlelink ><a href="game/minesweeper.asp" target="frm">超级工兵</a></Th>
    <th id=tabletitlelink ><a href="game/rubik.asp" target="frm">魔方</a></Th>
    <th id=tabletitlelink ><a href="game/tetris.asp" target="frm">俄罗斯方块</a></Th>
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