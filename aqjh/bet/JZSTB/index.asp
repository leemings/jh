<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"

aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../../error.asp?id=440"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
Function computerChooses()
Dim randomNum
Dim choice
randomize
randomNum = int(rnd*15)+1
If randomNum = 1 OR randomNum = 3 OR randomNum = 7 OR randomNum = 8 OR randomNum = 15 OR randomNum = 12 Then
choice = "R"
ElseIf randomNum = 2 OR randomNum = 6 OR randomNum = 11 OR randomNum = 13 Then
choice = "S"
Else
choice = "P"
End If
computerChooses = choice
End Function
Sub determineWinner(playerChoice, computerChoice)
Const Rock = "R"
Const Scissor = "S"
Const Paper = "P"
Dim tempPlayer, tempComputer
If playerChoice = Rock Then
If computerChoice = Scissor Then
conn.Execute "update 用户 set 银两=银两+50 where 姓名='"&aqjh_name&"'"
%>
<title>剪刀 石头 布</title><body background="../../jhimg/bk_hc3w.gif">
<P><CENTER>
<IMG SRC="jhimg/shit.gif"><BR>
你的石头捶扁了计算机的剪子！您得到50两银子。</CENTER>
<%
End If
ElseIf playerChoice = Scissor Then
If computerChoice = Paper Then
conn.Execute  "update 用户 set 银两=银两+50 where 姓名='"&aqjh_name&"'"
%>
<P><CENTER>
<IMG SRC="jhimg/jianz.gif"><BR>
你的剪子剪破了计算机的布！您得到50两银子。</CENTER>
<%
End If
ElseIf playerChoice = Paper Then
If computerChoice = Rock Then
conn.Execute  "update 用户 set 银两=银两+50 where 姓名='"&aqjh_name&"'"

%>
<P><CENTER>
<IMG SRC="jhimg/bu.gif"><BR>
你的布包住了计算机的石头！您得到50两银子。</CENTER>
<%
End If
ElseIf playerChoice = computerChoice Then
%>
<p><CENTER>
<IMG SRC="jhimg/tie.gif"><BR>
看来你们想到一起了。</CENTER>
<%
End If
If computerChoice = Rock Then
If playerChoice = Scissor Then
conn.Execute  "update 用户 set 银两=银两-50 where 姓名='"&aqjh_name&"'"

%>
<P><CENTER>
<IMG SRC="jhimg/shit.gif"><BR>
计算机的石头捶扁了你的剪子！您损失50两银子。</CENTER>
<%
ElseIf playerChoice = computerChoice Then
%>
<P><CENTER>
<IMG SRC="jhimg/tie.gif"><BR>
看来你们想到一起了。</CENTER>
<%
End If
ElseIf computerChoice = Scissor Then
If playerChoice = Paper Then
conn.Execute  "update 用户 set 银两=银两-50 where 姓名='"&aqjh_name&"'"

%>
<P><CENTER>
<IMG SRC="jhimg/jianz.gif"><BR>
计算机的剪子剪破了你的布！您损失50两银子。</CENTER>
<%
ElseIf playerChoice = computerChoice Then
%>
<P><CENTER>
<IMG SRC="jhimg/tie.gif"><BR>
看来你们想到一起了。</CENTER>
<%
End If
ElseIf computerChoice = Paper Then
If playerChoice = Rock Then
conn.Execute  "update 用户 set 银两=银两-50 where 姓名='"&aqjh_name&"'"
%>
<P><CENTER>
<IMG SRC="jhimg/bu.gif"><BR>
计算机的布包住了你的石头！您损失50两银子。</CENTER>
<%
ElseIf playerChoice = computerChoice Then
%>
<P><CENTER>
<IMG SRC="jhimg/tie.gif"><BR>
看来你们想到一起了。</CENTER>
<%
End If
ElseIf computerChoice = playerChoice Then
%>
<P>
<CENTER>
<font color="#0000FF"><IMG SRC="jhimg/tie.gif"><BR>
看来你们想到一起了。</font>
</CENTER>
<%
End If
End Sub
Sub playGame()
%>
<P align="center"> <font color="#0000FF" size="5" face="黑体">剪刀 石头 布</font>
<table width="390" border="1" cellspacing="10" cellpadding="2" align="center" bgcolor="#336633">
<tr bgcolor="#FFCCCC">
<td width="178" height="193" bgcolor="#CCCCFF">
<div align="center"><font size="2">请 选 择:</font> </div>
<form action="index.asp?action=winner" method="post">
<div align="center">
<table width="178" cellspacing="0" border="1" bgcolor="#CCCCCC" cellpadding="0" bordercolorlight="#000000" bordercolordark="#FFFFFF">
<tr valign=top>
<td width="75">
<div align="center"><font size="2">石 头</font></div>
</td>
<td width="63">
<div align="center">
<input type="radio" name="playerSelect" value="R">
</div>
</td>
</tr>
<tr valign=top>
<td width="75">
<div align="center"><font size="2">剪 子</font></div>
</td>
<td width="63">
<div align="center">
<input type="radio" name="playerSelect" value="S">
</div>
</td>
</tr>
<tr valignn=top>
<td width="75">
<div align="center"><font size="2">布</font></div>
</td>
<td width="63">
<div align="center">
<input type="radio" name="playerSelect" value="P">
</div>
</td>
</tr>
</table>
<input type="submit" value="开 始" name="submit">
</div>
</form>
<div align="center"> </div>
<p align="center"><font size="2"><a href="../betindex.asp">不玩了</a></font><br>
</td>
<td width="211" bgcolor="#FFFFFF" height="193">
<div align="center"><font size="2">《剪子、石头、布》游戏</font><font color="#33CCCC"><br>
祝你胜利！<br>
</font><img src="jhimg/bu.gif"><img src="jhimg/shit.gif"></div>
</td>
</tr>
</table>
<p align="center"><%
End Sub
Sub playAgain()
%>
<p>
<center>
您要再玩一边么？
</center>
<br>
<center>
<a href="index.asp"><font size="2">是的</font></a> <font size="2"><a href="../betindex.asp">不玩了</a>
</font>
</center>
<H1 align="center"><BR>
</h1>
<CENTER>
</CENTER>
<%
End Sub
Sub endGame()
Response.Buffer = true
End Sub
Dim player, computer
Dim gameAction
gameAction = Request.QueryString("action")
Select Case gameAction
Case "winner"
player = Request.Form("playerSelect")
computer = computerChooses
		
determineWinner player, computer
Response.Write "<BR><BR>"
playAgain

Case "again"
playAgain
Case "gameover"
endGame
Case Else
playGame
End Select
%>
<div align="center">版权所有≮爱情江湖总站≯</div>
</body>
