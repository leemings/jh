<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"

sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
if sjjh_name="" then Response.Redirect "../../error.asp?id=440"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
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
conn.Execute "update �û� set ����=����+50 where ����='"&sjjh_name&"'"
%>
<title>���� ʯͷ ��</title><body background="../../jhimg/bk_hc3w.gif">
<P><CENTER>
<IMG SRC="jhimg/shit.gif"><BR>
���ʯͷ�����˼�����ļ��ӣ����õ�50�����ӡ�</CENTER>
<%
End If
ElseIf playerChoice = Scissor Then
If computerChoice = Paper Then
conn.Execute  "update �û� set ����=����+50 where ����='"&sjjh_name&"'"
%>
<P><CENTER>
<IMG SRC="jhimg/jianz.gif"><BR>
��ļ��Ӽ����˼�����Ĳ������õ�50�����ӡ�</CENTER>
<%
End If
ElseIf playerChoice = Paper Then
If computerChoice = Rock Then
conn.Execute  "update �û� set ����=����+50 where ����='"&sjjh_name&"'"

%>
<P><CENTER>
<IMG SRC="jhimg/bu.gif"><BR>
��Ĳ���ס�˼������ʯͷ�����õ�50�����ӡ�</CENTER>
<%
End If
ElseIf playerChoice = computerChoice Then
%>
<p><CENTER>
<IMG SRC="jhimg/tie.gif"><BR>
���������뵽һ���ˡ�</CENTER>
<%
End If
If computerChoice = Rock Then
If playerChoice = Scissor Then
conn.Execute  "update �û� set ����=����-50 where ����='"&sjjh_name&"'"

%>
<P><CENTER>
<IMG SRC="jhimg/shit.gif"><BR>
�������ʯͷ��������ļ��ӣ�����ʧ50�����ӡ�</CENTER>
<%
ElseIf playerChoice = computerChoice Then
%>
<P><CENTER>
<IMG SRC="jhimg/tie.gif"><BR>
���������뵽һ���ˡ�</CENTER>
<%
End If
ElseIf computerChoice = Scissor Then
If playerChoice = Paper Then
conn.Execute  "update �û� set ����=����-50 where ����='"&sjjh_name&"'"

%>
<P><CENTER>
<IMG SRC="jhimg/jianz.gif"><BR>
������ļ��Ӽ�������Ĳ�������ʧ50�����ӡ�</CENTER>
<%
ElseIf playerChoice = computerChoice Then
%>
<P><CENTER>
<IMG SRC="jhimg/tie.gif"><BR>
���������뵽һ���ˡ�</CENTER>
<%
End If
ElseIf computerChoice = Paper Then
If playerChoice = Rock Then
conn.Execute  "update �û� set ����=����-50 where ����='"&sjjh_name&"'"
%>
<P><CENTER>
<IMG SRC="jhimg/bu.gif"><BR>
������Ĳ���ס�����ʯͷ������ʧ50�����ӡ�</CENTER>
<%
ElseIf playerChoice = computerChoice Then
%>
<P><CENTER>
<IMG SRC="jhimg/tie.gif"><BR>
���������뵽һ���ˡ�</CENTER>
<%
End If
ElseIf computerChoice = playerChoice Then
%>
<P>
<CENTER>
<font color="#0000FF"><IMG SRC="jhimg/tie.gif"><BR>
���������뵽һ���ˡ�</font>
</CENTER>
<%
End If
End Sub
Sub playGame()
%>
<P align="center"> <font color="#0000FF" size="5" face="����">���� ʯͷ ��</font>
<table width="390" border="1" cellspacing="10" cellpadding="2" align="center" bgcolor="#336633">
<tr bgcolor="#FFCCCC">
<td width="178" height="193" bgcolor="#CCCCFF">
<div align="center"><font size="2">�� ѡ ��:</font> </div>
<form action="index.asp?action=winner" method="post">
<div align="center">
<table width="178" cellspacing="0" border="1" bgcolor="#CCCCCC" cellpadding="0" bordercolorlight="#000000" bordercolordark="#FFFFFF">
<tr valign=top>
<td width="75">
<div align="center"><font size="2">ʯ ͷ</font></div>
</td>
<td width="63">
<div align="center">
<input type="radio" name="playerSelect" value="R">
</div>
</td>
</tr>
<tr valign=top>
<td width="75">
<div align="center"><font size="2">�� ��</font></div>
</td>
<td width="63">
<div align="center">
<input type="radio" name="playerSelect" value="S">
</div>
</td>
</tr>
<tr valignn=top>
<td width="75">
<div align="center"><font size="2">��</font></div>
</td>
<td width="63">
<div align="center">
<input type="radio" name="playerSelect" value="P">
</div>
</td>
</tr>
</table>
<input type="submit" value="�� ʼ" name="submit">
</div>
</form>
<div align="center"> </div>
<p align="center"><font size="2"><a href="../betindex.asp">������</a></font><br>
</td>
<td width="211" bgcolor="#FFFFFF" height="193">
<div align="center"><font size="2">�����ӡ�ʯͷ��������Ϸ</font><font color="#33CCCC"><br>
ף��ʤ����<br>
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
��Ҫ����һ��ô��
</center>
<br>
<center>
<a href="index.asp"><font size="2">�ǵ�</font></a> <font size="2"><a href="../betindex.asp">������</a>
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
<div align="center"><%=Application("sjjh_chatroomname")%>��Ȩ����</div>
</body>
