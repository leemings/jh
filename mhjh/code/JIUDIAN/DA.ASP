<!--#include file="dadata.asp"-->
<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
nickname=session("yx8_mhjh_username")
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
connstr=Application("yx8_mhjh_connstr")
conn.open connstr
sql="select * from 用户 where 姓名='" & nickname & "'"
set rs=conn.execute(sql)
If rs.bof Or rs.EOF Then
    response.write "你不是江湖中人，不能进参加宴会"
    conn.Close
    response.End
End If

dj = rs("等级")
shenf = rs("身份")
peiou = rs("配偶")
xiaoq = rs("小妾")
mypai = rs("门派")
sharen=rs("杀人")
If dj < 1 Then
response.write "你还是江湖小辈，不能参加宴会"
conn.Close
response.End
End If
If sharen>9 Then
response.write "通缉犯来参加宴会了,没人理你,你只好离开"
conn.Close
response.End
End If
If rs("体力")>rs("等级")*1000000 Then
response.write "你已经吃的这么饱了,还想来这里拣便宜?回去消化消化吧"
conn.Close
response.End
End If
If rs("内力")>rs("等级")*100000 Then
response.write "你的力气足的都要爆炸了,还想来这里拣便宜?回去消化消化吧"
conn.Close
response.End
End If
id = request("id")
sql = "select * from 宴会列表 where ID=" & id
Set rs = connt.execute(sql)
yyou = rs("拥有者")
mc = rs("宴会名")
nl = rs("内力")
tl = rs("体力")
jb = rs("资质")
lx = rs("类型")
ls = rs("数量")
shihe = rs("说明")

errjoin = ""

Select Case shihe
Case "官府"
If dj < 6 Or mypai <> "官府" Then errjoin = "唉，人家请的是官府中人，你就别进去了。"
Case "首领"
If shenf <> "掌门" And shenf <> "副掌门" And shenf <> "长老" And shenf <> "护法" And shenf <> "堂主" Then errjoin = "这里都是帮派中有头有脸的人啊，堂主以下的人不要乱冯！"
Case "情侣"
If peiou <> yyou And xiaoq <> yyou Then errjoin = "你不是" & yyou & "的配偶或情人啊，不能打扰人家好吗？"
Case "所有"
If dj < 2 Then errjoin = "你还是江湖小辈，不能参加宴会。"
Case Else
If mypai <> shihe Then errjoin = "你不是" & shihe & "派的人啊，不能参加他们的酒宴。"
End select

If errjoin <> "" Then
response.write errjoin
conn.Close
connt.Close
Set conn = Nothing
Set connt = Nothing
response.End
End If


If ls < 1 Then
    sql1 = "delete * from 宴会列表 where ID=" & id
    connt.execute sql1
    sql1 = "delete * from 宴会者 where 宴会名=" & id
    connt.execute sql1
    response.write "你来晚，宴会已经结束"
    conn.Close
    connt.Close
    Set conn = Nothing
    Set connt = Nothing
    response.End
Else
    sql = "select * from 宴会者 where 参加者='" & nickname & "' and 宴会名=" & id
    Set rs = connt.execute(sql)
        If rs.EOF Or rs.bof Then
            sql2 = "insert into 宴会者(参加者,宴会名) values ('" & nickname & "' , " & id & " )"
            connt.execute sql2
            sql1="update 用户 set 内力=内力+"&nl&",体力=体力+"&tl&",资质=资质+"&jb&" where 姓名='"&nickname&"' "
            conn.execute sql1
            sql1 = "update 宴会列表 set 数量=数量-1 where ID=" & id
            connt.execute sql1
            
                If ls = 1 Then
                    sql1 = "delete * from 宴会列表 where ID=" & id
                    connt.execute sql1
                    sql1 = "delete * from 宴会者 where 宴会名=" & id
                    connt.execute sql1
                End If
            
            conn.Close
            connt.Close
            Set conn = Nothing
            Set connt = Nothing
            Set rs = Nothing
dim newtalkarr(600) 
   talkarr=Application("yx8_mhjh_talkarr") 
		talkpoint=Application("yx8_mhjh_talkpoint") 
		j=1 
		for i=11 to 600 step 10 
			newtalkarr(j)=talkarr(i) 
			newtalkarr(j+1)=talkarr(i+1) 
			newtalkarr(j+2)=talkarr(i+2) 
			newtalkarr(j+3)=talkarr(i+3) 
			newtalkarr(j+4)=talkarr(i+4) 
			newtalkarr(j+5)=talkarr(i+5) 
			newtalkarr(j+6)=talkarr(i+6) 
			newtalkarr(j+7)=talkarr(i+7) 
			newtalkarr(j+8)=talkarr(i+8) 
			newtalkarr(j+9)=talkarr(i+9) 
			j=j+10 
		next 
		newtalkarr(591)=talkpoint+1 
		newtalkarr(592)=1 
		newtalkarr(593)=0 
		newtalkarr(594)=username 
		newtalkarr(595)="大家" 
		newtalkarr(596)="" 
		newtalkarr(597)="000000" 
		newtalkarr(598)="000000" 
		newtalkarr(599)="<font color=red>【江湖消息】</font><font color=blue>"&nickname&"参加了"&yyou&"在魔幻大酒店举行的※<font color=red>"&mc&"</font>※宴会,内力增加"&nl&",体力增加"&tl&",资质增加"&jb&" </font>" 
newtalkarr(600)=Session("yx8_mhjh_userchatroomsn")
Application.Lock
Application("yx8_mhjh_talkpoint")=talkpoint+1
Application("yx8_mhjh_talkarr")=newtalkarr
Application.UnLock
erase newtalkarr
erase talkarr
        %>
        <%else
        conn.Close
        connt.Close
        Set conn = Nothing
        Set connt = Nothing
        %>
            <script>
            alert('你已参加过这个宴会了，怎么还来啊。');
            window.close();
            </script>
        <%end if%>
<%end if%>

<html>

<head>
<title>参加宴会成功</title>
<style type="text/css">
<!--
table {  border: #000000; border-style: solid; border-top-width: 1px; border-right-width: 1px; border-bottom-width: 1px; border-left-width: 1px}
font {  font-size: 12px}
.unnamed1 {  font-size: 9pt}
-->
</style>

</head>

<body bgcolor="#CCCCCC">
<p>&nbsp;</p>
<table width="90%" border="0" height="156" bordercolor="#330033" cellspacing="0" cellpadding="0" align="center">
  <tr> 
    <td height="17" bgcolor="#996633" align="center"><font color="#FFFFFF">参加酒宴成功</font></td>
  </tr>
  <tr bgcolor="#66FF66"> 
    <td align="center" height="157"> 
      <p><font> <img src="jd.files/ka1.gif" width="204" height="251"></font></p>
    </td>
  </tr>
  <tr bgcolor="#0033CC"> 
    <td align="center" height="20" class="unnamed1"><b><font color="#FF3333">你经过了一番狼吞虎咽，喝的上面一个模样，增加内力<%=nl%>，体力<%=tl%>,领到资质<%=jb%></font></b></td>
  </tr>
</table>
<p><p align=center><a href="#" onclick=javascript:history.back(); onmouseover="window.status='返回';return true;" onmouseout="window.status='';return true;">返回</a></p></p>
</body>
</html>
