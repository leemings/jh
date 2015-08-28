<%
function qiuhun(un,st,mg)
rst.Open "select 性别,配偶,道德,形态 from 用户 where 姓名='"&st&"'",conn
if rst("配偶")<>"无" then
Response.Write "<script language=JavaScript>{alert('对方已经是有家室的人了呀！');}</script>"
Response.End
exit function 
end if
if rst("道德")<100 then
Response.Write "<script language=JavaScript>{alert('对方道德低于100，你也要娶?');}</script>"
Response.End
exit function 
end if
if st="武术家" or st="魔法师" or st="爱情鸟" then
Response.Write "<script language=JavaScript>{alert('你想和机器人结婚啊?除非你也变成机器人!');}</script>"
Response.End
exit function 
end if
sex=rst("性别")
rst.close
rst.Open "select 性别,配偶,道德,银两 from 用户 where 姓名='"&un&"'",conn
if rst("性别")=sex then
Response.Write "<script language=JavaScript>{alert('同性恋患者?送你进疯人院要吗?');}</script>"
Response.End
exit function 
end if
if rst("道德")<100 then
Response.Write "<script language=JavaScript>{alert('道德低于1000不能结婚!');}</script>"
Response.End
exit function 
end if
if rst("银两")<1000000 then
Response.Write "<script language=JavaScript>{alert('结婚需要100万银两!');}</script>"
Response.End
exit function 
end if
if rst("配偶")<>"无" then
Response.Write "<script language=JavaScript>{alert('你已经是有家室的人了呀！');}</script>"
Response.End
exit function 
end if
qiuhun="<font color=FF0000>【求婚】</font><font color=red>##慢慢地走到向%%面前,温柔地说："&mg&"!</font><input type=button value='同意' onclick=""javascript:window.open('qiuhunok.asp?name="&st&"','sendfrm','width=380,height=210')"" style=""background-color:#86A231;color:FFFFFF;border: 1 double"" onMouseOver=""this.style.color='FFFF00'"" onMouseOut=""this.style.color='FFFFFF'"" name=""button2231""><input type=button value='拒绝' onclick=""javascript:window.open('qiuhunno.asp?name="&st&"','sendfrm','width=380,height=210')"" style=""background-color:#86A231;color:FFFFFF;border: 1 double"" onMouseOver=""this.style.color='FFFF00'"" onMouseOut=""this.style.color='FFFFFF'"" name=""button223"">"	
Application.Lock
application("yx8_mhjh_qiuhun")=st&"|"&un
Application.UnLock
end function
%>
z