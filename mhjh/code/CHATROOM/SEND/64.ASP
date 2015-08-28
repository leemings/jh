<%
'红包
function hongbao(un,st,gr,mg)
if gr<15 then
Response.Write "<script language=javascript>{alert('你的等级太低了，需要15级才能送红包！');}</script>"
 Response.End
 exit function
end if
if not isnumeric(mg) then
Response.Write "<script language=javascript>{alert('##想你的确是派红包"&mg&"给%%吗？');}</script>"
  Response.End
  exit function
else
mg=abs(int(clng(mg)))
end if
if mg<100000 then
Response.Write "<script language=javascript>{alert('你这么小气啊！100000以内也送得出!');}</script>"
 Response.End
 exit function
end if
if mg>20000000 then
Response.Write "<script language=javascript>{alert('你也太爽快了吧！一出手就" & mg & "两白银！红包一次最多是2千万两哦!');}</script>"
 Response.End
 exit function
end if
if mg>=2000000 then
bao="bao2"
else
bao="bao"
end if 
rst.Open "select 银两,操作时间,配偶,性别 from 用户 where 姓名='"&un&"' and 银两>="&mg,conn
if rst.EOF or rst.BOF then
hongbao="<font color=FF0000>【派红包】</font>%%笑着对##说：'先谢谢你了，等你有了"&mg&"两银子再来送我红包好吗？'"
rst.close
else
sj=DateDiff("s",rst("操作时间"),now())
pei=rst("配偶")
xing=rst("性别")
if xing="男" then 
ch="哥哥"
else
ch="姐姐"
end if
rst.close
if sj<1 then
 ss=1-sj
 Response.Write "<script language=javascript>{alert('请你等上["& ss &"]秒再操作！');}</script>"
 Response.End
exit function
 end if
conn.execute "update 用户 set 银两=银两+"&mg&" where 姓名='"&st&"'"
conn.execute "update 用户 set 操作时间=now(),银两=银两-"&mg&" where 姓名='"&un&"'"
if pei<>"无" then
hongbao="<font color=FF0000>【派红包】</font>##笑着对%%说：<a href=../hb.asp?yl="&mg&"&dx="&un&"&ddd="&st&" target=optfrm><img border=0 src=../chatroom/image/"&bao&".gif title=打开红包></a>这是我和"&pei&"给你的大红包！抓稳点,别丢了,祝你天天开心,新年快乐！<bgsound src=../mid/thanks.wav loop=1>"
else
hongbao="<font color=FF0000>【派红包】</font>##笑着对%%说：<a href=hb.asp?yl="&mg&"&dx="&un&"&ddd="&st&" target=optfrm><img border=0 src=../chatroom/image/"&bao&".gif title=打开红包></a>这是我给你的大红包！唉,"&ch&"还没有结婚呢,不过还是祝福你快高长大！<bgsound src=../mid/thanks.wav loop=1>"
end if
end if
end function
%>