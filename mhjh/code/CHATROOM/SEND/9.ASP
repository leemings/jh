<%
function qiqi(un,st,mg)
if mg="" then
response.write "<script language='javascript'>alert('请输入抢亲的理由');</script>"
response.end
end if
rst.Open "select 等级,性别,配偶 from 用户 where 姓名='"&st&"'",conn
if rst("等级")>5 or rst("性别")<>"女" or rst("配偶")<>"无" then
response.write "<script language='javascript'>alert('对方不符合抢亲的条件,对方需要等级>5、性别=女、配偶=无');</script>"
response.end
else
rst.Close
rst.Open "select 等级,性别,小妾,精力 from 用户 where 姓名='"&un&"'",conn
if rst("等级")>8 and rst("性别")="男" and rst("小妾")="无" then
randomize()
odds=rnd()*10000
if odds>1999 then
uhp=rst("精力")	
rst.Close
if odds*0.2>uhp then odds=uhp
conn.execute "update 用户 set 体力=体力-"&odds*100&",道德=道德-"&odds*100&" where 姓名='"&un&"'"
qiqi="<font color=FF0000>【抢亲】</font>##因对%%见色起心，大叫一声"&mg&"冲上前调戏,有违江湖公道，被大家打成了花痴！体力和道德下降"&clng(odds*100)&"，"&newaberrant(un,"大家","疯狂",180)
else
conn.execute "update 用户 set 小妾='"&st&"' where 姓名='"&un&"'"
conn.execute "update 用户 set 配偶='"&un&"' where 姓名='"&st&"'"
qiqi="<font color=FF0000>【抢亲】</font>##因对%%见色起心，口中叫着"&mg&"经过百折不饶的努力,终于将%%抢回了家里当了老婆!<bgsound src='../mid/Mg.wav' loop=1>"
end if
else
response.write "<script language='javascript'>alert('你不符合抢亲的条件,你需要8级以上,而且没小妾!');</script>"
response.end
rst.close
set rst=nothing
end if
end if
end function
%>