<%
function burglemp(un,st)
randomize()
odds=rnd()*5000
rst.Open "select * from 用户 where 姓名='"&st&"'",conn
if rst.EOF or rst.BOF then
burglemp="<font color=FF0000>【吸星大法】</font>##你想对谁吸星大法啊，"&newaberrant(un,"大家","中毒",360)
elseif rst("保护")=True then
burglemp="<font color=FF0000>【吸星大法】</font>##太傻了,居然在%%已经开启保护的情况下实施吸星大法，被%%的保护真气震得晕头转向，"&newaberrant(un,"大家","中毒",360)
elseif rst("内力")<100000 then
burglemp="<font color=FF0000>【吸星大法】</font>##太不道德了,%%连10万内力都没有了还实施吸星大法，被大家打晕了，"&newaberrant(un,"大家","疯狂",180)
else
rst.Close
rst.Open "select 积分,保护 from 用户 where 姓名='"&un&"'",conn
uhp=rst("积分")
bh=rst("保护")
rst.Close
if uhp<10000 then
burglemp="<font color=FF0000>【吸星大法】</font>##还是个小孩子,居然敢在公开场合对%%实施吸星大法，有违江湖公道，被大家打傻了，"&newaberrant(un,"大家","麻痹",180)
elseif bh=true then
burglemp="<font color=FF0000>【吸星大法】</font>##处于保护中,居然蒙着眼睛对%%实施吸星大法，真是愚蠢啊，被大家打傻了，"&newaberrant(un,"大家","麻痹",180)
elseif odds>1500 then
if odds>uhp then odds=uhp
conn.execute "update 用户 set 体力=体力-"&odds&",道德=道德-"&odds&" where 姓名='"&un&"'"
burglemp="<font color=FF0000>【吸星大法】</font>##因在公众场合对%%滥施吸星大法，有违江湖公道，被大家打得动弹不得！体力和道德下降"&odds&"，"&newaberrant(un,"大家","麻痹",120)
elseif odds>0 and odds<1500 then
odds=clng(odds*0.5)
if odds>uhp then odds=uhp
conn.execute "update 用户 set 内力=内力+"&odds*1000&" where 姓名='"&un&"'"
conn.execute "update 用户 set 内力=内力-"&odds*1000&" where 姓名='"&st&"'"
burglemp="<font color=FF0000>【吸星大法】</font>##偷偷使用了江湖中最忌讳的吸星大法，吸走了%%"&odds*100&"的内力<bgsound src='../mid/xixing.wav' loop=1>"
else
conn.execute "update 用户 set 体力=体力-"&odds&",道德=道德-"&odds&" where 姓名='"&un&"'"
burglemp="<font color=FF0000>【吸星大法】</font>##因在公众场合对%%滥施吸星大法，有违江湖公道，被大家打得动弹不得！体力和道德下降"&odds*2&"，"&newaberrant(un,"大家","麻痹",120)
rst.Close
end if
end if
end function
%>
