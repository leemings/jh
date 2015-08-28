<%
function seeequip(un,st)
rst.Open "select 性别,门派,身份,等级,配偶,师傅,徒弟,结拜兄弟,头像,头衔,国家,职位 from 用户 where 姓名='"&st&"'",conn
seeequip="<img src=../chatroom/Image/gb.gif>##得到关于%%的情报。<img border=0 src='"&rst("头像")&"' width='20' height='20'> 性别："&rst("性别")&"；国家："&rst("国家")&"；职位："&rst("职位")&"；门派："&rst("门派")&"；身份："&rst("身份")&"；配偶："&rst("配偶")&"；师傅："&rst("师傅")&"；徒弟："&rst("徒弟")&"；结拜："&rst("结拜兄弟")&"；外号："&rst("头衔")&"；"
if rst("等级")<3 then
seeequip=seeequip&"看起来象个新手。"
elseif rst("等级")<10 and  rst("等级")>=3 then
seeequip=seeequip&"看起来有点儿实力。"
elseif rst("等级")<15 and  rst("等级")>=10 then
seeequip=seeequip&"看起来来头不小。"
elseif rst("等级")<20 and  rst("等级")>=15 then
seeequip=seeequip&"看起来是个高手。"
elseif rst("等级")<25 and  rst("等级")>=20 then
seeequip=seeequip&"看起来是个顶级高手。"
else
seeequip=seeequip&"这种人太厉害了武功已经出神入化,最好不要惹。"
end if
rst.Close
end function
%>
