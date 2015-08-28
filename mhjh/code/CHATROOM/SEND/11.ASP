<%
function bomb(un,st,co,gr,sn)
if (co="官府" and gr>=Application("yx8_mhjh_bombright")) then
aberrantlist=Application("yx8_mhjh_aberrantlist")
aberrantlistubd=ubound(aberrantlist)
dim newaberrantlist()
newaberrantname=";"
j=1
tt=now()
addsucf=false
for i=1 to aberrantlistubd step 4
if aberrantlist(i)=st and aberrantlist(i+2)="炸弹" then
redim preserve newaberrantlist(j),newaberrantlist(j+1),newaberrantlist(j+2),newaberrantlist(j+3)
newaberrantlist(j)=st
newaberrantlist(j+1)=un
newaberrantlist(j+2)="炸弹"
newaberrantlist(j+3)=dateadd("n",15,tt)
newaberrantname=newaberrantname&st&";"
j=j+4
addsucf=true
elseif datediff("s",aberrantlist(i+3),tt)<0 then
redim preserve newaberrantlist(j),newaberrantlist(j+1),newaberrantlist(j+2),newaberrantlist(j+3)
newaberrantlist(j)=aberrantlist(i)
newaberrantlist(j+1)=aberrantlist(i+1)
newaberrantlist(j+2)=aberrantlist(i+2)
newaberrantlist(j+3)=aberrantlist(i+3)
j=j+4
newaberrantname=newaberrantname&aberrantlist(i)&";"
end if
next
if addsucf=false then
redim preserve newaberrantlist(j),newaberrantlist(j+1),newaberrantlist(j+2),newaberrantlist(j+3)
newaberrantlist(j)=st
newaberrantlist(j+1)=un
newaberrantlist(j+2)="炸弹"
newaberrantlist(j+3)=dateadd("n",15,tt)
newaberrantname=newaberrantname&st&";"
end if
Application.Lock
Application("yx8_mhjh_aberrantlist")=newaberrantlist
Application("yx8_mhjh_aberrantname")=newaberrantname
Application.UnLock
erase newaberrantlist
erase aberrantlist
bomb="<font color=FF0000>【炸弹】</font>##将%%恭恭敬敬地炸上了天堂，并认认真真的考虑年年祭拜的可能性。"
else
bomb="<font color=FF0000>【炸弹】</font>##本是想向%%多扔几颗炸弹的，无耐官府中无人撑腰。"
end if
end function
%>
