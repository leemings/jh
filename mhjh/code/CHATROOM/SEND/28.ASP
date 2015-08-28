<%
function alert(un,st,co,gr,sn,mg)
if (co="官府" and gr>=Application("yx8_mhjh_kickguestright")) or (Application("yx8_mhjh_systemname"&sn)=co and gr>7) then
alert="<font color=FF0000 size=3>【警告】%%,网管慎重警告你:"&mg&"</font>"
else
alert="<font color=FF0000>【警告】</font>%%,##向你发无效警告:"&mg
end if
end function
%>
