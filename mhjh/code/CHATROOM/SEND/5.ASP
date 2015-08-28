<%
function huanying(un,st)
rst.Open "SELECT 门派 FROM 用户 WHERE 姓名='" & un & "' and (身份='掌门' or 身份='副掌门')",conn
if rst.EOF or rst.BOF then
huanying="<font color=FF0000>【邀请入派】</font>官府人员对##说：不是掌门或者副掌门,无权拉人入派！"
else
men=rst("门派")
huanying="<font color=FF0000>【邀请入派】</font><font color=blue> ##对%%说:欢迎你入住魔幻江湖，如果你还没有加入任何门派，那么就请加入我们<input type=button value='"&men&"' onclick=""javascript:window.open('joincorp1.asp?mg="&men&"','newwindow','width=380,height=210')"" style=""background-color:#FF0000;color:FFFFFF;border: 1 double"" onMouseOver=""this.style.color='FFFF00'"" onMouseOut=""this.style.color='FFFFFF'"" name=""button2231"">吧，这里的工资更高，而且同门兄弟都会帮助你的！希望这里能成为你永远的家！如果你同意，就点这个红色的我派的名字吧!</font>"
end if
rst.close
end function
%>