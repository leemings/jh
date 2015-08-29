<%
	dim Meipai_inout_rate,maxtitleid
	dim isZangmen		'是否掌门人
	isZangmen=false
	founderr=false
	set rs=conn.execute("select * from [GroupName] where Zangmen='"&membername&"'")
	if not(rs.eof and rs.bof) then isZangmen=true
	
	Meipai_inout_rate=split(conn.execute("select Meipai_inout_rate from [menpai]")(0),"|")
	maxtitleid=conn.execute("select top 1 UserTitleID from [UserTitle] order by UserTitleID desc")(0)
'加入 社团 扣除 金钱，增加 经验，魄力，威望 的百分比	1		Meipai_inout_rate(0～3) 		Meipai_inout_rate((menpai_setting(0)-1)*4+0～3))
'加入 魔教 增加 金钱，扣除 经验，魄力，威望 的百分比	2		Meipai_inout_rate(4～7)			Meipai_inout_rate((menpai_setting(0)-1)*4+0～3))

'退出 社团 增加 金钱，扣除 经验，魄力，威望 的百分比	1		Meipai_inout_rate(8～11)		Meipai_inout_rate(menpai_setting(0)*4+5～7)) 		
'退出 魔教 扣除 金钱，增加 经验，魄力，威望 的百分比	2		Meipai_inout_rate(12～15) 		Meipai_inout_rate(menpai_setting(0)*4+5～7))

'============================ 门派通用函数============================
'获取用户等级名          
function GetUserTitle(UserClass)
	GetUserTitle=conn.execute("select usertitle from usertitle where userclass="&UserClass&"")(0)
end function
'获取用户等级数字
function GetUserClass(UserTitle)
	GetUserClass=conn.execute("select UserClass from usertitle where UserTitle='"&UserTitle&"'")(0)
end function
'获取门派属性
function MenpaiAttri(expression)
	if expression=0 or expression="" then
		MenpaiAttri="明门正派"
	elseif expression=1 then
		MenpaiAttri="社团"
	elseif expression=2 then
		MenpaiAttri="魔教"
	elseif expression=3 then
		MenpaiAttri="<font color=navy>中立派</font>"
	else
		MenpaiAttri="明门正派"
	end if
end function
'获取门派状态
function MenpaiState(expression)
	if expression=2 then
		MenpaiState="<font color=red>已注销</font>"		
	elseif expression=3 then
		MenpaiState="<font color=red>等待审核</font>"
	else
		MenpaiState="正常"
	end if	
end function
'获取加入门派时各项属性值增减的百分率
function GetMeipai_in_rate(menpaiattri) 
	if cint(menpaiattri)=1 then
		GetMeipai_in_rate="" 
	elseif cint(menpaiattri)=2 then
		GetMeipai_in_rate=""	
	else
		GetMeipai_in_rate="0|0|0|0|0|0"	
	end if	
end function
'计算字符串长度
function strLength(str)
'功能：1个英文算1个单位长度，1个中文算2个单位长度  我来了注释
       ON ERROR RESUME NEXT
       dim WINNT_CHINESE
       WINNT_CHINESE    = (len("论坛")=2)
       if WINNT_CHINESE then
          dim l,t,c
          dim i
          l=len(str)
          t=l
          for i=1 to l
             c=asc(mid(str,i,1))
             if c<0 then c=c+65536
             if c>255 then
                t=t+1
             end if
          next
          strLength=t
       else 
          strLength=len(str)
       end if
       if err.number<>0 then err.clear
end function

'-------------------------------出错信息-------------------------------
sub menpai_err() 
%>
<table cellspacing=1 cellpadding=3 align=center class=tableborder1 style="width:75%">
	<tr><th height=25>门派错误信息</th></tr>
	<tr><td width="100%" class=tablebody1><b>产生错误的可能原因：</b><br><br><li>您是否仔细阅读了门派说明和门派规则，可能您还没有登录或者不具有使用当前功能的权限
	<font color=<%=Forum_body(8)%>><%=errmsg%></font><br></td></tr>
	<tr><td align=center height=26 class=tablebody2><a href="javascript:history.go(-1)">返回上一页</a></td></tr> 
</table>
<%
end sub

'-------------------------------成功信息-------------------------------
sub menpai_suc() 
%>
<table cellspacing=1 cellpadding=3 align=center class=tableborder1 style="width:75%">
	<tr><th height=25>门派操作成功</th></tr>
	<tr><td width="100%" class=tablebody1><b>操作成功</b><br><font color=navy><%=sucmsg%></font><br></td></tr>
	<tr><td align=center height=26 class=tablebody2><a href=<%=Request.ServerVariables("HTTP_REFERER")%>>返回上一页</a></td></tr>
</table>
<%
end sub
%>