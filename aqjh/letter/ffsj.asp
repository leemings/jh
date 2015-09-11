<%
function numer(a,b)	'判断是否为数字,a为数据，b为错误时的提示话语
	if not isnumeric(a) then
		Response.Write "<script Language=Javascript>alert('提示："&b&"');location.href = 'javascript:history.go(-1)';</script>"
		response.end
	end if
end function
function charjc(a,b)	'检测是否有非法值，a为数据,b为错误时的提示话语
	a=lcase(trim(a))
	if instr(a,"or")<>0 or instr(a,"'")<>0 or instr(a,chr(13))<>0 or instr(a,chr(34))<>0 or instr(a,chr(39))<>0 or instr(a,"|")<>0 or instr(a,"&")<>0 or instr(a,"java")<>0 or instr(a,"union")<>0 or instr(a,"(")<>0 or instr(a,")")<>0 or instr(a,"=")<>0 or instr(a,">")<>0 or instr(a,"<")<>0 then
		Response.Write "<script Language=Javascript>alert('提示："&b&"！');location.href = 'javascript:history.go(-1)';</script>"
		response.end
	end if
end function
sub subend(tssay)	'中止执行子程序
	Response.Write "<script Language=Javascript>alert('"&tssay&"');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end sub
sub myclose(qingchu,sfgb)	'关闭数据库并中止程序执行,qingchu为中止后的提示话语，sfgb为是否关闭窗口，1为关闭,0为返回上一页面。
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	if sfgb=1 then
		Response.Write "<script Language=Javascript>alert('提示："&qingchu&"');window.close();</script>"
	else
		Response.Write "<script Language=Javascript>alert('提示："&qingchu&"');location.href = 'javascript:history.go(-1)';</script>"
	end if
	response.end
end sub
function jc(a)	'替换掉非法字符及拉人字符
	a=replace(a,"'","“")
	a=replace(a,chr(34),"，")
	a=replace(a,",","，")
	a=replace(a,chr(39),"‘")
	a=replace(a,"&#","—#")
	a=replace(a,"=","＝")
	a=replace(a,">","〉")
	a=replace(a,"<","〈")
	a=replace(a,":","：")
	a=replace(a,"+","＋")
	a=replace(a,"-","－")
	a=replace(a,";","；")
	a=replace(a,"gt","》")
	a=replace(a,"lt","《")
	a=replace(a,"%","％")
	a=replace(a,"http","")
	a=replace(a,"w","")
	a=replace(a,"j","")
	a=replace(a,"h","")
	a=replace(a,"江湖","")
	a=replace(a,"c","")
	a=replace(a,"o","")
	a=replace(a,"m","")
	a=replace(a,"n","")
	a=replace(a,"e","")
	a=replace(a,"t","")
	a=replace(a,"r","")
	a=replace(a,"g","")
	a=replace(a,".","、")
	a=replace(a,"ｏ","")
	a=replace(a,"ｒ","")
	a=replace(a,"ｇ","")
	a=replace(a,"ｃ","")
	a=replace(a,"ｍ","")
	a=replace(a,"ｎ","")
	a=replace(a,"ｅ","")
	a=replace(a,"ｔ","")
	a=replace(a,"ｗ","")
	a=replace(a,"ｊ","")
	a=replace(a,"ｈ","")
	a=replace(a,"Ｃ","")
	a=replace(a,"Ｏ","")
	a=replace(a,"Ｍ","")
	a=replace(a,"Ｒ","")
	a=replace(a,"Ｇ","")
	a=replace(a,"Ｎ","")
	a=replace(a,"Ｅ","")
	a=replace(a,"Ｔ","")
	a=replace(a,"Ｊ","")
	a=replace(a,"Ｈ","")
	a=replace(a,"q","")
	a=replace(a,"ｑ","")
	a=replace(a,"Ｑ","")
	a=replace(a,"加我","")
	a=Replace(a,"\x3c","")
	a=Replace(a,"\x3e","")
	a=Replace(a,"\074","")
	a=Replace(a,"\74","")
	a=Replace(a,"\75","")
	a=Replace(a,"\76","")
	badstr="射精|奸|去死|吃屎|吴鑫|振元|你妈|你娘|日你|尻|操你|干死你|王八|逼|傻B|贱人|狗娘|婊子|表子|靠你|叉你|叉死|插你|插死|干你|干死|日死|鸡巴|睾丸|死去 |爬你达来蛋|撅你达来蛋|死你达来蛋|包皮|龟头|屄|赑|妣|肏|奶子|尻|屌|作爱|做爱|床上|抱抱|鸡八|处女|打炮|十八摸|你爷|你爸|我儿|操你|妈|逼|网管|掌门"
	bad=split(badstr,"|")
	for i=0 to UBound(bad)
		a=Replace(a,bad(i),"**")
	next
	jc=a
end function
sub cls()	'关闭数据库
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
end sub
%>
