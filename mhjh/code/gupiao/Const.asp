<!--#include file="Function.asp"-->
<%
	Dim ChengJiaoNum,ChengJiaoMoney,Gupiao_Setting,Trade_Setting,StopReadme,AI_Setting,Custom_Setting		'���ճɽ�����,���ճɽ����,��Ʊ����
	Dim membername,master,userhidden
	Dim MyUserID,MyCash,MyBiShu,MyStats		'�û�ID,���õ��ʽ�,������������,״̬  
	Dim HaveAccount '�Ƿ��й�Ʊ�˺� 
	Dim errmess,sucmess,rUrl,founderr,i,stats
		
	
	membername=session("yx8_mhjh_username")
	userhidden=""
	userhidden=iif(userhidden="",2,userhidden)
	master=false 
	
	if membername="" then
		response.write "<title>��Ʊ������-��ʾ��Ϣ</title><style type=text/css>td {  font-family: ����; font-size: 9pt}</style>"
		errmess="<li>����û�е�½�����½��������"
		call endinfo(0)
		response.end
	end if
	
'	set rsbbs=connbbs.execute("select id from [admin] where adduser='"&membername&"'")
'	if rsbbs.eof and rsbbs.bof then
'		master=false
'	else
'		master=true
'	end if
'	rsbbs.close
        if Session("yx8_mhjh_grade")=10 and instr(Application("yx8_mhjh_admin"),Session("yx8_mhjh_username"))<>0 then
           master=true
        else
	   master=false        
        end if 
			
	REM ��ȡ������Ϣ 
	set rs=conn.execute("select top 1 * from [GupiaoConfig] order by id") 
	'if datediff("d",rs("TodayDate"),Now())<>0 then
	'������͹�Ʊ����
	if datediff("h",rs("TodayDate"),Now())<>0 then
		conn.execute("update [�ͻ�] set ��������=0,��������=0")
		conn.execute("update [GupiaoConfig] set TodayBuy=0,TodaySale=0,TodayTotal=0,TodayDate=now()")  '
		call History()
		ChengJiaoNum=0
		ChengJiaoMoney=0
	else
		ChengJiaoNum=rs("TodayBuy")+rs("TodaySale")
		ChengJiaoMoney=rs("TodayTotal")
	end if
	StopReadme=rs("StopReadme")
	Gupiao_Setting=split(rs("Gupiao_Setting"),",")
	Trade_Setting =split(rs("Trade_Setting"),"|")
	AI_Setting    =split(rs("AI_Setting"),"|")
	Custom_Setting=split(rs("Custom_Setting"),"||")
	rs.close
	
	REM �Ƿ���ʱ���չ���/�Ƿ���ʱ�رչ���
	if not master then
		if cint(Gupiao_Setting(0))=0 then
			call CloseGuPiao(0)
			response.end	
		elseif cint(Gupiao_Setting(3))=1 then  				'��Ʊ �Ƿ� ���� 
			if hour(time)<cint(split(Gupiao_Setting(4),"||")(0)) or hour(time)>=cint(split(Gupiao_Setting(4),"||")(1)) then
				call CloseGuPiao(1)
				response.end
			end if		
		end if
	end if
	 
	REM ��ˢ�¹��� 
	if cint(Gupiao_Setting(6))=1 and (not master) then
		Dim SplitReflashPage
		Dim DoReflashPage
		Dim ScriptName
		DoReflashPage=false
		ScriptName=lcase(request.ServerVariables("PATH_INFO"))
		if trim(Gupiao_Setting(8))<>"" then
			SplitReflashPage=split(Gupiao_Setting(8),"|")
			for i=0 to ubound(SplitReflashPage)
				if instr(scriptname,SplitReflashPage(i))>0 then
					DoReflashPage=true
					exit for
				end if
			next
		end if
		if (not isnull(session("GP_ReflashTime"))) and cint(Gupiao_Setting(7))>0 and DoReflashPage then
			if DateDiff("s",session("GP_ReflashTime"),Now())<cint(Gupiao_Setting(7)) then
				'response.write "<META http-equiv=Content-Type content=text/html; charset=gb2312><meta HTTP-EQUIV=REFRESH CONTENT="&Gupiao_Setting(7)&">��ҳ�������˷�ˢ�»��ƣ��벻Ҫ��<font color=red>"&Gupiao_Setting(7)&"</font>��������ˢ�±�ҳ��<BR><font color=blue>"&Gupiao_Setting(7)&"</font>��֮�󽫻��Զ���ҳ�棬���Ժ󡭡�"
				call CloseGuPiao(2)
				response.end
			else
				session("GP_ReflashTime")=Now()
			end if
		elseif isnull(session("GP_ReflashTime")) and cint(Gupiao_Setting(7))>0 and DoReflashPage then
			Session("GP_ReflashTime")=Now()
		end if
	end if
			
	REM ��ȡ�û���Ϣ
	set rs=conn.execute("select ����,�ʽ�,��������,��������,id,���ʽ� from [�ͻ�] where �ʺ�='"&membername&"'"  )
	if rs.eof and rs.bof then 
		HaveAccount=false
		MyCash=0
		MyBiShu=0
		MyUserID=0
	elseif rs(0)=1 then 	
		rs.close
		response.write "<title>"&Gupiao_Setting(5)&"-�˺ű�����</title><style type=text/css>td {  font-family: ����; font-size: 9pt}</style>"
		errmess="<li>���Ĺ�Ʊ�˻������ᣬ�������Ա��ϵ��"
		call endinfo(0)
		response.end
	else
		HaveAccount=True	'�Ƿ��Ѿ����� 
		MyCash=rs(1)		'���õ��ʽ� 
		MyBiShu=rs(2)+rs(3) '������������ 
		MyUserID=rs(4)		'UserID 
	end if
	rs.close	
%>