<%
function numer(a,b)	'�ж��Ƿ�Ϊ����,aΪ���ݣ�bΪ����ʱ����ʾ����
	if not isnumeric(a) then
		Response.Write "<script Language=Javascript>alert('��ʾ��"&b&"');location.href = 'javascript:history.go(-1)';</script>"
		response.end
	end if
end function
function charjc(a,b)	'����Ƿ��зǷ�ֵ��aΪ����,bΪ����ʱ����ʾ����
	a=lcase(trim(a))
	if instr(a,"or")<>0 or instr(a,"'")<>0 or instr(a,chr(13))<>0 or instr(a,chr(34))<>0 or instr(a,chr(39))<>0 or instr(a,"|")<>0 or instr(a,"&")<>0 or instr(a,"java")<>0 or instr(a,"union")<>0 or instr(a,"(")<>0 or instr(a,")")<>0 or instr(a,"=")<>0 or instr(a,">")<>0 or instr(a,"<")<>0 then
		Response.Write "<script Language=Javascript>alert('��ʾ��"&b&"��');location.href = 'javascript:history.go(-1)';</script>"
		response.end
	end if
end function
sub subend(tssay)	'��ִֹ���ӳ���
	Response.Write "<script Language=Javascript>alert('"&tssay&"');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end sub
sub myclose(qingchu,sfgb)	'�ر����ݿⲢ��ֹ����ִ��,qingchuΪ��ֹ�����ʾ���sfgbΪ�Ƿ�رմ��ڣ�1Ϊ�ر�,0Ϊ������һҳ�档
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	if sfgb=1 then
		Response.Write "<script Language=Javascript>alert('��ʾ��"&qingchu&"');window.close();</script>"
	else
		Response.Write "<script Language=Javascript>alert('��ʾ��"&qingchu&"');location.href = 'javascript:history.go(-1)';</script>"
	end if
	response.end
end sub
function jc(a)	'�滻���Ƿ��ַ��������ַ�
	a=replace(a,"'","��")
	a=replace(a,chr(34),"��")
	a=replace(a,",","��")
	a=replace(a,chr(39),"��")
	a=replace(a,"&#","��#")
	a=replace(a,"=","��")
	a=replace(a,">","��")
	a=replace(a,"<","��")
	a=replace(a,":","��")
	a=replace(a,"+","��")
	a=replace(a,"-","��")
	a=replace(a,";","��")
	a=replace(a,"gt","��")
	a=replace(a,"lt","��")
	a=replace(a,"%","��")
	a=replace(a,"http","")
	a=replace(a,"w","")
	a=replace(a,"j","")
	a=replace(a,"h","")
	a=replace(a,"����","")
	a=replace(a,"c","")
	a=replace(a,"o","")
	a=replace(a,"m","")
	a=replace(a,"n","")
	a=replace(a,"e","")
	a=replace(a,"t","")
	a=replace(a,"r","")
	a=replace(a,"g","")
	a=replace(a,".","��")
	a=replace(a,"��","")
	a=replace(a,"��","")
	a=replace(a,"��","")
	a=replace(a,"��","")
	a=replace(a,"��","")
	a=replace(a,"��","")
	a=replace(a,"��","")
	a=replace(a,"��","")
	a=replace(a,"��","")
	a=replace(a,"��","")
	a=replace(a,"��","")
	a=replace(a,"��","")
	a=replace(a,"��","")
	a=replace(a,"��","")
	a=replace(a,"��","")
	a=replace(a,"��","")
	a=replace(a,"��","")
	a=replace(a,"��","")
	a=replace(a,"��","")
	a=replace(a,"��","")
	a=replace(a,"��","")
	a=replace(a,"q","")
	a=replace(a,"��","")
	a=replace(a,"��","")
	a=replace(a,"����","")
	a=Replace(a,"\x3c","")
	a=Replace(a,"\x3e","")
	a=Replace(a,"\074","")
	a=Replace(a,"\74","")
	a=Replace(a,"\75","")
	a=Replace(a,"\76","")
	badstr="�侫|��|ȥ��|��ʺ|����|��Ԫ|����|����|����|��|����|������|����|��|ɵB|����|����|���|����|����|����|����|����|����|����|����|����|����|غ��|��ȥ |���������|���������|���������|��Ƥ|��ͷ|��|�P|��|�H|����|��|��|����|����|����|����|����|��Ů|����|ʮ����|��ү|���|�Ҷ�|����|��|��|����|����"
	bad=split(badstr,"|")
	for i=0 to UBound(bad)
		a=Replace(a,bad(i),"**")
	next
	jc=a
end function
sub cls()	'�ر����ݿ�
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
end sub
%>
