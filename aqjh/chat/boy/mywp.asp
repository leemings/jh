<%'
function mywpsl(wpdata,wpname)'ȡ��Ʒ����������
wpdata=trim(lcase(wpdata))
if isnull(wpdata) or wpdata="" then 
	mywpsl=0
	exit function
end if
data1=split(wpdata,";")
data2=UBound(data1)
mywpsl=0
for i=0 to data2-1
	data3=split(data1(i),"|")
	sl=clng(data3(1))
		if data3(0)=wpname then
			mywpsl=sl
		end if
erase data3		
next 
erase data1
end function

function add(wpdata,wpname,wpsl)'������Ʒ�����ݡ���Ʒ��������
wpdata=trim(lcase(wpdata))
wpname=trim(lcase(wpname))
wpsl=int(abs(wpsl))
if isnull(wpdata) or wpdata="" then 
	add=wpname&"|"&wpsl&";"
	exit function
end if
data1=split(wpdata,";")
data2=UBound(data1)
if data2>25 then
	 if IsObject(rs) then
		set rs=nothing
	 end if
	 if IsObject(conn) then
	 	set conn=nothing
	 end if
	Response.Write "<script Language=Javascript>alert('��ʾ�����Ĵ�����Ʒ�Ѿ�����25��,\n�벻Ҫ�ٲ�������');</script>"
	response.end
end if
temp=null
tempcz=0
for i=0 to data2-1
	data3=split(data1(i),"|")
	sl=clng(data3(1))
		if data3(0)=wpname then
			tempcz=1
			temp=temp&wpname&"|"&sl+wpsl&";"
		else
			temp=temp&data3(0)&"|"&sl&";"
		end if
erase data3		
next 
erase data1
if tempcz=0 then
	temp=temp&wpname&"|"&wpsl&";"
end if
if isnull(temp) or temp="" then 
	add=null
else
	add=temp
end if
end function

function del(wpdata,wpname)'ɾ����Ʒ������,��Ʒ��
wpdata=trim(lcase(wpdata))
wpname=trim(lcase(wpname))
if isnull(wpdata) or wpdata="" then 
	 if IsObject(rs) then
		set rs=nothing
	 end if
	 if IsObject(conn) then
	 	set conn=nothing
	 end if
	Response.Write "<script Language=Javascript>alert('��ʾ����Ʒ��Ϊ�ղ��ܲ�������');</script>"
	response.end
'	del=null
'	exit function
end if
data1=split(wpdata,";")
data2=UBound(data1)
temp=null
for i=0 to data2-1
	data3=split(data1(i),"|")
	sl=clng(data3(1))
		if data3(0)<>wpname then
			temp=temp&data3(0)&"|"&sl&";"
		end if
erase data3		
next 
erase data1
if isnull(temp) or temp="" then 
	del=null
else
	del=temp
end if
end function

function abate(wpdata,wpname,wpsl)'������Ʒ�����ݡ���Ʒ��������
wpdata=trim(lcase(wpdata))
wpname=trim(lcase(wpname))
wpsl=int(abs(wpsl))
if isnull(wpdata) or wpdata="" then 
	 if IsObject(rs) then
		set rs=nothing
	 end if
	 if IsObject(conn) then
	 	set conn=nothing
	 end if
	Response.Write "<script Language=Javascript>alert('��ʾ����Ʒ��Ϊ�ղ��ܲ�������');</script>"
	response.end
end if
data1=split(wpdata,";")
data2=UBound(data1)
temp=null
tempcz=0
for i=0 to data2-1
	data3=split(data1(i),"|")
	sl=clng(data3(1))
		if data3(0)=wpname then
			if wpsl>sl then
				 if IsObject(rs) then
					set rs=nothing
				 end if
				 if IsObject(conn) then
				 	set conn=nothing
				 end if
				Response.Write "<script Language=Javascript>alert('��ʾ����Ʒ["&wpname&"]ֻ��("&sl&")��,\n�����ܽ��в�������');</script>"
				response.end
			end if
			tempcz=1
			if sl-wpsl>0 then
				temp=temp&wpname&"|"&sl-wpsl&";"
			end if
		else
			temp=temp&data3(0)&"|"&sl&";"
		end if
erase data3
next
erase data1
if tempcz=0 then
		 if IsObject(rs) then
			set rs=nothing
		 end if
		 if IsObject(conn) then
		 	set conn=nothing
		 end if
		Response.Write "<script Language=Javascript>alert('��ʾ����Ʒ["&wpname&"]��������,\n����ˢ��IE�������');</script>"
		response.end
end if
if isnull(temp) or temp="" then 
	abate=null
else
	abate=temp
end if
end function

function iswp(wpdata,wpname)'�鿴��Ʒ�Ƿ���ڣ����ݡ���Ʒ��������Ϊ1��������Ϊ0
wpdata=trim(lcase(wpdata))
wpname=trim(lcase(wpname))
if isnull(wpdata) or wpdata="" then 
	iswp=0
	exit function
end if
data1=split(wpdata,";")
data2=UBound(data1)
temp=null
iswp=0
for i=0 to data2-1
	data3=split(data1(i),"|")
	sl=clng(data3(1))
		if data3(0)=wpname and sl>0 then
			iswp=1
		end if
erase data3
next
erase data1
end function
%>