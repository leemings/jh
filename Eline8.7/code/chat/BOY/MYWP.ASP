<%'快乐江湖物品操作函数♀wWw.happyjh.com♀
function mywpsl(wpdata,wpname)'取物品数量：数据
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

function add(wpdata,wpname,wpsl)'增加物品：数据、物品名、数量
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
	Response.Write "<script Language=Javascript>alert('提示：您的此类物品已经超过25种,\n请不要再操作！！');</script>"
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

function del(wpdata,wpname)'删除物品：数据,物品名
wpdata=trim(lcase(wpdata))
wpname=trim(lcase(wpname))
if isnull(wpdata) or wpdata="" then 
	 if IsObject(rs) then
		set rs=nothing
	 end if
	 if IsObject(conn) then
	 	set conn=nothing
	 end if
	Response.Write "<script Language=Javascript>alert('提示：物品库为空不能操作！！');</script>"
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

function abate(wpdata,wpname,wpsl)'减少物品：数据、物品名、数量
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
	Response.Write "<script Language=Javascript>alert('提示：物品库为空不能操作！！');</script>"
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
				Response.Write "<script Language=Javascript>alert('提示：物品["&wpname&"]只有("&sl&")个,\n您不能进行操作！！');</script>"
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
		Response.Write "<script Language=Javascript>alert('提示：物品["&wpname&"]并不存在,\n请您刷新IE浏览器！');</script>"
		response.end
end if
if isnull(temp) or temp="" then 
	abate=null
else
	abate=temp
end if
end function

function iswp(wpdata,wpname)'查看物品是否存在：数据、物品名，存在为1，不存在为0
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
