<%@ LANGUAGE=VBScript codepage ="936" %>
<%Server.ScriptTimeout=999999
On Error resume Next 
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
Response.Expires=0
function jm(jmstr,jmsn)
jtemp=""
for t=1 to len(jmstr)
jtemp=jtemp &chr((asc(mid(jmstr,t,1)) xor jmsn))
next 
jm=jtemp
end function
Function GetURL(url) 
    Set Retrieval = CreateObject("Microsoft.XMLHTTP") 
    With Retrieval 
        .Open "GET", url, False
        .Send 
        GetURL = .responsebody 
	if len(.responsebody)<100 then
		response.write "获取远程文件 <a href="&url&" target=_blank>"&url&"</a> 失败。"
		response.end
	end if

    End With 
    Set Retrieval = Nothing 
End Function

Function bytes2BSTR(vIn) 
strReturn = "" 
For i = 1 To LenB(vIn) 
ThisCharCode = AscB(MidB(vIn,i,1)) 
If ThisCharCode < &H80 Then 
strReturn = strReturn & Chr(ThisCharCode) 
Else 
NextCharCode = AscB(MidB(vIn,i+1,1)) 
strReturn = strReturn & Chr (CLng(ThisCharCode) * &H100 + CInt(NextCharCode)) 
i = i + 1 
End If 
Next 
bytes2BSTR = strReturn 
End Function 
xx=0
ys=1			'页数
sssxm="501"		'商品id号
for kk=1 to ys
TakenHTML=GetURL("http://game.tencent.com/qqshow/index_m.shtml?shopid="& sssxm &"&nowpage="& kk &"&allpage="& ys &"&sex=0&pageno="& kk)
'TakenHTML=Bytes2BStr(TakenHTML)
'Response.Write TakenHTML
'Response.End
ll=lenb(TakenHTML)
TakenHTML=Bytes2BStr(TakenHTML)
'TakenHTML=replace(TakenHTML,chr(13),"")
'TakenHTML=replace(TakenHTML,chr(10),"")
'TakenHTML=replace(TakenHTML,chr(9),"")
abc=instr(TakenHTML,"<table id=qqshowGoods")
TakenHTML=mid(TakenHTML,abc,len(TakenHTML)-abc)
abc=instr(TakenHTML,"function GetMallPre()")-10
TakenHTML=mid(TakenHTML,1,abc)
temp=""
okstr=""
do while instr(TakenHTML,"PrintSex")<>0
abc=instr(TakenHTML,"<img src="&chr(34)&"http")
cde=instr(TakenHTML,"PrintSex")
temp=mid(TakenHTML,abc,(cde-abc-9))
TakenHTML=mid(TakenHTML,(cde+23),(len(TakenHTML)-cde))
abc=instr(temp,"itemgender")
cde=instr(temp,"border")
fgh=instr(temp,"<b>")
ll=len(temp)
okstr=okstr & mid(temp,abc,(cde-abc)) &"=" & right(temp,(ll-fgh-2))&"|"
loop
okstr=replace(okstr,"itemgender=","")
okstr=replace(okstr,"itemno","")
okstr=replace(okstr,"layerno","")
okstr=replace(okstr," ","")
okstr=replace(okstr,chr(34),"")
'Response.Write okstr&"<br><br><br>"
'Response.End
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("up.mdb")
show=split(okstr,"|")
for x=0 to UBound(show)-1
	data=split(show(x),"=",-1)
	Response.Write data(0) & "," & data(1) & "," & data(2) & "," & data(3)  &"<br>"
	rs.open "SELECT * FROM data WHERE i='"& data(1) &"' and a='"& data(2) &"' and d='"& data(0) &"' and b='" & data(3) & "'",conn,1,3
		if rs.Eof and rs.Bof then
			xx=xx+1
			rs.addnew
			'rs("a")=jm(data(2),((data(1) mod 15)+1))
			rs("a")=data(2)
			rs("b")=data(3)
			rs("c")=data(3)
			rs("d")=data(0)
			rs("f")=12
			rs("g")=8
			rs("h")=30
			rs("i")=data(1)
			rs("j")=sssxm
			rs("k")=now()
			rs("l")=0
			rs.update
		end if
		rs.close
next
set rs=nothing
conn.close
set conn=nothing
next
Response.Write "<br><br>总共更新："& xx &"个记录!<br>请去修改分类!"
%>