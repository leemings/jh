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
fl=5				'大的分类号
flname="其它"		'小分类的名字
ys=5				'页数
sssxm="406"			'商品id号
xx=0


for kk=1 to ys
TakenHTML=GetURL("http://game.qq.com/qqshow/index_m.shtml?/qqshow/shop_page/shop_"& sssxm &"_"& kk &"_0.htm")
'TakenHTML=GetURL("http://game.tencent.com/qqshow/index_m.shtml?shopid="& sssxm &"&nowpage="& kk &"&allpage="& ys &"&sex=0&pageno="& kk)
'TakenHTML=Bytes2BStr(TakenHTML)
'Response.Write TakenHTML
'Response.End
ll=lenb(TakenHTML)
TakenHTML=Bytes2BStr(TakenHTML)
'TakenHTML=replace(TakenHTML,chr(13),"")
'TakenHTML=replace(TakenHTML,chr(10),"")
'TakenHTML=replace(TakenHTML,chr(9),"")
abc=instr(TakenHTML,"id=qqshowGoods")
TakenHTML=mid(TakenHTML,abc,len(TakenHTML)-abc)
abc=instr(TakenHTML,"<form name="& chr(34) &"FORM")-10
TakenHTML=mid(TakenHTML,1,abc)
TakenHTML=replace(TakenHTML,"id=qqshowGoods width=""581"" border=""0"" cellspacing=""0"" cellpadding=""0"">","")
TakenHTML=replace(TakenHTML,"</tr>","")
TakenHTML=replace(TakenHTML,"<tr>","")
TakenHTML=replace(TakenHTML,"</td>","")
TakenHTML=replace(TakenHTML,"<td>","")
TakenHTML=replace(TakenHTML,"</table>","")
TakenHTML=replace(TakenHTML,"<br>","")
TakenHTML=replace(TakenHTML," ","")
TakenHTML=replace(TakenHTML,chr(13),"")
TakenHTML=replace(TakenHTML,chr(10),"")
TakenHTML=replace(TakenHTML,chr(8),"")
TakenHTML=replace(TakenHTML,"</script>","")
TakenHTML=replace(TakenHTML,"<script>","")
TakenHTML=replace(TakenHTML,"showTd(","")
TakenHTML=replace(TakenHTML,");",";")
TakenHTML=replace(TakenHTML,"<tablewidth=""582""border=""0""cellspacing=""0""cellpadding=""0""style=""BORDER-BOTTOM:#EBEBEB1pxdotted"">","")
TakenHTML=replace(TakenHTML,"'qqshow-item.tencent.com',","")
TakenHTML=replace(TakenHTML,"'","")
''香奈儿-星星闪烁-','5016','F','9_8','3.0','3.0','0','1.5','0','30','15','0';
'0 b c 商品名　1 i 商品号 2 d 性别　3 a　数据　4　价钱


'Response.Write TakenHTML 
'Response.End

Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("up.mdb")
rs.open "SELECT * FROM fl where flname='"& flname &"'",conn,1,3
if rs.Eof and rs.Bof then
	rs.addnew
	rs("fl")=fl
	rs("flname")=flname
	rs("fltype")=sssxm
	rs.update
end if
rs.close
show=split(TakenHTML,";")
for x=0 to UBound(show)-1
	data=split(show(x),",",-1)
	Response.Write data(0) & "," & data(1) & "," & data(2) & "," & data(3)  &"<br>"
	rs.open "SELECT * FROM data WHERE i='"& data(1) &"' and a='"& data(3) &"' and d='"& data(2) &"' and b='" & data(0) & "'",conn,1,3
		if rs.Eof and rs.Bof then
			xx=xx+1
			rs.addnew
			''香奈儿-星星闪烁-','5016','F','9_8','3.0','3.0','0','1.5','0','30','15','0';
			'0 b c 商品名　1 i 商品号 2 d 性别　3 a　数据　4　价钱
			'rs("a")=jm(data(2),((data(1) mod 15)+1))
			rs("a")=data(3)
			rs("b")=data(0)
			rs("c")=data(0)
			rs("d")=data(2)
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