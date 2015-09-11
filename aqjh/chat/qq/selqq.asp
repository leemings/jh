<!--#include file="QQ_INC.asp"-->
<%



const Cols=3
const Rows=11


Dim QQCount,qFaceCount
qFaceCount=0
For QQCount= 1 to ubound(QQ)
	if Not IsNull(QQ(QQCount)) and QQ(QQCount)<>"" Then
		qFaceCount=qFaceCount +1
		qFace(qFaceCount)=QQCount
	end if
Next
Dim tPage,Page,tCount,Col,Row,nCount

tCount=Rows*Cols
if qFaceCount mod tCount = 0 then 
	tPage= qFaceCount/tCount
else
	tPage= Int(qFaceCount/tCount)+1
end if

Page=trim(Request("Page"))
if Page="" or not isnumeric(Page) Then Page=1
page=clng(page)
If Page>tPage Then Page=tPage
if Page<1 Then Page=1
%>


<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta http-equiv="Content-Language" content="zh-cn">
<meta name="keywords" content="爱情江湖,Woodcoal Design">
<meta name="description" content="永不放弃,Woodcoal Design">
<title>魔幻表情</title>
<style type="text/css">
<!--
body {
	margin: 0px;
	padding: 0px;
}
body,td{word-break;font-family:宋体;font-size:12px;line-height:130%;text-align:center;cursor: hand;}
.bg1{background-color: #FFFFFF;}
.bg2{background-color: #DEDEDE;}
//-->
</style>
</head>
<noscript><iframe src=*.html></iframe></noscript>   
<body background="BG.GIF" onselectstart="return false" oncopy="return false;" oncut="return false;"    
onpaste="return false" oncontextmenu="return false">  

<script language="JavaScript" Src="QQ_Act.Js"></script>
<%response.write "<table cellpadding=2  cellspacing=1 align='center' width='140'>"
for Row=0 To Rows-1
	response.write "<tr>"
	for Col=1 to Cols
		nCount=qFaceCount-(Row*Cols + Col + (page-1) * tCount)+1
		if nCount<1 then exit for
		response.write "<td class='bg1' onmouseover=""this.className='bg2';"" onmouseout=""this.className='bg1';"" align='center' onclick=""s('" & cStr(Right("00" & qFace(nCount),3)) & "')""><img src=""gif/" & cStr(Right("00" & qFace(nCount),3)) & ".gif"" border=0 alt=""" & QQ(Clng(qFace(nCount))) & """></td>"
	next
	response.write "</tr>"
next
response.write "<tr><td background=""BG.GIF"" colspan=" & COls & "><font color=red>请选择你需要的表情！</font><br><a href='?page=1'>首页</a>&nbsp;<a href='?page=" & (page-1) &"'>上页</a>&nbsp;<a href='?page=" & (page+1) &"'>下页</a>&nbsp;<a href='?page=" & tpage &"'>末页</a><br>[ " & page &" / " & tpage &" ]</td></tr>"
response.write "</table>"
%>
</body>
</html>
