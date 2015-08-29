<!--#include file="conn.asp"-->
<!--#include file="inc/const.asp"-->
<!--#include file="z_counter_conn.asp"-->
<%
Dim CurAspName,IP,ID,TempValue,Style,Pos1,Pos2
Dim Total,ThisYear,ThisMonth,ThisWeek,Today,HourArray(23),TopMonth,TopWeek,TopDay,TopHour,SinceFrom
Dim CalcMethod,CalcFlag

CalcMethod=0			' 0 - 所有页面分别统计， 1 - 所有页面合并统计， 2 - 只统计index.asp
Style=Request("style")
if isnull(Style) or Style="" then
	Style=1
elseif not IsNumeric(Style) then
	Style=1
elseif cint(Style)<0 then
	Style=1
elseif not hasASPPainter() and Cint(Style)=0 then
	Style=1
else
	Style=Cint(Style)
end if
CurAspName=trim(lcase(request.ServerVariables("HTTP_REFERER")))
if CurAspName<>"" then 
	Pos1=InstrRev(CurAspName,"/")
	if Pos1>0 then
		Pos2=Instr(Pos1+1,CurAspName,"?")
		if Pos2>Pos1 then
			CurAspName=left(request.ServerVariables("PATH_INFO"),len(request.ServerVariables("PATH_INFO"))-13)&Mid(CurAspName,Pos1+1,Pos2-Pos1-1)
		else
			CurAspName=left(request.ServerVariables("PATH_INFO"),len(request.ServerVariables("PATH_INFO"))-13)&right(CurAspName,len(CurAspName)-Pos1)
		end if
	else
		CurAspName=""
	end if
end if
CalcFlag=True
if CalcMethod=1 then
	if CurAspName<>"" then
		Pos1=InstrRev(CurAspName,"/")
		if Pos1>0 then
			CurAspName=Left(CurAspName,Pos1)&"index.asp"
		else
			CurAspName="index.asp"
		end if
	end if
elseif CalcMethod=2 then
	if CurAspName<>"" then
		Pos1=InstrRev(CurAspName,"/")
		if Pos1>0 then
			if CurAspName<>Left(CurAspName,Pos1)&"index.asp" then
				CalcFlag=False
			end if
			CurAspName=Left(CurAspName,Pos1)&"index.asp"
		else
			if CurAspName<>"index.asp" then
				CalcFlag=False
			end if
			CurAspName="index.asp"
		end if
	end if
end if
if CurAspName="" then response.end
SET rs = Server.CreateObject("ADODB.Recordset")
sql ="SELECT * From [Counter] where URL='"&CurAspName&"'"
RS.open sql,CntConn,1,3
IF RS.Bof and RS.Eof THEN
	RS.AddNew
	RS("URL")=CurAspName
	RS.Update
end if
TempValue=DateDiff("yyyy",RS("DateAndTime"),Date())
IF TempValue<>0 THEN
	IF TempValue=1 THEN RS("LastYear") = RS("ThisYear") ELSE RS("LastYear")=0
	IF RS("ThisYear")>RS("TopYear") THEN RS("TopYear")=RS("ThisYear")
	RS("ThisYear") = 0
END IF
TempValue=DateDiff("m",RS("DateAndTime"),Date())
IF TempValue<>0 THEN
	IF TempValue=1 THEN RS("LastMonth") = RS("ThisMonth") ELSE RS("LastMonth")=0
	IF RS("ThisMonth")>RS("TopMonth") THEN RS("TopMonth")=RS("ThisMonth")
	RS("ThisMonth") = 0
END IF
TempValue=DateDiff("d",RS("DateAndTime"),Date())
IF TempValue>=7 or Weekday(RS("DateAndTime"))>Weekday(Date()) THEN
	IF (TempValue<7 and Weekday(RS("DateAndTime"))>Weekday(Now())) or (TempValue>=7 and TempValue<14 and Weekday(RS("DateAndTime"))<=Weekday(Now())) THEN RS("LastWeek") = RS("ThisWeek") ELSE RS("LastWeek")=0
	IF RS("ThisWeek")>RS("TopWeek") THEN RS("TopWeek")=RS("ThisWeek")
	RS("ThisWeek") = 0
END IF
TempValue=DateDiff("d",RS("DateAndTime"),Date())
IF TempValue<>0 THEN
	IF TempValue=1 THEN RS("Yestoday") = RS("Today") ELSE RS("Yestoday")=0
	IF RS("Today")>RS("TopDay") THEN RS("TopDay")=RS("Today")
	RS("Today") = 0
	For i=0 to 23
		RS("Hour"&RIGHT("00"&i,2))=0
	Next
END IF
if CalcFlag then
	RS("Hour"&RIGHT("00"&Hour(Now()),2))=RS("Hour"&RIGHT("00"&Hour(Now()),2))+1
end if
For i=0 to 23
	IF RS("Hour"&RIGHT("00"&i,2))>RS("TopHour") THEN RS("TopHour")=RS("Hour"&RIGHT("00"&i,2))
	HourArray(i)=RS("Hour"&RIGHT("00"&i,2))
Next
RS("DateAndTime") = DATE()
if CalcFlag then
	RS("Total") =RS("Total")+1
	RS("ThisYear")=RS("ThisYear")+1
	RS("ThisMonth")=RS("ThisMonth")+1
	RS("ThisWeek")=RS("ThisWeek")+1
	RS("Today")=RS("Today")+1
end if
RS.Update
ID=RS("ID")
Total=RS("Total")
ThisYear=RS("ThisYear")
ThisMonth=RS("ThisMonth")
ThisWeek=RS("ThisWeek")
Today=RS("Today")
TopMonth=RS("TopMonth")
TopWeek=RS("TopWeek")
TopDay=RS("TopDay")
TopHour=RS("TopHour")
SinceFrom=RS("AddDate")
RS.Close
IP=Request.ServerVariables("REMOTE_HOST")
SET RS=Nothing
call GCounter(Total,ThisYear,ThisMonth,ThisWeek,Today,HourArray,TopMonth,TopWeek,TopDay,TopHour,SinceFrom,IP,Style)
call CloseCounterDatabase

Sub GCounter(Total,ThisYear,ThisMonth,ThisWeek,Today,HourArray,TopMonth,TopWeek,TopDay,TopHour,SinceFrom,IP,Style)
	Dim S, i, G,MaxHour
	S = Right("0000000000"&Total,10)
	if Style>0 then
		For i = 1 to Len(S)
			G = G&"<IMG SRC=images/img_counter/Style"&Style&"/"&Mid(S, i, 1)&".gif border=0>"
		Next
		response.write "document.write("""&G&""");"
	else
		dim pic
		Set pic = CreateObject("ASPPainter.Pictures.1")
		pic.SetFormat 3
		dim PanelColor
		if Instr(Forum_body(5),"#")>0 then
			PanelColor=Mid(Forum_body(5),Instr(Forum_body(5),"#"),7)
		else
			PanelColor="#CECECE"
		end if
		pic.SetBkColor Hex2Dec(mid(PanelColor,2,2)),Hex2Dec(mid(PanelColor,4,2)),Hex2Dec(mid(PanelColor,6,2)),255
		pic.Create 563,48
		pic.SetImageIndex 1
		pic.SetFormat 3
		pic.LoadFile Server.MapPath("images/img_counter/panel.gif")
		pic.Merge 0,1,0,0,0,0,pic.Width,pic.Height,100
		pic.Destroy
		pic.SetImageIndex 0
		pic.SetFontName "Times New Roman"
		pic.SetFontSize 7
		pic.SetColor 0,0,0,255
		pic.SetFontOrientation 0
		pic.TextOut 12,36,S
		pic.TextOut 119,31,right("000000000"&Today,9)
		pic.TextOut 230,33,right("0000000"&ThisMonth,7)
		pic.TextOut 289,33,right("0000000"&ThisYear,7)
		pic.TextOut 361,2,right("00000"&TopMonth,5)
		pic.TextOut 361,12,right("00000"&TopWeek,5)
		pic.TextOut 361,22,right("00000"&TopDay,5)
		pic.TextOut 361,32,right("00000"&TopHour,5)
		Dim MinValue,MidValue,MaxValue,TopValue,RotDeg
		MinValue=0
		if ThisMonth<10 then
			MidValue=5
			MaxValue=10
			TopValue=10
		else
			if cint(left(ThisMonth,1))<5 then
				MidValue=25
				MaxValue=50
				TopValue=CLng("5"&string(len(ThisMonth)-1,"0"))
			else
				MidValue=50
				MaxValue=100
				TopValue=CLng("1"&string(len(ThisMonth),"0"))
			end if
		end if
		RotDeg=ThisMonth*146/TopValue+17
		RotDeg=RotDeg*3.14159265/180
		pic.SetColor 255,255,255,255
		pic.SetFontSize 6
		pic.TextOut 229,24,MinValue
		if len(MidValue)=1 then
			pic.TextOut 247,4,MidValue
		else
			pic.TextOut 244,4,MidValue
		end if
		if len(MaxValue)=2 then
			pic.TextOut 259,24,MaxValue
		else
			pic.TextOut 254,24,MaxValue
		end if
		pic.SetColor 165,148,255,255
		pic.SetLineWidth 1
		pic.DrawRectangle 247,25,250,28
		pic.SetColor 0,0,0,255
		pic.SetPixel 247,25
		pic.SetPixel 247,28
		pic.SetPixel 250,25
		pic.SetPixel 250,28
		pic.SetColor 255,204,0,255
		pic.SetLineWidth 2
		pic.DrawLine 248,27,248-16*Cos(RotDeg),27-16*Sin(RotDeg)
		MinValue=0
		if ThisYear<10 then
			MidValue=5
			MaxValue=10
			TopValue=10
		else
			if cint(left(ThisYear,1))<5 then
				MidValue=25
				MaxValue=50
				TopValue=CLng("5"&string(len(ThisYear)-1,"0"))
			else
				MidValue=50
				MaxValue=100
				TopValue=CLng("1"&string(len(ThisYear),"0"))
			end if
		end if
		RotDeg=ThisYear*146/TopValue+17
		RotDeg=RotDeg*3.14159265/180
		pic.SetColor 255,255,255,255
		pic.TextOut 288,24,MinValue
		if len(MidValue)=1 then
			pic.TextOut 306,4,MidValue
		else
			pic.TextOut 303,4,MidValue
		end if
		if len(MaxValue)=2 then
			pic.TextOut 318,24,MaxValue
		else
			pic.TextOut 313,24,MaxValue
		end if
		pic.SetColor 165,148,255,255
		pic.SetLineWidth 1
		pic.DrawRectangle 306,25,309,28
		pic.SetColor 0,0,0,255
		pic.SetPixel 306,25
		pic.SetPixel 306,28
		pic.SetPixel 309,25
		pic.SetPixel 309,28
		pic.SetColor 255,204,0,255
		pic.SetLineWidth 2
		pic.DrawLine 307,27,307-16*Cos(RotDeg),27-16*Sin(RotDeg)
		MaxHour=0
		for i=0 to 23
			if HourArray(i)>MaxHour then MaxHour=HourArray(i)
		next
		MinValue=0
		if MaxHour<10 then
			MidValue=5
			MaxValue=10
		else
			if cint(left(MaxHour,1))<8 then
				MaxValue=CLng((cint(left(MaxHour,1))+(2-cint(left(MaxHour,1)) mod 2))&string(len(MaxHour)-1,"0"))
				MidValue=MaxValue/2
			else
				MaxValue=CLng("1"&string(len(MaxHour),"0"))
				MidValue=MaxValue/2
			end if
		end if
		pic.SetColor 0,0,0,255
		pic.SetFontSize 7
		pic.TextOut 193,24,MinValue
		pic.TextOut 193,12,MidValue
		pic.TextOut 193,0,MaxValue
		Dim LastX,LastY,CurX,CurY
		LastX=92
		LastY=32
		pic.SetLineWidth 1
		for i=0 to 23
			CurX=92+100*(i+1)/24
			CurY=32-28*HourArray(i)/MaxValue
			pic.DrawLine LastX,LastY,CurX,CurY
			LastX=CurX
			LastY=CurY
		next

		pic.SetFontSize 8
		pic.SetFontOrientation 270
		pic.TextOut 549,43,Right(year(SinceFrom),2)&right("00"&month(SinceFrom),2)&right("00"&day(SinceFrom),2)
		pic.SetFontOrientation 0
		pic.SetFontRectangle 420,11,527,22
		pic.setTextAlign 1
		pic.SetFontSize 7
		pic.TextOut 0,0,FormatDateTime(Now(),2)&" "&FormatDateTime(Now(),4)
		pic.ClearFontRectangle
		pic.SetFontRectangle 420,33,527,44
		pic.setTextAlign 1
		pic.SetFontSize 7
		pic.TextOut 0,0,IP
		pic.ClearFontRectangle
		response.contentType="image/GIF"
		response.BinaryWrite pic.SaveToStream
		pic.DestroyAll
		Set pic = Nothing
	end if
End Sub

Function hasASPPainter()
	Dim pic
	on error resume next
	Set pic = CreateObject("ASPPainter.Pictures.1")
	if err.Number=0 then
		on error goto 0
		hasASPPainter=true
	else
		err.clear
		on error goto 0
		hasASPPainter=false
	end if
End Function

function Hex2Dec(HexStr)
	dim Result,i,ch
	Result=0
	for i=1 to len(trim(HexStr))
		ch=mid(trim(HexStr),i,1)
		if ucase(ch)>="A" and ucase(ch)<="F" then
			Result=Result*16+ASC(Ucase(ch))-55
		elseif ucase(ch)>="0" and ucase(ch)<="9" then
			Result=Result*16+ASC(Ucase(ch))-48
		end if
	next
	Hex2Dec=Result
end function%>
