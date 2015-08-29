<%
' Company: Sabra Inc
' Author: Dave Hoffenberg 
' Function: Finds a value within a delimited list
' Freeware

Function ListFind(value,list,delim)

If list <> "" Then

arr = split(list,delim)

For i=0 to ubound(arr)

If arr(i) = value Then
Match = i
Exit For
Else
Match = 0 
End If

Next

ListFind = Match

Else

ListFind = 0

End If

End Function

strList = "t,b,q"
response.write ListFind("q","a,b,c,q",",")
'$remove(abcdefg, cd)
response.write replace(strlist,",b","")

%>