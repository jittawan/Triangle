<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Jittawan Triangle</title>
</head>

<body>
<%
Dim myobj,f,fn
Set myobj = CreateObject("Scripting.FileSystemObject")
Set objTextFile = myobj.OpenTextFile("D:\triangle1.txt")

row = split(objTextFile.ReadAll,vbCrLf)
For i = UBound(row) To 0 Step -1 
	row(i) = Split(row(i)," ")
	If i < UBound(row) Then
		For j = 0 To UBound(row(i))
			If (row(i)(j) + row(i+1)(j)) > (row(i)(j) + row(i+1)(j+1)) Then
				row(i)(j) = CInt(row(i)(j)) + CInt(row(i+1)(j))
			Else
				row(i)(j) = CInt(row(i)(j)) + CInt(row(i+1)(j+1))
			End If
		Next
	End If
Next

Response.Write "Answer is  "&row(0)(0)

objTextFile.close
set myobj  = nothing

%>
</body>
</html>
