<%

Class AspSessionManager

	Private mFullPathToFile
	
	Private mIsDeserialized
	
	Private Sub Class_Initialize
	
		path = "D:\app\EREC\Archivo\Integracion\"
		
		mFullPathToFile  = path & Replace(Replace(Request.ServerVariables("REMOTE_ADDR"), ".", ""), ":", "") & ".txt"
		
		mIsDeserialized = False
		
	End Sub
	
	Function GetElementByKey(key)
	
		GetElementByKey = Session(key)
	
	End Function
	
	Sub SetElementByKey(key, value)
	
		Session(key) = value
	
	End Sub
	
	Public Property Get Elements()
	
		If mIsDeserialized Then
		
			Dim dictionary
			
			Set dictionary = Server.CreateObject("Scripting.Dictionary")
			
			Dim element
		
			For Each element in Session.Contents
			
				dictionary.Add element, Session(element)
				
			Next
			
			Set Elements = dictionary
			
		Else
		
			Set Elements = Nothing
			
		End If
	
	End Property
	
	Sub SerializeElements()
	
		Dim oFSO, oTextFile
		
		Set oFSO = CreateObject("Scripting.FileSystemObject")
		
		Set oTextFile = oFSO.CreateTextFile(mFullPathToFile, True)
		
		Dim element
		
		For Each element in Session.Contents
		
			oTextFile.WriteLine(element & "=" & Session(element))
			
		Next
		
		oTextFile.Close
		
		Set oFSO = Nothing
		
		Set oTextFile = Nothing
		
	End Sub
	
	
	Sub DeserializeElements()
	
		Dim oFSO, oTextFile
		
		Set oFSO = CreateObject("Scripting.FileSystemObject")
		
		Set oTextFile = oFSO.OpenTextFile(mFullPathToFile)
		
		Dim element
		
		While Not oTextFile.AtEndOfStream
		
			element = oTextFile.ReadLine
			
			Session(Split(element, "=")(0)) = Split(element, "=")(1)
		
		Wend
		
		oTextFile.Close
		
		Set oFSO = Nothing
		
		Set oTextFile = Nothing
		
		mIsDeserialized = True
		
	End Sub

End Class

%>