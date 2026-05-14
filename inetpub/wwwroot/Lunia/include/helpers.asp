<%
	dim filePrefix,ranksPerPage,MythStageGroups
	filePrefix = "update_"
	ranksPerPage = 20
	MythStageGroups = Array(196618,474497,543980,4165390,10116849,12579698,19141002,27110626,35290273,60356942,31516281,47931489,42127076,42127075,42067494,42067493,42037702,42007911,41189791,41184987,41184986,41184025,41183064,31325819,27709335,26727571,25011321,20967636,20966676,20966675,20965714,20964753,20494160,19296572,19295613,19295612,19295611,16178618,12556796,9734933,4998893,4165390,1238614)
	
	function roundUp(Number)
		roundUp = Int(Number)
		if roundUp <> Number then
			roundUp = roundUp + 1
		end if
	end function

	function getFileCalled()
		aux = Split(Request.ServerVariables("URL"),"/")
		getFileCalled = aux(UBound(aux))
	end function
	
	function getCacheFile()
		getCacheFile = Replace(Server.MapPath("./") & "\cache\"&Replace(getFileCalled(),".asp","")&".xml",filePrefix,"")
	end function
	
	Function InArray(theArray,theValue)
		dim i, fnd
		fnd = 0 ' 0 instead of False
		For each item in theArray
			'response.write( item & "=" & theValue & "<br>")
			If item = theValue Then
				fnd = 1 ' 1 instead of True
				Exit For
			End If
		Next
		InArray = fnd
	End Function
	
	function countNodes(xmlObj)
		dim i : i = 0
		for each item in xmlObj
			i = i + 1
		next
		countNodes = i
	end function
%>