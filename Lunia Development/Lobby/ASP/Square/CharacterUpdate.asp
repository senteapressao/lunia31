<!--#include file="../common.asp"-->
<!-- METADATA TYPE="typelib" FILE="C:\Program Files\Common Files\System\ado\msado15.dll" -->
<%
	Call Init()
	' request  : (POST)
	'			characterName,classNumber,stageLevel,stageExp,pvpLevel,pvpExp,warLevel,warExp,gameMoney,bankMoney,skillPoint,instantStateFlags,addedSkillPoint,rebirthCount,storedLevel,storedSkillPoint
	'			, delete Licenses(subseparated: stageGroupHash)
	'			, update Licenses(subseparated: stageGroupHash,accessLevel)
	'			, delete Items(subseparated: bagNumber,positionNumber)
	'			, update Items(subseparated: itemHash,bagNumber,positionNumber,stackedCount,itemSerial,instance,itemExpire)
	'			, delete Skills(subseparated: skillGroupHash)
	'			, update Skills(subseparated: skillGroupHash,skillLevel)
	'			, delete Quickslots(subseparated: position)
	'			, update Quickslots(subseparated: hash,position,isSkill,instance,itemExpire)
	'			, delete WorkingQuests(subseparated: questHash)
	'			, update WorkingQuests(subseparated: questHash,currentState,param1,param2,param3,expiredDate)
	'			, update CompletedQuests(subseparated: questHash,stageGroupHash,stageLevel,nextQuestStageGroupHash,nextQuestStageLevel,completeCount)
	'			, change onwer Pets(subseparated: petSerial)
	'			, delete Pets(subseparated: petSerial)
	'			, update Pets(subseparated: petSerial,petName,petLevel,petExp,fullValue,isRare,enchantSerial)	
	'			, delete PetItems(subseparated: petSerial,bagNumber,positionNumber)
	'			, update PetItems(subseparated: petSerial,bagNumber,positionNumber,itemHash,stackedCount,itemSerial,instance,itemExpire)
	'			, delete PetTraining(subseparated: petSerial)
	'			, insert PetTraining(subseparated: petSerial,itemHash,stackedCount,instance,itemExpire,expFactor,startTime,endTime)
	'			, callBack ( 0:none, 1:terminate, 2:rebirth )
	' response : characterName,callBack ( 0:none, 1:terminate, 2:rebirth )
	
	Dim i,j,k
	Dim sql,conn,rs
	Dim retString : retString = ""
	
	Dim rowData
	'rowData="Teles1399279500751451100000308900000545914418000153048101034505171183991861025Zack41089010.018661216239.843750334954.34375024505181627602258736Oracle10000.00000051443.101563151927.20312504505181627602375817Oracle123720000.08782744231.351563134201.98437504505130885171374770das10010.000000583438.312500747356.68750004505178999442517637lime156000.00000082003.765625172519.90625001"
	'REQUESTERID="SquareDEV_TELES"
	rowData=getConvAscii(Request.BinaryRead(Request.TotalBytes))
	Parameters=Split(rowData,SEPARATOR)
	
	' preparing request parameters
	Dim characterName,classNumber,stageLevel,stageExp,pvpLevel,pvpExp,warLevel,warExp,gameMoney,bankMoney,skillPoint,instantStateFlags _
		,addedSkillPoint,rebirthCount,storedLevel,storedSkillPoint
	Dim deleteLicenses,updateLicenses _
		,deleteItems,updateItems _
		,deleteSkills,updateSkills _
		,deleteQuickslots,updateQuickslots _
		,deleteWorkingQuests,updateWorkingQuests _
		,updateCompletedQuests _
		,updatePetOwner,updatePets,deletePets _
		,deletePetItems,updatePetItems _
		,deletePetTraining,insertPetTraining _
		,updatePetEnchantSerial
	Dim callBack
	Dim arr

	If UBound(Parameters)<34 Then Call Error("not enough parameter : ("& UBound(Parameters) &");"& rowData)
	
	characterName			= Parameters(0)
	classNumber				= Parameters(1)
	stageLevel				= Parameters(2)
	stageExp				= Parameters(3)
	pvpLevel				= Parameters(4)
	pvpExp					= Parameters(5)
	warLevel				= Parameters(6)
	warExp					= Parameters(7)
	gameMoney				= Parameters(8)
	bankMoney				= Parameters(9)
	skillPoint				= Parameters(10)
	instantStateFlags		= Parameters(11)
	
	addedSkillPoint			= Parameters(12)
	rebirthCount			= Parameters(13)
	storedLevel				= Parameters(14)
	storedSkillPoint		= Parameters(15)
	
	deleteLicenses			= Parameters(16)
	updateLicenses			= Parameters(17)
	
	deleteItems				= Parameters(18)
	updateItems				= Parameters(19)
	
	deleteSkills			= Parameters(20)
	updateSkills			= Parameters(21)
	
	deleteQuickslots		= Parameters(22)
	updateQuickslots		= Parameters(23)
	
	deleteWorkingQuests		= Parameters(24)
	updateWorkingQuests		= Parameters(25)
	
	updateCompletedQuests	= Parameters(26)
	
	updatePetOwner			= Parameters(27)
	
	deletePets				= Parameters(28)
	updatePets				= Parameters(29)
	
	deletePetItems			= Parameters(30)
	updatePetItems			= Parameters(31)
	
	deletePetTraining		= Parameters(32)
	insertPetTraining		= Parameters(33)
	
	callBack				= Parameters(34)

	Dim setCount : setCount = 0
 
	' basic info
	sql = sql &"exec UpdateCharacter"&_
		" "& SqlQuot(characterName) &_
		","& stageLevel &_
		","& stageExp &_
		","& pvpLevel &_
		","& pvpExp &_
		","& warLevel &_
		","& warExp &_
		","& gameMoney &_
		","& bankMoney &_
		","& skillPoint &_
		","& instantStateFlags &_
		","& addedSkillPoint &_
		","& rebirthCount &_
		","& storedLevel &_
		","& storedSkillPoint &_
		",@nError output, @strSP output"&_
		vbCrLf
	setCount = setCount +1
	sql = sql & "if(@nError not in (0)) goto lblFail"& vbCrLf
	
	' update Licenses
	arr = Split(updateLicenses,SUBSEPARATOR)
	If Ubound(arr)>=0 Then
		If ((UBound(arr)+1) Mod 2)<>0 Then
			Call Error("invalid number of item parameter")
		End If
		
		For i=0 To UBound(arr) Step 2
			sql = sql &"exec UpdateLicense"&_
				" "& SqlQuot(characterName) &_
				","& arr(i) &_
				","& arr(i+1) &_
				",@nError output, @strSP output"&_
				vbCrLf
			setCount = setCount +1
			sql = sql & "if(@nError not in (0)) goto lblFail"& vbCrLf
		Next
	End If

	' delete Items
	arr = Split(deleteItems,SUBSEPARATOR)
	If Ubound(arr)>=0 Then
		If ((UBound(arr)+1) Mod 2)<>0 Then
			Call Error("invalid number of item parameter")
		End If

		For i=0 To UBound(arr) Step 2
			sql = sql &"exec DeleteItem"&_
				" "& SqlQuot(characterName) &_
				","& arr(i) &_
				","& arr(i+1) &_
				",@nError output, @strSP output"&_
				vbCrLf
			setCount = setCount +1
			sql = sql & "if(@nError not in (0)) goto lblFail"& vbCrLf
		Next
	End If

	' update Items
	arr = Split(updateItems,SUBSEPARATOR)
	If Ubound(arr)>=0 Then
		If ((UBound(arr)+1) Mod 7)<>0 Then
			Call Error("invalid number of item parameter")
		End If

		For i=0 To UBound(arr) Step 7
			sql = sql &"exec UpdateItem"&_
				" "& SqlQuot(characterName) &_
				","& arr(i+1) &_
				","& arr(i+2) &_
				","& arr(i) &_
				","& arr(i+3)
			If IsNumeric(arr(i+4)) Then
				sql = sql &","& arr(i+4)
			Else
				sql = sql &",null"
			End If				
			sql = sql &","& arr(i+5) &_
				","& SqlQuot(arr(i+6)) &_
				",@nError output, @strSP output"&_
				vbCrLf
			setCount = setCount +1
			sql = sql & "if(@nError not in (0)) goto lblFail"& vbCrLf
		Next
	End If

	' delete Skills
	arr = Split(deleteSkills,SUBSEPARATOR)
	If Ubound(arr)>=0 Then
		For i=0 To UBound(arr) Step 1
			sql = sql &"exec DeleteSkill"&_
				" "& SqlQuot(characterName) &_
				","& arr(i) &_
				",@nError output, @strSP output"&_
				vbCrLf
			setCount = setCount +1
			sql = sql & "if(@nError not in (0)) goto lblFail"& vbCrLf	
		Next
	End If

	' update Skills
	arr = Split(updateSkills,SUBSEPARATOR)
	If Ubound(arr)>=0 Then
		If ((UBound(arr)+1) Mod 2)<>0 Then
			Call Error("invalid number of item parameter")
		End If
		
		For i=0 To UBound(arr) Step 2
			sql = sql &"exec UpdateSkill"&_
				" "& SqlQuot(characterName) &_
				","& arr(i) &_
				","& arr(i+1) &_
				",@nError output, @strSP output"&_
				vbCrLf
			setCount = setCount +1
			sql = sql & "if(@nError not in (0)) goto lblFail"& vbCrLf
		Next
	End If

	' delete WorkingQuests
	arr = Split(deleteWorkingQuests,SUBSEPARATOR)
	If Ubound(arr)>=0 Then
		For i=0 To UBound(arr) Step 1
			sql = sql &"exec DeleteWorkingQuest"&_
				" "& SqlQuot(characterName) &_
				","& arr(i) &_
				",@nError output, @strSP output"&_
				vbCrLf
			setCount = setCount +1
			sql = sql & "if(@nError not in (0)) goto lblFail"& vbCrLf
		Next
	End If

	' update WorkingQuests
	arr = Split(updateWorkingQuests,SUBSEPARATOR)
	If Ubound(arr)>=0 Then
		If ((UBound(arr)+1) Mod 6)<>0 Then
			Call Error("invalid number of item parameter")
		End If

		For i=0 To UBound(arr) Step 6
			sql = sql &"exec UpdateWorkingQuest"&_
				" "& SqlQuot(characterName) &_
				","& arr(i) &_
				","& arr(i+1) &_
				","& arr(i+2) &_
				","& arr(i+3) &_
				","& arr(i+4) &_
				","& SqlQuot(arr(i+5)) &_
				",@nError output, @strSP output"&_
				vbCrLf
			setCount = setCount +1
			sql = sql & "if(@nError not in (0)) goto lblFail"& vbCrLf
		Next
	End If

	' update CompletedQuests
	arr = Split(updateCompletedQuests,SUBSEPARATOR)
	If Ubound(arr)>=0 Then
		If ((UBound(arr)+1) Mod 6)<>0 Then
			Call Error("invalid number of item parameter")
		End If

		For i=0 To UBound(arr) Step 6
			sql = sql &"exec UpdateCompletedQuest"&_
				" "& SqlQuot(characterName) &_
				","& arr(i) &_
				","& arr(i+1) &_
				","& arr(i+2) &_
				","& arr(i+3) &_
				","& arr(i+4) &_
				","& arr(i+5) &_
				",@nError output, @strSP output"&_
				vbCrLf
			setCount = setCount +1
			sql = sql & "if(@nError not in (0)) goto lblFail"& vbCrLf
		Next
	End If

	' delete QuickSlots
	arr = Split(deleteQuickslots,SUBSEPARATOR)
	If Ubound(arr)>=0 Then
		For i=0 To UBound(arr) Step 1
			sql = sql &"exec DeleteQuickSlot"&_
				" "& SqlQuot(characterName) &_
				","& arr(i) &_
				",@nError output, @strSP output"&_
				vbCrLf
			setCount = setCount +1
			sql = sql & "if(@nError not in (0)) goto lblFail"& vbCrLf
		Next
	End If

	' update QuickSlots
	arr = Split(updateQuickslots,SUBSEPARATOR)
	If Ubound(arr)>=0 Then
		If ((UBound(arr)+1) Mod 5)<>0 Then
			Call Error("invalid number of item parameter")
		End If

		For i=0 To UBound(arr) Step 5
			sql = sql &"exec UpdateQuickSlot"&_
				" "& SqlQuot(characterName) &_
				","& arr(i+1) &_
				","& arr(i) &_
				","& int(arr(i+2)) &_
				","& arr(i+3) &_
				","& SqlQuot(arr(i+4)) &_
				",@nError output, @strSP output"&_
				vbCrLf
			setCount = setCount +1
			sql = sql & "if(@nError not in (0)) goto lblFail"& vbCrLf
		Next
	End If

	' update Pets Owner
	arr = Split(updatePetOwner,SUBSEPARATOR)
	If Ubound(arr)>=0 Then
		For i=0 To UBound(arr) Step 1
			sql = sql &"exec UpdatePetOwner"&_
				" "& arr(i) &_
				","& SqlQuot(characterName) &_				
				",@nError output, @strSP output"&_
				vbCrLf
			setCount = setCount +1
			sql = sql & "if(@nError not in (0)) goto lblFail"& vbCrLf
		Next
	End If

	' delete Pets
	arr = Split(deletePets,SUBSEPARATOR)
	If Ubound(arr)>=0 Then
		For i=0 To UBound(arr) Step 1
			sql = sql &"exec DeletePet"&_
				" "& SqlQuot(characterName) &_
				","& arr(i) &_
				",@nError output, @strSP output"&_
				vbCrLf
			setCount = setCount +1
			sql = sql & "if(@nError not in (0)) goto lblFail"& vbCrLf
		Next
	End If

	' update Pets
	arr = Split(updatePets,SUBSEPARATOR)
	If Ubound(arr)>=0 Then
		If ((UBound(arr)+1) Mod 10)<>0 Then
			Call Error("invalid number of item parameter")
		End If

		For i=0 To UBound(arr) Step 10
			sql = sql &"exec UpdatePet"&_
				" "& SqlQuot(characterName) &_
				","& arr(i) &_
				","& SqlQuot(arr(i+1)) &_
				","& arr(i+2) &_
				","& arr(i+3) &_
				","& arr(i+4) &_
				","& int(arr(i+5)) &_
				","& arr(i+6) &_
				","& arr(i+7) &_
				","& arr(i+8) &_
				","& arr(i+9) &_
				",@nError output, @strSP output"&_
				vbCrLf
			setCount = setCount +1
			sql = sql & "if(@nError not in (0)) goto lblFail"& vbCrLf
		Next
	End If	

	' delete PetItems
	arr = Split(deletePetItems,SUBSEPARATOR)
	If Ubound(arr)>=0 Then
		If ((UBound(arr)+1) Mod 3)<>0 Then
			Call Error("invalid number of item parameter")
		End If

		For i=0 To UBound(arr) Step 3
			sql = sql &"exec DeletePetItem"&_
				" "& SqlQuot(characterName) &_
				","& arr(i) &_
				","& arr(i+1) &_
				","& arr(i+2) &_
				",@nError output, @strSP output"&_
				vbCrLf
			setCount = setCount +1
			sql = sql & "if(@nError not in (0)) goto lblFail"& vbCrLf
		Next
	End If

	' update PetItems
	arr = Split(updatePetItems,SUBSEPARATOR)
	If Ubound(arr)>=0 Then
		If ((UBound(arr)+1) Mod 8)<>0 Then
			Call Error("invalid number of item parameter")
		End If

		For i=0 To UBound(arr) Step 8			
			sql = sql &"exec UpdatePetItem"&_
				" "& SqlQuot(characterName) &_
				","& arr(i) &_
				","& arr(i+1) &_
				","& arr(i+2) &_
				","& arr(i+3) &_
				","& arr(i+4)
			If IsNumeric(arr(i+5)) Then
				sql = sql &","& arr(i+5)
			Else
				sql = sql &",null"
			End If				
			sql = sql &","& arr(i+6) &_
				","& SqlQuot(arr(i+7)) &_
				",@nError output, @strSP output"&_
				vbCrLf
			setCount = setCount +1
			sql = sql & "if(@nError not in (0)) goto lblFail"& vbCrLf
		Next
	End If

	' delete PetTraining
	arr = Split(deletePetTraining,SUBSEPARATOR)
	If Ubound(arr)>=0 Then
		For i=0 To UBound(arr) Step 1
			sql = sql &"exec DeletePetTraining"&_
				" "& SqlQuot(characterName) &_
				","& arr(i) &_
				",@nError output, @strSP output"&_
				vbCrLf
			setCount = setCount +1
			sql = sql & "if(@nError not in (0)) goto lblFail"& vbCrLf
		Next
	End If	

	' insert PetTraining
	arr = Split(insertPetTraining,SUBSEPARATOR)
	If Ubound(arr)>=0 Then
		If ((UBound(arr)+1) Mod 8)<>0 Then
			Call Error("invalid number of item parameter")
		End If
		
		For i=0 To UBound(arr) Step 8			
			sql = sql &"exec InsertPetTraining"&_
				" "& SqlQuot(characterName) &_
				","& arr(i) &_
				","& arr(i+1) &_
				","& arr(i+2) &_
				","& arr(i+3) &_
				","& SqlQuot(arr(i+4)) &_
				","& arr(i+5) &_
				","& SqlQuot(arr(i+6)) &_
				","& SqlQuot(arr(i+7)) &_
				",@nError output, @strSP output"&_
				vbCrLf
			setCount = setCount +1
			sql = sql & "if(@nError not in (0)) goto lblFail"& vbCrLf
		Next
	End If

	
	Set conn = Server.CreateObject("ADODB.Connection")
	conn.Open(characterDBconnectionString)
	conn.Execute(sql)
	conn.Close()
	
	Call Ok(characterName & SEPARATOR & callBack)
%>