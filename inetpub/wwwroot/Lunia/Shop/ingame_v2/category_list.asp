<%	Option Explicit %>
<!--#include virtual="/include/connect.asp"-->
<!--#include virtual="/include/common.asp"-->
<%
	Dim i,j,k
	Dim arrCat
	Dim retString : retString = ""
	Dim rCodePage
	If Ubound(Parameters)<0 Then Call Error("not enough parameter")
	
	rCodePage	= Parameters(0)
	
	If rCodePage = "Portuguese" Then
		arrCat = Array( _		
			Array(0,1,0,	"Free",null) _
		,	Array(0,2,0,	"Renascimento",null) _
		,	Array(0,3,0,	"Visual Rare",null) _
		,	Array(0,7,0,	"Visual INT/VIT/DEX",null) _
		
		,	Array(1,7,0,	"INT",null) _
		,	Array(2,7,10,	"Arma",null) _
		,	Array(2,7,11,	"Cabeça",null) _
		,	Array(2,7,12,	"Peito",null) _
		,	Array(2,7,13,	"Perna",null) _
		,	Array(2,7,14,	"Mão",null) _
		,	Array(2,7,15,	"Pés",null) _
		,	Array(2,7,16,	"Suporte",null) _
		,	Array(2,7,17,	"Máscara",null) _
		,	Array(2,7,18,	"Costas",null) _
		,	Array(2,7,19,	"Cintura",null) _
		,	Array(2,7,20,	"ETC1",null) _
		,	Array(2,7,21,	"ETC2",null) _
		
		,	Array(1,11,0,	"VIT",null) _
		,	Array(2,11,10,	"Arma",null) _
		,	Array(2,11,11,	"Cabeça",null) _
		,	Array(2,11,12,	"Peito",null) _
		,	Array(2,11,13,	"Perna",null) _
		,	Array(2,11,14,	"Mão",null) _
		,	Array(2,11,15,	"Pés",null) _
		,	Array(2,11,16,	"Suporte",null) _
		,	Array(2,11,17,	"Máscara",null) _
		,	Array(2,11,18,	"Costas",null) _
		,	Array(2,11,19,	"Cintura",null) _
		,	Array(2,11,20,	"ETC1",null) _
		,	Array(2,11,21,	"ETC2",null) _
		
		,	Array(1,12,0,	"DEX",null) _
		,	Array(2,12,10,	"Arma",null) _
		,	Array(2,12,11,	"Cabeça",null) _
		,	Array(2,12,12,	"Peito",null) _
		,	Array(2,12,13,	"Perna",null) _
		,	Array(2,12,14,	"Mão",null) _
		,	Array(2,12,15,	"Pés",null) _
		,	Array(2,12,16,	"Suporte",null) _
		,	Array(2,12,17,	"Máscara",null) _
		,	Array(2,12,18,	"Costas",null) _
		,	Array(2,12,19,	"Cintura",null) _
		,	Array(2,12,20,	"ETC1",null) _
		,	Array(2,12,21,	"ETC2",null) _
		
		,	Array(1,13,0,	"BOX",null) _
		
		,	Array(0,8,0,	"Visual",null) _
		,	Array(0,9,0,	"Raid Visual",null) _
		,	Array(0,10,0,	"Dummy Visual",null) _
		,	Array(0,4,0,	"Utilidade",null) _
		,	Array(0,5,0,	"Reborn Customs",null) _
		,	Array(0,6,0,	"Mascotes",null) _
		,	Array(0,14,0,	"Emojis",null) _
		,	Array(0,15,0,	"Asas",null) _
		,	Array(0,16,0,	"Passe de Batalha",null) _
		,	Array(0,17,0,	"Pacote de Iniciante",null) _
		
		,	Array(0,18,1,	"Cash Exclusivo BP",null) _
		
		,	Array(1,18,0,	"INT",null) _
		,	Array(2,18,10,	"Arma",null) _
		,	Array(2,18,11,	"Cabeça",null) _
		,	Array(2,18,12,	"Peito",null) _
		,	Array(2,18,13,	"Perna",null) _
		,	Array(2,18,14,	"Mão",null) _
		,	Array(2,18,15,	"Pés",null) _
		,	Array(2,18,16,	"Suporte",null) _
		,	Array(2,18,17,	"Rosto",null) _
		,	Array(2,18,18,	"Costas",null) _
		,	Array(2,18,19,	"Cintura",null) _
		,	Array(2,18,20,	"ETC1",null) _
		,	Array(2,18,21,	"ETC2",null) _
		
		,	Array(1,19,0,	"VIT",null) _
		,	Array(2,19,10,	"Arma",null) _
		,	Array(2,19,11,	"Cabeça",null) _
		,	Array(2,19,12,	"Peito",null) _
		,	Array(2,19,13,	"Perna",null) _
		,	Array(2,19,14,	"Mão",null) _
		,	Array(2,19,15,	"Pés",null) _
		,	Array(2,19,16,	"Suporte",null) _
		,	Array(2,19,17,	"Rosto",null) _
		,	Array(2,19,18,	"Costas",null) _
		,	Array(2,19,19,	"Cintura",null) _
		,	Array(2,19,20,	"ETC1",null) _
		,	Array(2,19,21,	"ETC2",null) _
		
		,	Array(1,20,00,	"DEX",null) _
		,	Array(2,20,10,	"Arma",null) _
		,	Array(2,20,11,	"Cabeça",null) _
		,	Array(2,20,12,	"Peito",null) _
		,	Array(2,20,13,	"Perna",null) _
		,	Array(2,20,14,	"Mão",null) _
		,	Array(2,20,15,	"Pés",null) _
		,	Array(2,20,16,	"Suporte",null) _
		,	Array(2,20,17,	"Rosto",null) _
		,	Array(2,20,18,	"Costas",null) _
		,	Array(2,20,19,	"Cintura",null) _
		,	Array(2,20,20,	"ETC1",null) _
		,	Array(2,20,21,	"ETC2",null) _
		
		)
	Else
		arrCat = Array( _		
			Array(0,1,0,	"Free",null) _
		,	Array(0,2,0,	"Rebirths",null) _
		,	Array(0,3,0,	"Visual Rare",null) _
		,	Array(0,7,0,	"Visual INT/VIT/DEX",null) _
		
		,	Array(1,7,0,	"INT",null) _
		,	Array(2,7,10,	"Weapon",null) _
		,	Array(2,7,11,	"Head",null) _
		,	Array(2,7,12,	"Chest",null) _
		,	Array(2,7,13,	"Leg",null) _
		,	Array(2,7,14,	"Hand",null) _
		,	Array(2,7,15,	"Foot",null) _
		,	Array(2,7,16,	"Support",null) _
		,	Array(2,7,17,	"Mask",null) _
		,	Array(2,7,18,	"Back",null) _
		,	Array(2,7,19,	"Hip",null) _
		,	Array(2,7,20,	"ETC1",null) _
		,	Array(2,7,21,	"ETC2",null) _
		
		,	Array(1,11,0,	"VIT",null) _
		,	Array(2,11,10,	"Weapon",null) _
		,	Array(2,11,11,	"Head",null) _
		,	Array(2,11,12,	"Chest",null) _
		,	Array(2,11,13,	"Leg",null) _
		,	Array(2,11,14,	"Hand",null) _
		,	Array(2,11,15,	"Foot",null) _
		,	Array(2,11,16,	"Support",null) _
		,	Array(2,11,17,	"Mask",null) _
		,	Array(2,11,18,	"Back",null) _
		,	Array(2,11,19,	"Hip",null) _
		,	Array(2,11,20,	"ETC1",null) _
		,	Array(2,11,21,	"ETC2",null) _
		
		,	Array(1,12,0,	"DEX",null) _
		,	Array(2,12,10,	"Weapon",null) _
		,	Array(2,12,11,	"Head",null) _
		,	Array(2,12,12,	"Chest",null) _
		,	Array(2,12,13,	"Leg",null) _
		,	Array(2,12,14,	"Hand",null) _
		,	Array(2,12,15,	"Foot",null) _
		,	Array(2,12,16,	"Support",null) _
		,	Array(2,12,17,	"Mask",null) _
		,	Array(2,12,18,	"Back",null) _
		,	Array(2,12,19,	"Hip",null) _
		,	Array(2,12,20,	"ETC1",null) _
		,	Array(2,12,21,	"ETC2",null) _
		
		,	Array(1,13,0,	"BOX",null) _
		
		,	Array(0,8,0,	"Visual",null) _
		,	Array(0,9,0,	"Raid Visual",null) _
		,	Array(0,10,0,	"Dummy Visual",null) _
		,	Array(0,4,0,	"Utility",null) _
		,	Array(0,5,0,	"Reborn Customs",null) _
		,	Array(0,6,0,	"Pets",null) _
		,	Array(0,14,0,	"Emojis",null) _
		,	Array(0,15,0,	"Wings",null) _
		,	Array(0,16,0,	"Battle Pass",null) _
		,	Array(0,17,0,	"Starter Pack",null) _
		
		,	Array(0,18,1,	"Exclusive Cash BP",null) _
		
		,	Array(1,18,0,	"INT",null) _
		,	Array(2,18,10,	"Weapon",null) _
		,	Array(2,18,11,	"Head",null) _
		,	Array(2,18,12,	"Chest",null) _
		,	Array(2,18,13,	"Leg",null) _
		,	Array(2,18,14,	"Hand",null) _
		,	Array(2,18,15,	"Foot",null) _
		,	Array(2,18,16,	"Support",null) _
		,	Array(2,18,17,	"Mask",null) _
		,	Array(2,18,18,	"Back",null) _
		,	Array(2,18,19,	"Hip",null) _
		,	Array(2,18,20,	"ETC1",null) _
		,	Array(2,18,21,	"ETC2",null) _
		
		,	Array(1,19,0,	"VIT",null) _
		,	Array(2,19,10,	"Weapon",null) _
		,	Array(2,19,11,	"Head",null) _
		,	Array(2,19,12,	"Chest",null) _
		,	Array(2,19,13,	"Leg",null) _
		,	Array(2,19,14,	"Hand",null) _
		,	Array(2,19,15,	"Foot",null) _
		,	Array(2,19,16,	"Support",null) _
		,	Array(2,19,17,	"Mask",null) _
		,	Array(2,19,18,	"Back",null) _
		,	Array(2,19,19,	"Hip",null) _
		,	Array(2,19,20,	"ETC1",null) _
		,	Array(2,19,21,	"ETC2",null) _
		
		,	Array(1,20,00,	"DEX",null) _
		,	Array(2,20,10,	"Weapon",null) _
		,	Array(2,20,11,	"Head",null) _
		,	Array(2,20,12,	"Chest",null) _
		,	Array(2,20,13,	"Leg",null) _
		,	Array(2,20,14,	"Hand",null) _
		,	Array(2,20,15,	"Foot",null) _
		,	Array(2,20,16,	"Support",null) _
		,	Array(2,20,17,	"Mask",null) _
		,	Array(2,20,18,	"Back",null) _
		,	Array(2,20,19,	"Hip",null) _
		,	Array(2,20,20,	"ETC1",null) _
		,	Array(2,20,21,	"ETC2",null) _
		)
	End If
	
	' Array(lv,nCategory1,nCategory2,strName,classNumber)
	
	Dim n1,n2,n3
	n1 = 0 : n2 = 0 : n3 = 0
	For i=0 To Ubound(arrCat)
		If IsArray(arrCat(i)) Then
			If retString<>"" Then retString = retString & SEPARATOR
			
			If arrCat(i)(0)=0 Then
				n1 = n1 + 1
				n2 = 0 : n3 = 0
			ElseIf arrCat(i)(0)=1 Then
				n2 = n2 + 1
				n3 = 0
			ElseIf arrCat(i)(0)=2 Then
				n3 = n3 + 1
			End If
			
			retString = retString &_
				n1 & SUBSEPARATOR & n2 & SUBSEPARATOR & n3 &_
				SUBSEPARATOR & arrCat(i)(1) &_
				SUBSEPARATOR & arrCat(i)(2) &_
				SUBSEPARATOR & arrCat(i)(3) &_
				SUBSEPARATOR & arrCat(i)(4)
		End If
	Next
	
	Call Ok(retString)
%>