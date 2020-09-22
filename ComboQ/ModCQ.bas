Attribute VB_Name = "ModCQ"
Option Explicit

Global Const TotalPerson = 5
Global Const DefaultText = 0

Type Person
    pName As String
    pCountry As String
    pAge As Byte
End Type

Global MyPerson() As Person '*** Array Declaration ***
Public Sub GetMyPersons()
'*** Simple loading imaginary person data ***

    MyPerson(0).pName = "Sarvan"
    MyPerson(0).pCountry = "Inida"
    MyPerson(0).pAge = 23
    
    MyPerson(1).pName = "Zindal"
    MyPerson(1).pCountry = "France"
    MyPerson(1).pAge = 21
    
    MyPerson(2).pName = "Alex"
    MyPerson(2).pCountry = "Russia"
    MyPerson(2).pAge = 40
    
    MyPerson(3).pName = "Daniel"
    MyPerson(3).pCountry = "Luxembourg"
    MyPerson(3).pAge = 42
    
    MyPerson(4).pName = "Sukanto"
    MyPerson(4).pCountry = "India"
    MyPerson(4).pAge = 51
    
    MyPerson(5).pName = "Ive"
    MyPerson(5).pCountry = "Belgium"
    MyPerson(5).pAge = 25
    
    
End Sub


