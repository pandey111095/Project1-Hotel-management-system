Attribute VB_Name = "Module1"
Public C As ADODB.Connection
Public R As ADODB.Recordset
Public CC As ADODB.Connection
Public RR As ADODB.Recordset
Public SS As String
Public S As String
Public P As String
Public T As String
Public NOOFEMP As Integer
Public NOOFGUEST As Integer
Public USER As String
Public Function CON()
Set C = New ADODB.Connection
C.Open "Provider=MSDAORA.1;User ID=SANDEEP/KOHLI;Persist Security Info=True"
Set R = New ADODB.Recordset
End Function

