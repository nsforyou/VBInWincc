Dim fullName,dataBaseName
' Description: This script retrieves the full name of a database from a tag and extracts the database name by removing the last character.
fullName = HMIRuntime.Tags("@DatasourceNameRT").Read
dataBaseName = mid(fullName,1,len(fullName)-1)

' The connection string is constructed using the extracted database name.
' It uses the SQLOLEDB provider to connect to a SQL Server database with integrated security
Dim sqlConnectionString
sqlConnectionString = "Provider=SQLOLEDB;Data Source=.\Wincc;Initial Catalog=" & dataBaseName & ";Integrated Security=SSPI;"


Dim conn
Set conn = CreateObject("ADODB.Connection")

conn.ConnectionString = sqlConnectionString
conn.CursorLocation = 3 'adUseClient
conn.Open

Dim sql
sql = "SELECT name FROM PW_User WHERE GRPID > 0"

Dim rs
Set rs = CreateObject("ADODB.Recordset")

rs.open sql, conn, 1, 3 'adOpenKeyset, adLockOptimistic

Dim userCount
userCount = rs.RecordCount

Dim i,cmb
Set cmb = ScreenItems("cmoUserList")
cmb.NumberLines = userCount

If userCount > 0 Then
    rs.MoveFirst
    For i = 1 To userCount
        cmb.SelIndex = i
        cmb.SelText = Trim(rs.Fields("name").Value)
        rs.MoveNext
    Next
Else
End If
cmb.SelIndex = 1
rs.Close
conn.Close
Set rs = Nothing
Set conn = Nothing
Set cmb = Nothing
' The script ends here, having populated the combo box with user names from the database.
' It is important to ensure that the database connection is properly closed and objects are set to Nothing
