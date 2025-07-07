Dim fullname,dataBasename
fullname = HMIRunTime.Tags("@DatasourceNameRT").Read
msgbox ("数据源全名：" & fullname)
dataBasename = Mid(fullname,1,len(fullname)-1)
msgbox ("数据源名称：" & dataBasename)

Dim sqlConnectionString

sqlConnectionString = "Provider=SQLOLEDB;Data Source=.\Wincc;Initial Catalog=" & dataBasename & ";Integrated Security=SSPI;"
msgbox ("连接字符串：" & sqlConnectionString)
Dim conn

set conn = CreateObject("ADODB.Connection")
conn.ConnectionString = sqlConnectionString
conn.cursorlocation = 3
conn.Open

msgbox ("连接成功！")