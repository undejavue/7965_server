Dim s 'As SQLDMO.SQLServer
Dim d 'As SQLDMO.Database

Set s = CreateObject("SQLDMO.SQLServer")

s.LoginSecure = True
s.LoginTimeout = 10
s.Connect ".\WinCC"

For Each d In s.Databases
    If Left(d.Name, 3) = "CC_" Then
        d.DBOption.RecoveryModel = 0 ' SQLDMORECOVERY_Simple
        d.Shrink 1, 0                ' SQLDMOShrink_Default
        d.DBOption.RecoveryModel = 2 ' SQLDMORECOVERY_Full
    End If
Next
MsgBox "End of Shrink"