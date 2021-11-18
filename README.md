# subject-sql-group-by

sqlcmd

```
sp_configure 'show advanced options', 1;
go
RECONFIGURE;
go
sp_configure 'Ad Hoc Distributed Queries', 1;
go
RECONFIGURE;
EXEC master.dbo.sp_MSset_oledb_prop N'Microsoft.ACE.OLEDB.12.0', N'AllowInProcess', 1
go
EXEC master.dbo.sp_MSset_oledb_prop N'Microsoft.ACE.OLEDB.12.0', N'DynamicParameters', 1 
go
```
