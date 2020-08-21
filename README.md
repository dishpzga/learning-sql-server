# sql-server-learning

## 1 高级功能 使用BCP导出数据到Excel
- 项目中经常要数据导出Excel，可以使用BCP 调用DOS命令`xp_cmdshell`实现导出
- 出于系统安全考虑 SQL Server默认关闭`xp_cmdshell`，使用前需要开启，使用后关闭
- 为了使用方便，我们创建存储过程 ExportDataToExcel 需要导出时直接调用存储过程即可
- 注意：导出SQL 最好写在一行（亲测：最后一个字段与from table 必须在一行），否则会报错
- 创建存储过程
```tsql

---导出到Excel
---使用说明：
--        1.执行时所连接的服务器决定文件存放在哪个服务器
--        2.远程查询语句中，要加上数据库名
--更新：
--        2013.01.05:增加csv文件的支持

ALTER PROC ExportDataToExcel
     @QuerySql VARCHAR(max)
    ,@Server VARCHAR(20) 
    ,@User VARCHAR(20) = 'sa'
    ,@Password VARCHAR(20)
    ,@FilePath NVARCHAR(100) = 'D:\ExportFile.csv'
AS
    DECLARE @tmp VARCHAR(50) = '[##Table' + CONVERT(VARCHAR(36),NEWID())+']'
    BEGIN TRY
        DECLARE @Sql VARCHAR(max),@DataSource VARCHAR(max)='';
        --判断是否为远程服务器
        IF @Server <> '.' AND @Server <> '127.0.0.1'
            SET @DataSource = 'OPENDATASOURCE(''SQLOLEDB'',''Data Source='+@Server+';User ID='+@User+';Password='+@Password+''').'
        --将结果集导出到指定的数据库
        SET @Sql = REPLACE(@QuerySql,' from ',' into '+@tmp+ ' from ' + @DataSource)
        PRINT @Sql
        EXEC(@Sql)
        
        DECLARE @Columns VARCHAR(max) = '',@Data NVARCHAR(max)=''
        SELECT @Columns = @Columns + ',''' + name +''''--获取列名（xp_cmdshell导出文件没有列名）
            ,@Data = @Data + ',Convert(Nvarchar,[' + name +'])'--将结果集所在的字段更新为nvarchar（避免在列名和数据union的时候类型冲突）
        FROM tempdb.sys.columns WHERE object_id = OBJECT_ID('tempdb..'+@tmp)
        SELECT @Data  = 'SELECT ' + SUBSTRING(@Data,2,LEN(@Data)) + ' FROM ' + @tmp
        SELECT @Columns =  'Select ' + SUBSTRING(@Columns,2,LEN(@Columns))
        --使用xp_cmdshell的bcp命令将数据导出
        EXEC sp_configure 'xp_cmdshell',1
        RECONFIGURE
        DECLARE @cmd NVARCHAR(4000) = 'bcp "' + @Columns+' Union All ' + @Data+'" queryout ' + @FilePath + ' -c' + CASE WHEN RIGHT(@FilePath,4) = '.csv' THEN ' -t,' ELSE '' END + ' -T'
        PRINT @cmd
        exec sys.xp_cmdshell @cmd
        EXEC sp_configure 'xp_cmdshell',0
        RECONFIGURE
        EXEC('DROP TABLE ' + @tmp)
    END TRY
    BEGIN CATCH
        --处理异常
        IF OBJECT_ID('tempdb..'+@tmp) IS NOT NULL
            EXEC('DROP TABLE ' + @tmp)
        EXEC sp_configure 'xp_cmdshell',0
        RECONFIGURE
        
        SELECT ERROR_MESSAGE()
    END CATCH
```
- 使用示例：
```tsql
-- 第一步 开启高级功能
EXEC sp_configure 'show advanced options', 1; 
RECONFIGURE; 
EXEC sp_configure 'xp_cmdshell',1
RECONFIGURE;

-- 第二部 执行导出
declare @sql varchar(max)
declare @from varchar(20) = '2020/08/01 00:00:00'
declare @to varchar(20) = '2020/08/21 00:00:00'
set @sql = 'select wp.ProInfo_ResVar4 as ''工单流水号'' ,convert(varchar(19),wp.PrcInfo_CreatTime,120) as ''创建时间'' from super.dbo.TWork_ProcessInfo wp where wp.PrcInfo_CreatTime >= '''+@from+''' and wp.PrcInfo_CreatTime < '''+@to+''''
exec dbo.ExportDataToExcel @QuerySql=@sql, @Server='127.0.0.1',@Password='123456',@FilePath='D:\ProcessInfo.xls'

--第三步，关闭高级功能
EXEC sp_configure 'xp_cmdshell',0
RECONFIGURE;

EXEC sp_configure 'show advanced options', 0; 
    
RECONFIGURE; 
```