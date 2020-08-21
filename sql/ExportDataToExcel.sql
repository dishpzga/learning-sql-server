/**
-- ʹ��ʵ��

-- ��һ�� �����߼�����
EXEC sp_configure 'show advanced options', 1; 
RECONFIGURE; 
EXEC sp_configure 'xp_cmdshell',1
RECONFIGURE;

-- �ڶ��� ִ�е���
declare @sql varchar(5000)
declare @from varchar(20) = '2020/08/01 00:00:00'
declare @to varchar(20) = '2020/08/21 00:00:00'
set @sql = 'select wp.ProInfo_ResVar4 as ''������ˮ��'' 
,area.Area_Name as ''������'' 
,dType.dict_name as ''��������'' 
,dFrom.dict_name as ''������Դ'' 
,dLevel.dict_name as ''��������'' 
,wp.PrcInfo_Info as ''��������'' 
, convert(varchar(19),wp.PrcInfo_PrcEndTime,120) as ''��ֹʱ��'' 
,wp.ProInfo_ResVar3 as ''������'' ,wp.ProInfo_ResVar8 as ''�ظ��绰'' 
,wp.ProInfo_ResVar9 as ''������ע'' ,u.UserInfo_name as ''������'' 
,wp.ProInfo_ResVar5 as ''������'' 
, convert(varchar(19),wp.ProInfo_ResDate3,120)  as ''�ظ�ʱ��'' 
,wp.ProInfo_ResVar10 as ''��ע'' 
,cu.UserInfo_name as ''������'' 
,convert(varchar(19),wp.PrcInfo_CreatTime,120) as ''����ʱ��'' from super.dbo.TWork_ProcessInfo wp 
left join super.dbo.TManager_Area area on area.Area_ID = wp.PrcInfo_AreaID 
left join super.dbo.TBBase_UserInfo cu on cu.UserInfo_ID = wp.PrcInfo_CUserID 
left join super.dbo.dict dType on dType.dict_id = wp.ProInfo_ResInt1 and dType.dict_type = ''WorkitemType'' 
left join super.dbo.dict dFrom on dFrom.dict_id = wp.ProInfo_ResInt2 and dFrom.dict_type = ''WorkitemFrom'' 
left join super.dbo.dict dLevel on dLevel.dict_id = wp.ProInfo_ResInt3 and dLevel.dict_type = ''WorkitemLevel'' 
left join super.dbo.TBBase_UserInfo u on u.UserInfo_ID = wp.ProInfo_PaiUserID 
where wp.PrcInfo_CreatTime >= '''+@from+''' and wp.PrcInfo_CreatTime < '''+@to+''''
exec dbo.ExportDataToExcel @QuerySql=@sql, @Server='127.0.0.1',@Password='123456',@FilePath='D:\ProcessInfo.xls'


--���������رո߼�����
EXEC sp_configure 'xp_cmdshell',0
RECONFIGURE;

EXEC sp_configure 'show advanced options', 0; 
    
RECONFIGURE; 

*/

---������Excel
---ʹ��˵����
--        1.ִ��ʱ�����ӵķ����������ļ�������ĸ�������
--        2.Զ�̲�ѯ����У�Ҫ�������ݿ���
--���£�
--        2013.01.05:����csv�ļ���֧��

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
        --�ж��Ƿ�ΪԶ�̷�����
        IF @Server <> '.' AND @Server <> '127.0.0.1'
            SET @DataSource = 'OPENDATASOURCE(''SQLOLEDB'',''Data Source='+@Server+';User ID='+@User+';Password='+@Password+''').'
        --�������������ָ�������ݿ�
        SET @Sql = REPLACE(@QuerySql,' from ',' into '+@tmp+ ' from ' + @DataSource)
        PRINT @Sql
        EXEC(@Sql)
        
        DECLARE @Columns VARCHAR(max) = '',@Data NVARCHAR(max)=''
        SELECT @Columns = @Columns + ',''' + name +''''--��ȡ������xp_cmdshell�����ļ�û��������
            ,@Data = @Data + ',Convert(Nvarchar,[' + name +'])'--����������ڵ��ֶθ���Ϊnvarchar������������������union��ʱ�����ͳ�ͻ��
        FROM tempdb.sys.columns WHERE object_id = OBJECT_ID('tempdb..'+@tmp)
        SELECT @Data  = 'SELECT ' + SUBSTRING(@Data,2,LEN(@Data)) + ' FROM ' + @tmp
        SELECT @Columns =  'Select ' + SUBSTRING(@Columns,2,LEN(@Columns))
        --ʹ��xp_cmdshell��bcp������ݵ���
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
        --�����쳣
        IF OBJECT_ID('tempdb..'+@tmp) IS NOT NULL
            EXEC('DROP TABLE ' + @tmp)
        EXEC sp_configure 'xp_cmdshell',0
        RECONFIGURE
        
        SELECT ERROR_MESSAGE()
    END CATCH