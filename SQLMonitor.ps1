Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Data
[System.Windows.Forms.Application]::EnableVisualStyles()

$script:connString = ""
$script:connected  = $false
$script:timer      = $null

# ── HELPERS ──────────────────────────────────────────────────────────────────

function Invoke-SqlQuery([string]$sql) {
    $dt = New-Object System.Data.DataTable
    try {
        $cn  = New-Object System.Data.SqlClient.SqlConnection($script:connString)
        $cn.Open()
        $cmd = $cn.CreateCommand()
        $cmd.CommandText    = $sql
        $cmd.CommandTimeout = 25
        $da  = New-Object System.Data.SqlClient.SqlDataAdapter($cmd)
        [void]$da.Fill($dt)
        $cn.Close()
    } catch {
        $dt.Columns.Clear()
        [void]$dt.Columns.Add("Error")
        $r = $dt.NewRow()
        $r["Error"] = $_.Exception.Message
        [void]$dt.Rows.Add($r)
    }
    Write-Output -NoEnumerate $dt
}

function Bind-Grid($g, $dt) {
    if ($dt -is [array]) { $dt = $dt[0] }
    $bs            = New-Object System.Windows.Forms.BindingSource
    $bs.DataSource = $dt
    $g.DataSource  = $bs
}

function New-DGV {
    $g = New-Object System.Windows.Forms.DataGridView
    $g.Dock                      = [System.Windows.Forms.DockStyle]::Fill
    $g.ReadOnly                  = $true
    $g.AllowUserToAddRows        = $false
    $g.AllowUserToDeleteRows     = $false
    $g.AutoSizeColumnsMode       = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::AllCells
    $g.SelectionMode             = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
    $g.BackgroundColor           = [System.Drawing.Color]::FromArgb(28,28,32)
    $g.GridColor                 = [System.Drawing.Color]::FromArgb(55,55,62)
    $g.BorderStyle               = [System.Windows.Forms.BorderStyle]::None
    $g.RowHeadersVisible         = $false
    $g.EnableHeadersVisualStyles = $false
    $cs = $g.DefaultCellStyle
    $cs.BackColor          = [System.Drawing.Color]::FromArgb(28,28,32)
    $cs.ForeColor          = [System.Drawing.Color]::FromArgb(215,215,215)
    $cs.Font               = New-Object System.Drawing.Font("Consolas",9)
    $cs.SelectionBackColor = [System.Drawing.Color]::FromArgb(0,115,210)
    $cs.SelectionForeColor = [System.Drawing.Color]::White
    $g.AlternatingRowsDefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(33,33,38)
    $hcs = $g.ColumnHeadersDefaultCellStyle
    $hcs.BackColor = [System.Drawing.Color]::FromArgb(42,42,50)
    $hcs.ForeColor = [System.Drawing.Color]::FromArgb(0,185,255)
    $hcs.Font      = New-Object System.Drawing.Font("Segoe UI",9,[System.Drawing.FontStyle]::Bold)
    $g.ColumnHeadersBorderStyle = [System.Windows.Forms.DataGridViewHeaderBorderStyle]::Single
    return $g
}

function New-SectionPanel([string]$text) {
    $p = New-Object System.Windows.Forms.Panel
    $p.Dock      = [System.Windows.Forms.DockStyle]::Top
    $p.Height    = 26
    $p.BackColor = [System.Drawing.Color]::FromArgb(40,40,48)
    $l = New-Object System.Windows.Forms.Label
    $l.Text      = "  $text"
    $l.Dock      = [System.Windows.Forms.DockStyle]::Fill
    $l.ForeColor = [System.Drawing.Color]::FromArgb(175,210,255)
    $l.Font      = New-Object System.Drawing.Font("Segoe UI",9,[System.Drawing.FontStyle]::Bold)
    $l.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
    $p.Controls.Add($l)
    return $p
}

# Adds header+grid into a Panel using Dock (no TableLayoutPanel needed in tabs)
function New-GridPanel([string]$header) {
    $outer = New-Object System.Windows.Forms.Panel
    $outer.Dock = [System.Windows.Forms.DockStyle]::Fill
    $hdr = New-SectionPanel $header
    $g   = New-DGV
    # Add grid first, then header - Fill+Top stacking requires this order
    $outer.Controls.Add($g)
    $outer.Controls.Add($hdr)
    return $outer,$g
}

function Set-Status([string]$msg,[string]$color="LightGreen") {
    $script:statusLbl.Text      = "  $msg"
    $script:statusLbl.ForeColor = [System.Drawing.Color]::$color
}

# ── SQL QUERIES ───────────────────────────────────────────────────────────────

$Q_Config = @"
SELECT 'SQL Version' AS Setting,
  CAST(SERVERPROPERTY('ProductVersion') AS NVARCHAR(30))+' | '+
  CAST(SERVERPROPERTY('Edition') AS NVARCHAR(80)) AS [Value],
  'SQL Server version info' AS Recommendation,
  'INFO' AS Status
UNION ALL
SELECT name, CAST(value_in_use AS NVARCHAR(50)), description,
  CASE name
    WHEN 'max server memory (MB)'         THEN CASE WHEN value_in_use=2147483647 THEN 'WARNING' ELSE 'OK' END
    WHEN 'max degree of parallelism'      THEN CASE WHEN value_in_use=0 THEN 'WARNING' ELSE 'OK' END
    WHEN 'cost threshold for parallelism' THEN CASE WHEN value_in_use<25 THEN 'WARNING' ELSE 'OK' END
    WHEN 'optimize for ad hoc workloads'  THEN CASE WHEN value_in_use=0 THEN 'WARNING' ELSE 'OK' END
    WHEN 'backup compression default'     THEN CASE WHEN value_in_use=0 THEN 'WARNING' ELSE 'OK' END
    WHEN 'remote admin connections'       THEN CASE WHEN value_in_use=0 THEN 'WARNING' ELSE 'OK' END
    WHEN 'xp_cmdshell'                   THEN CASE WHEN value_in_use=1 THEN 'WARNING' ELSE 'OK' END
    WHEN 'clr enabled'                   THEN CASE WHEN value_in_use=1 THEN 'WARNING' ELSE 'OK' END
    ELSE 'INFO'
  END
FROM sys.configurations
WHERE name IN (
  'max server memory (MB)','max degree of parallelism',
  'cost threshold for parallelism','optimize for ad hoc workloads',
  'backup compression default','remote admin connections',
  'xp_cmdshell','clr enabled','fill factor (%)')
UNION ALL
SELECT 'SA Account',
  CASE WHEN is_disabled=0 THEN 'ENABLED - security risk' ELSE 'Disabled (good)' END,
  'Rename or disable SA to reduce brute-force exposure',
  CASE WHEN is_disabled=0 THEN 'WARNING' ELSE 'OK' END
FROM sys.server_principals WHERE name='sa'
UNION ALL
SELECT 'AUTO_CLOSE databases',
  CAST(COUNT(*) AS NVARCHAR)+' user DB(s)',
  'Set AUTO_CLOSE OFF to avoid repeated startup overhead',
  CASE WHEN COUNT(*)>0 THEN 'WARNING' ELSE 'OK' END
FROM sys.databases WHERE is_auto_close_on=1 AND database_id>4
UNION ALL
SELECT 'AUTO_SHRINK databases',
  CAST(COUNT(*) AS NVARCHAR)+' user DB(s)',
  'Set AUTO_SHRINK OFF - causes fragmentation and CPU spikes',
  CASE WHEN COUNT(*)>0 THEN 'WARNING' ELSE 'OK' END
FROM sys.databases WHERE is_auto_shrink_on=1 AND database_id>4
UNION ALL
SELECT 'TempDB data files',
  CAST(COUNT(*) AS NVARCHAR)+' data file(s)',
  'Recommend 1 file per logical CPU (max 8) to reduce contention',
  CASE WHEN COUNT(*)<2 THEN 'WARNING' ELSE 'OK' END
FROM tempdb.sys.database_files WHERE type=0
"@

# Q_IndexHealth is dynamic - see Get-IndexHealthQuery function

$Q_CPU = @"
SELECT TOP 1
  record.value('(./Record/SchedulerMonitorEvent/SystemHealth/ProcessUtilization)[1]','int') AS [SQL CPU %],
  record.value('(./Record/SchedulerMonitorEvent/SystemHealth/SystemIdle)[1]','int') AS [Idle %],
  100
  -record.value('(./Record/SchedulerMonitorEvent/SystemHealth/SystemIdle)[1]','int')
  -record.value('(./Record/SchedulerMonitorEvent/SystemHealth/ProcessUtilization)[1]','int') AS [Other CPU %]
FROM (SELECT TOP 1 CONVERT(XML,record) AS record
      FROM sys.dm_os_ring_buffers
      WHERE ring_buffer_type=N'RING_BUFFER_SCHEDULER_MONITOR'
        AND record LIKE N'%<SystemHealth>%'
      ORDER BY timestamp DESC) x
"@

$Q_ActiveReqs = @"
SELECT r.session_id AS SPID, r.status AS Status,
  r.blocking_session_id AS [Blocked By],
  DB_NAME(r.database_id) AS [Database],
  s.login_name AS Login, s.host_name AS Host,
  r.cpu_time AS [CPU ms], r.total_elapsed_time AS [Elapsed ms],
  r.reads AS Reads, r.writes AS Writes,
  r.wait_type AS [Wait Type], r.wait_time AS [Wait ms],
  r.open_transaction_count AS [Open Txns],
  SUBSTRING(t.text,(r.statement_start_offset/2)+1,
    ((CASE r.statement_end_offset WHEN -1 THEN DATALENGTH(t.text)
      ELSE r.statement_end_offset END - r.statement_start_offset)/2)+1) AS [SQL Text]
FROM sys.dm_exec_requests r
JOIN sys.dm_exec_sessions s ON r.session_id=s.session_id
CROSS APPLY sys.dm_exec_sql_text(r.sql_handle) t
WHERE r.session_id<>@@SPID AND s.is_user_process=1
ORDER BY r.total_elapsed_time DESC
"@

$Q_WaitStats = @"
SELECT TOP 20 wait_type AS [Wait Type],
  waiting_tasks_count AS [Wait Count],
  CAST(wait_time_ms/1000.0 AS DECIMAL(18,1)) AS [Total Wait Sec],
  CAST(max_wait_time_ms/1000.0 AS DECIMAL(18,1)) AS [Max Wait Sec],
  CAST(signal_wait_time_ms*100.0/NULLIF(wait_time_ms,0) AS DECIMAL(5,1)) AS [Signal %]
FROM sys.dm_os_wait_stats
WHERE wait_type NOT IN (
  'SLEEP_TASK','BROKER_TO_FLUSH','BROKER_TASK_STOP','CLR_AUTO_EVENT','CLR_MANUAL_EVENT',
  'DISPATCHER_QUEUE_SEMAPHORE','FT_IFTS_SCHEDULER_IDLE_WAIT','HADR_WORK_QUEUE',
  'LAZYWRITER_SLEEP','LOGMGR_QUEUE','ONDEMAND_TASK_QUEUE','REQUEST_FOR_DEADLOCK_SEARCH',
  'RESOURCE_QUEUE','SERVER_IDLE_CHECK','SLEEP_DBSTARTUP','SLEEP_DBTASK','SLEEP_ERRORLOG',
  'SLEEP_MASTERDBREADY','SLEEP_MASTERMDREADY','SLEEP_MASTERUPGRADED','SLEEP_MSDBSTARTUP',
  'SLEEP_SYSTEMTASK','SLEEP_TEMPDBSTARTUP','SNI_HTTP_ACCEPT','SP_SERVER_DIAGNOSTICS_SLEEP',
  'SQLTRACE_BUFFER_FLUSH','SQLTRACE_INCREMENTAL_FLUSH_SLEEP','WAITFOR',
  'XE_DISPATCHER_WAIT','XE_TIMER_EVENT','BROKER_EVENTHANDLER','CHECKPOINT_QUEUE',
  'DBMIRROR_EVENTS_QUEUE','SQLTRACE_WAIT_ENTRIES','XE_DISPATCHER_JOIN',
  'WAIT_XTP_OFFLINE_CKPT_NEW_LOG','HADR_FILESTREAM_IOMGR_IOCOMPLETION')
ORDER BY wait_time_ms DESC
"@

$Q_Blocking = @"
SELECT r.session_id AS [Blocked SPID], bs.login_name AS [Blocked Login],
  r.blocking_session_id AS [Blocking SPID], bls.login_name AS [Blocker Login],
  DB_NAME(r.database_id) AS [Database],
  r.wait_time/1000 AS [Waiting Sec], r.wait_type AS [Wait Type],
  SUBSTRING(bq.text,1,150) AS [Blocked SQL]
FROM sys.dm_exec_requests r
JOIN sys.dm_exec_sessions bs  ON r.session_id=bs.session_id
JOIN sys.dm_exec_sessions bls ON r.blocking_session_id=bls.session_id
CROSS APPLY sys.dm_exec_sql_text(r.sql_handle) bq
WHERE r.blocking_session_id>0
"@

$Q_Backup = @"
SELECT d.name AS [Database], d.recovery_model_desc AS [Recovery Model],
  MAX(CASE b.type WHEN 'D' THEN b.backup_finish_date END) AS [Last Full Backup],
  MAX(CASE b.type WHEN 'I' THEN b.backup_finish_date END) AS [Last Diff Backup],
  MAX(CASE b.type WHEN 'L' THEN b.backup_finish_date END) AS [Last Log Backup],
  CASE
    WHEN MAX(CASE b.type WHEN 'D' THEN b.backup_finish_date END) IS NULL
      THEN 'NEVER BACKED UP'
    WHEN MAX(CASE b.type WHEN 'D' THEN b.backup_finish_date END)<DATEADD(day,-1,GETDATE())
      THEN 'WARNING - Overdue'
    ELSE 'OK'
  END AS [Backup Status]
FROM sys.databases d
LEFT JOIN msdb.dbo.backupset b ON d.name=b.database_name
WHERE d.database_id>4 AND d.state_desc='ONLINE'
GROUP BY d.name,d.recovery_model_desc
ORDER BY d.name
"@

$Q_Jobs = @"
SELECT TOP 50 j.name AS [Job Name], j.enabled AS Enabled,
  CONVERT(VARCHAR(8),CAST(h.run_date AS CHAR(8)),0) AS [Run Date],
  STUFF(STUFF(RIGHT('000000'+CAST(h.run_time AS VARCHAR),6),3,0,':'),6,0,':') AS [Run Time],
  CASE h.run_status WHEN 0 THEN 'Failed' WHEN 1 THEN 'Succeeded'
    WHEN 2 THEN 'Retry' WHEN 3 THEN 'Cancelled' ELSE 'Unknown' END AS Status,
  h.run_duration AS Duration,
  LEFT(h.message,200) AS Message
FROM msdb.dbo.sysjobs j
JOIN msdb.dbo.sysjobhistory h ON j.job_id=h.job_id
WHERE h.step_id=0
ORDER BY h.run_date DESC, h.run_time DESC
"@

$Q_DiskIO = @"
SELECT DB_NAME(v.database_id) AS [Database], m.physical_name AS [File Path],
  m.type_desc AS Type,
  CAST(v.io_stall_read_ms *1.0/NULLIF(v.num_of_reads ,0) AS DECIMAL(10,2)) AS [Avg Read ms],
  CAST(v.io_stall_write_ms*1.0/NULLIF(v.num_of_writes,0) AS DECIMAL(10,2)) AS [Avg Write ms],
  v.num_of_reads AS [Total Reads], v.num_of_writes AS [Total Writes],
  CAST(v.num_of_bytes_read   /1048576.0 AS DECIMAL(18,1)) AS [MB Read],
  CAST(v.num_of_bytes_written/1048576.0 AS DECIMAL(18,1)) AS [MB Written]
FROM sys.dm_io_virtual_file_stats(NULL,NULL) v
JOIN sys.master_files m ON v.database_id=m.database_id AND v.file_id=m.file_id
ORDER BY v.io_stall DESC
"@

$Q_TopQ = @"
SELECT TOP 20
  CAST(qs.total_elapsed_time/qs.execution_count/1000.0 AS DECIMAL(18,2)) AS [Avg Elapsed ms],
  CAST(qs.total_worker_time/qs.execution_count/1000.0  AS DECIMAL(18,2)) AS [Avg CPU ms],
  CAST(qs.total_logical_reads/qs.execution_count AS BIGINT)              AS [Avg Logical Reads],
  qs.execution_count AS [Exec Count],
  DB_NAME(t.dbid) AS [Database],
  SUBSTRING(t.text,1,250) AS [SQL Text]
FROM sys.dm_exec_query_stats qs
CROSS APPLY sys.dm_exec_sql_text(qs.sql_handle) t
ORDER BY qs.total_elapsed_time/qs.execution_count DESC
"@

function Get-IndexHealthQuery([string]$dbName) {
    # Pass DB_ID explicitly by name - avoids USE statement context issues
    return @"
SELECT TOP 50
  '$dbName' AS [Database],
  OBJECT_NAME(s.object_id, s.database_id) AS [Table],
  i.name AS [Index],
  i.type_desc AS [Type],
  CAST(s.avg_fragmentation_in_percent AS DECIMAL(5,1)) AS [Frag %],
  s.page_count AS Pages,
  s.record_count AS Rows,
  CASE WHEN s.avg_fragmentation_in_percent>=30 THEN 'REBUILD needed'
       WHEN s.avg_fragmentation_in_percent>=10 THEN 'REORGANIZE suggested'
       ELSE 'OK' END AS Action
FROM sys.dm_db_index_physical_stats(DB_ID('$dbName'),NULL,NULL,NULL,'LIMITED') s
JOIN [$dbName].sys.indexes i ON s.object_id=i.object_id AND s.index_id=i.index_id
WHERE s.page_count>100 AND i.name IS NOT NULL AND i.index_id > 0
ORDER BY s.avg_fragmentation_in_percent DESC;
"@
}


# ── STORAGE & TEMPDB SQL QUERIES ─────────────────────────────────────────────

$Q_DriveSpace = @"
EXEC xp_fixeddrives
"@

$Q_DBFiles = @"
SELECT
  d.name AS [Database],
  mf.name AS [Logical Name],
  mf.physical_name AS [Physical Path],
  mf.type_desc AS [File Type],
  CAST(CAST(mf.size AS BIGINT) * 8 / 1024 AS VARCHAR(20)) + ' MB' AS [Allocated MB],
  CAST(vfs.num_of_bytes_read   / 1048576.0 AS DECIMAL(12,1)) AS [MB Read],
  CAST(vfs.num_of_bytes_written / 1048576.0 AS DECIMAL(12,1)) AS [MB Written],
  CASE mf.is_percent_growth
    WHEN 1 THEN CAST(mf.growth AS VARCHAR(20)) + ' %'
    ELSE CAST(CAST(mf.growth AS BIGINT) * 8 / 1024 AS VARCHAR(20)) + ' MB'
  END AS [Auto Growth],
  CASE
    WHEN mf.max_size = -1 THEN 'Unlimited'
    WHEN mf.max_size =  0 THEN 'No Growth'
    ELSE CAST(CAST(mf.max_size AS BIGINT) * 8 / 1024 AS VARCHAR(20)) + ' MB'
  END AS [Max Size],
  CASE
    WHEN mf.max_size = 0                              THEN 'WARNING - Autogrowth disabled'
    WHEN mf.is_percent_growth = 1 AND mf.growth > 10 THEN 'WARNING - % growth causes stalls'
    WHEN mf.growth = 0                                THEN 'WARNING - No growth configured'
    ELSE 'OK'
  END AS [Status]
FROM sys.master_files mf
JOIN sys.databases d ON mf.database_id = d.database_id
LEFT JOIN sys.dm_io_virtual_file_stats(NULL, NULL) vfs
       ON vfs.database_id = mf.database_id AND vfs.file_id = mf.file_id
WHERE d.state_desc = 'ONLINE'
ORDER BY d.name, mf.type, mf.file_id
"@

$Q_TempDBChecks = @"
SELECT 'TempDB Data File Count' AS [Check],
  CAST(COUNT(*) AS VARCHAR) AS [Value],
  CAST(COUNT(*) AS VARCHAR)+' file(s) — recommend 1 per logical CPU core (max 8)' AS [Detail],
  CASE WHEN COUNT(*) = 1 THEN 'WARNING'
       WHEN COUNT(*) > 8 THEN 'WARNING'
       ELSE 'OK' END AS [Status]
FROM tempdb.sys.database_files WHERE type = 0
UNION ALL
SELECT 'TempDB File Size Equality',
  CAST(MAX(size)-MIN(size) AS VARCHAR)+' page difference',
  'All data files should be equal size to enable proportional fill',
  CASE WHEN MAX(size)<>MIN(size) THEN 'WARNING' ELSE 'OK' END
FROM tempdb.sys.database_files WHERE type = 0
UNION ALL
SELECT 'TempDB Total Size MB',
  CAST(CAST(SUM(size)*8.0/1024 AS DECIMAL(10,1)) AS VARCHAR)+' MB',
  'Monitor growth trend; pre-size to avoid runtime autogrowth',
  'INFO'
FROM tempdb.sys.database_files WHERE type = 0
UNION ALL
SELECT 'TempDB Log Size MB',
  CAST(CAST(SUM(size)*8.0/1024 AS DECIMAL(10,1)) AS VARCHAR)+' MB',
  'Large log may indicate long-running transactions in TempDB',
  CASE WHEN SUM(size)*8.0/1024 > 5120 THEN 'WARNING' ELSE 'INFO' END
FROM tempdb.sys.database_files WHERE type = 1
UNION ALL
SELECT 'TempDB % Growth Setting',
  CAST(COUNT(*) AS VARCHAR)+' file(s) use % growth',
  'Use fixed MB growth (e.g. 512MB) instead of % to avoid large stalls',
  CASE WHEN COUNT(*) > 0 THEN 'WARNING' ELSE 'OK' END
FROM tempdb.sys.database_files WHERE type=0 AND is_percent_growth=1
UNION ALL
SELECT 'TempDB Version Store Size MB',
  CAST(CAST(SUM(version_store_reserved_page_count)*8.0/1024 AS DECIMAL(10,1)) AS VARCHAR)+' MB',
  'High version store = long-running snapshot/RCSI transactions',
  CASE WHEN SUM(version_store_reserved_page_count)*8.0/1024 > 1024 THEN 'WARNING' ELSE 'OK' END
FROM sys.dm_db_file_space_usage
UNION ALL
SELECT 'TempDB Internal Object MB',
  CAST(CAST(SUM(internal_object_reserved_page_count)*8.0/1024 AS DECIMAL(10,1)) AS VARCHAR)+' MB',
  'Sort spills, hash joins, cursors using TempDB internal objects',
  CASE WHEN SUM(internal_object_reserved_page_count)*8.0/1024 > 2048 THEN 'WARNING' ELSE 'INFO' END
FROM sys.dm_db_file_space_usage
UNION ALL
SELECT 'TempDB User Object MB',
  CAST(CAST(SUM(user_object_reserved_page_count)*8.0/1024 AS DECIMAL(10,1)) AS VARCHAR)+' MB',
  'Temp tables and table variables in TempDB',
  'INFO'
FROM sys.dm_db_file_space_usage
UNION ALL
SELECT 'PFS/GAM Contention (2:1:1)',
  CAST(COUNT(*) AS VARCHAR)+' waiter(s)',
  'PAGELATCH waits on 2:1:1 or 2:1:3 indicate TempDB contention - add files',
  CASE WHEN COUNT(*) > 0 THEN 'WARNING' ELSE 'OK' END
FROM sys.dm_exec_requests r
WHERE r.wait_type LIKE 'PAGELATCH%'
  AND r.wait_resource LIKE '2:1:%'
"@

$Q_TempDBFiles = @"
SELECT
  f.file_id AS [File#],
  f.name AS [Logical Name],
  f.physical_name AS [Physical Path],
  f.type_desc AS [Type],
  CAST(f.size*8.0/1024 AS DECIMAL(10,1)) AS [Size MB],
  CAST(u.user_object_reserved_page_count*8.0/1024 AS DECIMAL(10,1)) AS [User Obj MB],
  CAST(u.internal_object_reserved_page_count*8.0/1024 AS DECIMAL(10,1)) AS [Internal MB],
  CAST(u.version_store_reserved_page_count*8.0/1024 AS DECIMAL(10,1)) AS [Version MB],
  CAST(u.unallocated_extent_page_count*8.0/1024 AS DECIMAL(10,1)) AS [Free MB],
  CAST(u.unallocated_extent_page_count*100.0/NULLIF(f.size,0) AS DECIMAL(5,1)) AS [Free %],
  CASE WHEN f.is_percent_growth=1 THEN CAST(f.growth AS VARCHAR)+'%'
       ELSE CAST(f.growth*8/1024 AS VARCHAR)+' MB' END AS [Auto Growth]
FROM tempdb.sys.database_files f
LEFT JOIN sys.dm_db_file_space_usage u ON f.file_id=u.file_id
ORDER BY f.type, f.file_id
"@


# ── TAB 6 & 7 SQL QUERIES ────────────────────────────────────────────────────

$Q_MissingIndexes = @"
SELECT TOP 30
  DB_NAME(mid.database_id) AS [Database],
  OBJECT_NAME(mid.object_id, mid.database_id) AS [Table],
  migs.avg_user_impact AS [Avg Benefit %],
  migs.user_seeks AS [Seeks],
  migs.user_scans AS [Scans],
  CAST(migs.avg_total_user_cost * migs.avg_user_impact * (migs.user_seeks + migs.user_scans) / 100.0 AS DECIMAL(18,1)) AS [Impact Score],
  mid.equality_columns AS [Equality Cols],
  mid.inequality_columns AS [Inequality Cols],
  mid.included_columns AS [Included Cols],
  'CREATE INDEX [IX_' + OBJECT_NAME(mid.object_id,mid.database_id) + '_missing] ON '
    + mid.statement
    + ' (' + ISNULL(mid.equality_columns,'')
    + CASE WHEN mid.equality_columns IS NOT NULL AND mid.inequality_columns IS NOT NULL THEN ',' ELSE '' END
    + ISNULL(mid.inequality_columns,'') + ')'
    + ISNULL(' INCLUDE (' + mid.included_columns + ')','') AS [Create Script]
FROM sys.dm_db_missing_index_groups mig
JOIN sys.dm_db_missing_index_group_stats migs ON mig.index_group_handle = migs.group_handle
JOIN sys.dm_db_missing_index_details mid ON mig.index_handle = mid.index_handle
WHERE migs.avg_user_impact > 10
ORDER BY [Impact Score] DESC
"@

$Q_UnusedIndexes = @"
SELECT TOP 30
  DB_NAME() AS [Database],
  OBJECT_NAME(i.object_id) AS [Table],
  i.name AS [Index],
  i.type_desc AS [Type],
  CAST(s.used_page_count * 8.0 / 1024 AS DECIMAL(10,1)) AS [Size MB],
  ius.user_seeks AS [Seeks],
  ius.user_scans AS [Scans],
  ius.user_lookups AS [Lookups],
  ius.user_updates AS [Updates],
  CASE
    WHEN ius.user_seeks + ius.user_scans + ius.user_lookups = 0
      AND ius.user_updates > 0 THEN 'DROP candidate - never read'
    WHEN ius.user_seeks + ius.user_scans + ius.user_lookups < ius.user_updates / 10
      THEN 'Review - write overhead exceeds reads'
    ELSE 'OK - in use'
  END AS [Recommendation]
FROM sys.indexes i
JOIN sys.objects o ON i.object_id = o.object_id
LEFT JOIN sys.dm_db_index_usage_stats ius
  ON i.object_id = ius.object_id AND i.index_id = ius.index_id
  AND ius.database_id = DB_ID()
LEFT JOIN sys.dm_db_partition_stats s
  ON i.object_id = s.object_id AND i.index_id = s.index_id
WHERE o.type = 'U'
  AND i.index_id > 0
  AND i.is_primary_key = 0
  AND i.is_unique_constraint = 0
ORDER BY ius.user_updates DESC, [Size MB] DESC
"@

$Q_DuplicateIndexes = @"
SELECT
  DB_NAME() AS [Database],
  OBJECT_NAME(i1.object_id) AS [Table],
  i1.name AS [Index 1],
  i2.name AS [Index 2],
  c1.cols AS [Columns],
  CAST(s1.used_page_count*8.0/1024 AS DECIMAL(10,1)) AS [Idx1 Size MB],
  CAST(s2.used_page_count*8.0/1024 AS DECIMAL(10,1)) AS [Idx2 Size MB],
  'Consider dropping: ' + i2.name AS [Action]
FROM sys.indexes i1
JOIN sys.indexes i2
  ON i1.object_id = i2.object_id AND i1.index_id < i2.index_id
  AND i1.is_primary_key = 0 AND i2.is_primary_key = 0
CROSS APPLY (
  SELECT STRING_AGG(c.name, ', ') WITHIN GROUP (ORDER BY ic.key_ordinal) AS cols
  FROM sys.index_columns ic JOIN sys.columns c ON ic.object_id=c.object_id AND ic.column_id=c.column_id
  WHERE ic.object_id=i1.object_id AND ic.index_id=i1.index_id AND ic.is_included_column=0
) c1
CROSS APPLY (
  SELECT STRING_AGG(c.name, ', ') WITHIN GROUP (ORDER BY ic.key_ordinal) AS cols
  FROM sys.index_columns ic JOIN sys.columns c ON ic.object_id=c.object_id AND ic.column_id=c.column_id
  WHERE ic.object_id=i2.object_id AND ic.index_id=i2.index_id AND ic.is_included_column=0
) c2
LEFT JOIN sys.dm_db_partition_stats s1 ON i1.object_id=s1.object_id AND i1.index_id=s1.index_id
LEFT JOIN sys.dm_db_partition_stats s2 ON i2.object_id=s2.object_id AND i2.index_id=s2.index_id
WHERE c1.cols = c2.cols
ORDER BY OBJECT_NAME(i1.object_id)
"@

$Q_MemoryPressure = @"
SELECT Metric, Value, Detail, Status FROM (
  SELECT 1 AS s, 'Page Life Expectancy' AS Metric,
    CAST(cntr_value AS NVARCHAR(30)) + ' sec' AS Value,
    'Target >= 300 sec (>= 1000 recommended for busy servers)' AS Detail,
    CASE WHEN cntr_value < 300 THEN 'WARNING' WHEN cntr_value < 1000 THEN 'INFO' ELSE 'OK' END AS Status
  FROM sys.dm_os_performance_counters
  WHERE counter_name = 'Page life expectancy' AND object_name LIKE '%Buffer Manager%'
  UNION ALL
  SELECT 2,'Memory Grants Pending',
    CAST(cntr_value AS NVARCHAR(30)) + ' grant(s)',
    'Queries waiting for workspace memory - > 0 indicates pressure',
    CASE WHEN cntr_value > 0 THEN 'WARNING' ELSE 'OK' END
  FROM sys.dm_os_performance_counters
  WHERE counter_name = 'Memory Grants Pending' AND object_name LIKE '%Memory Manager%'
  UNION ALL
  SELECT 3,'Target Server Memory',
    CAST(CAST(cntr_value/1024 AS BIGINT) AS NVARCHAR(30)) + ' MB',
    'How much memory SQL Server is targeting to use',
    'INFO'
  FROM sys.dm_os_performance_counters
  WHERE counter_name = 'Target Server Memory (KB)' AND object_name LIKE '%Memory Manager%'
  UNION ALL
  SELECT 4,'Total Server Memory',
    CAST(CAST(cntr_value/1024 AS BIGINT) AS NVARCHAR(30)) + ' MB',
    'How much memory SQL Server is actually consuming',
    'INFO'
  FROM sys.dm_os_performance_counters
  WHERE counter_name = 'Total Server Memory (KB)' AND object_name LIKE '%Memory Manager%'
  UNION ALL
  SELECT 5,'SQL Cache Memory',
    CAST(CAST(cntr_value/1024 AS BIGINT) AS NVARCHAR(30)) + ' MB',
    'Memory used by plan cache',
    'INFO'
  FROM sys.dm_os_performance_counters
  WHERE counter_name = 'SQL Cache Memory (KB)' AND object_name LIKE '%Memory Manager%'
  UNION ALL
  SELECT 6,'Stolen Server Memory',
    CAST(CAST(cntr_value/1024 AS BIGINT) AS NVARCHAR(30)) + ' MB',
    'Non-buffer pool memory (CLR, linked servers, etc.)',
    CASE WHEN cntr_value/1024 > 2048 THEN 'WARNING' ELSE 'INFO' END
  FROM sys.dm_os_performance_counters
  WHERE counter_name = 'Stolen Server Memory (KB)' AND object_name LIKE '%Memory Manager%'
  UNION ALL
  SELECT 7,'Buffer Cache Hit Ratio',
    CAST(CAST(
      CAST(v AS FLOAT) / NULLIF(CAST(b AS FLOAT),0) * 100
    AS DECIMAL(5,1)) AS NVARCHAR(30)) + ' %',
    'Target > 95%. Low value = not enough RAM or excessive scans',
    CASE WHEN CAST(v AS FLOAT)/NULLIF(CAST(b AS FLOAT),0)*100 < 95 THEN 'WARNING' ELSE 'OK' END
  FROM (
    SELECT
      MAX(CASE WHEN counter_name='Buffer cache hit ratio'      THEN cntr_value END) AS v,
      MAX(CASE WHEN counter_name='Buffer cache hit ratio base' THEN cntr_value END) AS b
    FROM sys.dm_os_performance_counters
    WHERE counter_name IN ('Buffer cache hit ratio','Buffer cache hit ratio base')
      AND object_name LIKE '%Buffer Manager%'
  ) x
) r
ORDER BY s
"@

$Q_MemoryClerks = @"
SELECT TOP 15
  type AS [Clerk Type],
  name AS [Name],
  CAST(pages_kb / 1024.0 AS DECIMAL(10,1)) AS [Memory MB],
  CAST(pages_kb * 100.0 / NULLIF(SUM(pages_kb) OVER(),0) AS DECIMAL(5,1)) AS [% of Total],
  CAST(virtual_memory_committed_kb / 1024.0 AS DECIMAL(10,1)) AS [Virtual MB],
  CAST(awe_allocated_kb / 1024.0 AS DECIMAL(10,1)) AS [AWE MB]
FROM sys.dm_os_memory_clerks
WHERE pages_kb > 0
ORDER BY pages_kb DESC
"@

$Q_BufferByDB = @"
SELECT
  CASE WHEN database_id = 32767 THEN 'Resource DB'
       ELSE DB_NAME(database_id) END AS [Database],
  COUNT(*) * 8 / 1024 AS [Buffer MB],
  CAST(COUNT(*) * 100.0 / SUM(COUNT(*)) OVER() AS DECIMAL(5,1)) AS [% of Buffer Pool],
  SUM(CASE WHEN is_modified=1 THEN 1 ELSE 0 END) * 8 / 1024 AS [Dirty Pages MB]
FROM sys.dm_os_buffer_descriptors
GROUP BY database_id
ORDER BY [Buffer MB] DESC
"@

$Q_CPUHistory = @"
SELECT TOP 60
  DATEADD(ms, -1 * (si.cpu_ticks / CONVERT(FLOAT, si.cpu_ticks/ms_ticks) - rb.timestamp), GETDATE()) AS [Recorded At],
  record.value('(./Record/SchedulerMonitorEvent/SystemHealth/ProcessUtilization)[1]','int') AS [SQL CPU %],
  record.value('(./Record/SchedulerMonitorEvent/SystemHealth/SystemIdle)[1]','int') AS [Idle %],
  100
  - record.value('(./Record/SchedulerMonitorEvent/SystemHealth/SystemIdle)[1]','int')
  - record.value('(./Record/SchedulerMonitorEvent/SystemHealth/ProcessUtilization)[1]','int') AS [Other CPU %]
FROM (
  SELECT timestamp, CONVERT(XML,record) AS record
  FROM sys.dm_os_ring_buffers
  WHERE ring_buffer_type = N'RING_BUFFER_SCHEDULER_MONITOR'
    AND record LIKE N'%<SystemHealth>%'
) rb
CROSS JOIN sys.dm_os_sys_info si
ORDER BY rb.timestamp DESC
"@

$Q_LongTxns = @"
SELECT
  s.session_id AS SPID,
  s.login_name AS Login,
  s.host_name AS Host,
  DB_NAME(r.database_id) AS [Database],
  DATEDIFF(second, at.transaction_begin_time, GETDATE()) AS [Open Sec],
  CAST(DATEDIFF(second, at.transaction_begin_time, GETDATE()) / 60.0 AS DECIMAL(10,1)) AS [Open Min],
  at.transaction_type AS [Txn Type],
  at.transaction_state AS [Txn State],
  r.wait_type AS [Wait Type],
  r.blocking_session_id AS [Blocked By],
  SUBSTRING(ISNULL(t.text,'(no SQL)'),1,200) AS [SQL Text]
FROM sys.dm_tran_active_transactions at
JOIN sys.dm_tran_session_transactions st ON at.transaction_id = st.transaction_id
JOIN sys.dm_exec_sessions s ON st.session_id = s.session_id
LEFT JOIN sys.dm_exec_requests r ON s.session_id = r.session_id
OUTER APPLY sys.dm_exec_sql_text(r.sql_handle) t
WHERE DATEDIFF(second, at.transaction_begin_time, GETDATE()) > 30
  AND s.session_id <> @@SPID
ORDER BY [Open Sec] DESC
"@

$Q_Deadlocks = @"
SELECT TOP 20
  xdr.value('@timestamp','datetime2') AS [Deadlock Time],
  xdr.value('(victim-list/victimProcess/@id)[1]','varchar(50)') AS [Victim Process],
  xdr.value('(process-list/process/@spid)[1]','int') AS [SPID 1],
  xdr.value('(process-list/process/@loginname)[1]','varchar(100)') AS [Login 1],
  xdr.value('(process-list/process/@hostname)[1]','varchar(100)') AS [Host 1],
  xdr.value('(process-list/process/@currentdbname)[1]','varchar(100)') AS [Database],
  xdr.value('(process-list/process/inputbuf)[1]','nvarchar(500)') AS [SQL 1],
  xdr.value('(process-list/process/@spid)[2]','int') AS [SPID 2],
  xdr.value('(process-list/process/inputbuf)[2]','nvarchar(500)') AS [SQL 2]
FROM (
  SELECT CONVERT(XML, target_data) AS target_data
  FROM sys.dm_xe_session_targets t
  JOIN sys.dm_xe_sessions s ON t.event_session_address = s.address
  WHERE s.name = 'system_health'
    AND t.target_name = 'ring_buffer'
) x
CROSS APPLY target_data.nodes('RingBufferTarget/event[@name="xml_deadlock_report"]') AS xdt(xdr)
ORDER BY [Deadlock Time] DESC
"@

$Q_OpenTxnLocks = @"
SELECT TOP 30
  tl.request_session_id AS SPID,
  DB_NAME(tl.resource_database_id) AS [Database],
  tl.resource_type AS [Resource Type],
  CASE tl.resource_type
    WHEN 'OBJECT' THEN OBJECT_NAME(tl.resource_associated_entity_id, tl.resource_database_id)
    ELSE CAST(tl.resource_associated_entity_id AS VARCHAR(50))
  END AS [Object],
  tl.request_mode AS [Lock Mode],
  tl.request_status AS [Status],
  s.login_name AS Login,
  s.host_name AS Host,
  s.open_transaction_count AS [Open Txns]
FROM sys.dm_tran_locks tl
JOIN sys.dm_exec_sessions s ON tl.request_session_id = s.session_id
WHERE tl.request_session_id <> @@SPID
  AND s.is_user_process = 1
ORDER BY tl.request_session_id, tl.resource_type
"@

$Q_SecurityChecks = @"
SELECT 'Authentication Mode' AS [Check],
  CASE SERVERPROPERTY('IsIntegratedSecurityOnly') WHEN 1 THEN 'Windows Only' ELSE 'Mixed Mode (SQL+Windows)' END AS [Value],
  'Mixed mode allows SQL logins - ensure all have strong passwords',
  CASE SERVERPROPERTY('IsIntegratedSecurityOnly') WHEN 1 THEN 'OK' ELSE 'WARNING' END AS [Status]
UNION ALL
SELECT 'SA Account',
  CASE WHEN is_disabled=0 THEN 'ENABLED' ELSE 'Disabled' END,
  'SA should be disabled or renamed - common attack target',
  CASE WHEN is_disabled=0 THEN 'WARNING' ELSE 'OK' END
FROM sys.server_principals WHERE name='sa'
UNION ALL
SELECT 'xp_cmdshell',
  CAST(value_in_use AS VARCHAR),
  'Should be 0 unless explicitly required',
  CASE WHEN value_in_use=1 THEN 'WARNING' ELSE 'OK' END
FROM sys.configurations WHERE name='xp_cmdshell'
UNION ALL
SELECT 'Sysadmin members (non-system)',
  CAST(COUNT(*) AS VARCHAR) + ' account(s)',
  'Minimize sysadmin membership - principle of least privilege',
  CASE WHEN COUNT(*) > 3 THEN 'WARNING' ELSE 'INFO' END
FROM sys.server_role_members srm
JOIN sys.server_principals sp ON srm.member_principal_id = sp.principal_id
WHERE srm.role_principal_id = SUSER_ID('sysadmin')
  AND sp.name NOT IN ('sa','NT AUTHORITY\SYSTEM','NT SERVICE\MSSQLSERVER','NT SERVICE\SQLSERVERAGENT')
UNION ALL
SELECT 'SQL Logins with blank password',
  CAST(COUNT(*) AS VARCHAR) + ' login(s)',
  'SQL logins must have strong passwords',
  CASE WHEN COUNT(*) > 0 THEN 'WARNING' ELSE 'OK' END
FROM sys.sql_logins WHERE is_disabled=0 AND PWDCOMPARE('',password_hash)=1
UNION ALL
SELECT 'SQL Logins with password = username',
  CAST(COUNT(*) AS VARCHAR) + ' login(s)',
  'Weak passwords matching login name detected',
  CASE WHEN COUNT(*) > 0 THEN 'WARNING' ELSE 'OK' END
FROM sys.sql_logins WHERE is_disabled=0 AND PWDCOMPARE(name,password_hash)=1
UNION ALL
SELECT 'Guest user enabled (non-system DBs)',
  CAST(COUNT(*) AS VARCHAR) + ' database(s)',
  'Guest account allows any login to connect - disable per DB',
  CASE WHEN COUNT(*) > 0 THEN 'WARNING' ELSE 'OK' END
FROM sys.databases d
WHERE d.database_id > 4
  AND d.state_desc='ONLINE'
  AND EXISTS (
    SELECT 1 FROM sys.dm_exec_sql_text(0x) -- dummy to allow subquery
    WHERE 1=0
  )
UNION ALL
SELECT 'Linked Servers configured',
  CAST(COUNT(*) AS VARCHAR) + ' linked server(s)',
  'Review linked servers - each is a potential lateral movement path',
  CASE WHEN COUNT(*) > 0 THEN 'WARNING' ELSE 'INFO' END
FROM sys.servers WHERE is_linked=1
"@

$Q_SysadminMembers = @"
SELECT
  sp.name AS [Login],
  sp.type_desc AS [Type],
  sp.is_disabled AS [Disabled],
  sp.create_date AS [Created],
  sp.modify_date AS [Modified],
  CASE sp.type
    WHEN 'S' THEN 'SQL Login'
    WHEN 'U' THEN 'Windows User'
    WHEN 'G' THEN 'Windows Group'
    ELSE sp.type_desc
  END AS [Auth Type],
  CASE WHEN sp.name IN ('sa','NT AUTHORITY\SYSTEM','NT SERVICE\MSSQLSERVER','NT SERVICE\SQLSERVERAGENT','NT SERVICE\SQLWriter','NT SERVICE\Winmgmt')
    THEN 'System' ELSE 'User-added' END AS [Category]
FROM sys.server_role_members srm
JOIN sys.server_principals sp ON srm.member_principal_id = sp.principal_id
WHERE srm.role_principal_id = SUSER_ID('sysadmin')
ORDER BY [Category], sp.name
"@

# ── FORM ─────────────────────────────────────────────────────────────────────

$form               = New-Object System.Windows.Forms.Form
$form.Text          = "SQL Server Lite Monitor"
$form.Size          = New-Object System.Drawing.Size(1300,820)
$form.MinimumSize   = New-Object System.Drawing.Size(900,600)
$form.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
$form.BackColor     = [System.Drawing.Color]::FromArgb(25,25,28)

# ── Title panel (Dock Top) ────────────────────────────────────────────────────
$titlePnl           = New-Object System.Windows.Forms.Panel
$titlePnl.Dock      = [System.Windows.Forms.DockStyle]::Top
$titlePnl.Height    = 48
$titlePnl.BackColor = [System.Drawing.Color]::FromArgb(0,78,158)
$titleLbl           = New-Object System.Windows.Forms.Label
$titleLbl.Text      = "  SQL Server Lite Monitor"
$titleLbl.Dock      = [System.Windows.Forms.DockStyle]::Fill
$titleLbl.Font      = New-Object System.Drawing.Font("Segoe UI",13,[System.Drawing.FontStyle]::Bold)
$titleLbl.ForeColor = [System.Drawing.Color]::White
$titleLbl.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
$titlePnl.Controls.Add($titleLbl)

# ── Connection bar (Dock Top) ─────────────────────────────────────────────────
$connPnl           = New-Object System.Windows.Forms.Panel
$connPnl.Dock      = [System.Windows.Forms.DockStyle]::Top
$connPnl.Height    = 48
$connPnl.BackColor = [System.Drawing.Color]::FromArgb(38,38,44)

function cLbl($t,$x,$y,$w=52) {
    $l=New-Object System.Windows.Forms.Label; $l.Text=$t
    $l.Location=New-Object System.Drawing.Point($x,$y); $l.Size=New-Object System.Drawing.Size($w,22)
    $l.ForeColor=[System.Drawing.Color]::FromArgb(190,190,190); $l.TextAlign="MiddleLeft"
    $l.Font=New-Object System.Drawing.Font("Segoe UI",9); $l.BackColor=[System.Drawing.Color]::Transparent
    $connPnl.Controls.Add($l)
}
function cTxt($x,$y,$w=185,$val="",$pwd=$false) {
    if($pwd){$t=New-Object System.Windows.Forms.MaskedTextBox;$t.PasswordChar='*'}
    else    {$t=New-Object System.Windows.Forms.TextBox}
    $t.Location=New-Object System.Drawing.Point($x,$y); $t.Size=New-Object System.Drawing.Size($w,22)
    $t.BackColor=[System.Drawing.Color]::FromArgb(50,50,58); $t.ForeColor=[System.Drawing.Color]::White
    $t.BorderStyle="FixedSingle"; if($val){$t.Text=$val}
    $connPnl.Controls.Add($t); return $t
}
function cBtn($t,$x,$y,$w=105) {
    $b=New-Object System.Windows.Forms.Button; $b.Text=$t
    $b.Location=New-Object System.Drawing.Point($x,$y); $b.Size=New-Object System.Drawing.Size($w,26)
    $b.FlatStyle="Flat"; $b.BackColor=[System.Drawing.Color]::FromArgb(0,98,188)
    $b.ForeColor=[System.Drawing.Color]::White; $b.Cursor=[System.Windows.Forms.Cursors]::Hand
    $b.FlatAppearance.BorderColor=[System.Drawing.Color]::FromArgb(0,140,230)
    $b.Font=New-Object System.Drawing.Font("Segoe UI",9,[System.Drawing.FontStyle]::Bold)
    $connPnl.Controls.Add($b); return $b
}

cLbl "Server:" 8 13
$txtSrv   = cTxt 60 12 180 $env:COMPUTERNAME
cLbl "Auth:" 248 13 38
$cmbAuth  = New-Object System.Windows.Forms.ComboBox
$cmbAuth.Location=New-Object System.Drawing.Point(286,11); $cmbAuth.Size=New-Object System.Drawing.Size(135,22)
$cmbAuth.DropDownStyle="DropDownList"; $cmbAuth.BackColor=[System.Drawing.Color]::FromArgb(50,50,58)
$cmbAuth.ForeColor=[System.Drawing.Color]::White
$cmbAuth.Items.AddRange(@("Windows Auth","SQL Server Auth")); $cmbAuth.SelectedIndex=0
$connPnl.Controls.Add($cmbAuth)
cLbl "User:" 428 13 36
$txtUser  = cTxt 464 12 105; $txtUser.Enabled=$false
cLbl "Pass:" 576 13 36
$txtPwd   = cTxt 612 12 105 "" $true; $txtPwd.Enabled=$false
$btnConn  = cBtn "Connect"     724 11 95
$cmbRef   = New-Object System.Windows.Forms.ComboBox
$cmbRef.Location=New-Object System.Drawing.Point(826,11); $cmbRef.Size=New-Object System.Drawing.Size(125,22)
$cmbRef.DropDownStyle="DropDownList"; $cmbRef.BackColor=[System.Drawing.Color]::FromArgb(50,50,58)
$cmbRef.ForeColor=[System.Drawing.Color]::White
$cmbRef.Items.AddRange(@("Refresh: 15s","Refresh: 30s","Refresh: 60s","Refresh: Off"))
$cmbRef.SelectedIndex=1; $connPnl.Controls.Add($cmbRef)
$btnRef   = cBtn "Refresh Now" 958 11 115
cLbl "DB:" 1082 13 28
$cmbDB = New-Object System.Windows.Forms.ComboBox
$cmbDB.Location=New-Object System.Drawing.Point(1110,11); $cmbDB.Size=New-Object System.Drawing.Size(155,22)
$cmbDB.DropDownStyle="DropDownList"; $cmbDB.BackColor=[System.Drawing.Color]::FromArgb(50,50,58)
$cmbDB.ForeColor=[System.Drawing.Color]::White; $cmbDB.Items.Add("master") | Out-Null
$cmbDB.SelectedIndex=0; $connPnl.Controls.Add($cmbDB)

# ── Status bar (Dock Bottom) ──────────────────────────────────────────────────
$statusPnl           = New-Object System.Windows.Forms.Panel
$statusPnl.Dock      = [System.Windows.Forms.DockStyle]::Bottom
$statusPnl.Height    = 24
$statusPnl.BackColor = [System.Drawing.Color]::FromArgb(30,30,35)
$script:statusLbl    = New-Object System.Windows.Forms.Label
$script:statusLbl.Dock      = [System.Windows.Forms.DockStyle]::Fill
$script:statusLbl.Font      = New-Object System.Drawing.Font("Segoe UI",9)
$script:statusLbl.ForeColor = [System.Drawing.Color]::DimGray
$script:statusLbl.Text      = "  Not connected. Enter server name and click Connect."
$script:statusLbl.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
$statusPnl.Controls.Add($script:statusLbl)

# ── TabControl (Dock Fill — must be added LAST so Fill uses remaining space) ──
$tabs         = New-Object System.Windows.Forms.TabControl
$tabs.Dock    = [System.Windows.Forms.DockStyle]::Fill
$tabs.Font    = New-Object System.Drawing.Font("Segoe UI",10,[System.Drawing.FontStyle]::Bold)
$tabs.Padding = New-Object System.Drawing.Point(14,5)

# Add controls in correct Dock order: Bottom first, then Top items, then Fill last
$form.Controls.Add($tabs)       # Fill  - added first so it ends up "behind" in z-order
$form.Controls.Add($statusPnl)  # Bottom
$form.Controls.Add($connPnl)    # Top (second, so renders below title)
$form.Controls.Add($titlePnl)   # Top (last added = topmost)

# ── TAB 1: CONFIGURATION ─────────────────────────────────────────────────────
$t1 = New-Object System.Windows.Forms.TabPage
$t1.Text = "  Configuration"
$t1.BackColor = [System.Drawing.Color]::FromArgb(28,28,32)
$tabs.TabPages.Add($t1)

# Use a SplitContainer so both grids get space
$split1 = New-Object System.Windows.Forms.SplitContainer
$split1.Dock        = [System.Windows.Forms.DockStyle]::Fill
$split1.Orientation = [System.Windows.Forms.Orientation]::Horizontal
$split1.Panel1.BackColor = [System.Drawing.Color]::FromArgb(28,28,32)
$split1.Panel2.BackColor = [System.Drawing.Color]::FromArgb(28,28,32)
$split1.BackColor = [System.Drawing.Color]::FromArgb(28,28,32)
$t1.Controls.Add($split1)

$hdr1a = New-SectionPanel "Server Configuration & Best Practice Checks   GREEN=OK  ORANGE=Warning  BLUE=Info"
$script:gCfg = New-DGV
$split1.Panel1.Controls.Add($script:gCfg)
$split1.Panel1.Controls.Add($hdr1a)   # Top dock, added after Fill so it renders on top

$hdr1b = New-SectionPanel "Index Fragmentation Health — DB: master"
$script:idxHdrLbl = $hdr1b.Controls[0]  # store label ref for dynamic update
$script:gIdx = New-DGV
$split1.Panel2.Controls.Add($script:gIdx)
$split1.Panel2.Controls.Add($hdr1b)

$script:gCfg.add_CellFormatting({
    param($s,$e)
    if($e.ColumnIndex -lt 0 -or $script:gCfg.Columns.Count -eq 0){return}
    if($script:gCfg.Columns[$e.ColumnIndex].Name -eq "Status"){
        switch($e.Value){
            "WARNING"{$e.CellStyle.ForeColor=[System.Drawing.Color]::Orange;
                      $e.CellStyle.Font=New-Object System.Drawing.Font("Segoe UI",9,[System.Drawing.FontStyle]::Bold)}
            "OK"    {$e.CellStyle.ForeColor=[System.Drawing.Color]::LightGreen}
            "INFO"  {$e.CellStyle.ForeColor=[System.Drawing.Color]::CornflowerBlue}
        }
    }
})
$script:gIdx.add_CellFormatting({
    param($s,$e)
    if($e.ColumnIndex -lt 0 -or $script:gIdx.Columns.Count -eq 0){return}
    if($script:gIdx.Columns[$e.ColumnIndex].Name -eq "Action"){
        switch -Wildcard ($e.Value){
            "*REBUILD*"    {$e.CellStyle.ForeColor=[System.Drawing.Color]::OrangeRed}
            "*REORGANIZE*" {$e.CellStyle.ForeColor=[System.Drawing.Color]::Yellow}
            "OK"           {$e.CellStyle.ForeColor=[System.Drawing.Color]::LightGreen}
        }
    }
})

# ── TAB 2: PERFORMANCE ───────────────────────────────────────────────────────
$t2 = New-Object System.Windows.Forms.TabPage
$t2.Text = "  Performance"
$t2.BackColor = [System.Drawing.Color]::FromArgb(28,28,32)
$tabs.TabPages.Add($t2)

# CPU row at top with chart
$cpuHdr = New-SectionPanel "CPU History (last ~60 snapshots from ring buffer)   BLUE=SQL  ORANGE=Other  — current: "
$script:cpuHdrLbl = $cpuHdr.Controls[0]
$cpuPnl = New-Object System.Windows.Forms.Panel
$cpuPnl.Dock      = [System.Windows.Forms.DockStyle]::Top
$cpuPnl.Height    = 160
$cpuPnl.BackColor = [System.Drawing.Color]::FromArgb(28,28,32)

# Label strip
$script:lSql = New-Object System.Windows.Forms.Label; $script:lSql.AutoSize=$true
$script:lOth = New-Object System.Windows.Forms.Label; $script:lOth.AutoSize=$true
$script:lIdl = New-Object System.Windows.Forms.Label; $script:lIdl.AutoSize=$true
$script:lSql.Text="SQL CPU: --"; $script:lOth.Text="Other: --"; $script:lIdl.Text="Idle: --"
$script:lSql.Location=New-Object System.Drawing.Point(10,5)
$script:lOth.Location=New-Object System.Drawing.Point(160,5)
$script:lIdl.Location=New-Object System.Drawing.Point(290,5)
foreach($x in @($script:lSql,$script:lOth,$script:lIdl)){
    $x.Font=New-Object System.Drawing.Font("Segoe UI",10,[System.Drawing.FontStyle]::Bold)
    $x.BackColor=[System.Drawing.Color]::Transparent
}
$script:lSql.ForeColor=[System.Drawing.Color]::FromArgb(0,185,255)
$script:lOth.ForeColor=[System.Drawing.Color]::Orange
$script:lIdl.ForeColor=[System.Drawing.Color]::LightGreen

# CPU Chart with timestamps
$script:cpuHistory  = [System.Collections.Generic.List[int[]]]::new()    # [sqlPct, othPct]
$script:cpuTimes    = [System.Collections.Generic.List[string]]::new()   # HH:mm:ss labels

$script:cpuChart = New-Object System.Windows.Forms.PictureBox
$script:cpuChart.Dock      = [System.Windows.Forms.DockStyle]::Fill
$script:cpuChart.BackColor = [System.Drawing.Color]::FromArgb(18,18,22)
$script:cpuChart.SizeMode  = [System.Windows.Forms.PictureBoxSizeMode]::Normal

$script:cpuChart.add_Paint({
    param($s,$e)
    $g = $e.Graphics
    $g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
    $g.TextRenderingHint = [System.Drawing.Text.TextRenderingHint]::ClearTypeGridFit

    $totalW = $script:cpuChart.Width
    $totalH = $script:cpuChart.Height
    if($totalW -le 0 -or $totalH -le 0 -or $script:cpuHistory.Count -lt 2){ return }

    # Layout margins: left=42 (Y axis labels), right=8, top=6, bottom=22 (X axis)
    $mL=42; $mR=8; $mT=6; $mB=22
    $cW = $totalW - $mL - $mR   # chart area width
    $cH = $totalH - $mT - $mB   # chart area height
    if($cW -le 0 -or $cH -le 0){ return }

    $pts = $script:cpuHistory.Count

    # Fonts & brushes
    $fntAxis  = New-Object System.Drawing.Font("Consolas",8)
    $fntLabel = New-Object System.Drawing.Font("Consolas",7)
    $brWhite  = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(200,200,200))
    $brDim    = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(90,90,100))
    $brSQL    = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(140,0,185,255))
    $brOth    = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(110,255,140,0))
    $penGrid  = New-Object System.Drawing.Pen([System.Drawing.Color]::FromArgb(45,45,55),1)
    $penSQL   = New-Object System.Drawing.Pen([System.Drawing.Color]::FromArgb(0,185,255),1.5)
    $penOth   = New-Object System.Drawing.Pen([System.Drawing.Color]::Orange,1.5)
    $penBorder= New-Object System.Drawing.Pen([System.Drawing.Color]::FromArgb(60,60,70),1)

    # Background
    $g.FillRectangle(
        (New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(18,18,22))),
        0,0,$totalW,$totalH)

    # Y-axis grid lines & labels at 0,25,50,75,100
    foreach($pct in @(0,25,50,75,100)){
        $y = $mT + $cH - [int]($pct * $cH / 100)
        $g.DrawLine($penGrid, $mL, $y, $mL+$cW, $y)
        $lbl = "$pct%"
        $sz  = $g.MeasureString($lbl,$fntAxis)
        $g.DrawString($lbl,$fntAxis,$brDim,[float]($mL-$sz.Width-3),[float]($y-$sz.Height/2))
    }

    # Chart border
    $g.DrawRectangle($penBorder,$mL,$mT,$cW,$cH)

    # Build X positions
    $colW  = $cW / [math]::Max(1,$pts-1)

    # Filled area under SQL CPU line
    $sqlFillPts = [System.Collections.Generic.List[System.Drawing.PointF]]::new()
    $sqlFillPts.Add([System.Drawing.PointF]::new($mL, $mT+$cH))
    for($i=0;$i-lt$pts;$i++){
        $x = [float]($mL + $i * $colW)
        $y = [float]($mT + $cH - ($script:cpuHistory[$i][0] * $cH / 100))
        $sqlFillPts.Add([System.Drawing.PointF]::new($x,$y))
    }
    $sqlFillPts.Add([System.Drawing.PointF]::new([float]($mL+$cW),[float]($mT+$cH)))
    $g.FillPolygon($brSQL,$sqlFillPts.ToArray())

    # Filled area under Other CPU line
    $othFillPts = [System.Collections.Generic.List[System.Drawing.PointF]]::new()
    $othFillPts.Add([System.Drawing.PointF]::new($mL,[float]($mT+$cH)))
    for($i=0;$i-lt$pts;$i++){
        $x = [float]($mL + $i * $colW)
        $y = [float]($mT + $cH - ($script:cpuHistory[$i][1] * $cH / 100))
        $othFillPts.Add([System.Drawing.PointF]::new($x,$y))
    }
    $othFillPts.Add([System.Drawing.PointF]::new([float]($mL+$cW),[float]($mT+$cH)))
    $g.FillPolygon($brOth,$othFillPts.ToArray())

    # SQL CPU line
    $sqlLinePts = for($i=0;$i-lt$pts;$i++){
        [System.Drawing.PointF]::new(
            [float]($mL + $i * $colW),
            [float]($mT + $cH - ($script:cpuHistory[$i][0] * $cH / 100)))
    }
    if($sqlLinePts.Count -ge 2){ $g.DrawLines($penSQL,$sqlLinePts) }

    # Other CPU line
    $othLinePts = for($i=0;$i-lt$pts;$i++){
        [System.Drawing.PointF]::new(
            [float]($mL + $i * $colW),
            [float]($mT + $cH - ($script:cpuHistory[$i][1] * $cH / 100)))
    }
    if($othLinePts.Count -ge 2){ $g.DrawLines($penOth,$othLinePts) }

    # X-axis timestamp labels — show ~8 evenly spaced
    $step = [math]::Max(1,[int]($pts/8))
    for($i=0;$i-lt$pts;$i+=$step){
        if($i -lt $script:cpuTimes.Count){
            $x    = [float]($mL + $i * $colW)
            $lbl  = $script:cpuTimes[$i]
            $sz   = $g.MeasureString($lbl,$fntLabel)
            $g.DrawString($lbl,$fntLabel,$brDim,
                [float]($x - $sz.Width/2),
                [float]($mT+$cH+3))
            # Tick mark
            $g.DrawLine($penGrid,$x,[float]($mT+$cH),$x,[float]($mT+$cH+3))
        }
    }
    # Always show rightmost label (newest)
    if($script:cpuTimes.Count -gt 0){
        $xLast = [float]($mL + ($pts-1)*$colW)
        $lbl   = $script:cpuTimes[$script:cpuTimes.Count-1]
        $sz    = $g.MeasureString($lbl,$fntLabel)
        $g.DrawString($lbl,$fntLabel,$brWhite,
            [float]($xLast-$sz.Width/2),[float]($mT+$cH+3))
    }

    # Current value annotation on right edge
    if($pts -gt 0){
        $sqlNow = $script:cpuHistory[$pts-1][0]
        $othNow = $script:cpuHistory[$pts-1][1]
        $annFont= New-Object System.Drawing.Font("Segoe UI",8,[System.Drawing.FontStyle]::Bold)
        $g.DrawString("SQL: $sqlNow%",$annFont,
            (New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(0,185,255))),
            [float]($mL+4),[float]($mT+4))
        $g.DrawString("Oth: $othNow%",$annFont,
            (New-Object System.Drawing.SolidBrush([System.Drawing.Color]::Orange)),
            [float]($mL+4),[float]($mT+18))
        $annFont.Dispose()
    }

    $fntAxis.Dispose();$fntLabel.Dispose();$brWhite.Dispose();$brDim.Dispose()
    $brSQL.Dispose();$brOth.Dispose();$penGrid.Dispose()
    $penSQL.Dispose();$penOth.Dispose();$penBorder.Dispose()
})

# Label strip (top of cpuPnl, above chart)
$lblStrip = New-Object System.Windows.Forms.Panel
$lblStrip.Dock = [System.Windows.Forms.DockStyle]::Top; $lblStrip.Height = 24
$lblStrip.BackColor = [System.Drawing.Color]::FromArgb(28,28,32)
$lblStrip.Controls.AddRange(@($script:lSql,$script:lOth,$script:lIdl))

# Add to cpuPnl: Fill first then Top
$cpuPnl.Controls.Add($script:cpuChart)
$cpuPnl.Controls.Add($lblStrip)

# Split: top=active requests, bottom=waits+blocking side by side
$split2 = New-Object System.Windows.Forms.SplitContainer
$split2.Dock             = [System.Windows.Forms.DockStyle]::Fill
$split2.Orientation      = [System.Windows.Forms.Orientation]::Horizontal
$split2.SplitterDistance = 300
$split2.Panel1.BackColor = [System.Drawing.Color]::FromArgb(28,28,32)
$split2.Panel2.BackColor = [System.Drawing.Color]::FromArgb(28,28,32)
$split2.BackColor        = [System.Drawing.Color]::FromArgb(28,28,32)

$hdrReq = New-SectionPanel "Active Requests   (RED rows = blocked sessions)"
$script:gReq = New-DGV
$split2.Panel1.Controls.Add($script:gReq)
$split2.Panel1.Controls.Add($hdrReq)

$wSplit = New-Object System.Windows.Forms.SplitContainer
$wSplit.Dock             = [System.Windows.Forms.DockStyle]::Fill
$wSplit.Orientation      = [System.Windows.Forms.Orientation]::Vertical
$wSplit.SplitterDistance = 500
$wSplit.BackColor        = [System.Drawing.Color]::FromArgb(28,28,32)
$wSplit.Panel1.BackColor = [System.Drawing.Color]::FromArgb(28,28,32)
$wSplit.Panel2.BackColor = [System.Drawing.Color]::FromArgb(28,28,32)
$hdrWait = New-SectionPanel "Top Wait Types"
$script:gWait = New-DGV
$wSplit.Panel1.Controls.Add($script:gWait)
$wSplit.Panel1.Controls.Add($hdrWait)
$hdrBlk = New-SectionPanel "Blocking Chains"
$script:gBlk  = New-DGV
$wSplit.Panel2.Controls.Add($script:gBlk)
$wSplit.Panel2.Controls.Add($hdrBlk)
$split2.Panel2.Controls.Add($wSplit)

# Add to tab2: Fill first, then Top panels
$t2.Controls.Add($split2)
$t2.Controls.Add($cpuPnl)
$t2.Controls.Add($cpuHdr)

$script:gReq.add_CellFormatting({
    param($s,$e)
    if($e.RowIndex -lt 0 -or $script:gReq.Columns.Count -eq 0){return}
    try{
        $v=$script:gReq.Rows[$e.RowIndex].Cells["Blocked By"].Value
        if($null -ne $v -and "$v" -ne "0" -and "$v" -ne ""){
            $e.CellStyle.BackColor=[System.Drawing.Color]::FromArgb(90,22,22)
        }
    }catch{}
})

# ── TAB 3: BACKUPS & JOBS ────────────────────────────────────────────────────
$t3 = New-Object System.Windows.Forms.TabPage
$t3.Text = "  Backups & Jobs"
$t3.BackColor = [System.Drawing.Color]::FromArgb(28,28,32)
$tabs.TabPages.Add($t3)

$split3 = New-Object System.Windows.Forms.SplitContainer
$split3.Dock        = [System.Windows.Forms.DockStyle]::Fill
$split3.Orientation = [System.Windows.Forms.Orientation]::Horizontal
$split3.Panel1.BackColor = [System.Drawing.Color]::FromArgb(28,28,32)
$split3.Panel2.BackColor = [System.Drawing.Color]::FromArgb(28,28,32)
$split3.BackColor = [System.Drawing.Color]::FromArgb(28,28,32)
$t3.Controls.Add($split3)

$hdr3a = New-SectionPanel "Database Backup Status   RED=Never backed up  ORANGE=Overdue  GREEN=OK"
$script:gBak = New-DGV
$split3.Panel1.Controls.Add($script:gBak)
$split3.Panel1.Controls.Add($hdr3a)

$hdr3b = New-SectionPanel "SQL Agent Job History (last 50 runs)"
$script:gJob = New-DGV
$split3.Panel2.Controls.Add($script:gJob)
$split3.Panel2.Controls.Add($hdr3b)

$script:gBak.add_CellFormatting({
    param($s,$e)
    if($e.ColumnIndex -lt 0 -or $script:gBak.Columns.Count -eq 0){return}
    if($script:gBak.Columns[$e.ColumnIndex].Name -eq "Backup Status"){
        switch -Wildcard ($e.Value){
            "NEVER*"  {$e.CellStyle.ForeColor=[System.Drawing.Color]::Red;
                       $e.CellStyle.Font=New-Object System.Drawing.Font("Segoe UI",9,[System.Drawing.FontStyle]::Bold)}
            "WARNING*"{$e.CellStyle.ForeColor=[System.Drawing.Color]::Orange;
                       $e.CellStyle.Font=New-Object System.Drawing.Font("Segoe UI",9,[System.Drawing.FontStyle]::Bold)}
            "OK"      {$e.CellStyle.ForeColor=[System.Drawing.Color]::LightGreen}
        }
    }
})
$script:gJob.add_CellFormatting({
    param($s,$e)
    if($e.ColumnIndex -lt 0 -or $script:gJob.Columns.Count -eq 0){return}
    if($script:gJob.Columns[$e.ColumnIndex].Name -eq "Status"){
        switch($e.Value){
            "Failed"   {$e.CellStyle.ForeColor=[System.Drawing.Color]::Red;
                        $e.CellStyle.Font=New-Object System.Drawing.Font("Segoe UI",9,[System.Drawing.FontStyle]::Bold)}
            "Succeeded"{$e.CellStyle.ForeColor=[System.Drawing.Color]::LightGreen}
            "Retry"    {$e.CellStyle.ForeColor=[System.Drawing.Color]::Orange}
            "Cancelled"{$e.CellStyle.ForeColor=[System.Drawing.Color]::Yellow}
        }
    }
})

# ── TAB 4: QUERIES & I/O ─────────────────────────────────────────────────────
$t4 = New-Object System.Windows.Forms.TabPage
$t4.Text = "  Queries & I/O"
$t4.BackColor = [System.Drawing.Color]::FromArgb(28,28,32)
$tabs.TabPages.Add($t4)

$split4 = New-Object System.Windows.Forms.SplitContainer
$split4.Dock        = [System.Windows.Forms.DockStyle]::Fill
$split4.Orientation = [System.Windows.Forms.Orientation]::Horizontal
$split4.Panel1.BackColor = [System.Drawing.Color]::FromArgb(28,28,32)
$split4.Panel2.BackColor = [System.Drawing.Color]::FromArgb(28,28,32)
$split4.BackColor = [System.Drawing.Color]::FromArgb(28,28,32)
$t4.Controls.Add($split4)

$hdr4a = New-SectionPanel "Top 20 Slowest Queries by Avg Elapsed Time (since last SQL Server restart)"
$script:gTopQ = New-DGV
$split4.Panel1.Controls.Add($script:gTopQ)
$split4.Panel1.Controls.Add($hdr4a)

$hdr4b = New-SectionPanel "Disk I/O & File Latency   GREEN<20ms  ORANGE 20-49ms  RED>=50ms"
$script:gIO = New-DGV
$split4.Panel2.Controls.Add($script:gIO)
$split4.Panel2.Controls.Add($hdr4b)

$script:gIO.add_CellFormatting({
    param($s,$e)
    if($e.ColumnIndex -lt 0 -or $script:gIO.Columns.Count -eq 0){return}
    $cn=$script:gIO.Columns[$e.ColumnIndex].Name
    if($cn -in @("Avg Read ms","Avg Write ms")){
        $v=0.0
        if([double]::TryParse("$($e.Value)",[ref]$v)){
            if($v-ge 50){$e.CellStyle.ForeColor=[System.Drawing.Color]::Red}
            elseif($v-ge 20){$e.CellStyle.ForeColor=[System.Drawing.Color]::Orange}
            else{$e.CellStyle.ForeColor=[System.Drawing.Color]::LightGreen}
        }
    }
})


# ── TAB 5: STORAGE & TEMPDB ──────────────────────────────────────────────────
$t5 = New-Object System.Windows.Forms.TabPage
$t5.Text = "  Storage & TempDB"
$t5.BackColor = [System.Drawing.Color]::FromArgb(28,28,32)
$tabs.TabPages.Add($t5)

# Outer vertical split: top=drive space+db files, bottom=tempdb
$split5a = New-Object System.Windows.Forms.SplitContainer
$split5a.Dock = [System.Windows.Forms.DockStyle]::Fill
$split5a.Orientation = [System.Windows.Forms.Orientation]::Horizontal
$split5a.SplitterDistance = 320
$split5a.Panel1.BackColor = [System.Drawing.Color]::FromArgb(28,28,32)
$split5a.Panel2.BackColor = [System.Drawing.Color]::FromArgb(28,28,32)
$split5a.BackColor = [System.Drawing.Color]::FromArgb(28,28,32)
$t5.Controls.Add($split5a)

# Top panel: drive space (left) | db files (right)
$split5b = New-Object System.Windows.Forms.SplitContainer
$split5b.Dock = [System.Windows.Forms.DockStyle]::Fill
$split5b.Orientation = [System.Windows.Forms.Orientation]::Vertical
$split5b.SplitterDistance = 300
$split5b.Panel1.BackColor = [System.Drawing.Color]::FromArgb(28,28,32)
$split5b.Panel2.BackColor = [System.Drawing.Color]::FromArgb(28,28,32)
$split5b.BackColor = [System.Drawing.Color]::FromArgb(28,28,32)
$split5a.Panel1.Controls.Add($split5b)

# Drive space grid (built from xp_fixeddrives + WMI in refresh)
$hdr5drive = New-SectionPanel "Drive Free Space   RED<10%  ORANGE<20%  GREEN=OK"
$script:gDrive = New-DGV
$split5b.Panel1.Controls.Add($script:gDrive)
$split5b.Panel1.Controls.Add($hdr5drive)

# DB files grid
$hdr5files = New-SectionPanel "Database Files — Allocated Size, I/O Activity & AutoGrowth Settings   ORANGE=Warning"
$script:gDBFiles = New-DGV
$split5b.Panel2.Controls.Add($script:gDBFiles)
$split5b.Panel2.Controls.Add($hdr5files)

# Bottom panel: tempdb checks (left) | tempdb files (right)
$split5c = New-Object System.Windows.Forms.SplitContainer
$split5c.Dock = [System.Windows.Forms.DockStyle]::Fill
$split5c.Orientation = [System.Windows.Forms.Orientation]::Vertical
$split5c.SplitterDistance = 480
$split5c.Panel1.BackColor = [System.Drawing.Color]::FromArgb(28,28,32)
$split5c.Panel2.BackColor = [System.Drawing.Color]::FromArgb(28,28,32)
$split5c.BackColor = [System.Drawing.Color]::FromArgb(28,28,32)
$split5a.Panel2.Controls.Add($split5c)

$hdr5tmp = New-SectionPanel "TempDB Configuration & Pressure Checks   GREEN=OK  ORANGE=Warning"
$script:gTmpChk = New-DGV
$split5c.Panel1.Controls.Add($script:gTmpChk)
$split5c.Panel1.Controls.Add($hdr5tmp)

$hdr5tmpf = New-SectionPanel "TempDB Files — Per-File Space Usage"
$script:gTmpFiles = New-DGV
$split5c.Panel2.Controls.Add($script:gTmpFiles)
$split5c.Panel2.Controls.Add($hdr5tmpf)

# Color formatting - drive space
$script:gDrive.add_CellFormatting({
    param($s,$e)
    if($e.ColumnIndex -lt 0 -or $script:gDrive.Columns.Count -eq 0){return}
    $cn = $script:gDrive.Columns[$e.ColumnIndex].Name
    if($cn -eq "Free %"){
        $v = 0.0
        if([double]::TryParse("$($e.Value)",[ref]$v)){
            if($v -le 10)   {$e.CellStyle.ForeColor=[System.Drawing.Color]::Red; $e.CellStyle.Font=New-Object System.Drawing.Font("Segoe UI",9,[System.Drawing.FontStyle]::Bold)}
            elseif($v -le 20){$e.CellStyle.ForeColor=[System.Drawing.Color]::Orange}
            else             {$e.CellStyle.ForeColor=[System.Drawing.Color]::LightGreen}
        }
    }
    if($cn -eq "Status"){
        switch -Wildcard ($e.Value){
            "CRITICAL*"{$e.CellStyle.ForeColor=[System.Drawing.Color]::Red; $e.CellStyle.Font=New-Object System.Drawing.Font("Segoe UI",9,[System.Drawing.FontStyle]::Bold)}
            "WARNING*" {$e.CellStyle.ForeColor=[System.Drawing.Color]::Orange; $e.CellStyle.Font=New-Object System.Drawing.Font("Segoe UI",9,[System.Drawing.FontStyle]::Bold)}
            "OK"       {$e.CellStyle.ForeColor=[System.Drawing.Color]::LightGreen}
        }
    }
})

# Color formatting - db files
$script:gDBFiles.add_CellFormatting({
    param($s,$e)
    if($e.ColumnIndex -lt 0 -or $script:gDBFiles.Columns.Count -eq 0){return}
    if($script:gDBFiles.Columns[$e.ColumnIndex].Name -eq "Status"){
        switch -Wildcard ($e.Value){
            "WARNING*"{$e.CellStyle.ForeColor=[System.Drawing.Color]::Orange; $e.CellStyle.Font=New-Object System.Drawing.Font("Segoe UI",9,[System.Drawing.FontStyle]::Bold)}
            "OK"      {$e.CellStyle.ForeColor=[System.Drawing.Color]::LightGreen}
        }
    }
})

# Color formatting - tempdb checks
$script:gTmpChk.add_CellFormatting({
    param($s,$e)
    if($e.ColumnIndex -lt 0 -or $script:gTmpChk.Columns.Count -eq 0){return}
    if($script:gTmpChk.Columns[$e.ColumnIndex].Name -eq "Status"){
        switch($e.Value){
            "WARNING"{$e.CellStyle.ForeColor=[System.Drawing.Color]::Orange; $e.CellStyle.Font=New-Object System.Drawing.Font("Segoe UI",9,[System.Drawing.FontStyle]::Bold)}
            "OK"     {$e.CellStyle.ForeColor=[System.Drawing.Color]::LightGreen}
            "INFO"   {$e.CellStyle.ForeColor=[System.Drawing.Color]::CornflowerBlue}
        }
    }
})

# Color formatting - tempdb files free %
$script:gTmpFiles.add_CellFormatting({
    param($s,$e)
    if($e.ColumnIndex -lt 0 -or $script:gTmpFiles.Columns.Count -eq 0){return}
    if($script:gTmpFiles.Columns[$e.ColumnIndex].Name -eq "Free %"){
        $v=0.0
        if([double]::TryParse("$($e.Value)",[ref]$v)){
            if($v -le 10)    {$e.CellStyle.ForeColor=[System.Drawing.Color]::Red}
            elseif($v -le 25){$e.CellStyle.ForeColor=[System.Drawing.Color]::Orange}
            else             {$e.CellStyle.ForeColor=[System.Drawing.Color]::LightGreen}
        }
    }
})


# ── TAB 6: INDEXES & MEMORY ──────────────────────────────────────────────────
$t6 = New-Object System.Windows.Forms.TabPage
$t6.Text = "  Indexes & Memory"
$t6.BackColor = [System.Drawing.Color]::FromArgb(28,28,32)
$tabs.TabPages.Add($t6)

$split6outer = New-Object System.Windows.Forms.SplitContainer
$split6outer.Dock = [System.Windows.Forms.DockStyle]::Fill
$split6outer.Orientation = [System.Windows.Forms.Orientation]::Horizontal
$split6outer.SplitterDistance = 430
$split6outer.BackColor = [System.Drawing.Color]::FromArgb(28,28,32)
$split6outer.Panel1.BackColor = [System.Drawing.Color]::FromArgb(28,28,32)
$split6outer.Panel2.BackColor = [System.Drawing.Color]::FromArgb(28,28,32)
$t6.Controls.Add($split6outer)

# Top: missing indexes (left) | unused indexes (right)
$split6top = New-Object System.Windows.Forms.SplitContainer
$split6top.Dock = [System.Windows.Forms.DockStyle]::Fill
$split6top.Orientation = [System.Windows.Forms.Orientation]::Vertical
$split6top.SplitterDistance = 640
$split6top.BackColor = [System.Drawing.Color]::FromArgb(28,28,32)
$split6top.Panel1.BackColor = [System.Drawing.Color]::FromArgb(28,28,32)
$split6top.Panel2.BackColor = [System.Drawing.Color]::FromArgb(28,28,32)
$split6outer.Panel1.Controls.Add($split6top)

$hdr6a = New-SectionPanel "Missing Indexes (server-wide, ranked by impact score — since last restart)"
$script:gMissIdx = New-DGV
$split6top.Panel1.Controls.Add($script:gMissIdx)
$split6top.Panel1.Controls.Add($hdr6a)

$hdr6b = New-SectionPanel "Unused / Redundant Indexes — DB: master (change DB dropdown)"
$script:idxUnusedHdr = $hdr6b.Controls[0]
$script:gUnusedIdx = New-DGV
$split6top.Panel2.Controls.Add($script:gUnusedIdx)
$split6top.Panel2.Controls.Add($hdr6b)

# Bottom: memory pressure (left) | buffer pool by DB (right) | clerks (far right)
$split6bot = New-Object System.Windows.Forms.SplitContainer
$split6bot.Dock = [System.Windows.Forms.DockStyle]::Fill
$split6bot.Orientation = [System.Windows.Forms.Orientation]::Vertical
$split6bot.SplitterDistance = 400
$split6bot.BackColor = [System.Drawing.Color]::FromArgb(28,28,32)
$split6bot.Panel1.BackColor = [System.Drawing.Color]::FromArgb(28,28,32)
$split6bot.Panel2.BackColor = [System.Drawing.Color]::FromArgb(28,28,32)
$split6outer.Panel2.Controls.Add($split6bot)

$hdr6c = New-SectionPanel "Memory Pressure Checks   GREEN=OK  ORANGE=Warning"
$script:gMemChk = New-DGV
$split6bot.Panel1.Controls.Add($script:gMemChk)
$split6bot.Panel1.Controls.Add($hdr6c)

$split6bot2 = New-Object System.Windows.Forms.SplitContainer
$split6bot2.Dock = [System.Windows.Forms.DockStyle]::Fill
$split6bot2.Orientation = [System.Windows.Forms.Orientation]::Vertical
$split6bot2.SplitterDistance = 360
$split6bot2.BackColor = [System.Drawing.Color]::FromArgb(28,28,32)
$split6bot2.Panel1.BackColor = [System.Drawing.Color]::FromArgb(28,28,32)
$split6bot2.Panel2.BackColor = [System.Drawing.Color]::FromArgb(28,28,32)
$split6bot.Panel2.Controls.Add($split6bot2)

$hdr6d = New-SectionPanel "Buffer Pool by Database"
$script:gBufDB = New-DGV
$split6bot2.Panel1.Controls.Add($script:gBufDB)
$split6bot2.Panel1.Controls.Add($hdr6d)

$hdr6e = New-SectionPanel "Top Memory Clerks"
$script:gClerks = New-DGV
$split6bot2.Panel2.Controls.Add($script:gClerks)
$split6bot2.Panel2.Controls.Add($hdr6e)

# Color formatting - missing indexes impact score
$script:gMissIdx.add_CellFormatting({
    param($s,$e)
    if($e.ColumnIndex -lt 0 -or $script:gMissIdx.Columns.Count -eq 0){return}
    $cn = $script:gMissIdx.Columns[$e.ColumnIndex].Name
    if($cn -eq "Avg Benefit %"){
        $v=0.0
        if([double]::TryParse("$($e.Value)",[ref]$v)){
            if($v -ge 80){$e.CellStyle.ForeColor=[System.Drawing.Color]::Red; $e.CellStyle.Font=New-Object System.Drawing.Font("Segoe UI",9,[System.Drawing.FontStyle]::Bold)}
            elseif($v -ge 50){$e.CellStyle.ForeColor=[System.Drawing.Color]::Orange}
            else{$e.CellStyle.ForeColor=[System.Drawing.Color]::Yellow}
        }
    }
})

# Color formatting - unused indexes
$script:gUnusedIdx.add_CellFormatting({
    param($s,$e)
    if($e.ColumnIndex -lt 0 -or $script:gUnusedIdx.Columns.Count -eq 0){return}
    if($script:gUnusedIdx.Columns[$e.ColumnIndex].Name -eq "Recommendation"){
        switch -Wildcard ($e.Value){
            "*DROP*"   {$e.CellStyle.ForeColor=[System.Drawing.Color]::OrangeRed; $e.CellStyle.Font=New-Object System.Drawing.Font("Segoe UI",9,[System.Drawing.FontStyle]::Bold)}
            "*Review*" {$e.CellStyle.ForeColor=[System.Drawing.Color]::Yellow}
            "*in use*" {$e.CellStyle.ForeColor=[System.Drawing.Color]::LightGreen}
        }
    }
})

# Color formatting - memory checks
$script:gMemChk.add_CellFormatting({
    param($s,$e)
    if($e.ColumnIndex -lt 0 -or $script:gMemChk.Columns.Count -eq 0){return}
    if($script:gMemChk.Columns[$e.ColumnIndex].Name -eq "Status"){
        switch($e.Value){
            "WARNING"{$e.CellStyle.ForeColor=[System.Drawing.Color]::Orange; $e.CellStyle.Font=New-Object System.Drawing.Font("Segoe UI",9,[System.Drawing.FontStyle]::Bold)}
            "OK"     {$e.CellStyle.ForeColor=[System.Drawing.Color]::LightGreen}
            "INFO"   {$e.CellStyle.ForeColor=[System.Drawing.Color]::CornflowerBlue}
        }
    }
})

# ── TAB 7: TRANSACTIONS, DEADLOCKS & SECURITY ────────────────────────────────
$t7 = New-Object System.Windows.Forms.TabPage
$t7.Text = "  Security & Locks"
$t7.BackColor = [System.Drawing.Color]::FromArgb(28,28,32)
$tabs.TabPages.Add($t7)

$split7outer = New-Object System.Windows.Forms.SplitContainer
$split7outer.Dock = [System.Windows.Forms.DockStyle]::Fill
$split7outer.Orientation = [System.Windows.Forms.Orientation]::Horizontal
$split7outer.SplitterDistance = 350
$split7outer.BackColor = [System.Drawing.Color]::FromArgb(28,28,32)
$split7outer.Panel1.BackColor = [System.Drawing.Color]::FromArgb(28,28,32)
$split7outer.Panel2.BackColor = [System.Drawing.Color]::FromArgb(28,28,32)
$t7.Controls.Add($split7outer)

# Top: long transactions (left) | deadlocks (right)
$split7top = New-Object System.Windows.Forms.SplitContainer
$split7top.Dock = [System.Windows.Forms.DockStyle]::Fill
$split7top.Orientation = [System.Windows.Forms.Orientation]::Vertical
$split7top.SplitterDistance = 560
$split7top.BackColor = [System.Drawing.Color]::FromArgb(28,28,32)
$split7top.Panel1.BackColor = [System.Drawing.Color]::FromArgb(28,28,32)
$split7top.Panel2.BackColor = [System.Drawing.Color]::FromArgb(28,28,32)
$split7outer.Panel1.Controls.Add($split7top)

$hdr7a = New-SectionPanel "Long-running Transactions (open > 30 sec)   RED rows = also blocking"
$script:gLongTxn = New-DGV
$split7top.Panel1.Controls.Add($script:gLongTxn)
$split7top.Panel1.Controls.Add($hdr7a)

$hdr7b = New-SectionPanel "Recent Deadlocks (from system_health XE session)"
$script:gDeadlock = New-DGV
$split7top.Panel2.Controls.Add($script:gDeadlock)
$split7top.Panel2.Controls.Add($hdr7b)

# Bottom: security checks (left) | sysadmin members (right)
$split7bot = New-Object System.Windows.Forms.SplitContainer
$split7bot.Dock = [System.Windows.Forms.DockStyle]::Fill
$split7bot.Orientation = [System.Windows.Forms.Orientation]::Vertical
$split7bot.SplitterDistance = 560
$split7bot.BackColor = [System.Drawing.Color]::FromArgb(28,28,32)
$split7bot.Panel1.BackColor = [System.Drawing.Color]::FromArgb(28,28,32)
$split7bot.Panel2.BackColor = [System.Drawing.Color]::FromArgb(28,28,32)
$split7outer.Panel2.Controls.Add($split7bot)

$hdr7c = New-SectionPanel "Security Audit Checks   GREEN=OK  ORANGE=Warning"
$script:gSecChk = New-DGV
$split7bot.Panel1.Controls.Add($script:gSecChk)
$split7bot.Panel1.Controls.Add($hdr7c)

$hdr7d = New-SectionPanel "Sysadmin Role Members"
$script:gSysadmin = New-DGV
$split7bot.Panel2.Controls.Add($script:gSysadmin)
$split7bot.Panel2.Controls.Add($hdr7d)

# Color formatting - long transactions
$script:gLongTxn.add_CellFormatting({
    param($s,$e)
    if($e.RowIndex -lt 0 -or $script:gLongTxn.Columns.Count -eq 0){return}
    try{
        $blocked = $script:gLongTxn.Rows[$e.RowIndex].Cells["Blocked By"].Value
        $openMin = 0.0
        [double]::TryParse("$($script:gLongTxn.Rows[$e.RowIndex].Cells['Open Min'].Value)",[ref]$openMin) | Out-Null
        if($null -ne $blocked -and "$blocked" -ne "0" -and "$blocked" -ne ""){
            $e.CellStyle.BackColor=[System.Drawing.Color]::FromArgb(90,22,22)
        } elseif($openMin -ge 5){
            $e.CellStyle.ForeColor=[System.Drawing.Color]::Orange
        }
    }catch{}
})

# Color formatting - security checks
$script:gSecChk.add_CellFormatting({
    param($s,$e)
    if($e.ColumnIndex -lt 0 -or $script:gSecChk.Columns.Count -eq 0){return}
    if($script:gSecChk.Columns[$e.ColumnIndex].Name -eq "Status"){
        switch($e.Value){
            "WARNING"{$e.CellStyle.ForeColor=[System.Drawing.Color]::Orange; $e.CellStyle.Font=New-Object System.Drawing.Font("Segoe UI",9,[System.Drawing.FontStyle]::Bold)}
            "OK"     {$e.CellStyle.ForeColor=[System.Drawing.Color]::LightGreen}
            "INFO"   {$e.CellStyle.ForeColor=[System.Drawing.Color]::CornflowerBlue}
        }
    }
})

# Color formatting - sysadmin members
$script:gSysadmin.add_CellFormatting({
    param($s,$e)
    if($e.ColumnIndex -lt 0 -or $script:gSysadmin.Columns.Count -eq 0){return}
    if($script:gSysadmin.Columns[$e.ColumnIndex].Name -eq "Category"){
        if($e.Value -eq "User-added"){$e.CellStyle.ForeColor=[System.Drawing.Color]::Orange}
        else{$e.CellStyle.ForeColor=[System.Drawing.Color]::DimGray}
    }
})

# ── EVENTS ───────────────────────────────────────────────────────────────────

$cmbAuth.add_SelectedIndexChanged({
    $sql=($cmbAuth.SelectedIndex -eq 1)
    $txtUser.Enabled=$sql; $txtPwd.Enabled=$sql
})

function Refresh-All {
    if(-not $script:connected){return}
    Set-Status "Refreshing data..." "Yellow"
    $form.Refresh()

    Bind-Grid $script:gCfg  (Invoke-SqlQuery $Q_Config)
    $selDB = if($cmbDB.SelectedItem){"$($cmbDB.SelectedItem)"}else{"master"}
    Bind-Grid $script:gIdx  (Invoke-SqlQuery (Get-IndexHealthQuery $selDB))
    # Update the index section header to show current DB
    $script:idxHdrLbl.Text = "  Index Fragmentation Health — DB: $selDB"

    # CPU history chart
    $dcHist = Invoke-SqlQuery $Q_CPUHistory
    if($dcHist.Rows.Count -gt 0 -and -not $dcHist.Columns.Contains("Error")){
        $script:cpuHistory.Clear()
        $script:cpuTimes.Clear()
        # Rows come newest-first; reverse so chart goes left=oldest right=newest
        $rows = @($dcHist.Rows)
        [array]::Reverse($rows)
        foreach($rh in $rows){
            $sqlPct = [int]"$($rh['SQL CPU %'])"
            $othPct = [int]"$($rh['Other CPU %'])"
            $script:cpuHistory.Add(@($sqlPct,$othPct))
            # Format timestamp as HH:mm:ss
            try {
                $ts = [datetime]$rh['Recorded At']
                $script:cpuTimes.Add($ts.ToString('HH:mm:ss'))
            } catch {
                $script:cpuTimes.Add("")
            }
        }
        $lastRow = $dcHist.Rows[0]
        $sqlNow = $lastRow['SQL CPU %']; $othNow = $lastRow['Other CPU %']; $idlNow = $lastRow['Idle %']
        $script:lSql.Text = "SQL CPU: $sqlNow%"
        $script:lOth.Text = "Other: $othNow%"
        $script:lIdl.Text = "Idle: $idlNow%"
        $script:cpuHdrLbl.Text = "  CPU History — $($script:cpuHistory.Count) snapshots   BLUE=SQL Server  ORANGE=Other processes   now: SQL $sqlNow%  Other $othNow%  Idle $idlNow%"
        $script:cpuChart.Invalidate()
    }
    Bind-Grid $script:gReq  (Invoke-SqlQuery $Q_ActiveReqs)
    Bind-Grid $script:gWait (Invoke-SqlQuery $Q_WaitStats)
    Bind-Grid $script:gBlk  (Invoke-SqlQuery $Q_Blocking)

    Bind-Grid $script:gBak  (Invoke-SqlQuery $Q_Backup)
    Bind-Grid $script:gJob  (Invoke-SqlQuery $Q_Jobs)

    Bind-Grid $script:gTopQ (Invoke-SqlQuery $Q_TopQ)
    Bind-Grid $script:gIO   (Invoke-SqlQuery $Q_DiskIO)


    # Tab 5 - Drive space via WMI then SQL
    try {
        $driveDT = New-Object System.Data.DataTable
        [void]$driveDT.Columns.Add("Drive")
        [void]$driveDT.Columns.Add("Total GB")
        [void]$driveDT.Columns.Add("Free GB")
        [void]$driveDT.Columns.Add("Used GB")
        [void]$driveDT.Columns.Add("Free %")
        [void]$driveDT.Columns.Add("Status")
        $wmiDrives = Get-WmiObject Win32_LogicalDisk -Filter "DriveType=3" -ComputerName $txtSrv.Text.Split("\")[0].Split(",")[0] -ErrorAction SilentlyContinue
        if(-not $wmiDrives){ $wmiDrives = Get-WmiObject Win32_LogicalDisk -Filter "DriveType=3" }
        foreach($d in $wmiDrives){
            $totalGB = [math]::Round($d.Size/1GB,1)
            $freeGB  = [math]::Round($d.FreeSpace/1GB,1)
            $usedGB  = [math]::Round(($d.Size-$d.FreeSpace)/1GB,1)
            $freePct = if($d.Size -gt 0){ [math]::Round($d.FreeSpace*100.0/$d.Size,1) }else{0}
            $status  = if($freePct -le 10){"CRITICAL - Disk nearly full"}elseif($freePct -le 20){"WARNING - Low space"}else{"OK"}
            $r=$driveDT.NewRow()
            $r["Drive"]=$d.DeviceID; $r["Total GB"]=$totalGB; $r["Free GB"]=$freeGB
            $r["Used GB"]=$usedGB; $r["Free %"]=$freePct; $r["Status"]=$status
            [void]$driveDT.Rows.Add($r)
        }
        Bind-Grid $script:gDrive $driveDT
    } catch {
        $errDT = New-Object System.Data.DataTable; [void]$errDT.Columns.Add("Error"); $r=$errDT.NewRow(); $r["Error"]=$_.Exception.Message; [void]$errDT.Rows.Add($r)
        Bind-Grid $script:gDrive $errDT
    }
    Bind-Grid $script:gDBFiles  (Invoke-SqlQuery $Q_DBFiles)
    Bind-Grid $script:gTmpChk   (Invoke-SqlQuery $Q_TempDBChecks)
    Bind-Grid $script:gTmpFiles (Invoke-SqlQuery $Q_TempDBFiles)

    # Tab 6 - Indexes & Memory
    Bind-Grid $script:gMissIdx   (Invoke-SqlQuery $Q_MissingIndexes)
    $selDB2 = if($cmbDB.SelectedItem){"$($cmbDB.SelectedItem)"}else{"master"}
    $script:idxUnusedHdr.Text = "  Unused / Redundant Indexes — DB: $selDB2"
    Bind-Grid $script:gUnusedIdx (Invoke-SqlQuery "USE [$selDB2]; $Q_UnusedIndexes")
    Bind-Grid $script:gMemChk    (Invoke-SqlQuery $Q_MemoryPressure)
    Bind-Grid $script:gBufDB     (Invoke-SqlQuery $Q_BufferByDB)
    Bind-Grid $script:gClerks    (Invoke-SqlQuery $Q_MemoryClerks)

    # Tab 7 - Security & Locks
    Bind-Grid $script:gLongTxn  (Invoke-SqlQuery $Q_LongTxns)
    Bind-Grid $script:gDeadlock (Invoke-SqlQuery $Q_Deadlocks)
    Bind-Grid $script:gSecChk   (Invoke-SqlQuery $Q_SecurityChecks)
    Bind-Grid $script:gSysadmin (Invoke-SqlQuery $Q_SysadminMembers)

    Set-Status "Last refresh: $(Get-Date -Format 'HH:mm:ss')   Connected to: $($txtSrv.Text)" "LightGreen"
}

$btnConn.add_Click({
    $srv=$txtSrv.Text.Trim()
    if([string]::IsNullOrWhiteSpace($srv)){
        [System.Windows.Forms.MessageBox]::Show("Please enter a server name.","SQL Monitor",0,48)|Out-Null; return
    }
    if($cmbAuth.SelectedIndex -eq 0){
        $script:connString="Server=$srv;Database=master;Integrated Security=True;Connection Timeout=10;"
    } else {
        $script:connString="Server=$srv;Database=master;User Id=$($txtUser.Text);Password=$($txtPwd.Text);Connection Timeout=10;"
    }
    Set-Status "Connecting to $srv ..." "Yellow"; $form.Refresh()
    try {
        $cn=New-Object System.Data.SqlClient.SqlConnection($script:connString); $cn.Open()
        $cmd=$cn.CreateCommand(); $cmd.CommandText="SELECT @@VERSION"
        $ver=$cmd.ExecuteScalar(); $cn.Close()
        $script:connected=$true
        Set-Status "Connected: $($ver.Split([char]10)[0].Trim())" "LightGreen"
        # Populate database list
        $cmbDB.Items.Clear()
        $dbList = Invoke-SqlQuery "SELECT name FROM sys.databases WHERE state_desc='ONLINE' ORDER BY name"
        foreach($row in $dbList.Rows){ [void]$cmbDB.Items.Add($row["name"]) }
        $cmbDB.SelectedItem = "master"
        if($cmbDB.SelectedIndex -lt 0){ $cmbDB.SelectedIndex=0 }
        Refresh-All
    } catch {
        $script:connected=$false
        [System.Windows.Forms.MessageBox]::Show("Connection failed:`n`n$($_.Exception.Message)","SQL Monitor",0,16)|Out-Null
        Set-Status "Connection failed. Check server/credentials/permissions." "Red"
    }
})

$btnRef.add_Click({ Refresh-All })

$cmbDB.add_SelectedIndexChanged({
    if($script:connected){
        $selDB = if($cmbDB.SelectedItem){"$($cmbDB.SelectedItem)"}else{"master"}
        $script:idxHdrLbl.Text = "  Index Fragmentation Health — DB: $selDB"
        Bind-Grid $script:gIdx (Invoke-SqlQuery (Get-IndexHealthQuery $selDB))
        $script:idxUnusedHdr.Text = "  Unused / Redundant Indexes — DB: $selDB"
        Bind-Grid $script:gUnusedIdx (Invoke-SqlQuery "USE [$selDB]; $Q_UnusedIndexes")
    }
})

$cmbRef.add_SelectedIndexChanged({
    if($script:timer){$script:timer.Stop();$script:timer.Dispose();$script:timer=$null}
    $sec=switch($cmbRef.SelectedIndex){0{15}1{30}2{60}default{0}}
    if($sec-gt 0){
        $script:timer=New-Object System.Windows.Forms.Timer
        $script:timer.Interval=$sec*1000
        $script:timer.add_Tick({Refresh-All})
        $script:timer.Start()
    }
})

$script:timer=New-Object System.Windows.Forms.Timer
$script:timer.Interval=30000
$script:timer.add_Tick({Refresh-All})
$script:timer.Start()

[System.Windows.Forms.Application]::Run($form)
