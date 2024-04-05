IF EXISTS (SELECT * FROM sys.views WHERE object_id = OBJECT_ID(N'[dbo].[vw_CallTraceFilesToDisplay_New_TestV2]'))
DROP VIEW [dbo].[vw_CallTraceFilesToDisplay_New_TestV2];
GO

CREATE VIEW [dbo].[vw_CallTraceFilesToDisplay_New_TestV2] AS

WITH RunStatusForJobs AS (
   SELECT DISTINCT CT_JobID, FileID 
   FROM dbo.CT_JobRun 
   WHERE RunStatus NOT IN ('Neptune Complete', 'In QC', 'Verified', 'Cancelled', 'Complete')
),
PreparedData AS (
    SELECT TOP 1000 
        JNS.JobStatus AS Job_Number,
        v.fileid, 
        v.Source,
        v.Folder,
        v.filename,
        FORMAT(v.Supplieddate, 'yyyy-MM-dd HH:mm') AS ReceivedDate,
        v.filesize,
        LC.CanLaunch,
        JB.JobActive
    FROM dbo.vw_ExistingFiles_New v
    LEFT JOIN dbo.CT_RejectedFiles r ON r.fileID = v.FileID AND r.FileSource = v.Source AND r.UnRejectedDate IS NULL
    LEFT JOIN dbo.CT_JobRun JR ON JR.fileID = v.FileID AND JR.FileSource = v.Source 
    LEFT JOIN RunStatusForJobs FR ON FR.FileID = v.FileID
    LEFT JOIN [DataBureauDataLoadAudit].[dbo].[vw_CT_JobNumberStatus_Test] JNS ON JNS.fileid = v.FileID
    LEFT JOIN [DataBureauDataLoadAudit].[dbo].[CT_Jobs] JB ON JNS.MaxJobNumber = JB.Job_Number
    LEFT JOIN [DataBureauDataLoadAudit].[dbo].[vw_CT_LaunchCapability_Test] LC ON LC.fileid = v.FileID
    WHERE r.fileid IS NULL 
          AND JR.fileid IS NULL 
          AND v.FileID IS NOT NULL 
)
SELECT 
    Job_Number,
    fileid, 
    Source,
    Folder,
    filename,
    ReceivedDate,
    filesize,
    CASE 
        WHEN Job_Number IN ('No Matches', 'Multiple Matches') OR Job_Number LIKE '%-NA' THEN 'False'
        ELSE CanLaunch
    END AS CanLaunch,
    JobActive
FROM PreparedData
ORDER BY UpdatedDate;

GO
