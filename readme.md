WITH RunStatusForJobs AS (
   SELECT DISTINCT CT_JobID, FileID 
   FROM dbo.CT_JobRun 
   WHERE RunStatus NOT IN ('Neptune Complete', 'In QC', 'Verified', 'Cancelled', 'Complete')
)
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
LEFT JOIN [DataBureauDataLoadAudit].[dbo].[CT_Jobs] JB ON JNS.MaxJobNumber = JB.Job_Number
LEFT JOIN JobNumberStatus JNS ON JNS.fileid = v.FileID
LEFT JOIN LaunchCapability LC ON LC.fileid = v.FileID
WHERE r.fileid IS NULL 
      AND JR.fileid IS NULL 
      AND v.FileID IS NOT NULL 
GROUP BY 
    v.fileid, 
    v.Source, 
    v.Folder, 
    v.filename, 
    v.createddate, 
    v.updateddate, 
    v.Supplieddate, 
    v.filesize,
    LC.CanLaunch,
    JB.JobActive
ORDER BY v.Updateddate;



CREATE VIEW JobNumberStatus AS
SELECT
    v.fileid,
    (CASE WHEN COUNT(JN.Job_Number) = 0 THEN 'No Matches' 
          WHEN COUNT(JN.Job_Number) > 1 THEN 'Multiple Matches' 
          WHEN COUNT(JN.Job_Number) = 1 AND JB.JobActive <> 1 THEN MAX(JN.Job_Number) + '-NA' -- Not active
          ELSE MAX(JN.Job_Number) 
     END) AS JobStatus,
    COUNT(JN.Job_Number) AS JobNumberCount,
    MAX(JN.Job_Number) AS MaxJobNumber,
    JB.JobActive
FROM dbo.vw_ExistingFiles_New v
LEFT JOIN dbo.vw_CallTraceJobList JN ON jn.InboundFolder = v.Folder AND v.FileName LIKE jn.InboundFileSpec
LEFT JOIN [DataBureauDataLoadAudit].[dbo].[CT_Jobs] JB ON JN.Job_Number = JB.Job_Number
GROUP BY v.fileid, JB.JobActive;




CREATE VIEW LaunchCapability AS
SELECT
    v.fileid,
    (CASE
        WHEN EXISTS (SELECT 1 FROM dbo.CT_Config WHERE ConfigItem = 'Launch Enabled' AND ConfigValue = 'False') THEN 'False'
        WHEN CB.JobType = 'DPP' AND (
            SELECT [Enabled]
            FROM Neptune.dbo.NeptuneVariables nv
            INNER JOIN Cobra.Environments ce ON nv.varName = 'DPP_Default_BatchWorkflowEndpoint' AND nv.varValue = ce.Environment
        ) <> 1 THEN 'False'
        ELSE 'True'
     END) AS CanLaunch,
    CB.JobType
FROM dbo.vw_ExistingFiles_New v
LEFT JOIN [dbo].[vw_CallTraceCurrentJobs] CB ON JN.Job_Number = CB.JobNo
LEFT JOIN dbo.vw_CallTraceJobList JN ON jn.InboundFolder = v.Folder AND v.FileName LIKE jn.InboundFileSpec;
