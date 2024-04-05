CREATE View [dbo].[vw_CallTraceFilesToDisplay_New_TestV2] as


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
LEFT JOIN [DataBureauDataLoadAudit].[dbo].[vw_CT_JobNumberStatus_Test] JNS ON JNS.fileid = v.FileID
LEFT JOIN [DataBureauDataLoadAudit].[dbo].[CT_Jobs] JB ON JNS.MaxJobNumber = JB.Job_Number
LEFT JOIN [DataBureauDataLoadAudit].[dbo].[vw_CT_LaunchCapability_Test] LC ON LC.fileid = v.FileID
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

GO



Column 'DataBureauDataLoadAudit.dbo.vw_CT_JobNumberStatus_Test.JobStatus' is invalid in the select list because it is not contained in either an aggregate function or the GROUP BY clause.
