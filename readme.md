CREATE VIEW LaunchCapability AS
SELECT
    v.fileid,
    (CASE
        WHEN EXISTS (SELECT 1 FROM dbo.CT_Config WHERE ConfigItem = 'Launch Enabled' AND ConfigValue = 'False') THEN 'False'
        WHEN CB.JobType = 'DPP' AND (
            SELECT TOP 1 [Enabled]
            FROM Neptune.dbo.NeptuneVariables nv
            INNER JOIN Cobra.Environments ce ON nv.varName = 'DPP_Default_BatchWorkflowEndpoint' AND nv.varValue = ce.Environment
        ) <> 1 THEN 'False'
        ELSE 'True'
     END) AS CanLaunch,
    CB.JobType
FROM dbo.vw_ExistingFiles_New v
LEFT JOIN dbo.vw_CallTraceJobList JN ON v.Folder = JN.InboundFolder AND v.FileName LIKE JN.InboundFileSpec
LEFT JOIN [dbo].[vw_CallTraceCurrentJobs] CB ON CB.JobNo = JN.Job_Number;
