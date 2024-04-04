WITH RunStatusForJobs AS (
   select distinct CT_JobID,FileID from dbo.CT_JobRun where RunStatus not in ('Neptune Complete', 'In QC', 'Verified','Cancelled', 'Complete')
)


--DPPLaunch AS (
--	select nv.varValue as Environment, ce.[Enabled]
--		from		Neptune.dbo.NeptuneVariables nv
--		inner join	Cobra.Environments ce
--			on nv.varName = 'DPP_Default_BatchWorkflowEndpoint' and nv.varValue = ce.Environment
--)

select top 1000 
	CASE WHEN COUNT(JN.Job_Number) = 0 THEN 'No Matches' 
	WHEN COUNT(JN.Job_Number) > 1 THEN 'Multiple Matches' 
	WHEN COUNT(JN.Job_Number) = 1 AND JB.JobActive <> 1 THEN  MAX(JN.Job_Number) + '-NA' --Not active
	ELSE MAX(JN.Job_Number) 
END AS Job_Number,
	v.fileid, 
	v.Source, 
	Folder,
	v.filename,
	format(v.Supplieddate, 'yyyy-MM-dd HH:mm') as ReceivedDate,
	v.filesize,
	CASE
		WHEN EXISTS (Select 1 from dbo.CT_Config where ConfigItem = 'Launch Enabled' and ConfigValue = 'False') THEN 'False'
		WHEN CB.JobType = 'DPP' and (
			select [Enabled]
			from		Neptune.dbo.NeptuneVariables nv
			inner join	Cobra.Environments ce
			on nv.varName = 'DPP_Default_BatchWorkflowEndpoint' and nv.varValue = ce.Environment
			) <> 1 Then 'False'
		ELSE 'True'
	END AS CanLaunch,
	JB.JobActive
from dbo.vw_ExistingFiles_New v
left join dbo.CT_RejectedFiles r
	on r.fileID = v.FileID and r.FileSource = v.Source and r.UnRejectedDate is null
left join dbo.CT_JobRun JR
	on jr.fileID = v.FileID and jr.FileSource = v.Source 
left join dbo.vw_CallTraceJobList JN
	on jn.InboundFolder = Folder and v.FileName like jn.InboundFileSpec
left join RunStatusForJobs FR
	on fr.FileID = v.FileID
left join [dbo].[vw_CallTraceCurrentJobs] CB
	on JN.Job_Number = CB.JobNo
left join [DataBureauDataLoadAudit].[dbo].[CT_Jobs] JB
	on JN.Job_Number = JB.Job_Number
where r.fileid is null and jr.fileid is null and v.FileID is not null 
GROUP BY 
	v.fileid, 
	v.Source, 
	Folder, 
	v.filename, 
	v.createddate, 
	v.updateddate, 
	v.Supplieddate, 
	v.filesize,
	CB.JobType,
	JB.JobActive
ORDER BY Updateddate;
