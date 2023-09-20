select top 1000 
	JR.Job_Number,
	v.fileid, 
	v.Source,
	Folder,
	v.filename,
	v.createddate,
	v.updateddate,
	v.Supplieddate,
	v.filesize 
	from dbo.vw_ExistingFiles v
	left join dbo.CT_RejectedFiles r
		on r.fileID = v.FileID and r.FileSource = v.Source and r.UnRejectedDate is null
	left join dbo.CT_JobRun JR
		on jr.fileID = v.FileID and jr.FileSource = v.Source 
	left join dbo.vw_CallTraceJobList JN
		on jn.InboundFolder = Folder
	where r.fileid is null and jr.fileid is null
	order by Updateddate







 SELECT TOP 1000 JR.JobNumber, v.fileid, v.Source, Folder, v.filename, v.createddate, v.updateddate, v.Supplieddate, v.filesize
FROM dbo.vw_ExistingFiles v
LEFT JOIN dbo.CT_RejectedFiles r ON r.fileID = v.FileID AND r.FileSource = v.Source AND r.UnRejectedDate IS NULL
LEFT JOIN dbo.CT_JobRun JR ON jr.fileID = v.FileID AND jr.FileSource = v.Source
LEFT JOIN dbo.vw_CallTraceJobList JN ON jn.InboundFolder = Folder
WHERE r.fileid IS NULL AND jr.fileid IS NULL
ORDER BY Updateddate;

