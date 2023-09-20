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
