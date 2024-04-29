Visual Updates:
	File List page:
•	In the File List window table, there an extra column “CanLaunch”, 1 – an operator is allowed to launch a job, 0 – cannot launch a job, probably launch is disabled on the database
•	“Launch Job” check box. Initially checked box would be selected, which means a file will be linked, set and launched automatically 
•	If a file has an inactive job then you will “-NA” at the end of a job number an you won’t be able to process a file.

Job Details page:
•	If  job was launched automatically the “Launch Job” button will be not at the page. If automatic launch would not be successful or the “Launch Job” check box wasn’t selected then the you would be able to launch a job from job details page.

Background Changes:
	Code was updated to allow launch jobs automatically, “JobLaunch” function was created.
•	An operator wouldn’t be allowed to launch a job, if job is inactive or “CanLaunch” column has “0” as value, according message box will be received
•	“CanLaunchJob” (does additional checks if a job can be launched) function was created in order to tidy up “JobLaunch” function.
•	For database updates there were created following views and functions- vw_CallTraceFilesToDisplay_New_2, vw_CT_JobNumberStatus, fn_CT_CanLaunch.
