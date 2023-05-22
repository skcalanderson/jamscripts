. .\delete_and_move_files.ps1
. .\SKC_BrightSpace_User_Query_Term_Year.ps1
. .\SKC_BrightSpace_Enrollment_Query_Year_Term.ps1

. python.exe .\user_difference.py
. python.exe .\enrollment_difference.py

. .\copy_user_enrollment_difference_files_to_upload.ps1
