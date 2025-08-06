@echo off
REM ---
REM Runs the VBScript and writes its output to a file.
REM The "//nologo" flag is used to suppress the cscript banner.
REM The "for /f" command processes the output of the command in the single quotes.
REM The output is then redirected to "meeting-time.txt".
REM After creating the file, it's added to git, committed, and pushed.
REM ---
for /f "tokens=*" %%a in ('cscript //nologo get-meeting-time.vbs') do (
    echo %%a > meeting-time.txt
)

REM Add the new file to the git staging area
git add meeting-time.txt

REM Commit the changes with a message
REM You might want to make this commit message more dynamic, e.g., by including the date.
git commit -m "Update meeting time"

REM Push the changes to the remote repository
git push
