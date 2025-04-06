"# Excel-Automation" 

Enhancing the code from YouTube tutorial (Project #3) TS 1:42:08
https://www.youtube.com/watch?v=PXMJ6FS7llk&t=9316s


- The main file: 'pivot-table.py'combines the scripts into a single one
- The pre-executable file: 'pivot-table_executable.py'is an alteration probable to initiate executable file
		w/ additional content:
			- References to execution directory
			- Getting the moth (for report name) from the input line prompt
			- Checking for the existence of 'Results' folder if missing (to avoid exception if missing)


NOTE: to create the executable file, run from cmd: >> pyinstaller --onefile pivot-table_executable.py