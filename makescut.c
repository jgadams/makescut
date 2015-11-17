/*
Copyright 2015 Justin Gregory Adams.

This program is free software: you can redistribute it and/or modify
it under the terms of the GNU General Public License as published by
the Free Software Foundation, version 3 only.

This program is distributed in the hope that it will be useful,
but WITHOUT ANY WARRANTY; without even the implied warranty of
MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
GNU General Public License for more details.

You should have received a copy of the GNU General Public License
along with this program.  If not, see <http://www.gnu.org/licenses/>.
*/


#include <stdio.h>
#include <stdlib.h>
#include <string.h>


/* Exit codes. */
#define MAKESCUT_USAGE               5
#define MAKESCUT_NULL_POINTER        7
#define MAKESCUT_FOPEN              11

#define PATHMAX 1024
#define COMMMAX (PATHMAX * 2)


int usage()
{
	fprintf(stderr, "Usage:\n");
	fprintf(stderr, "makescut  --lnk LNK_PATH  --target TARGET_PATH   [--icon ICON_PATH]  [--work WORKING_DIRECTORY]  [--args ARGUMENTS]\n");
	fprintf(stderr, "   Creates a shortcut file at LNK_PATH, pointing to TARGET_PATH. LNK_PATH must include \".lnk\" extension!\n");
	fprintf(stderr, "makescut  --help\n");
	fprintf(stderr, "makescut  --license\n\n");
	fprintf(stderr, "Example: makescut  --lnk C:\\Temp\\Google.lnk  --target \"C:\\Program Files (x86)\\Internet Explorer\\iexplore.exe\"  --icon C:\\Temp\\Google.ico  --args http://google.com/\n");
	fprintf(stderr, "Example: makescut  --lnk C:\\Temp\\program.lnk  --target \"C:\\Program Files\\Program\\program.exe\"  --icon \"C:\\Program Files\\Program\\program.ico\"  --args \"\\\"arg 1\\\" \\\"arg 2\\\" \\\"arg 3\\\"\"\n");
	exit(MAKESCUT_USAGE);
	return 0;
}


int license()
{
	fprintf(stderr, "Copyright 2015 Justin Gregory Adams.\n\n");

	fprintf(stderr, "This program is free software: you can redistribute it and/or modify\n");
	fprintf(stderr, "it under the terms of the GNU General Public License as published by\n");
	fprintf(stderr, "the Free Software Foundation, version 3 only.\n\n");

	fprintf(stderr, "This program is distributed in the hope that it will be useful,\n");
	fprintf(stderr, "but WITHOUT ANY WARRANTY; without even the implied warranty of\n");
	fprintf(stderr, "MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the\n");
	fprintf(stderr, "GNU General Public License for more details.\n\n");

	fprintf(stderr, "You should have received a copy of the GNU General Public License\n");
	fprintf(stderr, "along with this program.  If not, see <http://www.gnu.org/licenses/>.\n");

	exit(0);
	return 0;
}


/* Writes string with double quotes optionally escaped. Writes at most max characters from str (output may be more due to escapes).*/
int writeEscaped(FILE* file, char* str, int includeDQuotes, int max)
{
	int i, stop;

	if((file == NULL) || (str == NULL))
	{
		fprintf(stderr, "writeEscaped(): null pointer.\n");
		exit(MAKESCUT_NULL_POINTER);
	}

	stop= strnlen(str, PATHMAX);
	if(stop > max)
		stop= max;

	for(i= 0; i < stop; i++)
	{
		if(str[i] == '"')
		{
			if(includeDQuotes)
			{
				fputc('"', file);
				fputc('"', file);
			}
		}
		else
			fputc(str[i], file);
	}

	return 0;
}


/*
We will open a temp file (within the %TEMP% environment variable), write VBScript to it, and run the temp file.
Yes, it's lame. No, I don't have a better idea, other than writing the .lnk directly, which I don't have time for at the moment.
File is NOT deleted in case debugging is needed.
For VBScript, we need to escape double quotes. We will eliminate double quotes in target path, icon location,
and working directory, since we are going to double quote them anyway. We will need to keep and escape double quotes
in the arg list, as it may contain multiple args each delimited by double quotes.

To create a LNK, we must run VBScript similar to the following:

set WshShell = WScript.CreateObject("WScript.Shell")
set oShellLink = WshShell.CreateShortcut("C:\Path\to\Shortcut.lnk")
oShellLink.TargetPath = "C:\Path\to\the\target\fubar.exe"
oShellLink.IconLocation = "C:\Path\to\Icon.ico"
oShellLink.WorkingDirectory = "C:\Path\to\the\target"
oShellLink.Save
*/
int makescut(char* lnkPath, char* targetPath, char* iconPath, char* workingDirectory, char* arguments)
{
	FILE* vbs;
	char path[PATHMAX + 1];
	char comm[COMMMAX + 1];
	size_t workLen, n;

	/* lnkPath and targetPath are required */
	if((lnkPath == NULL) || (targetPath == NULL))
	{
		fprintf(stderr, "makescut(): null pointer.\n");
		exit(MAKESCUT_NULL_POINTER);
	}

	/* Open file. */
	path[PATHMAX]= 0;
	_snprintf(path, PATHMAX, "%s\\makescut.vbs", getenv("TEMP"));
	vbs= fopen(path, "w");
	if(vbs == NULL)
	{
		fprintf(stderr, "makescut(): could not open temp file %s\n", path);
		exit(MAKESCUT_FOPEN);
	}

	/* Begin file write. */
	fprintf(vbs, "set WshShell = WScript.CreateObject(\"WScript.Shell\")\n");

	/* LNK path */
	fprintf(vbs, "set oShellLink = WshShell.CreateShortcut(\"");
	writeEscaped(vbs, lnkPath, 0, PATHMAX);
	fprintf(vbs, "\")\n");

	/* Target path */
	fprintf(vbs, "oShellLink.TargetPath = \"");
	writeEscaped(vbs, targetPath, 0, PATHMAX);
	fprintf(vbs, "\"\n");

	/* Working directory */
	/* If no working directory is specified then use directory containing the target. */
	if(workingDirectory == NULL)
	{
		workingDirectory= targetPath;
		for(workLen= strlen(workingDirectory) - 1; workLen > 0; workLen--)
		{
			if(workingDirectory[workLen] == '\\')
			{
				/* Do not terminate if workingDirectory ends with backslash. However, do terminate if we have a string like 'C:\' */
				if((workLen == strlen(workingDirectory) - 1) && (workLen > 2))
					continue;
				else
					break;
			}
		}
	}
	else
	{
		workLen= strlen(workingDirectory);
	}
	fprintf(vbs, "oShellLink.WorkingDirectory = \"");
	writeEscaped(vbs, workingDirectory, 0, workLen);
	fprintf(vbs, "\"\n");


	if(iconPath != NULL)
	{
		fprintf(vbs, "oShellLink.IconLocation = \"");
		writeEscaped(vbs, iconPath, 0, PATHMAX);
		fprintf(vbs, "\"\n");
	}
	if(arguments != NULL)
	{
		fprintf(vbs, "oShellLink.Arguments = \"");
		writeEscaped(vbs, arguments, 1, PATHMAX); /* The only instance where we keep double quotes. */
		fprintf(vbs, "\"\n");
	}
	fprintf(vbs, "oShellLink.Save\n");
	fclose(vbs);

	/* Run VBScript file. */
	comm[COMMMAX]= 0;
	_snprintf(comm, COMMMAX, "CScript.exe %s", path);
	system(comm);

	return 0;
}


int main(int argc, char** argv)
{
	int i;
	char* lnkPath= NULL;
	char* targetPath= NULL;
	char* iconPath= NULL;
	char* workingDirectory= NULL;
	char* arguments= NULL;

	if(argc < 2)
		usage();

	if(strncmp(argv[1], "--help", 7) == 0)
	{
		usage();
	}

	if(strncmp(argv[1], "--license", 10) == 0)
	{
		license();
	}

	if((argc < 5) || (argc > 11) || (argc % 2 != 1))
	{
		usage();
	}

	for(i= 1; i < argc - 1; i++)
	{
		if(strncmp(argv[i], "--lnk", 6) == 0)
		{
			lnkPath= argv[i + 1];
			i++;
		}

		else if(strncmp(argv[i], "--target", 9) == 0)
		{
			targetPath= argv[i + 1];
			i++;
		}

		else if(strncmp(argv[i], "--icon", 7) == 0)
		{
			iconPath= argv[i + 1];
			i++;
		}

		else if(strncmp(argv[i], "--work", 7) == 0)
		{
			workingDirectory= argv[i + 1];
			i++;
		}

		else if(strncmp(argv[i], "--args", 7) == 0)
		{
			arguments= argv[i + 1];
			i++;
		}

		else
			usage();
	}

	if((lnkPath == NULL) || (targetPath == NULL))
	{
		fprintf(stderr, "makescut: LNK_PATH and TARGET_PATH are required.\n");
		usage();
	}

	makescut(lnkPath, targetPath, iconPath, workingDirectory, arguments);

	printf("\n");
	return 0;
}