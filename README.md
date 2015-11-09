# IsTagged 
Version 1.0
November 5, 2015
Copyright 2012 - 2015 by Jamal Mazrui
GNU Lesser General Public License (LGPL)

* * *

## Description

IsTagged is a stand-alone, 32-bit, Windows console-mode utility that checks the Tagged status of one ore more PDF documents.  If the Tagged status of a PDF has been set by its author, it means that tags have been added to the document to enable accessibility for persons with disabilities using assistive technology, e.g., a blind person using a screen reader program to operate the computer. Note that a positive Tagged status does not necessarily mean that a PDF is accessible, since tags may not have been applied correctly. A negative Tagged status, however, is probably a flag for inaccessibility. Among other sources, more information on Tagged PDF is at
http://www.planetpdf.com/enterprise/article.asp?ContentID=6067

## Operation

IsTagged.exe takes a single parameter on the command line indicating what file(s) to check. For example, to check a single file, run `IsTagged test.pdf`.

To check multiple files, provide a wildcard specification, e.g.,
`IsTagged c:\temp\*.pdf`.

Since Windows file paths are not case sensitive, a command may be input in all lower case for easier typing.

For each file checked, the program reports its status with a line of console output. A single character indicates the Tagged status: Y for yes, N for No, or O for Other -- meaning the file is something other than a valid PDF.  After the status character is a single space, followed by the full path of the file. For example, the console output might be as follows:

    Checking Tagged Status of *.*
    O c:\temp\source.docx
    N c:\temp\initial.pdf
    Y c:\temp\final.pdf

The program output may be captured in a text file using standard redirection, e.g., `IsTagged c:\temp\*.pdf >report.txt`.

The output file may be examined manually or parsed by another program, e.g., to extract lines that begin with the N status.

* * *

## Development Notes

The latest version of IsTagged is available at http://EmpowermentZone.com/IsTagged.zip

This documentation is also available online at [Empowerment Zone](http://EmpowermentZone.com/IsTagged.htm)

The IsTagged.bas file contains the main source code for the program, built with the [PowerBASIC compiler](http://PowerBASIC.com)

The [QuickPDF library](http://QuickPDFLibrary.com) is used by the program.

Note that these are commercial products needed to successfully compile a new version of the executable. The source code, however, is open and free to share according to the GNU Lesser General Public License (LGPL).  IsTagged stores temporary files in the user's temp directory, `IsTagged_temp.dll` and `IsTagged_temp.tlb`, which are deleted at the end of program execution.

I welcome feedback, which helps IsTagged improve over time. When reporting a problem, the more specifics, the better, including steps to reproduce the problem if possible.
