# eval-hta
VBScript REPL circa 2006, now with added JScript.

## Usage

The following objects exist for you to play with:

	| Name  | Type                       |
	|-------|----------------------------|
	| wmi   | WMI COM interface          |
	| fso   | Scripting.FileSystemObject |
	| shell | WScript.Shell              |
	| net   | WScript.Network            |
	| dict  | Scripting.Dictionary       |

There are custom functions to echo to the console, copy to the clipboard, and
clear the screen. In VBScript, these can be used without parentheses because...
VBScript. There are also commands to switch between JScript and VBScript and to
quit the interpreter. To see all of this, type help (or help() in JScript).
Output:

	' Functions:
	' Clip() - Copy to clipboard
	' Echo() - Print to console
	' Commands:
	' cls - Clear screen
	' jscript - Switch to JScript
	' vbscript - Switch to VBScript
	' exit/quit/bye Terminate

Output is meant to be easily pasted to a WSF or other file for convenience.
