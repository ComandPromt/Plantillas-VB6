System Monitor by Levi Lansing
                  puzzlesome@hotmail.com


System Monitor is a fairly simple program that uses only
API to monitor the CPU usage (%), total free memory (%),
and the amount of RAM free (% and MB) in small bar
graphs.


oops, made a little mistake last time- i accidenally used
the same file name as windows' system monitor, so it
accidently overwrote it- this one will save to sysmon2.exe
if you choose the autorun option. if you lost your sysmon.exe
program for windows, you can either re-install it, or just
get it from a friends computer. sorry about that.

Sorry about the small bug in the first version as well- the
2 memory graphs were not tested properly and did not
function properly, though the actual values were correct.


The API this program uses to get the CPU usage is
derived from that of the system monitor that comes
with windows.  I am not absolutely sure all the
registry functions are necessary, but they are
exactly as window's system monitor uses them and
work perfect thus far.


**Note** this monitor code only works for windows
versions 95, 98, and ME.  The CPU usage does NOT work
for win 2K, though the other 2 bar graphs *should*
this code has not been tested in any other operating
systems.