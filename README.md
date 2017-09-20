# Essbase.Pandora 
 

start-up example

 cscript svStressTool_exe.vbs RU3ACTIV.RUN

 ------------------
This set of scripts is designed to create stress test of APS and Essbase. This tool process a storm of connections, MDX requests, and calculations. It is developed in VBScript using the SmartView HTML XML API.

To get started, you need

  1) MDX scripts (RU3ACTIVХХ.MDX)
  
  2) Authorization file (RU3ACTIV.AUT)
  
  3) Scenario of a specific session (RU3ACTIV.scn01)
  
  4) Script startup file (RU3ACTIV.RUN)
  
  5) Essbase Calculation scripts 
 
 
a) MDX scripts (RU3ACTIVХХ.MDX) can be obtained by including audit in essbase.cfg (after creation they should be placed in the MDX directory.)
 
 ------------------
;TRACE_REPORT [<appname> [<dbname>] ] <number>
;Where the optional <appname> and <dbname> limit the applications and/or databases this feature is enabled for, and the <number> is a bitwise-or/sum of the following flags:
  
;    1 – SYNC – do fflush() after each record; this is useful for crashes, but has more penalty on performance

;    2 – POST – print query elapsed time after its execution (like TRACE_MDX does); refer to same query by the thread ID ;

;    4 – MDX – trace MDX reports

;    8 – GRID – trace Grid API reports

;    16 – DDB – trace queries coming via partition to the source

;    32 – FULL – trace MDX, GRID and DDB for performance analysis (somewhat heavy!)

;    64 – FULL_MEMBERS – traces MDX, GRID, DDB and member names for performance analysis (very heavy!)

  ------------------
 
b) The authorization file  (RU3ACTIV.AUT)  consists of instructions for setting up a connection and login with a password. For example

APS=http://serverAPS:13080/aps/SmartView

ESB=serverEssBase:1424

LOG=system

PAS=password

APP=sample

DBS=basic

 ------------------
c) The session script file (RU3ACTIV.scn01) consists of commands, connection, MDX start, calculations and
  pause. The lines that begin with "#" symbol are considered as comments and will be ignored. For example
 
SLEEP=1

CON=RU3ACTIV.AUT

SLEEP=1

MDX=RU3ACTIV_01.MDX

SLEEP=1

MDX=RU3ACTIV_02.MDX

#CSC=ACT_CALC_ALL

 ------------------

c) Command file (RU3ACTIV.RUN) for running scripts. Specify the name of the script file, the number of starts, wait for completion (sync), timeout for the next run
For example

#ScenarioName;times;mode;delay

RU3ACTIV.scn01;1;sync;2

RU3ACTIV.scn01;1;async;2

  
