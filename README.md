# Essbase.Pandora 
English description will be soon )

пример запуска 

 cscript svStressTool_exe.vbs RU3ACTIV.RUN

 ------------------
Данный набор скриптов предназначен для проведения стресс-тестирования APS и  Essbase. В процесе работы создается шторм подключений ,  запросов MDX, расчетов. Разработка создана на VBScript с использованием SmartView HTML XML API.

Для того что бы начать работу потребуется 
 1) MDX скрипты (RU3ACTIVХХ.MDX)
 2) Файл авторизации  (RU3ACTIV.AUT)
 3) Сценарий конкретной сессии (RU3ACTIV.scn01)
 4) Файл запуска сценариев (RU3ACTIV.RUN)
 
 
 a) MDX с крипты можно получить включив  аудит в essbase.cfg ( они должны быть размещены в директории MDX. )
 
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
 
b) Файл авторизации состоит из указаний настройки подключения и логина с паролем  Например 

APS=http://serverAPS:13080/aps/SmartView

ESB=serverEssBase:1424

LOG=system

PAS=password

APP=sample

DBS=basic

 ------------------
 c) Файл сценария работы сессии состоит из команд , подключения , запуска MDX (указывается файл , в котором хранится запрос), запуска расчетов 
 паузы. Строчки котороые начинаются с # считаются комментарифми и игнорируются . Например 
 
SLEEP=1

CON=RU3ACTIV.AUT

SLEEP=1

MDX=RU3ACTIV_01.MDX

SLEEP=1

MDX=RU3ACTIV_02.MDX

#CSC=ACT_CALC_ALL

 ------------------

в) командный файл запуска сценариев. Указывается имя файл сценария, количество запусков, ждать завершения (sync), тайм аут для следующего запуска 
Например 

#ScenarioName;times;mode;delay

RU3ACTIV.scn01;1;sunc;2

RU3ACTIV.scn01;1;asunc;2

  
