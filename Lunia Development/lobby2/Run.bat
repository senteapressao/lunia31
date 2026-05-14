:start


@ConsoleServer.exe LobbyConfig.xml

@IF %errorlevel% EQU 0 GOTO end

@GOTO start

:err
@ECHO UNABLE TO FIND FILE : %1

:end
@ECHO SERVICE STOPPED
@PAUSE
