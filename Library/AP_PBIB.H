#include "AP_GLOB.H"

BYTE TM_check_AWR (char* prg);	//GSC_NAME:TM_check_AWR	GSC_PROTOTYPE:\tm_check_awr.fct
BOOL tproj_ConnectToPWRT ();	//GSC_NAME:tproj_ConnectToPWRT	GSC_PROTOTYPE:\tproj_connecttopwrt.fct
SYSTEMTIME tproj_DifTime (SYSTEMTIME st, long double dwTimeShift);	//GSC_NAME:tproj_DifTime	GSC_PROTOTYPE:\tproj_diftime.fct
BOOL tproj_DisplayIfQualityCodeIsNotOK (char* tagname);	//GSC_NAME:tproj_DisplayIfQualityCodeIsNotOK	GSC_PROTOTYPE:\tproj_displayifqualitycodeisnotok.fct
BOOL tproj_DisplayIfQualityCodeIsOK (char* tagname);	//GSC_NAME:tproj_DisplayIfQualityCodeIsOK	GSC_PROTOTYPE:\tproj_displayifqualitycodeisok.fct
BOOL tproj_LoopInAlarm (char* pageCODE);	//GSC_NAME:tproj_LoopInAlarm	GSC_PROTOTYPE:\tproj_loopinalarm.fct
void tproj_navigation_byCode (long int panel, long int pageCode);	//GSC_NAME:tproj_navigation_byCode	GSC_PROTOTYPE:\tproj_navigation_bycode.fct
void tproj_navigation_byCodeExt (long int pageCode, long int BTNS, WORD MASK1, WORD MASK2);	//GSC_NAME:tproj_navigation_byCodeExt	GSC_PROTOTYPE:\tproj_navigation_bycodeext.fct
void tproj_navigation_byName (long int panel, char* pageName);	//GSC_NAME:tproj_navigation_byName	GSC_PROTOTYPE:\tproj_navigation_byname.fct
char* tproj_ReturnProjectPath ();	//GSC_NAME:tproj_ReturnProjectPath	GSC_PROTOTYPE:\tproj_returnprojectpath.fct
long int tproj_SensorColor (BOOL ah, BOOL wh, BOOL wl, BOOL al, BOOL man_on);	//GSC_NAME:tproj_SensorColor	GSC_PROTOTYPE:\tproj_sensorcolor.fct
long int tproj_SensorColorFullBySoftware (char* TagBaseName);	//GSC_NAME:tproj_SensorColorFullBySoftware	GSC_PROTOTYPE:\tproj_sensorcolorfullbysoftware.fct
void tproj_ShowAnyWindow (char* caption, char* DevModelName);	//GSC_NAME:tproj_ShowAnyWindow	GSC_PROTOTYPE:\tproj_showanywindow.fct
void tproj_ShowArhive (char* object, char* device);	//GSC_NAME:tproj_ShowArhive	GSC_PROTOTYPE:\tproj_showarhive.fct
void tproj_ShowDevice (char* TagBaseName, char* DevModelName);	//GSC_NAME:tproj_ShowDevice	GSC_PROTOTYPE:\tproj_showdevice.fct
void tproj_ShowDeviceSlider (char* TagBaseName, char* DevModelName);	//GSC_NAME:tproj_ShowDeviceSlider	GSC_PROTOTYPE:\tproj_showdeviceslider.fct
void tproj_ShowKeyboardFloat (char* prefix, char* lpszPictureName, char* lpszObjectName, char* Caption);	//GSC_NAME:tproj_ShowKeyboardFloat	GSC_PROTOTYPE:\tproj_showkeyboardfloat.fct
void tproj_ShowPoorWindow (char* TagBaseName, char* DevModelName);	//GSC_NAME:tproj_ShowPoorWindow	GSC_PROTOTYPE:\tproj_showpoorwindow.fct
BOOL tproj_ShowTrend (char* TLG_ARHNAME, char* TagBaseName);	//GSC_NAME:tproj_ShowTrend	GSC_PROTOTYPE:\tproj_showtrend.fct
BOOL tproj_ShowTrendSimple (char* TLG_ARHNAME, char* TagBaseName);	//GSC_NAME:tproj_ShowTrendSimple	GSC_PROTOTYPE:\tproj_showtrendsimple.fct
BOOL tproj_Users_Group ();	//GSC_NAME:tproj_Users_Group	GSC_PROTOTYPE:\tproj_users_group.fct
void tproj_Users_Level ();	//GSC_NAME:tproj_Users_Level	GSC_PROTOTYPE:\tproj_users_level.fct
BOOL tproj_Users_ReadLevel (DWORD level);	//GSC_NAME:tproj_Users_ReadLevel	GSC_PROTOTYPE:\tproj_users_readlevel.fct
void tproj_WriteAckMessageToLog (char* param, char* name, char* codename, char* description);	//GSC_NAME:tproj_WriteAckMessageToLog	GSC_PROTOTYPE:\tproj_writeackmessagetolog.fct
void tproj_WriteControlMessageToLog (char* param, char* name, char* codename, char* pageCODE);	//GSC_NAME:tproj_WriteControlMessageToLog	GSC_PROTOTYPE:\tproj_writecontrolmessagetolog.fct
void tproj_WriteOIMessageToLog (char* name, char* codename, char* razm, char* pageCODE, double lastValue, double newValue);	//GSC_NAME:tproj_WriteOIMessageToLog	GSC_PROTOTYPE:\tproj_writeoimessagetolog.fct
