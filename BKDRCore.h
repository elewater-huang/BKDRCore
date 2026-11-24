#pragma once
#ifdef __cplusplus
extern "C" {
#endif
#define BK __stdcall
/* Single connection; calls serialized by caller. 0=success, negative=error. */
int BK BKDR_OpenSerial(const char* port);
int BK BKDR_Close(void);
int BK BKDR_Query(const char* cmd,int lines,char* out,int capacity);
int BK BKDR_Write(const char* cmd);
int BK BKDR_GetID(char* out,int capacity);
int BK BKDR_GetVI(double* v,double* i);
int BK BKDR_GetSetVI(double* v,double* i);
int BK BKDR_GetSnapshot(double* v,double* i,double* sv,double* si);
int BK BKDR_SetVI(double v,double i);
int BK BKDR_SetVoltage(double v);
int BK BKDR_SetCurrent(double i);
int BK BKDR_GetVoltage(double* v);
int BK BKDR_GetCurrent(double* i);
int BK BKDR_GetOutput(int* on);
int BK BKDR_OutputOn(void);
int BK BKDR_OutputOff(void);
int BK BKDR_Local(void);
int BK BKDR_SafeClose(void);
int BK BKDR_OpenTCP(const char* host,int port);
int BK BKDR_GetStatus(char* out,int capacity);
int BK BKDR_GetLastError(char* out,int capacity);
#ifdef __cplusplus
}
#endif
