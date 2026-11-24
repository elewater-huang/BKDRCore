#define WIN32_LEAN_AND_MEAN
#include <windows.h>
#include "BKDRCore.h"
#include <string>
#include <sstream>
#include <iomanip>
#include <locale>
#include <cmath>
#include <cstring>
#include <cstdio>
#include <cctype>
static HANDLE h=INVALID_HANDLE_VALUE;
static std::string err;
static bool desync=false;
static int fail(int n,const char* s){err=s;return n;}
static int copy(const std::string& s,char* p,int n){if(!p||n<=0||s.size()>=(size_t)n)return fail(-4,"Buffer too small");memcpy(p,s.c_str(),s.size()+1);return 0;}
static bool valid(const char* s){return s&&*s&&!strchr(s,'\n')&&!strchr(s,'\r')&&strlen(s)<=1024;}
int BK BKDR_OpenSerial(const char* port){if(h!=INVALID_HANDLE_VALUE)return fail(-1,"Already connected");if(!port)return fail(-4,"Missing port");std::string p=port; if(p.empty())return fail(-4,"Missing port");if(isdigit((unsigned char)p[0]))p="COM"+p; for(char& c:p)c=(char)toupper((unsigned char)c);p="\\\\.\\"+p;h=CreateFileA(p.c_str(),GENERIC_READ|GENERIC_WRITE,0,0,OPEN_EXISTING,0,0);if(h==INVALID_HANDLE_VALUE)return fail(-1,"Cannot open COM port");DCB d={};d.DCBlength=sizeof(d);if(!GetCommState(h,&d)){BKDR_Close();return fail(-1,"GetCommState failed");}d.BaudRate=57600;d.ByteSize=8;d.Parity=NOPARITY;d.StopBits=ONESTOPBIT;d.fBinary=TRUE;d.fParity=FALSE;d.fOutxCtsFlow=FALSE;d.fOutxDsrFlow=FALSE;d.fDtrControl=DTR_CONTROL_DISABLE;d.fRtsControl=RTS_CONTROL_DISABLE;d.fOutX=FALSE;d.fInX=FALSE;d.fDsrSensitivity=FALSE;d.fAbortOnError=FALSE;COMMTIMEOUTS t={};t.ReadIntervalTimeout=MAXDWORD;t.ReadTotalTimeoutConstant=20;t.WriteTotalTimeoutConstant=1000;if(!SetCommState(h,&d)||!SetCommTimeouts(h,&t)){BKDR_Close();return fail(-1,"Serial configuration failed");}PurgeComm(h,PURGE_RXCLEAR|PURGE_TXCLEAR);desync=false;err.clear();return 0;}
int BK BKDR_Close(){if(h!=INVALID_HANDLE_VALUE){CloseHandle(h);h=INVALID_HANDLE_VALUE;}desync=false;return 0;}
static int send(const char* cmd){if(h==INVALID_HANDLE_VALUE)return fail(-1,"Not connected");if(!valid(cmd))return fail(-4,"Invalid command");std::string s=std::string(cmd)+"\n";DWORD n=0;if(!WriteFile(h,s.data(),(DWORD)s.size(),&n,0)||n!=s.size()){desync=true;return fail(-2,"Write failed");}return 0;}
int BK BKDR_Write(const char* cmd){if(desync)return fail(-2,"Reconnect required");if(cmd&&strchr(cmd,'?'))return fail(-4,"Use Query for queries");int r=send(cmd);if(!r)Sleep(80);return r;}
int BK BKDR_Query(const char* cmd,int lines,char* out,int cap){if(desync)return fail(-2,"Reconnect required");if(lines<1||lines>32||!out||cap<1)return fail(-4,"Invalid query arguments");out[0]=0;int r=send(cmd);if(r)return r;std::string s;int count=0;DWORD start=GetTickCount();while(GetTickCount()-start<2000){char c;DWORD n=0;if(!ReadFile(h,&c,1,&n,0)){desync=true;return fail(-2,"Read failed");}if(!n)continue;if(c=='\n'){++count;if(count==lines)return copy(s,out,cap);s+='\n';}else if(c!='\r'&&c!=0)s+=c;if(s.size()>32768){desync=true;return fail(-3,"Response too long");}}desync=true;return fail(-3,"Timeout or partial response; reconnect required");}
static int nums(const char* cmd,int n,double* a){char b[512];int r=BKDR_Query(cmd,n,b,sizeof(b));if(r)return r;std::istringstream s(b);s.imbue(std::locale::classic());for(int j=0;j<n;j++)if(!(s>>a[j])||!std::isfinite(a[j]))return fail(-5,"Invalid numeric response");std::string extra;if(s>>extra)return fail(-5,"Extra numeric response");return 0;}
static std::string fmt(double x,int p){std::ostringstream s;s.imbue(std::locale::classic());s<<std::fixed<<std::setprecision(p)<<x;return s.str();}
static bool vv(double v){return std::isfinite(v)&&v>=0&&v<=70&&fabs(v*1000-round(v*1000))<1e-6;}
static bool ii(double i){return std::isfinite(i)&&i>=.0005&&i<=3&&fabs(i*10000-round(i*10000))<1e-6;}
int BK BKDR_GetID(char* b,int n){return BKDR_Query("*IDN?",1,b,n);}
int BK BKDR_GetStatus(char* b,int n){return BKDR_Query("STATUS?",1,b,n);}
int BK BKDR_GetVI(double* v,double* i){if(!v||!i)return -4;double a[2];int r=nums("VOUT?;IOUT?",2,a);if(!r){*v=a[0];*i=a[1];}return r;}
int BK BKDR_GetSetVI(double* v,double* i){if(!v||!i)return -4;double a[2];int r=nums("VSET?;ISET?",2,a);if(!r){*v=a[0];*i=a[1];}return r;}
int BK BKDR_GetSnapshot(double* v,double* i,double* sv,double* si){if(!v||!i||!sv||!si)return -4;double a[4];int r=nums("VOUT?;IOUT?;VSET?;ISET?",4,a);if(!r){*v=a[0];*i=a[1];*sv=a[2];*si=a[3];}return r;}
int BK BKDR_GetVoltage(double* v){return v?nums("VOUT?",1,v):-4;}
int BK BKDR_GetCurrent(double* i){return i?nums("IOUT?",1,i):-4;}
int BK BKDR_GetOutput(int* on){if(!on)return -4;char b[32];int r=BKDR_Query("OUT?",1,b,sizeof(b));if(r)return r;if(!strcmp(b,"ON"))*on=1;else if(!strcmp(b,"OFF"))*on=0;else return fail(-5,"Unknown output state");return 0;}
static int off(){int on;int r=BKDR_GetOutput(&on);return r?r:on?fail(-6,"Switch output OFF before setting"):0;}
int BK BKDR_SetVI(double v,double i){if(!vv(v)||!ii(i))return fail(-4,"Outside conservative DR2003 limits or resolution");int r=off();if(r)return r;double a[2];r=nums(("VSET "+fmt(v,3)+";ISET "+fmt(i,4)+";VSET?;ISET?").c_str(),2,a);return r?r:(fabs(a[0]-v)>0.0001||fabs(a[1]-i)>0.00001)?fail(-5,"Setpoint mismatch; output remains OFF"):0;}
int BK BKDR_SetVoltage(double v){if(!vv(v))return -4;int r=off();if(r)return r;double a;r=nums(("VSET "+fmt(v,3)+";VSET?").c_str(),1,&a);return r?r:fabs(a-v)>0.0001?fail(-5,"Voltage mismatch"):0;}
int BK BKDR_SetCurrent(double i){if(!ii(i))return -4;int r=off();if(r)return r;double a;r=nums(("ISET "+fmt(i,4)+";ISET?").c_str(),1,&a);return r?r:fabs(a-i)>0.00001?fail(-5,"Current mismatch"):0;}
int BK BKDR_OutputOn(){int r=BKDR_Write("OUT ON");if(r)return r;int on;r=BKDR_GetOutput(&on);return r?r:on?0:fail(-5,"Output not ON");}
int BK BKDR_OutputOff(){if(desync){send("OUT OFF");return fail(-2,"OFF sent but unverified; check panel");}int r=BKDR_Write("OUT OFF");if(r)return r;int on;r=BKDR_GetOutput(&on);return r?r:!on?0:fail(-5,"Output not OFF");}
int BK BKDR_Local(){return fail(-7,"Not supported; stop queries and press panel decimal key");}
int BK BKDR_OpenTCP(const char*,int){return fail(-7,"TCP not implemented");}
int BK BKDR_SafeClose(){int r=BKDR_OutputOff();BKDR_Close();return r;}
int BK BKDR_GetLastError(char* b,int n){return copy(err,b,n);}
