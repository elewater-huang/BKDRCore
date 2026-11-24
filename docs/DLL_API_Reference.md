# BKDRCore V0.1 — DLL API Reference
# DLL API 完整說明書

版本：0.1｜文件日期：2026-11-24｜目標平台：Windows x86 / Windows 7、32-bit Office
非官方獨立專案。實機開發基準：BK DR2003-MO，韌體 1.16。使用者已確認 CLI 基本指令及 Excel/Access 範例可用；Local 尚未實作。DR2004-MO、Ethernet 未經實機驗證。本文不宣稱每個 API 均獨立完成硬體驗證。

## 1. 安裝與 ABI (Installation and ABI)

從 [Actions](https://github.com/elewater-huang/BKDRCore/actions) 下載 BKDRCore-x86，取得 BKDRCore.dll、BKDRCore.lib、BKDRCore.h、BKDRCLI.exe。DLL 與呼叫程式需同為 x86；64-bit Office/Python 不可載入此 DLL。Windows 7 實機需 CP210x 驅動。AccessPort 或其他程式須釋放 COM Port。使用可信任的絕對 DLL 路徑或已知 DLL 搜尋位置；Office 文件同目錄不必然是 Windows DLL 搜尋路徑。可將下述 VBA Declare 的 Lib 改為實際絕對路徑。

C linkage / __stdcall；公開未修飾函式名稱見 BKDRCore.def。int 為 32-bit signed，double 為 64-bit floating point，char* 為 ANSI/NUL 結尾字串。所有 API 回傳 int：0 成功，負數失敗。指標必須指向有效、足夠容量的呼叫者記憶體；DLL 不配置供呼叫端釋放的緩衝。字串容量以 bytes 計，含末端 NUL。ASCII SCPI 建議使用 4096 bytes 緩衝。禁止傳入無效指標。

DLL 每一程序只有一個全域連線；無 handle-based 多設備 API，無內建執行緒同步。呼叫端必須序列化所有 API，包含計時器、背景取樣及關閉操作。所有呼叫同步阻塞；不提供取消、callback 或自動重連。

## 2. USB protocol / 通訊協定

| Parameter | Value |
|---|---|
| Serial | 57600 baud, 8N1, no flow control |
| TX | ASCII + LF (0A), DLL 自動附加 |
| RX | CRLF observed; DLL strips CR/NUL and uses LF to count lines |
| Command length | 1–1024 bytes, no embedded CR/LF |
| Query response lines | 1–32 |
| Read deadline | about 2000 ms from after send; not configurable |
| Serial write timeout | 1000 ms |
| Write settling wait | 80 ms, no acknowledgement |
| Numeric locale | classic C locale, decimal point |

USB V/I 複合查詢平均約 77 ms；1000 組未發生錯誤。單筆約 63 ms、混合設定 V/I 約 148 ms（不含 DLL 額外 OFF 查詢）。四數值快照實驗單次約 87–94 ms，非長期保證。CLI record 支援 0.01–12 Hz，建議先用 5 Hz；實際取樣由通訊延遲、OS 排程及寫檔決定，並非 ADC 更新頻率。

## 3. Limits and safety / 限制與安全

| Quantity | Typed API constraint |
|---|---|
| Voltage | 0–70 V, increments of 0.001 V |
| Current | 0.0005–3 A, increments of 0.0001 A |
| Setting | OUTPUT OFF required |

上述為 DLL 驗證規則，並非完整設備額定範圍；未檢查檔位、功率包絡、型號、保護閾值或負載。必須自行核對面板檔位、額定值及接線。1 mV / 0.1 mA 為設定解析度而非準確度。實機 ISET 0 或 0.0001 A 會被限制到 0.0005 A，標準 DLL setter 則拒絕。停止輸出使用 OutputOff，不能以 SetCurrent(0) 替代。

**斷線、Close、CLI 結束均不保證停止供電。** OutputOn 失敗仍可能已開啟輸出。SafeClose 只嘗試關閉並確認，不是獨立硬體連鎖。錯誤後應獨立確認面板輸出狀態，必要時使用實體斷電/安全裝置。初次測試空載，維持輸出 OFF。Raw Write/Query 可繞過 setter 的所有檢查；不可將未受信任的字串直接轉交 DLL。OVP、OCP、LED mode、Slew Rate 及暫態輸出尚未完整驗證。

## 4. Return codes / 回傳碼

| Code | Meaning | Handling |
|---:|---|---|
| 0 | Success 成功 | 僅表示該 API 所定義的成功條件 |
| -1 | Connection error 連線錯誤 | 檢查 COM、驅動及占用 |
| -2 | I/O error or desync | 停止輪詢、確認輸出、重新連線 |
| -3 | Timeout / partial / oversized response | 同上，不直接重送下一個查詢 |
| -4 | Invalid argument or buffer | 核對指標、容量、範圍與步進 |
| -5 | Invalid response or readback mismatch | 不使用失敗讀值；確認設定和輸出 |
| -6 | Output active | 先停止輸出並確認，再設定 |
| -7 | Not supported | Local / TCP 尚未實作 |

GetLastError 提供輔助文字，不是每個錯誤都更新，成功也不一定清空。應優先使用函式回傳碼。通訊逾時或 I/O 錯誤會使 DLL 標記失同步，須 Close 後 OpenSerial。重連也不保證裝置沒有延遲回傳，先確認狀態再執行控制。錯誤行數或多餘回傳可能造成未被偵測的失同步。

## 5. Complete API reference / 21 個 API

下列所有宣告的回傳型別均為 int，呼叫慣例均為 __stdcall。0 為成功，其他依第 4 節判定。

### 5.1 BKDR_OpenSerial

```c
int __stdcall BKDR_OpenSerial(const char* port);
```

- **Parameters / 參數：** COM 名稱（COM14、com14 或 14）
- **Behavior / 行為：** 開啟獨占串列埠，設定 57600 8N1，清除收發緩衝；不發送識別指令、不改變輸出。已開啟時 -1。
- **Errors / 回傳：** 連線失敗 -1；空參數 -4。

### 5.2 BKDR_Close

```c
int __stdcall BKDR_Close(void);
```

- **Parameters / 參數：** 無
- **Behavior / 行為：** 釋放串列埠並清除失同步旗標；不關閉輸出、不解除 RMT。
- **Errors / 回傳：** 目前固定回傳 0，不檢查 CloseHandle 結果。

### 5.3 BKDR_Query

```c
int __stdcall BKDR_Query(const char* cmd, int lines, char* out, int capacity);
```

- **Parameters / 參數：** cmd：SCPI 字串；lines：預期回傳行數 1–32；out：可寫緩衝；capacity：位元組數含 NUL
- **Behavior / 行為：** 傳送 cmd + LF，讀到指定 LF 行數即返回；移除 CR、NUL，行間保留 LF，末行無 LF。約 2 秒整體讀取期限。
- **Errors / 回傳：** -1/-2/-3/-4；行數錯誤可能使後續回傳錯位。

### 5.4 BKDR_Write

```c
int __stdcall BKDR_Write(const char* cmd);
```

- **Parameters / 參數：** 不含問號、CR 或 LF 的 SCPI 字串
- **Behavior / 行為：** 傳送後固定等待 80 ms；不讀回、不確認設備接受。
- **Errors / 回傳：** -1/-2/-4；失同步時 -2。

### 5.5 BKDR_GetID

```c
int __stdcall BKDR_GetID(char* out, int capacity);
```

- **Parameters / 參數：** 字串緩衝及容量
- **Behavior / 行為：** *IDN?；回傳設備識別原始文字（經 Query 正規化）。
- **Errors / 回傳：** 同 Query。

### 5.6 BKDR_GetVI

```c
int __stdcall BKDR_GetVI(double* v, double* i);
```

- **Parameters / 參數：** 兩個可寫 double，V、A
- **Behavior / 行為：** VOUT?;IOUT?；兩個數值都解析成功後才更新輸出。
- **Errors / 回傳：** -4 空指標；其他查詢或解析錯誤。

### 5.7 BKDR_GetSetVI

```c
int __stdcall BKDR_GetSetVI(double* v, double* i);
```

- **Parameters / 參數：** 兩個可寫 double，設定 V、A
- **Behavior / 行為：** VSET?;ISET?；讀取設定值，不是實際輸出。
- **Errors / 回傳：** 同 GetVI。

### 5.8 BKDR_GetSnapshot

```c
int __stdcall BKDR_GetSnapshot(double* v, double* i, double* sv, double* si);
```

- **Parameters / 參數：** 四個可寫 double：實際 V、實際 A、設定 V、設定 A
- **Behavior / 行為：** VOUT?;IOUT?;VSET?;ISET?；依順序回傳，非同步硬體取樣。
- **Errors / 回傳：** -4 空指標；查詢或解析錯誤。

### 5.9 BKDR_SetVI

```c
int __stdcall BKDR_SetVI(double v, double i);
```

- **Parameters / 參數：** V、A；符合下述軟體範圍及步進
- **Behavior / 行為：** 先 OUT? 確認 OFF，再 VSET ...;ISET ...;VSET?;ISET?；讀回容差 V 0.0001、A 0.00001。非原子交易，失敗可能已改變部分設定。
- **Errors / 回傳：** -4/-6/-5 或通訊錯誤。

### 5.10 BKDR_SetVoltage

```c
int __stdcall BKDR_SetVoltage(double v);
```

- **Parameters / 參數：** V，0–70，1 mV 步進
- **Behavior / 行為：** 先確認 OFF；VSET ...;VSET?；讀回容差 0.0001 V。
- **Errors / 回傳：** -4/-6/-5 或通訊錯誤。

### 5.11 BKDR_SetCurrent

```c
int __stdcall BKDR_SetCurrent(double i);
```

- **Parameters / 參數：** A，0.0005–3，0.0001 A 步進
- **Behavior / 行為：** 先確認 OFF；ISET ...;ISET?；讀回容差 0.00001 A。
- **Errors / 回傳：** -4/-6/-5 或通訊錯誤。

### 5.12 BKDR_GetVoltage

```c
int __stdcall BKDR_GetVoltage(double* v);
```

- **Parameters / 參數：** 可寫 double，V
- **Behavior / 行為：** VOUT?；實際輸出電壓。
- **Errors / 回傳：** -4 或查詢/解析錯誤；失敗時不得使用輸出。

### 5.13 BKDR_GetCurrent

```c
int __stdcall BKDR_GetCurrent(double* i);
```

- **Parameters / 參數：** 可寫 double，A
- **Behavior / 行為：** IOUT?；實際輸出電流。
- **Errors / 回傳：** 同 GetVoltage。

### 5.14 BKDR_GetOutput

```c
int __stdcall BKDR_GetOutput(int* on);
```

- **Parameters / 參數：** 可寫 32-bit int；0=OFF、1=ON
- **Behavior / 行為：** OUT?；只接受精確 ON/OFF 字串。
- **Errors / 回傳：** -4/-5 或通訊錯誤。

### 5.15 BKDR_OutputOn

```c
int __stdcall BKDR_OutputOn(void);
```

- **Parameters / 參數：** 無
- **Behavior / 行為：** OUT ON，再 OUT? 確認 ON；不檢查負載、設定上限或保護狀態。
- **Errors / 回傳：** -5 或通訊錯誤；失敗不表示一定 OFF。

### 5.16 BKDR_OutputOff

```c
int __stdcall BKDR_OutputOff(void);
```

- **Parameters / 參數：** 無
- **Behavior / 行為：** OUT OFF，再 OUT? 確認 OFF。若已失同步仍嘗試送 OUT OFF，但固定回 -2，須人工確認。
- **Errors / 回傳：** -2/-5 或通訊錯誤。

### 5.17 BKDR_Local

```c
int __stdcall BKDR_Local(void);
```

- **Parameters / 參數：** 無
- **Behavior / 行為：** 預留；不傳送任何指令。停止所有輪詢後手按面板小數點鍵。
- **Errors / 回傳：** 固定 -7。

### 5.18 BKDR_SafeClose

```c
int __stdcall BKDR_SafeClose(void);
```

- **Parameters / 參數：** 無
- **Behavior / 行為：** 呼叫 OutputOff，再 Close；無論 OFF 成敗都關閉串列埠；不解除 RMT。
- **Errors / 回傳：** 回傳 OutputOff 的結果；非硬體安全保證。

### 5.19 BKDR_OpenTCP

```c
int __stdcall BKDR_OpenTCP(const char* host, int port);
```

- **Parameters / 參數：** 預留主機名稱/IP、TCP 埠
- **Behavior / 行為：** 未實作，不建立連線。
- **Errors / 回傳：** 固定 -7。

### 5.20 BKDR_GetStatus

```c
int __stdcall BKDR_GetStatus(char* out, int capacity);
```

- **Parameters / 參數：** 字串緩衝及容量
- **Behavior / 行為：** STATUS?；回傳狀態字串，例如 000021；DLL 不解碼位元。
- **Errors / 回傳：** 同 Query。

### 5.21 BKDR_GetLastError

```c
int __stdcall BKDR_GetLastError(char* out, int capacity);
```

- **Parameters / 參數：** 字串緩衝及容量
- **Behavior / 行為：** 複製 DLL 最近保存的錯誤文字，不發送 SCPI；成功操作未必清除舊錯誤，部分 -4 路徑不更新文字。
- **Errors / 回傳：** 0 或 -4；不是上一個 API 的數字回傳碼。


## 6. C/C++ example / 範例

安全空載讀取與設定範例，不開啟輸出。在 x86 developer command prompt 使用 cl /EHsc /MT example.cpp BKDRCore.lib 編譯；需 header、lib、DLL。

```cpp
#include "BKDRCore.h"
#include <cstdio>
int main() {
    int r = BKDR_OpenSerial("COM14");
    if (r) return 1;
    double v=0, i=0, sv=0, si=0;
    char id[512]={};
    if (!r) r=BKDR_GetID(id,sizeof id);
    if (!r) r=BKDR_OutputOff();
    if (!r) r=BKDR_SetVI(5.0,0.0013);
    if (!r) r=BKDR_GetSnapshot(&v,&i,&sv,&si);
    if (!r) std::printf("%s\n%.3f,%.4f,%.3f,%.4f\n",id,v,i,sv,si);
    if (r) {
        char err[512]={};
        BKDR_GetLastError(err,sizeof err);
        std::fprintf(stderr,"Error %d: %s\n",r,err);
    }
    int end=BKDR_SafeClose();
    if (end) std::fprintf(stderr,"OFF unverified; check panel (%d)\n",end);
    return (r||end)?1:0;
}
```

## 7. Excel / Access VBA (32-bit Office)

在標準模組加入以下完整宣告。Office 2003 不使用 PtrSafe；64-bit Office 不能以修改宣告方式載入 x86 DLL。

```vb
Option Explicit
Public Declare Function BKDR_OpenSerial Lib "BKDRCore.dll" (ByVal port As String) As Long
Public Declare Function BKDR_Close Lib "BKDRCore.dll" () As Long
Public Declare Function BKDR_Query Lib "BKDRCore.dll" (ByVal cmd As String, ByVal lines As Long, ByVal buf As String, ByVal capacity As Long) As Long
Public Declare Function BKDR_Write Lib "BKDRCore.dll" (ByVal cmd As String) As Long
Public Declare Function BKDR_GetID Lib "BKDRCore.dll" (ByVal buf As String, ByVal capacity As Long) As Long
Public Declare Function BKDR_GetVI Lib "BKDRCore.dll" (ByRef v As Double, ByRef i As Double) As Long
Public Declare Function BKDR_GetSetVI Lib "BKDRCore.dll" (ByRef v As Double, ByRef i As Double) As Long
Public Declare Function BKDR_GetSnapshot Lib "BKDRCore.dll" (ByRef v As Double, ByRef i As Double, ByRef sv As Double, ByRef si As Double) As Long
Public Declare Function BKDR_SetVI Lib "BKDRCore.dll" (ByVal v As Double, ByVal i As Double) As Long
Public Declare Function BKDR_SetVoltage Lib "BKDRCore.dll" (ByVal v As Double) As Long
Public Declare Function BKDR_SetCurrent Lib "BKDRCore.dll" (ByVal i As Double) As Long
Public Declare Function BKDR_GetVoltage Lib "BKDRCore.dll" (ByRef v As Double) As Long
Public Declare Function BKDR_GetCurrent Lib "BKDRCore.dll" (ByRef i As Double) As Long
Public Declare Function BKDR_GetOutput Lib "BKDRCore.dll" (ByRef onState As Long) As Long
Public Declare Function BKDR_OutputOn Lib "BKDRCore.dll" () As Long
Public Declare Function BKDR_OutputOff Lib "BKDRCore.dll" () As Long
Public Declare Function BKDR_Local Lib "BKDRCore.dll" () As Long
Public Declare Function BKDR_SafeClose Lib "BKDRCore.dll" () As Long
Public Declare Function BKDR_OpenTCP Lib "BKDRCore.dll" (ByVal host As String, ByVal port As Long) As Long
Public Declare Function BKDR_GetStatus Lib "BKDRCore.dll" (ByVal buf As String, ByVal capacity As Long) As Long
Public Declare Function BKDR_GetLastError Lib "BKDRCore.dll" (ByVal buf As String, ByVal capacity As Long) As Long

Public Function BKDRText(ByVal buf As String) As String
    Dim p As Long
    p = InStr(1, buf, vbNullChar)
    If p > 0 Then
        BKDRText = Left$(buf, p - 1)
    Else
        BKDRText = buf
    End If
End Function

Public Sub BKDRDemo()
    Dim rc As Long, endRc As Long
    Dim opened As Boolean
    Dim v As Double, i As Double, sv As Double, si As Double
    Dim buf As String
    On Error GoTo Failed
    rc = BKDR_OpenSerial("COM14")
    If rc <> 0 Then GoTo Failed
    opened = True
    buf = String$(512, vbNullChar)
    rc = BKDR_GetID(buf, Len(buf))
    If rc <> 0 Then GoTo Failed
    Debug.Print BKDRText(buf)
    rc = BKDR_OutputOff()
    If rc <> 0 Then GoTo Failed
    rc = BKDR_SetVI(5#, 0.0013)
    If rc <> 0 Then GoTo Failed
    rc = BKDR_GetSnapshot(v, i, sv, si)
    If rc <> 0 Then GoTo Failed
    Debug.Print v, i, sv, si
    GoTo Cleanup
Failed:
    Debug.Print "Error:", rc, Err.Number, Err.Description
Cleanup:
    On Error GoTo 0
    If opened Then
        endRc = BKDR_SafeClose()
        If endRc <> 0 Then MsgBox "Output OFF unverified. Check panel."
    End If
End Sub
```

Excel 可將 v、i、sv、si 寫入儲存格；Access 可寫入表單控制項及 DAO 資料表。功率 W = V × A。取樣使用 Application.OnTime 或 Form.Timer，應防止重入，先取消計時器再關閉連線。工作簿/表單異常退出仍不能保證 OFF。範例不會主動 OutputOn。

## 8. Python 3.7 x86 ctypes

使用 32-bit Python；WinDLL 對應 stdcall。不需 pyserial。

```python
import ctypes as c
from pathlib import Path
dll = c.WinDLL(str(Path("BKDRCore.dll").resolve()))
D = c.c_double
P = c.POINTER(D)
dll.BKDR_OpenSerial.argtypes = [c.c_char_p]
dll.BKDR_OpenSerial.restype = c.c_int
dll.BKDR_GetSnapshot.argtypes = [P, P, P, P]
dll.BKDR_GetSnapshot.restype = c.c_int
dll.BKDR_OutputOff.argtypes = []
dll.BKDR_OutputOff.restype = c.c_int
dll.BKDR_SetVI.argtypes = [D, D]
dll.BKDR_SetVI.restype = c.c_int
dll.BKDR_SafeClose.argtypes = []
dll.BKDR_SafeClose.restype = c.c_int

def check(rc):
    if rc != 0:
        raise RuntimeError("BKDR error {}".format(rc))

check(dll.BKDR_OpenSerial(b"COM14"))
try:
    check(dll.BKDR_OutputOff())
    check(dll.BKDR_SetVI(5.0, 0.0013))
    values = [D() for _ in range(4)]
    check(dll.BKDR_GetSnapshot(*(c.byref(x) for x in values)))
    print([x.value for x in values])
finally:
    rc = dll.BKDR_SafeClose()
    if rc:
        print("OFF unverified; check panel:", rc)
```

其他 API 請依第 5 節設定 argtypes/restype；可寫字串使用 create_string_buffer 與容量，GetOutput 使用 POINTER(c.c_int)。

## 9. CLI quick reference

```bat
BKDRCLI.exe --port COM14 id
BKDRCLI.exe --port COM14 off
BKDRCLI.exe --port COM14 setvi 5 0.0013
BKDRCLI.exe --port COM14 snapshot
BKDRCLI.exe --port COM14 status
BKDRCLI.exe --port COM14 query 1 "SYS:LED?"
BKDRCLI.exe --port COM14 record 100 5 > data.csv
BKDRCLI.exe --port COM14 local
```

snapshot 欄位依序為實際 V、實際 A、設定 V、設定 A。record CSV 標頭 elapsed_ms,voltage_V,current_A；COUNT 1–1000000 整數，HZ 0.01–12，可小數。elapsed_ms 從該次 record 開始計時，第一筆包含查詢耗時。CLI query 僅接受含問號且不含分號的字串；DLL Query 支援多行複合指令。on --confirm 可啟動輸出，只在核對負載與保護條件後使用。CLI 成功 exit 0，操作失敗 1，基本參數格式錯誤 2。每次 CLI 呼叫結束只執行 Close；record 結束不會自動 OFF。local 預期失敗（-7）。

## 10. Raw SCPI and troubleshooting

| SCPI | Expected lines | Purpose |
|---|---:|---|
| *IDN? | 1 | identification; HP E3634A compatibility string observed |
| VER? | 1 | firmware; embedded NUL removed |
| VOUT?;IOUT? | 2 | measured V/A |
| VSET?;ISET? | 2 | setpoints |
| OUT?;STATUS? | 2 | output and raw status |
| PROT:OVP? | 1 | observed OFF |
| PROT:OCP? | 1 | observed ON |
| OUT:SR:VOLT? | 1 | observed 7.000; units not established here |
| OUT:SR:CURR? | 1 | observed 0.6000; units not established here |
| SYS:ERR? | 1 | device error; unsupported commands may still report 0 |

SYS:LED? 可作為後續探索指令，本文未將其列為已驗證。不要把 OUT:MODE? 或 LED? 當作已支援指令（實測 timeout）。Raw Query 必須提供正確回傳行數；不得使用無回傳指令配合 Query。不要在同一連線上交錯執行其他查詢。

### RMT stays on
停止所有背景查詢，確認輸出處於預期安全狀態，Close，按面板小數點鍵解除 RMT。已實測拔 USB 或關 COM 不解除 RMT。Local API 暫回 -7。不要使用 *RST 代替 Local。

### Cannot load DLL / COM
確認 DLL 及呼叫程式位元數均為 32-bit、檔案路徑與驅動正常、COM 未被其他程式占用。VBA 的 DLL 找不到問題應優先檢查 Lib 絕對路徑，不能只假設資料庫目錄會被搜尋。

## 11. Release scope / 版本範圍

本階段結案：USB SCPI、Win32 DLL、CLI、Excel/Access 範例及 API 文件。不追加 GUI。TCP、軟體解除 RMT、實際輸出準確度、CV/CC、LED Mode、Slew Rate、保護功能完整驗證均不屬於 V0.1 已完成範圍。

## Source references / 文件依據

- [BKDRCore.h](../BKDRCore.h) — 0ed6d16d0ff8586eb88600b2822d589534bbf48a
- [BKDRCore.cpp](../BKDRCore.cpp) — 6da62298c3cbe3969e0785641f5e7faba5de0acf
- [BKDRCore.def](../BKDRCore.def) — 577f1368f3bb652f3a4ad7ed299cb907f577eebc
- [BKDRCLI.cpp](../BKDRCLI.cpp) — d97d03f656b668d9054216db4ad045b9bb8df85b
- 使用者提供之三輪 USB 實測報告及 CLI / Office 範例驗證結果。
