# BKDRCore v0.1
Unofficial native Windows x86 control library and CLI for BK DR programmable DC power supplies. Initial target: DR2003-MO firmware 1.16. Windows 7 x86 is a compatibility target, not yet verified. USB virtual COM only; Ethernet and remote-to-local are pending.

非官方 BK DR 系列可程式電源控制函式庫。初版針對 DR2003-MO，原生 Win32 DLL、CLI，靜態 CRT；預留 Ethernet。尚未實機驗證編譯後的 DLL。

## Build / 編譯
Run `build.bat` in an x86 Visual Studio developer command prompt, or push this directory (including `.github`) to GitHub and download the Actions artifact. Keep DLL and EXE together. Win7 compatibility must be checked on the target machine.

## CLI / 操作
Disconnect loads and verify output OFF before initial testing. Close AccessPort first.
```
BKDRCLI --port COM14 id
BKDRCLI --port COM14 snapshot
BKDRCLI --port COM14 off
BKDRCLI --port COM14 setvi 5 0.0013
BKDRCLI --port COM14 snapshot
BKDRCLI --port COM14 record 100 5 > data.csv
BKDRCLI --port COM14 query 1 SYS:LED?
BKDRCLI --port COM14 local
```
`on --confirm` explicitly enables output; closing CLI does NOT disable output. `off` verifies state. `local` returns unsupported; stop polling and press decimal key on the panel. Never assume loss of communication disables output. Raw DLL commands bypass typed validation; only trusted callers should use them. Never use software as the sole safety interlock.

## API
See BKDRCore.h. stdcall, undecorated exports, 32-bit VBA compatible; doubles in V/A, pointers to caller-owned buffers or doubles. Single session, caller must serialize calls. Return 0 on success, negative otherwise (-1 connection, -2 I/O/desync, -3 timeout, -4 invalid input/buffer, -5 invalid response/readback, -6 output active, -7 unsupported). Reconnect after timeout; verify panel independently. SafeClose attempts output off then closes, reports failure if unverified; it cannot release RMT. Close does not alter output. Typed setters require output OFF and conservatively limit 0–70 V, 0.5 mA–3 A; not a full model/range profile. SetVI is sequential, not atomic. Snapshot values are sequential, not simultaneous.

## Observed protocol / 實測
57600 8N1, LF transmit, CRLF receive, NUL in VER? response. V/I query pair ~77 ms; 1000 repetitions without errors. Setpoints: 1 mV, 0.1 mA observed granularity; minimum current 0.5 mA. These are not output accuracy or ADC update rate. Mixed set/query ~148 ms for V/I, before extra safety checks. Output transient, LED mode, CV/CC, Ethernet and software LOCAL remain unverified. No raw device serial numbers included.
