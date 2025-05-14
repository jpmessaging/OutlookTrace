<#
.NOTES
Copyright (c) 2021 Ryusuke Fujita

This software is released under the MIT License.
http://opensource.org/licenses/mit-license.php

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
#>

$Version = 'v2025-05-06'
#Requires -Version 3.0

# Outlook's ETW pvoviders
$Outlook2016Providers =
@'
"{9efff48f-728d-45a1-8001-536349a2db37}" 0xFFFFFFFFFFFFFFFF 64
"{f50d9315-e17e-43c1-8370-3edf6cc057be}" 0xFFFFFFFFFFFFFFFF 64
"{02fd33df-f746-4a10-93a0-2bc6273bc8e4}" 0x00000015         64
"{b866d7ae-7c99-4c20-aa98-278fc044fb98}" 0xFFFFFFFFFFFFFFFF 64
"{35150b03-b441-412b-ace4-291895289743}" 0xFFFFFFFFFFFFFFFF 64
"{d6dd4818-7123-4abb-ad96-b044c1387b49}" 0xFFFFFFFFFFFFFFFF 64
"{5457acd3-0d66-4e53-b9e0-3be59f8b9f9d}" 0xFFFFFFFFFFFFFFFF 64
"{02cac15f-d4be-400e-9127-d54982aa4ae9}" 0xFFFFFFFFFFFFFFFF 64
"{d8d0510d-3f14-4da9-a096-b9c7ad386da0}" 0xFFFFFFFFFFFFFFFF 64
"{aa8fa310-0939-4ce3-b9bb-ae05b2695110}" 0xFFFFFFFFFFFFFFFF 64
"{31c1f514-1937-40ce-b0bf-2db7cb1b6a17}" 0xFFFFFFFFFFFFFFFF 64
"{6b6b571b-f4e3-4fbb-a83f-0790d11d19ab}" 0xFFFFFFFFFFFFFFFF 64
"{c911b508-e06d-4f76-8835-ea1b78e2f66d}" 0xFFFFFFFFFFFFFFFF 64
"{f762ce39-ac6c-4e1c-b55f-0e11586e6d07}" 0xFFFFFFFFFFFFFFFF 64
"{691e1c12-2693-4d4a-852c-7478657bbe6e}" 0xFFFFFFFFFFFFFFFF 64
"{081f51e8-2528-44af-ad0b-6e2e5c7242ad}" 0xFFFFFFFFFFFFFFFF 64
"{284b8d30-4aa6-4a0f-9143-ce2e8e1f10f0}" 0xFFFFFFFFFFFFFFFF 64
"{265f23e0-615d-4082-8e17-ddcd7e6f7eb4}" 0xFFFFFFFFFFFFFFFF 64
"{11adbd74-7df2-4e8e-802b-b3bcbfd04a78}" 0xFFFFFFFFFFFFFFFF 64
"{287bf315-5a11-4b2f-b069-b761ade25a49}" 0xFFFFFFFFFFFFF7FF 64
"{0dae1c38-7bfb-4960-8ea5-54139b54b751}" 0xFFFFFFFFFFFFFFFF 64
"{13967ee5-6b23-4bcd-a496-1d788449a8cf}" 0xFFFFFFFFFFFFFFFF 64
"{31b56255-5883-4f3e-8350-d7d6d88a4908}" 0xFFFFFFFFFFFFFFFF 64
"{059b2f1f-fc6d-4236-8c06-4357a91b17a1}" 0xFFFFFFFFFFFFFFFF 64
"{ad58872e-4df6-4b26-9841-5d7887a1c7a5}" 0xFFFFFFFFFFFFFFFF 64
"{03b1de06-84f4-4fa7-ba4c-cc1b82b56004}" 0xFFFFFFFFFFFFFFFF 64
"{daf0b914-9c1c-450a-81b2-fea7244f6ffa}" 0xFFFFFFFFFFFFFFFF 64
"{bb00e856-a12f-4ab7-b2c8-4e80caea5b07}" 0xFFFFFFFFFFFFFFFF 64
"{a1b69d49-2195-4f59-9d33-bdf30c0fe473}" 0xFFFFFFFFFFFFFFFF 64
"{b4f150b4-67db-4742-8846-2cd7b16ee60e}" 0xFFFFFFFFFFFFFFFF 64
"{8736922d-e8b2-47eb-8564-23e77e728cf3}" 0x00000414         64
'@

$Outlook2013Providers =
@'
"{284b8d30-4aa6-4a0f-9143-ce2e8e1f10f0}" 0xFFFFFFFFFFFFFFFF 64
"{02cac15f-d4be-400e-9127-d54982aa4ae9}" 0xFFFFFFFFFFFFFFFF 64
"{aa8fa310-0939-4ce3-b9bb-ae05b2695110}" 0xFFFFFFFFFFFFFFFF 64
"{6b6b571b-f4e3-4fbb-a83f-0790d11d19ab}" 0xFFFFFFFFFFFFFFFF 64
"{c911b508-e06d-4f76-8835-ea1b78e2f66d}" 0xFFFFFFFFFFFFFFFF 64
"{f762ce39-ac6c-4e1c-b55f-0e11586e6d07}" 0xFFFFFFFFFFFFFFFF 64
"{691e1c12-2693-4d4a-852c-7478657bbe6e}" 0xFFFFFFFFFFFFFFFF 64
"{11adbd74-7df2-4e8e-802b-b3bcbfd04a78}" 0xFFFFFFFFFFFFFFFF 64
"{287bf315-5a11-4b2f-b069-b761ade25a49}" 0xFFFFFFFFFFFFFFFF 64
"{265f23e0-615d-4082-8e17-ddcd7e6f7eb4}" 0xFFFFFFFFFFFFFFFF 64
"{31c1f514-1937-40ce-b0bf-2db7cb1b6a17}" 0xFFFFFFFFFFFFFFFF 64
"{d8d0510d-3f14-4da9-a096-b9c7ad386da0}" 0xFFFFFFFFFFFFFFFF 64
"{b9522d9f-e2cd-44d4-b567-0d5182060e55}" 0xFFFFFFFFFFFFFFFF 64
"{96991e14-71db-4799-a66c-270004757fd8}" 0xFFFFFFFFFFFFFFFF 64
"{8736922d-e8b2-47eb-8564-23e77e728cf3}" 0x0000014         64
"{464a42fb-36bd-4749-a67c-02138387138c}" 0xFFFFFFFFFFFFFFFF 64
"{02fd33df-f746-4a10-93a0-2bc6273bc8e4}" 0xFFFFFFFFFFFFFFFF 64
'@

$Outlook2010Providers =
@'
"{f94cbe33-31c2-492d-9bf8-573beff84c94}" 0x0FB7FFEF 64
"{e3c8312d-b20c-4831-995e-5ec5f5522215}" 0x00124586 64
'@

$WamProviders =
@'
{2A3C6602-411E-4DC6-B138-EA19D64F5BBA} 0xFFFF 0xff
{EF98103D-8D3A-4BEF-9DF2-2156563E64FA} 0xFFFF 0xff
{FB6A424F-B5D6-4329-B9B5-A975B3A93EAD} 0x000003FF
{D93FE84A-795E-4608-80EC-CE29A96C8658} 0x7FFFFFFF 0xff
{3F8B9EF5-BBD2-4C81-B6C9-DA3CDB72D3C5} 0x7 0xff
{B1108F75-3252-4b66-9239-80FD47E06494} 0x2FF 0xff
{C10B942D-AE1B-4786-BC66-052E5B4BE40E} 0x3FF 0xff
{82c7d3df-434d-44fc-a7cc-453a8075144e} 0x2FF 0xff
{05f02597-fe85-4e67-8542-69567ab8fd4f} 0xFFFFFFFF 0xff
{3C49678C-14AE-47FD-9D3A-4FEF5D796DB9} 0xFFFFFFFF 0xff
{077b8c4a-e425-578d-f1ac-6fdf1220ff68} 0xFFFFFFFF 0xff
{7acf487e-104b-533e-f68a-a7e9b0431edb} 0xFFFFFFFF 0xff
{5836994d-a677-53e7-1389-588ad1420cc5} 0xFFFFFFFF 0xff
{4DE9BC9C-B27A-43C9-8994-0915F1A5E24F} 0xFFFFFFFF 0xff
{bfed9100-35d7-45d4-bfea-6c1d341d4c6b} 0xFFFFFFFF 0xff
{9EBB3B15-B094-41B1-A3B8-0F141B06BADD} 0xFFF 0xff
{6ae51639-98eb-4c04-9b88-9b313abe700f} 0xFFFFFFFF 0xff
{7B79E9B1-DB01-465C-AC8E-97BA9714BDA2} 0xFFFFFFFF 0xff
{86510A0A-FDF4-44FC-B42F-50DD7D77D10D} 0xFFFFFFFF 0xff
{08B15CE7-C9FF-5E64-0D16-66589573C50F} 0xFFFFFF7F 0xff
{63b6c2d2-0440-44de-a674-aa51a251b123} 0xFFFFFFFF 0xff
{4180c4f7-e238-5519-338f-ec214f0b49aa} 0xFFFFFFFF 0xff
{EB65A492-86C0-406A-BACE-9912D595BD69} 0xFFFFFFFF 0xff
{d49918cf-9489-4bf1-9d7b-014d864cf71f} 0xFFFFFFFF 0xff
{5AF52B0D-E633-4ead-828A-4B85B8DAAC2B} 0xFFFF 0xff
{2A6FAF47-5449-4805-89A3-A504F3E221A6} 0xFFFF 0xff
{EC3CA551-21E9-47D0-9742-1195429831BB} 0xFFFFFFFF 0xff
{bb8dd8e5-3650-5ca7-4fea-46f75f152414} 0xFFFFFFFF 0xff
{7fad10b2-2f44-5bb2-1fd5-65d92f9c7290} 0xFFFFFFFF 0xff
{74D91EC4-4680-40D2-A213-45E2D2B95F50} 0xFFFFFFFF 0xff
{556045FD-58C5-4A97-9881-B121F68B79C5} 0xFFFFFFFF 0xff
{5A9ED43F-5126-4596-9034-1DCFEF15CD11} 0xFFFFFFFF 0xff
{F7C77B8D-3E3D-4AA5-A7C5-1DB8B20BD7F0} 0xFFFFFFFF 0xff
'@

$Win32Interop = @'
namespace Win32
{
    using System;
    using System.Runtime.InteropServices;
    using System.Collections.Generic;
    using System.ComponentModel;
    using Microsoft.Win32.SafeHandles;

    public static class Advapi32
    {
        // https://learn.microsoft.com/en-us/windows/win32/api/processthreadsapi/nf-processthreadsapi-openprocesstoken
        [DllImport("advapi32.dll", ExactSpelling = true, SetLastError = true)]
        public static extern bool OpenProcessToken(
            SafeProcessHandle ProcessToken,
            System.Security.Principal.TokenAccessLevels DesiredAccess,
            out IntPtr TokenHandle);

        [DllImport("advapi32.dll", ExactSpelling = true, SetLastError = true)]
        public static extern bool OpenProcessToken(
            SafeProcessHandle ProcessToken,
            System.Security.Principal.TokenAccessLevels DesiredAccess,
            out SafeProcessTokenHandle TokenHandle);

        // https://learn.microsoft.com/en-us/windows/win32/api/securitybaseapi/nf-securitybaseapi-gettokeninformation
        [DllImport("advapi32.dll", ExactSpelling = true, SetLastError = true)]
        public static extern bool GetTokenInformation(
            SafeProcessTokenHandle TokenHandle,
            int TokenInformationClass,
            IntPtr TokenInformation,
            uint TokenInformationLength,
            out uint ReturnLength);

        [DllImport("advapi32.dll", ExactSpelling = true, SetLastError = true)]
        public static extern bool GetTokenInformation(
            SafeProcessTokenHandle TokenHandle,
            int TokenInformationClass,
            SafeCoTaskMemFreeHandle TokenInformation,
            uint TokenInformationLength,
            out uint ReturnLength);

        [DllImport("advapi32.dll", ExactSpelling = true, SetLastError = true)]
        public static extern bool GetTokenInformation(
            SafeProcessTokenHandle TokenHandle,
            int TokenInformationClass,
            out TOKEN_ELEVATION TokenElevation,
            uint TokenInformationLength,
            out uint ReturnLength);

        // https://learn.microsoft.com/en-us/windows/win32/api/winbase/nf-winbase-lookupprivilegenamew
        [DllImport("advapi32.dll", ExactSpelling = true, SetLastError = true, CharSet = CharSet.Unicode)]
        public static extern bool LookupPrivilegeNameW(
            string lpSystemName,
            IntPtr pLUID,
            [Out] char[] lpName,
            out uint cchName);

        // https://learn.microsoft.com/en-us/windows/win32/api/winnt/ns-winnt-token_elevation
        [StructLayout(LayoutKind.Sequential)]
        public struct TOKEN_ELEVATION
        {
            public uint TokenIsElevated;
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct TOKEN_PRIVILEGES
        {
            public int PrivilegeCount;
            public LUID_AND_ATTRIBUTES Privileges;
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct LUID_AND_ATTRIBUTES
        {
            public LUID Luid;
            public uint Attributes;
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct LUID
        {
            public int LowPart;
            public int HighPart;
        }

        // https://docs.microsoft.com/en-us/windows/win32/api/evntrace/ns-evntrace-event_trace_properties
        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
        public struct EVENT_TRACE_PROPERTIES
        {
            public WNODE_HEADER Wnode;
            public uint BufferSize;
            public uint MinimumBuffers;
            public uint MaximumBuffers;
            public uint MaximumFileSize;
            public uint LogFileMode;
            public uint FlushTimer;
            public uint EnableFlags;
            public int AgeLimit;
            public uint NumberOfBuffers;
            public uint FreeBuffers;
            public uint EventsLost;
            public uint BuffersWritten;
            public uint LogBuffersLost;
            public uint RealTimeBuffersLost;
            public IntPtr LoggerThreadId;
            public int LogFileNameOffset;
            public int LoggerNameOffset;
        }

        // https://docs.microsoft.com/en-us/windows/win32/etw/wnode-header
        [StructLayout(LayoutKind.Sequential)]
        public struct WNODE_HEADER
        {
            public uint BufferSize;
            public uint ProviderId;
            public ulong HistoricalContext;
            public ulong KernelHandle;
            public Guid Guid;
            public uint ClientContext;
            public uint Flags;
        }

        public struct EventTraceProperties
        {
            public EVENT_TRACE_PROPERTIES Properties;
            public string SessionName;
            public string LogFileName;

            public EventTraceProperties(EVENT_TRACE_PROPERTIES properties, string sessionName, string logFileName)
            {
                Properties = properties;
                SessionName = sessionName;
                LogFileName = logFileName;
            }
        }

        //https://docs.microsoft.com/en-us/windows/win32/api/evntrace/nf-evntrace-queryalltracesw
        [DllImport("Advapi32.dll", ExactSpelling = true)]
        public static extern int QueryAllTracesW(IntPtr[] PropertyArray, uint PropertyArrayCount, ref int LoggerCount);

        //https://docs.microsoft.com/en-us/windows/win32/api/evntrace/nf-evntrace-stoptracew
        [DllImport("Advapi32.dll", CharSet = CharSet.Unicode, ExactSpelling = true)]
        public static extern int StopTraceW(ulong TraceHandle, string InstanceName, IntPtr Properties); // TRACEHANDLE is defined as ULONG64

        const int MAX_SESSIONS = 64;
        const int MAX_NAME_COUNT = 1024; // max char count for LogFileName & SessionName
        const uint ERROR_SUCCESS = 0;

        // https://docs.microsoft.com/en-us/windows/win32/etw/wnode-header
        // > The size of memory must include the room for the EVENT_TRACE_PROPERTIES structure plus the session name string and log file name string that follow the structure in memory.
        static readonly int PropertiesSize = Marshal.SizeOf(typeof(EVENT_TRACE_PROPERTIES)) + 2 * sizeof(char) * MAX_NAME_COUNT; // EVENT_TRACE_PROPERTIES + LogFileName & LoggerName
        static readonly int LoggerNameOffset = Marshal.SizeOf(typeof(EVENT_TRACE_PROPERTIES));
        static readonly int LogFileNameOffset = LoggerNameOffset + sizeof(char) * MAX_NAME_COUNT;

        public static List<EventTraceProperties> QueryAllTraces()
        {
            List<EventTraceProperties> eventProperties = null;
            int BufferSize = PropertiesSize * MAX_SESSIONS;

            // Wrap the native memory in SafeHandle-derived class
            using (Win32.SafeCoTaskMemFreeHandle safeHandle = new Win32.SafeCoTaskMemFreeHandle(Marshal.AllocCoTaskMem(BufferSize)))
            {
                IntPtr pBuffer = safeHandle.DangerousGetHandle();
                Win32.Kernel32.RtlZeroMemory(pBuffer, BufferSize);
                IntPtr[] sessions = new IntPtr[64];

                for (int i = 0; i < 64; ++i)
                {
                    //sessions[i] = pBuffer + (i * PropertiesSize); // This does not compile in .NET 2.0
                    sessions[i] = new IntPtr(pBuffer.ToInt64() + (i * PropertiesSize));

                    // Marshal from managed to native
                    EVENT_TRACE_PROPERTIES props = new EVENT_TRACE_PROPERTIES();
                    props.Wnode.BufferSize = (uint)PropertiesSize;
                    props.LoggerNameOffset = LoggerNameOffset;
                    props.LogFileNameOffset = LogFileNameOffset;
                    Marshal.StructureToPtr(props, sessions[i], false);
                }

                int loggerCount = 0;
                int status = QueryAllTracesW(sessions, MAX_SESSIONS, ref loggerCount);

                if (status != ERROR_SUCCESS)
                {
                    throw new Win32Exception(status);
                }

                eventProperties = new List<EventTraceProperties>();
                for (int i = 0; i < loggerCount; ++i)
                {
                    // Marshal back from native to managed.
                    EVENT_TRACE_PROPERTIES props = (EVENT_TRACE_PROPERTIES)Marshal.PtrToStructure(sessions[i], typeof(EVENT_TRACE_PROPERTIES));
                    string sessionName = Marshal.PtrToStringUni(new IntPtr(sessions[i].ToInt64() + LoggerNameOffset));
                    string logFileName = Marshal.PtrToStringUni(new IntPtr(sessions[i].ToInt64() + LogFileNameOffset));

                    //eventProperties.Add(new EventTraceProperties { Properties = props, SessionName = sessionName, LogFileName = logFileName });
                    eventProperties.Add(new EventTraceProperties(props, sessionName, logFileName));
                }
            }

            return eventProperties;
        }

        public static EventTraceProperties StopTrace(string SessionName)
        {
            using (var safeHandle = new SafeCoTaskMemFreeHandle(Marshal.AllocCoTaskMem(PropertiesSize)))
            {
                IntPtr pProps = safeHandle.DangerousGetHandle();
                Win32.Kernel32.RtlZeroMemory(pProps, PropertiesSize);

                EVENT_TRACE_PROPERTIES props = new EVENT_TRACE_PROPERTIES();
                props.Wnode.BufferSize = (uint)PropertiesSize;
                props.LoggerNameOffset = LoggerNameOffset;
                props.LogFileNameOffset = LogFileNameOffset;
                Marshal.StructureToPtr(props, pProps, false);

                int status = StopTraceW(0, SessionName, pProps);
                if (status != ERROR_SUCCESS)
                {
                    throw new Win32Exception(status);
                }

                props = (EVENT_TRACE_PROPERTIES)Marshal.PtrToStructure(pProps, typeof(EVENT_TRACE_PROPERTIES));
                string sessionName = Marshal.PtrToStringUni(new IntPtr(pProps.ToInt64() + LoggerNameOffset));
                string logFileName = Marshal.PtrToStringUni(new IntPtr(pProps.ToInt64() + LogFileNameOffset));

                return new EventTraceProperties(props, sessionName, logFileName);
            }
        }
    } // end of class Advapi32

    public static class Dbghelp
    {
        [Flags]
        public enum MINIDUMP_TYPE {
            MiniDumpNormal                         = 0x00000000,
            MiniDumpWithDataSegs                   = 0x00000001,
            MiniDumpWithFullMemory                 = 0x00000002,
            MiniDumpWithHandleData                 = 0x00000004,
            MiniDumpFilterMemory                   = 0x00000008,
            MiniDumpScanMemory                     = 0x00000010,
            MiniDumpWithUnloadedModules            = 0x00000020,
            MiniDumpWithIndirectlyReferencedMemory = 0x00000040,
            MiniDumpFilterModulePaths              = 0x00000080,
            MiniDumpWithProcessThreadData          = 0x00000100,
            MiniDumpWithPrivateReadWriteMemory     = 0x00000200,
            MiniDumpWithoutOptionalData            = 0x00000400,
            MiniDumpWithFullMemoryInfo             = 0x00000800,
            MiniDumpWithThreadInfo                 = 0x00001000,
            MiniDumpWithCodeSegs                   = 0x00002000,
            MiniDumpWithoutAuxiliaryState          = 0x00004000,
            MiniDumpWithFullAuxiliaryState         = 0x00008000,
            MiniDumpWithPrivateWriteCopyMemory     = 0x00010000,
            MiniDumpIgnoreInaccessibleMemory       = 0x00020000,
            MiniDumpWithTokenInformation           = 0x00040000,
            MiniDumpWithModuleHeaders              = 0x00080000,
            MiniDumpFilterTriage                   = 0x00100000,
            MiniDumpWithAvxXStateContext           = 0x00200000,
            MiniDumpWithIptTrace                   = 0x00400000,
            MiniDumpValidTypeFlags                 = 0x007fffff,
        }

        [DllImport("Dbghelp.dll", SetLastError=true)]
        public static extern bool MiniDumpWriteDump(
            SafeHandle hProcess,
            uint ProcessId,
            SafeHandle hFile,
            uint DumpType,
            IntPtr ExceptionParam,
            IntPtr UserStreamParam,
            IntPtr CallbackParam
        );
    } // end of class Dbghelp

    public static class Kernel32
    {
        // Need GlobalFree to free the memory allocated for WINHTTP_PROXY_INFO & WINHTTP_CURRENT_USER_IE_PROXY_CONFIG
        [DllImport("kernel32.dll", CallingConvention = CallingConvention.StdCall)]
        public static extern IntPtr GlobalFree(IntPtr hMem);

        [DllImport("kernel32.dll", ExactSpelling = true, SetLastError = true)]
        public static extern bool IsWow64Process(SafeHandle hProcess, out bool wow64Process);

        [DllImport("kernel32.dll", ExactSpelling = true)]
        public static extern void RtlZeroMemory(IntPtr dst, int length);

        [DllImport("kernel32.dll", CharSet = CharSet.Unicode, ExactSpelling = true)]
        public static extern SafeFileMappingHandle OpenFileMappingW(uint dwDesiredAccess, bool bInheritHandle, string lpName);

        [DllImport("kernel32.dll", ExactSpelling = true, SetLastError = true)]
        public static extern bool CloseHandle(IntPtr handle);

        [DllImport("Kernel32.dll", ExactSpelling = true)]
        public static extern uint GetACP();

        [DllImport("kernel32.dll", ExactSpelling = true, SetLastError = true)]
        public static extern SafeProcessHandle OpenProcess(
            uint dwDesiredAccess,
            [MarshalAs(UnmanagedType.Bool)] bool bInheritHandle,
            int dwProcessId);

        [DllImport("kernel32.dll", ExactSpelling = true, SetLastError = true)]
        public static extern IntPtr GetConsoleWindow();

        [DllImport("kernel32.dll", CharSet = CharSet.Unicode, ExactSpelling = true, SetLastError = true)]
        public static extern IntPtr LoadLibraryW(string lpLibFileName);

        [DllImport("kernel32.dll", ExactSpelling = true, SetLastError = true)]
        public static extern bool FreeLibrary(IntPtr hModule);

        public const uint PROCESS_QUERY_LIMITED_INFORMATION = 0x1000;
        public const uint PROCESS_QUERY_INFORMATION = 0x0400;
    }

    // ETW Logging Mode Constants for logman
    // https://docs.microsoft.com/en-us/windows/win32/etw/logging-mode-constants
    public static class Logman
    {
        public static class Mode
        {
            public static string EVENT_TRACE_FILE_MODE_SEQUENTIAL = "sequential";
            public static string EVENT_TRACE_FILE_MODE_CIRCULAR   = "circular";
            public static string EVENT_TRACE_FILE_MODE_APPEND     = "append";
            public static string EVENT_TRACE_FILE_MODE_NEWFILE    = "newfile";
            public static string EVENT_TRACE_USE_GLOBAL_SEQUENCE  = "globalsequence";
            public static string EVENT_TRACE_USE_LOCAL_SEQUENCE   = "localsequence";
        }
    }

    namespace Mapi
    {
        // https://learn.microsoft.com/en-us/office/client-developer/outlook/mapi/pidtagresourceflags-canonical-property
        [Flags]
        public enum ResourceFlags
        {
            None = 0,
            STATUS_DEFAULT_OUTBOUND = 1,
            STATUS_DEFAULT_STORE = 2,
            STATUS_PRIMARY_IDENTITY = 4,
            STATUS_SIMPLE_STORE = 8,
            STATUS_XP_PREFER_LAST = 0x10,
            STATUS_NO_PRIMARY_IDENTITY = 0x20,
            STATUS_NO_DEFAULT_STORE = 0x40,
            STATUS_TEMP_SECTION = 0x80,
            STATUS_OWN_STORE = 0x100,
            STATUS_NEED_IPM_TREE = 0x800,
            STATUS_PRIMARY_STORE = 0x1000,
            STATUS_SECONDARY_STORE = 0x2000
        }

        public enum CacheSyncMode
        {
            Headers = 1,
            FullItems = 2,
            Drizzle = 3
        }

        public enum OfflineState
        {
            Offline = 1,
            Online = 2,
            MASK = 3
        }

        [Flags]
        public enum ProfileConfigFlags
        {
            CONFIG_SERVICE = 1,
            CONFIG_SHOW_STARTUP_UI = 2,
            CONFIG_SHOW_CONNECT_UI = 4,
            CONFIG_PROMPT_FOR_CREDENTIALS = 8,
            CONFIG_NO_AUTO_DETECT = 0x10,
            CONFIG_OST_CACHE_ONLY = 0x20,
            CONFIG_USE_SMTP_ADDRESSES = 0x40,
            CONFIG_OST_CACHE_PRIVATE = 0x180,
            CONFIG_OST_DISASTER_RECOVERY = 0x200,
            CONFIG_OST_CACHE_PUBLIC = 0x400,
            CONFIG_OST_CACHE_DELEGATE_PIM = 0x800,
            CONFIG_PUB_FOLDERS_ALIVE = 0x1000,
            CONFIG_PUB_FOLDERS_DEAD = 0x2000,
        }

        public enum SharedCalProfileConfigFlags
        {
            None = 0,
            Rest = 1,
            Mapi = 2
        }
    }

    public static class Netapi32
    {
        public enum NETSETUP_JOIN_STATUS
        {
            NetSetupUnknownStatus = 0,
            NetSetupUnjoined,
            NetSetupWorkgroupName,
            NetSetupDomainName
        }

        [DllImport("Netapi32.dll", CharSet = CharSet.Unicode, ExactSpelling = true)]
        public static extern uint NetGetJoinInformation(string server, out IntPtr name, out uint status);

        [DllImport("Netapi32.dll", ExactSpelling = true)]
        public static extern uint NetApiBufferFree(IntPtr Buffer);
    }

    public static class Netlistmgr
    {
        // From netlistmgr.h
        public const string CLSID_NetworkListManager = "DCB00C01-570F-4A9B-8D69-199FDBA5723B";

        // # NLM_CONNECTIVITY enumeration
        // # https://docs.microsoft.com/en-us/windows/win32/api/netlistmgr/ne-netlistmgr-nlm_connectivity
        [Flags]
        public enum NLM_CONNECTIVITY
        {
            NLM_CONNECTIVITY_DISCONNECTED      = 0,
            NLM_CONNECTIVITY_IPV4_NOTRAFFIC    = 1,
            NLM_CONNECTIVITY_IPV6_NOTRAFFIC    = 2,
            NLM_CONNECTIVITY_IPV4_SUBNET       = 0x10,
            NLM_CONNECTIVITY_IPV4_LOCALNETWORK = 0x20,
            NLM_CONNECTIVITY_IPV4_INTERNET     = 0x40,
            NLM_CONNECTIVITY_IPV6_SUBNET       = 0x100,
            NLM_CONNECTIVITY_IPV6_LOCALNETWORK = 0x200,
            NLM_CONNECTIVITY_IPV6_INTERNET     = 0x400
        }

        // # NLM_CONNECTION_COST enumeration
        // # https://docs.microsoft.com/en-us/windows/win32/api/netlistmgr/ne-netlistmgr-nlm_connection_cost
        [Flags]
        public enum NLM_CONNECTION_COST
        {
            NLM_CONNECTION_COST_UNKNOWN              = 0,
            NLM_CONNECTION_COST_UNRESTRICTED         = 0x1,
            NLM_CONNECTION_COST_FIXED                = 0x2,
            NLM_CONNECTION_COST_VARIABLE             = 0x4,
            NLM_CONNECTION_COST_OVERDATALIMIT        = 0x10000,
            NLM_CONNECTION_COST_CONGESTED            = 0x20000,
            NLM_CONNECTION_COST_ROAMING              = 0x40000,
            NLM_CONNECTION_COST_APPROACHINGDATALIMIT = 0x80000
        }

        [ComImport, Guid("DCB00008-570F-4A9B-8D69-199FDBA5723B"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        public interface INetworkCostManager
        {
            int GetCost(out NLM_CONNECTION_COST pCost, IntPtr DestIPAddr);
            int GetDataPlanStatus(IntPtr pDataPlanStatus, IntPtr pDestIPAddr);
            int SetDestinationAddresses(uint length, IntPtr pDestIPAddrList, bool bAppend);
        }

        public static NLM_CONNECTION_COST GetMeteredNetworkCost()
        {
            Type nlmType = Type.GetTypeFromCLSID(new Guid(CLSID_NetworkListManager));
            INetworkCostManager ncm = Activator.CreateInstance(nlmType) as INetworkCostManager;

            if (ncm == null)
            {
                throw new InvalidOperationException("Failed to obtain INetworkCostManager");
            }

            try
            {
                NLM_CONNECTION_COST cost = NLM_CONNECTION_COST.NLM_CONNECTION_COST_UNKNOWN;
                int hr = ncm.GetCost(out cost, IntPtr.Zero);
                Marshal.ThrowExceptionForHR(hr);
                return cost;
            }
            finally
            {
                Marshal.FinalReleaseComObject(ncm);
            }
        }
    }

    public static class Ole32
    {
        [DllImport("ole32.dll", CharSet = CharSet.Unicode, ExactSpelling = true)]
        public static extern uint CLSIDFromProgID(string progOd, out Guid clsid);

        [DllImport("ole32.dll", CharSet = CharSet.Unicode, ExactSpelling = true)]
        public static extern uint StringFromCLSID([MarshalAs(UnmanagedType.LPStruct)] Guid refclsid, out IntPtr pClsidString);
    }

    public class SafeFileMappingHandle : SafeHandleZeroOrMinusOneIsInvalid
    {
        public SafeFileMappingHandle() : base(true) {}

        override protected bool ReleaseHandle()
        {
            return Win32.Kernel32.CloseHandle(handle);
        }
    }

    // SafeHandle-derived class for the native memory that should be freed by GlobalFree.
    public class SafeGlobalFreeHandle: SafeHandleZeroOrMinusOneIsInvalid
    {
        public SafeGlobalFreeHandle(): base(true) {}
        public SafeGlobalFreeHandle(bool ownsHandle): base(ownsHandle) {}
        public SafeGlobalFreeHandle(IntPtr handle, bool ownsHandle = true): base(ownsHandle)
        {
            SetHandle(handle);
        }

        override protected bool ReleaseHandle()
        {
            return Win32.Kernel32.GlobalFree(handle) == IntPtr.Zero;
        }
    }

    // SafeHandle-derived class for the native Unicode string that should be freed by GlobalFree
    public class SafeGlobalFreeString: SafeGlobalFreeHandle
    {
        public SafeGlobalFreeString(): base(true) {}
        public SafeGlobalFreeString(bool ownsHandle): base(ownsHandle) {}
        public SafeGlobalFreeString(IntPtr handle, bool ownsHandle = true): base(ownsHandle)
        {
            SetHandle(handle);
        }

        public override string ToString()
        {
            return Marshal.PtrToStringUni(handle);
        }
    }

    // SafeHandle-derived class for the native memory that should be freed by CoTaskMemFree.
    public class SafeCoTaskMemFreeHandle: SafeHandle
    {
        public SafeCoTaskMemFreeHandle(): base(IntPtr.Zero, true) {}

        public SafeCoTaskMemFreeHandle(IntPtr handle, bool ownsHandle = true): base(IntPtr.Zero, ownsHandle)
        {
            SetHandle(handle);
        }

        public SafeCoTaskMemFreeHandle(int size) : this(Marshal.AllocCoTaskMem(size)) {}

        public static SafeCoTaskMemFreeHandle InvalidHandle
        {
            get
            {
                return new SafeCoTaskMemFreeHandle();
            }
        }

        public override bool IsInvalid
        {
            get { return IsClosed || handle == IntPtr.Zero; }
        }

        protected override bool ReleaseHandle()
        {
            Marshal.FreeCoTaskMem(handle);
            return true;
        }

        public void Reset(IntPtr handle)
        {
            if (!IsInvalid)
            {
                ReleaseHandle();
            }

            this.handle = handle;
        }
    }

    public class SafeProcessTokenHandle : SafeHandleZeroOrMinusOneIsInvalid
    {
        public SafeProcessTokenHandle() : base(true) {}

        public SafeProcessTokenHandle(IntPtr handle) : base(true)
        {
            SetHandle(handle);
        }

        protected override bool ReleaseHandle()
        {
            return Win32.Kernel32.CloseHandle(handle);
        }
    }

    public static class Shell32
    {
        // https://learn.microsoft.com/en-us/windows/win32/api/shellapi/ne-shellapi-query_user_notification_state
        public enum QUERY_USER_NOTIFICATION_STATE
        {
            QUNS_NOT_PRESENT             = 1,
            QUNS_BUSY                    = 2,
            QUNS_RUNNING_D3D_FULL_SCREEN = 3,
            QUNS_PRESENTATION_MODE       = 4,
            QUNS_ACCEPTS_NOTIFICATIONS   = 5,
            QUNS_QUIET_TIME              = 6,
            QUNS_APP                     = 7
        };

        // https://learn.microsoft.com/en-us/windows/win32/api/shellapi/nf-shellapi-shqueryusernotificationstate
        [DllImport("shell32.dll")]
        public static extern uint SHQueryUserNotificationState(out QUERY_USER_NOTIFICATION_STATE pquns);
    }

    public static class User32
    {
        [DllImport("user32.dll")]
        public static extern bool IsHungAppWindow(IntPtr hWnd);

        [DllImport("user32.dll", SetLastError=true, CharSet=CharSet.Auto)]
        public static extern uint SendMessageTimeoutW(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam, uint fuFlags, uint uTimeout, out IntPtr lpdwResult);

        [DllImport("user32.dll", SetLastError=true)]
        public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        [StructLayout(LayoutKind.Sequential)]
        public struct KeyboardInput
        {
            public ushort wVk;
            public ushort wScan;
            public uint dwFlags;
            public uint time;
            public IntPtr dwExtraInfo;
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct MouseInput
        {
            public int dx;
            public int dy;
            public uint mouseData;
            public uint dwFlags;
            public uint time;
            public IntPtr dwExtraInfo;
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct HardwareInput
        {
            public uint uMsg;
            public ushort wParamL;
            public ushort wParamH;
        }

        [StructLayout(LayoutKind.Explicit)]
        public struct InputUnion
        {
            [FieldOffset(0)] public MouseInput mi;
            [FieldOffset(0)] public KeyboardInput ki;
            [FieldOffset(0)] public HardwareInput hi;
        }

        public struct Input
        {
            public InputType type;
            public InputUnion u;
        }

        public enum InputType
        {
            Mouse = 0,
            Keyboard = 1,
            Hardware = 2
        }

        [Flags]
        public enum KeyEventF
        {
            KeyDown = 0x0000,
            ExtendedKey = 0x0001,
            KeyUp = 0x0002,
            Unicode = 0x0004,
            Scancode = 0x0008
        }

        [DllImport("user32.dll", SetLastError = true)]
        public static extern uint SendInput(int nInputs, Input[] pInputs, int cbSize);

        public static void SendKeyboardInput(ushort vkey, KeyEventF flags)
        {
            Input[] inputs = new Input[1];
            inputs[0].type = InputType.Keyboard;
            inputs[0].u.ki.wVk = vkey;
            inputs[0].u.ki.dwFlags = (uint)flags;

            var count = SendInput(1, inputs, Marshal.SizeOf(typeof(Input)));

            if (count == 0)
            {
                throw new Win32Exception(Marshal.GetLastWin32Error());
            }
        }

        public static void SendCtrl5()
        {
            Input[] inputs = new Input[4];

            inputs[0].type = InputType.Keyboard;
            inputs[0].u.ki.wVk = 0x11; // VK_CONTROL

            inputs[1].type = InputType.Keyboard;
            inputs[1].u.ki.wVk = 0x35; // 5

            inputs[2].type = InputType.Keyboard;
            inputs[2].u.ki.wVk = 0x35; // 5
            inputs[2].u.ki.dwFlags = (uint)(KeyEventF.KeyUp);

            inputs[2].type = InputType.Keyboard;
            inputs[2].u.ki.wVk = 0x11; // VK_CONTROL
            inputs[2].u.ki.dwFlags = (uint)(KeyEventF.KeyUp);

            var count = SendInput(inputs.Length, inputs, Marshal.SizeOf(typeof(Input)));

            if (count == 0)
            {
                throw new Win32Exception(Marshal.GetLastWin32Error());
            }
        }
    }

    public static class WinHttp
    {
        // Some error codes from winhttp.h
        public enum Error
        {
            WINHTTP_ERROR_BASE                      = 12000,
            ERROR_WINHTTP_INTERNAL_ERROR            = WINHTTP_ERROR_BASE + 4,
            ERROR_WINHTTP_NAME_NOT_RESOLVED         = WINHTTP_ERROR_BASE + 7,
            ERROR_WINHTTP_INCORRECT_HANDLE_TYPE     = WINHTTP_ERROR_BASE + 18,
            ERROR_WINHTTP_INCORRECT_HANDLE_STATE    = WINHTTP_ERROR_BASE + 19,
            ERROR_WINHTTP_AUTO_PROXY_SERVICE_ERROR  = WINHTTP_ERROR_BASE + 178,
            ERROR_WINHTTP_BAD_AUTO_PROXY_SCRIPT     = WINHTTP_ERROR_BASE + 166,
            ERROR_WINHTTP_UNABLE_TO_DOWNLOAD_SCRIPT = WINHTTP_ERROR_BASE + 167,
            ERROR_WINHTTP_UNHANDLED_SCRIPT_TYPE     = WINHTTP_ERROR_BASE + 176,
            ERROR_WINHTTP_SCRIPT_EXECUTION_ERROR    = WINHTTP_ERROR_BASE + 177,
            ERROR_WINHTTP_NOT_INITIALIZED           = WINHTTP_ERROR_BASE + 172,
            ERROR_WINHTTP_SECURE_FAILURE            = WINHTTP_ERROR_BASE + 175,
            ERROR_WINHTTP_AUTODETECTION_FAILED      = WINHTTP_ERROR_BASE + 180
        }

        // https://docs.microsoft.com/en-us/windows/win32/api/winhttp/ns-winhttp-winhttp_proxy_info
        public struct WINHTTP_PROXY_INFO
        {
            public ProxyAccessType dwAccessType;
            public IntPtr lpszProxy;
            public IntPtr lpszProxyBypass;
        }

        // https://docs.microsoft.com/en-us/windows/win32/api/winhttp/ns-winhttp-winhttp_current_user_ie_proxy_config
        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
        public struct WINHTTP_CURRENT_USER_IE_PROXY_CONFIG
        {
            public bool fAutoDetect;
            public IntPtr lpszAutoConfigUrl;
            public IntPtr lpszProxy;
            public IntPtr lpszProxyBypass;
        }

        // From winhttp.h
        // WinHttpOpen dwAccessType values (also for WINHTTP_PROXY_INFO::dwAccessType)
        public enum ProxyAccessType
        {
            WINHTTP_ACCESS_TYPE_DEFAULT_PROXY = 0,
            WINHTTP_ACCESS_TYPE_NO_PROXY = 1,
            WINHTTP_ACCESS_TYPE_NAMED_PROXY = 3,
            WINHTTP_ACCESS_TYPE_AUTOMATIC_PROXY = 4
        }

        [Flags]
        public enum AutoDetectType
        {
            WINHTTP_AUTO_DETECT_TYPE_DHCP = 0x1,
            WINHTTP_AUTO_DETECT_TYPE_DNS_A = 0x2
        }

        // https://docs.microsoft.com/en-us/windows/win32/api/winhttp/nf-winhttp-winhttpgetdefaultproxyconfiguration
        [DllImport("winhttp.dll", SetLastError = true)]
        public static extern bool WinHttpGetDefaultProxyConfiguration(out WINHTTP_PROXY_INFO proxyInfo);

        // https://docs.microsoft.com/en-us/windows/win32/api/winhttp/nf-winhttp-winhttpgetieproxyconfigforcurrentuser
        [DllImport("winhttp.dll", SetLastError = true)]
        public static extern bool WinHttpGetIEProxyConfigForCurrentUser(out WINHTTP_CURRENT_USER_IE_PROXY_CONFIG proxyConfig);

        // https://docs.microsoft.com/en-us/windows/win32/api/winhttp/nf-winhttp-winhttpdetectautoproxyconfigurl
        [DllImport("winhttp.dll", SetLastError = true)]
        public static extern bool WinHttpDetectAutoProxyConfigUrl(AutoDetectType dwAutoDetectFlags, out Win32.SafeGlobalFreeString autoConfigUrlHandle);
    } // end of class WinHttp

    public static class WinInet
    {
        // From wininet.h
        [Flags]
        public enum PER_CONN_FLAGS
        {
            PROXY_TYPE_DIRECT         = 1,
            PROXY_TYPE_PROXY          = 2,
            PROXY_TYPE_AUTO_PROXY_URL = 4,
            PROXY_TYPE_AUTO_DETECT    = 8
        }
    }

    public static class Wscapi
    {
        public enum WSC_SECURITY_PROVIDER_HEALTH
        {
            WSC_SECURITY_PROVIDER_HEALTH_GOOD,
            WSC_SECURITY_PROVIDER_HEALTH_NOTMONITORED,
            WSC_SECURITY_PROVIDER_HEALTH_POOR,
            WSC_SECURITY_PROVIDER_HEALTH_SNOOZE
        }

        // https://docs.microsoft.com/en-us/windows/win32/api/wscapi/ne-wscapi-wsc_security_provider
        [Flags]
        public enum WSC_SECURITY_PROVIDER
        {
            // Represents the aggregation of all firewalls for this computer.
            WSC_SECURITY_PROVIDER_FIREWALL =                   0x1,
            // Represents the Automatic updating settings for this computer.
            WSC_SECURITY_PROVIDER_AUTOUPDATE_SETTINGS  =       0x2,
            // Represents the aggregation of all antivirus products for this comptuer.
            WSC_SECURITY_PROVIDER_ANTIVIRUS =                  0x4,
            // Represents the aggregation of all antispyware products for this comptuer.
            WSC_SECURITY_PROVIDER_ANTISPYWARE =                0x8,
            // Represents the settings that restrict the access of web sites in each of the internet zones.
            WSC_SECURITY_PROVIDER_INTERNET_SETTINGS =          0x10,
            // Represents the User Account Control settings on this machine.
            WSC_SECURITY_PROVIDER_USER_ACCOUNT_CONTROL =       0x20,
            // Represents the running state of the Security Center service on this machine.
            WSC_SECURITY_PROVIDER_SERVICE =                    0x40,

            WSC_SECURITY_PROVIDER_NONE =                       0,

            // Aggregates all of the items that Security Center monitors.
            WSC_SECURITY_PROVIDER_ALL =                             WSC_SECURITY_PROVIDER_FIREWALL |
                                                                    WSC_SECURITY_PROVIDER_AUTOUPDATE_SETTINGS |
                                                                    WSC_SECURITY_PROVIDER_ANTIVIRUS |
                                                                    WSC_SECURITY_PROVIDER_ANTISPYWARE |
                                                                    WSC_SECURITY_PROVIDER_INTERNET_SETTINGS |
                                                                    WSC_SECURITY_PROVIDER_USER_ACCOUNT_CONTROL |
                                                                    WSC_SECURITY_PROVIDER_SERVICE
        }

        // https://docs.microsoft.com/en-us/windows/win32/api/wscapi/nf-wscapi-wscgetsecurityproviderhealth
        [DllImport("Wscapi.dll", SetLastError = true)]
        public static extern int WscGetSecurityProviderHealth(uint Providers, out int pHealth);
    }

    public static class WamInterop
    {
        [DllImport("WamInterop.dll")]
        public static extern int RequestToken(IntPtr hwnd, IntPtr request, out IntPtr result);

        [DllImport("WamInterop.dll")]
        public static extern IntPtr CreateAnchorWindow();
    }
}
'@

function Open-Log {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$Path,
        [switch]$AutoFlush,
        [switch]$WithBOM
    )

    if ($Script:LogWriter) {
        Close-Log
    }

    # Open a file & add header
    try {
        $utf8Encoding = New-Object System.Text.UTF8Encoding -ArgumentList $WithBOM.IsPresent
        $Script:LogWriter = New-Object System.IO.StreamWriter -ArgumentList $Path, <# append #> $true, $utf8Encoding

        if ($AutoFlush) {
            $Script:LogWriter.AutoFlush = $true
        }

        $Script:LogWriter.WriteLine("datetime,thread_relative_delta,thread,function,category,message")
    }
    catch {
        Write-Error -ErrorRecord $_
    }
}

function Write-Log {
    [CmdletBinding()]
    param (
        [Parameter(ValueFromPipeline = $true)]
        [string]$Message,
        [Parameter(ValueFromPipeline = $true)]
        [System.Management.Automation.ErrorRecord]$ErrorRecord,
        [ValidateSet('Information', 'Warning', 'Error')]
        $Category = 'Information',
        # Output the given ErrorRecord
        [Switch]$PassThru
    )

    process {
        # If ErrorRecord is provided, use it.
        if ($ErrorRecord) {
            $errorDetails = $null

            if ($ErrorRecord.ErrorDetails.Message -ne $ErrorRecord.Exception.Message) {
                $errorDetails = $ErrorRecord.ErrorDetails.Message
            }

            $Message = "$Message; [ErrorRecord] $(if ($errorDetails) { "ErrorDetails:$errorDetails, " })ExceptionType:$($ErrorRecord.Exception.GetType().Name), Exception.Message:$($ErrorRecord.Exception.Message), InvocationInfo.Line:'$($ErrorRecord.InvocationInfo.Line.Trim())', ScriptStackTrace:$($ErrorRecord.ScriptStackTrace.Replace([Environment]::NewLine, ' '))"
        }

        # Ignore null or an empty string.
        if (-not $Message) {
            return
        }

        # If Open-Log is not called beforehand, just output to verbose.
        if (-not $Script:LogWriter) {
            Write-Verbose $Message
            return
        }

        # If LogWriter exists but disposed already, something went wrong.
        if (-not $Script:LogWriter.BaseStream.CanWrite) {
            Write-Warning "LogWriter has been disposed already"
            return
        }

        $currentTime = Get-Date
        $currentTimeFormatted = $currentTime.ToString('o')

        # Delta time is relative to thread.
        # Each thread has it's own copy of LastLogTime now.
        [TimeSpan]$delta = 0

        if ($Script:LastLogTime) {
            $delta = $currentTime.Subtract($Script:LastLogTime)
        }

        $caller = '<ScriptBlock>'
        $caller = Get-PSCallStack | Select-Object -Skip 1 | & {
            process {
                if (-not $_.Command.StartsWith('<ScriptBlock>')) {
                    $_.Command
                }
            }
        } | Select-Object -First 1

        # Format as CSV:
        $sb = New-Object System.Text.StringBuilder
        $null = $sb.Append($currentTimeFormatted).Append(',')
        $null = $sb.Append($delta).Append(',')
        $null = $sb.Append([System.Threading.Thread]::CurrentThread.ManagedThreadId).Append(',')
        $null = $sb.Append($caller).Append(',')

        $categoryEmoji = switch ($Category) {
            'Information' { $Script:Emoji.Information; break }
            'Warning' { $Script:Emoji.Warning; break }
            'Error' { $Script:Emoji.Error; break }
        }

        $null = $sb.Append($categoryEmoji).Append(',')

        $null = $sb.Append('"').Append($Message.Replace('"', "'")).Append('"')

        # Protect from concurrent write
        [System.Threading.Monitor]::Enter($Script:LogWriter)

        try {
            $Script:LogWriter.WriteLine($sb.ToString())
        }
        finally {
            [System.Threading.Monitor]::Exit($Script:LogWriter)
        }

        $sb = $null
        $Script:LastLogTime = $currentTime

        if ($PassThru) {
            $ErrorRecord
        }
    }
}

function Close-Log {
    if ($Script:LogWriter) {
        if ($Script:LogWriter.BaseStream.CanWrite) {
            Write-Log "Closing LogWriter"
            $Script:LogWriter.Close()
        }

        $Script:LogWriter = $null
        $Script:LastLogTime = $null
    }
}

function Get-Timestamp {
    [CmdletBinding()]
    [OutputType([long])]
    param()

    [System.Diagnostics.Stopwatch]::GetTimestamp()
}

function Get-Elapsed {
    [CmdletBinding()]
    [OutputType([TimeSpan])]
    param(
        [Parameter(Mandatory, ValueFromPipeline)]
        [long]$StartingTimestamp,
        [long]$EndingTimestamp
    )

    process {
        if ($PSBoundParameters.ContainsKey('EndingTimestamp')) {
            [TimeSpan]::FromTicks($EndingTimestamp - $StartingTimestamp)
        }
        else {
            [TimeSpan]::FromTicks([System.Diagnostics.Stopwatch]::GetTimestamp() - $StartingTimestamp)
        }
    }
}

<#
.SYNOPSIS
Create a runspace pool so that Start-Task commands can use it.
Make sure to call Close-TaskRunspace to dispose the runspace pool.
#>
function Open-TaskRunspace {
    [CmdletBinding()]
    param(
        [ValidateRange(1, [int]::MaxValue)]
        [int]$MinRunspaces = 1,
        # Maximum number of runspaces that pool creates
        [ValidateRange(1, [int]::MaxValue)]
        [int]$MaxRunspaces = $env:NUMBER_OF_PROCESSORS,
        # PowerShell modules to import to InitialSessionState.
        [string[]]$Modules,
        # Variable to import to InitialSessionState.
        [System.Management.Automation.PSVariable[]]$Variables,
        # Import all non-const script-scoped variables to InitialSessionState.
        [switch]$IncludeScriptVariables
    )

    if ($Script:RunspacePool) {
        return
    }

    Write-Log "Setting up a Runspace Pool with an initialSessionState. MinRunspaces:$MinRunspaces, MaxRunspaces:$MaxRunspaces"
    $initialSessionState = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault()

    # Add functions from this script module. This will find all the functions including non-exported ones.
    # Note: I just want to call "ImportPSModule". It works, but emits "WARNING: The names of some imported commands ...".
    # Just to avoid this, I'm manually adding each command.
    #   $initialSessionState.ImportPSModule($MyInvocation.MyCommand.Module.Path)
    if ($MyInvocation.MyCommand.Module) {
        Get-Command -Module $MyInvocation.MyCommand.Module | & {
            process {
                $initialSessionState.Commands.Add((New-Object System.Management.Automation.Runspaces.SessionStateFunctionEntry($_.Name, $_.ScriptBlock)))
            }
        }
    }

    # Import extra modules.
    if ($Modules) {
        $initialSessionState.ImportPSModule($Modules)
    }

    # Import Script-scoped variables.
    if ($IncludeScriptVariables) {
        Get-Variable -Scope Script | & {
            process {
                if ($_.Options -notmatch 'Constant' -and $_.Value) {
                    $initialSessionState.Variables.Add((New-Object System.Management.Automation.Runspaces.SessionStateVariableEntry -ArgumentList $_.Name, $_.Value, <# description #> $null))
                }
            }
        }
    }

    # Import given variables
    foreach ($_ in $Variables) {
        $initialSessionState.Variables.Add((New-Object System.Management.Automation.Runspaces.SessionStateVariableEntry -ArgumentList $_.Name, $_.Value, <# description #> $null))
    }

    $Script:RunspacePool = [runspacefactory]::CreateRunspacePool($MinRunspaces, $MaxRunspaces, $initialSessionState, $Host)
    $Script:RunspacePool.Open()

    Write-Log "RunspacePool ($($Script:RunspacePool.InstanceId.ToString())) is opened"
}

function Close-TaskRunspace {
    [CmdletBinding()]
    param()

    if (-not $Script:RunspacePool) {
        return
    }

    $id = $Script:RunspacePool.InstanceId.ToString()
    $Script:RunspacePool.Close()
    $Script:RunspacePool = $null
    Write-Log "RunspacePool ($id) is closed"
}

<#
.SYNOPSIS
Start a task to run a command or scriptblock to run asynchronously.

.EXAMPLE
$t = Start-Task { Invoke-LongRunning }
if (Wait-Task $t -Timeout 00:01:00) {
    $t | Receive-Task
}
else {
    Write-Error "Timeout"
}

.EXAMPLE
$t = Start-Task {param ($data) Invoke-LongRunning -Data $data} -ArgumentList $data
Note: Start-Task takes ScriptBlock and ArgumentList, just like Invoke-Command.

.EXAMPLE
Start-Task { Get-ChildItem C:\ } | Receive-Task -AutoRemoveTask
Note: Receive-Task waits for the task to complete and returns the result (and errors too).
#>
function Start-Task {
    [CmdletBinding(PositionalBinding = $false)]
    param (
        # Command to execute.
        [Parameter(ParameterSetName = 'Command', Mandatory = $true, Position = 0)]
        [string]$Command,
        # ScriptBlock to execute.
        [Parameter(ParameterSetName = 'ScriptBlock', Mandatory = $true, Position = 0)]
        [ScriptBlock]$ScriptBlock,
        # ArgumentList (Hashtable or list of argument values)
        $ArgumentList,
        # Optional name of task
        [string]$Name
    )

    if (-not $Script:RunspacePool) {
        Write-Error -Message "Open-TaskRunspace must be called in advance"
        return
    }

    # Create a PowerShell instance and set paramters if any.
    [PowerShell]$ps = [PowerShell]::Create()
    $ps.RunspacePool = $Script:RunspacePool

    switch ($PSCmdlet.ParameterSetName) {
        'Command' {
            $null = $ps.AddCommand($Command)
            break
        }

        'ScriptBlock' {
            $null = $ps.AddScript($ScriptBlock)
            break
        }
    }

    if ($ArgumentList -is [System.Collections.IDictionary]) {
        $null = $ps.AddParameters($ArgumentList)
    }
    else {
        foreach ($arg in $ArgumentList) {
            $null = $ps.AddArgument($arg)
        }
    }

    # Start the command
    $ar = $ps.BeginInvoke()

    # Give a name to this task
    if (-not $Name) {
        if ($Command) {
            $Name = $Command
        }
        else {
            # If ScriptBlock is from a function, its Ast.Name is the function name
            $Name = $ScriptBlock.Ast.Name

            if (-not $Name) {
                if ($ScriptBlock.Ast) {
                    $Name = $ScriptBlock.Ast.ToString()
                }
                else {
                    $Name = "{$ScriptBlock}"
                }
            }
        }
    }

    [PSCustomObject]@{
        AsyncResult  = $ar
        PowerShell   = $ps
        # These are for diagnostic purpose
        ScriptBlock  = $ScriptBlock
        ArgumentList = $ArgumentList
        Name         = $Name
    }
}

<#
.SYNOPSIS
Wait for a task with optional timeout. By default, it waits indefinitely.
It returns the task object if the task completes before the timeout.
When timeout occurs, there is no output.
#>
function Wait-Task {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
        $Task,
        # By default, it waits indefinitely
        # TimeSpan that represents -1 milliseconds is to wait indefinitely.
        [TimeSpan]$Timeout = [System.Threading.Timeout]::InfiniteTimeSpan
    )

    process {
        foreach ($t in $Task) {
            [IAsyncResult]$ar = $t.AsyncResult
            try {
                if ($ar.AsyncWaitHandle.WaitOne($Timeout)) {
                    $t
                }
            }
            catch {
                Write-Error -ErrorRecord $_
            }
        }
    }
}

# Helper function to format an error message from an ErrorRecord of task.
function Format-TaskError {
    [OutputType([string])]
    param(
        $Task,
        [System.Management.Automation.ErrorRecord]$ErrorRecord,
        [switch]$Terminating
    )

    $msg = New-Object System.Text.StringBuilder "Task $($Task.Name) had a $(if (-not $Terminating) {'non-'})terminating error "
    $null = $msg.Append($ErrorRecord.ScriptStackTrace.Split([System.Environment]::NewLine)[0]).Append('; ')

    if ($ErrorRecord.ErrorDetails.Message) {
        $null = $msg.Append($ErrorRecord.ErrorDetails.Message)
    }
    else {
        $null = $msg.Append($ErrorRecord.Exception.Message)
    }

    $msg.ToString()
}

function Receive-Task {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
        $Task,
        [switch]$AutoRemoveTask,
        [string]$TaskErrorVariable
    )

    process {
        foreach ($t in $Task) {
            [powershell]$ps = $t.PowerShell
            [IAsyncResult]$ar = $t.AsyncResult

            # To support Ctrl+C, wake up once in while.
            while ($true) {
                if ($ar.AsyncWaitHandle.WaitOne(2000)) {
                    break
                }
            }

            try {
                $ps.EndInvoke($ar)
            }
            catch {
                # "Real" ErrorRecord is inside InnerException (The outermost exception points to EndInvoke() above)
                $errorMessage = Format-TaskError -Task $t -ErrorRecord $_.Exception.InnerException.ErrorRecord -Terminating
                Write-Error -Message $errorMessage -Exception $_.Exception
            }

            foreach ($_ in $ps.Streams.Error) {
                $errorMessage = Format-TaskError -Task $t -ErrorRecord $_
                Write-Error -Message $errorMessage -Exception $_.Exception
            }

            if ($TaskErrorVariable -and $ps.Streams.Error.Count -gt 0) {
                # Scope 1 is the parent scope, but it's not necessarily the caller scope.
                # If the caller is a function in this module, then scope 1 is the caller function.
                # However, if it's called from outside of module, scope 1 is the module's script scope. Thus the caller does not get the error.
                # Because this function is meant to be moudule-internal and should be called only within the moudle, Scope 1 is ok for now.
                New-Variable -Name $TaskErrorVariable -Value $($ps.Streams.Error.ReadAll()) -Scope 1 -Force
            }

            if ($AutoRemoveTask) {
                Remove-Task $t
            }
        }
    }
}

function Remove-Task {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
        $Task
    )

    process {
        foreach ($t in $Task) {
            [powershell]$ps = $t.PowerShell
            [IAsyncResult]$ar = $t.AsyncResult

            # Note: Disposing PowerShell instance will stop the currently running command & its thread.
            # So there's no need to call EndInvoke() if you don't need the result.
            $ps.Dispose()
            $ar.AsyncWaitHandle.Close()
        }
    }
}

function Stop-Task {
    param (
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
        $Task
    )

    process {
        foreach ($t in $Task) {
            [powershell]$ps = $t.PowerShell
            $ps.Stop()
        }
    }
}

<#
.SYNOPSIS
    Convert a PSPath to a path without prefix (sucn as "Microsoft.PowerShell.Core\FileSystem::", Microsoft.PowerShell.Core\Registry::)
#>
function ConvertFrom-PSPath {
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory, ValueFromPipeline)]
        [string]$Path,
        # Keep provider name such as "FileSystem::" or "Registry::"
        [switch]$KeepProvider
    )

    process {
        if ($Path -match '(?<Prefix>^.*::)(?<Rest>.*)') {
            if ($KeepProvider) {
                $pathWithoutPrefix = $Matches['Rest']
                if ($Matches['Prefix'] -match '(?<Provider>\w+::)') {
                    "$($Matches['Provider'])$pathWithoutPrefix"
                }
            }
            else {
                $Matches['Rest']
            }
        }
    }
}

function Test-RunAsAdministrator {
    [CmdletBinding()]
    [OutputType([bool])]
    param()

    Test-ProcessElevated $PID
}

function Test-ProcessElevated {
    [CmdletBinding()]
    [OutputType([bool])]
    param(
        # Process ID. Support pipeline from both Get-Process & WMI Win32_Process
        [Parameter(Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName, Position = 0)]
        [Alias('ProcessId')]
        # Note about ArgumentCompleter: CompletionResult is not used here, because CompletionResult does not work well for PowerShell (not ISE) when there are a lot of items. It shows "Display all ... possiblities?" (it shows the list, but it ends the command input)
        # [ArgumentCompleter({ Get-Process | Sort-Object Id | Select-Object -ExpandProperty Id })]
        [int]$Id
    )

    begin {
        # https://learn.microsoft.com/en-us/windows/win32/api/winnt/ne-winnt-token_information_class
        $TokenElevation = 20

        $cbSize = [System.Runtime.InteropServices.Marshal]::SizeOf([type]'Win32.Advapi32+TOKEN_ELEVATION')
        $elevation = New-Object Win32.Advapi32+TOKEN_ELEVATION
    }

    process {
        $hProcess = $null
        $hToken = New-Object Win32.SafeProcessTokenHandle

        try {
            $hProcess = [Win32.Kernel32]::OpenProcess([Win32.Kernel32]::PROCESS_QUERY_LIMITED_INFORMATION, $false, $Id)

            if (-not $hProcess -or $hProcess.IsInvalid) {
                $ex = New-Object System.ComponentModel.Win32Exception -ArgumentList ([System.Runtime.InteropServices.Marshal]::GetLastWin32Error())
                Write-Error -Message "OpenProcess failed for PID $Id. $($ex.Message) (NativeErrorCode:$($ex.NativeErrorCode))" -Exception $ex
                return
            }

            if (-not [Win32.Advapi32]::OpenProcessToken($hProcess, [System.Security.Principal.TokenAccessLevels]::Query, [ref]$hToken)) {
                $ex = New-Object System.ComponentModel.Win32Exception -ArgumentList ([System.Runtime.InteropServices.Marshal]::GetLastWin32Error())
                Write-Error -Message "OpenProcessToken failed for PID $Id. $($ex.Message) (NativeErrorCode:$($ex.NativeErrorCode))" -Exception $ex
                return
            }

            if (-not [Win32.Advapi32]::GetTokenInformation(
                    $hToken,
                    $TokenElevation,
                    [ref]$elevation,
                    $cbSize,
                    [ref]$cbSize)) {
                $ex = New-Object System.ComponentModel.Win32Exception -ArgumentList ([System.Runtime.InteropServices.Marshal]::GetLastWin32Error())
                Write-Error -Message "GetTokenInformation failed for PID $Id. $($ex.Message) (NativeErrorCode:$($ex.NativeErrorCode))" -Exception $ex
                return
            }

            # https://learn.microsoft.com/en-us/windows/win32/api/winnt/ns-winnt-token_elevation
            # > "A nonzero value if the token has elevated privileges; otherwise, a zero value"
            $elevation.TokenIsElevated -ne 0
        }
        finally {
            if ($hToken) {
                $hToken.Dispose()
            }

            if ($hProcess) {
                $hProcess.Dispose()
            }
        }
    }
}

function Get-Privilege {
    [CmdletBinding()]
    param()

    # A pseudo handle for the current process (no need to close)
    $hProcess = New-Object Microsoft.Win32.SafeHandles.SafeProcessHandle -ArgumentList (New-Object IntPtr -ArgumentList -1), $false

    $hToken = New-Object Win32.SafeProcessTokenHandle
    $buffer = $null

    try {
        if (-not [Win32.Advapi32]::OpenProcessToken($hProcess, [System.Security.Principal.TokenAccessLevels]::Query, [ref]$hToken)) {
            $ex = New-Object System.ComponentModel.Win32Exception -ArgumentList ([System.Runtime.InteropServices.Marshal]::GetLastWin32Error())
            Write-Error -Message "OpenProcessToken failed for PID $Id. $($ex.Message) (NativeErrorCode:$($ex.NativeErrorCode))" -Exception $ex
            return
        }

        # https://learn.microsoft.com/en-us/windows/win32/api/winnt/ne-winnt-token_information_class
        $TokenPrivileges = 3

        # Get the necessary buffer size (in bytes)
        $cbSize = 0

        $null = [Win32.Advapi32]::GetTokenInformation(
            $hToken,
            $TokenPrivileges,
            [Win32.SafeCoTaskMemFreeHandle]::InvalidHandle,
            $cbSize,
            [ref]$cbSize)

        # Expected to get ERROR_INSUFFICIENT_BUFFER
        $ERROR_INSUFFICIENT_BUFFER = 0x7a
        $errorCode = [System.Runtime.InteropServices.Marshal]::GetLastWin32Error()

        if ($errorCode -ne $ERROR_INSUFFICIENT_BUFFER) {
            $ex = New-Object System.ComponentModel.Win32Exception -ArgumentList ([System.Runtime.InteropServices.Marshal]::GetLastWin32Error())
            Write-Error -Message "GetTokenInformation failed. $($ex.Message) (NativeErrorCode:$($ex.NativeErrorCode))" -Exception $ex
            return
        }

        # Allocate buffer & retry
        # Note: The following line does not work in PSv4 due to an overload resolutuion issue:
        #
        #   $buffer = New-Object Win32.SafeCoTaskMemFreeHandle -ArgumentList $cbSize
        #
        # It fails with IndexOutOfRangeException at System.Management.Automation.Adapter.CompareOverloadCandidates()
        # To workaround, manually assign a raw IntPtr to SafeCoTaskMemFreeHandle
        $buffer = New-Object Win32.SafeCoTaskMemFreeHandle
        $buffer.Reset([System.Runtime.InteropServices.Marshal]::AllocCoTaskMem($cbSize))

        if (-not [Win32.Advapi32]::GetTokenInformation(
                $hToken,
                $TokenPrivileges,
                $buffer,
                $cbSize,
                [ref]$cbSize)) {
            $ex = New-Object System.ComponentModel.Win32Exception -ArgumentList ([System.Runtime.InteropServices.Marshal]::GetLastWin32Error())
            Write-Error -Message "GetTokenInformation failed. $($ex.Message) (NativeErrorCode:$($ex.NativeErrorCode))" -Exception $ex
            return
        }

        # Map the pointer to TOKEN_PRIVILEGES
        # For usage of [Type], see https://support.microsoft.com/en-us/topic/exceptions-in-windows-powershell-other-dynamic-languages-and-dynamically-executed-c-code-when-code-that-targets-the-net-framework-calls-some-methods-680ca719-0782-1052-7999-183d00e9cc93
        $privileges = [System.Runtime.InteropServices.Marshal]::PtrToStructure($buffer.DangerousGetHandle(), [Type][Win32.Advapi32+TOKEN_PRIVILEGES])

        # This is the pointer to LUID_ATTRIBUTES[PrivilegeCount].
        # LUID_ATTRIBUTES[] is at offset 4 of TOKEN_PRIVILEGES (after "DWORD PrivilegeCount").
        $pPrivilges = [IntPtr]::Add($buffer.DangerousGetHandle(), 4)

        # Size of each privilege
        $privSize = [System.Runtime.InteropServices.Marshal]::SizeOf([Type][Win32.Advapi32+LUID_AND_ATTRIBUTES])

        $nameBuffer = New-Object char[] -ArgumentList 256
        $SE_PRIVILEGE_ENABLED = 0x00000002

        for ($i = 0; $i -lt $privileges.PrivilegeCount; ++$i) {
            # Move to the current privilege (type is LUID_AND_ATTRIBUTES)
            $pCurrent = [IntPtr]::Add($pPrivilges, $i * $privSize)
            $priv = [System.Runtime.InteropServices.Marshal]::PtrToStructure($pCurrent, [Type][Win32.Advapi32+LUID_AND_ATTRIBUTES])

            # Get Privilege name
            $cchName = $nameBuffer.Length

            if (-not [Win32.Advapi32]::LookupPrivilegeNameW($null, $pCurrent, $nameBuffer, [ref]$cchName)) {
                Write-Error "LookupPrivilegeNameW failed for $($pCurrent.Luid) with $([System.Runtime.InteropServices.Marshal]::GetLastWin32Error())"
                continue
            }

            $privName = New-Object string -ArgumentList $nameBuffer, 0, $cchName
            $isEnabled = ($priv.Attributes -band $SE_PRIVILEGE_ENABLED) -ne 0

            [PSCustomObject]@{
                Name    = $privName
                Enabled = $isEnabled
            }
        }
    }
    finally {
        if ($buffer) {
            $buffer.Dispose()
        }

        if ($hToken) {
            $hToken.Dispose()
        }
    }
}

function Test-DebugPrivilege {
    [CmdletBinding()]
    [OutputType([bool])]
    param()

    $debugPrivilege = Get-Privilege | Where-Object { $_.Name -eq 'SeDebugPrivilege' }

    if (-not $debugPrivilege) {
        Write-Verbose "The user does not have DebugPrivilege"
        $false
    }
    elseif (-not $debugPrivilege.Enabled) {
        Write-Verbose "DebugPrivilege is not enabled"
        $false
    }
    else {
        $true
    }
}

function Enable-DebugPrivilege {
    [CmdletBinding()]
    param()

    try {
        [System.Diagnostics.Process]::EnterDebugMode()
    }
    catch {
        $msg = "Failed to enable Debug Privilege"
        $runAsAdmin = Test-RunAsAdministrator

        if (-not $runAsAdmin) {
            $msg += " because the process is not running as Administrator"
        }
        else {
            $debugPrivilege = Get-Privilege | Where-Object { $_.Name -eq 'SeDebugPrivilege' }

            if (-not $debugPrivilege) {
                $msg += " because the user does not have DebugPrivilege"
            }
        }

        Write-Error -Message $msg -Exception $_.Exception
    }
}

function Disable-DebugPrivilege {
    [CmdletBinding()]
    param()

    try {
        [System.Diagnostics.Process]::LeaveDebugMode()
    }
    catch {
        Write-Error -Message "Failed to disable Debug Privilege. $_" -Exception $_.Exception
    }
}

function Download-File {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseApprovedVerbs', '')]
    [CmdletBinding(PositionalBinding = $false)]
    param (
        [Parameter(Mandatory, Position = 0)]
        # Uri to download
        [string]$Uri,
        [Parameter(Mandatory)]
        # Destination file path
        [string]$OutFile,
        # Used for Write-Progress's Activity parameter. If skipped, no progress report.
        [string]$Activity
    )

    $webClient = $null
    $progressChangeId = $null
    $completedId = $null

    $dir = Split-Path $OutFile

    # Directory must exist otherwize WebClient.DownloadFile(Async) fails.
    if (-not (Test-Path $dir)) {
        $null = New-Item -Path $dir -ItemType Directory -ErrorAction Stop
    }

    try {
        $webClient = New-Object System.Net.WebClient
        $webClient.UseDefaultCredentials = $true
        $start = Get-TimeStamp

        if (-not $Activity) {
            $webClient.DownloadFile($Uri, $OutFile)
        }
        else {
            Write-Progress -Activity $Activity -Status "Starting" -PercentComplete 0
            $waitInterval = [TimeSpan]::FromSeconds(1)
            $progressChangeId = [Guid]::NewGuid().ToString()
            $completedId = [Guid]::NewGuid().ToString()

            Register-ObjectEvent -InputObject $webClient -EventName 'DownloadProgressChanged' -SourceIdentifier $progressChangeId
            Register-ObjectEvent -InputObject $webClient -EventName 'DownloadFileCompleted' -SourceIdentifier $completedId

            $webClient.DownloadFileAsync($Uri, $OutFile)

            while ($true) {
                [System.ComponentModel.AsyncCompletedEventArgs]$completed = Get-Event -SourceIdentifier $completedId -ErrorAction SilentlyContinue `
                | Select-Object -Last 1 -ExpandProperty 'SourceEventArgs'

                if ($completed) {
                    if ($completed.Cancelled -or $completed.Error) {
                        # Remove imcomplete file.
                        Remove-Item -Path $OutFile -ErrorAction SilentlyContinue

                        if ($completed.Cancelled) {
                            Write-Error -Message "Download was cancelled"
                        }
                        elseif ($completed.Error) {
                            Write-Error -Message "WebClient.DownloadFileAsync() failed. $($completed.Error)" -Exception $completed.Error
                        }

                        return
                    }

                    break
                }

                [System.Net.DownloadProgressChangedEventArgs]$progressChanged = Get-Event -SourceIdentifier $progressChangeId -ErrorAction SilentlyContinue `
                | Select-Object -Last 1 -ExpandProperty 'SourceEventArgs'

                if ($progressChanged) {
                    Write-Progress -Activity $Activity -Status "$($progressChanged.ProgressPercentage)% ($($progressChanged.BytesReceived)/$($progressChanged.TotalBytesToReceive))" -PercentComplete $progressChanged.ProgressPercentage
                }

                Start-Sleep -Seconds $waitInterval.TotalSeconds
            }
        }

        [PSCustomObject]@{
            Uri     = $Uri
            OutFile = $OutFile
            Elapsed = Get-Elapsed $start
        }
    }
    catch {
        Write-Error -Message "Failed to download $Uri. $_" -Exception $_.Exception
    }
    finally {
        if ($Activity) {
            Write-Progress -Activity $Activity -Status "Done" -Completed
        }

        if ($progressChangeId) {
            Unregister-Event -SourceIdentifier $progressChangeId -ErrorAction SilentlyContinue
        }

        if ($completedId) {
            Unregister-Event -SourceIdentifier $completedId -ErrorAction SilentlyContinue
        }

        if ($webClient) {
            $webClient.Dispose()
        }
    }
}
<#
.SYNOPSIS
    Save files
.NOTES
    When Copy-Item fails on a single file, it fails the entire operation. This function will continue to copy other files.
#>
function Save-Item {
    [CmdletBinding(PositionalBinding = $false)]
    param(
        [Parameter(Mandatory, Position = 0)]
        # Path to the items to copy
        [string]$Path,
        # Destination folder path (Renaming by a file path is not supported)
        [Parameter(Mandatory)]
        [string]$Destination,
        # Filter (Forwarded to Get-ChildItem)
        [string]$Filter,
        # Include (Forwarded to Get-ChildItem)
        [string[]]$Include,
        # Exclude (Forwarded to Get-ChildItem)
        [string[]]$Exclude,
        # Include hidden or system files(Add -Force to Get-ChildItem)
        [switch]$IncludeHidden,
        # Recurse (Forwarded to Get-ChildItem)
        [switch]$Recurse,
        [switch]$PassThru
    )

    # Get source directory when recursing so that I know when to create a subdirectory in the destination
    if ($Recurse) {
        $srcItem = Get-Item -Path $Path -ErrorAction SilentlyContinue | Select-Object -First 1

        if (-not $srcItem) {
            Write-Error "$Path is not found"
            return
        }

        if ($srcItem.PSIsContainer) {
            $src = $srcItem.FullName
        }
        else {
            $src = $srcItem.DirectoryName
        }

        $src = $src.TrimEnd('\')
    }

    # Get destination full path.
    # Do not use Covert-Path because it fails when the path does not exist.
    # Do not just use [IO.Path]::GetFullPath() only because it resolves to .NET's current directory, not PowerShell's.
    $Destination = [IO.Path]::GetFullPath([IO.Path]::Combine($PWD.ProviderPath, $Destination)).TrimEnd('\')

    # Without Recurse, Path needs a trailing * in order to use Include or Exclude
    if (-not $Recurse -and ($Include -or $Exclude)) {
        $leaf = Split-Path $Path -Leaf

        if ($leaf.IndexOf('*') -eq -1) {
            $Path = Join-Path $Path -ChildPath '*'
        }
    }

    Get-ChildItem -Path $Path -Filter $Filter -Include $Include -Exclude $Exclude -Force:$IncludeHidden -Recurse:$Recurse -File | & {
        param(
            [Parameter(ValueFromPipeline)]
            [System.IO.FileInfo]$file
        )
        process {
            # If this file is from the destination directory, skip it.
            if ($file.DirectoryName.StartsWith($Destination)) {
                Write-Verbose "Skipping $($file.FullName)"
                return
            }

            $dest = $Destination

            if ($Recurse) {
                $childPath = $file.DirectoryName.SubString($src.Length)

                if ($childPath) {
                    $dest = Join-Path $Destination -ChildPath $childPath
                }
            }

            if (-not (Test-Path $dest)) {
                $null = New-Item -ItemType Directory -Path $dest -ErrorAction Stop
            }

            try {
                Copy-Item -LiteralPath $file.FullName -Destination $dest -PassThru:$PassThru
            }
            catch {
                Write-Error -Message "Failed to copy '$($file.FullName)'. $_" -Exception $_.Exception
            }
        }
    }
}

function Compress-Folder {
    [CmdletBinding()]
    param(
        # Folder path to compress
        [Parameter(Mandatory = $true)]
        [string]$Path,
        # Destination folder path
        [Parameter(Mandatory = $true)]
        [string]$Destination,
        # Filter for items in $Path
        [string[]]$Filter,
        # DateTime filters
        [DateTime]$FromDateTime,
        [DateTime]$ToDateTime,
        [ValidateSet('Zip', 'Cab')]
        [string]$ArchiveType = 'Zip'
    )

    <#
    .SYNOPSIS
        Create a Zip file using .NET's System.IO.Compression.
    #>
    function New-Zip {
        [CmdletBinding()]
        param(
            # Folder path to compress
            [Parameter(Mandatory = $true)]
            [string]$Path,
            # Destination folder path
            [Parameter(Mandatory = $true)]
            [string]$Destination,
            # Filter for items in $Path
            [string[]]$Filter,
            # DateTime filters
            [DateTime]$FromDateTime,
            [DateTime]$ToDateTime
        )

        if (Test-Path $Path) {
            $Path = Convert-Path -LiteralPath $Path
        }
        else {
            Write-Error "$Path is not found"
            return
        }

        if (-not (Get-Item $Path).PSIsContainer) {
            Write-Error "$Path is not a container"
            return
        }

        # Check if .NET Framework's compression is avaiable.
        try {
            Add-Type -AssemblyName System.IO.Compression -ErrorAction Stop
        }
        catch {
            Write-Error -Message "System.IO.Compression is not available. $_" -Exception $_.Exception
            return
        }

        if (Test-Path $Destination) {
            $Destination = Convert-Path -LiteralPath $Destination
        }
        else {
            $Destination = New-Item $Destination -ItemType Directory -ErrorAction Stop | Select-Object -ExpandProperty FullName
        }

        $files = @(
            $Filter | & {
                # Apply filename filters if any. Note: Even if Filter is null, the pipeline will run (unlike foreach keyword)
                param ([Parameter(ValueFromPipeline)]$filter)
                process {
                    Get-ChildItem -LiteralPath $Path -File -Recurse -Force -Filter $filter
                }
            } | & {
                param ([Parameter(ValueFromPipeline)]$file)
                process {
                    if ($FromDateTime -and $file.LastWriteTime -lt $FromDateTime) {
                        return
                    }

                    if ($ToDateTime -and $file.LastWriteTime -gt $ToDateTime) {
                        return
                    }

                    $file
                }
            }
        )

        # If there are no files after filters are applied, bail.
        if ($files.Count -eq 0) {
            Write-Error "There are no files after filters are applied. Server: $env:COMPUTERNAME, Path: $Path, Filter: $Filter, FromDateTime: $FromDateTime, ToDateTime: $ToDateTime"
            return
        }

        # Create a ZIP file
        $zipFileName = Split-Path $Path -Leaf
        $zipFilePath = Join-Path $Destination -ChildPath "$zipFileName.zip"

        if (Test-Path $zipFilePath) {
            # Append a random string to the zip file name.
            $zipFileName = $zipFileName + "_" + [IO.Path]::GetRandomFileName().Substring(0, 8) + '.zip'
            $zipFilePath = Join-Path $Destination $zipFileName
        }

        $zipStream = $zipArchive = $null

        try {
            $null = New-Item $zipFilePath -ItemType file

            $zipStream = New-Object System.IO.FileStream -ArgumentList $zipFilePath, ([IO.FileMode]::Open)
            $zipArchive = New-Object System.IO.Compression.ZipArchive -ArgumentList $zipStream, ([IO.Compression.ZipArchiveMode]::Create)

            $progressInterval = 10
            $prevProgress = - $progressInterval
            $activity = "Creating a zip file $zipFilePath"

            [long]$totalBytes = $files | Measure-Object -Property 'Length' -Sum | Select-Object -ExpandProperty 'Sum'
            [long]$currentBytes = 0

            foreach ($file in $files) {
                $progress = 100 * $currentBytes / $totalBytes

                if ($progress -ge $prevProgress + $progressInterval) {
                    Write-Progress -Activity $activity -Status "Please wait" -PercentComplete $progress
                    $prevProgress = $progress
                }

                $fileStream = $zipEntryStream = $null

                try {
                    $fileStream = New-Object System.IO.FileStream -ArgumentList $file.FullName, ([IO.FileMode]::Open), ([IO.FileAccess]::Read), ([IO.FileShare]::ReadWrite)
                    $zipEntry = $zipArchive.CreateEntry($file.FullName.Substring($Path.TrimEnd('\').Length + 1))
                    $zipEntryStream = $zipEntry.Open()
                    $fileStream.CopyTo($zipEntryStream)

                    $currentBytes += $file.Length
                }
                catch {
                    Write-Error -Message "Failed to add $($file.FullName). $_" -Exception $_.Exception
                }
                finally {
                    if ($fileStream) {
                        $fileStream.Dispose()
                    }

                    if ($zipEntryStream) {
                        $zipEntryStream.Dispose()
                    }
                }
            }
        }
        finally {
            if ($zipArchive) {
                $zipArchive.Dispose()
            }

            if ($zipStream) {
                $zipStream.Dispose()
            }

            Write-Progress -Activity $activity -Status "Done" -Completed
            $archivePath = $zipFilePath
        }

        New-Object PSCustomObject -Property @{
            ArchivePath = $archivePath
        }
    }

    <#
    .SYNOPSIS
        Create a Zip file using Shell.Application COM
    #>
    function New-ZipShell {
        [CmdletBinding()]
        param(
            # Folder path to compress
            [Parameter(Mandatory = $true)]
            [string]$Path,
            # Destination folder path
            [Parameter(Mandatory = $true)]
            [string]$Destination,
            # Filter for items in $Path
            [string[]]$Filter,
            # DateTime filters
            [DateTime]$FromDateTime,
            [DateTime]$ToDateTime
        )

        if (Test-Path $Path) {
            $Path = Convert-Path -LiteralPath $Path
        }
        else {
            Write-Error "$Path is not found"
            return
        }

        if (-not (Get-Item $Path).PSIsContainer) {
            Write-Error "$Path is not a container"
            return
        }

        if (Test-Path $Destination) {
            $Destination = Convert-Path -LiteralPath $Destination
        }
        else {
            $Destination = New-Item $Destination -ItemType Directory -ErrorAction Stop | Select-Object -ExpandProperty FullName
        }

        # If there are no filters to apply, archive the given Path.
        # Otherwise, apply filters and copy the filtered files to a temporary path and archive it.
        if (-not $PSBoundParameters.ContainsKey('Filter') -and -not $PSBoundParameters.ContainsKey('FromDateTime') -and -not $PSBoundParameters.ContainsKey('ToDateTime')) {
            $targetPath = $Path
        }
        else {
            # Apply filters.
            if ($Filter.Count) {
                $files = @(foreach ($f in $Filter) { Get-ChildItem -LiteralPath $Path -Filter $f -Recurse -Force | Where-Object { -not $_.PSIsContainer } })
            }
            else {
                $files = @(Get-ChildItem -LiteralPath $Path -Recurse -Force | Where-Object { -not $_.PSIsContainer })
            }

            if ($PSBoundParameters.ContainsKey('FromDateTime') -and $FromDateTime -ne [DateTime]::MinValue) {
                $files = @($files | Where-Object { $_.LastWriteTime -ge $FromDateTime })
            }

            if ($PSBoundParameters.ContainsKey('ToDateTime') -and $ToDateTime -ne [DateTime]::MaxValue) {
                $files = @($files | Where-Object { $_.LastWriteTime -le $ToDateTime })
            }

            # Remove duplicate by Fullname
            $files = @($files | Group-Object -Property 'FullName' | ForEach-Object { $_.Group | Select-Object -First 1 })

            # If there are no files after filters are applied, bail.
            if ($files.Count -eq 0) {
                Write-Error "There are no files after filters are applied. Server:$env:COMPUTERNAME, Path:$Path, Filter:$Filter, FromDateTime:$FromDateTime, ToDateTime:$ToDateTime"
                return
            }

            # Copy filtered files to a temporary folder
            $tempPath = Join-Path $env:TEMP ([IO.Path]::GetRandomFileName().Substring(0, 8))
            $null = New-Item $tempPath -ItemType Directory

            foreach ($file in $files) {
                $dest = $tempPath
                $subPath = $file.DirectoryName.SubString($Path.Length)
                if ($subPath) {
                    $dest = Join-Path $tempPath $subPath
                    if (-not (Test-Path -LiteralPath $dest)) {
                        $null = New-Item -ItemType Directory -Path $dest
                    }
                }

                try {
                    Copy-Item -LiteralPath $file.FullName -Destination $dest
                }
                catch {
                    Write-Error -Message "Failed to copy $($file.FullName) to a temporary path $dest. $_" -Exception $_.Exception
                }
            }

            $dest = $null
            $targetPath = $tempPath
        }

        Write-Verbose "targetPath:$targetPath"

        # Form the zip file name
        $archiveName = Split-Path $Path -Leaf
        $archivePath = Join-Path $Destination -ChildPath "$archiveName.zip"

        if (Test-Path $archivePath) {
            # Append a random string to the zip file name.
            $archiveName = $archiveName + "_" + [IO.Path]::GetRandomFileName().Substring(0, 8) + '.zip'
            $archivePath = Join-Path $Destination $archiveName
        }

        # Use Shell.Application COM.
        # Create a Zip file manually
        $shellApp = New-Object -ComObject Shell.Application
        Set-Content -LiteralPath $archivePath -Value ("PK" + [char]5 + [char]6 + ("$([char]0)" * 18)) -Encoding Ascii
        (Get-Item $archivePath).IsReadOnly = $false
        $zipFile = $shellApp.NameSpace($archivePath)

        $zipFile.CopyHere($targetPath)

        # Now wait and poll
        $delayMs = 200
        $inProgress = $true
        [System.IO.FileStream]$fileStream = $null
        #Start-Sleep -Milliseconds 3000

        while ($inProgress) {
            Start-Sleep -Milliseconds $delayMs

            $fileStream = $null

            try {
                $fileStream = [IO.File]::Open($archivePath, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::None)
                $inProgress = $false
            }
            catch {
                # ignore
            }
            finally {
                if ($fileStream) {
                    $fileStream.Dispose()
                }
            }
        }

        if ($tempPath) {
            Remove-Item -LiteralPath $tempPath -Force -Recurse
        }

        $null = [System.Runtime.Interopservices.Marshal]::FinalReleaseComObject($shellApp)

        New-Object PSCustomObject -Property @{
            ArchivePath = $archivePath
        }
    }

    # https://docs.microsoft.com/en-us/windows-server/administration/windows-commands/makecab
    # https://docs.microsoft.com/en-us/previous-versions/bb417343(v=msdn.10)
    function New-Cab {
        [CmdletBinding()]
        param(
            # Folder path to compress
            [Parameter(Mandatory = $true)]
            [string]$Path,
            # Destination folder path
            [Parameter(Mandatory = $true)]
            [string]$Destination,
            # Filter for items in $Path
            [string[]]$Filter,
            # DateTime filters
            [DateTime]$FromDateTime,
            [DateTime]$ToDateTime,
            [ValidateSet('MSZIP', 'LZX')]
            [string]$CompressionType = 'LZX'
        )

        if (Test-Path -LiteralPath $Path) {
            $Path = Convert-Path -LiteralPath $Path
        }
        else {
            Write-Error "Failed to find $Path"
            return
        }

        if (-not (Get-Item $Path).PSIsContainer) {
            Write-Error "$Path is not a container"
            return
        }

        if (Test-Path $Destination) {
            $Destination = Convert-Path -LiteralPath $Destination
        }
        else {
            $Destination = New-Item $Destination -ItemType Directory -ErrorAction Stop | Select-Object -ExpandProperty FullName
        }

        if ($Filter.Count) {
            $files = @(foreach ($f in $Filter) { Get-ChildItem -LiteralPath $Path -Filter $f -Recurse -Force | Where-Object { -not $_.PSIsContainer } })
        }
        else {
            $files = @(Get-ChildItem -LiteralPath $Path -Recurse -Force | Where-Object { -not $_.PSIsContainer })
        }

        if ($PSBoundParameters.ContainsKey('FromDateTime') -and $FromDateTime -ne [DateTime]::MinValue) {
            $files = @($files | Where-Object { $_.LastWriteTime -ge $FromDateTime })
        }

        if ($PSBoundParameters.ContainsKey('ToDateTime') -and $ToDateTime -ne [DateTime]::MaxValue) {
            $files = @($files | Where-Object { $_.LastWriteTime -le $ToDateTime })
        }

        # Remove duplicate by Fullname
        $files = @($files | Group-Object -Property 'FullName' | ForEach-Object { $_.Group | Select-Object -First 1 })

        if ($files.Count -eq 0) {
            Write-Error "There are no files after filters are applied. Server:$env:COMPUTERNAME, Path:$Path, Filter:$Filter, FromDateTime:$FromDateTime, ToDateTime:$ToDateTime"
            return
        }

        # Create a directive file (ddf)
        $ddfFile = Join-Path $env:TEMP $([IO.Path]::GetRandomFileName().Substring(0, 8) + ".ddf")
        $ddfStream = [IO.File]::OpenWrite($ddfFile)
        $ddfStream.Position = 0
        $ddfWriter = New-Object System.IO.StreamWriter($ddfStream)
        $ddfWrittenCount = 0
        $currentDir = $Path

        foreach ($file in $files) {
            # Make sure the file not locked by another process. Otherwise makecab would fail.
            $skip = $false
            try {
                $fileStream = [IO.File]::OpenRead($file.FullName)
            }
            catch {
                $skip = $true
            }
            finally {
                if ($fileStream) {
                    $fileStream.Dispose()
                }
            }

            if ($skip) {
                continue
            }

            if ($file.DirectoryName -ne $currentDir) {

                $subPath = $file.DirectoryName.SubString($Path.TrimEnd('\').Length + 1)
                $ddfWriter.WriteLine(".Set DestinationDir=`"$subPath`"")
                $currentDir = $file.DirectoryName
            }

            $ddfWriter.WriteLine("`"$($file.FullName)`"")
            $ddfWrittenCount++
        }

        if ($ddfWriter) {
            $ddfWriter.Dispose()
        }

        # There are no files to archive. This is not necessarily an error, but write as an error for the caller.
        if ($ddfWrittenCount -eq 0) {
            Write-Error -Message "There are $($files.Count) files in $Path, but none can be opened"
            return
        }

        $cabName = Split-Path $Path -Leaf
        $cabFilePath = Join-Path $Destination -ChildPath "$cabName.cab"

        if (Test-Path $cabFilePath) {
            # Append a random string to the cab file name.
            $cabName = $cabName + "_" + [IO.Path]::GetRandomFileName().Substring(0, 8)
            $cabFilePath = Join-Path $Destination "$cabName.cab"
        }

        Write-Progress -Activity "Creating a cab file" -Status "Please wait" -PercentComplete -1
        $err = $($stdout = & makecab.exe /D CompressionType=$CompressionType /D CabinetNameTemplate="$cabName.cab" /D DiskDirectoryTemplate=CDROM /D DiskDirectory1=$Destination /D MaxDiskSize=0 /D RptFileName=nul /D InfFileName=nul /F $ddfFile) 2>&1
        Remove-Item $ddfFile -Force
        Write-Progress -Activity "Creating a cab file" -Status "Done" -Completed

        if ($LASTEXITCODE -ne 0) {
            Write-Error "MakeCab.exe failed; exitCode:$LASTEXITCODE; stdout:`"$stdout`"; Error:$err"
            return
        }

        New-Object PSCustomObject -Property @{
            ArchivePath = $cabFilePath
            # Message = $stdout
        }
    }

    # Here's main body of Compress-Folder
    if ($FromDateTime -and $ToDateTime -and $FromDateTime -gt $ToDateTime) {
        Write-Error "FromDateTime must be less than or equal to ToDateTime"
        return
    }

    switch ($ArchiveType) {
        'Zip' {
            if ($PSVersionTable.PSVersion.Major -gt 2) {
                $compressCmd = Get-Command 'New-Zip'
            }
            else {
                $compressCmd = Get-Command 'New-ZipShell'
            }
            break
        }
        'Cab' {
            $compressCmd = Get-Command 'New-Cab'
            break
        }
    }

    $params = @{}

    foreach ($_ in $PSBoundParameters.GetEnumerator()) {
        if ($compressCmd.Parameters.ContainsKey($_.Key)) {
            $params.Add($_.Key, $_.Value)
        }
    }

    & $compressCmd @params
}

<#
.SYNOPSIS
    Enable Event Log. This command returns a previous configuration
#>
function Enable-EventLog {
    [CmdletBinding(PositionalBinding = $false)]
    param(
        [Parameter(Mandatory = $true, Position = 0)]
        [string]$EventName,
        # Max size in bytes (Default:20MB)
        [int]$MaxSize = 20971520
    )

    # Get the current configuration
    $([xml]$configXml = wevtutil.exe get-log $EventName /format:xml) 2>&1 | ForEach-Object {
        Write-Error -ErrorRecord $_
    }

    if (-not $configXml) {
        return
    }

    # Values are not parsed and kept as string on purpose
    $config = [PSCustomObject]@{
        Name       = $EventName
        Enabled    = $configXml.channel.enabled
        Retention  = $configXml.channel.logging.retention
        AutoBackup = $configXml.channel.logging.autoBackup
        MaxSize    = $configXml.channel.logging.maxSize
    }

    $errs = $($null = wevtutil.exe set-log $EventName /enabled:true /retention:false /maxsize:$MaxSize /quiet:true) 2>&1

    if ($errs) {
        $errs | ForEach-Object { Write-Error -ErrorRecord $_ }
        return
    }

    $config
}

function Disable-EventLog {
    [CmdletBinding(PositionalBinding = $false)]
    param(
        [Parameter(Mandatory = $true, Position = 0)]
        [string]$EventName
    )

    $($null = wevtutil.exe set-log $EventName /enabled:false /retention:false /quiet:true) 2>&1 | ForEach-Object {
        Write-Error -ErrorRecord $_
    }
}

<#
.SYNOPSIS
    Restore Event Log configuration using the output of Enable-EventLog
#>
function Restore-EventLog {
    [CmdletBinding(PositionalBinding = $false)]
    param(
        [Parameter(Mandatory, ValueFromPipeline)]
        $Configuration
    )

    process {
        $evtArgs = @(
            'set-log'
            $Configuration.Name

            # Use interporated strings here so that there is no space inside "/key:value"
            "/enabled:$($Configuration.Enabled)"
            "/retention:$($Configuration.Retention)"
            "/maxsize:$($Configuration.MaxSize)"
            '/quiet:true'
        )

        $errs = $($null = wevtutil.exe $evtArgs) 2>&1

        if ($errs) {
            $err | ForEach-Object { Write-Error -ErrorRecord $_ }
            return
        }

        Write-Log "Eveng Log config is restored for $($Configuration.Name). Enabled:$($Configuration.Enabled), MaxSize:$($Configuration.MaxSize)"
    }
}

function Add-EventLogConfigCache {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, ValueFromPipeline)]
        $Configuration
    )

    begin {
        if ($null -eq $Script:EventLogConfigCache) {
            $Script:EventLogConfigCache = @{}
        }
    }

    process {
        if (-not $Script:EventLogConfigCache.ContainsKey($_.Name)) {
            $Script:EventLogConfigCache.Add($_.Name, $_)
        }
    }
}

function Get-EventLogConfigCache {
    [CmdletBinding()]
    param()

    if ($Script:EventLogConfigCache.Count -gt 0) {
        $Script:EventLogConfigCache.Values
    }
}

function Clear-EventLogConfigCache {
    if ($Script:EventLogConfigCache) {
        $Script:EventLogConfigCache.Clear()
    }
}

function Enable-WamEventLog {
    [CmdletBinding(PositionalBinding = $false)]
    param()

    & {
        Enable-EventLog 'Microsoft-Windows-WebAuth/Operational'
        Enable-EventLog 'Microsoft-Windows-WebAuthN/Operational'
        Enable-EventLog 'Microsoft-Windows-AAD/Operational'
        Enable-EventLog 'Microsoft-Windows-AAD/Analytic'
    } | Add-EventLogConfigCache
}

function Disable-WamEventLog {
    [CmdletBinding(PositionalBinding = $false)]
    param()

    Disable-EventLog 'Microsoft-Windows-WebAuth/Operational'
    Disable-EventLog 'Microsoft-Windows-WebAuthN/Operational'
    Disable-EventLog 'Microsoft-Windows-AAD/Analytic'
    Disable-EventLog 'Microsoft-Windows-AAD/Operational'
}

function Start-WamTrace {
    [CmdletBinding(PositionalBinding = $false)]
    param(
        [Parameter(Mandatory = $true, Position = 0)]
        [string]$Path,
        [string]$FileName = "WAM_$(Get-DateTimeString).etl",
        [string]$SessionName = 'WamTrace',
        [ValidateSet('NewFile', 'Circular')]
        [string]$LogFileMode = 'NewFile',
        [ValidateRange(1, [int]::MaxValue)]
        [int]$MaxFileSizeMB = 256
    )

    if (-not (Test-Path $Path)) {
        $null = New-Item $Path -ItemType Directory -ErrorAction Stop
    }

    $Path = Convert-Path -LiteralPath $Path

    # Create a provider listing
    $providerFile = Join-Path $Path -ChildPath 'wam.prov'
    Set-Content -LiteralPath $providerFile -Value $WamProviders -Encoding Ascii -ErrorAction Stop

    switch ($LogFileMode) {
        'NewFile' {
            $mode = @([Win32.Logman+Mode]::EVENT_TRACE_USE_GLOBAL_SEQUENCE, [Win32.Logman+Mode]::EVENT_TRACE_FILE_MODE_NEWFILE) -join ','

            # In order to use newfile, file name must contain "%d"
            if ($FileName -notlike "*%d*") {
                $FileName = [System.IO.Path]::GetFileNameWithoutExtension($FileName) + "_%d.etl"
            }
            break
        }

        'Circular' {
            $mode = @([Win32.Logman+Mode]::EVENT_TRACE_USE_GLOBAL_SEQUENCE, [Win32.Logman+Mode]::EVENT_TRACE_FILE_MODE_CIRCULAR) -join ','

            if (-not $PSBoundParameters.ContainsKey('MaxFileSizeMB')) {
                $MaxFileSizeMB = 2048
            }
            break
        }
    }

    $traceFile = Join-Path $Path -ChildPath $FileName

    Write-Log "Starting a WAM trace"
    $err = $($stdout = Invoke-Command {
            $ErrorActionPreference = 'Continue'
            & logman.exe start trace $SessionName -pf $providerFile -o $traceFile -bs 128 -max $MaxFileSizeMB -mode $mode -ets
        }) 2>&1

    if ($err -or $LASTEXITCODE -ne 0) {
        Write-Error "logman failed. exit code:$LASTEXITCODE; stdout:`"$stdout`"; error:`"$err`""
        return
    }
}

function Stop-WamTrace {
    [CmdletBinding()]
    param(
        $SessionName = 'WamTrace'
    )

    Write-Log "Stopping $SessionName"
    $null = Stop-EtwSession $SessionName
}

function Start-OutlookTrace {
    [CmdletBinding(SupportsShouldProcess = $true, PositionalBinding = $false)]
    param(
        [Parameter(Mandatory = $true, Position = 0)]
        [string]$Path,
        [string]$FileName = "Outlook_$(Get-DateTimeString).etl",
        [string]$SessionName = 'OutlookTrace',
        [ValidateSet('NewFile', 'Circular')]
        [string]$LogFileMode = 'NewFile',
        [ValidateRange(1, [int]::MaxValue)]
        [int]$MaxFileSizeMB = 256
    )

    if (-not (Test-Path $Path)) {
        $null = New-Item $Path -ItemType Directory -ErrorAction Stop
    }

    $Path = Convert-Path -Literal $Path
    $providerFile = Join-Path $Path -ChildPath 'Office.prov'
    $officeInfo = Get-OfficeInfo -ErrorAction Stop
    $major = $officeInfo.Version.Split('.')[0] -as [int]
    Write-Log "Creating a provider listing according to the version $major"

    $providers = switch ($major) {
        14 { $Outlook2010Providers; break }
        15 { $Outlook2013Providers; break }
        16 { $Outlook2016Providers; break }
        default { throw "Couldn't find the version from $_" }
    }

    Set-Content -LiteralPath $providerFile -Value $providers -Encoding Ascii -ErrorAction Stop

    # Configure log file mode, filename, and max file size if necessary.
    switch ($LogFileMode) {
        'NewFile' {
            $mode = @([Win32.Logman+Mode]::EVENT_TRACE_USE_GLOBAL_SEQUENCE, [Win32.Logman+Mode]::EVENT_TRACE_FILE_MODE_NEWFILE) -join ','

            # In order to use newfile, file name must contain "%d"
            if ($FileName -notlike "*%d*") {
                $FileName = [System.IO.Path]::GetFileNameWithoutExtension($FileName) + "_%d.etl"
            }
            break
        }

        'Circular' {
            $mode = @([Win32.Logman+Mode]::EVENT_TRACE_USE_GLOBAL_SEQUENCE, [Win32.Logman+Mode]::EVENT_TRACE_FILE_MODE_CIRCULAR) -join ','

            if (-not $PSBoundParameters.ContainsKey('MaxFileSizeMB')) {
                $MaxFileSizeMB = 2048
            }
            break
        }
    }

    $traceFile = Join-Path $Path -ChildPath $FileName

    if ($PSCmdlet.ShouldProcess($env:COMPUTERNAME, $logmanCommand)) {
        Write-Log "Starting an Outlook trace. SessionName:`"$SessionName`", traceFile:`"$traceFile`", logFileMode:`"$mode`", maxFileSize:`"$MaxFileSizeMB`""

        $err = $($stdout = Invoke-Command {
                $ErrorActionPreference = 'Continue'
                & logman.exe start trace $SessionName -pf $providerFile -o $traceFile -bs 128 -max $MaxFileSizeMB -mode $mode -ets
            }) 2>&1

        if ($err -or $LASTEXITCODE -ne 0) {
            Write-Error "logman failed. exit code:$LASTEXITCODE; stdout:`"$stdout`"; error:`"$err`""
            return
        }
    }
}

function Stop-OutlookTrace {
    [CmdletBinding()]
    param(
        $SessionName = 'OutlookTrace'
    )

    Write-Log "Stopping $SessionName"
    $null = Stop-EtwSession $SessionName
}

function Start-NetshTrace {
    [CmdletBinding()]
    param (
        [parameter(Mandatory = $true)]
        $Path,
        $FileName = "Netsh_$(Get-DateTimeString).etl",
        [ValidateSet('None', 'Mini', 'Full')]
        $ReportMode = 'None'
    )

    if (-not (Test-Path $Path)) {
        $null = New-Item $Path -ItemType Directory -ErrorAction Stop
    }

    $Path = Convert-Path -LiteralPath $Path

    # Use "InternetClient_dbg" for Win10
    $win32OS = Get-CimInstance Win32_OperatingSystem
    $osMajor = $win32OS.Version.Split(".")[0] -as [int]
    $win32OS.Dispose()

    if ($osMajor -ge 10) {
        $scenario = "InternetClient_dbg"
    }
    else {
        $scenario = "InternetClient"
    }

    if ($env:PROCESSOR_ARCHITEW6432) {
        $netshexe = Join-Path $env:SystemRoot 'SysNative\netsh.exe'
    }
    else {
        $netshexe = Join-Path $env:SystemRoot 'System32\netsh.exe'
    }

    if (-not (Get-Command $netshexe -ErrorAction SilentlyContinue)) {
        Write-Error "Cannot find $netshexe"
        return
    }

    Write-Log "Clearing dns cache"
    $null = & ipconfig.exe /flushdns

    Write-Log "Starting a netsh trace"
    $traceFile = Join-Path $Path -ChildPath $FileName
    $err = $($stdout = Invoke-Command {
            $ErrorActionPreference = 'Continue'
            & $netshexe trace start scenario=$scenario capture=yes tracefile="`"$traceFile`"" overwrite=yes maxSize=2048 # correlation=yes
        }) 2>&1

    if ($err -or $LASTEXITCODE -ne 0) {
        Write-Error "netsh failed.`nexit code:$LASTEXITCODE; stdout:$stdout; error:$err"
        return
    }

    # Even with "report=no" (by default), "HKEY_CURRENT_USER\System\CurrentControlSet\Control\NetTrace\Session\MiniReportEnabled" might be set to 1.
    # (This depends on Win10 version with a scenario. For InternetClient_dbg scenario, Win10 2004 and above does not generate mini report).
    # In order to suppress generating a minireport (i.e. C:\Windows\System32\gatherNetworkInfo.vbs), set MiniReportEnabled to 0 before netsh trace stop.
    # * You could set "report=disabled", but if you want the mini report specifically (not Full report), you need to manually configure the registry value.
    $netshRegPath = 'HKCU:\System\CurrentControlSet\Control\NetTrace\Session\'
    switch ($ReportMode) {
        'None' { Set-ItemProperty -Path $netshRegPath -Name 'MiniReportEnabled' -Type DWord -Value 0; break }
        'Mini' { Set-ItemProperty -Path $netshRegPath -Name 'MiniReportEnabled' -Type DWord -Value 1; break }
        'Full' { Set-ItemProperty -Path $netshRegPath -Name 'ReportEnabled' -Type DWord -Value 1; break }
    }

    Write-Log "ReportMode $ReportMode is configured"
}

function Stop-NetshTrace {
    [CmdletBinding()]
    param (
        $SessionName = "NetTrace"
    )

    # Netsh session might not be found right after it started. So repeat with some delay (currently 1 + 2 + 3 = 6 seconds max).
    $maxRetry = 3
    $retry = 0
    $sessionFound = $false

    while ($retry -le $maxRetry -and -not $sessionFound) {
        if ($retry) {
            Write-Log "$SessionName was not found. Retrying after $retry seconds"
            Start-Sleep -Seconds $retry
        }

        $sessions = @(Get-EtwSession | Where-Object { $_.SessionName -like "*$SessionName*" })
        if ($sessions.Count -eq 1) {
            $SessionName = $sessions[0].SessionName
            $sessionFound = $true
            break
        }
        elseif ($sesionNames.Count -gt 1) {
            Write-Error "Found multiple sessions matching $SessionName"
            return
        }

        ++$retry
    }

    if (-not $sessionFound) {
        Write-Error "Cannot find a netsh trace session"
        return
    }

    # Get a netsh trace report mode
    $sessionProps = Get-ItemProperty 'HKCU:\System\CurrentControlSet\Control\NetTrace\Session\'
    $reportMode = 'None'

    if ($sessionProps.ReportEnabled) {
        $reportMode = 'Full'
    }
    elseif ($sessionProps.MiniReportEnabled) {
        $reportMode = 'Mini'
    }

    Write-Log "ReportMode $reportMode is found"

    if ($reportMode -ne 'None') {
        Write-Progress -Activity "Stopping netsh trace" -Status "This might take a while. Generating a $reportMode Report" -PercentComplete -1
    }

    if ($env:PROCESSOR_ARCHITEW6432) {
        $netshexe = Join-Path $env:SystemRoot 'SysNative\netsh.exe'
    }
    else {
        $netshexe = Join-Path $env:SystemRoot 'System32\netsh.exe'
    }

    if (-not (Get-Command $netshexe -ErrorAction SilentlyContinue)) {
        Write-Error "Cannot find $netshexe"
        return
    }

    Write-Log "Stopping $SessionName with netsh trace stop"

    $err = $($stdout = Invoke-Command {
            $ErrorActionPreference = 'Continue'
            & $netshexe trace stop
        }) 2>&1

    if ($err -or $LASTEXITCODE -ne 0) {
        Write-Log "Failed to stop netsh trace ($SessionName). exit code:$LASTEXITCODE; stdout:$stdout; error:$err" -Category Warning
        Write-Log "Stopping with Stop-EtwSession"
        $null = Stop-EtwSession -SessionName $SessionName
    }

    if ($reportMode -ne 'None') {
        Write-Progress -Activity "Stopping netsh trace" -Status "Done" -Completed
    }
}

function Get-EtwSession {
    [CmdletBinding()]
    param()

    try {
        [Win32.Advapi32]::QueryAllTraces()
    }
    catch {
        Write-Error -Message "QueryAllTraces failed. $_" -Exception $_.Exception
    }
}

function Stop-EtwSession {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$SessionName
    )

    try {
        [Win32.Advapi32]::StopTrace($SessionName)
    }
    catch {
        Write-Error -Message "StopTrace for $SessionName failed. $_" -Exception $_.Exception
    }
}

function Start-PSR {
    param(
        [parameter(Mandatory = $true)]
        $Path,
        $FileName = "PSR.mht",
        [switch]$ShowGUI
    )

    if (-not (Get-Command 'psr.exe' -ErrorAction SilentlyContinue)) {
        Write-Error "psr.exe is not available"
        return
    }

    # Make sure psr isn't running already.
    $processes = @(Get-Process psr -ErrorAction SilentlyContinue)

    if ($processes.Count -gt 0) {
        Write-Error "PSR is already running (PID:$($processes.ID -join ','))"
        return
    }

    if (-not (Test-Path $Path -ErrorAction Stop)) {
        $null = New-Item -ItemType Directory $Path -ErrorAction Stop
    }

    $Path = Convert-Path -LiteralPath $Path

    # File name must be ***.mht
    if ([IO.Path]::GetExtension($FileName) -ne ".mht") {
        $FileName = [IO.Path]::GetFileNameWithoutExtension($FileName) + '.mht'
    }

    # For Win7, maxsc is 100
    $maxScreenshotCount = 100

    $win32OS = Get-CimInstance Win32_OperatingSystem
    $osMajor = $win32OS.Version.Split(".")[0] -as [int]
    $osMinor = $win32OS.Version.Split(".")[1] -as [int]
    $win32OS.Dispose()

    if ($osMajor -gt 6 -or ($osMajor -eq 6 -and $osMinor -ge 3)) {
        $maxScreenshotCount = 300
    }

    $outputFile = Join-Path $Path -ChildPath $FileName

    if ($outputFile.IndexOf(' ') -ge 0) {
        $outputFile = "`"$outputFile`""
    }

    $psrArgs = @(
        '/start', '/maxsc', $maxScreenshotCount, '/maxlogsize', '10', '/output', $outputFile, '/exitonsave', '1', '/noarc', '1'

        if (-not $ShowGUI) {
            '/gui 0'
        }
    )

    $err = $($process = Start-Process 'psr' -ArgumentList $psrArgs -PassThru) 2>&1 | Select-Object -First 1

    if (-not $process -or $process.HasExited) {
        Write-Error -Message "PSR failed to start. $err" -Exception $err.Exception
        return
    }

    # Why access Handle? To make ExitTime, ExitCode etc available. See blow:
    # https://stackoverflow.com/questions/10262231/obtaining-exitcode-using-start-process-and-waitforexit-instead-of-wait/23797762#23797762
    $null = $process.Handle

    Write-Log "PSR (PID:$($process.Id)) started $(if ($ShowGUI) {'with UI'} else {'without UI'}). maxScreenshotCount:$maxScreenshotCount"

    [PSCustomObject]@{
        Process = $process
    }
}

function Stop-PSR {
    [CmdletBinding()]
    param (
        # Object returned from Start-PSR
        $StartResult
    )

    if ($StartResult) {
        $currentInstance = $StartResult.Process

        try {
            # WaitForExit(0) is used here instead of HasExited in order to detect the following 2 conditions with one shot:
            # 1. The actual process has exited already, but System.Diagnostics.Process has not been disposed yet.
            # 2. System.Diagnostics.Process has been disposed (i.e. invalid input). WaitForExit() throws an exception "No process is associated with this object".
            if ($currentInstance.WaitForExit(0)) {
                Write-Error "psr.exe (PID:$($currentInstance.Id)) has already exited. ExitTime:$($currentInstance.ExitTime), ExitCode:$($currentInstance.ExitCode)"
                $currentInstance.Dispose()
                return
            }
        }
        catch {
            Write-Error -Message "Process object passed by StartResult parameter has been disposed already. $_" -Exception $_.Exception
            return
        }
    }
    else {
        $processes = @(Get-Process -Name psr -ErrorAction SilentlyContinue)

        if ($processes.Count -eq 0) {
            Write-Error 'Cannot find psr.exe process'
            return
        }
        elseif ($processes.Count -eq 1) {
            $currentInstance = $processes[0]
            $null = $currentInstance.Handle
        }
        elseif ($processes.Count -gt 1) {
            # Unexpected to find multiple psr.exe instances.
            Write-Log "There are $($processes.Count) instances of psr.exe (PID:$($processes.ID -join ','))"
            $processes | ForEach-Object { if ($_.Dispose) { $_.Dispose() } }
            return
        }
    }

    # "psr /stop" creates a new instance of psr.exe and it stops the instance currently running.
    $err = $($stopInstance = Start-Process 'psr' -ArgumentList '/stop' -PassThru) 2>&1 | Select-Object -First 1

    if (-not $stopInstance) {
        Write-Error -Message "Failed to run psr.exe /stop. $err" -Exception $err.Exception
        return
    }

    # Do not use Wait-Process here because it can fail with Access Denied when running as non-admin
    $currentInstance.WaitForExit()
    Write-Log "PSR (PID:$($currentInstance.Id)) stopped. ExitCode:$($currentInstance.ExitCode)"

    if ($currentInstance.Dispose) {
        $currentInstance.Dispose()
    }

    try {
        if ($stopInstance.WaitForExit(1000)) {
            return
        }

        # When there were no clicks, the instance of 'psr /stop' remains after the existing instance exits. This causes a hung.
        # The existing instance is supposed to signal an event which 'psr /stop' instance waits for. But it seems this does not happen when there were no clicks.
        # So to avoid this, the following code manually signals the event so that 'psr /stop' shuts down.
        if (-not $stopInstance.HasExited) {
            $PSR_CLEANUP_COMPLETED = '{CD3E5009-5C9D-4E9B-B5B6-CAE1D8799AE3}'
            $h = [System.Threading.EventWaitHandle]::OpenExisting($PSR_CLEANUP_COMPLETED)
            $null = $h.Set()
            Write-Log "PSR_CLEANUP_COMPLETED was manually signaled"
            $stopInstance.WaitForExit()
        }
    }
    catch {
        Write-Log -ErrorRecord $_
    }
    finally {
        if ($stopInstance.Dispose) {
            $stopInstance.Dispose()
        }
    }
}

function Save-EventLog {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        $Path
    )

    # Need admin rights to archive event logs.
    if (-not (Test-RunAsAdministrator)) {
        Write-Error "Please run as administrator"
        return
    }

    # If this command is run by itself (not from Collect-OutlookInfo), need to create a runspace pool.
    if (-not $Script:RunspacePool) {
        Open-TaskRunspace
        $runspaceOpened = $true
    }

    if (-not (Test-Path $Path -ErrorAction Stop)) {
        $null = New-Item -ItemType directory $Path
    }
    $Path = Convert-Path -LiteralPath $Path

    $logs = @(
        'Application'
        'System'
        (wevtutil el) -match "Microsoft-Windows-Windows Firewall With Advanced Security|AAD|Microsoft-Windows-Bits-Client|WebAuth|CAPI2|AppLocker|AppxPackaging|AppXDeployment/|AppXDeploymentServer/"
    )

    $tasks = @(
        foreach ($log in $logs) {
            $fileName = $log.Replace('/', '_') + '.evtx'
            $filePath = Join-Path $Path -ChildPath $fileName
            Write-Log "Saving $log to $filePath"

            Start-Task -Name 'EventLogExportTask' -ScriptBlock {
                param ($Log, $FilePath)
                wevtutil export-log $Log $FilePath /ow
                wevtutil archive-log $FilePath
            } -ArgumentList @{Log = $log; FilePath = $filePath }
        }
    )

    $tasks | Receive-Task -AutoRemoveTask

    if ($Local:runspaceOpened) {
        Close-TaskRunspace
    }
}

function Get-InstalledUpdate {
    [CmdletBinding()]
    param()

    # Ask items in AppUpdatesFolder from Shell
    # FOLDERID_AppUpdates == a305ce99-f527-492b-8b1a-7e76fa98d6e4
    $shell = $appUpdates = $null

    try {
        $shell = New-Object -ComObject Shell.Application
        $appUpdates = $shell.NameSpace('Shell:AppUpdatesFolder')
        if ($null -eq $appUpdates) {
            Write-Log "Cannot obtain Shell:AppUpdatesFolder. Probabliy 32bit PowerShell is used on 64bit OS"
            Write-Error "Cannot obtain Shell:AppUpdatesFolder"
            return
        }

        $items = $appUpdates.Items()

        foreach ($item in $items) {
            # Raw installedOn includes 0x0e20 (0x200E Left-to-Right char). Remove them.
            $installedOnRaw = $appUpdates.GetDetailsOf($item, 12)
            $installedOn = $null

            if ($installedOnRaw) {
                try {
                    $installedOn = New-Object string -ArgumentList (, $($installedOnRaw.ToCharArray() | Where-Object { $_ -lt 128 }))
                }
                catch {
                    Write-Error -ErrorRecord $_
                }
            }

            # https://docs.microsoft.com/en-us/windows/win32/shell/folder-getdetailsof
            [PSCustomObject]@{
                Name        = $item.Name
                Program     = $appUpdates.GetDetailsOf($item, 2)
                Version     = $appUpdates.GetDetailsOf($item, 3)
                Publisher   = $appUpdates.GetDetailsOf($item, 4)
                URL         = $appUpdates.GetDetailsOf($item, 7)
                InstalledOn = $installedOn
            }

            $null = [System.Runtime.Interopservices.Marshal]::FinalReleaseComObject($item)
        }
    }
    finally {
        if ($appUpdates) {
            $null = [System.Runtime.Interopservices.Marshal]::FinalReleaseComObject($appUpdates)
        }

        if ($shell) {
            $null = [System.Runtime.Interopservices.Marshal]::FinalReleaseComObject($shell)
        }
    }
}

function Resolve-User {
    [CmdletBinding(PositionalBinding = $false)]
    param(
        [Parameter(Position = 0, Mandatory = $true)]
        # User Name or SID
        [string]$Identity
    )

    if ($null -eq $Script:ResolveCache) {
        $Script:ResolveCache = @{}
    }

    # Return the cached entry if available.
    if ($Script:ResolveCache.ContainsKey($Identity)) {
        $Script:ResolveCache[$Identity]
        return
    }

    # Note:WMI Win32_UserAccount can be very slow. I'm avoiding here.
    # Get-WmiObject -Class Win32_UserAccount -Filter "Name = '$userName'"

    $sid = $account = $null

    # Is SID?
    try {
        $sid = New-Object System.Security.Principal.SecurityIdentifier($Identity)
        $account = $sid.Translate([System.Security.Principal.NTAccount])
    }
    catch {
        # Ignore
    }

    # If not SID, then must be the account name
    if (-not $sid) {
        try {
            $account = New-Object System.Security.Principal.NTAccount($Identity)
            $sid = $account.Translate([System.Security.Principal.SecurityIdentifier])

            # Translate from SID to acccount so that the account name is more complete (domain\name)
            $account = $sid.Translate([System.Security.Principal.NTAccount])
        }
        catch {
            # Ignore
        }
    }

    if ($null -eq $sid -or $null -eq $account) {
        Write-Error "Cannot resolve $Identity"
        return
    }

    $resolved = [PSCustomObject]@{
        Name = $account.ToString()
        Sid  = $sid.ToString()
    } | Add-Member -MemberType ScriptMethod -Name 'ToString' -Value { $this.Name } -Force -PassThru

    # Add to cache
    if (-not $Script:ResolveCache.ContainsKey($resolved.Name)) {
        $Script:ResolveCache.Add($resolved.Name, $resolved)
    }

    if (-not $Script:ResolveCache.ContainsKey($resolved.Sid)) {
        $Script:ResolveCache.Add($resolved.Sid, $resolved)
    }

    $resolved
}

function Get-LogonUser {
    [CmdletBinding()]
    param()

    # Find unique users of explorer.exe instances.
    Get-CimInstance Win32_Process -Filter 'Name = "explorer.exe"' | & {
        param([Parameter(ValueFromPipeline)]$win32Process)

        begin { $usersCache = @{} }

        process {
            try {
                $owner = $win32Process | Get-ProcessOwner

                if (-not $owner) {
                    Write-Verbose "Cannot obtain the owner of explorer (PID $($win32Process.ProcessId)). Probably you are runnning without admin privilege"
                    return
                }

                if ($usersCache.ContainsKey($owner.Sid)) {
                    return
                }

                $usersCache.Add($owner.Sid, $null)
                $owner
            }
            finally {
                $win32Process.Dispose()
            }
        }
    }
}

function Get-ProcessOwner {
    [CmdletBinding()]
    param(
        [Parameter(ParameterSetName = 'Id', Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName)]
        [Alias('ProcessId')]
        [int]$Id,
        [Parameter(ParameterSetName = 'Win32Process', Mandatory, ValueFromPipeline)]
        # CimInstance of Win32_Process (Technically this parameter is not necessary because "ProcessId" alias can match, but it helps to speed up by skipping Get-CimInstance, which is quite slow)
        [Microsoft.Management.Infrastructure.CimInstance]$Win32Process
    )

    process {
        if (-not $Win32Process) {
            $Win32Process = Get-CimInstance Win32_Process -Filter "ProcessId = $Id"
            $needDispose = $true
        }

        # Note: If the process is null or has exited, GetOwnerSid emits non-terminating error. Ignore this error.
        $null = $($owner = Invoke-CimMethod -InputObject $Win32Process -MethodName 'GetOwnerSid') 2>&1

        if ($owner.ReturnValue -eq 0) {
            Resolve-User $owner.Sid
        }

        if ($Win32Process -and $needDispose) {
            $Win32Process.Dispose()
        }
    }
}

<#
.SYNOPSIS
Get a given local user's registry root. If User is empty, it just returns HKCU.
#>
function Get-UserRegistryRoot {
    [CmdletBinding()]
    param(
        # User name or SID
        [string]$User
    )

    if ($User) {
        $resolvedUser = Resolve-User $User

        if (-not $resolvedUser) {
            return
        }

        $userRegRoot = "Registry::HKEY_USERS\$($resolvedUser.Sid)"

        if (-not (Test-Path $userRegRoot)) {
            Write-Error "Cannot find $userRegRoot"
            return
        }
    }
    else {
        Write-Log "User is empty. Use Registry::HKCU"
        $userRegRoot = 'Registry::HKCU'
    }

    $userRegRoot
}

<#
.SYNOPSIS
Get a given user's profile path (i.e. same as USERPROFILE environment variable)
#>
function Get-UserProfilePath {
    [CmdletBinding()]
    param(
        # User name or SID
        [string]$User
    )

    if (-not $User) {
        $User = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
    }

    $resolvedUser = Resolve-User $User

    if (-not $resolvedUser) {
        return
    }

    # Get the value of ProfileImagePath
    $userProfile = Get-ItemProperty "Registry::HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList\$($resolvedUser.Sid)\"
    $userProfile.ProfileImagePath
}

<#
.SYNOPSIS
Get a given user's shell folder path (e.g. "LocalAppData", "Desktop" etc.)
#>
function Get-UserShellFolder {
    [CmdletBinding()]
    param(
        # User name or SID
        [string]$User,
        [parameter(Mandatory = $true)]
        # Shell folder name (e.g. "AppData", "Desktop", "Local AppData" etc.)
        [ValidateSet('AppData', 'Desktop', 'Local AppData', 'Programs', 'Personal', 'Startup')]
        [string]$ShellFolderName
    )

    $userRegRoot = Get-UserRegistryRoot -User $User

    if (-not $userRegRoot) {
        return
    }

    # Do not use Get-ItemProperty here because it'd expand environment variable.
    $shellFolders = Get-Item $(Join-Path $userRegRoot "SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders")
    $folderPath = $shellFolders.GetValue($ShellFolderName, $null, [Microsoft.Win32.RegistryValueOptions]::DoNotExpandEnvironmentNames)
    $shellFolders.Dispose()

    if (-not $folderPath) {
        return
    }

    # Folder path is like "%USERPROFILE%\AppData\Local". Replace USERPROFILE.
    $userProfile = Get-UserProfilePath $User
    $folderPath.Replace('%USERPROFILE%', $userProfile)
}

<#
.SYNOPSIS
Get a given user's TEMP folder path
#>
function Get-UserTempFolder {
    [CmdletBinding()]
    param(
        $User
    )

    $userRegRoot = Get-UserRegistryRoot -User $User

    if (-not $userRegRoot) {
        return
    }

    $tempPath = Join-Path $userRegRoot 'Environment' | Get-ItemProperty -Name 'TEMP' | Select-Object -ExpandProperty 'TEMP'
    $tempPath
}

function Save-OfficeRegistry {
    [CmdletBinding()]
    param (
        [parameter(Mandatory = $true)]
        $Path,
        [string]$User
    )

    if (-not (Test-Path $Path)) {
        $null = New-Item $Path -ItemType directory -ErrorAction Stop
    }

    $registryKeys = @(
        # HKCU
        'HKCU\Software\Microsoft\Exchange'
        'HKCU\Software\Microsoft\Office'
        'HKCU\Software\Wow6432Node\Microsoft\Office'

        'HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings'
        'HKCU\Software\Classes\Local Settings\Software\Microsoft\MSIPC'
        'HKCU\Software\IM Providers'
        'HKCU\Software\Microsoft\Windows\CurrentVersion\Notifications'
        'HKCU\Software\Microsoft\AuthN' # for Alternate Login ID. https://docs.microsoft.com/en-us/windows-server/identity/ad-fs/operations/configuring-alternate-login-id
        'HKCU\Software\Microsoft\IdentityCRL' # WAM related

        # HKLM
        'HKLM\SOFTWARE\Microsoft\Office'
        'HKLM\SOFTWARE\WOW6432Node\Microsoft\Office'
        'HKLM\SOFTWARE\IM Providers'
        'HKLM\SOFTWARE\Wow6432Node\IM Providers'
        'HKLM\SOFTWARE\Microsoft\MSIPC'
        'HKLM\SOFTWARE\WOW6432Node\Microsoft\MSIPC'
        'HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Internet Settings\WinHttp'
        'HKLM\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Internet Settings\WinHttp'
        'HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies'
        'HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\AeDebug'
        'HKLM\SOFTWARE\WOW6432Node\Microsoft\Windows NT\CurrentVersion\AeDebug'
        'HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\AppModelUnlock'
        'HKLM\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols'
        'HKLM\SOFTWARE\Microsoft\WindowsUpdate\Orchestrator'
        'HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Orchestrator'

        # Policies
        'HKCU\Software\Policies'
        'HKCU\Software\Wow6432Node\Policies'
        'HKLM\SOFTWARE\Policies'
        'HKLM\SOFTWARE\WOW6432Node\Policies'
    )

    $userRegRoot = Get-UserRegistryRoot $User | Convert-Path -ErrorAction SilentlyContinue

    if ($userRegRoot) {
        $registryKeys = $registryKeys | & { process { $_.Replace("HKCU", $userRegRoot).TrimEnd('\') } }
    }

    # Make sure NOT to use WOW64 version of reg.exe when running on 32bit PowerShell on 64bit OS.
    # I could use "/reg:64" option of reg.exe, but it's not available for Win7.
    if ($env:PROCESSOR_ARCHITEW6432) {
        $regexe = Join-Path $env:SystemRoot 'SysNative\reg.exe'
    }
    else {
        $regexe = Join-Path $env:SystemRoot 'System32\reg.exe'
    }

    # If, for some reason, reg.exe is not available, bail.
    if (-not (Get-Command $regexe -ErrorAction SilentlyContinue)) {
        Write-Error "$regexe is not avaialble"
        return
    }

    foreach ($key in $registryKeys) {
        $err = $($queryResult = Invoke-Command {
                $ErrorActionPreference = 'Continue'
                & $regexe Query $key
            }) 2>&1

        if ($null -eq $queryResult) {
            Write-Log "$key does not exist"
            continue;
        }

        # Cannot use Test-Path because when running 32bit PS on 64bit OS, HKLM\Software is redirected to WOW6432Node
        # if (-not (Test-Path $key)) {
        #     Write-Log "$key does not exist"
        #     continue
        # }

        $filePath = Join-Path $Path -ChildPath "$($key.Replace('\','_')).reg"

        if (Test-Path $filePath) {
            Remove-Item $filePath -Force
        }

        Write-Log "Saving $key to $filePath"
        $err = $(Invoke-Command {
                $ErrorActionPreference = 'Continue'
                $null = & $regexe export $key $filePath
            }) 2>&1

        if ($LASTEXITCODE -ne 0) {
            Write-Error "$key is not exported. exit code = $LASTEXITCODE. $err"
        }
    }
}

function Save-OSConfiguration {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        $Path,
        [Parameter(Mandatory = $true)]
        $User,
        [System.Threading.CancellationToken]$CancellationToken
    )

    if (-not (Test-Path $Path)) {
        $null = New-Item $Path -ItemType directory -ErrorAction Stop
    }

    & {
        @{ScriptBlock = { Get-CimInstance -ClassName Win32_ComputerSystem }; FileName = 'Win32_ComputerSystem.xml' }
        @{ScriptBlock = { Get-CimInstance -ClassName Win32_OperatingSystem }; FileName = 'Win32_OperatingSystem.xml' }
        @{ScriptBlock = { Get-CimInstance -ClassName Win32_Processor }; FileName = 'Win32_Processor.xml' }
        @{ScriptBlock = { Get-CimInstance -Namespace root\SecurityCenter2 -ClassName AntiVirusProduct }; FileName = 'AntiVirusProduct.xml' }
        @{ScriptBlock = { Get-ComputerInfo } }
        @{ScriptBlock = { Get-WinHttpDefaultProxy } }
        @{ScriptBlock = { Get-NLMConnectivity } }
        @{ScriptBlock = { Get-MeteredNetworkCost } }
        @{ScriptBlock = { Get-WSCAntivirus } }
        @{ScriptBlock = { Get-InstalledUpdate } }
        @{ScriptBlock = { Get-JoinInformation } }
        @{ScriptBlock = { Get-DeviceJoinStatus }; FileName = 'DeviceJoinStatus.txt' }
        @{ScriptBlock = { Get-ImageFileExecutionOptions } }
        @{ScriptBlock = { Get-SessionManager } }
        @{ScriptBlock = { Get-WinSystemLocale } }
        @{ScriptBlock = { Get-Service } }
        @{ScriptBlock = { Get-SmbMapping } }
        @{ScriptBlock = { Get-AnsiCodePage } }
        @{ScriptBlock = { Get-Volume } }
        @{ScriptBlock = { Get-NetFrameworkVersion } }
        @{ScriptBlock = { cmdkey /list }; FileName = 'cmdkey.txt' }

        $userArg = @{ User = $User }
        @{ScriptBlock = { param($User) Get-WebView2 @PSBoundParameters }; ArgumentList = $userArg }
        @{ScriptBlock = { param($User) Get-AppContainerRegistryAcl @PSBoundParameters }; ArgumentList = $userArg }
        @{ScriptBlock = { param($User) Get-StructuredQuerySchema @PSBoundParameters }; ArgumentList = $userArg }

        # These are moved to Collect-OutlookInfo so that they run before Fiddler is started
        # @{ScriptBlock = { param($User) Get-WinInetProxy @PSBoundParameters }; ArgumentList = $userArg }
        # @{ScriptBlock = { param($User) Get-ProxyAutoConfig @PSBoundParameters }; ArgumentList = $userArg }

        @{ScriptBlock = { Get-AppxPackage -AllUsers } }
        @{ScriptBlock = { Get-AppxProvisionedPackage -Online } }

        # These are just for troubleshooting.
        @{ScriptBlock = { Get-ChildItem 'Registry::HKEY_USERS' | Select-Object 'Name' }; FileName = 'Users.xml' }
        @{ScriptBlock = { whoami.exe /USER }; FileName = 'whoami.txt' }
        @{ScriptBlock = { Get-Privilege } }
    } | & {
        process {
            if ($CancellationToken.IsCancellationRequested) {
                return
            }

            Invoke-ScriptBlock @_ -Path $Path
        }
    }
}

function Save-NetworkInfo {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        $Path,
        [Threading.CancellationToken]$CancellationToken
    )

    if (-not (Test-Path $Path)) {
        $null = New-Item $Path -ItemType directory -ErrorAction Stop
    }

    # These are from C:\Windows\System32\gatherNetworkInfo.vbs with some extra.
    & {
        @{ScriptBlock = { Get-NetAdapter -IncludeHidden } }
        @{ScriptBlock = { Get-NetAdapterAdvancedProperty } }
        @{ScriptBlock = { Get-NetAdapterBinding -IncludeHidden } }
        @{ScriptBlock = { Get-NetIpConfiguration -Detailed } }
        @{ScriptBlock = { Get-DnsClientNrptPolicy } }
        # @{ScriptBlock = {Resolve-DnsName 'bing.com'}}
        # @{ScriptBlock = {ping bing.com -4}}
        # @{ScriptBlock = {ping bing.com -6}}
        # @{ScriptBlock = {Test-NetConnection 'bing.com' -InformationLevel Detailed}}
        # @{ScriptBlock = {Test-NetConnection 'bing.com' -InformationLevel Detailed -CommonTCPPort HTTP}}
        @{ScriptBlock = { Get-NetRoute } }
        @{ScriptBlock = { Get-NetIPAddress } }
        # @{ScriptBlock = {Get-NetLbfoTeam}}
        # @{ScriptBlock = {Get-Service -Name:VMMS}}
        # @{ScriptBlock = {Get-VMSwitch}}
        # @{ScriptBlock = {Get-VMNetworkAdapter -all}}
        # @{ScriptBlock = {Get-WindowsOptionalFeature -Online}}
        # @{ScriptBlock = {Get-Service}}
        # @{ScriptBlock = {Get-PnpDevice | Get-PnpDeviceProperty -KeyName DEVPKEY_Device_InstanceId,DEVPKEY_Device_DevNodeStatus,DEVPKEY_Device_ProblemCode}}
        @{ScriptBlock = { Get-NetIPInterface } }
        @{ScriptBlock = { Get-NetConnectionProfile } }
        @{ScriptBlock = { Get-NetFirewallProfile } }
        @{ScriptBlock = { Show-NetFirewallRule } }
        @{ScriptBlock = { Get-Content $(Join-Path $env:SystemRoot 'System32\drivers\etc\hosts') }; FileName = 'hosts.txt' }
        @{ScriptBlock = { ipconfig /all } }

        # Dump Windows Firewall config
        @{ScriptBlock = { netsh advfirewall monitor show currentprofile } }
        @{ScriptBlock = { netsh advfirewall monitor show firewall } }
        @{ScriptBlock = { netsh advfirewall monitor show consec } }
        @{ScriptBlock = { netsh advfirewall firewall show rule name=all verbose } }
        @{ScriptBlock = { netsh advfirewall consec show rule name=all verbose } }
        @{ScriptBlock = { netsh advfirewall monitor show firewall rule name=all } }
        @{ScriptBlock = { netsh advfirewall monitor show consec rule name=all } }
    } | & {
        process {
            if ($CancellationToken.IsCancellationRequested) {
                return
            }

            Invoke-ScriptBlock @_ -Path $Path
        }
    }
}

<#
.DESCRIPTION
Run a given script block. If Path is given, save the result there. Any errors (terminating or non-terminating) will be written by Write-Log.
If FileName is given, it's used for the file name for saving the result. If its extension is not ".xml", Set-Content will be used. Otherwise Export-CliXml will be used.
If FileName is not give, the file name will be auto-decided. If the command is an application, then Set-Content will be used. Otherwise Export-CliXml will be used.
#>
function Invoke-ScriptBlock {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true, Position = 0)]
        [ScriptBlock]$ScriptBlock,
        $ArgumentList,
        # Destination folder path
        $Path,
        # File name used for saving
        [string]$FileName
    )

    $result = $null
    $start = Get-Timestamp

    $scriptBlockName = $ScriptBlock.Ast.Name

    if (-not $scriptBlockName) {
        if ($ScriptBlock.Ast) {
            $scriptBlockName = $ScriptBlock.Ast.ToString()
        }
        else {
            $scriptBlockName = "{$ScriptBlock}"
        }
    }

    # Wrap in an array in case a single object is passed as ArgumentList (otherwise, splatting does not work as expected)
    if ($null -ne $ArgumentList -and $ArgumentList -isnot [System.Collections.ICollection]) {
        $ArgumentList = @($ArgumentList)
    }

    # Suppress progress that may be written by the script block
    $savedProgressPreference = $ProgressPreference
    $ProgressPreference = "SilentlyContinue";

    try {
        # To redirect error, call operator (&) is used, instead of $ScriptBlock.InvokeReturnAsIs().
        if ($ArgumentList) {
            $err = $($result = & $ScriptBlock @ArgumentList) 2>&1
        }
        else {
            $err = $($result = & $ScriptBlock) 2>&1
        }

        foreach ($e in $err) {
            Write-Log "$scriptBlockName $(if ($ArgumentList) { "with $(ConvertFrom-ArgumentList $ArgumentList) " })had a non-terminating error. $e" -ErrorRecord $e -Category Warning
        }
    }
    catch {
        Write-Log "$scriptBlockName $(if ($ArgumentList) { "with $(ConvertFrom-ArgumentList $ArgumentList) " })threw a terminating error. $_" -ErrorRecord $_ -Category Error
    }
    finally {
        $ProgressPreference = $savedProgressPreference
    }

    $elapsed = Get-Elapsed $start
    Write-Log "$scriptBlockName took $($elapsed.TotalMilliseconds) ms.$(if ($null -eq $result) {" It returned nothing"})"

    if ($null -eq $result) {
        return
    }

    if (-not $Path) {
        $result
        return
    }

    # If Path is given, save the result.
    if (-not $FileName) {
        # Decide the file name.
        # 1. If the ScriptBlock has FunctionDefinitionAst, use its Name (that is, the function name)
        # 2. Otherwise, search the first statement from ProcessBlock -> EndBlock -> BeginBlock
        # Note that a simple script block such as '{ Get-Foo }' has only EndBlock.
        # When the command type is Application (e.g, netsh), use the entire statement as the file name (so that command with different args are saved as indivisual files)
        # This is not perfect but if the scriptblock is more complicated, caller should supply FileName parameter.
        $commandName = $ScriptBlock.Ast.Name

        if (-not $commandName) {
            # Get the first statement
            $statement = & {
                $ScriptBlock.Ast.ProcessBlock
                $ScriptBlock.Ast.EndBlock
                $ScriptBlock.Ast.BeginBlock
            } | & {
                process {
                    if ($_) {
                        $_.Statements[0].Extent.Text
                    }
                }
            } | Select-Object -First 1
        }

        $commandName = ([RegEx]::Match($statement, '[\w-\.]+')).Value.Trim()

        if ($commandName) {
            if ($command = Get-Command $commandName -ErrorAction SilentlyContinue) {
                if ($command.CommandType -eq 'Application') {
                    $FileName = $statement.Replace('/', '-') + ".txt"
                }
                else {
                    $FileName = $command.Noun + '.xml'
                }
            }
        }

        if (-not $FileName) {
            $FileName = [Guid]::NewGuid().ToString() + ".xml"
            Write-Log "Cannot determine command name from $scriptBlockName. Saving with a random name $FileName" -Category Error
        }
    }

    if (-not (Test-Path $Path)) {
        $null = New-Item $Path -ItemType Directory -ErrorAction SilentlyContinue
    }

    if ([IO.Path]::GetExtension($FileName) -eq '.xml') {
        $result | Export-Clixml -LiteralPath (Join-Path $Path $FileName)
    }
    else {
        $result | Set-Content -LiteralPath (Join-Path $Path $FileName)
    }

    # Dispose if necessary
    foreach ($_ in $result) {
        if ($_.Dispose) {
            $_.Dispose()
        }
    }
}

<#
.SYNOPSIS
    Get WinInet proxy settings for a user.

.DESCRIPTION
    Get WinInet proxy settings for a user. If User is not give, the current user is used.

.EXAMPLE
    Get-WinInetProxy -User user01

    ProxySettingsPerUser :
    User                 : Admin
    Connection           : DefaultConnectionSettings
    AutoDetect           : True
    AutoConfigUrl        :
    Proxy                : myproxy2:8081
    ProxyBypass          : <local>
#>
function Get-WinInetProxy {
    [CmdletBinding(PositionalBinding = $false)]
    param(
        [Parameter(Position = 0)]
        [string]$User
    )

    if (-not $User) {
        $User = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
    }

    # For now, I want to include the result of WinHttpGetIEProxyConfigForCurrentUser because it automatically gets the WinInet proxy setting of "acitve" connection.
    # I do not know how to determine which connection is active yet.
    $props = [ordered]@{}
    $winInetProxy = New-Object Win32.WinHttp+WINHTTP_CURRENT_USER_IE_PROXY_CONFIG

    if ([Win32.WinHttp]::WinHttpGetIEProxyConfigForCurrentUser([ref] $winInetProxy)) {
        $props['fAutoDetect'] = $winInetProxy.fAutoDetect
        # Wrap the native string data in SafeHandle-derived class so that the memory will be properly freed (By GlobalFree in this case) when GC collects them.
        $props['lpszAutoConfigUrl'] = (New-Object Win32.SafeGlobalFreeString -ArgumentList $winInetProxy.lpszAutoConfigUrl).ToString()
        $props['lpszProxy'] = (New-Object Win32.SafeGlobalFreeString -ArgumentList $winInetProxy.lpszProxy).ToString()
        $props['lpszProxyBypass'] = (New-Object Win32.SafeGlobalFreeString -ArgumentList $winInetProxy.lpszProxyBypass).ToString()
        $props['User'] = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
        $currentUserActiveConnProxy = [PSCustomObject]$props
    }
    else {
        Write-Error -Message ("Win32 WinHttpGetIEProxyConfigForCurrentUser failed with 0x{0:x8}" -f [System.Runtime.InteropServices.Marshal]::GetLastWin32Error())
    }

    # If ProxySettingsPerUser is 0, then check HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Internet Settings\Connections, instead of the user's registry.
    $proxySettingsPerUser = Get-ItemProperty 'Registry::HKLM\SOFTWARE\Policies\Microsoft\Windows\CurrentVersion\Internet Settings' -Name 'ProxySettingsPerUser' -ErrorAction SilentlyContinue `
    | Select-Object -ExpandProperty 'ProxySettingsPerUser'

    if ($proxySettingsPerUser -eq 0) {
        $regRoot = 'Registry::HKLM'
    }
    else {
        $err = $($regRoot = Get-UserRegistryRoot -User $User) 2>&1 | Select-Object -First 1

        if (-not $regRoot) {
            Write-Error -Message "Cannot get user $User's registry root. $err" -Exception $err.Exception
            return
        }
    }

    # There might be multiple connections besides "DefaultConnectionSettings" if there are VPNs.
    $connectionsKey = Join-Path $regRoot 'SOFTWARE\Microsoft\Windows\CurrentVersion\Internet Settings\Connections'
    $connections = @(Get-Item $connectionsKey | Select-Object -ExpandProperty Property)

    # It's possible that there is no connection at all (maybe IE has never been started).
    # In this case, return the default configuration (This is what WinHttpGetIEProxyConfigForCurrentUser does anyway).
    if ($connections.Count -eq 0) {
        Write-Log "No connections are found under $connectionsKey. Returning a default setting"

        [PSCustomObject]@{
            ProxySettingsPerUser  = $proxySettingsPerUser
            User                  = $User
            Connection            = 'DefaultConnectionSettings'
            AutoDetect            = $true
            AutoConfigUrl         = $null
            Proxy                 = $null
            ProxyBypass           = $null
            ActiveConnectionProxy = $currentUserActiveConnProxy
        }

        return
    }

    foreach ($connection in $connections) {
        # Skip SavedLegacySettings & WinHttpSettings (in HKLM)
        if ($connection -eq 'SavedLegacySettings' -or $connection -eq 'WinHttpSettings') {
            continue
        }

        $raw = $null
        try {
            # This line could throw a terminating error.
            $raw = Get-ItemProperty $connectionsKey -Name $connection -ErrorAction SilentlyContinue | Select-Object -ExpandProperty $connection
        }
        catch {
            Write-Error -Message "Get-ItemProperty failed for a connection, $connection. $_" -Exception $_.Exception
        }

        if (-not $raw) {
            continue
        }

        # Parse the data
        $structversion = [BitConverter]::ToInt32($raw, 0)
        $settingsVersion = [BitConverter]::ToInt32($raw, 4)
        $flags = [BitConverter]::ToInt32($raw, 8)

        $proxySize = [BitConverter]::ToInt32($raw, 12)
        $proxy = [Text.Encoding]::ASCII.GetString($raw, 16, $proxySize)
        $position = 16 + $proxySize

        $proxyBypassSize = [BitConverter]::ToInt32($raw, $position)
        $proxyBypass = [Text.Encoding]::ASCII.GetString($raw, $position + 4, $proxyBypassSize)
        $position += 4 + $proxyBypassSize

        $autoConfigUrlSize = [BitConverter]::ToInt32($raw, $position)
        $autoConfigUrl = [Text.Encoding]::ASCII.GetString($raw, $position + 4, $autoConfigUrlSize)

        $winInetProxy = [PSCustomObject]@{
            StructVersion   = $structversion
            SettingsVersion = $settingsVersion
            Flags           = $flags -as [Win32.WinInet+PER_CONN_FLAGS]
            Proxy           = $proxy
            ProxyBypass     = $proxyBypass
            AutoConfigUrl   = $autoConfigUrl
        }

        $props = [ordered]@{}
        $props['ProxySettingsPerUser'] = $proxySettingsPerUser
        $props['User'] = $User
        $props['Connection'] = $connection

        $props['AutoDetect'] = ($winInetProxy.Flags -band [Win32.WinInet+PER_CONN_FLAGS]::PROXY_TYPE_AUTO_DETECT) -as [bool]
        $props['AutoConfigUrl'] = if ($winInetProxy.Flags -band [Win32.WinInet+PER_CONN_FLAGS]::PROXY_TYPE_AUTO_PROXY_URL -and $winInetProxy.AutoConfigUrl) { $winInetProxy.AutoConfigUrl }
        $props['Proxy'] = if ($winInetProxy.Flags -band [Win32.WinInet+PER_CONN_FLAGS]::PROXY_TYPE_PROXY -and $winInetProxy.Proxy) { $winInetProxy.Proxy }
        $props['ProxyBypass'] = if ($winInetProxy.Flags -band [Win32.WinInet+PER_CONN_FLAGS]::PROXY_TYPE_PROXY -and $winInetProxy.ProxyBypass) { $winInetProxy.ProxyBypass }

        # This data is temporary.
        if (-not $activeConnAdded -and $currentUserActiveConnProxy) {
            $props['ActiveConnectionProxy'] = $currentUserActiveConnProxy
            $activeConnAdded = $true
        }

        [PSCustomObject]$props
    }
}

<#
.SYNOPSIS
    Helper function to marshal an unmanaged string to a managed string.
    This function will GlobaFree the given pointer.

.Notes
    Not used currently. I'll let SafeHandle-derived classes to take care of resource release and string data marshaling.
    See SafeGlobalHandle defined in Win32Interop type definition.
#>
function MarshalString {
    [CmdletBinding(PositionalBinding = $false)]
    param(
        [Parameter(Position = 0)]
        [IntPtr]$Ptr,
        [ValidateSet('Ansi', 'Unicode')]
        [string]$UnmanagedStringType = 'Unicode'
    )

    switch ($UnmanagedStringType) {
        'Ansi' { [Runtime.InteropServices.Marshal]::PtrToStringAnsi($Ptr); break }
        'Unicode' { [Runtime.InteropServices.Marshal]::PtrToStringUni($Ptr); break }
    }

    # Don't use [Runtime.InteropServices.Marshal]::FreeHGlobal($Ptr) here because it uses LocalFree(), not GlobalFree().
    $null = [Win32.Kernel32]::GlobalFree($Ptr)
}

<#
.SYNOPSIS
    Get proxy auto config (PAC) URL & file of WinInet proxy settings.
    It tries both manual PAC URL and WPAD protocol.

.Link
    Web Proxy Auto-Discovery Protocol
    https://datatracker.ietf.org/doc/html/draft-ietf-wrec-wpad-01
#>
function Get-ProxyAutoConfig {
    [CmdletBinding(PositionalBinding = $false)]
    param(
        # Only detect wpad URL, skip downloading.
        [Parameter(Position = 0)]
        [string]$User,
        [switch]$SkipDownload
    )

    # Helper function to download a PAC file.
    # Not using Invoke-RestMethod to avoid garbled text, for a pac file sometimes contains DBCS without correct charset spec.
    function Get-PAC {
        [CmdletBinding()]
        param($Url)

        $result = @{ Url = $Url }

        if ($SkipDownload) {
            $result
            return
        }

        Write-Log "Running Get-PAC with $Url"

        [System.Net.HttpWebRequest]$request = [System.Net.WebRequest]::Create($Url)
        $request.UserAgent = 'Mozilla/5.0 (Windows NT; Windows NT 10.0)'
        $request.Timeout = 10000
        $request.UseDefaultCredentials = $true
        $response = $copied = $null

        try {
            [System.Net.HttpWebResponse]$response = $request.GetResponse()
            $body = $response.GetResponseStream()
            $copied = New-Object System.IO.MemoryStream
            $body.CopyTo($copied)
            $rawBody = $copied.ToArray()

            # Try decoding the data
            if ($response.ContentType -match 'charset=(?<charset>[\w-]+)') {
                $charset = $Matches['charset']
            }
            else {
                $charset = 'utf-8'
            }

            $bodyString = $null
            try {
                [System.Text.Encoding]$encoding = [System.Text.Encoding]::GetEncoding($charset)
                $bodyString = $encoding.GetString($rawBody)

                if ($bodyString.Contains([char]::ConvertFromUtf32(0x0000FFFD))) {
                    # might be a Shift-JIS string.
                    $bodyString = [System.Text.Encoding]::GetEncoding('shift-jis').GetString($rawBody)
                }
            }
            catch {
                Write-Error $_
            }

            $result.Add('Headers', $response.Headers.ToString())
            $result.Add('RawBody', $rawBody)
            $result.Add('Pac', $bodyString)
        }
        catch {
            Write-Error -Message "Failed to download a PAC file from $Url" -Exception $_.Exception
        }
        finally {
            if ($response) { $response.Dispose() }
            if ($copied) { $copied.Dispose() }
        }

        $result
    }

    Get-WinInetProxy -User $User | & {
        param([Parameter(ValueFromPipeline)]$proxy)
        begin {
            # Cache PAC URLs that are already tried.
            $urlCache = @{}
        }

        process {
            if ($proxy.AutoDetect) {
                [Win32.SafeGlobalFreeString]$wpadUrl = $null

                if ([Win32.WinHttp]::WinHttpDetectAutoProxyConfigUrl([Win32.WinHttp+AutoDetectType] 'WINHTTP_AUTO_DETECT_TYPE_DHCP, WINHTTP_AUTO_DETECT_TYPE_DNS_A', [ref]$wpadUrl)) {
                    $pacUrl = $wpadUrl.ToString().ToLowerInvariant()

                    if ($urlCache.ContainsKey($pacUrl)) {
                        Write-Log "Skipped a PAC URL '$pacUrl' (Connection:$($proxy.Connection)) because it's already tried"
                    }
                    else {
                        $urlCache.Add($pacUrl, $true)
                        $pac = Get-PAC $pacUrl
                        $pac.Add('IsWpad', $true)
                        $pac.Add('User', $proxy.User)
                        [PSCustomObject]$pac
                    }
                }
                else {
                    $ec = [System.Runtime.InteropServices.Marshal]::GetLastWin32Error()
                    $winhttpError = $ec -as [Win32.WinHttp+Error]

                    if ($winhttpError) {
                        Write-Error "WinHttpDetectAutoProxyConfigUrl failed with $winhttpError ($($winhttpError.value__)) for connection $($proxy.Connection)"
                    }
                    else {
                        Write-Error "WinHttpDetectAutoProxyConfigUrl failed with $ec for connection $($proxy.Connection)"
                    }
                }
            }

            if ($proxy.AutoConfigUrl) {
                $pacUrl = $proxy.AutoConfigUrl.ToLowerInvariant()

                if ($urlCache.ContainsKey($pacUrl)) {
                    Write-Log "Skipped $pacUrl because it's already tried"
                }
                else {
                    $urlCache.Add($pacUrl, $true)
                    $pac = Get-PAC $pacUrl
                    $pac.Add('IsWpad', $false)
                    $pac.Add('User', $proxy.User)
                    [PSCustomObject]$pac
                }
            }
        }
    }
}

<#
.SYNOPSIS
    Get WinHttp's default proxy
#>
function Get-WinHttpDefaultProxy {
    [CmdletBinding(PositionalBinding = $false)]
    param()

    $props = [ordered]@{}
    $proxyInfo = New-Object Win32.WinHttp+WINHTTP_PROXY_INFO

    if ([Win32.WinHttp]::WinHttpGetDefaultProxyConfiguration([ref] $proxyInfo)) {
        $props['AccessType'] = $proxyInfo.dwAccessType

        # Wrap the native string data in SafeHandle-derived class so that the memory will be properly freed (By GlobalFree in this case) when GC collects them.
        $props['Proxy'] = (New-Object Win32.SafeGlobalFreeString -ArgumentList $proxyInfo.lpszProxy).ToString()
        $props['ProxyBypass'] = (New-Object Win32.SafeGlobalFreeString -ArgumentList $proxyInfo.lpszProxyBypass).ToString()
        #$props['WINHTTP_PROXY_INFO'] = $proxyInfo # for debugging purpuse
    }
    else {
        Write-Error ("Win32 WinHttpGetDefaultProxyConfiguration failed with 0x{0:x8}" -f [System.Runtime.InteropServices.Marshal]::GetLastWin32Error())
    }

    [PSCustomObject]$props
}

<#
.SYNOPSIS
    Get WinHttp default proxy and the user's WinInet proxy settings.

.OUTPUTS
    "WinInet***" properties correspond to WINHTTP_CURRENT_USER_IE_PROXY_CONFIG obtained by WinHttpGetIEProxyConfigForCurrentUser. See https://docs.microsoft.com/en-us/windows/win32/api/winhttp/ns-winhttp-winhttp_proxy_info"
    "WinHttp***" properties correspond to WINHTTP_PROXY_INFO obtained by WinHttpGetDefaultProxyConfiguration. See https://docs.microsoft.com/en-us/windows/win32/api/winhttp/ns-winhttp-winhttp_current_user_ie_proxy_config"

.NOTES
    This function is deprecated. Use Get-WinInetProxy & Get-WinHttpDefaultProxy instead.
#>
function Get-ProxySetting {
    [CmdletBinding(PositionalBinding = $false)]
    param(
        [string]$User
    )

    $currentUser = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
    Write-Log "Running as $currentUser"

    # props hold the return object properties.
    $props = [ordered]@{}

    # Get WebProxy class to get IE config
    # N.B. GetDefaultProxy won't be really needed, but I'm keeping it for now.
    # It's possible that [System.Net.WebProxy]::GetDefaultProxy() throws
    try {
        $props['WebProxyDefault'] = [System.Net.WebProxy]::GetDefaultProxy()
    }
    catch {
        Write-Log "$_"
    }

    # Get WinHttp's default proxy
    $proxyInfo = New-Object Win32.WinHttp+WINHTTP_PROXY_INFO

    if ([Win32.WinHttp]::WinHttpGetDefaultProxyConfiguration([ref] $proxyInfo)) {
        $props['WinHttpAccessType'] = $proxyInfo.dwAccessType
        $props['WinHttpProxy'] = if ($proxyInfo.lpszProxy) { MarshalString $proxyInfo.lpszProxy }
        $props['WinHttpProxyBypass'] = if ($proxyInfo.lpszProxyBypass) { MarshalString $proxyInfo.lpszProxyBypass }
    }
    else {
        Write-Error ("Win32 WinHttpGetDefaultProxyConfiguration failed with 0x{0:x8}" -f [System.Runtime.InteropServices.Marshal]::GetLastWin32Error())
    }

    # Get User's WinInet proxy
    # If User is not specified or the given user is the current user, then just use WinHttpGetIEProxyConfigForCurrentUser; otherwise use Get-WinInetProxy for the user
    $currentUserName = $currentUser.Split('\') | Select-Object -Last 1
    if (-not $User -or $User -match $currentUserName) {
        Write-Log "Invoking WinHttpGetIEProxyConfigForCurrentUser"
        $winInetProxy = New-Object Win32.WinHttp+WINHTTP_CURRENT_USER_IE_PROXY_CONFIG

        if ([Win32.WinHttp]::WinHttpGetIEProxyConfigForCurrentUser([ref] $winInetProxy)) {
            $props['WinInetAutoDetect'] = $winInetProxy.fAutoDetect
            $props['WinINetAutoConfigUrl'] = if ($winInetProxy.lpszAutoConfigUrl) { MarshalString $winInetProxy.lpszAutoConfigUrl }
            $props['WinInetProxy'] = if ($winInetProxy.lpszProxy) { MarshalString $winInetProxy.lpszProxy }
            $props['WinInetProxyBypass'] = if ($winInetProxy.lpszProxyBypass) { MarshalString $winInetProxy.lpszProxyBypass }
        }
        else {
            Write-Error ("Win32 WinHttpGetIEProxyConfigForCurrentUser failed with 0x{0:x8}" -f [System.Runtime.InteropServices.Marshal]::GetLastWin32Error())
        }
    }
    else {
        Write-Log "`"$User`" does not match current user `"$currentUserName`". Invoking Get-WinInetProxy"
        $winInetProxy = Get-WinInetProxy -User $User
        if ($winInetProxy) {
            $props['WinInetAutoDetect'] = $winInetProxy.AutoDetect
            $props['WinInetAutoConfigUrl'] = $winInetProxy.AutoConfigUrl
            $props['WinInetProxy'] = $winInetProxy.Proxy
            $props['WinInetProxyBypass'] = $winInetProxy.ProxyBypass
            $props['User'] = if ($User) { $User } else { $currentUser }
        }
    }

    [PSCustomObject]$props
}

function Get-NLMConnectivity {
    [CmdletBinding()]
    param()

    $type = [Type]::GetTypeFromCLSID([Win32.Netlistmgr]::CLSID_NetworkListManager)
    $nlm = [Activator]::CreateInstance($type)

    $isConnectedToInternet = $nlm.IsConnectedToInternet
    [Win32.Netlistmgr+NLM_CONNECTIVITY]$connectivity = $nlm.GetConnectivity()
    Write-Log ("INetworkListManager::GetConnectivity:$connectivity (0x$("{0:x8}" -f $connectivity.value__))")

    $refCount = [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($nlm)
    Write-Log "NetworkListManager COM object's remaining ref count:$refCount"
    $nlm = $null

    [PSCustomObject]@{
        TimeStamp             = [DateTimeOffset]::Now
        IsConnectedToInternet = $isConnectedToInternet
        Connectivity          = $connectivity
    }
}

function Get-MeteredNetworkCost {
    [CmdletBinding()]
    param()

    try {
        $cost = [Win32.Netlistmgr]::GetMeteredNetworkCost()
    }
    catch {
        Write-Error -Message "GetMeteredNetworkCost failed. $($_.Exception)" -Exception $_.Exception
        return
    }

    $highCost = $false
    $conservative = $false
    $approachingHighCost = $false

    if ($cost -band [Win32.Netlistmgr+NLM_CONNECTION_COST]::NLM_CONNECTION_COST_ROAMING) {
        $highCost = $true
    }

    if ($cost -band [Win32.Netlistmgr+NLM_CONNECTION_COST]::NLM_CONNECTION_COST_FIXED `
            -or $cost -band [Win32.Netlistmgr+NLM_CONNECTION_COST]::NLM_CONNECTION_COST_VARIABLE) {
        $conservative = $true

        if ($cost -band [Win32.Netlistmgr+NLM_CONNECTION_COST]::NLM_CONNECTION_COST_OVERDATALIMIT) {
            $highCost = $true
        }

        if ($cost -band [Win32.Netlistmgr+NLM_CONNECTION_COST]::NLM_CONNECTION_COST_APPROACHINGDATALIMIT) {
            $approachingHighCost = $true
        }
    }

    if ($highCost) {
        $meteredState = 'HIGH_COST'
    }
    elseif ($approachingHighCost) {
        $meteredState = 'APPROACHING_HIGH_COST'
    }
    elseif ($conservative) {
        $meteredState = 'CONSERVATIVE'
    }
    else {
        $meteredState = 'UNRESTRICTED'
    }

    [PSCustomObject]@{
        Cost         = $cost
        MeteredState = $meteredState
    }
}

function Get-WSCAntivirus {
    [CmdletBinding()]
    param()

    [Win32.Wscapi+WSC_SECURITY_PROVIDER_HEALTH]$health = [Win32.Wscapi+WSC_SECURITY_PROVIDER_HEALTH]::WSC_SECURITY_PROVIDER_HEALTH_POOR

    # This call could fail with a terminating error on the server OS since Wscapi.dll is not available.
    # Catch it and convert it a non-terminating error so that the caller can ignore with ErrorAction.
    try {
        $hr = [Win32.Wscapi]::WscGetSecurityProviderHealth([Win32.Wscapi+WSC_SECURITY_PROVIDER]::WSC_SECURITY_PROVIDER_ANTIVIRUS, [ref]$health)
        [PSCustomObject]@{
            HRESULT = $hr
            Health  = $health
        }
    }
    catch {
        Write-Error -ErrorRecord $_
    }
}

function Get-JoinInformation {
    [CmdletBinding()]
    param()

    [IntPtr]$pName = [IntPtr]::Zero
    [Win32.Netapi32+NETSETUP_JOIN_STATUS]$status = 'NetSetupUnknownStatus'

    $sc = [Win32.Netapi32]::NetGetJoinInformation([NullString]::Value, [ref]$pName, [ref]$status)

    if ($sc -ne 0) {
        Write-Error -Message "NetGetJoinInformation failed with $sc" -Exception (New-Object ComponentModel.Win32Exception($sc))
        return
    }

    $name = [System.Runtime.InteropServices.Marshal]::PtrToStringUni($pName)
    $sc = [Win32.Netapi32]::NetApiBufferFree($pName)

    if ($sc -ne 0) {
        Write-Error -Message "NetApiBufferFree failed with $sc" -Exception (New-Object ComponentModel.Win32Exception($sc))
        return
    }

    [PSCustomObject]@{
        Name       = $name
        JoinStatus = $status
    }
}

# ***********************
# MAPI related constants
# ***********************
# https://docs.microsoft.com/en-us/office/client-developer/outlook/auxiliary/iolkaccountmanager-enumerateaccounts
$AccountManagerCLSIDs = @{
    # Categories
    CLSID_OlkMail         = '{ED475418-B0D6-11D2-8C3B-00104B2A6676}'
    CLSID_OlkAddressBook  = '{ED475419-B0D6-11D2-8C3B-00104B2A6676}'
    CLSID_OlkStore        = '{ED475420-B0D6-11D2-8C3B-00104B2A6676}'

    # Account types
    CLSID_OlkPOP3Account  = '{ED475411-B0D6-11D2-8C3B-00104B2A6676}'
    CLSID_OlkIMAP4Account = '{ED475412-B0D6-11D2-8C3B-00104B2A6676}'
    CLSID_OlkMAPIAccount  = '{ED475414-B0D6-11D2-8C3B-00104B2A6676}'
}

$KnownSections = @{
    Global         = '0a0d020000000000c000000000000046'
    MapiProvider   = '9207f3e0a3b11019908b08002b2a56c2'
    AccountManager = '9375CFF0413111d3B88A00104B2A6676'
}

$PropTags = @{
    PR_ENTRYID                           = '01020fff'
    PR_LAST_OFFLINESTATE_OFFLINE         = '00030398'
    PR_SERVICE_UID                       = '01023d0c'
    PR_STORE_PROVIDERS                   = '01023d00'
    PR_RESOURCE_TYPE                     = '00033e03'
    PR_RESOURCE_FLAGS                    = '00033009'
    PR_DISPLAY_NAME                      = '001f3001'
    PR_PROFILE_USER_SMTP_EMAIL_ADDRESS   = '001f6641'
    PR_PROFILE_PST_PATH                  = '001f6700'
    PR_EMSMDB_SECTION_UID                = '01023d15'
    PR_CACHE_SYNC_MODE                   = '0003041f'
    PR_PROFILE_OFFLINE_STORE_PATH        = '001f6610'
    PR_EMSMDB_CRED_USERNAME              = '001f3d16'
    PR_EMSMDB_CRED_DOMAINNAME            = '001f3d17'
    PR_EMSMDB_IDENTITY_UNIQUEID          = '001f3d1d'
    PR_PROFILE_CONFIG_FLAGS              = '00036601'
    PR_PROFILE_CONFIG_FLAGS_EX           = '1003666e'
    PR_PROFILE_USER_FULL_NAME            = '001f663c'
    PR_PROFILE_SYNC_MONTHS               = '00036649'
    PR_PROFILE_SYNC_DAYS                 = '0003665a'
    PR_PROFILE_ALTERNATE_STORE_TYPE      = '001f65d0'
    PR_PROFILE_TENANT_ID                 = '001f6663'
    PR_PROFILE_OFFICE365_MAILBOX         = '000b6659'
    PR_PROFILE_EXCHANGE_CONSUMER_ACCOUNT = '000b665e'
    PR_PROFILE_USER_EMAIL_ADDRESSES      = '101f6637'
    PR_AB_SEARCH_PATH_CUSTOMIZATION      = '00033d1b'
}

function Get-OutlookProfile {
    [CmdletBinding(PositionalBinding = $false)]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseDeclaredVarsMoreThanAssignments', '')]
    param(
        # Target user
        [string]$User,
        # Profile names
        [string[]]$Name
    )

    if (-not $User) {
        $User = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
    }

    $userRegRoot = Get-UserRegistryRoot $User

    if (-not $userRegRoot) {
        return
    }

    Join-Path $userRegRoot 'Software\Microsoft\Office\' `
    | Get-ChildItem -ErrorAction SilentlyContinue | . {
        param ([Parameter(ValueFromPipeline)]$key)
        process {
            if ($key.Name -match '\d\d\.0') {
                if (-not $defaultProfile) {
                    $defaultProfile = Join-Path $key.PSPath 'Outlook' | Get-ItemProperty -Name 'DefaultProfile' -ErrorAction SilentlyContinue `
                    | Select-Object -ExpandProperty 'DefaultProfile'
                }

                $cachedModePolicy = Get-CachedModePolicy -OfficeVersionKey $key
                Get-ChildItem (Join-Path $key.PSPath '\Outlook\Profiles') -ErrorAction SilentlyContinue
            }

            $key.Close()
        }
    } | & {
        param([Parameter(ValueFromPipeline)]$prof)
        process {
            $profileName = $prof.PSChildName

            try {
                if ($Name.Count -gt 0 -and $profileName -notin $Name) {
                    return
                }

                $globalSection = Get-GlobalSection $prof
                $mailAccounts = Get-MailAccount $prof
                $storeProviders = Get-StoreProvider $prof

                # Check default account
                foreach ($store in $storeProviders) {
                    if ($store.ResourceFlags.HasFlag([Win32.Mapi.ResourceFlags]::STATUS_DEFAULT_STORE)) {
                        foreach ($account in $mailAccounts) {
                            if ($account.DisplayName -eq $store.DisplayName) {
                                $account.IsDefaultAccount = $true
                                break
                            }
                        }

                        break
                    }
                }

                # Apply cache mode policy to MAPI accounts
                $mailAccounts | Merge-CachedModePolicy -CachedModePolicy $cachedModePolicy

                # Create a flattened object for data files (mailAccounts could be null if Account Manager is missing CLSID_OlkMail (ED475418-...))
                $dataFiles = $null

                if ($mailAccounts) {
                    $dataFiles = $storeProviders | Get-DataFile -MailAccounts $mailAccounts
                }
                else {
                    Write-Log "Profile '$profileName' does not have mail accounts" -Category Warning
                }

                [PSCustomObject]@{
                    User                      = $User
                    Name                      = $profileName
                    Path                      = $prof.Name
                    IsDefault                 = $profileName -eq $defaultProfile
                    Accounts                  = $mailAccounts | Select-Object -Property * -ExcludeProperty 'EmsmdbUid'
                    StoreProviders            = $storeProviders
                    DataFiles                 = $dataFiles
                    OfflineState              = $globalSection.OfflineState
                    CacheSyncMode             = $globalSection.CacheSyncMode
                    ABSearchPathCustomization = $globalSection.ABSearchPathCustomization
                }
            }
            catch {
                Write-Error -Message "Error parsing a profile '$profileName'. $_" -Exception $_.Exception
            }
            finally {
                $prof.Close()
            }
        }
    }
}

# Helper function to create a flattened object for data files.
function Get-DataFile {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, ValueFromPipeline)]
        $StoreProvider,
        [Parameter(Mandatory)]
        $MailAccounts
    )

    begin {
        $emsmdbUidCache = @{}
    }

    process {
        if ($_.PstPath) {
            [PSCustomObject]@{
                Name      = $_.DisplayName
                IsDefault = $_.ResourceFlags.HasFlag([Win32.Mapi.ResourceFlags]::STATUS_DEFAULT_STORE)
                Location  = $_.PstPath
                Size      = $_.PstSize
            }
        }
        elseif ($emsmdbUid = $_.EmsmdbUid) {
            if ($emsmdbUidCache.ContainsKey($emsmdbUid)) {
                return
            }

            $emsmdbUidCache.Add($emsmdbUid, $true)
            $account = $MailAccounts | Where-Object { $_.EmsmdbUid -eq $emsmdbUid } | Select-Object -First 1

            if ($account.OstPath) {
                [PSCustomObject]@{
                    Name      = $_.DisplayName
                    IsDefault = $_.ResourceFlags.HasFlag([Win32.Mapi.ResourceFlags]::STATUS_DEFAULT_STORE)
                    Location  = $account.OstPath
                    Size      = $account.OstSize
                }
            }
        }
    }
}

function Get-CachedModePolicy {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        $OfficeVersionKey # e.g, HKCU\Software\Microsoft\Office\16.0
    )

    # Insert 'Policies' in the path and read registry values.
    $cachedModePolicy = $OfficeVersionKey | ConvertTo-PolicyPath `
    | Join-Path -ChildPath 'outlook\cached mode' `
    | Get-ItemProperty -ErrorAction SilentlyContinue

    $props = @{}

    if ($cachedModePolicy.Enable) {
        $props.Enable = $cachedModePolicy.Enable -eq 1
    }

    if ($cachedModePolicy.CachedExchangeMode) {
        $props.SyncMode = [Win32.Mapi.CacheSyncMode]([int]::Parse($cachedModePolicy.CachedExchangeMode))
    }

    if ($syncWindow = Get-SyncWindow -Days $cachedModePolicy.SyncWindowSettingDays -Months $cachedModePolicy.SyncWindowSetting) {
        $props.SyncWindow = $syncWindow
    }

    # Dump other values as-is
    & {
        'CacheOthersMail'
        'NoManualOnlineSync'
        'NoFullItems'
        'NoDrizzle'
        'NoHeaders'
        'NoSlowHeaders'
        'ANR Include Online GAL'
        'SpecifyOfflineAddressBookPath'
    } | & {
        process {
            if ($null -ne $cachedModePolicy.$_) {
                $props.$_ = $cachedModePolicy.$_
            }
        }
    }

    [PSCustomObject]$props
}

function Merge-CachedModePolicy {
    [CmdletBinding(PositionalBinding = $false)]
    param(
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
        $Account,
        [Parameter(Mandatory = $true)]
        $CachedModePolicy
    )

    process {
        if ($Account.AccountType -ne 'MAPI') {
            return
        }

        # If CachedModePolicy is not empty, save it as is; otherwise bail
        if ($CachedModePolicy | Get-Member -MemberType Properties) {
            $Account | Add-Member -NotePropertyName 'CachedModePolicy' -NotePropertyValue $CachedModePolicy
        }
        else {
            return
        }

        $merged = New-Object System.Collections.Generic.List[string]

        if ($null -ne $cachedModePolicy.Enable) {
            $Account.IsCachedMode = $cachedModePolicy.Enable
            $merged.Add('IsCachedMode')
        }

        if ($cachedModePolicy.SyncWindow) {
            $Account.SyncWindow = $cachedModePolicy.SyncWindow
            $merged.Add('SyncWindow')
        }

        if ($merged.Count -gt 0) {
            $Account | Add-Member -NotePropertyName 'CachedModePolicyOverrides' -NotePropertyValue $merged
        }
    }
}

function Get-SyncWindow {
    [CmdletBinding(PositionalBinding = $false)]
    param(
        $Days,
        $Months
    )

    if ($null -ne $Days) {
        # When Online-mode, the value is set to signed int max (0x7fffffff)
        if ($Days -eq [int]::MaxValue) {
            return "None (Online Mode)"
        }
        elseif ($Days -gt 0) {
            return "$Days Days"
        }
    }

    # If Days is explicitly set to 0 then default to 12 months.
    if ($Days -eq 0 -and $null -eq $Months) {
        $Months = 12
    }

    if ($null -ne $Months) {
        # When Online-mode, the value is set to signed int max (0x7fffffff)
        if ($Months -eq [int]::MaxValue) {
            return "None (Online Mode)"
        }
        elseif ($Months -eq 0) {
            return 'All'
        }
        elseif ($Months -lt 12) {
            return "$Months Month$(if ($Months -gt 1) { 's' })"
        }
        else {
            return "$($Months / 12) Year$(if ($Months -gt 12) { 's' })"
        }
    }
}

function Get-MailAccount {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        $Profile
    )

    $accountManager = Join-Path $Profile.PSPath $KnownSections.AccountManager | Get-ItemProperty -Name $AccountManagerCLSIDs.CLSID_OlkMail -ErrorAction SilentlyContinue
    $accountIdBin = $accountManager.$($AccountManagerCLSIDs.CLSID_OlkMail)
    $accountCount = $accountIdBin.Count / 4

    for ($i = 0; $i -lt $accountCount; ++$i) {
        $accountId = "{0:x8}" -f [BitConverter]::ToInt32($accountIdBin, $i * 4)
        $account = Join-Path $accountManager.PsPath $accountId | Get-ItemProperty

        $acct = switch ($account.clsid) {
            $AccountManagerCLSIDs.CLSID_OlkPOP3Account { Get-Pop3Account $account; break }
            $AccountManagerCLSIDs.CLSID_OlkIMAP4Account { Get-Imap4Account $account; break }
            $AccountManagerCLSIDs.CLSID_OlkMAPIAccount { Get-MapiAccount $account; break }
        }

        # if ($i -eq 0) {
        #     $acct.IsDefaultAccount = $true
        # }

        $acct.Profile = $Profile.PSChildName
        $acct
    }
}

function Get-Pop3Account {
    [CmdletBinding()]
    param(
        $Account
    )

    $Pop3DefaultPort = 110
    $SmtpDefaultPort = 25

    [PSCustomObject]@{
        Profile              = $null
        AccountName          = $Account.'Account Name'
        AccountType          = 'POP3'
        IsDefaultAccount     = $false
        DisplayName          = $Account.'Display Name'
        Email                = $Account.Email
        Pop3Server           = $Account.'POP3 Server'
        Pop3Port             = if ($Account.'Pop3 Port') { $Account.'Pop3 Port' } else { $Pop3DefaultPort }
        Pop3User             = $Account.'POP3 User'
        SmtpServer           = $Account.'SMTP Server'
        SmtpPort             = if ($Account.'SMTP Port') { $Account.'SMTP Port' } else { $SmtpDefaultPort }
        SmtpUser             = $Account.'SMTP User'
        SmtpUseAuth          = $Account.'SMTP Use Auth' -eq 1
        SmtpUseSPA           = $Account.'SMTP Use SPA' -eq 1
        SmtpSecureConnection = switch ($Account.'SMTP Secure Connection') { 0 { 'None'; break } 1 { 'SSL/TLS'; break } 2 { 'STARTTLS'; break } 3 { 'Auto'; break } }
    }
}

function Get-Imap4Account {
    [CmdletBinding()]
    param(
        $Account
    )

    $ImapDefaultPort = 143
    $SmtpDefaultPort = 25

    [PSCustomObject]@{
        Profile              = $null
        AccountName          = $Account.'Account Name'
        AccountType          = 'IMAP4'
        IsDefaultAccount     = $false
        DisplayName          = $Account.'Display Name'
        Email                = $Account.Email
        ImapServer           = $Account.'IMAP Server'
        ImapPort             = if ($Account.'IMAP Port') { $Account.'IMAP Port' } else { $ImapDefaultPort }
        ImapUser             = $Account.'IMAP User'
        ImapUseSPA           = $Account.'IMAP Use SPA' -eq 1
        SmtpServer           = $Account.'SMTP Server'
        SmtpPort             = if ($Account.'SMTP Port') { $Account.'SMTP Port' } else { $SmtpDefaultPort }
        SmtpUser             = $Account.'SMTP User'
        SmtpUseAuth          = $Account.'SMTP Use Auth' -eq 1
        SmtpUseSPA           = $Account.'SMTP Use SPA' -eq 1
        SmtpSecureConnection = switch ($Account.'SMTP Secure Connection') { 0 { 'None'; break } 1 { 'SSL/TLS'; break } 2 { 'STARTTLS'; break } 3 { 'Auto'; break } }
    }
}

function Get-GlobalSection {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        $Profile
    )

    $properties = @(
        $PropTags.PR_LAST_OFFLINESTATE_OFFLINE
        $PropTags.PR_CACHE_SYNC_MODE
        $PropTags.PR_AB_SEARCH_PATH_CUSTOMIZATION
    )

    $globalSection = Join-Path $Profile.PSPath $KnownSections.Global | Get-ItemProperty -Name $properties -ErrorAction SilentlyContinue

    # It's possible that Global Section does not have any of above properties and thus null.
    $offlineState = 'Unknown'
    $cacheSyncMode = [Win32.Mapi.CacheSyncMode]::FullItems
    $ABSearchPathCustomization = 1 # "Start with Global Address List"

    if ($globalSection) {
        if ($offlineStateBin = $globalSection.$($PropTags.PR_LAST_OFFLINESTATE_OFFLINE)) {
            # In rare circumstances, offlineStateBin can be 0. In this case, casting to Win32.Mapi.OfflineState fails.
            try {
                [Win32.Mapi.OfflineState]$offlineState = [BitConverter]::ToInt32($offlineStateBin, 0) -band [Win32.Mapi.OfflineState]::Mask
            }
            catch {
                # ignore
            }
        }

        if ($syncModeBin = $globalSection.$($PropTags.PR_CACHE_SYNC_MODE)) {
            $cacheSyncMode = [BitConverter]::ToInt32($syncModeBin, 0)
        }

        if ($ABSearchPathCustomizationBin = $globalSection.$($PropTags.PR_AB_SEARCH_PATH_CUSTOMIZATION)) {
            $ABSearchPathCustomization = [BitConverter]::ToInt32($ABSearchPathCustomizationBin, 0)
        }

        $ABSearchPathCustomizationName = switch ($ABSearchPathCustomization) {
            0 { 'Custom'; break }
            1 { 'Start with Global Address List'; break }
            2 { 'Start with contact folders'; break }
            default { 'Start with Global Address List'; break }
        }

        [PSCustomObject]@{
            DisplayName               = 'Outlook Global Section'
            Uid                       = $KnownSections.Global
            OfflineState              = $offlineState
            CacheSyncMode             = $cacheSyncMode
            ABSearchPathCustomization = $ABSearchPathCustomizationName
        }
    }
}

function Get-StoreProvider {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        $Profile
    )

    if ($Profile -is [string]) {
        if (-not $Profile.StartsWith('Registry::')) {
            $Local:Profile = 'Registry::' + $Profile
        }

        $Local:Profile = Get-Item $Profile
    }

    $accountManager = Join-Path $Profile.PSPath $KnownSections.AccountManager | Get-ItemProperty -Name $AccountManagerCLSIDs.CLSID_OlkStore -ErrorAction SilentlyContinue
    $storeBin = $accountManager | Select-Object -ExpandProperty $AccountManagerCLSIDs.CLSID_OlkStore
    $accountCount = $storeBin.Count / 4

    $storeProviderProps = @(
        $PropTags.PR_ENTRYID
        $PropTags.PR_DISPLAY_NAME
        $PropTags.PR_RESOURCE_FLAGS
        $PropTags.PR_PROFILE_PST_PATH
        $PropTags.PR_PROFILE_ALTERNATE_STORE_TYPE
        $PropTags.PR_PROFILE_USER_SMTP_EMAIL_ADDRESS
        $PropTags.PR_PROFILE_TENANT_ID
        $PropTags.PR_EMSMDB_SECTION_UID
    )

    for ($i = 0; $i -lt $accountCount; ++$i) {
        $accountId = "{0:x8}" -f [BitConverter]::ToInt32($storeBin, $i * 4)
        $account = Join-Path $accountManager.PSPath $accountId | Get-ItemProperty
        $serviceUid = [BitConverter]::ToString($account.'Service UID').Replace('-', '').ToLowerInvariant()
        $service = Join-Path $Profile.PSPath $serviceUid | Get-ItemProperty -Name $PropTags.PR_STORE_PROVIDERS -ErrorAction SilentlyContinue
        $storeProvidersBin = $service.$($PropTags.PR_STORE_PROVIDERS)
        $storeProvidersCount = $storeProvidersBin.Count / 16

        for ($j = 0; $j -lt $storeProvidersCount; ++$j) {
            $storeUid = [BitConverter]::ToString($storeProvidersBin, $j * 16, 16).Replace('-', '')
            $store = Join-Path $Profile.PSPath $storeUid | Get-ItemProperty -Name $storeProviderProps -ErrorAction SilentlyContinue

            $props = [ordered]@{}

            if ($displayNameBin = $store.$($PropTags.PR_DISPLAY_NAME)) {
                $props.DisplayName = Get-MapiString $displayNameBin
            }

            if ($resourceFlagsBin = $store.$($PropTags.PR_RESOURCE_FLAGS)) {
                $props.ResourceFlags = [Win32.Mapi.ResourceFlags][BitConverter]::ToUInt32($resourceFlagsBin, 0)
            }

            if ($alternateStoreTypeBin = $store.$($PropTags.PR_PROFILE_ALTERNATE_STORE_TYPE)) {
                $props.AlternateStoreType = Get-MapiString $alternateStoreTypeBin
            }

            if ($pstPath = $store.$($PropTags.PR_PROFILE_PST_PATH)) {
                $props.PstPath = Get-MapiString $pstPath
                $props.PstSize = 'Unknown'

                if ($props.PstPath -and (Test-Path $props.PstPath)) {
                    if ($size = Get-ItemProperty $props.PstPath -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Length | Format-ByteSize) {
                        $props.PstSize = $size
                    }
                }
                else {
                    $props.PstSize = 'Path does not exist'
                }
            }

            if ($userSmtpEmailAddressBin = $store.$($PropTags.PR_PROFILE_USER_SMTP_EMAIL_ADDRESS)) {
                $props.UserSmtpEmailAddress = Get-MapiString $userSmtpEmailAddressBin
            }

            if ($tenantIdBin = $store.$($PropTags.PR_PROFILE_TENANT_ID)) {
                $props.TenantId = Get-MapiString $tenantIdBin
            }

            if ($entryIdBin = $store.$($PropTags.PR_ENTRYID)) {
                $props.EntryId = [BitConverter]::ToSTring($entryIdBin).Replace('-', [String]::Empty).ToLowerInvariant()
            }

            if ($emsmdbUidBin = $store.$($PropTags.PR_EMSMDB_SECTION_UID)) {
                $props.EmsmdbUid = [BitConverter]::ToString($emsmdbUidBin).Replace('-', [String]::Empty).ToLowerInvariant()
            }

            [PSCustomObject]$props
        }
    }
}

function Get-MapiAccount {
    [CmdletBinding()]
    param(
        $Account
    )

    # Get Profile root path
    $profRoot = $Account.PSPath.SubString(0, $Account.PSPath.IndexOf($KnownSections.AccountManager))

    $serviceUid = [BitConverter]::ToString($Account.'Service UID').Replace('-', [String]::Empty)
    $service = Join-Path $profRoot $serviceUid | Get-ItemProperty -Name $PropTags.PR_EMSMDB_SECTION_UID -ErrorAction SilentlyContinue
    $emsmdbUid = [BitConverter]::ToString($service.$($PropTags.PR_EMSMDB_SECTION_UID)).Replace('-', '').ToLowerInvariant()

    # Get EMSMDB section properties
    $emsmdbProperties = @(
        $PropTags.PR_DISPLAY_NAME
        $PropTags.PR_EMSMDB_CRED_DOMAINNAME
        $PropTags.PR_EMSMDB_CRED_USERNAME
        $PropTags.PR_EMSMDB_IDENTITY_UNIQUEID
        $PropTags.PR_PROFILE_USER_FULL_NAME
        $PropTags.PR_PROFILE_OFFLINE_STORE_PATH
        $PropTags.PR_PROFILE_CONFIG_FLAGS
        $PropTags.PR_PROFILE_CONFIG_FLAGS_EX
        $PropTags.PR_PROFILE_SYNC_MONTHS
        $PropTags.PR_PROFILE_SYNC_DAYS
        $PropTags.PR_PROFILE_OFFICE365_MAILBOX
        $PropTags.PR_PROFILE_EXCHANGE_CONSUMER_ACCOUNT
        $PropTags.PR_PROFILE_USER_EMAIL_ADDRESSES
    )

    $emsmdb = Join-Path $profRoot $emsmdbUid | Get-ItemProperty -Name $emsmdbProperties -ErrorAction SilentlyContinue

    $props = [ordered]@{
        Profile          = $null
        AccountType      = 'MAPI'
        IsDefaultAccount = $false
        EmsmdbUid        = $emsmdbUid
    }

    if ($Account.'Delivery Folder EntryID') {
        $props.DeliveryFolderEntryID = [BitConverter]::ToString($Account.'Delivery Folder EntryID').Replace('-', [String]::Empty)
    }

    if ($displayNameBin = $emsmdb.$($PropTags.PR_DISPLAY_NAME)) {
        $props.DisplayName = Get-MapiString $displayNameBin
    }

    if ($credDomainName = $emsmdb.$($PropTags.PR_EMSMDB_CRED_DOMAINNAME)) {
        $props.CredentialDomainName = Get-MapiString $credDomainName
    }

    if ($credUserName = $emsmdb.$($PropTags.PR_EMSMDB_CRED_USERNAME)) {
        $props.CredentialUserName = Get-MapiString $credUserName
    }

    if ($identityUniqueIdBin = $emsmdb.$($PropTags.PR_EMSMDB_IDENTITY_UNIQUEID)) {
        $props.IdentityUniqueId = Get-MapiString $identityUniqueIdBin
    }

    if ($userFullNameBin = $emsmdb.$($PropTags.PR_PROFILE_USER_FULL_NAME)) {
        $props.UserFullName = Get-MapiString $userFullNameBin
    }

    if ($ostPath = $emsmdb.$($PropTags.PR_PROFILE_OFFLINE_STORE_PATH)) {
        $props.OstPath = Get-MapiString $ostPath
        $props.OstSize = 'Unknown'

        if ($props.OstPath -and (Test-Path $props.OstPath)) {
            if ($size = Get-ItemProperty $props.OstPath -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Length) {
                $props.OstSize = Format-ByteSize $size
            }
        }
        else {
            $props.OstSize = 'Path does not exist'
        }
    }

    if ($configFlagsBin = $emsmdb.$($PropTags.PR_PROFILE_CONFIG_FLAGS)) {
        [Win32.Mapi.ProfileConfigFlags]$configFlags = [BitConverter]::ToUInt32($configFlagsBin, 0)
        $props.IsCachedMode = $configFlags.HasFlag([Win32.Mapi.ProfileConfigFlags]::CONFIG_OST_CACHE_PRIVATE)
        $props.DownloadPublicFolderFavorites = $configFlags.HasFlag([Win32.Mapi.ProfileConfigFlags]::CONFIG_OST_CACHE_PUBLIC)
        $props.DownloadSharedFolders = $configFlags.HasFlag([Win32.Mapi.ProfileConfigFlags]::CONFIG_OST_CACHE_DELEGATE_PIM)
    }

    if ($isOffice365MailboxBin = $emsmdb.$($PropTags.PR_PROFILE_OFFICE365_MAILBOX)) {
        $props.IsOffice365Mailbox = [System.BitConverter]::ToInt16($isOffice365MailboxBin, 0) -eq 1
    }

    if ($isConsumerAccountBin = $emsmdb.$($PropTags.PR_PROFILE_EXCHANGE_CONSUMER_ACCOUNT)) {
        $props.IsConsumerAccount = [System.BitConverter]::ToInt16($isConsumerAccountBin, 0) -eq 1
    }

    if ($emailAddressesBin = $emsmdb.$($PropTags.PR_PROFILE_USER_EMAIL_ADDRESSES)) {
        # TODO: To be removed later
        try {
            $props.UserEmailAddresses = Convert-MVUnicode $emailAddressesBin
        }
        catch {
            Write-Error -Message "Convert-MVUnicode failed. $_" -Exception $_.Exception
        }
    }

    # Get Sync Window
    $syncMonths = $null
    $syncDays = $null

    if ($syncMonthsBin = $emsmdb.$($PropTags.PR_PROFILE_SYNC_MONTHS)) {
        $syncMonths = [System.BitConverter]::ToInt32($syncMonthsBin, 0)
    }

    if ($syncDaysBin = $emsmdb.$($PropTags.PR_PROFILE_SYNC_DAYS)) {
        $syncDays = [System.BitConverter]::ToInt32($syncDaysBin, 0)
    }

    $props.SyncWindow = Get-SyncWindow -Days $syncDays -Months $syncMonths

    # Shared Calendar (low 2 bytes of PR_PROFILE_CONFIG_FLAGS_EX)
    if ($configFlagsExBin = $emsmdb.$($PropTags.PR_PROFILE_CONFIG_FLAGS_EX)) {
        [Win32.Mapi.SharedCalProfileConfigFlags]$props.SharedCalendarOption = [BitConverter]::ToInt32($configFlagsExBin, 0) -band 0xffff
    }

    # Connected Experience
    if ($props.IdentityUniqueId) {
        $props.ConnectedExperienceEnabled = Get-ConnectedExperience $props.IdentityUniqueId | Select-Object -ExpandProperty 'Enabled'
    }

    [PSCustomObject]$props
}

function Format-ByteSize {
    [CmdletBinding()]
    [OutputType([string])]
    param (
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
        $Size
    )

    $suffix = "B", "KB", "MB", "GB", "TB", "PB"
    $index = 0

    while ($Size -gt 1kb) {
        $Size = $Size / 1kb
        $index++
    }

    "{0:N2} {1}" -f $Size, $suffix[$index]
}

<#
.SYNOPSIS
    Get a string from a byte array of PT_STRING8 or MAPI PT_UNICODE value.
    The data is interpreted as PT_UNICODE by default. Use "Ascii" switch for PT_STRING8.
#>
function Get-MapiString {
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory)]
        [byte[]]$Bin,
        # Interpret as PT_STRING8
        [switch]$Ascii
    )

    if ($Ascii) {
        # PT_STRING8 ends with a single NULL (0x00)
        $terminatingNullCount = 1
        [System.Text.Encoding]::ASCII.GetString($Bin, 0, $Bin.Length - $terminatingNullCount)
    }
    else {
        # PT_UNICODE ends with a double-byte NULL (0x00 00)
        $terminatingNullCount = 2
        [System.Text.Encoding]::Unicode.GetString($Bin, 0, $Bin.Length - $terminatingNullCount)
    }
}

<#
.SYNOPSIS
    Parse PT_MV_UNICODE value and return an array of strings
#>
function Convert-MVUnicode {
    [CmdletBinding()]
    [OutputType([string[]])]
    param(
        [Parameter(Mandatory)]
        [byte[]]$Bin
    )

    $reader = $null

    try {
        $stream = New-Object System.IO.MemoryStream -ArgumentList (, $Bin)
        $reader = New-Object System.IO.BinaryReader $stream

        # Number of strings
        $count = $reader.ReadInt32()

        # Next 4 or 8 bytes are offset to the start of each string (4 bytes for 32 bit Outlook, 8 bytes 64 bit Outlook)
        # Note: For some unknown reason, sometimes there are mix of both 4 bytes and 8 bytes in different accounts. So, can't just rely on the Office bitness.
        # Because these offsets must be less than 32-bit int max, read 8 bytes first, then if it's greater than 0xffffffff, assume 4 byte offsets.
        $offsets = @(
            for ($i = 0; $i -lt $count; ++$i) {
                $offset64 = $reader.ReadInt64()

                # suffix "L" is for long data type
                if ($offset64 -le 0xffffffffL) {
                    $offset64
                }
                else {
                    # Rewind and read as 4-byte int
                    $reader.BaseStream.Position -= 8
                    $reader.ReadInt32()
                }
            }
        )

        $reader.BaseStream.Position = $offsets[0]

        @(
            for ($i = 0; $i -lt $count; ++$i) {
                $currentOffset = $offsets[$i]

                if ($i -lt $count - 1) {
                    $nextOffset = $offsets[$i + 1]
                }
                else {
                    # For the last string, read to the end
                    $nextOffset = $reader.BaseStream.Length
                }

                $bytes = $reader.ReadBytes($nextOffset - $currentOffset)
                [System.Text.Encoding]::Unicode.GetString($bytes)
            }
        )
    }
    finally {
        if ($reader) {
            $reader.Dispose()
        }
    }
}

<#
.SYNOPSIS
    Get corrupt file information stored at 'HKCU\Software\Microsoft\Office\16.0\MAPI\Stores\CorruptFiles'
#>
function Get-MapiCorruptFiles {
    [CmdletBinding()]
    param(
        $User
    )

    $userRegRoot = Get-UserRegistryRoot -User $User

    if (-not $userRegRoot) {
        return
    }

    # https://github.com/MicrosoftDocs/OfficeDocs-DeployOffice/blob/public/DeployOffice/privacy/required-diagnostic-data.md
    $NdbType = @(
        'NdbUndefined' # 0
        'NdbSmall'     # 1
        'NdbLarge'     # 2
        'NdbTardis'    # 3
    )

    $corruptFiles = Join-Path $userRegRoot 'Software\Microsoft\Office\16.0\MAPI\Stores\CorruptFiles' | Get-ItemProperty -ErrorAction SilentlyContinue | Select-Object -Property * -ExcludeProperty 'PS*'

    if ($corruptFiles) {
        # Enumerate properties.
        # Reg value is Binary type (e.g. "2024-02-16-06:29::59.0609"=hex:03,00,....)
        foreach ($prop in ($corruptFiles | Get-Member -MemberType Properties)) {
            # Registry name is timestamp. But this string is not formatted in such a way that [DateTime] can parse. So just leave it as is.
            $time = $prop.Name
            $dataBin = $corruptFiles.$time
            $corruptState = [System.BitConverter]::ToUInt32($dataBin, 0)
            $corruptAction = $corruptState -shr 16
            $corruptType = $corruptState -band 0xffff
            $fileName = [System.Text.Encoding]::Unicode.GetString($dataBin[4..($dataBin.Length - 1)])

            [PSCustomObject]@{
                CorruptAction = $corruptAction
                CorruptType   = $NdbType[$corruptType]
                FileName      = $fileName
                TimeStamp     = $time
            }
        }
    }
}

<#
.SYNOPSIS
    Insert "Policies" after "Software" in the given registry path.
    e.g., 'HKCU:\Software\Microsoft\Office\' --> 'HKCU:\Software\Policies\Microsoft\Office\'

    If the given path is already a policy path, then it is returned as is.
    In both cases, the returned string will be prefixed with 'Registry::'

    * If the path is already prefixed with "Registry::" or "Microsoft.PowerShell.Core\Registry::", no prefix will be added.
#>
function ConvertTo-PolicyPath {
    [CmdletBinding()]
    [OutputType([string])]
    param (
        [Parameter(Mandatory, ValueFromPipeline)]
        [string]$Path
    )

    process {
        $policyPath = $Path

        if ($Path -match '(?<Head>^.*\\Software)\\(?<Tail>.*)') {
            # If it looks like "...\Software\Policies\...", then do nothing.
            if (-not $Matches['Tail'].StartsWith('Policies\')) {
                $policyPath = Join-Path $Matches['Head'] 'Policies' `
                | Join-Path -ChildPath $Matches['Tail']
            }
        }

        # Add prefix if necessary.
        # Note that no need to add prefix if it looks like "Microsoft.PowerShell.Core\Registry::"
        if (-not ($policyPath -match 'Registry::')) {
            $policyPath = "Registry::$policyPath"
        }

        $policyPath
    }
}

<#
.SYNOPSIS
    Test if a given registry path is a policy path ("...\Software\Policies\...")
#>
function Test-PolicyPath {
    [CmdletBinding()]
    [OutputType([bool])]
    param(
        [Parameter(Mandatory, ValueFromPipeline)]
        [string]$Path
    )

    process {
        $Path -match '\\Software\\Policies\\'
    }
}

function Get-OutlookOption {
    [CmdletBinding()]
    param (
        $User
    )

    function New-Option {
        param (
            [Parameter(Mandatory)]
            $Name,
            [Parameter(Mandatory)]
            $Description,
            [Parameter(Mandatory)]
            [ValidateSet('Mail', 'Calendar', 'General', 'Tasks', 'Advanced', 'Power', 'Security', 'Setup', 'Search')]
            $Category,
            $Value
        )

        [PSCustomObject]@{
            Name        = $Name
            Description = $Description
            Category    = $Category
            Value       = $Value
            Path        = $null
            IsPolicy    = $false
        }
    }

    function Set-Option {
        param (
            [Parameter(Mandatory)]
            # Name of registry value
            $Name,
            [Parameter(Mandatory)]
            $Property,
            # Default converter just converts non-zero to $true
            [ScriptBlock]$Converter = { param ($val) $val -ne 0 },
            [Parameter(Mandatory)]
            $Options
        )

        $regValue = $Property.$Name

        if ($null -ne $regValue) {
            $option = $Options | Where-Object { $_.Name -eq $Name } | Select-Object -First 1
            $option.Value = & $Converter $regValue $Name
            $option.Path = $Property | Convert-Path -ErrorAction SilentlyContinue

            if (Test-PolicyPath $option.Path) {
                $option.IsPolicy = $true
            }
        }
    }

    $userRegRoot = Get-UserRegistryRoot -User $User

    if (-not $userRegRoot) {
        return
    }

    $officeInfo = Get-OfficeInfo
    $major = $officeInfo.Version.Split('.')[0]

    $optionsPath = Join-Path $userRegRoot "Software\Microsoft\Office\$major.0\Outlook\Options"
    $prefPath = Join-Path $userRegRoot "Software\Microsoft\Office\$major.0\Outlook\Preferences"
    $powerPath = Join-Path $userRegRoot "Software\Microsoft\Office\$major.0\Outlook\Power"
    $securityPath = Join-Path $userRegRoot "Software\Microsoft\Office\$major.0\Outlook\Security"
    $setupPath = Join-Path $userRegRoot "Software\Microsoft\Office\$major.0\Outlook\Setup"
    $searchPath = Join-Path $userRegRoot "Software\Microsoft\Office\$major.0\Outlook\Search"

    # Set default values for options I'm interested in.
    $options = @(
        New-Option -Name 'Send Mail Immediately' -Description 'Send Mail Immediately' -Category Mail -Value $true
        New-Option -Name 'SaveAllMIMENotJustHeaders' -Description 'Save entire MIME message' -Category Mail -Value $false
        New-Option -Name 'EditorPreference' -Description 'Specifies editor to use for mail items' -Category Mail -Value $null
        New-Option -Name 'NewMailDesktopAlerts' -Description 'Display a Desktop Alert' -Category Mail -Value $true
        New-Option -Name 'NewMailDesktopAlertsDRMPreview' -Description 'Enable preview for Rights Protected messages' -Category Mail -Value $false
        New-Option -Name 'SaveSent' -Description 'Save copies of messages in the Sent Items folder' -Category Mail -Value $true
        New-Option -Name 'DelegateSentItemsStyle' -Description "When set to 1, items sent on behalf of a manager will now go to the manager's sent items box" -Category Mail -Value $false
        New-Option -Name 'EnableAppsInOutlook' -Description "Shows apps in Outlook" -Category General -Value $true
        New-Option -Name 'UseNewOutlook' -Description 'Use the new Outlook for Windows' -Category General -Value $false
        New-Option -Name 'HideNewOutlookToggle' -Description 'Hide the "Try the new Outlook" toggle in Outlook Desktop' -Category General -Value $false
        New-Option -Name 'ShowLegacySharingUX' -Description 'Turn off Calendar Sharing REST API and use Legacy UI' -Category Calendar -Value $false
        New-Option -Name 'EnableMeetingCopy' -Description 'Re-enable copying meetings:' -Category Calendar -Value $false
        New-Option -Name 'ShowLegacyRoomFinder' -Description 'Show legacy room finder' -Category Calendar -Value $false
        New-Option -Name 'CalendarEditorPreference' -Description "Specifies editor to use for calendar items" -Category Calendar -Value $null
        New-Option -Name 'OpenTasksWithToDoApp' -Description 'When opening from a reminder, open tasks with ToDo App' -Category Tasks -Value $false
        New-Option -Name 'Autodetect_CodePageOut' -Description 'Automatically select encoding for outgoing messages' -Category Advanced -Value $true
        New-Option -Name 'Default_CodePageOut' -Description 'Preferred encoding for outgoing messages' -Category Advanced -Value $null
        New-Option -Name 'HighCostMeteredNetworkBehavior' -Description 'Behavior on a high cost metered network' -Category Power -Value 'Default'
        New-Option -Name 'ConservativeMeteredNetworkBehavior' -Description 'Behavior on a conservative metered network' -Category Power -Value 'Default'
        New-Option -Name 'BatteryMode' -Description 'Battery mode' -Category Power -Value 'Default'
        New-Option -Name 'MarkInternalAsUnsafe' -Description 'Use Protected View for attachments received from internal senders' -Category Security -Value $false
        New-Option -Name 'EnableUnsafeClientMailRules' -Description 'Enable executing client side mail rules which start an application, or invoke a VBA macro' -Category Security -Value $false
        New-Option -Name 'DisableOffice365SimplifiedAccountCreation' -Description 'Using simplified account creation to add an account to Outlook' -Category Setup -Value $false
        New-Option -Name 'DisableServerAssistedSearch' -Description 'Disables Outlook from requesting and using Search results from Exchange for cached and non-cached mailbox items. Instead it will use search results from windows search service' -Category Search -Value $false
        New-Option -Name 'DisableServerAssistedSuggestions' -Description 'Disables Outlook from requesting search suggestions from Exchange' -Category Search -Value $false
    )

    $PSDefaultParameterValues['Set-Option:Options'] = $options

    & {
        Join-Path $optionsPath 'Mail'
        Join-Path $optionsPath 'Mail' | ConvertTo-PolicyPath
    } | Get-ItemProperty -ErrorAction SilentlyContinue | & {
        process {
            $PSDefaultParameterValues['Set-Option:Property'] = $_
            Set-Option -Name 'Send Mail Immediately'
            Set-Option -Name 'SaveAllMIMENotJustHeaders'

            Set-Option -Name 'EditorPreference' -Converter {
                param ($regValue)
                switch (($regValue -band 0x00ff0000) -shr 16) {
                    1 { 'PlainText'; break }
                    2 { 'HTML'; break }
                    3 { 'RTF'; break }
                    default { ("Unknown (0x{0:x})" -f $regValue) }
                }
            }
        }
    }

    & {
        Join-Path $optionsPath 'Calendar'
        Join-Path $optionsPath 'Calendar' | ConvertTo-PolicyPath
    } | Get-ItemProperty -ErrorAction SilentlyContinue | & {
        process {
            $PSDefaultParameterValues['Set-Option:Property'] = $_
            Set-Option -Name 'ShowLegacySharingUX'
            Set-Option -Name 'EnableMeetingCopy'
            Set-Option -Name 'ShowLegacyRoomFinder'

            Set-Option -Name 'CalendarEditorPreference' -Converter {
                param ($regValue)
                switch (($regValue -band 0x00ff0000) -shr 16) {
                    1 { 'PlainText'; break }
                    2 { 'HTML'; break }
                    3 { 'RTF'; break }
                    default { ("Unknown (0x{0:x})" -f $regValue) }
                }
            }
        }
    }

    & {
        Join-Path $optionsPath 'General'
        Join-Path $optionsPath 'General' | ConvertTo-PolicyPath
    } | Get-ItemProperty -ErrorAction SilentlyContinue | & {
        process {
            $PSDefaultParameterValues['Set-Option:Property'] = $_
            Set-Option -Name 'HideNewOutlookToggle'
        }
    }

    & {
        $prefPath
        $prefPath | ConvertTo-PolicyPath
    } | Get-ItemProperty -ErrorAction SilentlyContinue | & {
        process {
            $PSDefaultParameterValues['Set-Option:Property'] = $_
            Set-Option -Name 'NewMailDesktopAlerts'
            Set-Option -Name 'NewmailDesktopAlertsDRMPreview'
            Set-Option -Name 'OpenTasksWithToDoApp'
            Set-Option -Name 'SaveSent'
            Set-Option -Name 'DelegateSentItemsStyle'
            Set-Option -Name 'EnableAppsInOutlook'
            Set-Option -Name 'UseNewOutlook'
        }
    }

    & {
        Join-Path $optionsPath 'MSHTML\International\'
        Join-Path $optionsPath 'MSHTML\International\' | ConvertTo-PolicyPath
    } | Get-ItemProperty -ErrorAction SilentlyContinue | & {
        process {
            $PSDefaultParameterValues['Set-Option:Property'] = $_
            Set-Option -Name 'Autodetect_CodePageOut'
            Set-Option -Name 'Default_CodePageOut' -Converter { param ($regValue) [System.Text.Encoding]::GetEncoding($regValue).WebName }
        }
    }

    & {
        $powerPath
        $powerPath | ConvertTo-PolicyPath
    } | Get-ItemProperty -ErrorAction SilentlyContinue | & {
        process {
            $PSDefaultParameterValues['Set-Option:Property'] = $_

            $meteredNetworkBehaviorConverter = {
                param ($regValue, $regName)
                switch ($regValue) {
                    0 { 'Default'; break }
                    1 { 'Ignore'; break }
                    2 { if ($regName -eq 'ConservativeMeteredNetworkBehavior') { 'TreatAsHighCost' } else { 'Invalid' }; break }
                    default { 'Invalid'; break }
                }
            }

            Set-Option -Name 'HighCostMeteredNetworkBehavior' -Converter $meteredNetworkBehaviorConverter
            Set-Option -Name 'ConservativeMeteredNetworkBehavior' -Converter $meteredNetworkBehaviorConverter

            $batteryModeConverter = {
                param ($regValue)
                switch ($regValue) {
                    0 { 'Default'; break }
                    1 { 'Always'; break }
                    2 { 'Never'; break }
                    default { 'Invalid'; break }
                }
            }

            Set-Option -Name 'BatteryMode' -Converter $batteryModeConverter
        }
    }

    & {
        $securityPath
        $securityPath | ConvertTo-PolicyPath
    } | Get-ItemProperty -ErrorAction SilentlyContinue | & {
        process {
            $PSDefaultParameterValues['Set-Option:Property'] = $_
            Set-Option -Name 'MarkInternalAsUnsafe'
            Set-Option -Name 'EnableUnsafeClientMailRules'
        }
    }

    & {
        $setupPath
        $setupPath | ConvertTo-PolicyPath
    } | Get-ItemProperty -ErrorAction SilentlyContinue | & {
        process {
            $PSDefaultParameterValues['Set-Option:Property'] = $_
            Set-Option -Name 'DisableOffice365SimplifiedAccountCreation'
        }
    }

    & {
        $searchPath
        $searchPath | ConvertTo-PolicyPath
    } | Get-ItemProperty -ErrorAction SilentlyContinue | & {
        process {
            $PSDefaultParameterValues['Set-Option:Property'] = $_
            Set-Option -Name 'DisableServerAssistedSearch'
            Set-Option -Name 'DisableServerAssistedSuggestions'
        }
    }

    $options
}

<#
.SYNOPSIS
    Get Outlook's PickLogonProfile setting.
#>
function Get-PickLogonProfile {
    [CmdletBinding()]
    [OutputType([bool])]
    param(
        $User
    )

    $userRegRoot = Get-UserRegistryRoot -User $User

    if (-not $userRegRoot) {
        return
    }

    $pickLogonProfile = Get-ItemProperty (Join-Path $userRegRoot 'SOFTWARE\Microsoft\Exchange\Client\Options') -Name 'PickLogonProfile' -ErrorAction SilentlyContinue `
    | Select-Object -ExpandProperty 'PickLogonProfile'

    if ($pickLogonProfile -eq '1') {
        $true
    }
    else {
        $false
    }
}

<#
.SYNOPSIS
    Let Outlook prompt for a profile to be used.
#>
function Enable-PickLogonProfile {
    [CmdletBinding()]
    param(
        $User
    )

    $userRegRoot = Get-UserRegistryRoot -User $User

    if (-not $userRegRoot) {
        return
    }

    $optionsPath = Join-Path $userRegRoot 'SOFTWARE\Microsoft\Exchange\Client\Options'

    if (-not (Test-Path $optionsPath)) {
        $null = New-Item $optionsPath -Force -ErrorAction Stop
    }

    # Type is string
    Set-ItemProperty $optionsPath -Name 'PickLogonProfile' -Value '1'
}

<#
.SYNOPSIS
    Let Outlook not prompt for a profile to be used.
#>
function Disable-PickLogonProfile {
    [CmdletBinding()]
    param(
        $User
    )

    $userRegRoot = Get-UserRegistryRoot -User $User

    if (-not $userRegRoot) {
        return
    }

    $optionsPath = Join-Path $userRegRoot 'SOFTWARE\Microsoft\Exchange\Client\Options'

    if (Test-Path $optionsPath) {
        Set-ItemProperty $optionsPath -Name 'PickLogonProfile' -Value '0'
    }
}

function Disable-AccountSetupV2 {
    [CmdletBinding()]
    param(
        $User
    )

    $userRegRoot = Get-UserRegistryRoot -User $User

    if (-not $userRegRoot) {
        return
    }

    $officeInfo = Get-OfficeInfo
    $major = $officeInfo.Version.Split('.')[0]
    $setupPath = Join-Path $userRegRoot "Software\Microsoft\Office\$major.0\Outlook\Setup\"

    if (-not (Test-Path $setupPath)) {
        $null = New-Item $setupPath -ErrorAction Stop
    }

    Set-ItemProperty $setupPath -Name 'DisableOffice365SimplifiedAccountCreation' -Value 1 -Type DWord
}

function Enable-AccountSetupV2 {
    [CmdletBinding()]
    param(
        $User
    )

    $userRegRoot = Get-UserRegistryRoot -User $User

    if (-not $userRegRoot) {
        return
    }

    $officeInfo = Get-OfficeInfo
    $major = $officeInfo.Version.Split('.')[0]
    $setupPath = Join-Path $userRegRoot "Software\Microsoft\Office\$major.0\Outlook\Setup\"

    if (-not (Test-Path $setupPath)) {
        return
    }

    Remove-ItemProperty $setupPath -Name 'DisableOffice365SimplifiedAccountCreation'
}

function Get-WordMailOption {
    [CmdletBinding()]
    param(
        $User
    )

    $userRegRoot = Get-UserRegistryRoot -User $User

    if (-not $userRegRoot) {
        return
    }

    $officeInfo = Get-OfficeInfo
    $major = $officeInfo.Version.Split('.')[0]

    Join-Path $userRegRoot "Software\Microsoft\Office\$major.0\Word\Options\WordMail" `
    | Get-ItemProperty `
    | Select-Object -Property '*' -ExcludeProperty 'PSParentPath', 'PSChildName', 'PSProvider'
}

<#
.SYNOPSIS
Helper function to get the locations of cached Autodiscover XML files.
#>
function Get-CachedAutodiscoverLocation {
    [CmdletBinding()]
    param(
        $User
    )

    # Check %LOCALAPPDATA%\Microsoft\Outlook and path specified by "ForcePSTPath" registry value.

    # LOCALAPPDATA
    if ($localAppdata = Get-UserShellFolder -User $User -ShellFolderName 'Local AppData') {
        [PSCustomObject]@{
            Name = 'UnderLocalAppData'
            Path = Join-Path $localAppdata -ChildPath 'Microsoft\Outlook'
        }
    }

    # ForcePSTPath if any
    $userRegRoot = Get-UserRegistryRoot -User $User

    if (-not $userRegRoot) {
        return
    }

    $officeInfo = Get-OfficeInfo

    if (-not $officeInfo) {
        return
    }

    $ver = ($officeInfo.Version.Split('.')[0] -as [int]).ToString('00.0')
    $ForcePSTPath = 'ForcePSTPath'

    & {
        "SOFTWARE\Policies\Microsoft\Office\$ver\Outlook"
        "SOFTWARE\Microsoft\Office\$ver\Outlook"
    } `
    | Join-Path -Path $userRegRoot -ChildPath { $_ } `
    | Get-ItemProperty -Name $ForcePSTPath -ErrorAction SilentlyContinue | & {
        process {
            [PSCustomObject]@{
                Name = $ForcePSTPath
                Path = [System.Environment]::ExpandEnvironmentVariables($_.$ForcePSTPath)
            }
        }
    } `
    | Select-Object -First 1  # If ForcePSTPath is found in the policy key, no need to check the rest.
}

<#
.SYNOPSIS
Save cached Autodiscover XML files
#>
function Save-CachedAutodiscover {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        # Where to save
        $Path,
        # Target user name
        [string]$User
    )

    if ($User -and -not (Resolve-User $User)) {
        return
    }

    if (-not (Test-Path $Path)) {
        $null = New-Item $Path -ItemType Directory -ErrorAction Stop
    }

    $cachePaths = Get-CachedAutodiscoverLocation -User $User

    foreach ($cachePath in $cachePaths) {
        Write-Log "Searching $($cachePath.Name) $($cachePath.Path)"

        $saveArgs = @{
            Filter        = '*Autod*.xml'
            IncludeHidden = $true
            PassThru      = $true
        }

        & {
            Save-Item -Path $cachePath.Path -Destination $Path @saveArgs

            if ($cachePath.Name -eq 'UnderLocalAppData') {
                Save-Item -Path "$($cachePath.Path)\16" -Destination "$Path\16" @saveArgs
            }
        } | Remove-HiddenAttribute
    }
}

<#
.SYNOPSIS
Remove cached Autodiscover XML files
#>
function Remove-CachedAutodiscover {
    [CmdletBinding()]
    param(
        # Target user name
        [string]$User
    )

    if ($User -and -not (Resolve-User $User)) {
        return
    }

    Get-CachedAutodiscoverLocation -User $User | & {
        param ([Parameter(ValueFromPipeline)]$cachePath)
        process {
            Get-ChildItem -LiteralPath $cachePath.Path -Filter '*Autod*.xml' -Force -Recurse:($cachePath.Name -eq 'UnderLocalAppData') `
            | Remove-Item -Force
        }
    }
}

<#
.SYNOPSIS
Save cached Outlook/OPX config json files
#>
function Save-CachedOutlookConfig {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        # Where to save
        $Path,
        # Target user name
        [string]$User
    )

    $LocalAppData = 'Local AppData'
    $sourcePath = Get-UserShellFolder -User $User -ShellFolderName $LocalAppData | Join-Path -ChildPath 'Microsoft\Outlook\16'

    if (-not $sourcePath) {
        Write-Error "Cannot find $LocalAppData for $User"
        return
    }

    if (-not (Test-Path $Path)) {
        $null = New-Item $Path -ItemType Directory -ErrorAction Stop
    }

    $Path = Convert-Path -LiteralPath $Path

    Save-Item -Path $sourcePath -Filter '*Config*.json' -IncludeHidden -Destination $Path -PassThru | Remove-HiddenAttribute
}

function Remove-HiddenAttribute {
    [CmdletBinding()]
    param(
        [Parameter(ValueFromPipeline = $true)]
        $File
    )

    process {
        try {
            if ((Get-ItemProperty $File.FullName).Attributes -band [IO.FileAttributes]::Hidden) {
                (Get-ItemProperty $File.FullName).Attributes -= 'Hidden'
                return
            }
        }
        catch {
            # ignore
        }

        # This could fail if attributes other than Archive, Hidden, Normal, ReadOnly, or System are set (such as NotContentIndexed)
        Set-ItemProperty $File.Fullname -Name Attributes -Value ((Get-ItemProperty $File.FullName).Attributes -bxor [IO.FileAttributes]::Hidden)
    }
}

<#
.SYNOPSIS
Remove cached OutlookConfig/OPX json files
#>
function Remove-CachedOutlookConfig {
    [CmdletBinding()]
    param(
        # Target user name
        [string]$User
    )

    if ($localAppdata = Get-UserShellFolder -User $User -ShellFolderName 'Local AppData') {
        $sourcePath = Join-Path $localAppdata -ChildPath 'Microsoft\Outlook'
    }
    else {
        Write-Error "Cannot find LocalAppData for $User"
    }

    Get-ChildItem -LiteralPath $sourcePath -Filter '*Config*.json' -Force -Recurse | Remove-Item -Force
}

<#
.SYNOPSIS
Remove cached Identities & Tokens
#>
function Remove-IdentityCache {
    [CmdletBinding()]
    param(
        [string]$User
    )

    # You need to be elevated to restart TokenBroker service
    if (-not (Test-RunAsAdministrator)) {
        Write-Error "Please run as administrator"
        return
    }

    # Remove Office Identity registry sub keys
    $userRegRoot = Get-UserRegistryRoot -User $User

    if (-not $userRegRoot) {
        return
    }

    Join-Path $userRegRoot 'Software\Microsoft\Office\16.0\Common\Identity' `
    | Get-ChildItem -ErrorAction SilentlyContinue `
    | Remove-Item -Recurse -Force

    $localAppData = Get-UserShellFolder -User $User -ShellFolderName 'Local AppData'

    if (-not $localAppData) {
        return
    }

    $TokenBrokerService = 'TokenBroker'
    Set-Service $TokenBrokerService -StartupType Disabled
    Stop-Service $TokenBrokerService

    # Remove identity and token cache
    & {
        'Microsoft\OneAuth\accounts'

        # IdentityCache
        'Microsoft\IdentityCache'

        # TokenBroker Cache
        'Microsoft\TokenBroker\Cache'

        # Accounts
        'Packages\Microsoft.AAD.BrokerPlugin_cw5n1h2txyewy\ac\TokenBroker\Accounts'
        'Packages\Microsoft.Windows.CloudExperienceHost_cw5n1h2txyewy\AC\TokenBroker\Accounts'
    } `
    | Join-Path $localAppData -ChildPath { $_ } `
    | Where-Object { Test-Path $_ } `
    | Remove-Item -Recurse -Force

    # Restart TokenBroker service
    Write-Verbose "Restarting TokenBroker service"
    Set-Service $TokenBrokerService -StartupType Manual
    Start-Service $TokenBrokerService
}

function Start-LdapTrace {
    [CmdletBinding(PositionalBinding = $false)]
    param(
        [Parameter(Mandatory = $true, Position = 0, HelpMessage = "Directory for output file")]
        [string]$Path,
        [Parameter(Mandatory = $true, HelpMessage = "Process name to trace. e.g. Outlook.exe")]
        [Alias('TargetProcess')]
        [string]$TargetExecutable,
        [string]$FileName = "LDAP_$(Get-DateTimeString).etl",
        [string]$SessionName = 'LdapTrace',
        [ValidateSet('NewFile', 'Circular')]
        [string]$LogFileMode = 'NewFile',
        [ValidateRange(1, [int]::MaxValue)]
        [int]$MaxFileSizeMB = 256
    )

    if (-not (Test-Path $Path)) {
        $null = New-Item $Path -ItemType directory -ErrorAction Stop
    }

    $Path = Convert-Path -LiteralPath $Path

    # Process name must contain the extension such as "Outlook.exe", instead of "Outlook"
    $TargetExecutable = [IO.Path]::ChangeExtension($TargetExecutable, 'exe')

    # Create a registry key under HKLM\SYSTEM\CurrentControlSet\Services\ldap\tracing
    $keypath = "Registry::HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\ldap\tracing"

    if (-not (Test-Path $keypath)) {
        # Create "tracing" key
        $err = $($null = New-Item (Split-Path $keypath) -Name 'tracing') 2>&1

        if ($err) {
            Write-Error "Failed to create $keypath. Make sure to run as an administrator"
            return
        }
    }

    # Create a key under HKLM\SYSTEM\CurrentControlSet\Services\ldap\tracing
    $err = $($null = New-Item $keypath -Name $TargetExecutable) 2>&1

    if ($err) {
        Write-Error "Failed to create a key under $keypath. Make sure to run as an administrator"
        return
    }

    # Configure ETW session parameters
    switch ($LogFileMode) {
        'NewFile' {
            $mode = @([Win32.Logman+Mode]::EVENT_TRACE_USE_GLOBAL_SEQUENCE, [Win32.Logman+Mode]::EVENT_TRACE_FILE_MODE_NEWFILE) -join ','

            # In order to use newfile, file name must contain "%d"
            if ($FileName -notlike "*%d*") {
                $FileName = [System.IO.Path]::GetFileNameWithoutExtension($FileName) + "_%d.etl"
            }
            break
        }

        'Circular' {
            $mode = @([Win32.Logman+Mode]::EVENT_TRACE_USE_GLOBAL_SEQUENCE, [Win32.Logman+Mode]::EVENT_TRACE_FILE_MODE_CIRCULAR) -join ','

            if (-not $PSBoundParameters.ContainsKey('MaxFileSizeMB')) {
                $MaxFileSizeMB = 2048
            }
            break
        }
    }

    $traceFile = Join-Path $Path -ChildPath $FileName

    # Start ETW session
    Write-Log "Starting a LDAP trace for $TargetExecutable"
    $err = $($stdout = Invoke-Command {
            $ErrorActionPreference = 'Continue'
            & logman.exe create trace $SessionName -ow -o $traceFile -p Microsoft-Windows-LDAP-Client 0x1a59afa3 0xff -bs 1024 -mode $mode -max $MaxFileSizeMB -ets
        }) 2>&1

    if ($err -or $LASTEXITCODE -ne 0) {
        Write-Error "Failed to start LDAP trace. exit code:$LASTEXITCODE; stdout:`"$stdout`"; error:`"$err`""
        return
    }
}

function Stop-LdapTrace {
    [CmdletBinding()]
    param(
        $SessionName = 'LdapTrace',
        [Alias('TargetProcess')]
        $TargetExecutable
    )

    $session = Get-EtwSession | Where-Object { $_.SessionName -eq $SessionName }

    if ($session) {
        Write-Log "Stopping $SessionName"
        $null = Stop-EtwSession $SessionName
    }
    else {
        Write-Error "Cannot find an ETW session named `"$SessionName`""
    }

    $tracingKey = "Registry::HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\ldap\tracing\"

    if ($TargetExecutable) {
        # Process name must contain the extension such as "outlook.exe", instead of "outlook"
        $TargetExecutable = [IO.Path]::ChangeExtension($TargetExecutable, 'exe')

        $targetPath = Join-Path $tracingKey -ChildPath $TargetExecutable
    }
    else {
        $tracedApps = @(Get-ChildItem $tracingKey -ErrorAction SilentlyContinue)

        if ($tracedApps.Count -eq 0) {
            return
        }
        elseif ($tracedApps.Count -gt 1) {
            Write-Error "Multiple keys are found under $tracingKey. Please specify the target executable with TargetExecutable parameter."
            return
        }

        $targetPath = $tracedApps[0].PSPath
    }

    # Remove a registry key under HKLM\SYSTEM\CurrentControlSet\Services\ldap\tracing (ignore any errors)
    Write-Log "Removing $targetPath"
    $null = Remove-Item $targetPath -ErrorAction SilentlyContinue
}

function Get-OfficeModuleInfo {
    [CmdletBinding()]
    param (
        # Filter files by their name using -match (e.g. 'outlook.exe','mso\d\d.*\.dll'). These are treated as "OR".
        [string[]]$Filters,
        [Threading.CancellationToken]$CancellationToken
    )

    # If MS Office is not installed, bail.
    $officeInfo = Get-OfficeInfo

    if (-not $officeInfo) {
        return
    }

    $officePaths = @(
        $officeInfo.InstallPath

        if ($env:CommonProgramFiles) {
            Join-Path $env:CommonProgramFiles 'microsoft shared'
        }

        if (${env:CommonProgramFiles(x86)}) {
            Join-Path ${env:CommonProgramFiles(x86)} 'microsoft shared'
        }
    ) | Select-Object -Unique

    Write-Log "officePaths are $($officePaths -join ',')"

    # Get exe and dll
    # It's slightly faster to run gci twice with -Filter than running once with -Include *exe, *.dll
    $officePaths | & {
        process {
            if ($CancellationToken.IsCancellationRequested) {
                return
            }

            Get-ChildItem -Path $_ -Filter '*.exe' -Recurse -ErrorAction SilentlyContinue
            Get-ChildItem -Path $_ -Filter '*.dll' -Recurse -ErrorAction SilentlyContinue
        }
    } | & {
        # Apply filters if any
        param ([Parameter(ValueFromPipeline)]$file)
        process {
            if ($CancellationToken.IsCancellationRequested) {
                return
            }

            if ($Filters.Count -eq 0) {
                $file
                return
            }

            foreach ($filter in $Filters) {
                if ($file.Name -match $filter) {
                    $file
                    break
                }
            }
        }
    } | & {
        param ([Parameter(ValueFromPipeline)]$file)
        process {
            if ($file.VersionInfo.FileVersionRaw) {
                $fileVersion = $file.VersionInfo.FileVersionRaw
            }
            else {
                $fileVersion = $file.VersionInfo.FileVersion
            }

            $arch = Get-ImageInfo $file.FullName -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Architecture

            [PSCustomObject]@{
                Name         = $file.Name
                FullName     = $file.FullName
                FileVersion  = $fileVersion
                Architecture = $arch
            }
        }
    }
}

function Save-OfficeModuleInfo {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        $Path,
        # filter items by their name using -match (e.g. 'outlook.exe','mso\d\d.*\.dll'). These are treated as "OR".
        [string[]]$Filters,
        # Not implemented currently.
        [Threading.CancellationToken]$CancellationToken
    )

    $null = New-Item $Path -ItemType Directory -ErrorAction SilentlyContinue

    $null = $PSBoundParameters.Remove('Path')
    Get-OfficeModuleInfo @PSBoundParameters | Export-Clixml -LiteralPath (Join-Path $Path "$($MyInvocation.MyCommand.Noun).xml")
}

<#
This is an old implementation using a PowerShell Job.
Not used currently but I'm keeping it for a reference in future development.
This uses a named kernel event object for inter PS process (i.e. Job) communication.
#>
function Start-SavingOfficeModuleInfo_PSJob {
    [CmdletBinding()]
    param(
        [parameter(Mandatory = $true)]
        $Path,
        # filter items by their name using -match (e.g. 'outlook.exe','mso\d\d.*\.dll'). These are treated as "OR".
        [string[]]$Filters
    )

    if (-not (Test-Path $Path)) {
        $null = New-Item -ItemType Directory $Path -ErrorAction Stop
    }

    $Path = Convert-Path -LiteralPath $Path

    # If MS Office is not installed, bail.
    $officeInfo = Get-OfficeInfo -ErrorAction SilentlyContinue
    if (-not $officeInfo) {
        Write-Error "MS Office is not installed"
        return
    }

    # Create a named event for inter PS process communication
    $eventName = [Guid]::NewGuid().ToString()
    $namedEvent = New-Object System.Threading.EventWaitHandle($false, [Threading.EventResetMode]::ManualReset, $eventName)

    $job =
    Start-Job -ScriptBlock {
        param (
            $Path,
            $OfficeInfo,
            $Filters,
            $OutputFileName,
            $EventName
        )

        $namedEvent = [System.Threading.EventWaitHandle]::OpenExisting($EventName)

        $officePaths = @(
            $OfficeInfo.InstallPath

            if ($env:CommonProgramFiles) {
                Join-Path $env:CommonProgramFiles 'microsoft shared'
            }

            if (${env:CommonProgramFiles(x86)}) {
                Join-Path ${env:CommonProgramFiles(x86)} 'microsoft shared'
            }
        )

        # Get exe and dll
        # It's slightly faster to run gci twice with -Filter than running once with -Include *exe, *.dll
        $items = @(
            foreach ($officePath in $officePaths) {
                Get-ChildItem -Path $officePath -Filter *.exe -Recurse -ErrorAction SilentlyContinue
                Get-ChildItem -Path $officePath -Filter *.dll -Recurse -ErrorAction SilentlyContinue
            }
        )

        # Apply filters
        if ($Filters.Count) {
            # This is for PowerShell v2. PSv2 iterates a null collection.
            $items = $items | Where-Object {
                foreach ($filter in $Filters) {
                    if ($_.Name -match $filter) {
                        $true
                        break
                    }
                }
            }
        }

        @(
            foreach ($item in $items) {
                if ($item.VersionInfo.FileVersionRaw) {
                    $fileVersion = $item.VersionInfo.FileVersionRaw
                }
                else {
                    $fileVersion = $item.VersionInfo.FileVersion
                }

                [PSCustomObject]@{
                    Name        = $item.Name
                    FullName    = $item.FullName
                    #VersionInfo = $item.VersionInfo # too much info and FileVersionRaw is harder to find
                    FileVersion = $fileVersion
                }

                #  If event is signaled, finish
                if ($namedEvent.WaitOne(0)) {
                    break
                }
            }) | Export-Clixml -Depth 4 -Path $(Join-Path $Path -ChildPath "$OutputFileName.xml") -Encoding UTF8

        $namedEvent.Close()

    } -ArgumentList $Path, $officeInfo, $Filters, 'OfficeModuleInfo', $eventName

    [PSCustomObject]@{
        Job   = $job
        Event = $namedEvent # To be closed by Stop-SavingOfficeModuleInfo_PSJob
    }

    Write-Log "Job (ID:$($job.Id)) has started. A Named Event (Handle:$($namedEvent.Handle), Name:'$eventName') is created"
}

<#
This is an old implementation using a PowerShell Job. Counterpart of Start-SavingOfficeModuleInfo_PSJob
Not used currently but I'm keeping it for a reference in future development.
#>
function Stop-SavingOfficeModuleInfo_PSJob {
    [CmdletBinding()]
    param(
        # Returned from Start-SavingOfficeModuleInfo_PSJob
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
        $JobDescriptor,

        # Number of seconds to wait for the job.
        # Default value is -1 and this will block till the job completes
        [int]$TimeoutSecond = -1
    )

    process {
        $job = $JobDescriptor.job
        $namedEvent = $JobDescriptor.Event

        # Wait for the job up to timeout
        Write-Log "Waiting for the job (ID:$($job.Id)) up to $TimeoutSecond seconds"
        if (Wait-Job -Job $job -Timeout $TimeoutSecond) {
            # https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.core/wait-job
            # > This cmdlet returns job objects that represent the completed jobs. If the wait ends because the value of the Timeout parameter is exceeded, Wait-Job does not return any objects.
            Write-Log "Job was completed"
        }
        else {
            Write-Log "Job did not complete. It will be stopped by event signal"
        }

        # Signal the event and close
        try {
            $null = $namedEvent.Set()
            $namedEvent.Close()
            Write-Log "Event (Handle:$($namedEvent.Handle)) was closed"
        }
        catch {
            Write-Error -ErrorRecord $_
        }

        # Let the job finish
        $null = Wait-Job -Job $job
        Stop-Job -Job $job
        # Receive-Job -Job $job
        Remove-Job -Job $job
        Write-Log "Job (ID:$($job.Id)) was removed"
    }
}

function Start-CapiTrace {
    [CmdletBinding(PositionalBinding = $false)]
    param(
        [Parameter(Mandatory = $true, Position = 0)]
        [string]$Path,
        [string]$FileName = "CAPI_$(Get-DateTimeString).etl",
        [string]$SessionName = 'CapiTrace',
        [ValidateSet('NewFile', 'Circular')]
        [string]$LogFileMode = 'NewFile',
        [ValidateRange(1, [int]::MaxValue)]
        [int]$MaxFileSizeMB = 256
    )

    if (-not (Test-Path $Path)) {
        $null = New-Item $Path -ItemType directory -ErrorAction Stop
    }
    $Path = Convert-Path -LiteralPath $Path

    switch ($LogFileMode) {
        'NewFile' {
            $mode = @([Win32.Logman+Mode]::EVENT_TRACE_USE_GLOBAL_SEQUENCE, [Win32.Logman+Mode]::EVENT_TRACE_FILE_MODE_NEWFILE) -join ','

            # In order to use newfile, file name must contain "%d"
            if ($FileName -notlike "*%d*") {
                $FileName = [System.IO.Path]::GetFileNameWithoutExtension($FileName) + "_%d.etl"
            }
            break
        }

        'Circular' {
            $mode = @([Win32.Logman+Mode]::EVENT_TRACE_USE_GLOBAL_SEQUENCE, [Win32.Logman+Mode]::EVENT_TRACE_FILE_MODE_CIRCULAR) -join ','

            if (-not $PSBoundParameters.ContainsKey('MaxFileSizeMB')) {
                $MaxFileSizeMB = 2048
            }
            break
        }
    }

    $traceFile = Join-Path $Path -ChildPath $FileName

    Write-Log "Starting a CAPI trace"
    $logmanResult = & logman.exe create trace $SessionName -ow -o $traceFile -p "Security: SChannel" 0xffffffffffffffff 0xff -bs 1024 -mode $mode -max $MaxFileSizeMB -ets

    if ($LASTEXITCODE -ne 0) {
        Write-Error "logman failed. exit code:$LASTEXITCODE; stdout:$logmanResult"
        return
    }

    # Note:Depending on the OS version, not all providers are available.
    $logmanResult = & logman.exe update trace $SessionName -p "Schannel" 0xffffffffffffffff 0xff -ets
    $logmanResult = & logman.exe update trace $SessionName -p "{44492B72-A8E2-4F20-B0AE-F1D437657C92}" 0xffffffffffffffff 0xff -ets
    $logmanResult = & logman.exe update trace $SessionName -p "Microsoft-Windows-Schannel-Events" 0xffffffffffffffff 0xff -ets
}

function Stop-CapiTrace {
    [CmdletBinding()]
    param(
        $SessionName = 'CapiTrace'
    )

    Write-Log "Stopping $SessionName"
    $null = Stop-EtwSession $SessionName
}

function Start-FiddlerCap {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        $Path,
        # Do not start FiddlerCap.exe, but ensure it's avaiable (download it if necessary).
        [Switch]$CheckAvailabilityOnly
    )

    if (-not (Test-Path $Path -ErrorAction Stop)) {
        $null = New-Item -ItemType Directory $Path -ErrorAction Stop
    }

    $Path = Convert-Path -LiteralPath $Path
    $fiddlerPath = Join-Path $Path -ChildPath 'FiddlerCap'
    $fiddlerExe = Join-Path $fiddlerPath -ChildPath 'FiddlerCap.exe'

    #  FiddlerCap is not available.
    if (-not (Test-Path $fiddlerExe)) {
        $fiddlerCapUrl = 'https://telerik-fiddler.s3.amazonaws.com/fiddler/FiddlerCapSetup.exe'
        $fiddlerSetupFile = Join-Path $Path -ChildPath 'FiddlerCapSetup.exe'

        # Check if FiddlerCapSetup.exe is already available locally; Otherwise download the setup file and extract it.
        if (-not (Test-Path $fiddlerSetupFile)) {
            # If it's not connected to internet, bail.
            $connectivity = Get-NLMConnectivity
            if (-not $connectivity.IsConnectedToInternet) {
                Write-Error "It seems there is no connectivity to Internet. Please download FiddlerCapSetup.exe from `"$fiddlerCapUrl`" and place it `"$Path`". Then run again"
                return
            }

            Write-Log "Downloading FiddlerCapSetup.exe"
            $webClient = $null
            try {
                $webClient = New-Object System.Net.WebClient
                $webClient.UseDefaultCredentials = $true
                Write-Progress -Activity "Downloading FiddlerCap" -Status "Please wait" -PercentComplete -1
                $webClient.DownloadFile($fiddlerCapUrl, $fiddlerSetupFile)
            }
            catch {
                Write-Error -Message "Failed to download FiddlerCapSetup from $fiddlerCapUrl. $_" -Exception $_.Exception
                return
            }
            finally {
                Write-Progress -Activity "Downloading FiddlerCap" -Status "Done" -Completed

                if ($webClient) {
                    $webClient.Dispose()
                }
            }
        }

        # Silently extract. Path must be absolute.
        $process = $null
        try {
            Write-Log "Extracting from FiddlerCapSetup"
            Write-Progress -Activity "Extracting from FiddlerCapSetup" -Status "This may take a while. Please wait" -PercentComplete -1

            Unblock-File $fiddlerSetupFile -ErrorAction SilentlyContinue

            # To redirect & capture error even when this cmdlet is called with ErrorAction:SilentlyContinue, need "Continue" error action.
            # Usually you can simply specify ErrorAction:Continue to the cmdlet. However, Start-Process does not respect that. So, I need to manually set $ErrorActionPreference here.
            $err = $($process = Invoke-Command {
                    $ErrorActionPreference = 'Continue'
                    # Do not double-quote $fiddlerPath here like /D=`"$fiddlerPath`". FiddlerSetupCap.exe doesn't like it for some reason. It's ok to have spaces in the path.
                    Start-Process $fiddlerSetupFile -ArgumentList "/S /D=$fiddlerPath" -Wait -PassThru
                }) 2>&1

            if ($process.ExitCode -ne 0) {
                Write-Error "Failed to extract $fiddlerExe. $(if ($process.ExitCode) {"exit code = $($process.ExitCode)"}) $err"
                return
            }
        }
        finally {
            Write-Progress -Activity "Extracting from FiddlerCapSetup" -Status "Done" -Completed
            if ($process) {
                $process.Dispose()
            }
        }
    }

    if ($CheckAvailabilityOnly) {
        [PSCustomObject]@{ FiddlerPath = $fiddlerExe }
        return
    }

    # Start FiddlerCap.exe
    $process = $null
    try {
        Write-Log "Starting FiddlerCap"
        $err = $($process = Invoke-Command {
                $ErrorActionPreference = 'Continue'
                try {
                    Start-Process $fiddlerExe -PassThru
                }
                catch {
                    Write-Error -ErrorRecord $_
                }
            }) 2>&1

        if (-not $process -or $process.HasExited) {
            Write-Error "FiddlerCap failed to start or prematurely exited. $(if ($null -ne $process.ExitCode) {"exit code = $($process.ExitCode)"}) $err"
            return
        }
    }
    finally {
        if ($process) {
            $process.Dispose()
        }
    }

    [PSCustomObject]@{
        FiddlerPath = $fiddlerPath
    }
}

function Start-FiddlerEverywhereReporter {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        $Path,
        # Do not start FiddlerCap.exe, but ensure it's avaiable (download it if necessary).
        [Switch]$CheckAvailabilityOnly
    )

    if (-not (Test-Path $Path -ErrorAction Stop)) {
        $null = New-Item -ItemType Directory $Path -ErrorAction Stop
    }

    $Path = Convert-Path -LiteralPath $Path
    $fiddlerName = 'Fiddler Everywhere Reporter'

    if (-not (Test-Path $Path)) {
        $null = New-Item -ItemType Directory $Path -ErrorAction Stop
    }

    # Look for existing one.
    $fiddlerExe = Get-ChildItem -Path $Path -Filter "$fiddlerName*.exe" -ErrorAction SilentlyContinue | Select-Object -First 1 | Select-Object -ExpandProperty FullName

    if ($fiddlerExe) {
        Write-Log "Skip downloading because $fiddlerExe is found"
        $version = (Get-ItemProperty $fiddlerExe).VersionInfo.ProductVersion
    }
    else {
        Write-Log "Downloading $fiddlerName"
        $url = 'https://api.getfiddler.com/reporter/win/latest'

        # If it's not connected to internet, bail.
        $connectivity = Get-NLMConnectivity

        if (-not $connectivity.IsConnectedToInternet) {
            Write-Error "It seems there is no connectivity to Internet. Please download $fiddlerName from `"$url`" and place it `"$Path`". Then run again"
            return
        }

        $tempPath = Join-Path $Path -ChildPath "$fiddlerName.exe"

        $err = $($result = Download-File -Uri $url -OutFile $tempPath -Activity "Downloading $fiddlerName") 2>&1 | Select-Object -First 1

        if ($result) {
            Unblock-File $tempPath
            $version = (Get-ItemProperty $tempPath).VersionInfo.ProductVersion
            $newName = "$fiddlerName-$version.exe"
            Rename-Item $tempPath -NewName $newName
            $fiddlerExe = Join-Path (Split-Path $tempPath) $newName

            Write-Log "Successfully downloaded $fiddlerName. It took $($result.Elapsed)"
        }
        else {
            Write-Error -Message "Failed to download $fiddlerName from $url. $err" -Exception $err.Exception
            return
        }
    }

    if ($CheckAvailabilityOnly) {
        [PSCustomObject]@{ FiddlerPath = $fiddlerExe }
        return
    }

    # Start exe file.
    $activity = "Starting $fiddlerName"
    $status = "This may take a while. Please wait"
    $process = $null

    try {
        Write-Log "Starting $fiddlerName (Version: $version)"
        Write-Progress -Activity $activity -Status $status

        $err = $($process = Invoke-Command {
                $ErrorActionPreference = 'Continue'
                try {
                    Start-Process $fiddlerExe -PassThru
                }
                catch {
                    Write-Error -ErrorRecord $_
                }
            }) 2>&1

        if (-not $process -or $process.HasExited) {
            Write-Error "$fiddlerName failed to start or prematurely exited. $(if ($null -ne $process.ExitCode) {"exit code = $($process.ExitCode)"}) $err"
            return
        }

        # Wait for the UI.
        Write-Log "Waiting for $fiddlerName's UI to show up"
        $waitCount = 0
        $waitInterval = [TimeSpan]::FromSeconds(1)

        while ($true) {
            Write-Progress -Activity $activity -Status ($status + '.' * $waitCount)

            if ($process.HasExited) {
                Write-Log "$fiddlerName (PID:$($process.Id)) has prematurely exited before UI shows up" -Category Error
                Write-Error "Something went wrong after $fiddlerName was started."
                return
            }

            $processes = @(Get-Process -Name $fiddlerName -ErrorAction SilentlyContinue)
            $foundUIProcess = $false

            foreach ($proc in $processes) {
                if ($proc.MainWindowHandle -ne [IntPtr]::Zero) {
                    $foundUIProcess = $true
                }

                $proc.Dispose()
            }

            if ($foundUIProcess) {
                break
            }

            Start-Sleep -Seconds $waitInterval.TotalSeconds
            $waitCount = ++$waitCount % 10
        }

        Write-Log "$fiddlerName is successfully started"

        [PSCustomObject]@{
            FiddlerPath = $fiddlerExe
        }
    }
    finally {
        Write-Progress -Activity $activity -Completed

        if ($process) {
            $process.Dispose()
        }
    }
}

function Start-Procmon {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        $Path,
        [Parameter(Mandatory = $true)]
        # Look for existing procmon.exe before downloading
        $ProcmonSearchPath,
        $PmlFileName = "Procmon_$(Get-DateTimeString).pml"
    )

    # Explicitly check admin rights
    if (-not (Test-RunAsAdministrator)) {
        Write-Warning "Please run as administrator"
        return
    }

    if (-not (Test-Path $Path)) {
        $null = New-Item $Path -ItemType Directory -ErrorAction Stop
    }

    $Path = Convert-Path -LiteralPath $Path
    $procmonFile = $null

    # Search procmon.exe or procmon64.exe under $ProcmonSearchPath (including subfolders).
    if ($ProcmonSearchPath -and (Test-Path $ProcmonSearchPath)) {
        $files = Get-ChildItem -Path $ProcmonSearchPath -Filter 'procmon*.exe' -Exclude 'Procmon64a.exe' -Recurse

        # For x64, prefer Procmon64.exe if available. Otherwise, just use Procmon.exe
        if ($env:PROCESSOR_ARCHITECTURE -eq 'AMD64') {
            $procmonFile = $files | Where-Object { $_.Name -eq 'Procmon64.exe' } | Select-Object -First 1 -ExpandProperty FullName
        }

        if (-not $procmonFile) {
            $procmonFile = $files | Where-Object { $_.Name -eq 'Procmon.exe' } | Select-Object -First 1 -ExpandProperty FullName
        }

        if (-not $procmonFile) {
            Write-Log "ProcmonSearchPath '$ProcmonSearchPath' is provided, but coulnd't find Procmon.exe or Procmon64.exe"
        }
    }

    $procmonZipDownloaded = $false

    if ($procmonFile -and (Test-Path $procmonFile)) {
        Write-Log "$procmonFile is found. Skip searching & downloading ProcessMonitor.zip"
    }
    else {
        $procmonDownloadUrl = 'https://download.sysinternals.com/files/ProcessMonitor.zip'
        $procmonFolderPath = Join-Path $ProcmonSearchPath -ChildPath 'ProcessMonitor'
        $procmonZipFile = Join-Path $ProcmonSearchPath -ChildPath 'ProcessMonitor.zip'

        if ($env:PROCESSOR_ARCHITECTURE -eq 'AMD64') {
            $procmonFile = Join-Path $procmonFolderPath -ChildPath 'Procmon64.exe'
        }
        else {
            $procmonFile = Join-Path $procmonFolderPath -ChildPath 'Procmon.exe'
        }

        if (-not (Test-Path $procmonFolderPath)) {
            $null = New-Item $procmonFolderPath -ItemType Directory -ErrorAction Stop
        }

        if (Test-Path $procmonZipFile) {
            Write-Log "$procmonZipFile is found. Skip downloading"
        }
        else {
            # If 'ProcessMonitor.zip' isn't there, download it.
            $err = $($null = Download-File -Uri $procmonDownloadUrl -OutFile $procmonZipFile -Activity "Downloading Process Monitor") 2>&1 | Select-Object -First 1

            if (-not (Test-Path $procmonZipFile)) {
                Write-Error -ErrorRecord $err
                return
            }
        }

        # Unzip ProcessMonitor.zip
        try {
            Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction Stop
            $NETFileSystemAvailable = $true
        }
        catch {
            Write-Log "System.IO.Compression.FileSystem isn't found. Using alternate method"
        }

        if ($NETFileSystemAvailable) {
            # .NET 4 or later
            [System.IO.Compression.ZipFile]::ExtractToDirectory($procmonZipFile, $procmonFolderPath)
        }
        else {
            # Use Shell.Application COM
            # see https://docs.microsoft.com/en-us/previous-versions/windows/desktop/sidebar/system-shell-folder-copyhere
            $shell = New-Object -ComObject Shell.Application
            $shell.NameSpace($procmonFolderPath).CopyHere($shell.NameSpace($procmonZipFile).Items(), 4)
        }
    }

    if (-not ($procmonFile -and (Test-Path $procmonFile))) {
        Write-Error "Failed to find $procmonFile"
        return
    }

    if (-not $PmlFileName.EndsWith('.pml')) {
        $PmlFileName = "$PmlFileName.pml"
    }

    $pmlFile = Join-Path $Path -ChildPath $PmlFileName

    # Start procmon.exe or procmon64.exe depending on the native arch.
    Write-Log "Starting procmon"
    $process = $null
    $err = $($process = Invoke-Command {
            $ErrorActionPreference = 'Continue'
            try {
                Start-Process $procmonFile -ArgumentList "/AcceptEula /Minimized /Quiet /NoFilter /BackingFile `"$pmlFile`"" -PassThru
            }
            catch {
                Write-Error -ErrorRecord $_
            }
        }) 2>&1

    $processId = $process.Id

    try {
        if (-not $process -or $process.HasExited) {
            Write-Error "procmon failed to start or prematurely exited. $(if ($process.ExitCode) {"exit code = $($process.ExitCode)"}) $err"
            return
        }
    }
    finally {
        if ($process) {
            $process.Dispose()
        }
    }

    Write-Log "Procmon successfully started"
    [PSCustomObject]@{
        ProcmonPath          = $procmonFile
        ProcmonProcessId     = $processId
        PmlFile              = $pmlFile
        ProcmonZipDownloaded = $procmonZipDownloaded
        ProcmonFolderPath    = $procmonFolderPath
    }
}

function Stop-Procmon {
    [CmdletBinding()]
    param()

    $process = @(Get-Process -Name Procmon*)

    if ($process.Count -eq 0) {
        Write-Error "Cannot find procmon or procmon64"
        return
    }

    $procmonFile = $process[0].Path

    foreach ($p in $process) {
        $p.Dispose()
    }

    # Stop procmon
    Write-Log "Stopping procmon"

    $process = $null

    try {
        $err = $($process = Invoke-Command {
                $ErrorActionPreference = 'Continue'
                try {
                    Start-Process $procmonFile -ArgumentList "/Terminate" -Wait -PassThru
                }
                catch {
                    Write-Error -ErrorRecord $_
                }
            }) 2>&1

        if ($process.ExitCode -ne 0) {
            Write-Error "procmon failed to stop. $(if ($process.ExitCode) {"exit code = $($process.ExitCode)"}) $err"
        }
    }
    finally {
        if ($process) {
            $process.Dispose()
        }
    }
}

function Start-TcoTrace {
    [CmdletBinding()]
    param(
        [string]$User
    )

    $officeInfo = Get-OfficeInfo -ErrorAction Stop
    $majorVersion = $officeInfo.Version.Split('.')[0]

    # Create registry key & values. Ignore errors (might fail due to existing values)
    $userRegRoot = Get-UserRegistryRoot -User $User -ErrorAction Stop
    $keypath = Join-Path $userRegRoot "Software\Microsoft\Office\$majorVersion.0\Common\Debug"

    Write-Log "Using $keypath"

    if (-not (Test-Path $keypath)) {
        $null = New-Item $keypath -ErrorAction Stop
    }

    Write-Log "Starting a TCO trace by setting up $keypath"
    $null = New-ItemProperty $keypath -Name 'TCOTrace' -PropertyType DWORD -Value 7 -ErrorAction SilentlyContinue
    $null = New-ItemProperty $keypath -Name 'MsoHttpVerbose' -PropertyType DWORD -Value 1 -ErrorAction SilentlyContinue

    # If failed, throw a terminating error
    $null = Get-ItemProperty $keypath -Name 'TCOTrace' -ErrorAction Stop
    $null = Get-ItemProperty $keypath -Name 'MsoHttpVerbose' -ErrorAction Stop
}

function Stop-TcoTrace {
    [CmdletBinding()]
    param(
        [parameter(Mandatory = $true)]
        $Path,
        [string]$User
    )

    $officeInfo = Get-OfficeInfo -ErrorAction Stop
    $majorVersion = $officeInfo.Version.Split('.')[0]

    # Remove registry values
    $userRegRoot = Get-UserRegistryRoot -User $User -ErrorAction Stop
    $keypath = Join-Path $userRegRoot "Software\Microsoft\Office\$majorVersion.0\Common\Debug"

    if (-not (Test-Path $keypath)) {
        Write-Warning "$keypath does not exist"
        return
    }

    Write-Log "Stopping a TCO trace by removing TCOTrace & MsoHttpVerbose from $keypath"
    $null = Remove-ItemProperty $keypath -Name 'TCOTrace' -ErrorAction SilentlyContinue
    $null = Remove-ItemProperty $keypath -Name 'MsoHttpVerbose' -ErrorAction SilentlyContinue

    # TCO Trace logs are in %TEMP%
    Save-Item -Path $env:TEMP -Include "office.log", "*.exe.log" -Destination $Path
}

<#
.SYNOPSIS
    Returns a string representation of the given DateTime in UTC with a format somewhat close to ISO 8601. Without input, it uses the current time.
    Output string does not include commmas (':') in the time portion because Windows does not allow it in file names.
.Notes
    I am aware that Get-Date has FileDateTime & FileDateTimeUniversal with -Format. However I don't think they are easy to read (they use "yyyyMMddTHHmmssffff").
    Consolidate to this function to avoid inconsistency in the format.
#>
function Get-DateTimeString {
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [DateTime]$DateTime
    )

    if (-not $PSBoundParameters.ContainsKey('DateTime')) {
        $DateTime = [DateTime]::UtcNow
    }

    $DateTime.ToUniversalTime().ToString('yyyy-MM-ddTHHmmssZ')
}

function Get-ConnTimeout {
    [CmdletBinding()]
    param (
        $User
    )

    $userRegRoot = Get-UserRegistryRoot -User $User

    if (-not $userRegRoot) {
        return
    }

    $path = Join-Path $userRegRoot 'Software\Microsoft\Exchange\'
    $name = 'ConnTimeout'

    $prop = Get-ItemProperty $path -Name $name -ErrorAction SilentlyContinue

    [PSCustomObject]@{
        ConnTimeout = if ($prop) { [TimeSpan]::FromMilliseconds($prop.ConnTimeout) } else { $null }
    }
}

function Set-ConnTimeout {
    [CmdletBinding()]
    param(
        $User,
        [Parameter(Mandatory = $true)]
        [TimeSpan]$Value
    )

    $userRegRoot = Get-UserRegistryRoot -User $User

    if (-not $userRegRoot) {
        return
    }

    $path = Join-Path $userRegRoot 'Software\Microsoft\Exchange\'
    $name = 'ConnTimeout'

    if (-not (Test-Path $path)) {
        $null = New-Item -Path $path -Force -ErrorAction Stop
    }

    $null = Set-ItemProperty -Path $path -Name $name -Value $Value.TotalMilliseconds -Type ([Microsoft.Win32.RegistryValueKind]::DWord)
    Get-ConnTimeout
}

function Remove-ConnTimeout {
    [CmdletBinding()]
    param(
        $User
    )

    $userRegRoot = Get-UserRegistryRoot -User $User

    if (-not $userRegRoot) {
        return
    }

    $path = Join-Path $userRegRoot 'Software\Microsoft\Exchange\'
    $name = 'ConnTimeout'

    Remove-ItemProperty -Path $path -Name $name -ErrorAction SilentlyContinue
}

<#
.SYNOPSIS
There are 2 modes of execution:
1. Without OnLaunch switch
    Start tttracer.exe to launch and trace the given executable
    This mode starts the executable.

2. With OnLaunch switch
    Start tttracer.exe and begin monitoring the new process instance of the given executable.
    This mode does not start the executable.
#>
function Start-TTTracer {
    [CmdletBinding()]
    param(
        # Folder to save to.
        [Parameter(Mandatory = $true)]
        $Path,
        # Executable path (e.g. C:\Windows\System32\notepad.exe)
        [Parameter(Mandatory = $true)]
        $Executable,
        [switch]$OnLaunch
    )

    # Check if tttracer.exe is available (Win10 RS5 and above should include it)
    if (-not ($tttracer = Get-Command 'tttracer.exe' -ErrorAction SilentlyContinue)) {
        Write-Error "tttracer.exe is not available"
        return
    }

    # Make sure $Executable exists.
    if (-not (Test-Path $Executable)) {
        Write-Error "Cannot find $Executable"
        return
    }

    if (-not (Test-Path $Path)) {
        $null = New-Item $Path -ItemType Directory -ErrorAction Stop
    }

    $stdout = Join-Path $Path 'stdout.txt'
    $stderr = Join-Path $Path 'stderr.txt'

    if ($OnLaunch) {
        Write-Log "TTD monitoring $Executable"
        # trace file name must include a wildcard ("%") for OnLaunch recording
        $outPath = Join-Path $Path "$([IO.Path]::GetFileNameWithoutExtension($Executable))_$(Get-DateTimeString)_%.run"
        $process = Start-Process $tttracer -ArgumentList "-out `"$outPath`"", "-onLaunch `"$Executable`"", "-parent *" -PassThru -WindowStyle Hidden -RedirectStandardOutput $stdout -RedirectStandardError $stderr
    }
    else {
        Write-Log "TTD launching $Executable"
        $outPath = Join-Path $Path "$([IO.Path]::GetFileNameWithoutExtension($Executable))_$(Get-DateTimeString).run"
        $process = Start-Process $tttracer -ArgumentList "-out `"$outPath`"", "`"$Executable`"" -PassThru -WindowStyle Hidden -RedirectStandardOutput $stdout -RedirectStandardError $stderr
    }

    if (-not $process -or $process.HasExited) {
        Write-Error "tttracer.exe failed to start. ExitCode:$($process.ExitCode). $(Get-Content $stderr)"
        $process.Dispose()
        return
    }

    $targetProcess = $null

    if (-not $OnLaunch) {
        # Find out the new process instantiated by tttracer.exe. This might take a bit.
        # The new process starts as a child process of tttracer.exe.
        $targetName = [IO.Path]::GetFileNameWithoutExtension($Executable)
        $maxRetry = 3
        foreach ($i in 1..$maxRetry) {
            if ($newProcess = Get-CimInstance Win32_Process -Filter "Name='$targetName.exe' AND ParentProcessId='$($process.Id)'") {
                $targetProcess = Get-Process -Id $newProcess.ProcessId
                $newProcess.Dispose()
                break
            }

            Start-Sleep -Seconds $i
        }

        if (-not $targetProcess) {
            Write-Error "Cannot find the new instance of $targetName"
            return
        }

        Write-Log "Target process $($targetProcess.Name) (PID:$($targetProcess.Id)) has started"

        # To get ExitTime etc.
        $targetProcess.EnableRaisingEvents = $true
    }

    # Return a descriptor object with Dispose method.
    [PSCustomObject]@{
        TTTracerProcess = $process
        TargetProcess   = $targetProcess
        OutputFile      = $outPath
        OnLaunch        = $OnLaunch.IsPresent
    } | Add-Member -MemberType ScriptMethod -Name Dispose -Value {
        if ($this.TTTracerProcess) { $this.TTTracerProcess.Dispose() }
        if ($this.TargetProcess) { $this.TargetProcess.Dispose() }
    } -PassThru
}

function Stop-TTTracer {
    [CmdletBinding()]
    param(
        # The returned object of Stop-TTTracer
        [Parameter(Mandatory = $true)]
        $Descriptor,
        [switch]$AutoRemove
    )

    $tttracerProcess = $Descriptor.TTTracerProcess
    $targetProcess = $Descriptor.TargetProcess # could be null
    $onLaunch = $Descriptor.OnLaunch

    if (-not ($tttracer = Get-Command 'tttracer.exe' -ErrorAction SilentlyContinue)) {
        Write-Error "tttracer.exe is not available"
        return
    }

    if (-not ($tttracerProcess.ID)) {
        Write-Error "Invalid input. tttracer PID:$($tttracerProcess.ID), target process PID:$($targetProcess.ID)"
        return
    }

    $stopTarget = 'all'
    if (Get-Process -Id $targetProcess.Id -ErrorAction SilentlyContinue) {
        $stopTarget = $targetProcess.ID
    }
    else {
        Write-Log "Target process $($targetProcess.Name) (PID:$($targetProcess.Id)) does not exist; ExitCode:$($targetProcess.ExitCode), ExitTime:$(if ($targetProcess.ExitTime) {$targetProcess.ExitTime.ToString('o')}), ElapsedTime:$($targetProcess.ExitTime - $targetProcess.StartTime)"
    }

    $exitCode = 0
    [string[]]$message = $null
    $maxRetry = 3
    for ($i = 0; $i -le $maxRetry; $i++) {
        Write-Log "tttracer -stop $stopTarget $(if ($i) {"($i+1/$maxRetry retry)"})"
        $message = & $tttracer -stop $stopTarget
        $exitCode = $LASTEXITCODE

        # When timeout (0x800705b4) occurs, retry
        if ($exitCode -ne 0x800705b4) {
            break
        }
    }

    #  Non zero exitcode indicates an error.
    if ($exitCode -ne 0) {
        Write-Error $("'tttracer -stop' failed. ExitCode:0x{0:x}" -f $exitCode)
    }

    if ($onLaunch) {
        Write-Log "Killing tttracer (PID:$($tttracerProcess.Id)) running in OnLaunch mode"
        $tttracerProcess.Kill()
        $message += ";" + (& $tttracer -cleanup)
    }

    # Wait for the tttracer to exit.
    # Wait-Process writes a non-terminating error when the process has exited. Ignore this error.
    $null = $(Wait-Process -InputObject $tttracerProcess -ErrorAction SilentlyContinue) 2>&1

    [PSCustomObject]@{
        ExitCode = $exitCode  # This is the exit code of "tttracer -stop"
        Message  = $message # message of "tttracer -stop"
    }

    if ($AutoRemove) {
        $ttd.Dispose()
    }
}

function Attach-TTTracer {
    [CmdletBinding()]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseApprovedVerbs', '')]
    param(
        # Folder to save to.
        [Parameter(Mandatory = $true)]
        $Path,
        # ProcessID of the process to attach to.
        [Parameter(Mandatory = $true)]
        $ProcessID
    )

    # Check if tttracer.exe is available (Win10 RS5 and above should include it)
    if (-not ($tttracer = Get-Command 'tttracer.exe' -ErrorAction SilentlyContinue)) {
        Write-Error "tttracer.exe is not available"
        return
    }

    if ($targetProcess = Get-Process -Id $ProcessID -ErrorAction SilentlyContinue) {
        $targetName = $targetProcess.Name
    }
    else {
        Write-Error "Cannot find a process with PID $ProcessID"
        return
    }

    if (-not (Test-Path $Path)) {
        $null = New-Item $Path -ItemType Directory -ErrorAction Stop
    }

    # Form the output file name.
    $outPath = Join-Path $Path "$($targetName)_$(Get-DateTimeString).run"

    # If a folder path is used, it must not end with "\". If that's the case remove it.
    # $outPath = $Path
    # if ($outPath.EndsWith([IO.Path]::DirectorySeparatorChar)) {
    #     $outPath = $outPath.Substring(0, $outPath.Length - 1)
    # }

    $stdout = Join-Path $Path 'stdout.txt'
    $stderr = Join-Path $Path 'stderr.txt'

    $err = $($process = Invoke-Command {
            $ErrorActionPreference = 'Continue'
            Start-Process $tttracer -ArgumentList "-out `"$outPath`"", "-attach $ProcessID" -PassThru -WindowStyle Hidden -RedirectStandardOutput $stdout -RedirectStandardError $stderr
        }) 2>&1

    # Must wait for a little to see if tttracer succeeded.
    # Wait-Process writes a non-terminating error when timeout occurs. Note that timeout here is a good thing; tttracer is still running.
    $timeout = $(Wait-Process -InputObject $process -Timeout 3 <#seconds#> -ErrorAction Continue) 2>&1

    if ($timeout) {
        # To get ExitTime etc.
        $targetProcess.EnableRaisingEvents = $true

        [PSCustomObject]@{
            TTTracerProcess = $process
            TargetProcess   = $targetProcess
            OutputFile      = $outPath
            OnLaunch        = $false
        } | Add-Member -MemberType ScriptMethod -Name Dispose -Value {
            if ($this.TTTracerProcess) { $this.TTTracerProcess.Dispose() }
            if ($this.TargetProcess) { $this.TargetProcess.Dispose() }
        } -PassThru
    }
    else {
        if ($targetProcess) {
            $targetProcess.Dispose()
        }

        $stderrContent = Get-Content $stderr
        $exitCodeHex = "0x{0:x}" -f $process.ExitCode
        Write-Error "tttracer.exe failed to attach. ExitCode:$exitCodeHex; Error:$err.`n$stderrContent"
    }
}


<#
.SYNOPSIS
    Download TTD.appinstaller & its msixbundle file from https://aka.ms/ttd/download
#>
function Download-TTD {
    [CmdletBinding()]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseApprovedVerbs', '')]
    param(
        [Parameter(Mandatory)]
        # Path to download files
        [string]$Path,
        [switch]$SkipCache
    )

    $downloadUrl = 'https://aka.ms/ttd/download'

    if (-not (Test-Path $Path)) {
        $null = New-Item -Path $Path -ItemType Directory -ErrorAction Stop
    }

    $Path = Convert-Path -LiteralPath $Path

    if (-not $SkipCache) {
        # See if TTD.msixbundle exists locally.
        $msixBundlePath = Join-Path $Path 'TTD.msixbundle'

        if (Test-Path $msixBundlePath) {
            Write-Log "Found cached $msixBundlePath"

            [PSCustomObject]@{
                MsixBundlePath = $msixBundlePath
            }

            return
        }
    }

    # First, download appinstaller XML file.
    $appInstallerPath = Join-Path $Path 'TTD.appinstaller'

    if (-not (Test-Path $appInstallerPath)) {
        $err = $($null = Download-File -Uri $downloadUrl -OutFile $appInstallerPath) 2>&1 | Select-Object -First 1

        if (-not (Test-Path $appInstallerPath)) {
            Write-Error -Message "Failed to download TTD.appinstaller from $downloadUrl. $err" -Exception $err.Exception
            return
        }
    }

    # From appinstaller XML, extract MainBundle URI
    $ns = @{ ns = 'http://schemas.microsoft.com/appx/appinstaller/2018' }
    $mainBundle = Select-Xml -Path $appInstallerPath -Namespace $ns -XPath '//ns:MainBundle[@Name="Microsoft.TimeTravelDebugging"]' `
    | Select-Object -ExpandProperty Node

    if (-not $mainBundle.Uri) {
        Write-Error "Failed to find MainBundle element in TTD.appinstaller"
        return
    }

    # Download TTD.msixbundle
    $msixName = Split-Path $mainBundle.Uri -Leaf
    $msixBundlePath = Join-Path $Path $msixName

    $err = $($null = Download-File -Uri $mainBundle.Uri -OutFile $msixBundlePath) 2>&1 | Select-Object -First 1

    if (-not (Test-Path $msixBundlePath)) {
        Write-Error -Message "Failed to download $($mainBundle.Uri). $err" -Exception $err.Exception
        return
    }

    [PSCustomObject]@{
        MsixBundlePath = $msixBundlePath
    }
}

<#
.SYNOPSIS
    Instead of installing public TTD's MsixBundle, simply expand and extract the contents
#>
function Expand-TTDMsixBundle {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, ValueFromPipelineByPropertyName)]
        # Path to TTD.msixbundle file
        [string]$MsixBundlePath
    )

    if (-not (Test-Path $MsixBundlePath)) {
        Write-Error "Cannot find $MsixBundlePath"
        return
    }

    $MsixBundlePath = Convert-Path -LiteralPath $MsixBundlePath

    # Helper script block to rename & expand
    $expand = {
        param($Path, $Destination)

        # First rename to "***.zip"
        $dir = [IO.Path]::GetDirectoryName($Path)
        $originalName = [IO.Path]::GetFileName($Path)
        $newName = "$originalName.zip"
        $newFilePath = Join-Path $dir $newName

        if (-not (Test-Path $newFilePath)) {
            Rename-Item -LiteralPath $Path -NewName $newName
        }

        if (-not (Test-Path $Destination)) {
            $null = New-Item $Destination -ItemType Directory -ErrorAction Stop
        }

        try {
            Expand-Archive -LiteralPath $newFilePath -DestinationPath $Destination -Force -ErrorAction SilentlyContinue
        }
        finally {
            # Revert to the original name
            Rename-Item -LiteralPath $newFilePath -NewName $originalName
        }
    }

    # This is to suppress progress on Expand-Archive
    $savedProgressPreference = $Global:ProgressPreference
    $Global:ProgressPreference = "SilentlyContinue"

    try {
        # Expand destination
        $root = [IO.Path]::GetDirectoryName($MsixBundlePath)
        $dest = Join-Path $root 'TTD'

        # See if TTD.exe of the same architecture already exists
        $ttdPath = Get-ChildItem $dest -Filter 'TTD.exe' -File -Recurse | & {
            begin {
                # Map PROCESSOR_ARCHITECTURE to Get-ImageInfo's Architecture value (x86, x64, ARM64)
                $archMap = @{
                    'x86'   = 'x86'
                    'AMD64' = 'x64'
                    'ARM64' = 'ARM64'
                }
            }

            process {
                $image = Get-ImageInfo $_.FullName

                if ($image.Architecture -eq $archMap[$env:PROCESSOR_ARCHITECTURE]) {
                    $_.FullName
                }
            }
        } | Select-Object -First 1

        if ($ttdPath) {
            Write-Log "Skip expanding $MsixBundlePath because TTD.exe (for $env:PROCESSOR_ARCHITECTURE) already exists at $ttdPath"
        }
        else {
            $msixFileName = switch ($env:PROCESSOR_ARCHITECTURE) {
                'AMD64' { 'TTD-x64.msix'; break }
                'x86' { 'TTD-x86.msix'; break }
                'ARM64' { 'TTD-ARM64.msix'; break }
            }

            if (-not $msixFileName) {
                Write-Error "Unsupported Processor Archtecture:$($env:PROCESSOR_ARCHITECTURE)"
                return
            }

            & $expand -Path $MsixBundlePath -Destination $dest

            $msixFilePath = Join-Path $dest $msixFileName
            $dest = Join-Path $dest ([IO.Path]::GetFileNameWithoutExtension($msixFileName))

            & $expand -Path $msixFilePath -Destination $dest

            $ttdPath = Join-Path $dest 'TTD.exe'

            if (-not (Test-Path($ttdPath))) {
                Write-Error "Cannot find $ttdPath"
                return
            }
        }

        [PSCustomObject]@{
            TTDPath = $ttdPath
        }
    }
    finally {
        if ($savedProgressPreference) {
            $Global:ProgressPreference = $savedProgressPreference
        }
    }
}

function Install-TTD {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, ValueFromPipelineByPropertyName)]
        # Path to TTD.msixbundle file
        $MsixBundlePath
    )

    Write-Log "Installing TTD. Invoking 'Add-AppxPackage -Path $MsixBundlePath'"

    Invoke-Command {
        $ProgressPreference = "SilentlyContinue";
        Add-AppxPackage -Path $MsixBundlePath
    }
}

function Uninstall-TTD {
    [CmdletBinding()]
    param()

    $package = Get-AppxPackage -Name 'Microsoft.TimeTravelDebugging'

    if (-not $package) {
        Write-Error -Message "TTD is not installed. Cannot find AppxPackage 'Microsoft.TimeTravelDebugging'"
        return
    }

    Cleanup-TTD
    Write-Log "Uninstalling TTD. Invoking 'Remove-AppxPackage -Package $package'"

    Invoke-Command {
        $ProgressPreference = "SilentlyContinue";
        Remove-AppxPackage -Package $package
    }
}

function Start-TTDMonitor {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        # Path to TTD.exe
        $TTDPath,
        [Parameter(Mandatory)]
        # Output folder path
        $Path,
        # Name of executable (such as outlook.exe)
        [Parameter(Mandatory)]
        $ExecutableName,
        [Alias('CmdLineFilter')]
        [string]$CommandlineFilter,
        [string[]]$Modules,
        [switch]$ShowUI
    )

    if (-not (Test-Path $Path)) {
        $null = New-Item -Path $Path -ItemType Directory -ErrorAction Stop
    }

    $Path = Convert-Path -LiteralPath $Path

    # Make sure extension is ".exe"
    $ExecutableName = [IO.Path]::ChangeExtension($ExecutableName, 'exe')

    foreach ($module in $Modules) {
        if (-not [IO.Path]::GetExtension($module)) {
            Write-Error "Module name must have an extension. Invalid module name:`"$module`""
            return
        }
    }

    $outPath = $Path.ToString()

    if ($outPath.IndexOf(' ') -gt 0) {
        $outPath = "`"$outPath`""
    }

    $initCompleteEventName = [Guid]::NewGuid().ToString()
    $initCompleteEvent = New-Object System.Threading.EventWaitHandle -ArgumentList $false, ([System.Threading.EventResetMode]::ManualReset), $initCompleteEventName

    $ttdArgs = @(
        '-acceptEula'
        '-timestampFileName'
        '-out', $outPath
        '-monitor', $ExecutableName

        if ($CommandlineFilter) {
            # As of TTD 1.11.173.0, cmdLineFilter arg must not start with / or - (e.g., '/f' or '-cleanviews' does not work)
            if ($CommandlineFilter.StartsWith('/') -or $CommandlineFilter.StartsWith('-')) {
                $CommandlineFilter = $CommandlineFilter.Substring(1)
                Write-Log "CommandlineFilter cannot start with / or -. Modified to '$CommandlineFilter'"
            }

            '-cmdLineFilter', $CommandlineFilter
        }

        foreach ($module in $Modules) {
            '-module', $module
        }

        if (-not $ShowUI) {
            '-noUI'
        }

        '-onMonitorReadyEvent', $initCompleteEventName
    )

    $stderr = Join-Path $Path 'stderr.txt'
    Write-Log "Invoking 'TTD.exe $($ttdArgs -join ' ')'"

    $process = Start-Process $TTDPath -ArgumentList $ttdArgs -WindowStyle Hidden -RedirectStandardError $stderr -PassThru

    # Make sure TTD.exe started successfully
    $waitInterval = [TimeSpan]::FromSeconds(1)

    while ($true) {
        if (-not $process -or $process.HasExited) {
            # Something went wrong. Clean up and bail.
            $initCompleteEvent.Dispose()

            if ($process) {
                $process.Dispose()
            }

            if (Test-Path $stderr) {
                $errText = [IO.File]::ReadAllText($stderr)
            }

            Write-Error "TTD.exe failed to start. $errText"
            return
        }

        if ($initCompleteEvent.WaitOne($waitInterval)) {
            $initCompleteEvent.Dispose()
            break
        }
    }

    Write-Log "TTD.exe (PID:$($process.Id)) has successfully started"
    $null = $process.Handle

    [PSCustomObject]@{
        TTDPath       = $TTDPath # Need to remember this for Stop-TTDMonitor because if the process dies, System.Diagnostics.Process loses the path
        TTDProcess    = $process
        StandardError = $stderr
    }
}

function Stop-TTDMonitor {
    [CmdletBinding()]
    param(
        [Parameter(ValueFromPipeline)]
        # Output of Start-TTDMonitor
        $InputObject
    )

    if ($InputObject) {
        $ttdPath = $InputObject.TTDPath
        $ttdProcess = $InputObject.TTDProcess
    }
    else {
        # If InputObject is not given, find "TTD.exe -monitor" instance
        $ttdProcess = Get-CimInstance Win32_Process -Filter 'Name = "TTD.exe"' | & {
            process {
                if ($_.CommandLine -match '-monitor') {
                    Get-Process -Id $_.ProcessId
                }

                $_.Dispose()
            }
        } | Select-Object -First 1

        if ($ttdProcess) {
            $ttdPath = $ttdProcess.Path
        }
        else {
            # Without an instance of ttd.exe, "ttd.exe -stop all" fails with timeout anyway. So there's nothing I can do here.
            Write-Error "Cannot find TTD.exe -monitor instance"
            return
        }
    }

    # Stop monitoring
    # It's possible that this instance has died already.
    if ($ttdProcess -and -not $ttdProcess.HasExited) {
        # Stop current tracing, if any
        Write-Log "Invoking 'TTD.exe -stop all'"
        $null = & $ttdPath -stop all

        Write-Log "Stopping TTD.exe (PID:$($ttdProcess.Id))"
        Stop-Process -InputObject $ttdProcess -ErrorAction Stop

        $ttdProcess.WaitForExit()
        $ttdProcess.Dispose()
    }
    else {
        Write-Error "TTD -monitor instance (PID:$($ttdProcess.Id)) has already died (ExitTime:$($ttdProcess.ExitTime), ExitCode:$($ttdProcess.ExitCode))"
    }

    # If StandardError file exists but it's empty, remove it.
    if ($InputObject.StandardError -and (Test-Path $InputObject.StandardError) -and -not (Get-Content $InputObject.StandardError)) {
        Remove-Item $InputObject.StandardError 2>&1 | Write-Log -Category Warning
    }
}

function Cleanup-TTD {
    [CmdletBinding()]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseApprovedVerbs', '')]
    param(
        [Parameter(Mandatory, ValueFromPipelineByPropertyName)]
        $TTDPath
    )

    # Clean up ProcLaunchMon
    Write-Log "Invoking 'TTD.exe -cleanup'"
    $null = & $TTDPath -cleanup

    if ($LASTEXITCODE -ne 0) {
        Write-Error "TTD.exe -cleanup failed with $LASTEXITCODE"
    }
}

function Attach-TTD {
    [CmdletBinding()]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseApprovedVerbs', '')]
    param(
        [Parameter(Mandatory)]
        # Path to TTD.exe
        $TTDPath,
        [Parameter(Mandatory)]
        # Output folder path
        $Path,
        [Parameter(Mandatory)]
        # Target Process ID
        [ValidateRange(4, [int]::MaxValue)]
        [int]$ProcessId,
        # Module names to trace. Must have extension.
        [string[]]$Modules,
        [Switch]$ShowUI
    )

    # Validate input args
    if (0 -ne ($ProcessId % 4)) {
        Write-Error "Invalid Process ID:$ProcessId"
        return
    }

    if ($process = Get-Process -Id $ProcessId -ErrorAction SilentlyContinue) {
        $process.Dispose()
    }
    else {
        Write-Error "Cannot find a process with $ProcessId"
        return
    }

    foreach ($module in $Modules) {
        if (-not [IO.Path]::GetExtension($module)) {
            Write-Error "Module name must have an extension. Invalid module name:`"$module`""
            return
        }
    }

    # Make sure TTD.exe is available.
    if (-not (Test-Path $TTDPath -PathType Leaf)) {
        Write-Error "TTD is not available"
        return
    }

    if (-not (Test-Path $Path)) {
        $null = New-Item $Path -ItemType Directory -ErrorAction Stop
    }

    $Path = Convert-Path -LiteralPath $Path

    # If Path contains spaces, surround by double-quotes
    $outPath = $Path

    if ($outPath.IndexOf(' ') -gt 0) {
        $outPath = "`"$outPath`""
    }

    # Create a named event for onInitCompleteEvent parameter of TTD.exe
    $initCompleteEventName = [Guid]::NewGuid().ToString()
    $initCompleteEvent = New-Object System.Threading.EventWaitHandle -ArgumentList $false, ([System.Threading.EventResetMode]::ManualReset), $initCompleteEventName

    $ttdArgs = @(
        '-acceptEula'
        '-timestampFileName'
        '-out', $outPath
        '-attach', $ProcessId,
        '-onInitCompleteEvent', $initCompleteEventName

        foreach ($module in $Modules) {
            '-module', $module
        }

        if (-not $ShowUI) {
            '-noUI'
        }
    )

    $stderr = Join-Path $Path 'stderr.txt'
    Write-Log "Invoking TTD.exe $($ttdArgs -join ' ')"

    $process = Start-Process $TTDPath -ArgumentList $ttdArgs -WindowStyle Hidden -RedirectStandardError $stderr -PassThru

    $attachStart = Get-Timestamp
    $waitInterval = [TimeSpan]::FromSeconds(1)

    # Check if TTD.exe successfully attached. And if so, wait until TTD.exe signals initCompleteEvent
    while ($true) {
        if (-not $process -or $process.HasExited) {
            $initCompleteEvent.Dispose()

            if ($process) {
                $process.Dispose()
            }

            if (Test-Path $stderr) {
                $errText = [IO.File]::ReadAllText($stderr)
            }

            Write-Error "TTD.exe failed to attach to the target (PID:$ProcessId). $errText"
            return
        }

        if ($initCompleteEvent.WaitOne($waitInterval)) {
            # initCompleteEvent is signaled.
            $initCompleteEvent.Dispose()
            break
        }

        Start-Sleep -Seconds 1
    }

    $attachElapsed = Get-Elapsed $attachStart
    Write-Log "TTD (PID:$($process.Id)) successfully attached to the target (PID:$ProcessId). Attach Wait:$attachElapsed"

    [PSCustomObject]@{
        TTDPath         = $TTDPath
        TTDProcess      = $process
        TargetProcessId = $ProcessId
        IsAttached      = $true
        StandardError   = $stderr
    }
}

function Detach-TTD {
    [CmdletBinding(DefaultParameterSetName = 'InputObject')]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseApprovedVerbs', '')]
    param(
        [Parameter(ParameterSetName = 'InputObject', Mandatory, ValueFromPipeline)]
        # Output of Attach-TTD
        $InputObject
    )

    $ttdPath = $InputObject.TTDPath
    $targetPid = $InputObject.TargetProcessId

    # The target process might not exist anymore (e.g., it might have crashed)
    $targetProcess = $null

    try {
        $targetProcess = Get-Process -Id $targetPid -ErrorAction SilentlyContinue

        if (-not $targetProcess) {
            Write-Log "Target Process (PID:$targetPid) does not exist anymore"
            return
        }

        Write-Log "Detaching TTD. Invoking 'TTD.exe -stop $targetPid'"
        $null = & $ttdPath -stop $targetPid

        if ($LASTEXITCODE -ne 0) {
            Write-Error "Failed to detach TTD from $targetPid"
        }
    }
    finally {
        if ($targetProcess) {
            $targetProcess.Dispose()
        }

        # Wait for TTD.exe to finish.
        if ($InputObject.TTDProcess) {
            $InputObject.TTDProcess.WaitForExit()
            $InputObject.TTDProcess.Dispose()
        }

        # Remove StandardError file if empty
        if ($InputObject.StandardError -and (Test-Path $InputObject.StandardError) -and -not (Get-Content $InputObject.StandardError)) {
            Remove-Item $InputObject.StandardError 2>&1 | Write-Log -Category Warning
        }
    }
}

function Get-OfficeInfo {
    [CmdletBinding()]
    param(
        [switch]$IgnoreCache
    )

    # Use the cache if it's available
    if ($Script:OfficeInfoCache -and -not $IgnoreCache.IsPresent) {
        Write-Log "Returning a cache"
        return $Script:OfficeInfoCache
    }

    $officeInstallations = @(
        $hklm = $null
        try {
            if ('Microsoft.Win32.RegistryView' -as [type]) {
                Write-Log "Using OpenBaseKey with [Microsoft.Win32.RegistryView]::Registry64"
                $hklm = [Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::LocalMachine, [Microsoft.Win32.RegistryView]::Registry64);
            }
            elseif (-not $env:PROCESSOR_ARCHITEW6432) {
                # RegistryView is not available, but it's OK because no WOW64.
                Write-Log "Using OpenRemoteBaseKey"
                $hklm = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey([Microsoft.Win32.RegistryHive]::LocalMachine, [string]::Empty);
            }
            else {
                # This is the case where registry rediction takes place (32bit PowerShell on 64bit OS). Bail.
                Write-Error "32bit PowerShell 2.0 is running on 64bit OS. Please run 64bit PowerShell"
                return
            }

            $keysToSearch = @(
                'SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall',
                # 32bit MSI is under Wow6432Node.
                'SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall'
            )

            foreach ($key in $keysToSearch) {
                $uninstallKey = $hklm.OpenSubKey($key)

                if ($null -eq $uninstallKey) {
                    continue
                }

                foreach ($subKeyName in $uninstallKey.GetSubKeyNames()) {
                    if ($null -eq $subKeyName) {
                        continue
                    }

                    $subKey = $uninstallKey.OpenSubKey($subKeyName)
                    $displayName = $subKey.GetValue('DisplayName')
                    $displayIcon = $subKey.GetValue('DisplayIcon')
                    $modifyPath = $subKey.GetValue('ModifyPath')

                    if (($displayName -like "Microsoft Office*" -or $displayName -like "Microsoft 365 Apps*") -and $displayIcon -and $modifyPath -notlike "*MUI*") {
                        [PSCustomObject]@{
                            DisplayName = $displayName
                            Version     = $subKey.GetValue('DisplayVersion')
                            Location    = $subKey.GetValue('InstallLocation')
                            ModifyPath  = $modifyPath
                            DisplayIcon = $displayIcon
                        }
                    }

                    $subKey.Close()
                }

                $uninstallKey.Close()
            }
        }
        finally {
            if ($hklm) {
                $hklm.Close()
            }
        }
    )

    $displayName = $version = $installPath = $null

    if ($officeInstallations.Count -gt 0) {
        # There might be more than one version of Office installed.
        # Use the latest
        $latestOffice = $officeInstallations | Sort-Object -Property { [System.Version]$_.Version } -Descending | Select-Object -First 1
        $displayName = $latestOffice.DisplayName
        $version = $latestOffice.Version
        $installPath = $latestOffice.Location
    }
    else {
        Write-Log "Cannot find the Office installation from HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall. Fall back to HKLM:\SOFTWARE\Microsoft\Office"
        $keys = @(Get-ChildItem 'Registry::HKLM\SOFTWARE\Microsoft\Office\' -ErrorAction SilentlyContinue | Where-Object { [RegEx]::IsMatch($_.PSChildName, '\d\d\.0') -or $_.PSChildName -eq 'ClickToRun' })

        # If 'ClickToRun' exists, use its "InstallPath" & "VersionToReport".
        $clickToRun = $keys | Where-Object { $_.PSChildName -eq 'ClickToRun' }

        if ($clickToRun) {
            $installPath = Get-ItemProperty $clickToRun.PSPath | Select-Object -ExpandProperty 'InstallPath'
            $version = Get-ItemProperty (Join-Path $clickToRun.PSPath 'Configuration') | Select-Object -ExpandProperty 'VersionToReport'
        }
        else {
            # Otherwise, check "Common\InstallRoot" key's "Path"
            foreach ($key in ($keys | Sort-Object -Property PSChildName -Descending)) {
                $installPath = Get-ItemProperty (Join-Path $key.PSPath 'Common\InstallRoot') -ErrorAction SilentlyContinue | Select-Object -ExpandProperty 'Path'
                if ($installPath) {
                    $version = $key.PSChildName
                    break
                }
            }
        }
    }

    if (-not $installPath) {
        Write-Error "Microsoft Office is not installed"
        return
    }

    $outlookReg = Get-ItemProperty 'Registry::HKLM\SOFTWARE\Clients\Mail\Microsoft Outlook' -ErrorAction SilentlyContinue

    if ($outlookReg) {
        $mapiDll = Get-ItemProperty $outlookReg.DLLPathEx -ErrorAction SilentlyContinue
        $arch = Get-ImageInfo -Path $mapiDll | Select-Object -ExpandProperty Architecture
    }

    $Script:OfficeInfoCache = [PSCustomObject]@{
        DisplayName     = $displayName
        Version         = $version
        InstallPath     = $installPath
        MapiDllFileInfo = $mapiDll
        Architecture    = $arch
    }

    $Script:OfficeInfoCache
}

function Add-WerDumpKey {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        $Path, # Folder to save dump files
        [Parameter(ValueFromPipeline = $true)]
        [string]$TargetProcess # Target Process (e.g. Outlook.exe)
    )

    begin {
        # Need admin rights to modify HKLM registry values.
        if (-not (Test-RunAsAdministrator)) {
            Write-Error "Please run as administrator"
            return
        }

        if (-not (Test-Path $Path)) {
            $null = New-Item $Path -ItemType Directory -ErrorAction Stop
        }

        $Path = Convert-Path -LiteralPath $Path -ErrorAction Stop

        # Create a key 'LocalDumps' under HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\Windows Error Reporting\LocalDumps, if it doesn't exist
        $werKey = 'HKLM:\SOFTWARE\Microsoft\Windows\Windows Error Reporting'

        if (-not (Test-Path (Join-Path $werKey 'LocalDumps'))) {
            $null = New-Item $werKey -Name 'LocalDumps' -ErrorAction Stop
        }

        $localDumpsKey = Join-Path $werKey 'LocalDumps'

        # Remove global DumpCount if exits (If this is set to 0, no dumps generated)
        $DumpCount = 'DumpCount'
        $globalDumpCount = Get-ItemProperty $localDumpsKey -Name $DumpCount -ErrorAction SilentlyContinue | Select-Object -ExpandProperty $DumpCount

        if ($null -ne $globalDumpCount) {
            Write-Log "Removing global $DumpCount registry value (Value:$globalDumpCount)"
            Remove-ItemProperty $localDumpsKey -Name $DumpCount -ErrorAction SilentlyContinue
        }

        $beginBlockComplete = $true
    }

    process {
        # Bail if begin block failed.
        if (-not $Local:beginBlockComplete) {
            return
        }

        if ($TargetProcess) {
            $TargetProcess = [IO.Path]::ChangeExtension($TargetProcess, 'exe')

            # Create a ProcessName key under LocalDumps, if it doesn't exist.
            $targetKey = Join-Path $localDumpsKey $TargetProcess

            if (-not (Test-Path $targetKey)) {
                $null = New-Item $localDumpsKey -Name $TargetProcess -ErrorAction Stop
            }
        }
        else {
            $targetKey = $localDumpsKey
        }

        # Create "CustomDumpFlags", "DumpType", and "DumpFolder" values
        Write-Log "Setting up $targetKey with CustomDumpFlags:0x61826, DumpType:0, DumpFolder:$Path"
        # -Force will overwrite existing value
        # 0x61826 = MiniDumpWithTokenInformation | MiniDumpIgnoreInaccessibleMemory | MiniDumpWithThreadInfo (0x1000) | MiniDumpWithFullMemoryInfo (0x800) |MiniDumpWithUnloadedModules (0x20) | MiniDumpWithHandleData (4)| MiniDumpWithFullMemory (2)
        $null = New-ItemProperty $targetKey -Name 'CustomDumpFlags' -Value 0x00061826 -Force -ErrorAction Stop
        $null = New-ItemProperty $targetKey -Name 'DumpType' -Value 0 -PropertyType DWORD -Force -ErrorAction Stop
        $null = New-ItemProperty $targetKey -Name 'DumpFolder' -Value $Path -PropertyType String -Force -ErrorAction Stop

        $processBlockComplete = $true
    }

    end {
        # If none of process block gets completed, there is no need to configure the rest.
        if (-not $Local:processBlockComplete) {
            return
        }

        # Rename DW Installed keys to "_Installed" in order to disable it temporarily
        foreach ($_ in @(
                # For C2R
                'HKLM:\Software\Microsoft\Office\ClickToRun\REGISTRY\MACHINE\Software\Microsoft\PCHealth\ErrorReporting\DW\Installed'
                'HKLM:\Software\Microsoft\Office\ClickToRun\REGISTRY\MACHINE\Software\Wow6432Node\Microsoft\PCHealth\ErrorReporting\DW\Installed'

                # For MSI
                'HKLM:\Software\Microsoft\PCHealth\ErrorReporting\DW\Installed'
                'HKLM:\Software\Wow6432Node\Microsoft\PCHealth\ErrorReporting\DW\Installed'
            )) {
            if (Test-Path $_) {
                Write-Log "Temporarily renaming $_ to `"_Installed`""
                Rename-Item $_ -NewName '_Installed'
            }
        }

        Write-Log "Temporarily disabling dwwin"
        Disable-DWWin 2>&1 | Write-Log

        Write-Log "Temporarily disabling AeDebug Debugger if any"
        Disable-AeDebugDebugger 2>&1 | Write-Log
    }
}

function Remove-WerDumpKey {
    [CmdletBinding()]
    param (
        [Parameter(ValueFromPipeline = $true)]
        [string]$TargetProcess # Target Process (e.g. Outlook.exe)
    )

    begin {
        $werKey = 'HKLM:\SOFTWARE\Microsoft\Windows\Windows Error Reporting'
        $localDumpsKey = Join-Path $werKey 'LocalDumps'

        if (-not (Test-Path $localDumpsKey)) {
            Write-Log "Cannot find $localDumpsKey"
            return
        }

        $beginBlockComplete = $true
    }

    process {
        if (-not $Local:beginBlockComplete) {
            return
        }

        if ($TargetProcess) {
            $TargetProcess = [IO.Path]::ChangeExtension($TargetProcess, 'exe')
            $targetKey = Join-Path $localDumpsKey $TargetProcess
            Write-Log "Removing $targetKey"
            Remove-Item $targetKey
        }
        else {
            Write-Log "Removing values of $localDumpsKey"
            Remove-ItemProperty $localDumpsKey -Name 'CustomDumpFlags', 'DumpType', 'DumpFolder'
        }

        if (-not (Test-Path (Join-Path $localDumpsKey '*'))) {
            Write-Log "Removing $localDumpsKey because it has no subkeys"
            Remove-Item $localDumpsKey
        }
    }

    end {
        # Rename DW "_Installed" keys back to "Installed"
        foreach ($_ in @(
                # For C2R
                'HKLM:\Software\Microsoft\Office\ClickToRun\REGISTRY\MACHINE\Software\Microsoft\PCHealth\ErrorReporting\DW\_Installed'
                'HKLM:\Software\Microsoft\Office\ClickToRun\REGISTRY\MACHINE\Software\Wow6432Node\Microsoft\PCHealth\ErrorReporting\DW\_Installed'

                # For MSI
                'HKLM:\Software\Microsoft\PCHealth\ErrorReporting\DW\_Installed'
                'HKLM:\Software\Wow6432Node\Microsoft\PCHealth\ErrorReporting\DW\_Installed'
            )) {
            if (Test-Path $_) {
                Write-Log "Renaming $_ back to `"Installed`""
                Rename-Item $_ -NewName 'Installed'
            }
        }

        Write-Log "Re-enabling dwwin"
        Enable-DWWin 2>&1 | Write-Log

        Write-Log "Re-enabling AeDebug Debugger if it was disabled previously"
        Enable-AeDebugDebugger 2>&1 | Write-Log
    }
}

<#
Prevent dwwin.exe from lauching by adding a fake Debugger key in Image File Execution Options.
#>
function Disable-DWWin {
    [CmdletBinding()]
    param()

    $IFEO = 'Registry::HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options'
    $dwwin = 'dwwin.exe'
    $imageKeyPath = Join-Path $IFEO $dwwin

    if (-not (Test-Path $imageKeyPath)) {
        $null = New-Item $IFEO -Name $dwwin
    }

    # Create "Debugger" key (If the value already exists, an error will be put in verbose stream)
    $null = New-ItemProperty $imageKeyPath -Name 'Debugger' -Value ([Guid]::NewGuid().ToString())
}

function Enable-DWWin {
    [CmdletBinding()]
    param()

    $IFEO = 'Registry::HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options'
    $dwwin = 'dwwin.exe'
    $imageKeyPath = Join-Path $IFEO $dwwin

    if (Test-Path $imageKeyPath) {
        Remove-Item $imageKeyPath
    }
}

function Disable-AeDebugDebugger {
    [CmdletBinding()]
    param()

    $AeDebug = 'Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\AeDebug'
    $AeDebugWow64 = 'Registry::HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Microsoft\Windows NT\CurrentVersion\AeDebug'
    $Debugger = 'Debugger'
    $DebuggerTempName = '39f3719a-b064-465c-87c7-ccd09ba007df'

    @($AeDebug, $AeDebugWow64) | & {
        process {
            $value = Get-ItemProperty $_ -Name $Debugger -ErrorAction SilentlyContinue | Select-Object -ExpandProperty $Debugger

            if ($value) {
                Write-Log "AeDebug Debugger is found at $_ (value:$value). Renaming Debugger to $DebuggerTempName"
                Rename-ItemProperty $_ -Name $Debugger -NewName $DebuggerTempName
            }
        }
    }
}

function Enable-AeDebugDebugger {
    [CmdletBinding()]
    param()

    $AeDebug = 'Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\AeDebug'
    $AeDebugWow64 = 'Registry::HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Microsoft\Windows NT\CurrentVersion\AeDebug'
    $Debugger = 'Debugger'
    $DebuggerTempName = '39f3719a-b064-465c-87c7-ccd09ba007df'

    @($AeDebug, $AeDebugWow64) | & {
        process {
            if (Get-ItemProperty $_ -Name $DebuggerTempName -ErrorAction SilentlyContinue) {
                Write-Log "Placeholder is found at $_. Renaming to Debugger"
                Rename-ItemProperty $_ -Name $DebuggerTempName -NewName $Debugger
            }
        }
    }
}

function Enable-PageHeap {
    [CmdletBinding()]
    [OutputType([bool])]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ProcessName
    )

    if (-not (Test-RunAsAdministrator)) {
        Write-Error "Please run as administrator"
        return
    }

    $ProcessName = [IO.Path]::ChangeExtension($ProcessName, 'exe')

    Disable-PageHeap -ProcessName $ProcessName -ErrorAction SilentlyContinue

    $IFEO = 'Registry::HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options'
    $imageKeyPath = Join-Path $IFEO $ProcessName

    if (-not (Test-Path $imageKeyPath)) {
        $null = New-Item $IFEO -Name $ProcessName
    }

    $success = $true

    foreach ($kvp in @(@{Name = 'GlobalFlag'; Value = 0x2000000 }, @{Name = 'PageHeapFlags'; Value = 3 })) {
        $null = New-ItemProperty $imageKeyPath -Name $kvp.Name -Value $kvp.Value

        # Don't use -ErrorVariable because StopUpstreamCommandsException might populate it
        # https://github.com/PowerShell/PowerShell/pull/10840
        if (-not $?) {
            $success = $false
            break
        }
    }

    if ($success) {
        Write-Log "PageHeap is enabled for $ProcessName"
        $true
    }
}

function Disable-PageHeap {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ProcessName
    )

    if (-not (Test-RunAsAdministrator)) {
        Write-Error "Please run as administrator"
        return
    }

    $ProcessName = [IO.Path]::ChangeExtension($ProcessName, 'exe')

    $IFEO = 'Registry::HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options'
    $imageKeyPath = Join-Path $IFEO $ProcessName

    if (-not (Test-Path $imageKeyPath)) {
        # There's nothing to do.
        return
    }

    $success = $true

    foreach ($name in @('GlobalFlag', 'PageHeapFlags')) {
        Remove-ItemProperty $imageKeyPath -Name $name

        if (-not $?) {
            $success = $false
        }
    }

    if ($success) {
        Write-Log "PageHeap is disabled for $ProcessName"
    }
}

function Start-WfpTrace {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        $Path,
        [Parameter(Mandatory = $true)]
        [TimeSpan]$Interval,
        [TimeSpan]$MaxDuration = [TimeSpan]::FromHours(1)  # Just for safety, make sure to stop after a certaion period
    )

    if (-not (Test-RunAsAdministrator)) {
        Write-Error "Please run as administrator"
        return
    }

    if (-not (Test-Path $Path)) {
        $null = New-Item -ItemType directory $Path -ErrorAction Stop
    }

    $Path = Convert-Path -LiteralPath $Path

    if ($env:PROCESSOR_ARCHITEW6432) {
        $netshexe = Join-Path $env:SystemRoot 'SysNative\netsh.exe'
    }
    else {
        $netshexe = Join-Path $env:SystemRoot 'System32\netsh.exe'
    }

    # Dump some wfp show commands.
    $now = Get-DateTimeString

    $stateFile = Join-Path $Path "wfpstate_$now.xml"
    $null = & $netshexe wfp show state file=$stateFile

    $bootTimePolicyFile = Join-Path $Path "btpol_$now.xml"
    $null = & $netshexe wfp show boottimepolicy file=$bootTimePolicyFile

    $filterFilePath = Join-Path $Path "filters_$now.xml"
    $null = & $netshexe wfp show filters file=$filterFilePath verbose=on

    Write-Log "Starting WFP trace"
    $filePath = Join-Path $Path 'wfp'
    Start-Process $netshexe -ArgumentList "wfp capture start cab=OFF file=`"$filePath`"" -WindowStyle Hidden

    Write-Log "Starting a WFP job"

    $job = Start-Job -ScriptBlock {
        param($Path, $Interval, $MaxDuration, $GetDateTimeString)

        ${Get-DateTimeString} = [ScriptBlock]::Create($GetDateTimeString)

        if ($env:PROCESSOR_ARCHITEW6432) {
            $netshexe = Join-Path $env:SystemRoot 'SysNative\netsh.exe'
        }
        else {
            $netshexe = Join-Path $env:SystemRoot 'System32\netsh.exe'
        }

        $expiration = [DateTime]::Now.Add($MaxDuration)

        while ($true) {
            $now = [DateTime]::Now

            if ($now -gt $expiration) {
                "WfpTrace job was expired at $($expiration.ToString('HH:mm:ss')) after $MaxDuration"
                return
            }

            # dump netevents
            $eventFilePath = Join-Path $Path "netevents_$(& ${Get-DateTimeString} $now).xml"
            $null = & $netshexe wfp show netevents file="$eventFilePath" timewindow=$($Interval.TotalSeconds)

            Start-Sleep -Seconds $Interval.TotalSeconds
        }
    } -ArgumentList $Path, $Interval, $MaxDuration, ${Function:Get-DateTimeString}

    $job
}

function Stop-WfpTrace {
    [CmdletBinding()]
    [Parameter(Mandatory = $true)]
    param (
        $WfpJob
    )

    # Stop WFP trace
    if ($env:PROCESSOR_ARCHITEW6432) {
        $netshexe = Join-Path $env:SystemRoot 'SysNative\netsh.exe'
    }
    else {
        $netshexe = Join-Path $env:SystemRoot 'System32\netsh.exe'
    }

    $null = & $netshexe wfp capture stop

    Write-Log "Stopping a WFP job"
    Stop-Job -Job $WfpJob
    Receive-Job -Job $wfpJob -Wait -AutoRemoveJob | Write-Log
}

<#
.SYNOPSIS
    Save a user-mode memory dump file of a process.
    By default, this function automatically detects if a process is WOW6432 (i.e. 32bit process on 64bit OS), and it collects 32bit process dump in that case.
    To get a 64bit dump with WOW6432 layer, use SkipWow64Check switch parameter.
#>
function Save-Dump {
    [CmdletBinding(PositionalBinding = $false)]
    param(
        # Folder to save a dump file
        [Parameter(Mandatory = $true)]
        $Path,
        [Parameter(Mandatory = $true)]
        [ValidateRange(1, [int]::MaxValue)]
        [int]$ProcessId,
        # Skip WOW64 check.
        [switch]$SkipWow64Check
    )

    # Get the target process.
    $process = Get-Process -Id $ProcessId -ErrorAction SilentlyContinue

    if (-not $process) {
        Write-Error "Cannot find a process with PID $ProcessId"
        return
    }
    elseif (-not $process.SafeHandle) {
        # This scenario is possible for a system process.
        Write-Error "Cannot obtain the process handle of $($process.Name) (PID:$($process.Id))"
        $process.Dispose()
        return
    }

    if (-not (Test-Path $Path)) {
        $null = New-Item $Path -ItemType Directory -ErrorAction Stop
    }

    $Path = Convert-Path -LiteralPath $Path

    $wow64 = $false
    if (-not $SkipWow64Check) {
        # Check if the target process is WOW6432 (i.e. 32bit on 64bit OS)
        $null = [Win32.Kernel32]::IsWow64Process($process.SafeHandle, [ref]$wow64)
    }

    if ($wow64) {
        $ps32 = Join-Path $env:SystemRoot 'SysWOW64\WindowsPowerShell\v1.0\powershell.exe'
        $command = "& {Import-Module '$Script:MyModulePath' -DisableNameChecking; Save-Dump -Path '$Path' -ProcessId $ProcessId -SkipWow64Check}"
        Write-Log "Invoking $ps32 -NoLogo -NoProfile -OutputFormat XML -ExecutionPolicy Unrestricted -Command '$command'"

        $startInfo = New-Object System.Diagnostics.ProcessStartInfo
        $startInfo.FileName = $ps32
        $startInfo.RedirectStandardOutput = $true
        $startInfo.UseShellExecute = $false
        $startInfo.Arguments = "-NoLogo -NoProfile -OutputFormat XML -ExecutionPolicy Unrestricted -Command `"$command`""
        $startInfo.CreateNoWindow = $true

        $psProcess = $null

        try {
            $psProcess = New-Object System.Diagnostics.Process
            $psProcess.StartInfo = $startInfo
            $null = $psProcess.Start()

            $psProcess.WaitForExit()
            $stdOut = $psProcess.StandardOutput.ReadToEnd()

            if ($stdOut) {
                $saveDumpOutput = Join-Path $Path 'saveDumpOutput.xml'
                [IO.File]::AppendAllText($saveDumpOutput, $stdOut)
                Import-Clixml $saveDumpOutput
                Remove-Item $saveDumpOutput -Force
            }
        }
        finally {
            if ($psProcess) {
                $psProcess.Dispose()
            }
        }
    }
    else {
        $dumpFile = Join-Path $Path "$($process.Name)_PID$($ProcessId)_$(Get-DateTimeString).dmp"
        $dumpFileStream = [System.IO.File]::Create($dumpFile)
        $writeDumpSuccess = $false

        $dumpType = [Win32.Dbghelp+MINIDUMP_TYPE]'MiniDumpWithIptTrace, MiniDumpWithTokenInformation, MiniDumpIgnoreInaccessibleMemory, MiniDumpWithThreadInfo, MiniDumpWithFullMemoryInfo, MiniDumpWithUnloadedModules, MiniDumpWithHandleData, MiniDumpWithFullMemory'

        if ([Win32.DbgHelp]::MiniDumpWriteDump($process.SafeHandle, $ProcessId, $dumpFileStream.SafeFileHandle, $dumpType, [IntPtr]::Zero, [IntPtr]::Zero, [IntPtr]::Zero)) {
            [PSCustomObject]@{
                ProcessID   = $ProcessId
                ProcessName = $process.Name
                DumpFile    = $dumpFile
            }
            $writeDumpSuccess = $true
        }
        else {
            Write-Error ("Failed to save a memory dump of $process. Error = 0x{0:x}" -f [System.Runtime.InteropServices.Marshal]::GetLastWin32Error())
        }

        if ($dumpFileStream) {
            $dumpFileStream.Close()

            if (-not $writeDumpSuccess) {
                Remove-Item $dumpFile -Force -ErrorAction SilentlyContinue
            }
        }

        if ($process) {
            $process.Dispose()
        }
    }
}

function Save-HungDump {
    [CmdletBinding(PositionalBinding = $false)]
    param(
        # Folder to save dump files
        [Parameter(Mandatory = $true, Position = 0)]
        [string]$Path,
        # Target process ID
        [Parameter(Mandatory = $true, Position = 1)]
        [ValidateRange(1, [int]::MaxValue)]
        [int]$ProcessId,
        [TimeSpan]$Timeout = [TimeSpan]::FromSeconds(5),
        [int]$DumpCount = 1,
        [Threading.CancellationToken]$CancellationToken,
        [TimeSpan]$DumpInterval = [TimeSpan]::FromSeconds(10)
    )

    if (-not ($process = Get-Process -Id $ProcessId -ErrorAction SilentlyContinue)) {
        Write-Error "Cannnot find a process with PID $ProcessId"
        return
    }

    $WM_NULL = 0
    $SMTO_ABORTIFHUNG = 2
    $savedDumpCount = 0
    $interval = [TimeSpan]::FromSeconds(1)

    # Keep monitoring until one of the following is met:
    # - Cancellation is requested
    # - The target process exits
    # - DumpCount is reached.
    try {
        while ($true) {
            if ($CancellationToken.IsCancellationRequested) {
                Write-Log "Cancel request acknowledged"
                return
            }

            if ($process.HasExited) {
                Write-Log "$($process.Name) (PID:$ProcessId) has exited"
                return
            }

            if (-not $process.Handle) {
                Write-Error "Cannot obtain the process handle of $($process.Name) (PID:$($process.Id))"
                return
            }

            # Need to get the process object every time since MainWindowHandle can change during the life time of a process.
            if ($proc = Get-Process -Id $ProcessId -ErrorAction SilentlyContinue) {
                $hWnd = $proc.MainWindowHandle
                $proc.Dispose()

                # During start up and shut down, MainWindowHandle can be 0.
                if ($hWnd -eq 0) {
                    Start-Sleep -Seconds $interval.TotalSeconds
                    continue
                }
            }
            else {
                Write-Error "Cannnot find a process with PID $ProcessId"
                return
            }

            # > https://groups.google.com/g/microsoft.public.win32.programmer.kernel/c/b-r5qbLwUSA
            # > IsHungAppWindow is a private function used by User32 to determine things like whether SMTO_ABORTIFHUNG should abort or not, or whether it should give up on an application and force-paint its background because the app isn't painting. It is not for general use. (For example, it isn't supported on all Windows platforms)
            # > If you want to see if an app is hung, use SendMessageTimeout with SMTO_ABORTIFHUNG and the WM_NULL message.
            $result = [IntPtr]::Zero
            $ret = [Win32.User32]::SendMessageTimeoutW($hWnd, $WM_NULL, [IntPtr]::Zero, [IntPtr]::Zero, $SMTO_ABORTIFHUNG, $Timeout.TotalMilliseconds, [ref]$result)
            $isHung = $ret -eq 0

            if (-not $isHung) {
                Start-Sleep -Seconds $interval.TotalSeconds
                continue
            }

            # Hung detected. Save a dump file.
            Write-Log "SendMessageTimeoutW() detected a hung window with $($process.Name) (PID:$ProcessId, hWnd:$hWnd). $($savedDumpCount+1)/$DumpCount" -Category Warning
            $dumpResult = Save-Dump -Path $Path -ProcessId $ProcessId

            if ($dumpResult) {
                $savedDumpCount++
                Write-Log "Dump file is saved as '$($dumpResult.DumpFile)'"
            }

            if ($savedDumpCount -ge $DumpCount) {
                Write-Log "Dump count reached $DumpCount. Exiting"
                return
            }

            Write-Log "Pausing $($DumpInterval.TotalSeconds) seconds (or until canceled via CancellationToken) before next hung check"

            if ($CancellationToken) {
                $null = $CancellationToken.WaitHandle.WaitOne($DumpInterval)
            }
            else {
                Start-Sleep -Seconds $DumpInterval.TotalSeconds
            }

            Write-Log "Resuming hung check"
        }
    }
    finally {
        if ($process) {
            $process.Dispose()
        }
    }
}

function Save-MSIPC {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        $Path,
        $User,
        # Copy the entire "MSIPC" folder
        [switch]$All
    )

    # MSIPC info is in %LOCALAPPDATA%\Microsoft\MSIPC
    if ($localAppdata = Get-UserShellFolder -User $User -ShellFolderName 'Local AppData') {
        $msipcPath = Join-Path $localAppdata 'Microsoft\MSIPC\'

        if (-not (Test-Path $msipcPath)) {
            Write-Error "Cannot find path '$msipcPath'"
            return
        }
    }
    else {
        return
    }

    $saveArgs = @{
        Path          = $msipcPath
        IncludeHidden = $true
        Destination   = $Path
        Recurse       = $true
    }

    if ($All) {
        $saveArgs.Exclude = '*.lock', '*.drm'
    }
    else {
        $saveArgs.Filter = '*.ipclog'
    }

    Save-Item @saveArgs
}

function Save-MIP {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        $Path,
        $User
    )

    # MIP data is in %LOCALAPPDATA%\Microsoft\Outlook\MIPSDK\mip
    if ($localAppdata = Get-UserShellFolder -User $User -ShellFolderName 'Local AppData') {
        $mipPath = Join-Path $localAppdata 'Microsoft\Outlook\MIPSDK\mip'

        if (-not (Test-Path $mipPath)) {
            Write-Error "Cannot find path '$mipPath'"
            return
        }
    }
    else {
        return
    }

    $saveArgs = @{
        Path          = $mipPath
        Include       = 'AuthInfoStore.json', 'EngineInfoStore.json', 'EngineStore.json', 'mip.protection.json', 'TemplateStore.json'
        IncludeHidden = $true
        Destination   = $Path
        Recurse       = $true
    }

    Save-Item @saveArgs
}

<#
.SYNOPSIS
    Enable extended logging for MIP

.NOTES
    [HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\Common\DRM]
    "EnableExtendedLogging"=dword:00000001
#>
function Enable-DrmExtendedLogging {
    [CmdletBinding()]
    param (
        $User
    )

    $userRegRoot = Get-UserRegistryRoot -User $User

    if (-not $userRegRoot) {
        Write-Error "Cannot find user registry root for $User"
        return
    }

    $keyPath = Join-Path $userRegRoot 'SOFTWARE\Microsoft\Office\16.0\Common\DRM'

    if (-not (Test-Path $keyPath)) {
        $null = New-Item -Path $keyPath -ErrorAction Stop
    }

    Set-ItemProperty $keyPath -Name 'EnableExtendedLogging' -Value 1 -Type DWord
}

function Disable-DrmExtendedLogging {
    [CmdletBinding()]
    param (
        $User
    )

    $userRegRoot = Get-UserRegistryRoot -User $User

    if (-not $userRegRoot) {
        Write-Error "Cannot find user registry root for $User"
        return
    }

    Remove-ItemProperty -Path (Join-Path $userRegRoot 'SOFTWARE\Microsoft\Office\16.0\Common\DRM') -Name 'EnableExtendedLogging' -ErrorAction SilentlyContinue
}

function Get-DRMConfig {
    [CmdletBinding()]
    param (
        $User
    )

    $userRegRoot = Get-UserRegistryRoot -User $User

    if (-not $userRegRoot) {
        Write-Error "Cannot find user registry root for $User"
        return
    }

    $drmPath = Join-Path $userRegRoot 'SOFTWARE\Microsoft\Office\16.0\Common\DRM'

    & {
        $drmPath
        $drmPath | ConvertTo-PolicyPath
    } `
    | Get-ItemProperty -ErrorAction SilentlyContinue `
    | Split-ItemProperty
}

<#
.SYNOPSIS
    Save Outlook policy nudge files
#>
function Save-PolicyNudge {
    [CmdletBinding(PositionalBinding = $false)]
    param (
        [Parameter(Mandatory = $true)]
        # Destination folder path to save to
        [string]$Path,
        [string]$User
    )

    # Get the path to %LOCALAPPDATA%\Microsoft\Outlook.
    $localAppdata = Get-UserShellFolder -User $User -ShellFolderName 'Local AppData'

    if (-not $localAppdata) {
        Write-Error "Cannot find LocalAppData folder for User $User"
        return
    }

    $sourcePath = Join-Path $localAppdata -ChildPath 'Microsoft\Outlook\'
    $fileNameFilter = 'PolicyNudge*'

    Save-Item -Path $sourcePath -Destination $Path -Filter $fileNameFilter
}


<#
.SYNOPSIS
    Save $env:LOCALAPPDATA\Microsoft\Office\CLP
#>
function Save-CLP {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        # Destination folder path to save to
        [string]$Path,
        [string]$User
    )

    $localAppdata = Get-UserShellFolder -User $User -ShellFolderName 'Local AppData'

    if (-not $localAppdata) {
        Write-Error "Cannot find LocalAppData folder for User $User"
        return
    }

    $sourcePath = Join-Path $localAppdata -ChildPath 'Microsoft\Office\CLP'

    if (-not (Test-Path $sourcePath)) {
        Write-Log "Cannot find $sourcePath"
        return
    }

    Save-Item -Path $sourcePath -Destination $Path -Recurse
}

function Save-DLP {
    [CmdletBinding(PositionalBinding = $false)]
    param (
        [Parameter(Mandatory = $true)]
        # Destination folder path to save to
        [string]$Path,
        [string]$User
    )

    $localAppdata = Get-UserShellFolder -User $User -ShellFolderName 'Local AppData'

    if (-not $localAppdata) {
        Write-Error "Cannot find LocalAppData folder for User $User"
        return
    }

    $sourcePath = Join-Path $localAppdata -ChildPath 'Microsoft\Office\DLP'

    if (-not (Test-Path $sourcePath)) {
        Write-Log "Cannot find $sourcePath"
        return
    }

    Save-Item -Path $sourcePath -Destination $Path -Recurse
}

<#
.SYNOPSIS
This function returns an instance of Microsoft.Identity.Client.LogCallback delegate which calls the given scriptblock when LogCallback is invoked.
#>
function New-LogCallback {
    [CmdletBinding()]
    param (
        # Scriptblock to be called when MSAL invokes LogCallback
        [Parameter(Mandatory = $true)]
        [scriptblock]$Callback,

        # Remaining arguments to be passd to Callback scriptblock via $Event.MessageData
        [Parameter(ValueFromRemainingArguments = $true)]
        [object[]]$ArgumentList
    )

    # Class that exposes an event of type Microsoft.Identity.Client.LogCallback that Register-ObjectEvent can register to.
    $LogCallbackProxyType = @"
        using System;
        using System.Threading;
        using Microsoft.Identity.Client;

        public sealed class LogCallbackProxy
        {
            // This is the exposed event. The sole purpose is for Register-ObjectEvent to hook to.
            public event LogCallback Logging;

            // This is the LogCallback delegate instance.
            public LogCallback Callback
            {
                get { return new LogCallback(OnLogging); }
            }

            // Raise the event
            private void OnLogging(LogLevel level, string message, bool containsPii)
            {
                LogCallback temp = Volatile.Read(ref Logging);
                if (temp != null) {
                    temp(level, message, containsPii);
                }
            }
        }
"@

    if (-not ("LogCallbackProxy" -as [type])) {
        Add-Type $LogCallbackProxyType -ReferencedAssemblies (Join-Path (Split-Path $PSCommandPath) 'modules\Microsoft.Identity.Client.dll')
    }

    $proxy = New-Object LogCallbackProxy
    $null = Register-ObjectEvent -InputObject $proxy -EventName Logging -Action $Callback -MessageData $ArgumentList

    $proxy.Callback
}

<#
.SYNOPSIS
Obtains a modern auth token (maybe from a cached one if available).

.NOTES
You need the following MSAL.NET modules under "modules" sub folder:

 [MSAL.NET](https://www.nuget.org/packages/Microsoft.Identity.Client)
 [MSAL.NET Extensions](https://www.nuget.org/packages/Microsoft.Identity.Client.Extensions.Msal/)

 Folder structure should look like this:

    SomeFolder
    |  OutlookTrace.psm1
    |
    |- modules
          Microsoft.Identity.Client.dll
          Microsoft.Identity.Client.Extensions.Msal.dll

Note about proxy:
MSAL.NET uses System.Net.Http.HttpClient when calling RequestBase.ResolveAuthorityEndpointsAsync(), which reaches "/common/discovery/instance?api-version=1.1&authorization_endpoint=https%3A%2F%2Flogin.microsoftonline.com%2Fcommon%2Foauth2%2Fv2.0%2Fauthorize".
(And this data is cached by Microsoft.Identity.Client.Instance.Discovery.NetworkCacheMetadataProvider in memory. And it won't be fetched next time).
And it also uses System.Windows.Forms.WebBrowser-derived class (named "CustomWebBrowser") when calling InteractiveRequest.GetTokenResponseAsync() to reach "authorize" endpoint
e.g. "/common/oauth2/v2.0/authorize?scope=openid+profile+offline_access&response_type=code&client_id=d3590ed6-52b3-4102-aeff-aad2292ab01d&redirect_uri=https%3A%2F%2Flogin.microsoftonline.com%2Fcommon%2Foauth2%2Fnativeclient...
I can provide the builder's WithHttpClientFactory() with a IMsalHttpClientFactory of a HttpClient with a specific proxy. However I don't think I can do the same for the CustomWebBrowser.
Thus, in order to use a consistent proxy, it's best to configure the user's default proxy in WinInet.

.LINK
[AzureAD/microsoft-authentication-library-for-dotnet](https://github.com/AzureAD/microsoft-authentication-library-for-dotnet)
[AzureAD/microsoft-authentication-extensions-for-dotnet](https://github.com/AzureAD/microsoft-authentication-extensions-for-dotnet)

#>
function Get-Token {
    [CmdletBinding()]
    param(
        # Client ID (Application ID) of the registered application.
        [Parameter(Mandatory = $true)]
        [string]$ClientId,

        # Tenant ID. By default, it uses '/common' endpoint for multi-tenant app. For a single-tenant app, specify the tenant name or GUID (e.g. "contoso.com", "contoso.onmicrosoft.com", "333b3ed5-0ac4-4e75-a1cd-db9e8f593ff3")
        [string]$TenantId = 'common',

        # Array of scopes to request.  By default, "openid", "profile", and "offline_access" are included.
        [string[]]$Scopes,

        # Refirect URI for the application. When this is not given, "https://login.microsoftonline.com/common/oauth2/nativeclient" will be used.
        # Make sure to use the same URI as the one registered for the application.
        [string]$RedirectUri,

        # Clear the cached token and force to get a new token.
        [switch]$ClearCache,

        # Enable MSAL logging. Log file will be msal.log under the script folder.
        [switch]$EnableLogging
    )

    # Need MSAL.NET DLL under modules
    # https://github.com/AzureAD/microsoft-authentication-library-for-dotnet
    # [MSAL.NET](https://www.nuget.org/packages/Microsoft.Identity.Client)
    if (-not ('Microsoft.Identity.Client.AuthenticationResult' -as [type])) {
        try {
            Add-Type -Path (Join-Path (Split-Path $PSCommandPath) 'modules\Microsoft.Identity.Client.dll')
        }
        catch {
            Write-Error -ErrorRecord $_
            return
        }
    }

    # [MSAL.NET Extensions](https://www.nuget.org/packages/Microsoft.Identity.Client.Extensions.Msal/)
    if (-not ('Microsoft.Identity.Client.Extensions.Msal.MsalCacheHelper' -as [type])) {
        try {
            Add-Type -Path (Join-Path (Split-Path $PSCommandPath) 'modules\Microsoft.Identity.Client.Extensions.Msal.dll')
        }
        catch {
            Write-Error -ErrorRecord $_
            return
        }
    }

    # Configure & create a PublicClientApplication
    $builder = [Microsoft.Identity.Client.PublicClientApplicationBuilder]::Create($ClientId).WithAuthority((New-Object System.Uri "https://login.microsoftonline.com/$TenantId/"), <#validateAuthority#> $false)

    if ($RedirectUri) {
        $null = $builder.WithRedirectUri($RedirectUri)
    }
    else {
        # WithDefaultRedirectUri() makes the redirect_uri "https://login.microsoftonline.com/common/oauth2/nativeclient".
        # Without it, redirect_uri would be "urn:ietf:wg:oauth:2.0:oob".
        $null = $builder.WithDefaultRedirectUri()
    }

    $writer = $null

    if ($EnableLogging) {
        $logFile = Join-Path (Split-Path $PSCommandPath) 'msal.log'
        [IO.StreamWriter]$writer = [IO.File]::AppendText($logFile)
        Write-Verbose "MSAL Logging is enabled. Log file:$logFile"

        # Add a CSV header line
        $writer.WriteLine("datetime,level,containsPii,message");

        $null = $builder.WithLogging(
            # Microsoft.Identity.Client.LogCallback
            (New-LogCallback {
                param([Microsoft.Identity.Client.LogLevel]$level, [string]$message, [bool]$containsPii)
                $writer = $Event.MessageData[0]
                $writer.WriteLine("$((Get-Date).ToString('o')),$level,$containsPii,`"$message`"")
            } -ArgumentList $writer),
            [Microsoft.Identity.Client.LogLevel]::Verbose,
            $true, # enablePiiLogging
            $false # enableDefaultPla`tformLogging
        )
    }

    $publicClient = $builder.Build()

    # Configure caching
    $cacheFileName = "msalcache.bin"
    $cacheDir = Split-Path $PSCommandPath
    $storagePropertiesBuilder = New-Object Microsoft.Identity.Client.Extensions.Msal.StorageCreationPropertiesBuilder($cacheFileName, $cacheDir, $ClientId)
    $storageProperties = $storagePropertiesBuilder.Build()
    $cacheHelper = [Microsoft.Identity.Client.Extensions.Msal.MsalCacheHelper]::CreateAsync($storageProperties).GetAwaiter().GetResult()
    $cacheHelper.RegisterCache($publicClient.UserTokenCache)

    if ($ClearCache) {
        $cacheHelper.Clear()
    }

    # Get an account
    $firstAccount = $publicClient.GetAccountsAsync().GetAwaiter().GetResult() | Select-Object -First 1

    # By default, MSAL asks for scopes:openid, profile, and offline_access.
    try {
        $publicClient.AcquireTokenSilent($Scopes, $firstAccount).ExecuteAsync().GetAwaiter().GetResult()
    }
    catch [Microsoft.Identity.Client.MsalUiRequiredException] {
        try {
            $publicClient.AcquireTokenInteractive($Scopes).ExecuteAsync().GetAwaiter().GetResult()
        }
        catch {
            Write-Error -ErrorRecord $_
        }
    }
    catch {
        Write-Error -ErrorRecord $_
    }
    finally {
        if ($writer) {
            $writer.Dispose()
        }
    }
}

<#
.SYNOPSIS
This function makes an Autodiscover request.
#>
function Test-Autodiscover {
    [CmdletBinding()]
    param(
        # Server to send an Autodiscover request. For Exchange Online, use 'outlook.office365.com'
        # When not specified, "autodiscover.{SMTP domain}" will be tried.
        [string]$Server,

        # Target Email address for Autodiscover
        [Parameter(Mandatory = $true)]
        [string]$EmailAddress,

        # Legacy auth credential.
        [Parameter(ParameterSetName = 'LegacyAuth', Mandatory = $true)]
        [System.Management.Automation.PSCredential]
        $Credential,

        # Modern auth access token.
        # To mock an Office client, use ClientId 'd3590ed6-52b3-4102-aeff-aad2292ab01c' and Scope 'https://outlook.office.com/.default'
        # e.g. Get-Token -ClientId 'd3590ed6-52b3-4102-aeff-aad2292ab01c' -Scopes 'https://outlook.office.com/.default' -RedirectUri 'urn:ietf:wg:oauth:2.0:oob'
        [Parameter(ParameterSetName = 'ModernAuth', Mandatory = $true)]
        [string]$Token,

        # Proxy Server
        # e.g. "http://myproxy:8080"
        [string]$Proxy,

        # Skip adding "X-MapiHttpCapability: 1" to the header
        [switch]$SkipMapiHttpCapability,

        # Force Basic auth
        [switch]$ForceBasicAuth,

        # X-AnchorMailbox header value. If this parameter value is missing, $EmailAddress is used.
        # To explicitly turn off X-AnchorMailbox header, specify $null for this parameter.
        [string]$XAnchorMailbox
    )

    $body = @"
<?xml version="1.0" encoding="utf-8"?>
<Autodiscover xmlns="http://schemas.microsoft.com/exchange/autodiscover/outlook/requestschema/2006">
    <Request>
    <EMailAddress>$EmailAddress</EMailAddress>
    <AcceptableResponseSchema>http://schemas.microsoft.com/exchange/autodiscover/outlook/responseschema/2006a</AcceptableResponseSchema>
    </Request>
</Autodiscover>
"@

    $mailDomain = $EmailAddress.Substring($EmailAddress.IndexOf("@") + 1)

    # These are the URL to try (+ redirect URLs).
    # Note that the urls are tried in the reverse order because this is a stack.
    $urls = New-Object System.Collections.Generic.Stack[string](, [string[]]@(
            "http://autodiscover.$mailDomain/autodiscover/autodiscover.xml"
            "https://autodiscover.$mailDomain/autodiscover/autodiscover.xml"
            "https://$Server/autodiscover/autodiscover.xml"
        ))

    $step = 1

    while ($urls.Count -gt 0) {
        $url = $urls.Pop()

        # Check if URL is valid (it could be invalid if $Server is not provided).
        $uri = $null
        if (-not [Uri]::TryCreate($url, [UriKind]::Absolute, [ref]$uri)) {
            Write-Log "Skipping $url because it's invalid"
            continue
        }

        # Arguments for Invoke-WebRequest paramters
        if ($uri.Scheme -eq 'https') {
            $arguments = @{
                Method          = 'POST'
                Uri             = $uri
                Headers         = @{'Content-Type' = 'text/xml' }
                Body            = $body
                UseBasicParsing = $true
            }

            switch -Wildcard ($PSCmdlet.ParameterSetName) {
                'LegacyAuth' {
                    Write-Verbose "Credential is provided. Use it for legacy auth"

                    if ($ForceBasicAuth) {
                        $base64Cred = [System.Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes("$($Credential.UserName):$($Credential.GetNetworkCredential().Password)"))
                        $arguments['Headers'].Add('Authorization', "Basic $base64Cred")
                    }
                    else {
                        $arguments['Credential'] = $Credential
                    }
                    break
                }

                'ModernAuth' {
                    Write-Verbose "Token is provided. Use it for modern auth"
                    $arguments['Headers'].Add('Authorization', "Bearer $Token")
                    break
                }
            }

            if (-not $SkipMapiHttpCapability) {
                $arguments['Headers'].Add('X-MapiHttpCapability', '1')
            }

            # Add X-AnchorMailbox unless XAnchorMailbox param is explicitly given a null
            # By default, use EmailAddress as X-AnchorMailbox.
            if ($PSBoundParameters.ContainsKey('XAnchorMailbox')) {
                if ($XAnchorMailbox) {
                    $arguments['Headers'].Add('X-AnchorMailbox', $XAnchorMailbox)
                }
            }
            else {
                $arguments['Headers'].Add('X-AnchorMailbox', $EmailAddress)
            }
        }
        else {
            $arguments = @{
                Method             = 'GET'
                Uri                = $uri
                MaximumRedirection = 0 # Just get 302 and don't follow the redirect.
                UseBasicParsing    = $true
            }
        }

        if ($Proxy) {
            $arguments['Proxy'] = $Proxy
        }

        # Make a web request.
        Write-Log "Trying $($arguments.Method) $($arguments.Uri)"
        $result = $null
        $err = $($result = Invoke-WebRequest @arguments) 2>&1

        # Check result
        if ($result.StatusCode -eq 200) {
            [PSCustomObject]@{
                Step    = $step++
                URI     = $uri
                Success = $true
                Result  = $result
            }
            return
        }
        elseif ($uri.Scheme -eq 'http' -and $result.StatusCode -eq 302) {
            # See if we got 302 with Location header
            $redirectUrl = $null
            if ($result.StatusCode -eq 302) {
                $redirectUrl = $result.Headers['Location']
            }

            if ($redirectUrl) {
                $result | Add-Member -MemberType ScriptMethod -Name 'ToString' -Force -Value { "Received a redirect URL $($this.Headers['Location'])" }

                [PSCustomObject]@{
                    Step    = $step++
                    URI     = $uri
                    Success = $true
                    Result  = $result
                }

                # Try the given redirect uri next
                Write-Log "Received a redirect URL:$redirectUrl"
                $urls.Push($redirectUrl)
            }
            else {
                [PSCustomObject]@{
                    Step    = $step++
                    URI     = $uri
                    Success = $false
                    Result  = $err
                }
            }
        }
        else {
            [PSCustomObject]@{
                Step    = $step++
                URI     = $uri
                Success = $false
                Result  = $err
            }
        }
    }
}

<#
.SYNOPSIS
Convert a ProgID to CLSID
#>
function ConvertTo-CLSID {
    [CmdletBinding()]
    param(
        [parameter(Mandatory = $true)]
        [string]$ProgId,
        [string]$User
    )

    [uint32]$S_OK = 0

    [Guid]$CLSID = [Guid]::Empty
    [uint32]$hr = [Win32.Ole32]::CLSIDFromProgID($ProgId, [ref]$CLSID)
    $path = $null

    if ($hr -ne $S_OK) {
        $userRegRoot = Get-UserRegistryRoot -User $User

        if (-not $userRegRoot) {
            return
        }

        $locations = @(
            # ClickToRun Registry & the user's Classes
            "Registry::HKLM\SOFTWARE\Microsoft\Office\ClickToRun\REGISTRY\MACHINE\Software\Classes\"
            Join-Path $userRegRoot "SOFTWARE\Classes\"
        )

        foreach ($loc in $locations) {
            $clsidProp = Get-ItemProperty (Join-Path $loc "$ProgId\CLSID") -ErrorAction SilentlyContinue

            if (-not $clsidProp) {
                # See if CurVer key is available
                if ($curVerProp = Get-ItemProperty (Join-Path $loc "$ProgId\CurVer") -ErrorAction SilentlyContinue) {
                    $curProgId = $curVerProp.'(default)'
                    $clsidProp = Get-ItemProperty (Join-Path $loc "$curProgId\CLSID") -ErrorAction SilentlyContinue
                }
            }

            if ($clsidProp) {
                $CLSID = $clsidProp.'(default)'
                $path = $clsidProp | Convert-Path -ErrorAction SilentlyContinue
                break
            }
        }

        if ($CLSID -eq [Guid]::Empty) {
            Write-Error -Message $("CLSIDFromProgID for `"$ProgId`" failed with 0x{0:x}. Also, it was not found in the ClickToRun & user registry" -f $hr)
            return
        }
    }

    # CLSID found. Get its string representation.
    [IntPtr]$pClsIdString = [IntPtr]::Zero
    $hr = [Win32.Ole32]::StringFromCLSID($CLSID, [ref]$pCLSIDString)

    if ($hr -eq $S_OK -and $pCLSIDString) {
        $CLSIDString = [System.Runtime.InteropServices.Marshal]::PtrToStringUni($pCLSIDString)
        [System.Runtime.InteropServices.Marshal]::FreeCoTaskMem($pCLSIDString)
        $pCLSIDString = [IntPtr]::Zero
    }

    [PSCustomObject]@{
        GUID   = $CLSID
        String = $CLSIDString
        Path   = $path # Where CLSID is found. null indicates it's found by CLSIDFromProgID API
    }
}

<#
.SYNOPSIS
Get Outlook's COM addins
#>
function Get-OutlookAddin {
    [CmdletBinding()]
    param(
        # User name or SID
        [string]$User
    )

    $userRegRoot = Get-UserRegistryRoot $User

    if (-not $userRegRoot) {
        return
    }

    $officeInfo = Get-OfficeInfo

    # Depending on the arch type of Outlook/MAPI, change CLSID search paths order. If it is x86, search Wow6432 first.
    # The order of keys matters here for performance.
    # Checking sub key of HKEY_CLASSES_ROOT\CLSID\ & HKEY_CLASSES_ROOT\WOW6432Node\CLSID\ is quite slow when the path does not exist (> 100 ms). Thus they are checked later.
    $arch = $officeInfo.Architecture

    $clsIdSearchPaths = @(
        if ($arch -eq 'x86') {
            'Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\ClickToRun\REGISTRY\MACHINE\Software\Classes\Wow6432Node\CLSID'
            'Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\ClickToRun\REGISTRY\MACHINE\Software\Classes\CLSID\'
            'Registry::HKEY_CLASSES_ROOT\WOW6432Node\CLSID\'
            'Registry::HKEY_CLASSES_ROOT\CLSID\'
        }
        else {
            # Must be x64
            'Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\ClickToRun\REGISTRY\MACHINE\Software\Classes\CLSID\'
            'Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\ClickToRun\REGISTRY\MACHINE\Software\Classes\Wow6432Node\CLSID'
            'Registry::HKEY_CLASSES_ROOT\CLSID\'
            'Registry::HKEY_CLASSES_ROOT\WOW6432Node\CLSID\'
        })

    $LoadBehavior = @{
        0  = 'None'
        1  = 'NoneLoaded'
        2  = 'StartupUnloaded'
        3  = 'Startup'
        8  = 'LoadOnDemandUnloaded'
        9  = 'LoadOnDemand'
        16 = 'LoadAtNextStartupOnly'
    }

    @(
        'Registry::HKLM\SOFTWARE\Microsoft\Office\Outlook\Addins'
        'Registry::HKLM\SOFTWARE\WOW6432Node\Microsoft\Office\Outlook\Addins'
        'Registry::HKLM\SOFTWARE\Microsoft\Office\ClickToRun\REGISTRY\MACHINE\Software\Microsoft\Office\Outlook\Addins'
        'Registry::HKLM\SOFTWARE\Microsoft\Office\ClickToRun\REGISTRY\MACHINE\Software\Wow6432Node\Microsoft\Office\Outlook\AddIns'
        Join-Path $userRegRoot 'Software\Microsoft\Office\Outlook\Addins'
    ) | Get-ChildItem -ErrorAction SilentlyContinue | & {
        param([Parameter(ValueFromPipeline)]$addin)
        begin { $cache = @{} }
        process {
            try {
                $props = [ordered]@{
                    ProgId       = $addin.PSChildName
                    CLSID        = $null
                    # ToDo:text might get garbled in DBCS environment.
                    FriendlyName = $addin.GetValue('FriendlyName')
                    Description  = $addin.GetValue('Description')
                    Location     = $null
                    Path         = $addin.Name
                }

                # First check LoadBehavior and if it's missing, ignore this entry
                $loadBehaviorValue = $addin.GetValue('LoadBehavior')

                if ($loadBehaviorValue) {
                    $props.LoadBehavior = $LoadBehavior[$loadBehaviorValue]
                }
                else {
                    Write-Log "Skipping $($props.ProgId) because its LoadBehavior is null"
                    return
                }

                if ($cache.ContainsKey($props.ProgId)) {
                    Write-Log "Skipping $($props.ProgId) because it's already found"
                    return
                }
                else {
                    $cache.Add($props.ProgId, $null)
                }

                # Try to get CLSID.
                $clsidErr = $($clsid = ConvertTo-CLSID $props.ProgId -User $User -ErrorAction Continue) 2>&1

                if ($clsid) {
                    $props.CLSID = $clsid.String

                    if ($clsid.Path) {
                        Write-Log "CLSID of $($props.ProgId) is found at $($clsid.Path)"
                    }

                    # e.g. "...\CLSID\{C15AC6D0-15EE-49B3-9B2A-948320426B88}\InprocServer32"
                    # Check InprocServer32, LocalServer32, RemoteServer32
                    $null = & { 'InprocServer32'; 'LocalServer32'; 'RemoteServer32' } `
                    | & {
                        param([Parameter(ValueFromPipeline)]$comServerType)
                        process {
                            $clsIdSearchPaths | Join-Path -ChildPath $props.CLSID | Join-Path -ChildPath $comServerType
                        }
                    } `
                    | Get-ItemProperty -ErrorAction SilentlyContinue `
                    | & {
                        param([Parameter(ValueFromPipeline)]$comSpec)
                        process {
                            $props.Location = $comSpec.'(default)'
                            $props.ThreadingModel = $comSpec.ThreadingModel
                            $props.CodeBase = $comSpec.CodeBase
                            # Stop the pipeline
                            $true
                        }
                    } `
                    | Select-Object -First 1
                }
                elseif ($manifest = $addin.GetValue('Manifest')) {
                    # A managed addin does not have CLSID. Check "Manifest" instead.
                    $props.Location = $manifest
                    Write-Log "Manifest is found for $($props.ProgId). This is a VSTO addin"
                }
                else {
                    # If both CLSID & Manifest are missing, ignore this entry.
                    $clsidErr | Write-Log -Category Warning
                    return
                }

                [PSCustomObject]$props
            }
            finally {
                $addin.Close()
            }
        }
    } | Sort-Object -Property 'ProgId'
}

function Get-ClickToRunConfiguration {
    [CmdletBinding()]
    param()

    Get-ItemProperty Registry::HKLM\SOFTWARE\Microsoft\Office\ClickToRun\Configuration
}

function Get-WebView2 {
    [CmdletBinding(PositionalBinding = $false)]
    param (
        [Parameter(Position = 0)]
        $User
    )

    $userRegRoot = Get-UserRegistryRoot -User $User

    if (-not $userRegRoot) {
        return
    }

    # See https://docs.microsoft.com/en-us/microsoft-edge/webview2/concepts/distribution
    & {
        'Registry::HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Microsoft\EdgeUpdate\Clients\{F3017226-FE2A-4295-8BDF-00C3A9A7E4C5}'
        'Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\EdgeUpdate\Clients\{F3017226-FE2A-4295-8BDF-00C3A9A7E4C5}'
        Join-Path $userRegRoot 'Software\Microsoft\EdgeUpdate\Clients\{F3017226-FE2A-4295-8BDF-00C3A9A7E4C5}'
    } | Get-ItemProperty -ErrorAction SilentlyContinue
}

function Get-DeviceJoinStatus {
    [CmdletBinding()]
    param()

    $dsregcmd = 'dsregcmd.exe'

    if (Get-Command $dsregcmd -ErrorAction SilentlyContinue) {
        & $dsregcmd /status
    }
    else {
        Write-Log "$dsregcmd is not available"
    }
}

<#
This function just starts C:\Windows\System32\gatherNetworkInfo.vbs and returns a process
#>
function Start-GatherNetworkInfo {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        $Path
    )

    if (-not (Test-Path $Path)) {
        $null = New-Item -ItemType Directory -Path $Path -ErrorAction Stop
    }

    Invoke-Command {
        $ErrorActionPreference = 'Continue'
        Start-Process cscript.exe -ArgumentList "C:\windows\system32\gatherNetworkInfo.vbs" -WorkingDirectory $Path -PassThru
    }
}

function Stop-GatherNetworkInfo {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        $Process,
        [TimeSpan]$Timeout
    )

    Write-Log "Waiting for gatherNetworkInfo.vbs to finish $(if ($Timeout) {"with timeout $Timeout"})"
    if ($null -eq $Timeout) {
        $completed = $process.WaitForExit()
    }
    else {
        $completed = $process.WaitForExit($Timeout.TotalMilliseconds)
    }

    if (-not $completed) {
        Write-Log "gatherNetworkInfo reached timeout $Timeout"
        $process.Kill()
    }

    $process.Dispose()
}

function Start-PerfTrace {
    [CmdletBinding(PositionalBinding = $false)]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,
        [string]$FileName = "Perf_$(Get-DateTimeString)",
        [ValidateRange(1, [int]::MaxValue)]
        [int]$IntervalSecond = 1,
        [ValidateRange(1, [int]::MaxValue)]
        [int]$MaxFileSizeMB = 1024,
        [ValidateSet('NewFile', 'Circular')]
        [string]$LogFileMode = 'NewFile'
    )

    if (-not (Test-Path -LiteralPath $Path)) {
        $null = New-Item $Path -ItemType Directory -ErrorAction Stop
    }
    $Path = Convert-Path -LiteralPath $Path

    $counters = @(
        '\LogicalDisk(*)\*'
        '\Memory\*'
        '\Network Interface(*)\*'
        '\Paging File(*)\*'
        '\PhysicalDisk(*)\*'
        '\Process(*)\*'
        '\Processor Information(*)\*'
        '\Processor(*)\*'
        '\TCPv4\*'
        '\TCPv6\*'
        '\System\Processor Queue Length'
    )

    $configFile = Join-Path $Path "perf.config"
    # Note:Encoding must be Ascii here ('utf8' will write as UTF-8 with BOM, which does not work for logman)
    Set-Content -LiteralPath $configFile -Value $counters -Encoding Ascii

    $filePath = Join-Path $Path $FileName
    Write-Log "Staring PerfCounter. Mode:$LogFileMode, IntervalSecond:$IntervalSecond, MaxFileSizeMB:$MaxFileSizeMB, FilePath:$filePath"

    switch ($LogFileMode) {
        'NewFile' {
            $stdout = & logman.exe create counter -n 'PerfCounter' -cf $configFile -si $IntervalSecond -max $MaxFileSizeMB -o $filePath -ow --v -f 'bin' -cnf 0
            break
        }

        'Circular' {
            $stdout = & logman.exe create counter -n 'PerfCounter' -cf $configFile -si $IntervalSecond -max $MaxFileSizeMB -o $filePath -ow --v -f 'bincirc' # -cnf 0
            break
        }
    }

    if ($LASTEXITCODE -ne 0) {
        Write-Error "logman failed with 0x$('{0:x}' -f $LASTEXITCODE). $stdout"
        return
    }

    $stdout = & logman.exe start 'PerfCounter'

    if ($LASTEXITCODE -ne 0) {
        Write-Error "logman failed with 0x$('{0:x}' -f $LASTEXITCODE). $stdout"
    }
}

function Stop-PerfTrace {
    [CmdletBinding(PositionalBinding = $false)]
    param(
        [string]$DataCollectorSetName = 'PerfCounter'
    )

    $stdout = & logman query $DataCollectorSetName
    if ($LASTEXITCODE -ne 0) {
        Write-Error "logman query $DataCollectorSetName failed with 0x$('{0:x}' -f $LASTEXITCODE). $stdout"
        return
    }

    Write-Log "Stopping $DataCollectorSetName"
    $stdout = & logman.exe stop $DataCollectorSetName
    $stdout = & logman.exe delete $DataCollectorSetName

    if ($LASTEXITCODE -ne 0) {
        Write-Error "logman failed with 0x$('{0:x}' -f $LASTEXITCODE). $stdout"
        return
    }
}

<#
Get processes and its user (only for Outlook.exe & Fiddler*).
PowerShell 4's Get-Process has -IncludeUserName, but I'm using WMI here for now.
#>
function Save-Process {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        $Path,
        # Name of process whose owner info is also saved.
        [string[]]$Name
    )

    if (-not (Test-Path $Path)) {
        $null = New-Item -ItemType Directory -Path $Path -ErrorAction Stop
    }

    $outFileName = "Win32_Process_$(Get-DateTimeString).xml"

    Get-CimInstance Win32_Process | & {
        param([Parameter(ValueFromPipeline)]$win32Process)
        process {
            $processName = [IO.Path]::GetFileNameWithoutExtension($win32Process.ProcessName)

            # For processes specified in Name parameter, save its Owner & Environment Variables.
            if ($Name | Where-Object { $processName -like $_ } | Select-Object -First 1) {
                try {
                    # GetOwner() could fail if the process has exited. Not likely, but be defensive here.
                    $owner = Invoke-CimMethod -InputObject $win32Process -MethodName GetOwner
                    $win32Process | Add-Member -MemberType NoteProperty -Name 'User' -Value "$($owner.Domain)\$($owner.User)"

                    $ownerSid = Invoke-CimMethod -InputObject $win32Process -MethodName GetOwnerSid | Select-Object -ExpandProperty Sid
                    $win32Process | Add-Member -MemberType NoteProperty -Name 'UserSid' -Value $ownerSid

                    # Add Environment variables
                    if ($proc = Get-Process -Id $win32Process.ProcessId -ErrorAction SilentlyContinue) {
                        $win32Process | Add-Member -MemberType NoteProperty -Name 'EnvironmentVariables' -Value $proc.StartInfo.EnvironmentVariables
                        $proc.Dispose()
                    }
                }
                catch {
                    # Ignore
                    Write-Error -ErrorRecord $_
                }
            }

            $win32Process
            $win32Process.Dispose()
        }
    } | Export-Clixml -Path (Join-Path $Path $outFileName)

    Write-Log "Win32_Process saved as $outFileName"
}

<#
.SYNOPSIS
    Helper function to get a hash code of a Process (*), based on ProcessId & CreationDate.
    * Process can be either System.Diagnostics.Process or WIN32_PROCESS CimClass instance.

    Note: Be careful when passing a System.Diagnostics.Process because its StartTime might cause "Access is denied" if running without Debug Privilege
#>
function Get-ProcessHash {
    [OutputType([Int32])]
    param(
        [Parameter(ValueFromPipelineByPropertyName)]
        [Alias('ProcessId')]
        [int]$Id,
        [Parameter(ValueFromPipelineByPropertyName)]
        [Alias('CreationDate')]
        [DateTime]$StartTime
    )

    process {
        Get-CombinedHash $Id.GetHashCode(), $StartTime.GetHashCode()
    }
}

<#
.SYNOPSIS
    Helper function to combine hash values

.NOTES
    It'd be nice to be able to use HashCode.Combine(), but that's only available for .NET Core.
    The current implementation is from boost::hash_combine():
    https://www.boost.org/doc/libs/1_55_0/doc/html/hash/reference.html#boost.hash_combine
#>
function Get-CombinedHash {
    [OutputType([Int32])]
    [CmdletBinding()]
    param(
        [int[]]$HashValues
    )

    [UInt32]$hash = 0

    # The Golden Ratio (i.e. ([Math]::sqrt(5) - 1) / 2 * [Math]::Pow(2, 32)).
    # Suffix "L" for long. Avoid conversion to Double.
    [UInt64]$random = 0x9e3779b9L

    foreach ($val in $HashValues) {
        $hash = $hash -bxor ($val + $random + ($hash -shl 6) + ($hash -shr 2)) % 0x100000000
    }

    ConvertTo-Int32 $hash
}

<#
.SYNOPSIS
    Reinterpret Int32 to UInt32.
    Note that [Convert]::ToUInt32() does not work in general because it throws an OverflowException when the input is a negative value.
#>

function ConvertTo-UInt32 {
    [OutputType([UInt32])]
    param(
        [int]$i
    )

    [UInt32]$u = $i -band 0x7fffffff

    if ($i -band 0x80000000) {
        $u = $u -bor 0x80000000
    }

    $u
}

<#
.SYNOPSIS
    Reinterpret UInt32 to Int32.
    Note that [Convert]::Int32() does not work in general because it throws an OverflowException when the input is larger than [int]::MaxValue.
#>
function ConvertTo-Int32 {
    [OutputType([Int32])]
    param(
        [UInt32]$u
    )

    [int]$i = $u -band 0x7fffffff

    if ($u -band 0x80000000) {
        $i = $i -bor 0x80000000
    }

    $i
}

<#
.SYNOPSIS
    Start enumerating processes
.DESCRIPTION
    This command starts enumerating Win32 processes until canceled via a CancellationToken, and it repeats with the given interval.
    For processes whose name matches the NamePattern parameter, their User and Environment Variables are also retrieved.
#>
function Start-ProcessCapture {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        # Folder path to save process list
        [string]$Path,
        [Parameter(Mandatory = $true)]
        # Regex Pattern of process names to fetch details
        [string]$NamePattern,
        [System.Threading.CancellationToken]$CancellationToken,
        [TimeSpan]$Interval = [TimeSpan]::FromSeconds(1),
        [System.Threading.EventWaitHandle]$StartedEvent,
        # This is just for testing
        [Switch]$EnablePerfCheck
    )

    if (-not (Test-Path $Path)) {
        $null = New-Item $Path -ItemType Directory -ErrorAction Stop
    }

    # This is an optimization. Get-Process's IncludeUserName is faster than Win32_Process CimClass's GetOwnerSid method (although Win32_Process's GetOwner(Sid) works without Debug Privilege)
    $useIncludeUserName = (Get-Command Get-Process).Parameters.ContainsKey('IncludeUserName') -and (Test-DebugPrivilege)
    Write-Log "useIncludeUserName:$useIncludeUserName"

    $runAsAdmin = Test-RunAsAdministrator

    # HashTable of final objects to be returned (Using a HashTable is much faster than Where-Object)
    $win32ProcTable = @{}

    # HashTable for System.Diagnostics.Process
    $procTable = @{}

    # If StartedEvent is given, signal after the first iteration.
    $needSignal = $null -ne $StartedEvent

    while ($true) {
        $start = Get-Timestamp

        Get-CimInstance Win32_Process | & {
            param ([Parameter(ValueFromPipeline)]$win32Process)
            process {
                try {
                    # Don't use GetHashCode() because it changes for the same process in each iteration.
                    $key = $win32Process | Get-ProcessHash

                    if ($win32ProcTable.ContainsKey($key)) {
                        # Very rare, but cache collision could occur.
                        $cache = $win32ProcTable[$key]

                        if ($cache.Id -eq $win32Process.ProcessId -and $cache.StartTime -eq $win32Process.CreationDate) {
                            return
                        }
                    }

                    $obj = [ordered]@{
                        Name                 = $win32Process.Name
                        Id                   = $win32Process.ProcessId
                        Path                 = $win32Process.Path
                        CommandLine          = $win32Process.CommandLine
                        ParentProcessId      = $win32Process.ParentProcessId
                        StartTime            = $win32Process.CreationDate
                        ExitTime             = $null
                        Elevated             = $null
                        User                 = $null
                        EnvironmentVariables = $null
                    }

                    if ($useIncludeUserName) {
                        $proc = Get-Process -Id $win32Process.ProcessId -IncludeUserName -ErrorAction SilentlyContinue
                    }
                    else {
                        $proc = Get-Process -Id $win32Process.ProcessId -ErrorAction SilentlyContinue
                    }

                    if ($proc) {
                        $procTable.Add($key, $proc)
                        $obj.User = $proc.UserName

                        # To get ExitTime later, touch Handle property
                        $null = $proc.Handle
                    }

                    # Check if process is elevated, except for "System Idle Process" (PID 0) and "System" (PID 4)
                    if ($win32Process.ProcessId -gt 4) {
                        try {
                            $err = $($obj.Elevated = Test-ProcessElevated $win32Process.ProcessId) 2>&1 | Select-Object -First 1
                        }
                        catch {
                            $err = $_
                        }

                        if ($err) {
                            if (-not $runAsAdmin -and $err.Exception.NativeErrorCode -eq 5) {
                                # If not running as admin, Test-ProcessElevated is expected to fail with Access Denied (5) for some processes. No need to log this error.
                            }
                            else {
                                $errMsg = "Test-ProcessElevated failed for $($win32Process.Name) (PID:$($win32Process.ProcessId))"

                                # Maybe the process is gone already. In this case, OpenProcess would fail with ERROR_INVALID_PARAMETER (87).
                                if ($proc) {
                                    Write-Log $errMsg -ErrorRecord $err -Category Error
                                }
                                else {
                                    Write-Log "$errMsg because the process has already exited"
                                }
                            }
                        }
                    }

                    # For processes specified in NamePattern parameter, save its User, Modules, and EnvironmentVariables
                    if ($win32Process.Name -match $NamePattern) {
                        Write-Log "Found a new instance of $($win32Process.Name) (PID:$($win32Process.ProcessId), Elevated:$($obj.Elevated))"

                        # If not found by Get-Process, then retrieve from WIN32_PROCESS
                        if (-not $obj.User -and ($owner = $win32Process | Get-ProcessOwner)) {
                            $obj.User = $owner.Name
                        }

                        if ($proc) {
                            $obj.EnvironmentVariables = $proc.StartInfo.EnvironmentVariables

                            # To reduce the output size, include only necessary properties
                            $obj.Modules = $proc.Modules | Select-Object ModuleName, FileName, @{N = 'FileVersion'; E = { $_.FileVersionInfo.FileVersion } }, @{N = 'Language'; E = { $_.FileVersionInfo.Language } }
                        }
                    }

                    $win32ProcTable.Add($key, [PSCustomObject]$obj)
                }
                finally {
                    $win32Process.Dispose()
                }
            }
        }

        if ($needSignal) {
            $null = $StartedEvent.Set()
            $needSignal = $false
        }

        if ($EnablePerfCheck) {
            $elapsed = Get-Elapsed $start
            Write-Log "Processing Win32_Process took $($elapsed.TotalMilliseconds) ms"
        }

        if ($CancellationToken.IsCancellationRequested) {
            Write-Log "Cancel request acknowledged"
            break
        }

        Start-Sleep -Seconds $Interval.TotalSeconds
    }

    # Check ExitTime & dispose System.Diagnostics.Process instances
    foreach ($key in $procTable.Keys) {
        $proc = $procTable[$key]
        $err = $($win32ProcTable[$key].ExitTime = $proc.ExitTime) 2>&1

        # TODO: remove later
        if ($err) {
            Write-Log "Failed to get ExitTime of $($proc.Name) (PID:$($proc.Id))"
        }

        $proc.Dispose()
    }

    $win32ProcTable.Values | Export-Clixml -Path (Join-Path $Path 'Win32_Process.xml')
}

<#
.SYNOPSIS
    Start PSR as a task and restart after some time until canceled.
    This creates PSR_***.mht in $Path. When $Circular, only files writen within the last 1 hour will be kept.
#>
function Start-PsrMonitor {
    [CmdletBinding()]
    param(
        [string]$Path,
        [System.Threading.CancellationToken]$CancellationToken,
        [TimeSpan]$WaitInterval,
        [System.Threading.EventWaitHandle]$StartedEvent,
        [bool]$Circular
    )

    if ($StartedEvent) {
        $null = $StartedEvent.Set()
    }

    while ($true) {
        $retry = 0
        $maxRetry = 3
        $retryInterval = [TimeSpan]::FromSeconds(1)
        $startResult = $null

        while ($true) {
            $err = $($startResult = Start-PSR -Path $Path -FileName "PSR_$(Get-DateTimeString)") 2>&1 | Select-Object -First 1

            if ($startResult) {
                break
            }

            # PSR failed to start. Maybe an instance is already running? Retry up to maxRetry times.
            if ($retry -ge $maxRetry) {
                Write-Log -Message "PSR failed to start after $retry retries" -ErrorRecord $err -Category Error
                return
            }

            $retry++
            Write-Log -Message "PSR failed to start. $err. Retrying ($retry/$maxRetry)" -Category Warning
            Stop-PSR -ErrorAction SilentlyContinue
            Start-Sleep -Seconds $retryInterval.TotalSeconds
        }

        $canceled = $CancellationToken.WaitHandle.WaitOne($WaitInterval)
        Stop-PSR -StartResult $startResult

        if ($canceled) {
            Write-Log "PSR task cancellation is acknowledged"
            break
        }

        if ($circular) {
            # Remove mht files older than 1 hour
            Get-ChildItem $Path -Filter '*.mht' | & {
                begin {
                    $cutoff = [DateTime]::Now.AddHours(-1)
                    $removedCount = 0
                }

                process {
                    if ($_.LastWriteTime -lt $cutoff) {
                        Remove-Item $_.FullName
                        ++$removedCount
                    }
                }

                end {
                    if ($removedCount) {
                        Write-Log "$removedCount mht files were removed because they were older than 1 hour"
                    }
                }
            }
        }
    }
}

function Start-HungMonitor {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        # Destination folder path
        [string]$Path,
        [Parameter(Mandatory)]
        # Name of the target process (e.g. "Outlook")
        [string]$Name,
        [Parameter(Mandatory)]
        # Owner of the target process
        $User,
        [TimeSpan]$Timeout,
        [int]$DumpCount,
        [System.Threading.CancellationToken]$CancellationToken,
        [System.Threading.EventWaitHandle]$StartedEvent
    )

    # Remove the extension ".exe" if exits.
    $Name = [IO.Path]::GetFileNameWithoutExtension($Name)

    if (-not $User.Sid) {
        $User = Resolve-User $User

        if (-not $User) {
            return
        }
    }

    if ($StartedEvent) {
        $null = $StartedEvent.Set()
    }

    # Key:Process Hash, Value:true/false for "need to log".
    $procCache = @{}
    $interval = [TimeSpan]::FromSeconds(1)

    # The target process may restart while being monitored. Keep monitoring until canceled (via CancellationToken).
    while ($true) {
        # Wait for the target process ($Name) to come live.
        while ($true) {
            if ($CancellationToken.IsCancellationRequested) {
                return
            }

            # Find the target processes run by the specified user (There could be multiple process instances)
            $targetList = @(
                Get-CimInstance 'Win32_Process' -Filter "Name = '$Name.exe'" | & {
                    param([Parameter(ValueFromPipeline)]$win32Process)
                    process {
                        try {
                            # Don't use GetHashCode() because it changes for the same process in each iteration.
                            $key = $win32Process | Get-ProcessHash

                            if ($procCache.ContainsKey($key)) {
                                if ($procCache[$key]) {
                                    Write-Log "This instance of $($win32Process.Name) (PID:$($win32Process.ProcessId), CreationDate:$($win32Process.CreationDate)) has been seen already. This instance will be not be monitored"
                                    $procCache[$key] = $false
                                }

                                return
                            }

                            # Check if the owner matches the User
                            $owner = $win32Process | Get-ProcessOwner

                            if ($owner -and $owner.Sid -ne $User.Sid) {
                                Write-Log "This instance of $($win32Process.Name) (PID:$($win32Process.ProcessId)) has owner '$owner', and it is different from the target user '$User'. This instance will be not be monitored"
                                $procCache.Add($key, $false)
                                return
                            }

                            # Found a target process
                            $procCache.Add($key, $true)

                            [PSCustomObject]@{
                                Id           = $win32Process.ProcessId
                                CreationDate = $win32Process.CreationDate
                            }
                        }
                        finally {
                            $win32Process.Dispose()
                        }
                    }
                }
            )

            # If there are multiple processes available, pick the one that started earliest.
            $target = $targetList | Sort-Object CreationDate | Select-Object -First 1

            if ($targetList.Count -gt 1) {
                Write-Log "There are $($targetList.Count) instances of $Name found"
            }

            if ($target) {
                break
            }

            Start-Sleep -Seconds $interval.TotalSeconds
        }

        Write-Log "Found $Name (PID:$($target.Id), CreationDate:$($target.CreationDate)). Starting hung window monitoring"

        $hungDumpArgs = @{
            Path      = $Path
            ProcessId = $target.Id
        }

        if ($DumpCount) {
            $hungDumpArgs.DumpCount = $DumpCount
        }

        if ($Timeout) {
            $hungDumpArgs.Timeout = $Timeout
        }

        if ($CancellationToken) {
            $hungDumpArgs.CancellationToken = $CancellationToken
        }

        Save-HungDump @hungDumpArgs 2>&1 | Write-Log -Category Error -PassThru
    }
}

<#
Check GitHub's latest release and if it's newer, download and import it except if OutlookTrace is installed as module.
#>
function Invoke-AutoUpdate {
    [CmdletBinding()]
    param(
        [uri]$GitHubUri = 'https://api.github.com/repos/jpmessaging/OutlookTrace/releases/latest'
    )

    $autoUpdateSuccess = $false
    $message = $null
    $module = $PSCmdlet.MyInvocation.MyCommand.Module

    if ($module.Version.ToString() -ne '0.0') {
        $message = "Skipped autoupdate because OutlookTrace seems be installed as a module"
    }
    elseif (-not (Get-NLMConnectivity).IsConnectedToInternet) {
        $message = "Skipped autoupdate because there's no connectivity to internet"
    }
    else {
        try {
            Write-Progress -Activity "AutoUpdate" -Status 'Checking if a newer version is available. Please wait' -PercentComplete -1
            $release = Invoke-RestMethod -Uri $GitHubUri -UseDefaultCredentials -ErrorAction Stop

            if ($Version -ge $release.name) {
                $message = "Skipped because the current script ($Version) is newer than or equal to GitHub's latest release ($($release.name))"
            }
            else {
                Write-Verbose "Downloading the latest script"

                $response = Invoke-Command {
                    # Suppress progress on Invoke-WebRequest.
                    $ProgressPreference = "SilentlyContinue"
                    Invoke-WebRequest -Uri $release.assets.browser_download_url -UseDefaultCredentials -UseBasicParsing
                }

                # Rename the current script and replace with the latest one.
                $newName = [IO.Path]::GetFileNameWithoutExtension($PSCommandPath) + '_' + $Version + [IO.Path]::GetExtension($PSCommandPath)

                if (Test-Path (Join-Path ([IO.Path]::GetDirectoryName($PSCommandPath)) $newName)) {
                    $newName = [IO.Path]::GetFileNameWithoutExtension($PSCommandPath) + '_' + $Version + '_' + [IO.Path]::GetRandomFileName() + [IO.Path]::GetExtension($PSCommandPath)
                }

                Rename-Item -LiteralPath $PSCommandPath -NewName $newName -ErrorAction Stop
                [IO.File]::WriteAllBytes($PSCommandPath, $response.Content)

                Write-Verbose "Lastest script ($($release.name)) was successfully downloaded"
                Import-Module $PSCommandPath -DisableNameChecking -Force -ErrorAction Stop
                $autoUpdateSuccess = $true
            }
        }
        catch {
            $message = "Autoupdate failed. $_"
        }
        finally {
            Write-Progress -Activity "AutoUpdate" -Status "done" -Completed
        }
    }

    [PSCustomObject]@{
        Success    = $autoUpdateSuccess
        Message    = $message
        NewVersion = $release.name
    }
}

function Start-Wpr {
    [CmdletBinding(PositionalBinding = $false)]
    param(
        # Path to store temporary trace files
        [string]$Path,
        [ValidateSet('GeneralProfile', 'CPU', 'DiskIO', 'FileIO', 'Registry', 'Network', 'Heap', 'Pool', 'VirtualAllocation', 'Audio', 'Video', 'Power', 'InternetExplorer', 'EdgeBrowser', 'Minifilter', 'GPU', 'Handle', 'XAMLActivity', 'HTMLActivity', 'DesktopComposition', 'XAMLAppResponsiveness', 'HTMLResponsiveness', 'ReferenceSet', 'ResidentSet', 'XAMLHTMLAppMemoryAnalysis', 'UTC', 'DotNET', 'WdfTraceLoggingProvider', 'HeapSnapshot')]
        [string[]]$Profiles = @('GeneralProfile', 'CPU', 'DiskIO', 'FileIO', 'Registry', 'Network')
    )

    # wpr is available on Win10 and above
    if (-not (Get-Command 'wpr.exe' -ErrorAction SilentlyContinue)) {
        Write-Error "WPR is not available on this machine"
        return
    }

    if ($PSBoundParameters.ContainsKey('Path')) {
        if (-not (Test-Path $Path)) {
            $null = New-Item $Path -ItemType Directory -ErrorAction Stop
        }

        $Path = Convert-Path -LiteralPath $Path
    }

    $wprArgs = @(
        foreach ($prof in $Profiles) {
            '-start', $prof
        }

        '-filemode'
    )

    if ($Path) {
        # For some reason, if the path contains a space & is double-quoted & ends with a backslash, wpr fails with "Invalid temporary trace directory. Error code:0xc5586004"
        # Make sure to remove the last backslash.
        if ($Path.EndsWith('\')) {
            $Path = $Path.Remove($Path.Length - 1)
        }

        $wprArgs += '-RecordTempTo', $Path
    }

    Write-Log "Invoking wpr.exe $wprArgs"

    $errorMsg = $($null = wpr.exe $wprArgs) 2>&1 | & {
        process {
            $_.Exception.Message.Trim()
        }
    }

    if ($LASTEXITCODE -ne 0) {
        Write-Error "wpr failed to start. LASTEXITCODE:0x$('{0:x}' -f $LASTEXITCODE).$errorMsg"
    }
}

function Stop-Wpr {
    [CmdletBinding(PositionalBinding = $false)]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,
        [string]$FileName = 'wpr.etl'
    )

    # wpr is available on Win10 and above
    if (-not (Get-Command 'wpr.exe' -ErrorAction SilentlyContinue)) {
        Write-Error "WPR is not available on this machine"
        return
    }

    if (-not (Test-Path $Path)) {
        $null = New-Item $Path -ItemType Directory -ErrorAction Stop
    }

    $Path = Convert-Path -LiteralPath $Path
    $filePath = Join-Path $Path $FileName

    if ($filePath.IndexOf(' ') -gt 0) {
        $filePath = "`"$filePath`""
    }

    Write-Log "Stopping WPR trace"

    # Here Start-Process is used to suppress progress bar written by wpr.exe.
    $startProcArgs = @{
        FilePath     = 'wpr.exe'
        ArgumentList = @('-stop', $filePath, '-skipPdbGen')
        WindowStyle  = 'Hidden'
        PassThru     = $true
        Wait         = $true
    }

    $process = Start-Process @startProcArgs
    $exitCode = $process.ExitCode
    $process.Dispose()

    # If "Invalid command syntax", retry without -skipPdbGen because the option might not be available (e.g. W2019)
    if ($exitCode -eq 0xc5600602) {
        $startProcArgs.ArgumentList = @('-stop', $filePath)
        $process = Start-Process @startProcArgs
        $exitCode = $process.ExitCode
        $process.Dispose()
    }

    if ($exitCode -ne 0) {
        Write-Error "wpr failed to stop. ExitCode:0x$('{0:x}' -f $exitCode)"
    }
}

function Get-IMProvider {
    [CmdletBinding(PositionalBinding = $false)]
    param(
        $User
    )

    $root = Get-UserRegistryRoot $User

    if (-not $root) {
        return
    }

    $defaultIMApp = Join-Path $root 'SOFTWARE\IM Providers' | Get-ItemProperty -Name 'DefaultIMApp' -ErrorAction SilentlyContinue | Select-Object -ExpandProperty 'DefaultIMApp'

    if (-not $defaultIMApp) {
        Write-Error "There is no DefaultIMApp in $defaultIMApp"
        return
    }

    [Guid]$clsid = switch ($defaultIMApp) {
        'Teams' { '00425F68-FFC1-445F-8EDF-EF78B84BA1C7'; break }
        'Lync' { 'A0651028-BA7A-4D71-877F-12E0175A5806'; break }
        'MsTeams' { '88435F68-FFC1-445F-8EDF-EF78B84BA1C7'; break }
        default { Write-Error "Failed to get CLSID of DefaultIMApp '$defaultIMApp'"; return }
    }

    # The new Teams client's executable is "ms-teams.exe".
    $exeName = switch ($defaultIMApp) {
        'MsTeams' { 'ms-teams'; break }
        default { $defaultIMApp; break }
    }

    $isRunning = $false
    $process = Get-Process -Name $exeName -ErrorAction SilentlyContinue | Select-Object -First 1

    if ($process) {
        $isRunning = $true
        $process.Dispose()
    }

    # Bail if IM application is not runnning because instantiating its COM object will start the app and can take a long time.
    if (-not $isRunning) {
        [PSCustomObject]@{
            DefaultIMApp = $defaultIMApp
            IsRunning    = $isRunning
        }

        return
    }

    $imProvider = $null
    $punk = $pIUCOfficeIntegration = [IntPtr]::Zero

    try {
        # Create a COM instance
        $createInstance = $false
        $type = [Type]::GetTypeFromCLSID($clsid)
        $imProvider = [Activator]::CreateInstance($type)
        $createInstance = $true

        # Retrieve IUnknown
        $punk = [Runtime.InteropServices.Marshal]::GetIUnknownForObject($imProvider)

        # Get IUCOfficeIntegration
        [Guid]$IID_IUCOfficeIntegration = '6a222195-f65e-467f-8f77-eb180bd85288'
        $hr = [Runtime.InteropServices.Marshal]::QueryInterface($punk, [ref]$IID_IUCOfficeIntegration, [ref]$pIUCOfficeIntegration)

        if ($hr -ne 0) {
            Write-Error $("QueryInterface for IID $IID_IUCOfficeIntegration failed with 0x{0:x}" -f $hr)
        }

        # Call IUCOfficeIntegration->GetAuthenticationInfo()
        $authInfo = $imProvider.GetAuthenticationInfo('15.0.0.0')
    }
    catch {
        if (-not $imProvider) {
            Write-Error -Message "Failed to create an instance of $defaultIMApp (CLSID:{$clsid})" -Exception $_.Exception
        }
        elseif ($pIUCOfficeIntegration -eq [IntPtr]::Zero) {
            Write-Error -Message "Failed to obtain IUCOfficeIntegration interface" -Exception $_.Exception
        }
        else {
            Write-Error -ErrorRecord $_
        }
    }
    finally {
        if ($punk -ne [IntPtr]::Zero) {
            $null = [System.Runtime.InteropServices.Marshal]::Release($punk)
        }

        if ($pIUCOfficeIntegration -ne [IntPtr]::Zero) {
            $null = [System.Runtime.InteropServices.Marshal]::Release($pIUCOfficeIntegration)
        }

        if ($imProvider) {
            $null = [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($imProvider)
        }
    }

    [PSCustomObject]@{
        DefaultIMApp         = $defaultIMApp
        IsRunning            = $isRunning
        CreateInstance       = $createInstance
        IUCOfficeIntegration = $createInstance -and $pIUCOfficeIntegration -ne [IntPtr]::Zero
        AuthenticationInfo   = $authInfo
    }
}

<#
.SYNOPSIS
    Wait and get result of WinRT Async action/operation
.DESCRIPTION
    Wait and get result of WinRT Async action/operation.
    Currently only IAsyncAction and IAsyncOperation<TResult> are supported.
.EXAMPLE
    Receive-WinRTAsyncResult $asyncAction
.EXAMPLE
    Receive-WinRTAsyncResult $asyncOperation -TResult ([Windows.Security.Authentication.Web.Core.FindAllAccountsResult])
.EXAMPLE
    $result = $asyncOperation | Receive-WinRTAsyncResult -TResult ([Windows.Security.Authentication.Web.Core.FindAllAccountsResult])
#>
function Receive-WinRTAsyncResult {
    [CmdletBinding(PositionalBinding = $false)]
    param(
        [Parameter(Mandatory, Position = 0, ValueFromPipeline)]
        # COM object representing WinRT IAsyncInfo or its derived types (e.g. IAsyncOperation<TResult>, IAsyncAction)
        [System.__ComObject]$IAsyncInfo,
        # TResult of Async Operation
        [System.Reflection.TypeInfo]$TResult
    )

    # Assume that System.Runtime.WindowsRuntime is already loaded by "Add-Type -AssemblyName System.Runtime.WindowsRuntime"

    process {
        # Convert WinRT AsyncInfo to a .NET Task using WindowsRuntimeSystemExtensions's AsTask() method.
        # https://learn.microsoft.com/en-us/dotnet/api/system.windowsruntimesystemextensions
        if ($TResult) {
            # Use AsTask<TResult>(Windows.Foundation.IAsyncOperation<TResult>)
            if (-not $Script:AsTaskOfIAsyncOperation) {
                $methodInfo = [System.WindowsRuntimeSystemExtensions].GetMethods() | Where-Object {
                    $_.Name -eq 'AsTask' `
                        -and $_.GetParameters().Count -eq 1 `
                        -and $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncOperation`1'
                } | Select-Object -First 1

                $Script:AsTaskOfIAsyncOperation = $methodInfo
            }

            # Create AsTask<TResult> method
            $asTask = $Script:AsTaskOfIAsyncOperation.MakeGenericMethod($TResult)
        }
        else {
            # Use AsTask(Windows.Foundation.IAsyncAction)
            if (-not $Scipt:AsTaskofIAsyncAction) {
                $methodInfo = [System.WindowsRuntimeSystemExtensions].GetMethods() | Where-Object {
                    $_.Name -eq 'AsTask' `
                        -and $_.GetParameters().Count -eq 1 `
                        -and $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncAction'
                } | Select-Object -First 1

                $Script:AsTaskofIAsyncAction = $methodInfo
            }

            $asTask = $Script:AsTaskofIAsyncAction
        }

        # Create a Task and wait for it to complete.
        try {
            $task = $asTask.Invoke(<# obj #> $null, <# parameters #> @($IAsyncInfo))
            $result = $task.GetAwaiter().GetResult()

            # For IAsyncAction scenario, Task can be of type Task[System.VoidValueTypeParameter]. Do not return this result.
            if ($TResult) {
                $result
            }

            # Close IAsyncInfo
            if (-not $Script:AsyncInfoClose) {
                $Script:AsyncInfoClose = [Windows.Foundation.IAsyncInfo, Windows, ContentType = WindowsRuntime].GetMethods() | Where-Object Name -eq 'Close' | Select-Object -First 1
            }

            $null = $Script:AsyncInfoClose.Invoke($IAsyncInfo, $null)
        }
        catch {
            Write-Error -ErrorRecord $_
        }
    }
}

# WAM related constants
$WAM = @{
    ProviderId = @{
        # https://learn.microsoft.com/en-us/windows/uwp/security/web-account-manager
        AAD       = 'https://login.windows.net'
        Microsoft = 'https://login.microsoft.com'
        Local     = 'https://login.windows.local'
    }

    Authority  = @{
        Consumers     = 'consumers'
        Organizations = 'organizations'
    }

    ClientId   = @{
        # https://learn.microsoft.com/en-us/troubleshoot/entra/entra-id/governance/verify-first-party-apps-sign-in
        MSOffice  = 'd3590ed6-52b3-4102-aeff-aad2292ab01c'
        MSGraph   = '00000003-0000-0000-c000-000000000000'
        MSOutlook = '5d661950-3475-41cd-a2c3-d671a3162bc1'
    }

    Resource   = @{
        Outlook = 'https://outlook.office365.com/'
    }

    Scopes     = @{
        # Note: scopes are space-delimited strings:
        # https://datatracker.ietf.org/doc/html/rfc6749#section-3.3
        Default = 'https://outlook.office365.com//.default offline_access openid profile'
    }
}

<#
.SYNOPSIS
    Sign out of accounts persisted in WAM (Web Account Manager).
.DESCRIPTION
    This command invokes WebAuthenticationCoreManager's SignOutAsync() method to sign out of accounts that persist in WAM.
    When it finds account(s), it asks the user if s/he wants to sign out of each account, unless Force switch is specified in which case it automatically sign out of all accounts.
.EXAMPLE
    Invoke-WamSignOut
    By default, sign out of account associated with MS Office Client ID (d3590ed6-52b3-4102-aeff-aad2292ab01c).
.EXAMPLE
    Invoke-WamSignOut -Force
    Sign out of all the accounts found without user interaction.

.LINK
    WebAccount.SignOutAsync Method
    https://docs.microsoft.com/en-us/uwp/api/windows.security.credentials.webaccount.signoutasync
#>
function Invoke-WamSignOut {
    [CmdletBinding(PositionalBinding = $false)]
    param(
        # Application/Client ID
        [string]$ClientId,
        # Sign out of all the accounts without user interaction.
        [switch]$Force
    )

    if (-not ('Windows.Foundation.Metadata.ApiInformation' -as [type]) -or -not [Windows.Foundation.Metadata.ApiInformation, Windows, ContentType = WindowsRuntime]::IsMethodPresent('Windows.Security.Authentication.Web.Core.WebAuthenticationCoreManager', 'FindAllAccountsAsync')) {
        Write-Error "This script is not supported on this Windows version. You can use https://github.com/jpazureid/WPJ-cleanup-tool/blob/master/CleanupTool.zip"
        return
    }

    Add-Type -AssemblyName System.Runtime.WindowsRuntime
    $MS_OFFICE_CLIENTID = 'd3590ed6-52b3-4102-aeff-aad2292ab01c'

    # By default use MS Office Client ID.
    if (-not $ClientId) {
        $ClientId = $MS_OFFICE_CLIENTID
    }

    $provider = [Windows.Security.Authentication.Web.Core.WebAuthenticationCoreManager, Windows, ContentType = WindowsRuntime]::FindAccountProviderAsync('https://login.microsoft.com', 'organizations') `
    | Receive-WinRTAsyncResult -TResult ([Windows.Security.Credentials.WebAccountProvider, Windows, ContentType = WindowsRuntime])

    $findAllAccountsResult = [Windows.Security.Authentication.Web.Core.WebAuthenticationCoreManager, Windows, ContentType = WindowsRuntime]::FindAllAccountsAsync($provider, $ClientId) `
    | Receive-WinRTAsyncResult -TResult ([Windows.Security.Authentication.Web.Core.FindAllAccountsResult, Windows, ContentType = WindowsRuntime])

    $count = $findAllAccountsResult.Accounts | Measure-Object | Select-Object -ExpandProperty Count

    if ($count -eq 0) {
        Write-Log "No account found"
        return
    }

    Write-Log "$count account$(if ($count -gt 1) {'s'}) found"

    foreach ($account in $findAllAccountsResult.Accounts) {
        $accountId = "UserName:$($account.UserName), Id:$($account.Id)"
        $state = $account.State
        Write-Log "Account $accountId's State is $state"

        $signOutMsg = "Signing out an account; $accountId"

        # If Force is not specified, ask the user
        if (-not $Force) {
            $ans = Read-Host "Do you want to sign out the following account? (Y|N)`n  $accountId"

            if ($ans -like 'Y*') {
                Write-Host $signOutMsg -ForegroundColor Green
            }
            else {
                # Skip this account
                continue
            }
        }

        Write-Log $signOutMsg
        $account.SignOutAsync($ClientId) | Receive-WinRTAsyncResult
    }
}

<#
.SYNOPSIS
    Get a Web Account Provider.
.DESCRIPTION
    Get a Web Account Provider using WebAuthenticationCoreManager.FindAccountProviderAsync:
.LINK
    WebAuthenticationCoreManager.FindAccountProviderAsync
    https://learn.microsoft.com/en-us/uwp/api/windows.security.authentication.web.core.webauthenticationcoremanager.findaccountproviderasync?view=winrt-26100
#>
function Get-WebAccountProvider {
    [CmdletBinding(PositionalBinding = $false)]
    param(
        # The Id of the web account provider to find.
        [ValidateSet('https://login.windows.net', 'https://login.microsoft.com', 'https://login.windows.local')]
        [string]$ProviderId = $WAM.ProviderId.Microsoft,
        # The authority of the web account provider to find.
        [ValidateSet('organizations', 'consumers')]
        [string]$Authority = $WAM.Authority.Organizations
    )

    Add-Type -AssemblyName System.Runtime.WindowsRuntime

    [Windows.Security.Authentication.Web.Core.WebAuthenticationCoreManager, Windows, ContentType = WindowsRuntime]::FindAccountProviderAsync($ProviderId, $Authority) `
    | Receive-WinRTAsyncResult -TResult ([Windows.Security.Credentials.WebAccountProvider, Windows, ContentType = WindowsRuntime])
}

<#
.SYNOPSIS
    Get Web Accounts
.DESCRIPTION
    Get Web Accounts using WebAuthenticationCoreManager.FindAllAccountsAsync()
.LINK
    WebAuthenticationCoreManager.FindAllAccountsAsync Method
    https://learn.microsoft.com/en-us/uwp/api/windows.security.authentication.web.core.webauthenticationcoremanager.findallaccountsasync

.EXAMPLE
    Test-MyTestFunction -Verbose
    Explanation of the function or its result. You can include multiple examples with additional .EXAMPLE lines
#>
function Get-WebAccount {
    [CmdletBinding(PositionalBinding = $false)]
    param(
        # Web Account Provider ID
        [ValidateSet('https://login.windows.net', 'https://login.microsoft.com', 'https://login.windows.local')]
        [string]$ProviderId = $WAM.ProviderId.Microsoft,
        # The authority of the web account provider
        [ValidateSet('organizations', 'consumers')]
        [string]$Authority = $WAM.Authority.Organizations,
        [string]$ClientId = $WAM.ClientId.MSOffice
    )

    Add-Type -AssemblyName System.Runtime.WindowsRuntime

    $provider = Get-WebAccountProvider -ProviderId $ProviderId -Authority $Authority

    $findAllAccountsResult = [Windows.Security.Authentication.Web.Core.WebAuthenticationCoreManager, Windows, ContentType = WindowsRuntime]::FindAllAccountsAsync($provider, $ClientId) `
    | Receive-WinRTAsyncResult -TResult ([Windows.Security.Authentication.Web.Core.FindAllAccountsResult, Windows, ContentType = WindowsRuntime])

    if ($findAllAccountsResult.Status -ne [Windows.Security.Authentication.Web.Core.FindAllWebAccountsStatus]::Success) {
        Write-Error "FindAllAccountsAsync() failed with `"$($findAllAccountsResult.Status)`". ErrorCode:0x$("{0:x8}" -f $findAllAccountsResult.ProviderError.ErrorCode), ErrorMessage:$($findAllAccountsResult.ProviderError.ErrorMessage)"
        return
    }

    $findAllAccountsResult.Accounts
}

<#
.SYNOPSIS
    Invoke WebAccount.SignOutAsync()

.LINK
    WebAccount.SignOutAsync Method
    https://learn.microsoft.com/en-us/uwp/api/windows.security.credentials.webaccount.signoutasync

.EXAMPLE
    Get-WebAccount | Invoke-WebAccountSignOut
#>
function Invoke-WebAccountSignOut {
    [CmdletBinding(PositionalBinding = $false)]
    param(
        [Parameter(Mandatory, ValueFromPipeline, Position = 0)]
        $WebAccount
    )

    process {
        Write-Log "Invoking SignOutAsync() for WebAccount:$($WebAccount.Id)"
        $WebAccount.SignOutAsync() | Receive-WinRTAsyncResult
    }
}

<#
.SYNOPSIS
    Get a token using WebAuthenticationCoreManager.GetTokenSilentlyAsync()
.DESCRIPTION
    Get a token using WebAuthenticationCoreManager.GetTokenSilentlyAsync().
    If the Token is a JSON Web Token (JWT), its decoded Header & Payload are included in the output.
.LINK
    WebAuthenticationCoreManager.GetTokenSilentlyAsync Method
    https://learn.microsoft.com/en-us/uwp/api/windows.security.authentication.web.core.webauthenticationcoremanager.gettokensilentlyasync
.EXAMPLE
    Get-TokenSilently
.EXAMPLE
    Get-TokenSilently -Resource 'https://outlook.office365.com' -AddClaimCapability
.EXAMPLE
    Get-TokenSilently -Resource 'https://outlook.office365.com' -IncludeRawToken
#>
function Get-TokenSilently {
    [CmdletBinding()]
    param(
        # Web Account Provider ID
        [ValidateSet('https://login.windows.net', 'https://login.microsoft.com', 'https://login.windows.local')]
        [string]$ProviderId = $WAM.ProviderId.Microsoft,
        # The authority of the web account provider
        [ValidateSet('organizations', 'consumers')]
        [string]$Authority = $WAM.Authority.Organizations,
        # Application/Cliend ID (By default, MSOffice 'd3590ed6-52b3-4102-aeff-aad2292ab01c')
        [string]$ClientId = $WAM.ClientId.MSOffice,
        # Scopes are space-delimited strings:
        # https://datatracker.ietf.org/doc/html/rfc6749#section-3.3
        # e.g. "https://outlook.office365.com//.default offline_access openid profile"
        [string]$Scopes,
        # e.g. 'https://outlook.office365.com', 'https://graph.windows.net'
        [string]$Resource,
        # Add "wam_compat=2.0" to request
        [Switch]$AddWamCompat,
        # Add "claim={"access_token":{"xms_cc":{"values":["CP1"]}}}" to request
        [Switch]$AddClaimCapability,
        # You can use Get-WebAccount command to get a web account
        $WebAccount,
        # Include raw token in the output
        [Switch]$IncludeRawToken
    )

    Write-Log "Get-TokenSilently() called with ProviderId:$ProviderId, Authority:$Authority, ClientId:$ClientId, Scopes:$Scopes, Resource:$Resource, AddWamCompat:$AddWamCompat, AddClaimCapability:$AddClaimCapability, IncludeRawToken:$IncludeRawToken, WebAccount:$($WebAccount.Id)"

    $provider = Get-WebAccountProvider -ProviderId $ProviderId -Authority $Authority

    $promptType = [Windows.Security.Authentication.Web.Core.WebTokenRequestPromptType]::Default
    $request = [Windows.Security.Authentication.Web.Core.WebTokenRequest, Windows, ContentType = WindowsRuntime]::new($provider, $Scopes, $ClientId, $promptType)

    if ($null -eq $request.Properties) {
        Write-Error "WebTokenRequest.Properties is null. Why?"
        return
    }

    # IDictionary<string, string>'s Add() method for request.Properties (exposed as System.__ComObject)
    $addMethod = [System.Collections.Generic.IDictionary[string, string]].GetMethods() | Where-Object { $_.Name -eq 'Add' } | Select-Object -First 1

    If ($AddWamCompat) {
        $null = $addMethod.Invoke($request.Properties, @('wam_compat', '2.0'))
    }

    if ($Resource) {
        $null = $addMethod.Invoke($request.Properties, @('resource', $Resource))
    }

    if ($AddClaimCapability) {
        $null = $addMethod.Invoke($request.Properties, @('claims', '{"access_token":{"xms_cc":{"values":["CP1"]}}}'))
    }

    Write-Log "WebTokenRequest.Properties: $($request.Properties)"

    if ($WebAccount) {
        Write-Log "Invoking GetTokenSilentlyAsync() with WebAccount:$($WebAccount.Id)"

        $requestResult = [Windows.Security.Authentication.Web.Core.WebAuthenticationCoreManager, Windows, ContentType = WindowsRuntime]::GetTokenSilentlyAsync($request, $WebAccount) `
        | Receive-WinRTAsyncResult -TResult ([Windows.Security.Authentication.Web.Core.WebTokenRequestResult, Windows, ContentType = WindowsRuntime])
    }
    else {
        Write-Log "Invoking GetTokenSilentlyAsync() without WebAccount"

        $requestResult = [Windows.Security.Authentication.Web.Core.WebAuthenticationCoreManager, Windows, ContentType = WindowsRuntime]::GetTokenSilentlyAsync($request) `
        | Receive-WinRTAsyncResult -TResult ([Windows.Security.Authentication.Web.Core.WebTokenRequestResult, Windows, ContentType = WindowsRuntime])
    }

    if ($requestResult.ResponseStatus -ne [Windows.Security.Authentication.Web.Core.WebTokenRequestStatus]::Success) {
        Write-Error "GetTokenSilentlyAsync() failed with `"$($requestResult.ResponseStatus)`". ErrorCode:0x$("{0:x8}" -f $requestResult.ResponseError.ErrorCode), ErrorMessage:$($requestResult.ResponseError.ErrorMessage)"
        return
    }

    # Note: Do not use "$requestResult.ResponseData.Properties". It'd cause request.Properties to be null in the next invocation.
    foreach ($_ in $requestResult.ResponseData) {
        $result = [ordered]@{
            WebAccount = $_.WebAccount
            # Properties = [PSCustomObject]$props
        }

        if ($IncludeRawToken) {
            $result.Token = $_.Token
        }

        # If token is a JSON Web Token (JWT), decode it.
        $tokenParts = $_.Token.Split('.')

        if ($tokenParts.Count -eq 3) {
            $header = $tokenParts[0]
            $payload = $tokenParts[1]
            $result.JwtHeader = [System.Text.Encoding]::UTF8.GetString((Convert-Base64Url $header))
            $result.JwtPayload = [System.Text.Encoding]::UTF8.GetString((Convert-Base64Url $payload))
        }

        # Sice Properties is a System.__COMObject, pack them into a hash table
        $props = @{}

        foreach ($prop in $_.Properties) {
            $props.Add($prop.Key, $prop.Value)
        }

        $result.Properties = [PSCustomObject]$props

        [PSCustomObject]$result
    }
}

function Invoke-RequestToken {
    [CmdletBinding()]
    param(
        # Web Account Provider ID
        [ValidateSet('https://login.windows.net', 'https://login.microsoft.com', 'https://login.windows.local')]
        [string]$ProviderId = $WAM.ProviderId.Microsoft,
        # The authority of the web account provider
        [ValidateSet('organizations', 'consumers')]
        [string]$Authority = $WAM.Authority.Organizations,
        # Application/Cliend ID (By default, MSOffice 'd3590ed6-52b3-4102-aeff-aad2292ab01c')
        [string]$ClientId = $WAM.ClientId.MSOffice,
        # Scopes are space-delimited strings:
        # https://datatracker.ietf.org/doc/html/rfc6749#section-3.3
        # e.g. "https://outlook.office365.com//.default offline_access openid profile"
        [string]$Scopes,
        # e.g. 'https://outlook.office365.com', 'https://graph.windows.net'
        [string]$Resource,
        # Add "wam_compat=2.0" to request
        [Switch]$AddWamCompat,
        # Add "claim={"access_token":{"xms_cc":{"values":["CP1"]}}}" to request
        [Switch]$AddClaimCapability,
        # You can use Get-WebAccount command to get a web account
        $WebAccount,
        # Include raw token in the output
        $IncludeRawToken
    )

    $interopDllPath = Join-Path $env:TEMP 'WamInterop.dll'
    $hModule = [IntPtr]::Zero

    $currentProc = Get-Process -Id $PID
    $wamInteropDll = $currentProc.Modules | Where-Object { $_.FileName -eq $interopDllPath } | Select-Object -First 1
    $currentProc.Dispose()

    if (-not $wamInteropDll) {
        # Write WamInterop.dll
        $err = Save-WamInteropDll -Path $interopDllPath 2>&1

        if ($err) {
            Write-Error "Failed to save WamInterop.dll to $interopDllPath. $err"
            return
        }

        # I could add to $env:PATH but let's explicitly load it
        $hModule = [Win32.Kernel32]::LoadLibraryW($interopDllPath)

        if ($hModule -eq [IntPtr]::Zero) {
            Write-Error "Failed to load WamInterop.dll. ErrorCode:0x$("{0:x}" -f [System.Runtime.InteropServices.Marshal]::GetLastWin32Error())"
            return
        }
    }

    # Make sure that interop DLL is located at $PSScriptRoot
    # $WamInteropDllName = 'WamInterop.dll'
    # $interopDll = Get-ChildItem -Path $PSScriptRoot -Filter $WamInteropDllName -ErrorAction SilentlyContinue

    # if (-not $interopDll) {
    #     Write-Error "$WamInteropDllName is not found at $PSScriptRoot"
    #     return
    # }

    $provider = Get-WebAccountProvider -ProviderId $ProviderId -Authority $Authority
    $promptType = [Windows.Security.Authentication.Web.Core.WebTokenRequestPromptType]::ForceAuthentication
    $request = [Windows.Security.Authentication.Web.Core.WebTokenRequest, Windows, ContentType = WindowsRuntime]::new($provider, $Scopes, $ClientId, $promptType)

    if ($null -eq $request.Properties) {
        Write-Error "WebTokenRequest.Properties is null. Why?"
        return
    }

    # IDictionary<string, string>'s Add() method for request.Properties (exposed as System.__ComObject)
    $addMethod = [System.Collections.Generic.IDictionary[string, string]].GetMethods() | Where-Object { $_.Name -eq 'Add' } | Select-Object -First 1

    If ($AddWamCompat) {
        $null = $addMethod.Invoke($request.Properties, @('wam_compat', '2.0'))
    }

    if ($Resource) {
        $null = $addMethod.Invoke($request.Properties, @('resource', $Resource))
    }

    if ($AddClaimCapability) {
        $null = $addMethod.Invoke($request.Properties, @('claims', '{"access_token":{"xms_cc":{"values":["CP1"]}}}'))
    }

    $runSpaceOpened = $false
    $requestResult = $null

    try {
        if (-not $Script:RunspacePool) {
            Open-TaskRunspace
            $runSpaceOpened = $true
        }

        # Add-EnvPath $PSScriptRoot

        [IntPtr]$hwnd = [Win32.WamInterop]::CreateAnchorWindow()

        if ($hwnd -eq [IntPtr]::Zero) {
            Write-Error "CreateAnchorWindow failed"
            return
        }

        # RequestToken must be invoked on a different thread. This thread needs to process the anchor window msg loop.
        $task = Start-Task {
            param ($hwnd, $request)
            $ptr = [IntPtr]::Zero
            $requestPtr = [System.Runtime.InteropServices.Marshal]::GetIUnknownForObject($request)

            $hr = [Win32.WamInterop]::RequestToken($hwnd, $requestPtr, [ref]$ptr)

            $null = [System.Runtime.InteropServices.Marshal]::Release($requestPtr)

            if ($hr -ne 0 <# S_OK #>) {
                Write-Error "RequestToken() failed with 0x$("{0:x}" -f $hr)"
                return
            }

            [System.Runtime.InteropServices.Marshal]::GetObjectForIUnknown($ptr)
            $null = [System.Runtime.InteropServices.Marshal]::Release($ptr)
        } -ArgumentList $hwnd, $request

        $requestResult = $task | Receive-Task -AutoRemoveTask

    }
    catch {
        Write-Error -ErrorRecord $_
    }
    finally {
        if ($hwnd) {
            # Destroy anchor window
            $WM_DESTROY = 2
            $SMTO_ABORTIFHUNG = 2
            $Timeout = [TimeSpan]::FromSeconds(5)
            $result = [IntPtr]::Zero
            $ret = [Win32.User32]::SendMessageTimeoutW($hWnd, $WM_DESTROY, [IntPtr]::Zero, [IntPtr]::Zero, $SMTO_ABORTIFHUNG, $Timeout.TotalMilliseconds, [ref]$result)

            # > If the function succeeds, the return value is nonzero.
            if ($ret -ne 0) {
                Write-Log "Anchor Window is closed successfully"
            }
        }

        if ($runSpaceOpened) {
            Close-TaskRunspace
        }

        # Note that even if the interop DLL is unloaded here, C# interop will keep the module loaded.
        if ($hMdoule -ne [IntPtr]::Zero) {
            $null = [Win32.Kernel32]::FreeLibrary($hModule)
        }
    }

    if (-not $requestResult) {
        return
    }

    if ($requestResult.ResponseStatus -ne [Windows.Security.Authentication.Web.Core.WebTokenRequestStatus]::Success) {
        Write-Error "RequestTokenAsync() failed with `"$($requestResult.ResponseStatus)`". $(if ($requestResult.ResponseError) { "ErrorCode:0x$("{0:x8}" -f $requestResult.ResponseError.ErrorCode), ErrorMessage:$($requestResult.ResponseError.ErrorMessage)"})"
        return
    }


    foreach ($_ in $requestResult.ResponseData) {
        $result = [ordered]@{
            WebAccount = $_.WebAccount
        }

        if ($IncludeRawToken) {
            $result.Token = $_.Token
        }

        # If token is a JSON Web Token (JWT), decode it.
        $tokenParts = $_.Token.Split('.')

        if ($tokenParts.Count -eq 3) {
            $header = $tokenParts[0]
            $payload = $tokenParts[1]
            $result.JwtHeader = [System.Text.Encoding]::UTF8.GetString((Convert-Base64Url $header))
            $result.JwtPayload = [System.Text.Encoding]::UTF8.GetString((Convert-Base64Url $payload))
        }

        # Sice Properties is a System.__COMObject, pack them into a hash table
        $props = @{}

        foreach ($prop in $_.Properties) {
            $props.Add($prop.Key, $prop.Value)
        }

        $result.Properties = [PSCustomObject]$props

        [PSCustomObject]$result
    }
}

function Add-EnvPath {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Path
    )

    $found = $env:Path -split ';' | Where-Object { $_.TrimEnd() -eq $Path } | Select-Object -First 1

    if ($found) {
        return
    }

    $sb = New-Object System.Text.StringBuilder -ArgumentList $Path

    if (-not $env:Path.EndsWith(';')) {
        $null = $sb.Insert(0, ';')
    }

    $env:Path += $sb.ToString()
}

function Save-WamInteropDll {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        # Destination file path
        [string]$Path
    )

    try {
        $bytes = [Convert]::FromBase64String($Script:WamInteropBytes)
        [System.IO.File]::WriteAllBytes($Path, $bytes)
    }
    catch {
        Write-Error -ErrorRecord $_
    }
}

function Convert-Base64Url {
    [CmdletBinding()]
    [OutputType([byte[]])]
    param(
        [Parameter(Mandatory, ValueFromPipeline)]
        [string]$Base64Url
    )

    process {
        # Replace '-' with '+' and '_' with '/'
        $sb = New-Object System.Text.StringBuilder -ArgumentList $Base64Url
        $null = $sb.Replace('-', '+').Replace('_', '/')

        # Pad the string to make it a multiple of 4
        switch ($sb.Length % 4) {
            0 { break }
            2 { $null = $sb.Append('=='); break }
            3 { $null = $sb.Append('='); break }
            default { Write-Error "Invalid Base64Url string"; return }
        }

        [System.Convert]::FromBase64String($sb.ToString())
    }
}

<#
.SYNOPSIS
Helper command to recursively get registry key and its values.

.DESCRIPTION
Output object has the following properties:

- "KeyName"
- "Properties":PSCustomObject that contains key's properties (i.e. key values)
- Sub keys

Each subkey becomes a property.
e.g. For the following 'Outlook' key,

Outlook
    |- Profiles

The output object will have a property called "Profiles".
#>
function Get-RegistryChildItem {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        # Registry path
        [string]$Path
        # [switch]$IncludeRawItemProperty
    )

    if (-not (Test-Path $Path)) {
        Write-Error "Cannot find $Path"
        return
    }

    # Stack contains a hash table with "Parent" of type PSCustomObject and "Key" of type Microsoft.Win32.RegistryKey
    $stack = New-Object System.Collections.Generic.Stack[object]
    $stack.Push(@{Parent = $null; Key = (Get-Item $Path) })
    $root = $null

    while ($stack.Count -gt 0) {
        $node = $stack.Pop()
        $key = $node.Key
        $parent = $node.Parent

        $props = $key | Get-ItemProperty

        $obj = [PSCustomObject]@{
            PSPath     = $key.PSPath # $key.PSPath.SubString('Microsoft.PowerShell.Core\Registry::'.Length)
            Properties = $props
        }

        # Connect to its parent if exits; otherwise this is the root.
        if ($parent) {
            $(Add-Member -InputObject $parent -MemberType NoteProperty -Name $key.PSChildName -Value $obj) 2>&1 | Write-Log
        }
        else {
            $root = $obj
        }

        # Add child nodes with parent being the current object
        Get-ChildItem $key.PSPath | & {
            process { $stack.Push(@{Parent = $obj; Key = $_ }) }
        }

        $key.Dispose()
    }

    $root
}

function Get-OfficeIdentityConfig {
    [CmdletBinding()]
    param(
        [string]$User
    )

    $userRegRoot = Get-UserRegistryRoot $User

    if (-not $userRegRoot) {
        return
    }

    Join-Path $userRegRoot 'Software\Microsoft\Office\16.0\Common\Identity' `
    | Get-ItemProperty -ErrorAction SilentlyContinue `
    | Select-Object -Property '*' -ExcludeProperty 'PSParentPath', 'PSChildName', 'PSProvider'
}

function Get-OfficeIdentity {
    [CmdletBinding()]
    param(
        [string]$User
    )

    $userRegRoot = Get-UserRegistryRoot $User

    if (-not $userRegRoot) {
        return
    }

    $IdpMapping = @{
        0 = 'Unknown'
        1 = 'LiveId'
        2 = 'OrgId'
        3 = 'ActiveDirectory'
        4 = 'ADAL'
        5 = 'SSPI'
        6 = 'OAuth2'
        7 = 'Badger'
    }

    # Get the Office Identities
    $identities = Join-Path $userRegRoot 'Software\Microsoft\Office\16.0\Common\Identity\Identities\*' | Get-ItemProperty -ErrorAction SilentlyContinue

    if (-not $identities) {
        Write-Log "Cannot find Office Identities"
        return
    }

    # Add LastSwitchedTime in profile if avaialble.
    foreach ($id in $identities) {
        $lastSwitchedTime = Join-Path $userRegRoot 'Software\Microsoft\Office\16.0\Common\Identity\Profiles' | Join-Path -ChildPath $id.PSChildName `
        | Get-ItemProperty -Name 'LastSwitchedTime' -ErrorAction SilentlyContinue | Select-Object -ExpandProperty 'LastSwitchedTime'

        if ($lastSwitchedTime) {
            $id | Add-Member -NotePropertyName 'LastSwitchedTime' -NotePropertyValue $lastSwitchedTime
        }
    }

    # Find the one with the latest LastSwitchedTime if any.
    $activeIdentity = $identities | Where-Object { $_.SignedOut -ne 1 -and $_.LastSwitchedTime -and $_.LastSwitchedTime -ne '1601-01-01T00:01:00Z' } `
    | Sort-Object 'LastSwitchedTime' -Descending | Select-Object -First 1

    if ($activeIdentity) {
        Write-Log "Found active identity $($activeIdentity.EmailAddress) based on active profile"
    }
    else {
        # If there is no active profile, then pick one with LiveId, OrgId, or ADAL
        $activeIdentity = $identities `
        | Where-Object { $_.SignedOut -ne 1 -and ($IdpMapping[$_.IdP] -eq 'LiveId' -or $IdpMapping[$_.IdP] -eq 'OrgId' -or $IdpMapping[$_.IdP] -eq 'ADAL') } `
        | Select-Object -First 1

        if ($activeIdentity) {
            Write-Log "Found active identity $($activeIdentity.EmailAddress) based on IdP $($IdpMapping[$activeIdentity.IdP])"
        }
        else {
            $activeIdentity = $identities | Where-Object { $_.SignedOut -ne 1 } | Select-Object -First 1

            if ($activeIdentity) {
                Write-Log "Found active identity $($activeIdentity.EmailAddress) based on not SignedOut"
            }
            else {
                Write-Log "There is no active identity"
            }
        }
    }

    foreach ($identity in $identities) {
        [PSCustomObject]@{
            Profile          = $identity.PSChildName
            LastSwitchedTime = $identity.LastSwitchedTime
            IsActive         = $identity -eq $activeIdentity
            EmailAddress     = $identity.EmailAddress
            FriendlyName     = $identity.FriendlyName
            HomeTenantId     = $identity.HomeTenantId
            SigninName       = $identity.SigninName
            IdP              = $IdpMapping[$identity.IdP]
            SignedOut        = $identity.SignedOut -eq 1
            # Persisted      = $identity.Persisted -eq 1
        }
    }
}

# Get ConnectedExperience state from the Office's roaming settings.
function Get-ConnectedExperience {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Identity
    )

    if (-not $userRegRoot) {
        $userRegRoot = Get-UserRegistryRoot

        if (-not $userRegRoot) {
            return
        }
    }

    # Check Office's roaming settings.
    $roamingSettingsPath = Join-Path $userRegRoot 'Software\Microsoft\Office\16.0\Common\Roaming\Identities' `
    | Join-Path -ChildPath $Identity | Join-Path -ChildPath 'Settings\1272\{00000000-0000-0000-0000-000000000000}'

    if (-not (Test-Path $roamingSettingsPath)) {
        Write-Log "Cannot find roaming settings for $Identity" -Category Warning
        return
    }

    # If PendingChanges is available, use it.
    $pendingChanges = Join-Path $roamingSettingsPath 'PendingChanges' | Get-ItemProperty -Name 'Data' -ErrorAction SilentlyContinue

    if ($pendingChanges) {
        $data = $pendingChanges.Data
    }
    else {
        $roamingSetings = Get-ItemProperty $roamingSettingsPath -Name 'Data' -ErrorAction SilentlyContinue
        $data = $roamingSetings.Data
    }

    if ($data) {
        $value = [BitConverter]::ToInt32($data, 0)
    }
    else {
        Write-Log "There is no roaming data for $Identity"
    }

    # 1 == Enabled, 2 == Disabled
    switch ($value) {
        1 { $enabled = $true; break; }
        2 { $enabled = $false; break; }
        default { $enabled = $null }
    }

    [PSCustomObject]@{
        Enabled        = $enabled
        PendingChanges = $null -ne $pendingChanges
    }
}

function Get-OneAuthAccount {
    [CmdletBinding()]
    param(
        [string]$User
    )

    $localAppdata = Get-UserShellFolder -User $User -ShellFolderName 'Local AppData'

    if (-not $localAppdata) {
        return
    }

    Join-Path $localAppdata 'Microsoft\OneAuth\accounts\*' `
    | Get-ChildItem -ErrorAction SilentlyContinue -File -Exclude 'OneAuthAccounts.zip' | & {
        param([Parameter(ValueFromPipeline)]$file)
        process {
            try {
                Get-Content $file.FullName -Encoding UTF8 | ConvertFrom-Json
            }
            catch {
                Write-Error -Message "Failed to parse $($file.FullName)" -Exception $_.Exception
            }
        }
    }
}

function Save-OneAuthAccount {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Path,
        [string]$User
    )

    $localAppdata = Get-UserShellFolder -User $User -ShellFolderName 'Local AppData'

    if (-not $localAppdata) {
        return
    }

    $src = Join-Path $localAppdata 'Microsoft\OneAuth\accounts'
    Save-Item -Path $src -Destination $Path
}

function Remove-OneAuthAccount {
    [CmdletBinding()]
    param(
        [string]$User
    )

    $localAppdata = Get-UserShellFolder -User $User -ShellFolderName 'Local AppData'

    if (-not $localAppdata) {
        return
    }

    Join-Path $localAppdata 'Microsoft\OneAuth\accounts\*' | Remove-Item
}

<#
Use policy settings to manage privacy controls for Microsoft 365 Apps for enterprise
https://learn.microsoft.com/en-us/deployoffice/privacy/manage-privacy-controls

| Policy setting                                                                     | Registry setting                   | Values                          |
| ---------------------------------------------------------------------------------- | ---------------------------------- | ------------------------------- |
| Configure the level of client software diagnostic data sent by Office to Microsoft | SendTelemetry                      | 1=Required 2=Optional 3=Neither |
| Allow the use of connected experiences in Office that analyze content              | UserContentDisabled                | 1=Enabled 2=Disabled            |
| Allow the use of connected experiences in Office that download online content      | DownloadContentDisabled            | 1=Enabled 2=Disabled            |
| Allow the use of additional optional connected experiences in Office               | ControllerConnectedServicesEnabled | 1=Enabled 2=Disabled            |
| Allow the use of connected experiences in Office                                   | DisconnectedState                  | 1=Enabled 2=Disabled            |
#>
function Get-PrivacyPolicy {
    [CmdletBinding()]
    param(
        $User
    )

    $userRegRoot = Get-UserRegistryRoot -User $User

    if (-not $userRegRoot) {
        return
    }

    $privacyPolicyPath = Join-Path $userRegRoot 'Software\Policies\Microsoft\office\16.0\common\privacy'

    if (Test-Path $privacyPolicyPath) {
        $privacyPolicy = Get-ItemProperty $privacyPolicyPath -ErrorAction SilentlyContinue

        # Convert 1 -> $true, 2 -> $false, else $null
        $converter = {
            param($value)
            switch ($value) {
                1 { $true; break }
                2 { $false; break }
                default { $null }
            }
        }

        [PSCustomObject]@{
            ConnectedExperiencesEnabled        = & $converter $privacyPolicy.DisconnectedState
            # "Additional connected experiences"
            ControllerConnectedServicesEnabled = & $converter $privacyPolicy.ControllerConnectedServicesEnabled
            # Despite the name, 2 means disabled.
            DownloadedContentEnabled           = & $converter $privacyPolicy.DownloadContentDisabled
            UserContentEnabled                 = & $converter $privacyPolicy.UserContentDisabled
            Path                               = $privacyPolicyPath.Substring(10)
        }
    }
}

<#
.SYNOPSIS
Get Alternate ID support configuration.
https://docs.microsoft.com/en-us/windows-server/identity/ad-fs/operations/configuring-alternate-login-id#step-3-configure-registry-for-impacted-users-using-group-policy
#>
function Get-AlternateId {
    [CmdletBinding()]
    param(
        [string]$User
    )

    $userRegRoot = Get-UserRegistryRoot $User

    if (-not $userRegRoot) {
        return
    }

    $domainHint = Join-Path $userRegRoot 'Software\Microsoft\AuthN' `
    | Get-ItemProperty -Name 'DomainHint' -ErrorAction SilentlyContinue `
    | Select-Object -ExpandProperty 'DomainHint'

    $enableAlternateIdSupport = Join-Path $userRegRoot 'Software\Microsoft\Office\16.0\Common\Identity' `
    | Get-ItemProperty -Name 'EnableAlternateIdSupport' -ErrorAction SilentlyContinue `
    | Select-Object -ExpandProperty 'EnableAlternateIdSupport'

    if ($domainHint) {
        # Split (e.g., "fs.test.contoso.com" -> "contoso.com" & "fs.test")
        $domains = $domainHint.Split('.')

        if ($domains.Count -ge 2) {
            $root = "$($domains[$domains.Count - 2]).$($domains[$domains.Count - 1])"
        }
        else {
            $root = $domainHint
        }

        $domainZone = Join-Path $userRegRoot "Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Domains\$root" `
        | Get-ChildItem -ErrorAction SilentlyContinue `
        | Get-ItemProperty -ErrorAction SilentlyContinue `
        | & {
            process {
                $hostName = "$($_.PSChildName).$domainHint"
                $property = $_ | Split-ItemProperty | Select-Object -First 1
                $zoneId = $property.Value

                $zoneDisplayName = Join-Path $userRegRoot "Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones" `
                | Join-Path -ChildPath $zoneId `
                | Get-ItemProperty -Name 'DisplayName' -ErrorAction SilentlyContinue `
                | Select-Object -ExpandProperty 'DisplayName'

                [PSCustomObject]@{
                    HostName        = $hostName
                    Protocol        = $property.Name
                    ZoneId          = $property.Value
                    ZoneDisplayName = $zoneDisplayName
                }
            }
        }
    }

    [PSCustomObject]@{
        DomainHint                 = $domainHint
        EnableAlternateIdSupport   = $enableAlternateIdSupport
        InternetSettingsDomainZone = $domainZone
    }
}

function Get-UseOnlineContent {
    [CmdletBinding()]
    param (
        [string]$User
    )

    $userRegRoot = Get-UserRegistryRoot $User

    if (-not $userRegRoot) {
        return
    }

    $officeInfo = Get-OfficeInfo

    if (-not $officeInfo) {
        return
    }

    $major = $officeInfo.Version.Split('.')[0]

    & {
        "Software\Microsoft\Office\$major.0\Common\Internet"
        "Software\Policies\Microsoft\office\$major.0\common\Internet"
    }`
    | Join-Path -Path $userRegRoot -ChildPath { $_ } `
    | Get-ItemProperty -Name 'UseOnlineContent' -ErrorAction SilentlyContinue `
    | Split-ItemProperty
}

function Get-AutodiscoverConfig {
    [CmdletBinding()]
    param (
        [string]$User
    )

    $userRegRoot = Get-UserRegistryRoot $User

    if (-not $userRegRoot) {
        return
    }

    $officeInfo = Get-OfficeInfo

    if (-not $officeInfo) {
        return
    }

    $major = $officeInfo.Version.Split('.')[0]

    & {
        "Software\Microsoft\Office\$major.0\Outlook\AutoDiscover"
        "Software\Policies\Microsoft\Office\$major.0\Outlook\AutoDiscover"
        "Software\Microsoft\Exchange"
        "Software\Policies\Microsoft\Exchange"
    } `
    | Join-Path -Path $userRegRoot -ChildPath { $_ } `
    | Get-ItemProperty -Name 'Exclude*', 'Prefer*', '*Autodiscover*', 'ZeroConfigExchange*' -ErrorAction SilentlyContinue `
    | Split-ItemProperty
}

<#
.SYNOPSIS
Take an output of Get-ItemProperty and split its properties into objects with Name, Value, and Path.
#>
function Split-ItemProperty {
    [CmdletBinding()]
    param(
        [Parameter(ValueFromPipeline = $true)]
        [PSCustomObject]$Property,
        [string[]]$ExcludeProperty = @('PSPath', 'PSParentPath', 'PSChildName', 'PSProvider')
    )

    process {
        $Property | Get-Member -MemberType NoteProperty, Property | & {
            param([Parameter(ValueFromPipeline = $true)]$memberDefinition)
            process {
                if ($memberDefinition.Name -in $ExcludeProperty) {
                    return
                }

                [PSCustomObject]@{
                    Name  = $memberDefinition.Name
                    Value = $Property."$($memberDefinition.Name)"
                    Path  = $Property | Convert-Path -ErrorAction SilentlyContinue
                }
            }
        }
    }
}

function Get-SocialConnectorConfig {
    [CmdletBinding()]
    param (
        [string]$User
    )

    $userRegRoot = Get-UserRegistryRoot $User

    if (-not $userRegRoot) {
        return
    }

    & {
        'Software\Microsoft\Office\Outlook\SocialConnector'
        'Software\Policies\Microsoft\Office\Outlook\SocialConnector'
    } `
    | Join-Path -Path $userRegRoot -ChildPath { $_ } `
    | Get-ItemProperty -Name 'DownloadDetailsFromAd' -ErrorAction SilentlyContinue `
    | Split-ItemProperty
}

function Get-ImageFileExecutionOptions {
    [CmdletBinding()]
    param()

    Get-ItemProperty 'Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options\*' `
    | Select-Object -Property @{N = 'ImageName'; E = { $_.PSChildName } }, '*', @{N = 'Path'; E = { $_.PSPath.SubString(36) } } -ExcludeProperty 'PSParentPath', 'PSProvider', 'PSPath', 'PSChildName'
}

function Get-SessionManager {
    [CmdletBinding()]
    param()

    Get-ItemProperty 'Registry::HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Session Manager' `
    | Select-Object -Property '*' -ExcludeProperty 'PSParentPath', 'PSProvider'
}

function Set-ThreadCulture {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Culture
    )

    try {
        $newCulture = [System.Globalization.CultureInfo]::CreateSpecificCulture($Culture)

        # If CurrentUICulture is already the target culture, no need to change.
        if ($newCulture -eq [System.Threading.Thread]::CurrentThread.CurrentUICulture) {
            Write-Log "CurrentUICulture is already $Culture"
            return
        }

        # Save the current culture, but do not overwrite the saved value so that it can be reset to the original value later.
        if (-not $Script:SavedCulture) {
            $Script:SavedCulture = [System.Threading.Thread]::CurrentThread.CurrentCulture
        }

        if (-not $Script:SavedUICulture) {
            $Script:SavedUICulture = [System.Threading.Thread]::CurrentThread.CurrentUICulture
        }

        [System.Threading.Thread]::CurrentThread.CurrentCulture = $newCulture
        [System.Threading.Thread]::CurrentThread.CurrentUICulture = $newCulture

        # Changing CurrentThread.CurrentCulture & CurrentUICulture is not enough. NativeCultureResolver.m_Culture & m_uiCulture must be also changed.
        [System.Reflection.Assembly]::Load('System.Management.Automation').GetType('Microsoft.PowerShell.NativeCultureResolver').GetField('m_Culture', 'NonPublic, Static').SetValue($null, $newCulture)
        [System.Reflection.Assembly]::Load('System.Management.Automation').GetType('Microsoft.PowerShell.NativeCultureResolver').GetField('m_uiCulture', 'NonPublic, Static').SetValue($null, $newCulture)
    }
    catch {
        Write-Error -Message "Set-ThreadCulture failed" -Exception $_.Exception
    }
}

function Reset-ThreadCulture {
    [CmdletBinding()]
    param()

    try {
        if ($Script:SavedCulture) {
            [System.Threading.Thread]::CurrentThread.CurrentCulture = $Script:SavedCulture
            [System.Reflection.Assembly]::Load('System.Management.Automation').GetType('Microsoft.PowerShell.NativeCultureResolver').GetField('m_Culture', 'NonPublic, Static').SetValue($null, $Script:SavedCulture)
            $Script:SavedCulture = $null
        }

        if ($Script:SavedUICulture) {
            [System.Threading.Thread]::CurrentThread.CurrentUICulture = $Script:SavedUICulture
            [System.Reflection.Assembly]::Load('System.Management.Automation').GetType('Microsoft.PowerShell.NativeCultureResolver').GetField('m_uiCulture', 'NonPublic, Static').SetValue($null, $Script:SavedUICulture)
            $Script:SavedUICulture = $null
        }
    }
    catch {
        Write-Error -Message "Reset-ThreadCulture failed" -Exception $_.Exception
    }
}

<#
.SYNOPSIS
    Download ZoomIt
#>
function Download-ZoomIt {
    [CmdletBinding()]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseApprovedVerbs', '')]
    param(
        [Parameter(Mandatory = $true)]
        # Path to save zoomIt
        [string]$Path
    )

    $url = 'https://download.sysinternals.com/files/ZoomIt.zip'
    $zoomItZip = Join-Path $Path 'ZoomIt.zip'

    $err = $($null = Download-File -Uri $url -OutFile $zoomItZip) 2>&1 | Select-Object -First 1

    if (-not (Test-Path $zoomItZip)) {
        Write-Error -ErrorRecord $err
        return
    }

    # Expand ZIP file and remove it
    $err = $($null = Expand-Archive $zoomItZip -DestinationPath $Path) 2>&1

    if (-not $err) {
        Remove-Item $zoomItZip
    }
}

<#
.SYNOPSIS
    Start recording by using ZoomIt
#>
function Start-Recording {
    [CmdletBinding()]
    param(
        # Path to download ZoomIt if necessary
        [string]$ZoomItDownloadPath,
        # Path to look for zoomit executable, including subfolders
        [string]$ZoomItSearchPath
    )

    # OS version must be higher than Win10 1903 (Build 18362)
    $os = Get-CimInstance Win32_OperatingSystem

    try {
        if ($os.Version -match '(?<Major>\d+)\.\d+\.(?<Build>\d+)') {
            $major = $Matches.Major -as [int]
            $build = $Matches.Build -as [int]

            if (-not ($major -ge 10 -and $build -ge 18362)) {
                Write-Error "Windows version $($os.Version) is not supported"
                return
            }
        }
    }
    finally {
        $os.Dispose()
    }

    $zoomItExe = $null

    # Find the running instance of zoomit if exists and check if its version is >= 6 (Recording feature is available since version 6)
    $running = Get-Process 'ZoomIt*' | Select-Object -First 1

    if ($running) {
        $prop = Get-ItemProperty $running.Path -ErrorAction SilentlyContinue

        if ($prop.VersionInfo.FileVersion.ToString() -match '(?<Major>\d)\.') {
            $major = $Matches.Major -as [int]

            if ($major -ge 6) {
                $zoomItExe = $running.Path
                Write-Log "Found a running instance of $($running.Path) ($($prop.VersionInfo.FileVersion))"
            }
        }

        if (-not $zoomItExe) {
            Write-Error "$($running.Name) is running, but it's version is older than 6"
            return
        }
    }

    # Next, look for zoomIt executable under the given path (including subfolders)
    if (-not $zoomItExe -and $ZoomItSearchPath) {
        Get-ChildItem -Path (Join-Path $ZoomItSearchPath 'ZoomIt*.exe') -Recurse -ErrorAction SilentlyContinue | . {
            process {
                # Ignore one for ARM.
                if ($_.Name -eq 'ZoomIt64a.exe') {
                    return
                }

                if (-not $zoomItExe) {
                    $zoomItExe = $_.FullName
                }

                # For x64, prefer ZoomIt64.exe
                if ($env:PROCESSOR_ARCHITECTURE -eq 'AMD64' -and $_.Name -eq 'ZoomIt64.exe') {
                    $zoomItExe = $_.FullName
                }
            }
        }
    }

    # Still not found. Try to download.
    $downloaded = $false

    if (-not $zoomItExe -and $ZoomItDownloadPath) {
        Write-Log "Downloading ZoomIt"
        Download-ZoomIt -Path $ZoomItDownloadPath
        $downloaded = $true

        Get-ChildItem -Path (Join-Path $ZoomItDownloadPath 'ZoomIt*.exe') -ErrorAction SilentlyContinue | . {
            process {
                # Ignore one for ARM.
                if ($_.Name -eq 'ZoomIt64a.exe') {
                    return
                }

                if (-not $zoomItExe) {
                    $zoomItExe = $_.FullName
                }

                # For x64, prefer ZoomIt64.exe
                if ($env:PROCESSOR_ARCHITECTURE -eq 'AMD64' -and $_.Name -eq 'ZoomIt64.exe') {
                    $zoomItExe = $_.FullName
                }
            }
        }
    }

    if (-not $zoomItExe) {
        Write-Error "Cannot find ZoomIt executable"
        return
    }

    Unblock-File $zoomItExe
    $started = $false

    # If there was no running instance, start it.
    if (-not $running) {
        # Configure "OptionsShown" registry value so that zoomit's option won't be displayed.
        $zoomItPath = 'Registry::HKEY_CURRENT_USER\Software\Sysinternals\ZoomIt'

        if (-not (Test-Path $zoomItPath)) {
            $null = New-Item -Path $zoomItPath -Force
        }

        # Create or set "OptionsShown" to 1
        $null = Set-ItemProperty -Path $zoomItPath -Name 'OptionsShown' -Value 1

        # Start ZoomIt
        Write-Log "Starting $zoomItExe"
        $zoomIt = Start-Process $zoomItExe -ArgumentList '/AcceptEula' -PassThru
        $started = $true

        # Wait a little;otherwise zoomIt does not handle Ctrl+5
        # Start-Sleep -Seconds 1
    }

    # Send Ctrl+5 keybord input to start recording
    $success = $false
    $maxRetry = 5
    $interval = [TimeSpan]::FromMilliseconds(200)

    for ($i = 0; $i -lt $maxRetry; ++$i) {
        try {
            Write-Log "Sending Ctrl+5"
            [Win32.User32]::SendCtrl5()
        }
        catch {
            Write-Error -Message "Win32.User32.SendCtrl5 failed" -Exception $_.Exception
            break
        }

        Start-Sleep -Milliseconds $interval.TotalMilliseconds

        # Check if ZoomIt has started writing to %TEMP%\ZoomIt\zoomit.mp4. This file will be removed by ZoomIt when recording is stopped.
        if (Test-Path (Join-Path $env:TEMP 'ZoomIt\zoomit.mp4')) {
            $success = $true

            [PSCustomObject]@{
                Downloaded = $downloaded
                Started    = $started
            }

            break
        }
    }

    if ($zoomIt) {
        if (-not $success) {
            Write-Error "Failed to start recording"

            if ($started) {
                $zoomIt.Kill()
            }
        }

        $zoomIt.Dispose()
    }
}

<#
.SYNOPSIS
    Stop recording.

.NOTES
    This does not kill ZoomIt.
#>
function Stop-Recording {
    [CmdletBinding()]
    param()

    # Make sure zoomit is running.
    $zoomIt = Get-Process -Name 'ZoomIt*' | Select-Object -First 1

    # Send Ctrl+5 to stop recording.
    try {
        [Win32.User32]::SendCtrl5()
    }
    catch {
        Write-Error -Message "Win32.User32.SendCtrl5 failed" -Exception $_.Exception
    }

    if ($zoomIt) {
        $zoomIt.Dispose()
    }
}

function Get-ImageInfo {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        $Path
    )

    try {
        $stream = [System.IO.File]::OpenRead($Path)

        #  Bail if size is 0
        if ($stream.Length -eq 0) {
            Write-Error "The file size is 0"
            return
        }

        $reader = New-Object System.IO.BinaryReader $stream

        # Read IMAGE_DOS_HEADER. The first 2 bytes is "MZ" (0x5a4d)
        $magic = $reader.ReadUInt16()

        if ($magic -ne 0x5a4d) {
            Write-Error "This is not an executable image"
            return
        }

        # Get NT header offset at 0x3c from the beginning.
        $null = $reader.ReadBytes(0x3c - 2)
        $offsetToNTHeader = $reader.ReadUInt32()

        $reader.BaseStream.Position = $offsetToNTHeader

        # Make sure the signature is 0x00004550
        $signature = $reader.ReadUInt32()

        if ($signature -ne 0x4550) {
            Write-Error "Wrong signature for IMAGE_NT_HEADERS32"
            return
        }

        # The first 2 bytes of IMAGE_FILE_HEADER is the machine architecture
        # https://learn.microsoft.com/en-us/windows/win32/api/winnt/ns-winnt-image_file_header
        $machine = $reader.ReadUInt16()

        $arch = switch ($machine) {
            0x014c { 'x86'; break }
            0x0200 { 'IA64'; break }
            0x8664 { 'x64'; break }
            0xaa64 { 'ARM64'; break }
            default { 'Unknown'; break }
        }

        [PSCustomObject]@{
            Architecture = $arch
        }
    }
    catch {
        Write-Error -ErrorRecord $_
    }
    finally {
        if ($reader) {
            $reader.Close()
        }

        if ($stream) {
            $stream.Close()
        }
    }
}

function Get-PresentationMode {
    [CmdletBinding()]
    param()

    $state = [Win32.Shell32+QUERY_USER_NOTIFICATION_STATE]::QUNS_NOT_PRESENT
    $hr = [Win32.Shell32]::SHQueryUserNotificationState([ref]$state)

    if ($hr -gt 0) {
        Write-Error "SHQueryUserNotificationState failed with $hr"
        return
    }

    $PAGE_READONLY = 2;
    $fileMap = 'Local\FullScreenPresentationModeInfo'

    [Win32.SafeFileMappingHandle]$handle = [Win32.Kernel32]::OpenFileMappingW($PAGE_READONLY, $false, $fileMap)

    $fileMapExist = -not $handle.IsInvalid
    $handle.Dispose()

    if ($state -ne [Win32.Shell32+QUERY_USER_NOTIFICATION_STATE]::QUNS_ACCEPTS_NOTIFICATIONS -and $state -ne [Win32.Shell32+QUERY_USER_NOTIFICATION_STATE]::QUNS_QUIET_TIME) {
        $isPresentationMode = $true
    }
    else {
        $isPresentationMode = $fileMapExist
    }

    [PSCustomObject]@{
        User                     = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
        IsPresentationMode       = $isPresentationMode
        NotificationState        = $state
        PresentationFileMapExist = $fileMapExist
    }
}

function Get-AADBrokerPlugin {
    [CmdletBinding()]
    param()

    # Get-AppxPackage could throw a terminating error
    try {
        Get-AppxPackage -Name 'Microsoft.AAD.BrokerPlugin'
    }
    catch {
        Write-Error -Message "Get-AppxPackage Microsoft.AAD.BrokerPlugin threw a terminating error" -Exception $_.Exception
    }
}
function Add-LoopbackExempt {
    [CmdletBinding()]
    [OutputType([bool])]
    param(
        [Parameter(Mandatory = $true)]
        [string]$PackageFamiliyName
    )

    if (-not (Get-Command 'CheckNetIsolation.exe')) {
        Write-Error "CheckNetIsolation.exe is not available"
        return
    }

    # Check if it's already added.
    $found = CheckNetIsolation.exe LoopbackExempt -s | & {
        process {
            if (Select-String -InputObject $_ -Pattern $PackageFamiliyName -SimpleMatch -Quiet) {
                $true
            }
        }
    } | Select-Object -First 1

    if ($found) {
        Write-Error "$PackageFamiliyName is already in LoopbackExempt"
        return
    }

    # Add it (Note:package name MUST be double-quoted)
    $null = CheckNetIsolation.exe LoopbackExempt -a -n="$PackageFamiliyName"

    if ($LASTEXITCODE -eq 0) {
        Write-Log "$PackageFamiliyName is added to LoopbackExempt"
        $true
    }
    else {
        Write-Error "CheckNetIsolation failed with $LASTEXITCODE"
    }
}

function Remove-LoopbackExempt {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$PackageFamiliyName
    )

    $null = CheckNetIsolation.exe LoopbackExempt -d -n="$PackageFamiliyName"

    if ($LASTEXITCODE -eq 0) {
        Write-Log "$PackageFamiliyName is removed from LoopbackExempt"
    }
    else {
        Write-Error "CheckNetIsolation.exe failed with $LASTEXITCODE"
    }
}

function Get-AnsiCodePage {
    [CmdletBinding()]
    param()

    $acp = [Win32.Kernel32]::GetACP()
    [UInt32]$systemAcp = Get-ItemProperty 'Registry::HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Nls\CodePage' -Name 'ACP' | Select-Object -ExpandProperty 'ACP'

    [PSCustomObject]@{
        CurrentAnsiCodePage = $acp
        SystemAnsiCodePage  = $systemAcp
    }
}

function Save-USOSharedLog {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        $Path
    )

    $src = Join-Path $env:ProgramData 'USOShared\Logs\System'
    Save-Item -Path $src -Destination $Path -Filter 'MsoUsoCoreWorker.*.etl'
}

function Save-GPResult {
    [CmdletBinding(PositionalBinding = $false)]
    param(
        [Parameter(Mandatory)]
        # Destination folder path
        [string]$Path,
        [string]$User,
        [string]$FileName = 'GPResult',
        [ValidateSet('TEXT', 'HTML', 'XML')]
        [Parameter(Mandatory)]
        [string]$Format = 'TEXT',
        [Threading.CancellationToken]$CancellationToken
    )

    if (-not (Get-Command 'gpresult.exe' -ErrorAction SilentlyContinue)) {
        Write-Log "gpresult.exe is not available"
        return
    }

    $argList = New-Object System.Collections.Generic.List[string]

    if ($User) {
        $argList.Add('/USER')
        $argList.Add($User)
    }

    $fileName = [System.IO.Path]::GetFileNameWithoutExtension($FileName)

    switch ($Format) {
        'TEXT' {
            $filePath = Join-Path $Path "$fileName.txt"
            $argList.Add('/V')
            break
        }
        'HTML' {
            $filePath = Join-Path $Path "$fileName.htm"
            $argList.Add('/H')
            break
        }
        'XML' {
            $filePath = Join-Path $Path "$fileName.xml"
            $argList.Add('/X')
            break
        }
    }

    # Add file path argument for HTML & XML
    if ($Format -eq 'HTML' -or $Format -eq 'XML') {
        # If file path contains spaces, it must be double-quoted.
        if ($filePath.IndexOf(' ') -ge 0) {
            $filePath = "`"$filePath`""
        }

        $argList.Add($filePath)
    }

    $startProcArgs = @{
        ArgumentList = $argList
        WindowStyle  = 'Hidden'
        PassThru     = $true
    }

    if ($Format -eq 'TEXT') {
        # TODO:Using -RedirectStandardOutput is very slow. Refactor later by configuring System.Diagnostics.Process with StartInfo
        $startProcArgs.RedirectStandardOutput = $filePath
    }

    $start = Get-Timestamp
    Write-Log "Invoking 'gpresult.exe $argList'"

    $process = Start-Process 'gpresult.exe' @startProcArgs

    while ($true) {
        if ($process.WaitForExit(1000)) {
            # gpresult has finished
            break
        }

        # gpresult is still running. If cancellation is requested, bail.
        if ($CancellationToken.IsCancellationRequested) {
            Write-Log "Cancel request acknowledged"
            $process.Kill()
            break
        }
    }

    if ($process) {
        $process.Dispose()
    }

    $elapsed = Get-Elapsed $start
    Write-Log "gpresult.exe took $elapsed"
}

function Get-AppContainerRegistryAcl {
    [CmdletBinding(PositionalBinding = $false)]
    param(
        [Parameter(Position = 0)]
        [string]$User = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
    )

    $userRegRoot = Get-UserRegistryRoot -User $User

    if (-not $userRegRoot) {
        return
    }

    $appContainerPath = Join-Path $userRegRoot 'Software\Classes\Local Settings\Software\Microsoft\Windows\CurrentVersion\AppContainer'

    if (-not (Test-Path $appContainerPath)) {
        Write-Error "Cannot find $appContainerPath"
        return
    }

    $appContainerAcl = Get-Acl $appContainerPath

    # Includde "Mappings" key's ACL if avaialble.
    $mappingsPath = Join-Path $appContainerPath 'Mappings'

    if (Test-Path $mappingsPath) {
        $mappingsAcl = Get-Acl $mappingsPath
    }

    [PSCustomObject]@{
        User            = $User
        Path            = $appContainerPath
        AppContainerAcl = $appContainerAcl
        MappingsAcl     = $mappingsAcl
    }
}

<#
.SYNOPSIS
    Get StructuredQuery schema related information
#>
function Get-StructuredQuerySchema {
    [CmdletBinding(PositionalBinding = $false)]
    param(
        $User = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
    )

    $resolvedUser = Resolve-User $User
    $userRegRoot = Get-UserRegistryRoot -User $User

    if (-not $userRegRoot) {
        return
    }

    # Get user's Default UI Language (I'm not using Win32 GetUserDefaultUILanguage() or PowerShell's Get-WinUserLanguageList because the user running the command may be different from the target user
    # First, look for HKEY_CURRENT_USER\Control Panel\Desktop\PreferredUILanguages.
    # Note: To be more complete, the policy key (HKCU\Software\Policies\Microsoft\Control Panel\Desktop) should also be inspected, but skipping for now.
    $WinUILanguage = $userRegRoot | Join-Path -ChildPath 'Control Panel\Desktop' `
    | Get-ItemProperty -Name 'PreferredUILanguages' -ErrorAction SilentlyContinue | Select-Object -ExpandProperty 'PreferredUILanguages' | Select-Object -First 1

    if (-not $WinUILanguage) {
        Write-Log "PreferredUILanguages is missing. Looking for MachinePreferredUILanguages"

        # Look for HKEY_CURRENT_USER\Control Panel\Desktop\MuiCached\MachinePreferredUILanguages
        $WinUILanguage = $userRegRoot | Join-Path -ChildPath 'Control Panel\Desktop\MuiCached' `
        | Get-ItemProperty -Name 'MachinePreferredUILanguages' -ErrorAction SilentlyContinue | Select-Object -ExpandProperty 'MachinePreferredUILanguages' | Select-Object -First 1

        if (-not $WinUILanguage) {
            Write-Log "Cannot find Windows User PreferredUILanguage or MachinePreferredUILanguages"
        }
    }

    # See if QWORD registry value "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\StructuredQuer\SchemaChangedLast" exists.
    $schemaChangedLast = Get-ItemProperty 'Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\StructuredQuery' -Name 'SchemaChangedLast' -ErrorAction SilentlyContinue `
    | & {
        process {
            [DateTimeOffset]::FromFileTime($_.SchemaChangedLast)
        }
    }

    # Look for StructuredQuerySchema.bin under %LOCALAPPDATA%
    # e.g. "C:\Users\<user>\AppData\Local\Microsoft\Windows\1041\StructuredQuerySchema.bin"
    $localAppdata = Get-UserShellFolder -User $User -ShellFolderName 'Local AppData'

    if (-not $localAppdata) {
        return
    }

    $schemas = Join-Path $localAppdata 'Microsoft\Windows' `
    | Get-ChildItem -Filter 'StructuredQuerySchema.bin' -Recurse -ErrorAction SilentlyContinue `
    | & {
        param(
            [Parameter(ValueFromPipeline)]
            [System.IO.FileInfo]$fileInfo
        )

        process {
            $lcid = $null

            if ($fileInfo.FullName -match 'Windows\\(?<LCID>\d{4,5})') {
                $lcid = [int]::Parse($Matches['LCID'])
                $culture = New-Object System.Globalization.CultureInfo -ArgumentList $lcid
                $lastWriteTime = New-Object DateTimeOffset -ArgumentList $fileInfo.LastWriteTime

                [PSCustomObject]@{
                    LocaleName                   = $culture.Name
                    LocaleId                     = $lcid
                    Path                         = $fileInfo.FullName
                    LastWriteTime                = $lastWriteTime
                    IsNewerThanSchemaChangedLast = $schemaChangedLast -lt $lastWriteTime
                }
            }
            else {
                Write-Error "Failed to extract LCID from '$($fileInfo.FullName)'"
            }
        }
    }

    [PSCustomObject]@{
        User                  = $resolvedUser
        WindowsUserUILanguage = $WinUILanguage
        SchemaChangedLast     = $schemaChangedLast
        Schemas               = $schemas
    }
}

function Get-NetFrameworkVersion {
    [CmdletBinding()]
    param (
        [Parameter(ValueFromPipeline)]
        [string]$ComputerName = $env:COMPUTERNAME
    )

    process {
        # Read NDP registry
        $reg = $ndpKey = $null

        try {
            $reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey([Microsoft.Win32.RegistryHive]::LocalMachine, $ComputerName)
            $ndpKey = $reg.OpenSubKey("SOFTWARE\Microsoft\NET Framework Setup\NDP")

            if (-not $ndpKey) {
                Write-Error "OpenSubKey failed on 'SOFTWARE\Microsoft\NET Framework Setup\NDP'"
                return
            }

            @(
                foreach ($versionKeyName in $ndpKey.GetSubKeyNames()) {
                    # ignore "CDF" etc
                    if ($versionKeyName -notlike "v*") {
                        continue
                    }

                    $versionKey = $null

                    try {
                        $versionKey = $ndpKey.OpenSubKey($versionKeyName)

                        if (-not $versionKey) {
                            Write-Error "OpenSubKey failed on $versionKeyName. Skipping."
                            continue
                        }

                        $version = $versionKey.GetValue('Version', "")
                        $sp = $versionKey.GetValue('SP', "")
                        $install = $versionKey.GetValue('Install', "")

                        if ($version) {
                            [PSCustomObject]@{
                                Version      = $version
                                SP           = $sp
                                Install      = $install
                                SubKey       = $null
                                Release      = $release
                                NET45Version = $null
                                ComputerName = $ComputerName
                            }

                            continue
                        }

                        # for v4 and V4.0, check sub keys
                        foreach ($subKeyName in $versionKey.GetSubKeyNames()) {
                            if (-not $subKeyName) {
                                continue
                            }

                            $subKey = $null

                            try {
                                $subKey = $versionKey.OpenSubKey($subKeyName)

                                if (-not $subKey) {
                                    Write-Error "OpenSubKey failed on $subKeyName. Skipping."
                                    continue
                                }

                                $version = $subKey.GetValue("Version", "")
                                $install = $subKey.GetValue("Install", "")
                                $release = $subKey.GetValue("Release", "")

                                $NET45Version = $null

                                if ($release) {
                                    $NET45Version = Get-Net45Version $release
                                }

                                [PSCustomObject]@{
                                    Version      = $version
                                    SP           = $sp
                                    Install      = $install
                                    SubKey       = $subKeyName
                                    Release      = $release
                                    NET45Version = $NET45Version
                                    ComputerName = $ComputerName
                                }
                            }
                            finally {
                                if ($subKey) {
                                    $subKey.Close()
                                }
                            }
                        }
                    }
                    finally {
                        if ($versionKey) {
                            $versionKey.Close()
                        }
                    }
                }
            ) | Sort-Object -Property Version
        }
        catch {
            Write-Error -ErrorRecord $_
        }
        finally {
            if ($ndpKey) {
                $ndpKey.Close()
            }

            if ($reg) {
                $reg.Close()
            }
        }

    } # end of process block
}

function Get-Net45Version {
    [CmdletBinding()]
    [OutputType([string])]
    param (
        [Parameter(Mandatory)]
        $Release
    )

    switch ($Release) {
        { $_ -ge 533320 } { '4.8.1 or later'; break }
        { $_ -ge 528040 } { '4.8'; break }
        { $_ -ge 461808 } { '4.7.2'; break }
        { $_ -ge 460798 } { '4.7'; break }
        { $_ -ge 394802 } { "4.6.2"; break }
        { $_ -ge 394254 } { "4.6.1"; break }
        { $_ -ge 393295 } { "4.6"; break }
        { $_ -ge 379893 } { "4.5.2"; break }
        { $_ -ge 378675 } { '4.5.1'; break }
        { $_ -ge 378389 } { '4.5'; break }
        default { $null }
    }
}

<#
.SYNOPSIS
    Save Monarch related logs
#>
function Save-MonarchLog {
    [CmdletBinding()]
    param(
        # Destination folder path
        [Parameter(Mandatory)]
        [string]$Path,
        $User
    )

    # Collect data in %LOCALAPPDAT%\Microsoft\Olk
    $localAppdata = Get-UserShellFolder -User $User -ShellFolderName 'Local AppData'
    $olk = Join-Path $localAppdata -ChildPath 'Microsoft\Olk'

    if (-not (Test-Path $olk)) {
        Write-Log "Cannot find '$olk'"
        return
    }

    # For now, explicitly select the items to be copied.
    & {
        @{ Path = $olk; Destination = $Path; Filter = '*.log' }
        @{ Path = $olk; Destination = $Path; Filter = '*.json' }
        @{ Path = $olk; Destination = $Path; Filter = '*.txt' }
        @{ Path = "$olk\logs"; Destination = "$Path\logs"; Recurse = $true }
        @{ Path = "$olk\EBWebView\Crashpad"; Destination = "$Path\EBWebView\Crashpad"; Recurse = $true }
    } | & {
        process {
            Save-Item @_
        }
    }
}

function Save-MonarchSetupLog {
    [CmdletBinding()]
    param(
        # Destination folder path
        [Parameter(Mandatory)]
        [string]$Path,
        $User
    )

    # NewOutlookInstaller.exe writes "NewOutlookInstaller_***.log" in %TEMP%
    # setup.exe writes to Setup_***.log in %TEMP%
    # This may change in future, but for now we need to collect both.
    $temp = Get-UserTempFolder -User $User

    if (-not $temp) {
        Write-Error "Failed to find TEMP folder for $User"
        return
    }

    & {
        @{ Path = $temp; Filter = 'setup_*.log' }
        @{ Path = $temp; Filter = 'NewOutlookInstaller_*.log' }
    } | & {
        process {
            Save-Item @_ -Destination $Path
        }
    }
}

<#
.SYNOPSIS
    Enable DevTools for Monarch. This takes effect for subsequent launches of Monarch.

.NOTES
    There are several ways to enable DevTools for Monarch:

    1. Environment Variable, WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS with "--auto-open-devtools-for-tabs"

    2. Registry Key, HKLM|HKCU\SOFTWARE\Policies\Microsoft\Edge\WebView2\AdditionalBrowserArguments with name "olk.exe" with "--auto-open-devtools-for-tabs"

    Both options apply to any WebView2 apps as explained below:

    Globals
    https://learn.microsoft.com/en-us/microsoft-edge/webview2/reference/win32/webview2-idl?view=webview2-1.0.2535.41#createcorewebview2environmentwithoptions

    WebView2 browser flags
    https://learn.microsoft.com/en-us/microsoft-edge/webview2/concepts/webview-features-flags?tabs=dotnetcsharp

    3. config.json in %LOCALAPPDATA%\Microsoft\Olk
       This is Monarch only.

    For now, this command uses the 3rd method, but it might change in future.
#>
function Enable-MonarchDevTools {
    [CmdletBinding()]
    param(
        $User
    )

    $localAppdata = Get-UserShellFolder -User $User -ShellFolderName 'Local AppData'
    $olk = Join-Path $localAppdata -ChildPath 'Microsoft\Olk'
    $configJson = 'config.json'
    $config = Join-Path $olk -ChildPath $configJson

    $tempName = '39f3719a-b064-465c-87c7-ccd09ba007df'
    $tempFile = Join-Path $olk -ChildPath $tempName

    $content = '{ "edgeDevTools" : "autoopen" }'

    # SHA-1 hash for the config.json that we create (Make sure to update this value when the content is changed)
    $hash = 'D80733D2CAAE95D4748799192F0C10C92B123513'

    # If config.json exists, temporarily rename it and create a new one.
    if (Test-Path $config) {
        # If the content is same as the one we are going to write, skip.
        $fileHash = Get-FileHash -Path $config -Algorithm SHA1

        if ($fileHash.Hash -eq $hash) {
            Write-Log "config.json already exists with the expected content"
            return
        }

        # If the temp file already exists for some reason, remove it.
        if (Test-Path $tempFile) {
            Remove-Item $tempFile -Force
        }

        $err = $(Rename-Item -Path $config -NewName $tempName) 2>&1

        if ($err) {
            Write-Error "Failed to rename $config. $err"
            return
        }
    }

    $err = $($null = New-Item -Path $olk -Name $configJson -ItemType File) 2>&1

    if ($err) {
        Write-Error "Failed to create $config. $err"
        return
    }

    # Note: This file content must be written without BOM (thus "-Encoding Ascii")
    $err = $(Set-Content -Path $config -Value $content -Encoding Ascii) 2>&1

    if ($err) {
        Write-Error "Failed to write to $config. $err"
    }
}

<#
.SYNOPSIS
    Disable DevTools for Monarch
#>
function Disable-MonarchDevTools {
    [CmdletBinding()]
    param(
        $User
    )

    $localAppdata = Get-UserShellFolder -User $User -ShellFolderName 'Local AppData'
    $olk = Join-Path $localAppdata -ChildPath 'Microsoft\Olk'
    $configJson = 'config.json'
    $config = Join-Path $olk -ChildPath $configJson

    $tempName = '39f3719a-b064-465c-87c7-ccd09ba007df'

    # Remove the current config.json and restore temp file if exists.
    if (-not (Test-Path $config)) {
        Write-Error "Cannot find $config"
        return
    }

    $err = $(Remove-Item -Path $config -Force) 2>&1

    if ($err) {
        Write-Error "Failed to remove $config. $err"
        return
    }

    $tempFile = Join-Path $olk -ChildPath $tempName

    if (Test-Path $tempFile) {
        $err = $(Rename-Item -Path $tempFile -NewName $configJson) 2>&1

        if ($err) {
            Write-Error "Failed to rename $tempFile to $configJson. $err"
        }
    }
}

<#
.SYNOPSIS
    Enable Edge DevTools. This takes effect for subsequent launches of Monarch.

.NOTES
    There are several ways to enable DevTools for Monarch:

    1. Environment Variable, WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS with "--auto-open-devtools-for-tabs"

    2. Registry Key, HKLM|HKCU\SOFTWARE\Policies\Microsoft\Edge\WebView2\AdditionalBrowserArguments with name "olk.exe" with "--auto-open-devtools-for-tabs"

    Both options apply to any WebView2 apps as explained below:

    Globals
    https://learn.microsoft.com/en-us/microsoft-edge/webview2/reference/win32/webview2-idl?view=webview2-1.0.2535.41#createcorewebview2environmentwithoptions

    WebView2 browser flags
    https://learn.microsoft.com/en-us/microsoft-edge/webview2/concepts/webview-features-flags?tabs=dotnetcsharp

    3. config.json in %LOCALAPPDATA%\Microsoft\Olk
       This is Monarch only.

    This command uses the 2nd method.
#>
function Enable-WebView2DevTools {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        # Target executable name, such as "olk.exe"
        [string]$ExecutableName,
        $User
    )

    Add-WebView2Flags @PSBoundParameters -FlagNameAndValues @{ 'auto-open-devtools-for-tabs' = $null }
}

function Disable-WebView2DevTools {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        # Target executable name, such as "olk.exe"
        [string]$ExecutableName,
        $User
    )

    Remove-WebView2Flags @PSBoundParameters -FlagNames 'auto-open-devtools-for-tabs'
}

function Enable-WebView2Netlog {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        # Target executable name, such as "olk.exe"
        [string]$ExecutableName,
        $User,
        [Parameter(Mandatory)]
        # Folder file path
        [string]$Path,
        [string]$FileName = "netlog_$(Get-DateTimeString).json",
        [ValidateSet('Default', 'IncludeSensitive', 'Everything')]
        [string]$CaptureMode = 'Everything',
        [UInt32]$MaxFileSizeMB
    )

    if (-not (Test-Path $Path)) {
        $null = New-Item -Path $Path -ItemType Directory -ErrorAction Stop
    }

    $Path = Convert-Path -LiteralPath $Path

    $flags = @{
        'log-net-log'          = Join-Path $Path $FileName
        'net-log-capture-mode' = $CaptureMode
    }

    if ($PSBoundParameters.ContainsKey('MaxFileSizeMB')) {
        $flags.Add('net-log-max-size-mb', $MaxFileSizeMB.ToString())
    }

    Add-WebView2Flags -ExecutableName $ExecutableName -User $User -FlagNameAndValues $flags
}

function Disable-WebView2NetLog {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        # Target executable name, such as "olk.exe"
        [string]$ExecutableName,
        $User
    )

    Remove-WebView2Flags -ExecutableName $ExecutableName -User $User -FlagNames 'log-net-log', 'net-log-capture-mode', 'net-log-max-size-mb'
}

function Get-WebView2Flags {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$ExecutableName,
        $User
    )

    $userRegRoot = Get-UserRegistryRoot -User $User

    if (-not $userRegRoot) {
        return
    }

    $ExecutableName = [IO.Path]::ChangeExtension($ExecutableName, 'exe')

    $keyPath = Join-Path $userRegRoot 'SOFTWARE\Policies\Microsoft\Edge\WebView2\AdditionalBrowserArguments'
    $flags = @{}

    if (Test-Path $keyPath) {
        $flagsString = Get-ItemProperty $keyPath -Name $ExecutableName -ErrorAction SilentlyContinue | Select-Object -ExpandProperty $ExecutableName

        if ($flagsString) {
            # sample: --auto-open-devtools-for-tabs --log-net-log="C:\users\admin\desktop\netlog\net log.json" --net-log-capture-mode=Everything

            $current = -1
            $entries = New-Object System.Collections.Generic.List[[string]]

            # First extract entries "--***"
            while ($true) {
                $next = $flagsString.IndexOf('--', <# startIndex #> $current + 1)

                if ($next -eq -1) {
                    if ($current -ne -1) {
                        $entries.Add($flagsString.Substring($current + 2).Trim())
                    }

                    break
                }

                if ($current -ne -1) {
                    $entries.Add($flagsString.Substring($current + 2, $next - $current - 2).Trim())
                }

                $current = $next
            }

            # Next, split each with "name=value" pair. note: "=value" may be missing
            foreach ($entry in $entries) {
                $nameAndValue = $entry.Split('=')
                $name = $nameAndValue[0]
                $value = $null

                if ($nameAndValue.Count -gt 1) {
                    $value = $nameAndValue[1]
                }

                $flags.Add($name, $value)
            }
        }
    }

    [PSCustomObject]@{
        Path  = $keyPath # Make sure to keep "Registry::" prefix here because Add-WebView2Flags uses it with Test-Path
        Name  = $ExecutableName
        Flags = $flags
    }
}

<#
.SYNOPSIS
    Helper function for Add/Remove-WebView2Flags
#>
function Format-WebView2Flags {
    param(
        [Hashtable]$Flags
    )

    $sb = New-Object System.Text.StringBuilder

    foreach ($entry in $Flags.GetEnumerator()) {
        $null = $sb.Append("--$($entry.Key)")

        if ($entry.Value) {
            $value = $entry.Value

            # If there is a space in value, surround by double-quotations
            if ($value.IndexOf(' ') -ge 0) {
                $value = "`"$value`""
            }

            $null = $sb.Append("=$value")
        }

        $null = $sb.Append(' ')
    }

    $sb.ToString()
}

function Add-WebView2Flags {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$ExecutableName,
        [Parameter(Mandatory)]
        [Hashtable]$FlagNameAndValues,
        $User
    )

    # For now, support only the following flags:
    $validFlags = @('auto-open-devtools-for-tabs', 'disable-background-timer-throttling', 'log-net-log', 'net-log-capture-mode', 'net-log-max-size-mb')

    foreach ($flagName in $FlagNameAndValues.Keys) {
        if ($flagName -notin $validFlags) {
            Write-Error "Flag '$flagName' is not supported. Supported Flags: $($validFlags -join ',')"
            return
        }
    }

    $wv2Flags = Get-WebView2Flags -ExecutableName $ExecutableName -User $User

    $regValueName = $wv2Flags.Name
    $keyPath = $wv2Flags.Path
    $flags = $wv2Flags.Flags

    if (-not (Test-Path $keyPath)) {
        $err = $($key = New-Item $keyPath -Force) 2>&1 | Select-Object -First 1

        if ($key) {
            $key.Dispose()
        }
        else {
            Write-Error -Message "Failed to create $keyPath. $err" -Exception $err.Exception
            return
        }
    }

    # Consolidate current & new flags
    $isAdded = $false

    foreach ($entry in $FlagNameAndValues.GetEnumerator()) {
        # Add it if the flag does not exist or its value is different
        if (-not $flags.ContainsKey($entry.Key) -or $flags[$entry.Key] -ne $entry.Value) {
            $flags.Remove($entry.Key)
            $flags.Add($entry.Key, $entry.Value)
            $isAdded = $true
        }
    }

    if (-not $isAdded) {
        return
    }

    # Format flags to a string
    # e.g. "--auto-open-devtools-for-tabs --net-log-capture-mode=Everything"
    $flagsString = Format-WebView2Flags $flags

    # Set-ItemProperty either creates a new value or overwrites the existing value.
    $err = Set-ItemProperty $keyPath -Name $regValueName -Value $flagsString 2>&1 | Select-Object -First 1

    if ($err) {
        Write-Error -Message "Failed to set '$regValueName' in $keyPath. $err" -Exception $err.Exception
    }
    else {
        Write-Log "Flags set for $($regValueName): $flagsString"
    }
}

function Remove-WebView2Flags {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$ExecutableName,
        [Parameter(Mandatory)]
        [string[]]$FlagNames,
        $User
    )

    $wv2Flags = Get-WebView2Flags -ExecutableName $ExecutableName -User $User

    $regValueName = $wv2Flags.Name
    $keyPath = $wv2Flags.Path
    $flags = $wv2Flags.Flags
    $originalFlagCount = $flags.Count

    foreach ($flagName in $FlagNames) {
        $flags.Remove("$flagName")
    }

    if ($flags.Count -eq $originalFlagCount) {
        Write-Log "No flags is removed"
        return
    }

    if ($flags.Count) {
        # Format flags to a string
        $flagsString = Format-WebView2Flags $flags

        $err = Set-ItemProperty $keyPath -Name $regValueName -Value $flagsString 2>&1 | Select-Object -First 1

        if ($err) {
            Write-Error -Message "Failed to set '$regValueName' in $keyPath. $err" -Exception $err.Exception
        }
        else {
            Write-Log "Flags set: $flagsString"
        }
    }
    else {
        # If the resulting flags is empty, remove the registry value
        Remove-ItemProperty $keyPath -Name $regValueName -ErrorAction SilentlyContinue
    }
}

function Get-FileExtEditFlags {
    param(
        [Parameter(Mandatory)]
        # File extension such as ".txt", ".jpg"
        $Extension,
        $User
    )

    if (-not $Extension.StartsWith('.')) {
        $Extension = ".$Extension"
    }

    $obj = @{
        Extension = $Extension
    }

    # HKEY_CLASSES_ROOT\<extesion>
    $ext = Get-ItemProperty "Registry::HKEY_CLASSES_ROOT\$Extension" -Name '(default)' -ErrorAction SilentlyContinue

    if ($ext -and $ext.'(default)') {
        $handlerName = $ext.'(default)'
        $handler = Get-ItemProperty "Registry::HKEY_CLASSES_ROOT\$handlerName" -Name 'EditFlags' -ErrorAction SilentlyContinue

        if ($handler) {
            $obj.Path = $handler | Convert-Path -ErrorAction SilentlyContinue

            # EditFlags can be DWORD or BINARY
            if ($handler.EditFlags -is [byte[]]) {
                $obj.EditFlags = [System.BitConverter]::ToUInt32($handler.EditFlags, 0)
            }
            else {
                $obj.EditFlags = $handler.EditFlags
            }

            [PSCustomObject]$obj
        }
    }

    # HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FileExts\<extesion>
    $userRegRoot = Get-UserRegistryRoot -User $User
    $fileExt = Join-Path $userRegRoot "SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FileExts\$Extension" `
    | Get-ItemProperty -Name 'EditFlags' -ErrorAction SilentlyContinue

    if ($fileExt) {
        $obj.Path = $fileExt | Convert-Path -ErrorAction SilentlyContinue

        if ($fileExt.EditFlags -is [byte[]]) {
            $obj.EditFlags = [System.BitConverter]::ToUInt32($fileExt.EditFlags, 0)
        }
        else {
            $obj.EditFlags = $fileExt.EditFlags
        }

        [PSCustomObject]$obj
    }
}

<#
.SYNOPSIS
Check if this sript is too old.
It returns $false if ValidTimeSpan has passed since ReleaseDate.
#>
function Test-ScriptExpiration {
    [CmdletBinding()]
    [OutputType([bool])]
    param(
        [DateTime]$ReleaseDate = [DateTime]::Parse($Version.Substring(1)),
        [TimeSpan]$ValidTimeSpan = $Script:ValidTimeSpan
    )

    [DateTime]::Now - $ReleaseDate -le $ValidTimeSpan
}

function Disable-CtrlC {
    [CmdletBinding()]
    [OutputType([bool])]
    param()

    $success = $false

    # This does not work in some environments such as PowerShell ISE.
    try {
        [Console]::TreatControlCAsInput = $true
        $success = $true
    }
    catch {
        # ignore
    }

    $success
}

function Enable-CtrlC {
    [CmdletBinding()]
    [OutputType([bool])]
    param()

    $success = $false

    # This does not work in some environments such as PowerShell ISE.
    try {
        [Console]::TreatControlCAsInput = $false
        $success = $true
    }
    catch {
        # ignore
    }

    $success
}

function Get-ExperimentConfigs {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        # Application name such as "Outlook" (See HKCU\Software\Microsoft\Office\16.0\Common\ExperimentConfigs\Ecs)
        $AppName,
        $User,
        [switch]$SkipParsing
    )

    function Get-Value {
        param(
            [Parameter(Mandatory, ValueFromPipeline)]
            [string]$data,
            [switch]$Skip
        )

        process {
            if ($Skip) {
                return $data
            }

            # Note: the value can be empty (e.g. "std::wstring|")
            if ($data -Match '(?<DataType>[\w:]+)\|(?<Value>.*)') {
                $typeName = $Matches['DataType']
                $value = $Matches['Value']

                switch ($typeName) {
                    'bool' { if ($value -eq 0) { $false } else { $true } }
                    'Mso::AnyType' { $data } # as-is
                    default { $value }
                }
            }
            else {
                Write-Verbose "Cannot parse $data"
                $data
            }
        }
    }

    $userRegRoot = Get-UserRegistryRoot -User $User
    $key = Join-Path $userRegRoot "Software\Microsoft\Office\16.0\Common\ExperimentConfigs\Ecs\$AppName\ConfigContextData"

    # Read ChunkCount and VersionId
    $configContextData = Get-ItemProperty $key -Name 'ChunkCount', 'VersionId' -ErrorAction SilentlyContinue

    if (-not $configContextData) {
        Write-Error "Cannot find ChunkCount & VersionId in $($key | ConvertFrom-PSPath)"
        return
    }

    # ChunkCount is a string like "uint64_t|40". Extract the value
    $chunkCount = $null
    $versionId = $null

    if ($configContextData.ChunkCount -Match '(?<DataType>\w+)\|(?<Value>\w+)') {
        $chunkCount = $Matches['Value'] -as [int]
    }

    if (-not $chunkCount) {
        Write-Error "ChunkCount registry value is missing in $key"
        return
    }

    # VersionId is a string like "uint16_t|1". Extract the value
    if ($configContextData.VersionId -Match '(?<DataType>\w+)\|(?<Value>\w+)') {
        $versionId = $Matches['Value']
    }

    if (-not $versionId) {
        Write-Error "VersionId registry value is missing in $key"
        return
    }

    $sb = New-Object System.Text.StringBuilder

    # Read "1", "1.1", "1.2", ... "1.40"
    for ($chunkIndex = 0; $chunkIndex -le $chunkCount; ++$chunkIndex) {
        $name = $versionId

        if ($chunkIndex -gt 0) {
            $name += ".$chunkIndex"
        }

        $bin = Get-ItemProperty $key -Name $name | Select-Object -ExpandProperty $name
        $null = $sb.Append([System.Text.Encoding]::ASCII.GetString($bin))
    }

    $dataString = $sb.ToString()
    Write-Verbose "Raw string: `n$dataString"

    # Extract "EcsConfigResponseData" json value
    $keyword = 'EcsConfigResponseData|'
    $jsonData = $dataString.Substring($dataString.IndexOf($keyword) + $keyword.Length)

    $config = ConvertFrom-Json $jsonData

    # If SkipParsing is specified, make Get-Value a no-op
    $skipParamKey = 'Get-Value:Skip'

    if ($SkipParsing) {
        $PSDefaultParameterValues[$skipParamKey] = $true
    }

    # Parse FCMap
    # It's an array of kvp, where Key is "F" and Value is "V".
    # Note:fcMap itself should not be an object because it's possible that there are multiple properties with the same name with different cases
    # e.g. "Microsoft.Office.FileIO.NoDocumentVersionsCommandsForHistoricVersions" & "Microsoft.Office.FileIO.nodocumentversionscommandsforhistoricversions"
    $fcMap = @(
        foreach ($item in $config.FCMap) {
            [PSCustomObject]@{
                Name  = $item.F
                Value = $item.V | Get-Value
            }
        }
    )

    # Parse FCGroupMap
    # It's an object with properties called "FCGroupMap_1", "FCGroupMap_2" etc. and each group is an array of kvp, where Key is "F" and Value is "V".
    $fcGroupMap = @{}

    $config.FCGroupMap | Get-Member -MemberType Properties | & {
        process {
            $values = @(
                foreach ($item in $config.FCGroupMap.($_.Name)) {
                    [PSCustomObject]@{
                        Name  = $item.F
                        Value = $item.V | Get-Value
                    }
                }
            )

            $fcGroupMap.($_.Name) = $values
        }
    }

    [PSCustomObject]@{
        AppName               = $AppName
        Version               = $config.Ver | Get-Value
        ConfigIds             = $config.ConfIds | Get-Value
        CountryCode           = $config.CC | Get-Value
        DeferredConfigs       = $config.DefConfs | Get-Value
        ExpiryTime            = if ($SkipParsing) { $config.ExpTime } else { [DateTimeOffset]::FromUnixTimeSeconds(($config.ExpTime | Get-Value)) }
        ETag                  = $config.ETag | Get-Value
        FeatureConfigMap      = $fcMap
        GroupFeatureConfigMap = [PSCustomObject]$fCGroupMap
    }

    $PSDefaultParameterValues.Remove($skipParamKey)
}

function Get-CloudSettings {
    [CmdletBinding()]
    param(
        $User
    )

    $userRegRoot = Get-UserRegistryRoot -User $User

    if (-not $userRegRoot) {
        return
    }

    Join-Path $userRegRoot 'Software\Microsoft\Office\Outlook\Settings\Data' `
    | Get-ItemProperty -ErrorAction SilentlyContinue `
    | Split-ItemProperty `
    | & {
        process {
            # Try to convert JSON value to a PSCustomObject. It it fails, keep the original.
            $err = $($value = $_.Value | ConvertFrom-Json) 2>&1

            if ($err) {
                $value = $_.Value
            }

            [PSCustomObject]@{
                Name  = $_.Name
                Value = $value
            }
        }
    }
}

<#
.SYNOPSIS
Wait until user enters Enter key or Ctrl+C.
This is only possible when Console is available.
Console is not available in PowerShell ISE and in this case Ctrl+C will interrupt.
#>
function Wait-EnterOrControlC {
    [CmdletBinding()]
    param()

    # Check if a console is available, and if so, manually detect Enter key and Ctrl+C.
    $consoleAvailable = $false

    try {
        $Host.UI.RawUI.FlushInputBuffer()
        [Console]::TreatControlCAsInput = $true
        $consoleAvailable = $true
    }
    catch {
        # Ignore
    }

    if ($consoleAvailable) {
        $detectedKey = $null

        while ($true) {
            [ConsoleKeyInfo]$keyInfo = [Console]::ReadKey(<# intercept #> $true)

            # Enter or Ctrl+C exits the wait loop
            if ($keyInfo.Key -eq [ConsoleKey]::Enter) {
                Write-Log "Enter key is detected"
                $detectedKey = 'Enter'
            }
            elseif (($keyInfo.Modifiers -band [ConsoleModifiers]'Control') -and ($keyInfo.Key -eq [ConsoleKey]::C)) {
                Write-Log "Ctrl+C is detected" -Category Error
                $detectedKey = 'Ctrl+C'
            }

            if ($detectedKey) {
                break
            }
        }

        [Console]::TreatControlCAsInput = $false
        Write-Host
    }
    else {
        # Read-Host is not used here because it'd block background tasks.
        # When using UI.ReadLine(), Ctrl+C cannot be detected.
        # Note: PowerShell ISE does not implement $host.UI.RawUI.ReadKey().
        $null = $host.UI.ReadLine()
        $detectedKey = 'Enter'
    }

    [PSCustomObject]@{
        Key                = $detectedKey
        IsConsoleAvailable = $consoleAvailable
    }
}

<#
.SYNOPSIS
    Helper function that returns a string of command expression with given parameters.

.EXAMPLE
    Get-CommandExpression -Command Get-Process -Parameters @{ Name = 'Outlook' }

.EXAMPLE
    Get-CommandExpression -Invocation $MyInvocation

.NOTES
    This function does not check if the given parameters belong to the same ParameterSet.
    So, there is no guarantee that the output expression runs successfully.

    For example, the following returns "Get-Process -Name Outlook -Id 123", but "Name" & "Id" parameters cannot be used simultaneously.

    Get-CommandExpression -Command Get-Process -Parameters @{ Name = 'Outlook'; Id = '123' }
#>
function Get-CommandExpression {
    [CmdletBinding(PositionalBinding = $false)]
    [OutputType([string])]
    param(
        [Parameter(ParameterSetName = 'Command', Mandatory)]
        $Command,
        [Parameter(ParameterSetName = 'Command', Mandatory)]
        [Hashtable]$Parameters,
        [Parameter(ParameterSetName = 'Invocation', Mandatory)]
        [System.Management.Automation.InvocationInfo]
        $Invocation
    )

    if ($PSCmdlet.ParameterSetName -eq 'Invocation') {
        $Command = $Invocation.MyCommand
        $Parameters = $Invocation.BoundParameters
    }

    if ($Command -is [string]) {
        $Command = Get-Command $Command -ErrorAction SilentlyContinue

        if (-not $Command) {
            Write-Error "Cannot find $Command"
            return
        }
    }

    if ($Command -isnot [System.Management.Automation.CommandInfo]) {
        Write-Error "Need a CommandInfo for Command paramter"
        return
    }

    # It is expected to be passed a FunctionInfo or CmdletInfo. Anything else (such as ScriptInfo) is not really expected while it just returns an empty string
    if ($Command -isnot [System.Management.Automation.FunctionInfo] -or $Command -isnot [System.Management.Automation.CmdletInfo]) {
        Write-Verbose "Passed Command is of type $($Command.GetType().FullName)"
    }

    $sb = New-Object System.Text.StringBuilder -ArgumentList $Command.Name

    foreach ($param in $Parameters.GetEnumerator()) {
        # Skip if the given parameter name is not available
        if (-not $Command.Parameters.ContainsKey($param.Key)) {
            continue
        }

        $null = $sb.Append(" -$($param.Key)")

        # If this is a Switch parameter, no need to add value
        if ($Command.Parameters[$param.Key].SwitchParameter) {
            continue
        }

        $value = $param.Value

        if ($value -is [string] -and $value.IndexOf(' ') -ge 0) {
            $value = "'$value'"
        }
        elseif ($param.Value -is [System.Collections.ICollection]) {
            $value = $param.Value -join ', '
        }

        $null = $sb.Append(" $value")
    }

    $sb.ToString()
}

<#
.SYNOPSIS
    Helper function to convert an argument Hashtable or Array to a string representation
#>
function ConvertFrom-ArgumentList {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        $ArgumentList
    )

    if ($ArgumentList -is [Hashtable]) {
        if ($ArgumentList.Count -eq 0) {
            '@{}'
            return
        }

        $sb = New-Object System.Text.StringBuilder '@{'

        foreach ($entry in $ArgumentList.GetEnumerator()) {
            $null = $sb.Append(' ').Append($entry.Key).Append(' = ')

            if ($entry.Value -is [string] -and $entry.Value.IndexOf(' ') -ge 0) {
                $null = $sb.Append("'$($entry.Value)'")
            }
            else {
                $null = $sb.Append($entry.Value)
            }

            $null = $sb.Append(';')
        }

        # Remove the last ';' & close with '}'
        $null = $sb.Remove($sb.Length - 1, 1).Append(' }')
    }
    else {
        if ($null -ne $ArgumentList -and $ArgumentList -isnot [Array]) {
            $ArgumentList = @($ArgumentList)
        }

        if ($ArgumentList.Count -eq 0) {
            '@()'
            return
        }

        $sb = New-Object System.Text.StringBuilder '@('

        foreach ($entry in $ArgumentList) {
            $null = $sb.Append(' ').Append($entry).Append(',')
        }

        $null = $sb.Remove($sb.Length - 1, 1).Append(' )')
    }

    $sb.ToString()
}

<#
.SYNOPSIS
    Helper function to select a single process instance (owner is also checked when User parameter is given)
    If there are multipile processes, ask the user.
#>
function Get-SingleProcess {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        # Process name (such as 'outlook')
        [string]$Name,
        $User,
        [string]$CommandLineFilter
    )

    $processName = [IO.Path]::ChangeExtension($Name, 'exe')

    $win32Processes = @(Get-CimInstance Win32_Process -Filter "Name = '$processName'" | & {
            param([Parameter(ValueFromPipeline)]$win32Proc)
            process {
                # Drop if its owner is not the target user
                if ($User) {
                    $owner = $win32Proc | Get-ProcessOwner

                    if ($owner -and $owner.Sid -ne $User.Sid) {
                        Write-Log -Message "Skipping $processName (PID:$($win32Proc.ProcessId)) because its owner is `"$owner`", not the target user `"$User`"" -Category Warning
                        $win32Proc.Dispose()
                        return
                    }
                }

                # Drop if its commandline does not contain the given filter
                if ($CommandLineFilter) {
                    if ($win32Proc.CommandLine.IndexOf($CommandLineFilter, [StringComparison]::OrdinalIgnoreCase) -lt 0) {
                        Write-Log -Message "Skipping $processName (PID:$($win32Proc.ProcessId)) because its command line does not contain the filter `"$CommandLineFilter`"" -Category Warning
                        $win32Proc.Dispose()
                        return
                    }
                }

                $_
            }
        })

    $selectedProcess = $null

    try {
        switch ($win32Processes.Count) {
            0 { <# nothing to return #> break }
            1 { $selectedProcess = $win32Processes[0]; break }

            default {
                # Multiple instances are found. Ask the user to select one.
                $msg = New-Object System.Text.StringBuilder -ArgumentList "Multiple instances of $processName are found:`n`n"

                foreach ($win32Proc in $win32Processes) {
                    $null = $msg.Append("- $($win32Proc.Name) (PID:$($win32Proc.ProcessId))`n")
                }

                Write-Host $msg.ToString() -ForegroundColor Yellow

                while (-not $selectedProcess) {
                    Write-Host "Please enter the PID of the target process: " -NoNewline
                    $userInput = $host.UI.ReadLine()
                    $id = 0

                    if ([int]::TryParse($userInput, [ref]$id)) {
                        $selectedProcess = $win32Processes | Where-Object { $_.ProcessId -eq $id }

                        if (-not $selectedProcess) {
                            Write-Host "Cannot find $processName with PID:$id" -ForegroundColor Yellow
                        }
                    }
                    elseif ($userInput -eq 'q') {
                        break
                    }
                    else {
                        Write-Host "Invalid input `"$userInput`". Please enter a valid PID (to quit, press q)" -ForegroundColor Yellow
                    }
                }
            }
        }

        if ($selectedProcess) {
            # Return an object with minimal info so that the caller does not need to dispose the object
            [PSCustomObject]@{
                Id          = $selectedProcess.ProcessId
                Name        = $selectedProcess.Name
                CommandLine = $selectedProcess.CommandLine
            }
        }
    }
    finally {
        foreach ($win32Proc in $win32Processes) {
            $win32Proc.Dispose()
        }
    }
}

<#
.SYNOPSIS
    Get processes with the specified module. The module name is matched with a regex pattern

.EXAMPLE
    Get-ProcessWithModule -ModuleNamePattern 'olmapi32'
    Get processes with olmapi32
#>
function Get-ProcessWithModule {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$ModuleNamePattern
    )

    Get-Process | & {
        param(
            [Parameter(ValueFromPipeline)]
            [System.Diagnostics.Process]$Process
        )
        process {
            # See if this process contains the given module (name is regex matched)
            $mods = @(
                foreach ($module in $Process.Modules) {
                    if ($module.ModuleName -match $ModuleNamePattern) {
                        [PSCustomObject]@{
                            Name        = [IO.Path]::GetFileNameWithoutExtension($module.ModuleName)
                            Path        = $module.FileName
                            FileVersion = $module.FileVersionInfo.FileVersion # Note: This may not be available
                        }
                    }
                }
            )

            if ($mods.Count -eq 0) {
                return
            }

            [PSCustomObject]@{
                Id          = $Process.Id
                Name        = $Process.Name
                Path        = $Process.Path
                FileVersion = $Process.FileVersion
                Modules     = $mods
            }
        }
    }
}

<#
.SYNOPSIS
    Collect Microsoft Office Outlook related configuration & traces
.DESCRIPTION
    This will collect different kinds of traces & log files depending on the value specified in the "Component" parameter.
.EXAMPLE
    PS C:\> Collect-OutlookInfo -Path C:\temp -Component Configuration, Netsh, Outlook
    This will collect configuration data, Netsh trace, and Outlook ETW trace.
.LINK
    https://github.com/jpmessaging/OutlookTrace

#>
function Collect-OutlookInfo {
    [CmdletBinding(SupportsShouldProcess = $true, PositionalBinding = $false)]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseApprovedVerbs', '')]
    param (
        # Folder to place collected data
        [Parameter(Mandatory = $true, Position = 0)]
        $Path,
        # What to collect
        [Parameter(Mandatory = $true)]
        [ValidateSet('Outlook', 'Netsh', 'PSR', 'LDAP', 'CAPI', 'Configuration', 'Fiddler', 'TCO', 'Dump', 'CrashDump', 'HungDump', 'Procmon', 'WAM', 'WFP', 'TTD', 'Performance', 'WPR', 'Recording', 'NewOutlook', 'WebView2')]
        [array]$Component,
        # This controls the level of netsh trace report
        [ValidateSet('None', 'Mini', 'Full')]
        $NetshReportMode = 'None',
        # ETW trace file mode.
        [ValidateSet('NewFile', 'Circular')]
        [string]$LogFileMode = 'NewFile',
        # Max file size for ETW trace files. By default, 256 MB when NewFile and 2048 MB when Circular.
        [ValidateRange(1, [int]::MaxValue)]
        [int]$MaxFileSizeMB,
        # Archive type. Currently supports Zip or Cab. Zip is faster, but Cab is smaller.
        [ValidateSet('Zip', 'Cab')]
        [string]$ArchiveType = 'Zip',
        # Skip archiving
        [switch]$SkipArchive,
        # AutoFlush log file.
        [switch]$AutoFlush,
        # Skip running autoupdate of this script.
        [switch]$SkipAutoUpdate,
        # PSR recycle interval.
        [ValidateRange('00:01:00', '01:00:00')]
        [TimeSpan]$PsrRecycleInterval = [Timespan]::FromMinutes(10),
        # Target user whose configuration is collected. By default, it's the logon user (Note:Not necessarily the current user running the script).
        # [ArgumentCompleter({ Get-LogonUser })]
        [string]$User,
        # Timespan used to detect a hung window when "HungDump" is in Component.
        [ValidateRange('00:00:01', '00:01:00')]
        [TimeSpan]$HungTimeout = [TimeSpan]::FromSeconds(5),
        # Max number of hung dump files to be saved per process instance
        [ValidateRange(1, 10)]
        [int]$MaxHungDumpCount = 3,
        # Target process name (such as Outlook or olk). This is optional. 'Outlook' is used by default and 'olk' is used when 'NewOutlook' is specified in Component parameter.
        [string]$TargetProcessName,
        # Names of the target processes for crash dumps. When not specified, all processes will be the targets.
        [string[]]$CrashDumpTargets,
        # Switch to enable full page heap for the target process (With page heap, Outlook will consume a lot of memory and slow down)
        [switch]$EnablePageHeap,
        # Switch to add Microsoft.AAD.BrokerPlugin to Loopback Exempt
        [switch]$EnableLoopbackExempt,
        # Skip script version check.
        [switch]$SkipVersionCheck,
        # Command line filter for TTD monitor
        [string]$TTDCommandlineFilter,
        # Restrict TTD trace to specified modules (Must have extensions)
        [string[]]$TTDModules,
        # Switch to show TTD UI
        [switch]$TTDShowUI,
        [ValidateSet('GeneralProfile', 'CPU', 'DiskIO', 'FileIO', 'Registry', 'Network', 'Heap', 'Pool', 'VirtualAllocation', 'Audio', 'Video', 'Power', 'InternetExplorer', 'EdgeBrowser', 'Minifilter', 'GPU', 'Handle', 'XAMLActivity', 'HTMLActivity', 'DesktopComposition', 'XAMLAppResponsiveness', 'HTMLResponsiveness', 'ReferenceSet', 'ResidentSet', 'XAMLHTMLAppMemoryAnalysis', 'UTC', 'DotNET', 'WdfTraceLoggingProvider', 'HeapSnapshot')]
        # WPR profiles
        [string[]]$WprProfiles = @('GeneralProfile', 'CPU', 'DiskIO', 'FileIO', 'Registry', 'Network'),
        # Switch to remove cached identites & authentication tokens
        [switch]$RemoveIdentityCache
    )

    $runAsAdmin = Test-RunAsAdministrator

    # Explicitly check admin rights depending on the request.
    if (-not $runAsAdmin -and (($Component -join ' ') -match 'Outlook|Netsh|CAPI|LDAP|WAM|WPR|WFP|CrashDump|TTD' -or $EnablePageHeap -or $EnableLoopbackExempt -or $RemoveIdentityCache)) {
        Write-Warning "Please run as administrator"
        return
    }

    # Try to enable Debug Privilege if running as admin (It's possible that the user does not have SeDebugPrivilege)
    $debugPrivilegeEnabled = $false
    $debugPrivilegeError = $null

    if ($runAsAdmin) {
        $debugPrivilegeError = Enable-DebugPrivilege 2>&1
        $debugPrivilegeEnabled = $null -eq $debugPrivilegeError
    }

    if ($env:PROCESSOR_ARCHITEW6432) {
        Write-Error "32-bit PowerShell is running on 64-bit OS. Please use 64-bit PowerShell from C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe"
        return
    }

    if (-not $SkipVersionCheck -and -not (Test-ScriptExpiration)) {
        Write-Error "This script is too old. The script version is $Version and it has passed $($Script:ValidTimeSpan.Days) days.`nPlease download the latest version from https://github.com/jpmessaging/OutlookTrace.`nYou can skip this check by using -SkipVersionCheck switch"
        return
    }

    # MS Office must be installed to collect Outlook and TCO.
    # This is just a fail fast. Start-OutlookTrace/TCOTrace fail anyway.
    if ($Component -contains 'Outlook' -or $Component -contains 'TCO') {
        $err = $($null = Get-OfficeInfo -ErrorAction Continue) 2>&1

        if ($err) {
            Write-Error "Component `"Outlook`" or `"TCO`" is specified, but Microsoft Office is not installed"
            return
        }
    }

    $currentUser = Resolve-User ([System.Security.Principal.WindowsIdentity]::GetCurrent().Name)

    # If User is given, use it as the target user; Otherwise, use the logon user.
    if ($PSBoundParameters.ContainsKey('User')) {
        $targetUser = Resolve-User $User

        if (-not $targetUser) {
            return
        }
    }
    else {
        # Get logon users
        $logonUserErrors = $($logonUsers = @(Get-LogonUser)) 2>&1

        if ($logonUserErrors) {
            Write-Error "Failed to get logon users.`n$($logonUserErrors -join "`n")"
            return
        }

        if ($logonUsers.Count -eq 0) {
            Write-Error "Cannot find any logon user"
            return
        }
        elseif ($logonUsers.Count -eq 1) {
            $targetUser = $logonUsers[0]
        }
        else {
            # Multiple logon users are found. 'User' parameter needs to be used.
            Write-Error "Found multiple logon users ($($logonUsers.Count) users$(if ($logonUsers.Count -le 3) { "; $($logonUsers.Name -join ',')" })). Please specify the target user by `"-User`" parameter"
            return
        }
    }

    if ($Component -contains 'TTD') {
        # For TTD, the current user must have Debug Privilege (even for his/her own processes)
        if (-not $debugPrivilegeEnabled) {
            Write-Error "Cannot collect TTD because the current user `"$currentUser`" does not have Debug Privilege"
            return
        }

        # Validate TTDModules have extension
        $invalidModules = $TTDModules | Where-Object { -not [IO.Path]::GetExtension($_) }

        if ($invalidModules) {
            Write-Error "Module names in TTDModules parameter must have an extension. Invalid module names:$($invalidModules -join ', ')"
            return
        }
    }

    # Collecting Dump/HungDump is possible when the current user has DebugPrivilege or the target process's user is the same as the current user.
    # Otherwise process's handle is not available (crash dumps are OK because they are collected by WER)
    # Note: Even if this check passes, Save-Dump might still fail if the current user does not have Debug Privilege and the target process is running by a different user.
    if (($Component -contains "Dump" -or $Component -contains 'HungDump') -and ($targetUser.Sid -ne $currentUser.Sid -and -not $debugPrivilegeEnabled)) {
        Write-Error -Message "Cannot collect process dump files (Dump or HungDump) for the target user `"$targetUser`" because the current user `"$currentUser`" does not have Debug Privilege"
        return
    }

    if (-not $SkipAutoUpdate) {
        # Save a copy of Get-CommandExpression because it is not exported and it becomes unavailable after reloading the module.
        ${Get-CommandExpression} = (Get-Command Get-CommandExpression).ScriptBlock

        $autoUpdate = Invoke-AutoUpdate

        if ($autoUpdate.Success) {
            Write-Warning "$([IO.Path]::GetFileName($PSCommandPath)) was auto updated from $Version to $($autoUpdate.NewVersion)"

            # Run with a new powershell.exe or pwsh.exe instance instead of invoking again, because native interop might be updated in the new version.
            $PSBoundParameters.Add('SkipAutoUpdate', $true)
            $expression = & ${Get-CommandExpression} -Command $MyInvocation.MyCommand -Parameters $PSBoundParameters

            $currentProcess = Get-Process -Id $PID
            $powerShellExe = $currentProcess.Path
            $currentProcess.Dispose()
            $fileName = [IO.Path]::GetFileName($powerShellExe)

            # If running with other than powershell.exe or pwsh.exe (such as powershell_ise.exe), then warn the user to start over.
            if ($fileName -notmatch '(powershell|pwsh)\.exe') {
                Write-Warning "Please close $fileName and import the module again"
                return
            }

            & $powerShellExe -NoLogo -NoExit -NoProfile -NonInteractive -Command "& { Import-Module $PSCommandPath -DisableNameChecking -Force -ErrorAction Stop; $expression}"
            return
        }
    }

    if (-not (Test-Path $Path -ErrorAction Stop)) {
        $null = New-Item -ItemType Directory $Path -ErrorAction Stop
    }

    # Create a temporary folder to store data.
    $Path = Convert-Path -LiteralPath $Path
    $tempPath = Join-Path $Path -ChildPath $([Guid]::NewGuid().ToString())
    $null = New-Item $tempPath -ItemType directory -ErrorAction Stop

    # Start logging.
    Open-Log -Path (Join-Path $tempPath 'Log.csv') -WithBOM -AutoFlush:$AutoFlush -ErrorAction Stop
    Write-Log "Script Version:$Script:Version (Module Version $($MyInvocation.MyCommand.Module.Version.ToString())); PID:$pid"
    Write-Log "PSVersion:$($PSVersionTable.PSVersion); CLRVersion:$($PSVersionTable.CLRVersion)"
    Write-Log "PROCESSOR_ARCHITECTURE:$env:PROCESSOR_ARCHITECTURE; PROCESSOR_ARCHITEW6432:$env:PROCESSOR_ARCHITEW6432"
    Write-Log "Running as $($currentUser.Name) ($($currentUser.Sid)); RunningAsAdmin:$runAsAdmin; DebugPrivilegeEnabled:$debugPrivilegeEnabled"
    Write-Log "Target user:$($targetUser.Name) ($($targetUser.Sid))"
    Write-Log "AutoUpdate:$(if ($SkipAutoUpdate) { 'Skipped due to SkipAutoUpdate switch' } else { $autoUpdate.Message })"

    $invocation = Get-CommandExpression -Invocation $MyInvocation
    Write-Log "Invocation:$invocation"

    if ($runAsAdmin -and -not $debugPrivilegeEnabled) {
        Write-Log -Message "Running as admin, but failed to enable Debug Privilege. $debugPrivilegeError" -ErrorRecord $debugPrivilegeError -Category Error
    }

    # Determine TargetProcessName (must be without extension)
    if ($TargetProcessName) {
        $TargetProcessName = [IO.Path]::GetFileNameWithoutExtension($TargetProcessName)
    }
    else {
        if ($Component -contains 'NewOutlook') {
            $TargetProcessName = 'olk'
        }
        else {
            $TargetProcessName = 'outlook'
        }
    }

    Write-Log "TargetProcessName: $TargetProcessName"

    $ScriptInfo = [PSCustomObject]@{
        Version    = $Script:Version
        Invocation = $invocation
        RunAsUser  = "$($currentUser.Name) ($($currentUser.Sid))"
        TargetUser = "$($targetUser.Name) ($($targetUser.Sid))"
        Start      = [DateTimeOffSet]::Now
        WaitStart  = $null
        WaitStop   = $null
        End        = $null
    }

    try {
        # Set thread culture to en-US for consitent logging.
        Set-ThreadCulture 'en-US' 2>&1 | Write-Log -Category Error

        # Disable Ctrl+C temporarily while collecting data (this does not work in PowerShell ISE)
        $ctrlCDisabled = Disable-CtrlC

        if ($ctrlCDisabled) {
            Write-Log "Ctrl+C is successfully disabled"
        }

        # To use Start-Task, make sure to open runspaces first and close it when finished.
        # Currently MaxRunspaces is 8 or more because there are 8 tasks at most. 3 of them, processCaptureTask, psrTask, and hungMonitorTask are long running.
        $minimumMaxRunspacesCount = 8
        $vars = 'LogWriter', 'PSDefaultParameterValues', 'MyModulePath', 'Emoji' | Get-Variable
        Open-TaskRunspace -Variables $vars -MinRunspaces ([int]$env:NUMBER_OF_PROCESSORS) -MaxRunspaces ([math]::Max($minimumMaxRunspacesCount, (2 * [int]$env:NUMBER_OF_PROCESSORS)))

        # Configure log file mode and max file size for ETW traces (OutlookTrace, WAM, LDAP, and CAPI)
        $PSDefaultParameterValues['Start-*Trace:LogFileMode'] = $LogFileMode

        if ($PSBoundParameters.ContainsKey('MaxFileSizeMB')) {
            $PSDefaultParameterValues['Start-*Trace:MaxFileSizeMB'] = $MaxFileSizeMB
        }
        else {
            # MaxFileSizeMB is not specified by the user. Use default value depending on the log mode.
            if ($LogFileMode -eq 'NewFile') {
                $PSDefaultParameterValues['Start-*Trace:MaxFileSizeMB'] = 256
            }
            else {
                $PSDefaultParameterValues['Start-*Trace:MaxFileSizeMB'] = 2048
            }
        }

        if ($RemoveIdentityCache) {
            Invoke-ScriptBlock { param($User) Remove-IdentityCache @PSBoundParameters } -ArgumentList @{ User = $targetUser }
        }

        # Enable PageHeap for the target process
        if ($EnablePageHeap) {
            $pageHeapEnabled = Enable-PageHeap -ProcessName $TargetProcessName -ErrorAction Stop
        }

        # List of "Started" events for Wait & Close all at once.
        $startedEventList = New-Object System.Collections.Generic.List[System.Threading.EventWaitHandle]

        Write-Log "Starting traces"

        if ($Component -contains 'Configuration') {
            # Sub directories
            $ConfigDir = Join-Path $tempPath 'Configuration'
            $OSDir = Join-Path $ConfigDir 'OS'
            $OfficeDir = Join-Path $ConfigDir 'Office'
            $RegistryDir = Join-Path $ConfigDir 'Registry'
            $NetworkDir = Join-Path $ConfigDir 'Network'
            $MSIPCDir = Join-Path $ConfigDir 'MSIPC'
            $EventDir = Join-Path $ConfigDir 'EventLog'

            $null = New-Item -Path $OSDir -ItemType Directory -ErrorAction Stop
            $targetUser | Export-Clixml -Path (Join-Path $OSDir 'User.xml')

            $PSDefaultParameterValues['Write-Progress:Activity'] = 'Saving Configuration'
            $PSDefaultParameterValues['Write-Progress:Status'] = 'Please wait'
            Write-Progress -PercentComplete 0

            # First start tasks that might take a while.
            # Note:I could use ${Function:***}, but wrapping in a script block allows Write-Log to find the actual function name.
            # Also use $PSBoundParameters instead of $Args to forward arguments because $Args does not work for switch parameters.
            Write-Log "Starting OfficeModuleInfoTask"
            $officeModuleInfoTaskCts = New-Object System.Threading.CancellationTokenSource
            $officeModuleInfoTask = Start-Task -Name 'OfficeModuleInfoTask' -ScriptBlock { param($Path, $CancellationToken) Save-OfficeModuleInfo @PSBoundParameters } `
                -ArgumentList @{
                Path              = $OfficeDir
                CancellationToken = $officeModuleInfoTaskCts.Token
            }

            Write-Log "Starting NetworkInfoTask"
            $networkInfoTaskCts = New-Object System.Threading.CancellationTokenSource
            $networkInfoTask = Start-Task -Name 'NetworkInfoTask' -ScriptBlock { param($Path, $CancellationToken) Save-NetworkInfo @PSBoundParameters } `
                -ArgumentList @{
                Path              = $NetworkDir
                CancellationToken = $networkInfoTaskCts.Token
            }

            Write-Progress -PercentComplete 20

            Write-Log "Starting OfficeRegistryTask"
            $officeRegistryTask = Start-Task -Name 'OfficeRegistryTask' -ScriptBlock { param($Path, $User) Save-OfficeRegistry @PSBoundParameters } `
                -ArgumentList @{
                Path = $RegistryDir
                User = $targetUser
            }

            Write-Log "Starting OSConfigurationTask"
            $osConfigurationTaskCts = New-Object System.Threading.CancellationTokenSource
            $osConfigurationTask = Start-Task -Name 'OSConfigurationTask' -ScriptBlock { param($Path, $User, $CancellationToken) Save-OSConfiguration @PSBoundParameters } `
                -ArgumentList @{
                Path              = $OSDir
                User              = $targetUser
                CancellationToken = $osConfigurationTaskCts.Token
            }

            Write-Progress -PercentComplete 40

            Write-Log "Starting ProcessCaptureTask"
            $processCaptureTaskCts = New-Object System.Threading.CancellationTokenSource
            $processCaptureStartedEvent = New-Object System.Threading.EventWaitHandle($false, [Threading.EventResetMode]::ManualReset)
            $startedEventList.Add($processCaptureStartedEvent)
            $processCaptureTask = Start-Task -Name 'ProcessCaptureTask' -ScriptBlock { param($Path, $NamePattern, $CancellationToken, $StartedEvent) Start-ProcessCapture @PSBoundParameters } `
                -ArgumentList @{
                Path              = $OSDir
                NamePattern       = 'outlook|fiddler|explorer|backgroundTaskHost'
                CancellationToken = $processCaptureTaskCts.Token
                StartedEvent      = $processCaptureStartedEvent
            }

            Write-Log "Starting GPResultTask"
            $gpresultTaskCts = New-Object System.Threading.CancellationTokenSource
            $gpresultTask = Start-Task -Name 'GPResultTask' -ScriptBlock { param($Path, $User, $Format, $CancellationToken) Save-GPResult @PSBoundParameters } `
                -ArgumentList @{
                Path              = $OSDir
                User              = $targetUser
                Format            = 'HTML'
                CancellationToken = $gpresultTaskCts.Token
            }

            Write-Progress -PercentComplete 60

            $PSDefaultParameterValues['Invoke-ScriptBlock:Path'] = $OfficeDir
            Invoke-ScriptBlock { Get-OfficeInfo }
            Invoke-ScriptBlock { Get-ClickToRunConfiguration }
            Invoke-ScriptBlock { Get-PresentationMode }
            $PSDefaultParameterValues['Invoke-ScriptBlock:ArgumentList'] = @{ User = $targetUser }
            Invoke-ScriptBlock { param($User) Get-OutlookAddin @PSBoundParameters }
            Invoke-ScriptBlock { param($User) Get-OutlookOption @PSBoundParameters }
            Invoke-ScriptBlock { param($User) Get-AutodiscoverConfig @PSBoundParameters }
            Invoke-ScriptBlock { param($User) Get-SocialConnectorConfig @PSBoundParameters }
            Invoke-ScriptBlock { param($User) Get-IMProvider @PSBoundParameters }
            Invoke-ScriptBlock { param($User) Get-AlternateId @PSBoundParameters }
            Invoke-ScriptBlock { param($User) Get-UseOnlineContent @PSBoundParameters }
            Invoke-ScriptBlock { param($User) Get-OfficeIdentityConfig @PSBoundParameters }
            Invoke-ScriptBlock { param($User) Get-OfficeIdentity @PSBoundParameters }
            Invoke-ScriptBlock { param($User) Get-PrivacyPolicy @PSBoundParameters }
            Invoke-ScriptBlock { param($User) Get-MapiCorruptFiles @PSBoundParameters }
            Invoke-ScriptBlock { param($User) Get-ExperimentConfigs -AppName 'outlook' @PSBoundParameters }
            Invoke-ScriptBlock { param($User) Get-CloudSettings @PSBoundParameters }
            Invoke-ScriptBlock { param($User) Get-DRMConfig @PSBoundParameters }
            $PSDefaultParameterValues.Remove('Invoke-ScriptBlock:ArgumentList')
            $PSDefaultParameterValues.Remove('Invoke-ScriptBlock:Path')

            Write-Progress -PercentComplete 80

            Invoke-ScriptBlock { param($User, $Path) Save-CachedAutodiscover @PSBoundParameters } -ArgumentList @{ User = $targetUser; Path = Join-Path $OfficeDir 'Cached AutoDiscover' }
            Invoke-ScriptBlock { param($User, $Path) Save-CachedOutlookConfig @PSBoundParameters } -ArgumentList @{ User = $targetUser; Path = Join-Path $OfficeDir 'Cached OutlookConfig' }
            Invoke-ScriptBlock { param($User, $Path) Save-PolicyNudge @PSBoundParameters } -ArgumentList @{ User = $targetUser; Path = Join-Path $OfficeDir 'PolicyNudge' }
            Invoke-ScriptBlock { param($User, $Path) Save-DLP @PSBoundParameters } -ArgumentList @{ User = $targetUser; Path = Join-Path $OfficeDir 'DLP' }
            Invoke-ScriptBlock { param($User, $Path) Save-CLP @PSBoundParameters } -ArgumentList @{ User = $targetUser; Path = Join-Path $OfficeDir 'CLP' }

            # Gather WinINet related data in case Fiddler is started later.
            Invoke-ScriptBlock { param($User) Get-WinInetProxy @PSBoundParameters } -ArgumentList @{ User = $targetUser } -Path $OSDir
            Invoke-ScriptBlock { param($User) Get-ProxyAutoConfig @PSBoundParameters } -ArgumentList @{ User = $targetUser } -Path $OSDir

            Write-Progress -Completed
        }

        # Check if Microsoft.AAD.BrokerPlugin is available.
        $($brokerPlugin = Get-AADBrokerPlugin) 2>&1 | Write-Log -Category Warning

        # Add Microsoft.AAD.BrokerPlugin to Loopback Exempt list if that's appropriate. If it is already added, Add-LoopbackExempt does nothing.
        if ($brokerPlugin -and ($EnableLoopbackExempt -or $Component -contains 'Fiddler' -or $Component -contains 'Netsh')) {
            $($loopbackExemptAdded = Add-LoopbackExempt $brokerPlugin.PackageFamilyName) 2>&1 | Write-Log -Category Warning

            if ($loopbackExemptAdded -and $OSDir) {
                CheckNetIsolation.exe LoopbackExempt -s | Set-Content -Path (Join-Path $OSDir 'LoopbackExempt.txt')
            }
        }

        $PSDefaultParameterValues['Write-Progress:Activity'] = 'Starting traces'

        if ($Component -contains 'Fiddler') {
            Write-Progress -Status 'Starting Fiddler'

            if ($targetUser.Sid -eq $currentUser.Sid) {
                # $null = Start-FiddlerCap -Path $Path -ErrorAction Stop
                # Write-Warning "FiddlerCap has started. Please manually configure and start capture"
                $null = Start-FiddlerEverywhereReporter -Path $Path -ErrorAction Stop
                Write-Warning "Fiddler Everywhere Reporter has started. Please manually configure and start capture"
            }
            else {
                # If target user is different from current user, don't start FiddlerCap because it won't be able to capture (WinInet proxy needs to be configured for the target user).
                # $fiddler = Start-FiddlerCap -Path $Path -ErrorAction Stop -CheckAvailabilityOnly
                # Write-Warning "Let the user ($($targetUser.Name)) start $($fiddler.FiddlerPath)"

                $fiddler = Start-FiddlerEverywhereReporter -Path $Path -CheckAvailabilityOnly -ErrorAction Stop
                Write-Warning "Let the user ($($targetUser.Name)) start $($fiddler.FiddlerPath)"
            }

            $fiddlerStarted = $true
        }

        if ($Component -contains 'Netsh') {
            Write-Progress -Status 'Starting Netsh trace'

            # When netsh trace is run for the first time, it does not capture packets (even with "capture=yes").
            # To workaround, netsh is started and stopped immediately.
            $tempNetshName = 'netsh_test'
            Start-NetshTrace -Path (Join-Path $tempPath $tempNetshName) -FileName "$tempNetshName.etl" -ReportMode 'None' -ErrorAction Stop
            Stop-NetshTrace -ErrorAction Stop
            Remove-Item (Join-Path $tempPath $tempNetshName) -Recurse -Force -ErrorAction SilentlyContinue

            Start-NetshTrace -Path (Join-Path $tempPath 'Netsh') -ReportMode $NetshReportMode -ErrorAction Stop
            $netshTraceStarted = $true
        }

        if ($Component -contains 'Outlook') {
            Write-Progress -Status 'Starting Outlook trace'
            # Stop a lingering session if any.
            Stop-OutlookTrace -ErrorAction SilentlyContinue

            $err = Enable-DrmExtendedLogging -User $targetUser 2>&1

            if ($err) {
                Write-Log -Message "Enable-DrmExtendedLogging failed. $err" -ErrorRecord $err -Category Error
            }

            Start-OutlookTrace -Path (Join-Path $tempPath 'Outlook') -ErrorAction Stop
            $outlookTraceStarted = $true
        }

        if ($Component -contains 'NewOutlook') {
            Enable-WebView2DevTools -ExecutableName 'olk.exe' -User $targetUser -ErrorAction Stop
            $newOutlookTraceStarted = $true
        }

        if ($Component -contains 'WebView2') {
            Enable-WebView2Netlog -ExecutableName $TargetProcessName -User $targetUser -Path (Join-Path $tempPath 'WebView2') -MaxFileSizeMB 2048 -ErrorAction Stop
            $webView2TraceStarted = $true
        }

        if ($Component -contains 'PSR') {
            Write-Progress -Status 'Starting PSR'

            # Make sure psr isn't running already.
            $psrProcesses = @(Get-Process psr -ErrorAction SilentlyContinue)

            if ($psrProcesses.Count -gt 0) {
                Write-Error "PSR is already running (PID:$($psrProcesses.ID -join ',')).`nPlease stop PSR first and run again"
                return
            }

            $psrCts = New-Object System.Threading.CancellationTokenSource
            $psrStartedEvent = New-Object System.Threading.EventWaitHandle($false, [Threading.EventResetMode]::ManualReset)
            $startedEventList.Add($psrStartedEvent)
            Write-Log "Starting PSRTask. PsrRecycleInterval:$PsrRecycleInterval"

            $psrTask = Start-Task -Name 'PSRTask' -ScriptBlock { param($Path, $CancellationToken, $WaitInterval, $StartedEvent, $Circular) Start-PsrMonitor @PSBoundParameters } `
                -ArgumentList @{
                Path              = Join-Path $tempPath 'PSR'
                CancellationToken = $psrCts.Token
                WaitInterval      = $PsrRecycleInterval
                StartedEvent      = $psrStartedEvent
                Circular          = $LogFileMode -eq 'Circular'
            }

            $psrStarted = $true
        }

        if ($Component -contains 'LDAP') {
            Start-LDAPTrace -Path (Join-Path $tempPath 'LDAP') -TargetExecutable $TargetProcessName
            $ldapTraceStarted = $true
        }

        if ($Component -contains 'CAPI') {
            Enable-EventLog 'Microsoft-Windows-CAPI2/Operational' | Add-EventLogConfigCache
            Start-CAPITrace -Path (Join-Path $tempPath 'CAPI') -ErrorAction Stop
            $capiTraceStarted = $true
        }

        if ($Component -contains 'TCO') {
            Start-TCOTrace
            $tcoTraceStarted = $true
        }

        if ($Component -contains 'WAM') {
            Write-Progress -Status 'Starting WAM trace'
            Enable-WamEventLog -ErrorAction SilentlyContinue
            Stop-WamTrace -ErrorAction SilentlyContinue
            Start-WamTrace -Path (Join-Path $tempPath 'WAM') -ErrorAction Stop
            $wamTraceStarted = $true
        }

        if ($Component -contains 'Procmon') {
            Write-Progress -Status 'Starting Procmon'
            $null = Start-Procmon -Path (Join-Path $tempPath 'Procmon') -ProcmonSearchPath $Path -ErrorAction Stop
            $procmonStared = $true
        }

        if ($Component -contains 'WFP') {
            Write-Progress -Status 'Starting WFP trace'
            $wfpJob = Start-WfpTrace -Path (Join-Path $tempPath 'WFP') -Interval ([TimeSpan]::FromSeconds(15))
            $wfpStarted = $true
        }

        if ($Component -contains 'Performance') {
            Write-Progress -Status 'Starting performance trace'
            Start-PerfTrace -Path (Join-Path $tempPath 'Performance') -ErrorAction Stop
            $perfStarted = $true
        }

        if ($Component -contains 'WPR') {
            Write-Progress -Status 'Starting WPR trace'
            Start-Wpr -Path (Join-Path $tempPath 'WPR') -Profiles $WprProfiles -ErrorAction Stop
            $wprStarted = $true
        }

        if ($Component -contains 'CrashDump') {
            $CrashDumpTargets | Add-WerDumpKey -Path (Join-Path $tempPath 'WerDump')
            $crashDumpStarted = $true
        }

        if ($Component -contains 'Dump') {
            # Close the progress bar for now.
            Write-Progress -Completed

            # Ask a user when she/he wants to save a dump file
            while ($true) {
                Write-Host "Press enter to save a process dump of $TargetProcessName. To quit, press q: " -NoNewline
                $userInput = $host.UI.ReadLine()

                if ($userInput.ToLower().Trim() -eq 'q') {
                    break
                }

                $process = Get-SingleProcess -Name $TargetProcessName -User $targetUser

                if (-not $process) {
                    Write-Host "Cannot find $TargetProcessName. Please start $TargetProcessName" -ForegroundColor Yellow
                    continue
                }

                $activity = "Saving a process dump of $TargetProcessName"
                Write-Progress -Activity $activity -Status "Please wait" -PercentComplete -1

                $err = $($dumpResult = Save-Dump -Path (Join-Path $tempPath 'Dump') -ProcessId $process.Id) 2>&1

                Write-Progress -Activity $activity -Status "Done" -Completed

                if ($dumpResult) {
                    Write-Host "A dump file of $($dumpResult.ProcessName) (PID:$($dumpResult.ProcessId)) was successfully saved" -ForegroundColor Green
                    Write-Log "Saved a dump file of $($dumpResult.ProcessName) (PID:$($dumpResult.ProcessId)):$($dumpResult.DumpFile)"
                }
                else {
                    Write-Error "Failed to save a dump file of $($dumpResult.ProcessName) (PID:$($dumpResult.ProcessId)). $err"
                    Write-Log -Message "Save-Dump failed for $($dumpResult.ProcessName) (PID:$($dumpResult.ProcessId)). $err" -ErrorRecord $err -Category Error
                }
            }
        }

        if ($Component -contains 'HungDump') {
            Write-Progress -Status 'Starting HungMonitorTask'
            $hungDumpCts = New-Object System.Threading.CancellationTokenSource
            $monitorStartedEvent = New-Object System.Threading.EventWaitHandle($false, [Threading.EventResetMode]::ManualReset)
            $startedEventList.Add($monitorStartedEvent)

            $hungMonitorArgs = @{
                Path              = Join-Path $tempPath 'HungDump'
                Name              = $TargetProcessName
                User              = $targetUser
                Timeout           = $HungTimeout
                DumpCount         = $MaxHungDumpCount
                CancellationToken = $hungDumpCts.Token
                StartedEvent      = $monitorStartedEvent
            }

            Write-Log "Starting HungMonitorTask. TargetProcessName:$($hungMonitorArgs.Name), HungTimeout:$($hungMonitorArgs.Timeout), User:$targetUser"
            $hungMonitorTask = Start-Task -Name 'HungMonitorTask' -ScriptBlock { param($Path, $Name, $User, $Timeout, $DumpCount, $CancellationToken, $StartedEvent) Start-HungMonitor @PSBoundParameters } -ArgumentList $hungMonitorArgs
            $hungDumpStarted = $true
        }

        if ($Component -contains 'TTD') {
            Write-Progress -Status 'Starting TTD'

            # Download TTD. Instead of installing TTD msix, expand it.
            $ttdPath = Download-TTD -Path $Path -ErrorAction Stop | Expand-TTDMsixBundle | Select-Object -ExpandProperty 'TTDPath'

            # Log TTD version
            $version = Get-ItemProperty $ttdPath | Select-Object -ExpandProperty 'VersionInfo'
            Write-Log "Using $ttdPath (Version:$($version.FileVersion))"

            # Common args for Attach-TTD and Start-TTDMonitor.
            $ttdArgs = @{
                TTDPath = $ttdPath
                Path    = Join-Path $tempPath 'TTD'
                Modules = $TTDModules
                ShowUI  = $TTDShowUI
            }

            # If the target process (Outlook, olk, etc) is already running, attach to it. Otherwise, start monitoring
            $process = Get-SingleProcess -Name $TargetProcessName -User $targetUser -CommandLineFilter $TTDCommandlineFilter

            if ($process) {
                $logMsg = "Attaching TTD to $($process.Name) (PID:$($process.Id))"
                Write-Log $logMsg

                $ttdArgs.ProcessId = $process.Id

                # Attach TTD with a background task and asynchronously wait
                $attachTask = Start-Task -Name 'AttachTTD' -ScriptBlock { param($TTDPath, $Path, $ProcessID, $Modules, $ShowUI) Attach-TTD @PSBoundParameters } -ArgumentList $ttdArgs

                $waitMsg = "$logMsg. This might take a while. Please wait"
                $waitCount = 0
                $waitInterval = [TimeSpan]::FromSeconds(1)

                # Wait for the attach task to finish with a progress notification
                while (-not (Wait-Task -Task $attachTask -Timeout $waitInterval)) {
                    Write-Progress -Status ($waitMsg + '.' * $waitCount)
                    $waitCount = ++$waitCount % 10
                }

                $ttdProcess = Receive-Task -Task $attachTask -AutoRemoveTask -ErrorAction Stop
            }
            else {
                $ttdArgs.ExecutableName = $TargetProcessName
                $ttdArgs.CommandlineFilter = $TTDCommandlineFilter

                # If, for some reason, lingering "TTD.exe -monitor" is running, stop it first
                Stop-TTDMonitor -ErrorAction SilentlyContinue

                $ttdProcess = Start-TTDMonitor @ttdArgs -ErrorAction Stop
            }

            $ttdProcess | Add-Member -MemberType NoteProperty -Name 'TargetProcessName' -Value $TargetProcessName
            $ttdStarted = $true

            # Set ConnTimeout registry value
            $savedConnTimeout = Get-ConnTimeout
            $connTimeout = Set-ConnTimeout -User $targetUser -Value ([TimeSpan]::FromMinutes(5))
            Write-Log "ConnTimeout is set to $($connTimeout.ConnTimeout). Original value is $(if ($null -eq $savedConnTimeout.ConnTimeout) { "null" } else { $savedConnTimeout.ConnTimeout })"
        }

        if ($Component -contains 'Recording') {
            $recording = Start-Recording -ZoomItDownloadPath (Join-Path $Path 'ZoomIt') -ZoomItSearchPath $Path -ErrorAction Stop
            $recordingStarted = $true
        }

        # Wait all "Started" events
        if ($startedEventList.Count -gt 0) {
            Write-Progress -Status "Waiting for all tasks to start"
            Write-Log "Waiting for all tasks to start"

            foreach ($event in $startedEventList) {
                $null = $event.WaitOne()
                $event.Dispose()
            }
        }

        Write-Progress -Completed
        $waitStart = Get-Timestamp
        $waitResult = $null

        if ($netshTraceStarted -or $outlookTraceStarted -or $psrStarted -or $ldapTraceStarted -or $capiTraceStarted -or $tcoTraceStarted -or $fiddlerStarted -or $crashDumpStarted -or $procmonStared -or $wamTraceStarted -or $wfpStarted -or $ttdStarted -or $perfStarted -or $hungDumpStarted -or $wprStarted -or $recordingStarted -or $newOutlookTraceStarted -or $webView2TraceStarted) {
            Write-Log "Waiting for the user to stop"
            $ScriptInfo.WaitStart = [DateTimeOffset]::Now

            Write-Host 'Press enter to stop: ' -NoNewline
            $waitResult = Wait-EnterOrControlC

            if ($waitResult.Key -eq 'Ctrl+C') {
                Write-Warning "Ctrl+C is detected"
            }
        }

        if (-not $waitResult -or $waitResult.Key -eq 'Enter') {
            $startSuccess = $true
        }
    }
    catch {
        # Log & save the exception so that I can analyze later. Then rethrow.
        Write-Log "Terminating Error occured while staring traces" -ErrorRecord $_ -Category Error
        Export-CliXml -InputObject $_ -Path (Join-Path $tempPath 'Exception.xml')
        throw
    }
    finally {
        Write-Log "Stopping traces. $(if ($waitStart) { "Wait duration:$(Get-Elapsed $waitStart)" })"
        $ScriptInfo.WaitStop = [DateTimeOffset]::Now

        $PSDefaultParameterValues['Write-Progress:Activity'] = 'Stopping traces'

        if ($Local:recordingStarted) {
            # This will show the user a Save As dialog
            Stop-Recording
            Write-Host "Please save the recording (Save As dialog should appear). Then press enter to continue:" -ForegroundColor Yellow -NoNewline
            $null = $host.UI.ReadLine()

            # If the zoomit was started by above, then kill it.
            if ($recording.Started) {
                $zoomIt = Get-Process -Name 'ZoomIt*' | Select-Object -First 1

                if ($zoomIt) {
                    $zoomIt.Kill()
                    $zoomIt.Dispose()
                    Write-Log "ZoomIt instance was killed"
                }
            }
        }

        if ($Local:ttdStarted) {
            Write-Progress -Status 'Stopping TTD trace'

            if ($ttdProcess.IsAttached) {
                $err = $($ttdProcess | Detach-TTD) 2>&1 | Write-Log -Category Error -PassThru
            }
            else {
                $err = $($ttdProcess | Stop-TTDMonitor) 2>&1 | Write-Log -Category Error -PassThru
            }

            # Stopping or detaching TTD might fail if TTD.exe died during tracing. In this case, ask the user to shutdown Outlook manually so that trace file is fully written.
            if ($err) {
                $outlookProcess = @(Get-Process -Name $ttdProcess.TargetProcessName -ErrorAction SilentlyContinue | `
                        Where-Object { $_.Modules | Where-Object { $_.ModuleName -match 'TTDRecordCPU' } })

                if ($outlookProcess.Count) {
                    Write-Host "Please shutdown $($outlookProcess[0].Name) (PID:$($outlookProcess.Id -join ','))" -ForegroundColor Yellow

                    foreach ($proc in $outlookProcess) {
                        Write-Host "Waiting for $($proc.Name) (PID:$($proc.Id)) to shutdown ..." -ForegroundColor Yellow
                        $proc.WaitForExit()
                        $proc.Dispose()
                    }
                }
            }

            $ttdProcess | Cleanup-TTD
        }

        if ($Local:netshTraceStarted) {
            Write-Progress -Status 'Stopping Netsh trace'
            Stop-NetshTrace 2>&1 | Write-Log -Category Error -PassThru
        }

        if ($Local:outlookTraceStarted) {
            Write-Progress -Status 'Stopping Outlook trace'
            Stop-OutlookTrace 2>&1 | Write-Log -Category Error -PassThru
            Disable-DrmExtendedLogging -User $targetUser
        }

        if ($Local:newOutlookTraceStarted) {
            Disable-WebView2DevTools -ExecutableName 'olk.exe' -User $targetUser 2>&1 | Write-Log -Category Error -PassThru
            Save-MonarchLog -User $targetUser -Path (Join-Path $tempPath 'Monarch')  2>&1 | Write-Log -Category Error -PassThru
            Save-MonarchSetupLog -User $targetUser -Path (Join-Path $tempPath 'MonarchSetup')  2>&1 | Write-Log -Category Error -PassThru
        }

        if ($Local:webView2TraceStarted) {
            Disable-WebView2Netlog -ExecutableName $TargetProcessName -User $targetUser -ErrorAction Stop

            # The target process (and its WebView2 instances) must be shutdown so that the netlog will be written to the file.
            $processes = @(Get-Process -Name $TargetProcessName -ErrorAction SilentlyContinue)

            if ($processes.Count -gt 0) {
                Write-Host "Please shutdown $TargetProcessName (PID:$($processes.Id -join ','))" -ForegroundColor Yellow

                foreach ($proc in $processes) {
                    Write-Host "Waiting for $($proc.Name) (PID:$($proc.Id)) to shutdown ..." -ForegroundColor Yellow
                    $proc.WaitForExit()
                    $proc.Dispose()
                }
            }
        }

        if ($Local:ldapTraceStarted) {
            Write-Progress -Status 'Stopping LDAP trace'
            Stop-LDAPTrace -TargetExecutable $TargetProcessName 2>&1 | Write-Log -Category Error -PassThru
        }

        if ($Local:capiTraceStarted) {
            Write-Progress -Status 'Stopping CAPI trace'
            Stop-CAPITrace
        }

        if ($Local:tcoTraceStarted) {
            Stop-TcoTrace -Path (Join-Path $tempPath 'TCO')
        }

        if ($Local:wamTraceStarted) {
            Write-Progress -Status 'Stopping WAM trace'
            Stop-WamTrace
        }

        if ($Local:procmonStared) {
            Write-Progress -Status 'Stopping Procmon'
            Stop-Procmon 2>&1 | Write-Log -Category Error -PassThru
        }

        if ($Local:wfpStarted) {
            Write-Progress -Status 'Stopping WFP trace'
            Stop-WfpTrace $wfpJob
        }

        if ($Local:perfStarted) {
            Stop-PerfTrace 2>&1 | Write-Log -Category Error -PassThru
        }

        if ($Local:wprStarted) {
            Write-Progress -Status 'Stopping WPR trace'
            Stop-Wpr -Path (Join-Path $tempPath 'WPR') | Write-Log -Category Error -PassThru
        }

        if ($Local:hungDumpStarted) {
            $hungDumpCts.Cancel()
            Receive-Task $hungMonitorTask -AutoRemoveTask 2>&1 | Write-Log -Category Error
            $hungDumpCts.Dispose()
        }

        if ($Local:crashDumpStarted) {
            $CrashDumpTargets | Remove-WerDumpKey
        }

        if ($Local:fiddlerStarted) {
            # Write-Warning "Please stop FiddlerCap and save the capture manually"
            Write-Warning "Please stop Fiddler Everywhere Reporter and save the capture manually"
        }

        if ($Local:psrStarted) {
            Write-Progress -Status "Stopping PSR"
            $psrCts.Cancel()
            Receive-Task $psrTask -AutoRemoveTask 2>&1 | Write-Log -Category Error -PassThru
            $psrCts.Dispose()
        }

        # Restore Event Log configurations
        Get-EventLogConfigCache | Restore-EventLog 2>&1 | Write-Log -Category Error -PassThru
        Clear-EventLogConfigCache

        # Wait for the tasks started earlier and save the event logs
        if ($Component -contains 'Configuration') {
            if ($Local:processCaptureTask) {
                Write-Progress -Status 'Stopping Process capture task'
                $processCaptureTaskCts.Cancel()
                $processCaptureTask | Receive-Task -AutoRemoveTask 2>&1 | Write-Log -Category Error
                $processCaptureTaskCts.Dispose()
                Write-Log "$($processCaptureTask.Name) is complete"
            }

            Invoke-ScriptBlock { param($User) Get-OutlookProfile @PSBoundParameters } -ArgumentList @{ User = $targetUser } -Path $OfficeDir
            Invoke-ScriptBlock { param($User) Get-OneAuthAccount @PSBoundParameters } -ArgumentList @{ User = $targetUser } -Path $OfficeDir
            Invoke-ScriptBlock { param($Path, $User) Save-OneAuthAccount @PSBoundParameters } -ArgumentList @{ Path = Join-Path $OfficeDir 'OneAuthAccount'; User = $targetUser }

            if ($Local:startSuccess) {
                Write-Progress -Status 'Saving Event logs'
                Save-EventLog -Path $EventDir 2>&1 | Write-Log -Category Error

                Write-Progress -Status 'Saving MSIPC logs'
                Invoke-ScriptBlock { param($Path, $User, $All) Save-MSIPC @PSBoundParameters } -ArgumentList @{ Path = $MSIPCDir; User = $targetUser; All = $true }
                Invoke-ScriptBlock { param($Path, $User) Save-MIP @PSBoundParameters } -ArgumentList @{ User = $targetUser; Path = Join-Path $OfficeDir 'MIP' }
            }

            if ($Local:osConfigurationTask) {
                Write-Progress -Status 'Saving OS configuration'

                if (-not $Local:startSuccess) {
                    Write-Log "Canceling osConfigurationTask because startSuccess is false"
                    $osConfigurationTaskCts.Cancel()
                }

                $osConfigurationTask | Receive-Task -AutoRemoveTask 2>&1 | Write-Log -Category Error
                $osConfigurationTaskCts.Dispose()
                Write-Log "$($osConfigurationTask.Name) is complete"
            }

            if ($Local:officeRegistryTask) {
                Write-Progress -Status 'Saving Office Registry'
                $officeRegistryTask | Receive-Task -AutoRemoveTask 2>&1 | Write-Log -Category Error
                Write-Log "$($officeRegistryTask.Name) is complete"
            }

            if ($Local:networkInfoTask) {
                Write-Progress -Status 'Saving Network info'

                if (-not $Local:startSuccess) {
                    Write-Log "Canceling networkInfoTask because startSuccess is false"
                    $networkInfoTaskCts.Cancel()
                }

                $networkInfoTask | Receive-Task -AutoRemoveTask 2>&1 | Write-Log -Category Error
                $networkInfoTaskCts.Dispose()
                Write-Log "$($networkInfoTask.Name) is complete"
            }

            if ($Local:officeModuleInfoTask) {
                Write-Progress -Status "Saving Office module info"

                if ($Local:startSuccess) {
                    $timeout = [TimeSpan]::FromSeconds(30)

                    if (Wait-Task $officeModuleInfoTask -Timeout $timeout) {
                        Write-Log "$($officeModuleInfoTask.Name) is complete before timeout"
                    }
                    else {
                        Write-Log "$($officeModuleInfoTask.Name) timed out after $($timeout.TotalSeconds) seconds. Task will be canceled" -Category Warning
                        $officeModuleInfoTaskCts.Cancel()
                    }
                }
                else {
                    Write-Log "Canceling $($officeModuleInfoTask.Name) because startSuccess is false"
                    $officeModuleInfoTaskCts.Cancel()
                }

                $officeModuleInfoTask | Receive-Task -AutoRemoveTask 2>&1 | Write-Log -Category Error
                $officeModuleInfoTaskCts.Dispose()
                Write-Log "$($officeModuleInfoTask.Name) is complete"
            }

            if ($Local:gpresultTask) {
                Write-Progress -Status 'Saving Group Policy'

                if ($Local:startSuccess) {
                    # If collecting only Configuration, then do not timeout.
                    if ($Component.Count -eq 1) {
                        $timeout = [System.Threading.Timeout]::InfiniteTimeSpan
                    }
                    else {
                        $timeout = [TimeSpan]::FromSeconds(30)
                    }

                    if (Wait-Task -Task $gpresultTask -Timeout $timeout) {
                        Write-Log "$($gpresultTask.Name) is complete before timeout"
                    }
                    else {
                        Write-Log "$($gpresultTask.Name) timed out after $($timeout.TotalSeconds) seconds. Task will be canceled" -Category Warning
                        $gpresultTaskCts.Cancel()
                    }
                }
                else {
                    Write-Log "Canceling $($gpresultTask.Name) because startSuccess is false"
                    $gpresultTaskCts.Cancel()
                }

                $gpresultTask | Receive-Task -AutoRemoveTask 2>&1 | Write-Log -Category Error
                $gpresultTaskCts.Dispose()
                Write-Log "$($gpresultTask.Name) is complete"
            }
        }

        # Restore ConnTimeout
        if ($Local:savedConnTimeout) {
            # null value indicates the registry value did not exist
            if ($null -eq $savedConnTimeout.ConnTimeout) {
                Remove-ConnTimeout -User $targetUser
                Write-Log "ConnTimeout is removed"
            }
            else {
                $null = Set-ConnTimeout -User $targetUser -Value $savedConnTimeout.ConnTimeout
                Write-Log "ConnTimeout is restored to $($savedConnTimeout.ConnTimeout)"
            }
        }

        if ($Local:pageHeapEnabled) {
            Disable-PageHeap -ProcessName $TargetProcessName 2>&1 | Write-Log -Category Error -PassThru
        }

        if ($Local:loopbackExemptAdded) {
            Remove-LoopbackExempt $brokerPlugin.PackageFamilyName 2>&1 | Write-Log -Category Error -PassThru
        }

        if ($Local:debugPrivilegeEnabled) {
            Disable-DebugPrivilege 2>&1 | Write-Log -Category Error
        }

        Write-Progress -Completed
        Reset-ThreadCulture 2>&1 | Write-Log -Category Error
        Close-TaskRunspace 2>&1 | Write-Log -Category Error
        Close-Log

        if ($Local:ctrlCDisabled) {
            $null = Enable-CtrlC
        }

        $ScriptInfo.End = [DateTimeOffset]::Now
        $ScriptInfo | Export-CliXml (Join-Path $tempPath 'ScriptInfo.xml')
    }

    # Bail if something failed or user interruped with Ctrl+C.
    if (-not $Local:startSuccess) {
        Write-Warning "Temporary folder is `"$tempPath`""
        return
    }

    $archiveName = "Outlook_$($env:COMPUTERNAME)_$(Get-DateTimeString)"

    if ($SkipArchive) {
        # Rename with a job because it might take a while if Windows Search's SearchProtocolHost.exe opens the folder.
        $null = Start-Job { param($LiteralPath, $NewName) Rename-Item @PSBoundParameters } -ArgumentList $tempPath, $archiveName
        return
    }

    $archive = Compress-Folder -Path $tempPath -Destination $Path -ArchiveType $ArchiveType -ErrorAction Stop
    Rename-Item $archive.ArchivePath -NewName "$archiveName$([IO.Path]::GetExtension($archive.ArchivePath))"

    if (Test-Path $tempPath) {
        $null = Start-Job { param($LiteralPath) Remove-Item @PSBoundParameters -Recurse -Force } -ArgumentList $tempPath
    }

    Write-Host "The collected data is `"$(Join-Path $Path "$archiveName$([IO.Path]::GetExtension($archive.ArchivePath))")`"" -ForegroundColor Green
    Invoke-Item $Path
}

# Configure Export-Clixml, Out-File, Set-Content to use UTF8 by default.
if ($PSDefaultParameterValues -ne $null -and -not $PSDefaultParameterValues.Contains("Export-CliXml:Encoding")) {
    $PSDefaultParameterValues.Add("Export-Clixml:Encoding", 'UTF8')
}

if ($PSDefaultParameterValues -ne $null -and -not $PSDefaultParameterValues.Contains("Out-File:Encoding")) {
    $PSDefaultParameterValues.Add("Out-File:Encoding", 'utf8')
}

if ($PSDefaultParameterValues -ne $null -and -not $PSDefaultParameterValues.Contains("Set-Content:Encoding")) {
    $PSDefaultParameterValues.Add("Set-Content:Encoding", 'utf8')
}

# Some emoji chars (https://unicode.org/emoji/charts/full-emoji-list.html)
$Script:Emoji = @{
    Information = [Char]::ConvertFromUtf32(0x2139)
    Warning     = [Char]::ConvertFromUtf32(0x26A0)
    Error       = [Char]::ConvertFromUtf32(0x26D4) # This is actually "NoEntry" emoji
}

# Add type for Win32 interop
if (-not ('Win32.Kernel32' -as [type])) {
    Add-Type -TypeDefinition $Win32Interop
}

# Save this module path ("...\OutlookTrace.psm1") so that functions can easily find it when running in other runspaces.
$Script:MyModulePath = $PSCommandPath

$Script:ValidTimeSpan = [TimeSpan]::FromDays(90)

$Script:WamInteropBytes = @'
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BAUGBwABAgMGBwYHAQIDBgcFBgcAAgMGBwUGBwIDBgcEBQYHAAEDBgcFBgcBAwYHBAUGBwADBgcEBQYHAwYHAwQFBgcAAQIGBwUGBwECBgcEBQYHAAIGBwQFBgcCBgcDBAUGBwABBgcEBQYHAQYHAwQFBgcABgcDBAUGBwYHAgMEBQYHAAECAwQFBwcBAgMEBQcGBwACAwQFBwYHAgMEBQcFBgcAAQMEBQcGBwEDBAUHBQYHAAMEBQcFBgcDBAUHBAUGBwABAgQFBwYHAQIEBQcFBgcAAgQFBwUGBwIEBQcEBQYHAAEEBQcFBgcBBAUHBAUGBwAEBQcEBQYHBAUHAwQFBgcAAQIDBQcGBwECAwUHBQYHAAIDBQcFBgcCAwUHBAUGBwABAwUHBQYHAQMFBwQFBgcAAwUHBAUGBwMFBwMEBQYHAAECBQcFBgcBAgUHBAUGBwACBQcEBQYHAgUHAwQFBgcAAQUHBAUGBwEFBwMEBQYHAAUHAwQFBgcFBwIDBAUGBwABAgMEBwYHAQIDBAcFBgcAAgMEBwUGBwIDBAcEBQYHAAEDBAcFBgcBAwQHBAUGBwADBAcEBQYHAwQHAwQFBgcAAQIEBwUGBwECBAcEBQYHAAIEBwQFBgcCBAcDBAUGBwABBAcEBQYHAQQHAwQFBgcABAcDBAUGBwQHAgMEBQYHAAECAwcFBgcBAgMHBAUGBwACAwcEBQYHAgMHAwQFBgcAAQMHBAUGBwEDBwMEBQYHAAMHAwQFBgcDBwIDBAUGBwABAgcEBQYHAQIHAwQFBgcAAgcDBAUGBwIHAgMEBQYHAAEHAwQFBgcBBwIDBAUGBwAHAgMEBQYHBwECAwQFBgcAAQIDBAUGBwECAwQFBgYHAAIDBAUGBgcCAwQFBgUGBwABAwQFBgYHAQMEBQYFBgcAAwQFBgUGBwMEBQYEBQYHAAECBAUGBgcBAgQFBgUGBwACBAUGBQYHAgQFBgQFBgcAAQQFBgUGBwEEBQYEBQYHAAQFBgQFBgcEBQYDBAUGBwABAgMFBgYHAQIDBQYFBgcAAgMFBgUGBwIDBQYEBQYHAAEDBQYFBgcBAwUGBAUGBwADBQYEBQYHAwUGAwQFBgcAAQIFBgUGBwECBQYEBQYHAAIFBgQFBgcCBQYDBAUGBwABBQYEBQYHAQUGAwQFBgcABQYDBAUGBwUGAgMEBQYHAAECAwQGBgcBAgMEBgUGBwACAwQGBQYHAgMEBgQFBgcAAQMEBgUGBwEDBAYEBQYHAAMEBgQFBgcDBAYDBAUGBwABAgQGBQYHAQIEBgQFBgcAAgQGBAUGBwIEBgMEBQYHAAEEBgQFBgcBBAYDBAUGBwAEBgMEBQYHBAYCAwQFBgcAAQIDBgUGBwECAwYEBQYHAAIDBgQFBgcCAwYDBAUGBwABAwYEBQYHAQMGAwQFBgcAAwYDBAUGBwMGAgMEBQYHAAECBgQFBgcBAgYDBAUGBwACBgMEBQYHAgYCAwQFBgcAAQYDBAUGBwEGAgMEBQYHAAYCAwQFBgcGAQIDBAUGBwABAgMEBQYHAQIDBAUFBgcAAgMEBQUGBwIDBAUEBQYHAAEDBAUFBgcBAwQFBAUGBwADBAUEBQYHAwQFAwQFBgcAAQIEBQUGBwECBAUEBQYHAAIEBQQFBgcCBAUDBAUGBwABBAUEBQYHAQQFAwQFBgcABAUDBAUGBwQFAgMEBQYHAAECAwUFBgcBAgMFBAUGBwACAwUEBQYHAgMFAwQFBgcAAQMFBAUGBwEDBQMEBQYHAAMFAwQFBgcDBQIDBAUGBwABAgUEBQYHAQIFAwQFBgcAAgUDBAUGBwIFAgMEBQYHAAEFAwQFBgcBBQIDBAUGBwAFAgMEBQYHBQECAwQFBgcAAQIDBAUGBwECAwQEBQYHAAIDBAQFBgcCAwQDBAUGBwABAwQEBQYHAQMEAwQFBgcAAwQDBAUGBwMEAgMEBQYHAAECBAQFBgcBAgQDBAUGBwACBAMEBQYHAgQCAwQFBgcAAQQDBAUGBwEEAgMEBQYHAAQCAwQFBgcEAQIDBAUGBwABAgMEBQYHAQIDAwQFBgcAAgMDBAUGBwIDAgMEBQYHAAEDAwQFBgcBAwIDBAUGBwADAgMEBQYHAwECAwQFBgcAAQIDBAUGBwECAgMEBQYHAAICAwQFBgcCAQIDBAUGBwABAgMEBQYHAQECAwQFBgcAAQIDBAUGBwABAgMEBQYHCAcHBgcGBgUHBgYFBgUFBAcGBgUGBQUEBgUFBAUEBAMHBgYFBgUFBAYFBQQFBAQDBgUFBAUEBAMFBAQDBAMDAgcGBgUGBQUEBgUFBAUEBAMGBQUEBQQEAwUEBAMEAwMCBgUFBAUEBAMFBAQDBAMDAgUEBAMEAwMCBAMDAgMCAgEHBgYFBgUFBAYFBQQFBAQDBgUFBAUEBAMFBAQDBAMDAgYFBQQFBAQDBQQEAwQDAwIFBAQDBAMDAgQDAwIDAgIBBgUFBAUEBAMFBAQDBAMDAgUEBAMEAwMCBAMDAgMCAgEFBAQDBAMDAgQDAwIDAgIBBAMDAgMCAgEDAgIBAgEBAAABAgMEBQYHCAkKCwwNDg8CAwQFBgcICQoLDA0ODw4PAAEEBQYHCAkKCwwNDg8ODwQFBgcICQoLDA0ODwwNDg8AAQIDBgcICQoLDA0ODw4PAgMGBwgJCgsMDQ4PDA0ODwABBgcICQoLDA0ODwwNDg8GBwgJCgsMDQ4PCgsMDQ4PAAECAwQFCAkKCwwNDg8ODwIDBAUICQoLDA0ODwwNDg8AAQQFCAkKCwwNDg8MDQ4PBAUICQoLDA0ODwoLDA0ODwABAgMICQoLDA0ODwwNDg8CAwgJCgsMDQ4PCgsMDQ4PAAEICQoLDA0ODwoLDA0ODwgJCgsMDQ4PCAkKCwwNDg8AAQIDBAUGBwoLDA0ODw4PAgMEBQYHCgsMDQ4PDA0ODwABBAUGBwoLDA0ODwwNDg8EBQYHCgsMDQ4PCgsMDQ4PAAECAwYHCgsMDQ4PDA0ODwIDBgcKCwwNDg8KCwwNDg8AAQYHCgsMDQ4PCgsMDQ4PBgcKCwwNDg8ICQoLDA0ODwABAgMEBQoLDA0ODwwNDg8CAwQFCgsMDQ4PCgsMDQ4PAAEEBQoLDA0ODwoLDA0ODwQFCgsMDQ4PCAkKCwwNDg8AAQIDCgsMDQ4PCgsMDQ4PAgMKCwwNDg8ICQoLDA0ODwABCgsMDQ4PCAkKCwwNDg8KCwwNDg8GBwgJCgsMDQ4PAAECAwQFBgcICQwNDg8ODwIDBAUGBwgJDA0ODwwNDg8AAQQFBgcICQwNDg8MDQ4PBAUGBwgJDA0ODwoLDA0ODwABAgMGBwgJDA0ODwwNDg8CAwYHCAkMDQ4PCgsMDQ4PAAEGBwgJDA0ODwoLDA0ODwYHCAkMDQ4PCAkKCwwNDg8AAQIDBAUICQwNDg8MDQ4PAgMEBQgJDA0ODwoLDA0ODwABBAUICQwNDg8KCwwNDg8EBQgJDA0ODwgJCgsMDQ4PAAECAwgJDA0ODwoLDA0ODwIDCAkMDQ4PCAkKCwwNDg8AAQgJDA0ODwgJCgsMDQ4PCAkMDQ4PBgcICQoLDA0ODwABAgMEBQYHDA0ODwwNDg8CAwQFBgcMDQ4PCgsMDQ4PAAEEBQYHDA0ODwoLDA0ODwQFBgcMDQ4PCAkKCwwNDg8AAQIDBgcMDQ4PCgsMDQ4PAgMGBwwNDg8ICQoLDA0ODwABBgcMDQ4PCAkKCwwNDg8GBwwNDg8GBwgJCgsMDQ4PAAECAwQFDA0ODwoLDA0ODwIDBAUMDQ4PCAkKCwwNDg8AAQQFDA0ODwgJCgsMDQ4PBAUMDQ4PBgcICQoLDA0ODwABAgMMDQ4PCAkKCwwNDg8CAwwNDg8GBwgJCgsMDQ4PAAEMDQ4PBgcICQoLDA0ODwwNDg8EBQYHCAkKCwwNDg8AAQIDBAUGBwgJCgsODw4PAgMEBQYHCAkKCw4PDA0ODwABBAUGBwgJCgsODwwNDg8EBQYHCAkKCw4PCgsMDQ4PAAECAwYHCAkKCw4PDA0ODwIDBgcICQoLDg8KCwwNDg8AAQYHCAkKCw4PCgsMDQ4PBgcICQoLDg8ICQoLDA0ODwABAgMEBQgJCgsODwwNDg8CAwQFCAkKCw4PCgsMDQ4PAAEEBQgJCgsODwoLDA0ODwQFCAkKCw4PCAkKCwwNDg8AAQIDCAkKCw4PCgsMDQ4PAgMICQoLDg8ICQoLDA0ODwABCAkKCw4PCAkKCwwNDg8ICQoLDg8GBwgJCgsMDQ4PAAECAwQFBgcKCw4PDA0ODwIDBAUGBwoLDg8KCwwNDg8AAQQFBgcKCw4PCgsMDQ4PBAUGBwoLDg8ICQoLDA0ODwABAgMGBwoLDg8KCwwNDg8CAwYHCgsODwgJCgsMDQ4PAAEGBwoLDg8ICQoLDA0ODwYHCgsODwYHCAkKCwwNDg8AAQIDBAUKCw4PCgsMDQ4PAgMEBQoLDg8ICQoLDA0ODwABBAUKCw4PCAkKCwwNDg8EBQoLDg8GBwgJCgsMDQ4PAAECAwoLDg8ICQoLDA0ODwIDCgsODwYHCAkKCwwNDg8AAQoLDg8GBwgJCgsMDQ4PCgsODwQFBgcICQoLDA0ODwABAgMEBQYHCAkODwwNDg8CAwQFBgcICQ4PCgsMDQ4PAAEEBQYHCAkODwoLDA0ODwQFBgcICQ4PCAkKCwwNDg8AAQIDBgcICQ4PCgsMDQ4PAgMGBwgJDg8ICQoLDA0ODwABBgcICQ4PCAkKCwwNDg8GBwgJDg8GBwgJCgsMDQ4PAAECAwQFCAkODwoLDA0ODwIDBAUICQ4PCAkKCwwNDg8AAQQFCAkODwgJCgsMDQ4PBAUICQ4PBgcICQoLDA0ODwABAgMICQ4PCAkKCwwNDg8CAwgJDg8GBwgJCgsMDQ4PAAEICQ4PBgcICQoLDA0ODwgJDg8EBQYHCAkKCwwNDg8AAQIDBAUGBw4PCgsMDQ4PAgMEBQYHDg8ICQoLDA0ODwABBAUGBw4PCAkKCwwNDg8EBQYHDg8GBwgJCgsMDQ4PAAECAwYHDg8ICQoLDA0ODwIDBgcODwYHCAkKCwwNDg8AAQYHDg8GBwgJCgsMDQ4PBgcODwQFBgcICQoLDA0ODwABAgMEBQ4PCAkKCwwNDg8CAwQFDg8GBwgJCgsMDQ4PAAEEBQ4PBgcICQoLDA0ODwQFDg8EBQYHCAkKCwwNDg8AAQIDDg8GBwgJCgsMDQ4PAgMODwQFBgcICQoLDA0ODwABDg8EBQYHCAkKCwwNDg8ODwIDBAUGBwgJCgsMDQ4PAAECAwQFBgcICQoLDA0ODwIDBAUGBwgJCgsMDQwNDg8AAQQFBgcICQoLDA0MDQ4PBAUGBwgJCgsMDQoLDA0ODwABAgMGBwgJCgsMDQwNDg8CAwYHCAkKCwwNCgsMDQ4PAAEGBwgJCgsMDQoLDA0ODwYHCAkKCwwNCAkKCwwNDg8AAQIDBAUICQoLDA0MDQ4PAgMEBQgJCgsMDQoLDA0ODwABBAUICQoLDA0KCwwNDg8EBQgJCgsMDQgJCgsMDQ4PAAECAwgJCgsMDQoLDA0ODwIDCAkKCwwNCAkKCwwNDg8AAQgJCgsMDQgJCgsMDQ4PCAkKCwwNBgcICQoLDA0ODwABAgMEBQYHCgsMDQwNDg8CAwQFBgcKCwwNCgsMDQ4PAAEEBQYHCgsMDQoLDA0ODwQFBgcKCwwNCAkKCwwNDg8AAQIDBgcKCwwNCgsMDQ4PAgMGBwoLDA0ICQoLDA0ODwABBgcKCwwNCAkKCwwNDg8GBwoLDA0GBwgJCgsMDQ4PAAECAwQFCgsMDQoLDA0ODwIDBAUKCwwNCAkKCwwNDg8AAQQFCgsMDQgJCgsMDQ4PBAUKCwwNBgcICQoLDA0ODwABAgMKCwwNCAkKCwwNDg8CAwoLDA0GBwgJCgsMDQ4PAAEKCwwNBgcICQoLDA0ODwoLDA0EBQYHCAkKCwwNDg8AAQIDBAUGBwgJDA0MDQ4PAgMEBQYHCAkMDQoLDA0ODwABBAUGBwgJDA0KCwwNDg8EBQYHCAkMDQgJCgsMDQ4PAAECAwYHCAkMDQoLDA0ODwIDBgcICQwNCAkKCwwNDg8AAQYHCAkMDQgJCgsMDQ4PBgcICQwNBgcICQoLDA0ODwABAgMEBQgJDA0KCwwNDg8CAwQFCAkMDQgJCgsMDQ4PAAEEBQgJDA0ICQoLDA0ODwQFCAkMDQYHCAkKCwwNDg8AAQIDCAkMDQgJCgsMDQ4PAgMICQwNBgcICQoLDA0ODwABCAkMDQYHCAkKCwwNDg8ICQwNBAUGBwgJCgsMDQ4PAAECAwQFBgcMDQoLDA0ODwIDBAUGBwwNCAkKCwwNDg8AAQQFBgcMDQgJCgsMDQ4PBAUGBwwNBgcICQoLDA0ODwABAgMGBwwNCAkKCwwNDg8CAwYHDA0GBwgJCgsMDQ4PAAEGBwwNBgcICQoLDA0ODwYHDA0EBQYHCAkKCwwNDg8AAQIDBAUMDQgJCgsMDQ4PAgMEBQwNBgcICQoLDA0ODwABBAUMDQYHCAkKCwwNDg8EBQwNBAUGBwgJCgsMDQ4PAAECAwwNBgcICQoLDA0ODwIDDA0EBQYHCAkKCwwNDg8AAQwNBAUGBwgJCgsMDQ4PDA0CAwQFBgcICQoLDA0ODwABAgMEBQYHCAkKCwwNDg8CAwQFBgcICQoLCgsMDQ4PAAEEBQYHCAkKCwoLDA0ODwQFBgcICQoLCAkKCwwNDg8AAQIDBgcICQoLCgsMDQ4PAgMGBwgJCgsICQoLDA0ODwABBgcICQoLCAkKCwwNDg8GBwgJCgsGBwgJCgsMDQ4PAAECAwQFCAkKCwoLDA0ODwIDBAUICQoLCAkKCwwNDg8AAQQFCAkKCwgJCgsMDQ4PBAUICQoLBgcICQoLDA0ODwABAgMICQoLCAkKCwwNDg8CAwgJCgsGBwgJCgsMDQ4PAAEICQoLBgcICQoLDA0ODwgJCgsEBQYHCAkKCwwNDg8AAQIDBAUGBwoLCgsMDQ4PAgMEBQYHCgsICQoLDA0ODwABBAUGBwoLCAkKCwwNDg8EBQYHCgsGBwgJCgsMDQ4PAAECAwYHCgsICQoLDA0ODwIDBgcKCwYHCAkKCwwNDg8AAQYHCgsGBwgJCgsMDQ4PBgcKCwQFBgcICQoLDA0ODwABAgMEBQoLCAkKCwwNDg8CAwQFCgsGBwgJCgsMDQ4PAAEEBQoLBgcICQoLDA0ODwQFCgsEBQYHCAkKCwwNDg8AAQIDCgsGBwgJCgsMDQ4PAgMKCwQFBgcICQoLDA0ODwABCgsEBQYHCAkKCwwNDg8KCwIDBAUGBwgJCgsMDQ4PAAECAwQFBgcICQoLDA0ODwIDBAUGBwgJCAkKCwwNDg8AAQQFBgcICQgJCgsMDQ4PBAUGBwgJBgcICQoLDA0ODwABAgMGBwgJCAkKCwwNDg8CAwYHCAkGBwgJCgsMDQ4PAAEGBwgJBgcICQoLDA0ODwYHCAkEBQYHCAkKCwwNDg8AAQIDBAUICQgJCgsMDQ4PAgMEBQgJBgcICQoLDA0ODwABBAUICQYHCAkKCwwNDg8EBQgJBAUGBwgJCgsMDQ4PAAECAwgJBgcICQoLDA0ODwIDCAkEBQYHCAkKCwwNDg8AAQgJBAUGBwgJCgsMDQ4PCAkCAwQFBgcICQoLDA0ODwABAgMEBQYHCAkKCwwNDg8CAwQFBgcGBwgJCgsMDQ4PAAEEBQYHBgcICQoLDA0ODwQFBgcEBQYHCAkKCwwNDg8AAQIDBgcGBwgJCgsMDQ4PAgMGBwQFBgcICQoLDA0ODwABBgcEBQYHCAkKCwwNDg8GBwIDBAUGBwgJCgsMDQ4PAAECAwQFBgcICQoLDA0ODwIDBAUEBQYHCAkKCwwNDg8AAQQFBAUGBwgJCgsMDQ4PBAUCAwQFBgcICQoLDA0ODwABAgMEBQYHCAkKCwwNDg8CAwIDBAUGBwgJCgsMDQ4PAAECAwQFBgcICQoLDA0ODwABAgMEBQYHCAkKCwwNDg8QDg4MDgwMCg4MDAoMCgoIDgwMCgwKCggMCgoICggIBg4MDAoMCgoIDAoKCAoICAYMCgoICggIBgoICAYIBgYEDgwMCgwKCggMCgoICggIBgwKCggKCAgGCggIBggGBgQMCgoICggIBgoICAYIBgYECggIBggGBgQIBgYEBgQEAg4MDAoMCgoIDAoKCAoICAYMCgoICggIBgoICAYIBgYEDAoKCAoICAYKCAgGCAYGBAoICAYIBgYECAYGBAYEBAIMCgoICggIBgoICAYIBgYECggIBggGBgQIBgYEBgQEAgoICAYIBgYECAYGBAYEBAIIBgYEBgQEAgYEBAIEAgIAAAECAwQFBgcICQoLDA0ODwQFBgcICQoLDA0ODwwNDg8AAQIDCAkKCwwNDg8MDQ4PCAkKCwwNDg8ICQoLDA0ODwABAgMEBQYHDA0ODwwNDg8EBQYHDA0ODwgJCgsMDQ4PAAECAwwNDg8ICQoLDA0ODwwNDg8EBQYHCAkKCwwNDg8AAQIDBAUGBwgJCgsMDQ4PBAUGBwgJCgsICQoLDA0ODwABAgMICQoLCAkKCwwNDg8ICQoLBAUGBwgJCgsMDQ4PAAECAwQFBgcICQoLDA0ODwQFBgcEBQYHCAkKCwwNDg8AAQIDBAUGBwgJCgsMDQ4PAAECAwQFBgcICQoLDA0ODxAMDAgMCAgEDAgIBAgEBAAAAQIDBAUGBwECAwQFBgcHAAIDBAUGBwcCAwQFBgcGBwABAwQFBgcHAQMEBQYHBgcAAwQFBgcGBwMEBQYHBQYHAAECBAUGBwcBAgQFBgcGBwACBAUGBwYHAgQFBgcFBgcAAQQFBgcGBwEEBQYHBQYHAAQFBgcFBgcEBQYHBAUGBwABAgMFBgcHAQIDBQYHBgcAAgMFBgcGBwIDBQYHBQYHAAEDBQYHBgcBAwUGBwUGBwADBQYHBQYHAwUGBwQFBgcAAQIFBgcGBwECBQYHBQYHAAIFBgcFBgcCBQYHBAUGBwABBQYHBQYHAQUGBwQFBgcABQYHBAUGBwUGBwMEBQYHAAECAwQGBwcBAgMEBgcGBwACAwQGBwYHAgMEBgcFBgcAAQMEBgcGBwEDBAYHBQYHAAMEBgcFBgcDBAYHBAUGBwABAgQGBwYHAQIEBgcFBgcAAgQGBwUGBwIEBgcEBQYHAAEEBgcFBgcBBAYHBAUGBwAEBgcEBQYHBAYHAwQFBgcAAQIDBgcGBwECAwYHBQYHAAIDBgcFBgcCAwYHBAUGBwABAwYHBQYHAQMGBwQFBgcAAwYHBAUGBwMGBwMEBQYHAAECBgcFBgcBAgYHBAUGBwACBgcEBQYHAgYHAwQFBgcAAQYHBAUGBwEGBwMEBQYHAAYHAwQFBgcGBwIDBAUGBwABAgMEBQcHAQIDBAUHBgcAAgMEBQcGBwIDBAUHBQYHAAEDBAUHBgcBAwQFBwUGBwADBAUHBQYHAwQFBwQFBgcAAQIEBQcGBwECBAUHBQYHAAIEBQcFBgcCBAUHBAUGBwABBAUHBQYHAQQFBwQFBgcABAUHBAUGBwQFBwMEBQYHAAECAwUHBgcBAgMFBwUGBwACAwUHBQYHAgMFBwQFBgcAAQMFBwUGBwEDBQcEBQYHAAMFBwQFBgcDBQcDBAUGBwABAgUHBQYHAQIFBwQFBgcAAgUHBAUGBwIFBwMEBQYHAAEFBwQFBgcBBQcDBAUGBwAFBwMEBQYHBQcCAwQFBgcAAQIDBAcGBwECAwQHBQYHAAIDBAcFBgcCAwQHBAUGBwABAwQHBQYHAQMEBwQFBgcAAwQHBAUGBwMEBwMEBQYHAAECBAcFBgcBAgQHBAUGBwACBAcEBQYHAgQHAwQFBgcAAQQHBAUGBwEEBwMEBQYHAAQHAwQFBgcEBwIDBAUGBwABAgMHBQYHAQIDBwQFBgcAAgMHBAUGBwIDBwMEBQYHAAEDBwQFBgcBAwcDBAUGBwADBwMEBQYHAwcCAwQFBgcAAQIHBAUGBwECBwMEBQYHAAIHAwQFBgcCBwIDBAUGBwABBwMEBQYHAQcCAwQFBgcABwIDBAUGBwcBAgMEBQYHAAECAwQFBgcBAgMEBQYGBwACAwQFBgYHAgMEBQYFBgcAAQMEBQYGBwEDBAUGBQYHAAMEBQYFBgcDBAUGBAUGBwABAgQFBgYHAQIEBQYFBgcAAgQFBgUGBwIEBQYEBQYHAAEEBQYFBgcBBAUGBAUGBwAEBQYEBQYHBAUGAwQFBgcAAQIDBQYGBwECAwUGBQYHAAIDBQYFBgcCAwUGBAUGBwABAwUGBQYHAQMFBgQFBgcAAwUGBAUGBwMFBgMEBQYHAAECBQYFBgcBAgUGBAUGBwACBQYEBQYHAgUGAwQFBgcAAQUGBAUGBwEFBgMEBQYHAAUGAwQFBgcFBgIDBAUGBwABAgMEBgYHAQIDBAYFBgcAAgMEBgUGBwIDBAYEBQYHAAEDBAYFBgcBAwQGBAUGBwADBAYEBQYHAwQGAwQFBgcAAQIEBgUGBwECBAYEBQYHAAIEBgQFBgcCBAYDBAUGBwABBAYEBQYHAQQGAwQFBgcABAYDBAUGBwQGAgMEBQYHAAECAwYFBgcBAgMGBAUGBwACAwYEBQYHAgMGAwQFBgcAAQMGBAUGBwEDBgMEBQYHAAMGAwQFBgcDBgIDBAUGBwABAgYEBQYHAQIGAwQFBgcAAgYDBAUGBwIGAgMEBQYHAAEGAwQFBgcBBgIDBAUGBwAGAgMEBQYHBgECAwQFBgcAAQIDBAUGBwECAwQFBQYHAAIDBAUFBgcCAwQFBAUGBwABAwQFBQYHAQMEBQQFBgcAAwQFBAUGBwMEBQMEBQYHAAECBAUFBgcBAgQFBAUGBwACBAUEBQYHAgQFAwQFBgcAAQQFBAUGBwEEBQMEBQYHAAQFAwQFBgcEBQIDBAUGBwABAgMFBQYHAQIDBQQFBgcAAgMFBAUGBwIDBQMEBQYHAAEDBQQFBgcBAwUDBAUGBwADBQMEBQYHAwUCAwQFBgcAAQIFBAUGBwECBQMEBQYHAAIFAwQFBgcCBQIDBAUGBwABBQMEBQYHAQUCAwQFBgcABQIDBAUGBwUBAgMEBQYHAAECAwQFBgcBAgMEBAUGBwACAwQEBQYHAgMEAwQFBgcAAQMEBAUGBwEDBAMEBQYHAAMEAwQFBgcDBAIDBAUGBwABAgQEBQYHAQIEAwQFBgcAAgQDBAUGBwIEAgMEBQYHAAEEAwQFBgcBBAIDBAUGBwAEAgMEBQYHBAECAwQFBgcAAQIDBAUGBwECAwMEBQYHAAIDAwQFBgcCAwIDBAUGBwABAwMEBQYHAQMCAwQFBgcAAwIDBAUGBwMBAgMEBQYHAAECAwQFBgcBAgIDBAUGBwACAgMEBQYHAgECAwQFBgcAAQIDBAUGBwEBAgMEBQYHAAECAwQFBgcAAQIDBAUGByAcHBgcGBgUHBgYFBgUFBAcGBgUGBQUEBgUFBAUEBAMHBgYFBgUFBAYFBQQFBAQDBgUFBAUEBAMFBAQDBAMDAgcGBgUGBQUEBgUFBAUEBAMGBQUEBQQEAwUEBAMEAwMCBgUFBAUEBAMFBAQDBAMDAgUEBAMEAwMCBAMDAgMCAgEHBgYFBgUFBAYFBQQFBAQDBgUFBAUEBAMFBAQDBAMDAgYFBQQFBAQDBQQEAwQDAwIFBAQDBAMDAgQDAwIDAgIBBgUFBAUEBAMFBAQDBAMDAgUEBAMEAwMCBAMDAgMCAgEFBAQDBAMDAgQDAwIDAgIBBAMDAgMCAgEDAgIBAgEBAAAAQIDBAUGBwgJCgsMDQ4PCAkKCwwNDg8ICQoLDA0ODwABAgMEBQYHCAkKCwwNDg8AAQIDBAUGBwgJCgsMDQ4PEAgIAAAAAAAAAAAAAAAAAAABAgMEBQYHAgMEBQYHBgcAAQQFBgcGBwQFBgcEBQYHAAECAwYHBgcCAwYHBAUGBwABBgcEBQYHBgcCAwQFBgcAAQID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W10AAAAAAABvcGVyYXRvcgAAAAAtPgAAKgAAACsrAAAtLQAALQAAACsAAAAmAAAALT4qAC8AAAAlAAAAPAAAADw9AAA+AAAAPj0AACwAAAAoKQAAfgAAAF4AAAB8AAAAJiYAAHx8AAAqPQAAKz0AAC09AAAvPQAAJT0AAD4+PQA8PD0AJj0AAHw9AABePQAAYHZmdGFibGUnAAAAAAAAAGB2YnRhYmxlJwAAAAAAAABgdmNhbGwnAGB0eXBlb2YnAAAAAAAAAABgbG9jYWwgc3RhdGljIGd1YXJkJwAAAABgc3RyaW5nJwAAAAAAAAAAYHZiYXNlIGRlc3RydWN0b3InAAAAAAAAYHZlY3RvciBkZWxldGluZyBkZXN0cnVjdG9yJwAAAABgZGVmYXVsdCBjb25zdHJ1Y3RvciBjbG9zdXJlJwAAAGBzY2FsYXIgZGVsZXRpbmcgZGVzdHJ1Y3RvcicAAAAAYHZlY3RvciBjb25zdHJ1Y3RvciBpdGVyYXRvcicAAABgdmVjdG9yIGRlc3RydWN0b3IgaXRlcmF0b3InAAAAAGB2ZWN0b3IgdmJhc2UgY29uc3RydWN0b3IgaXRlcmF0b3InAAAAAABgdmlydHVhbCBkaXNwbGFjZW1lbnQgbWFwJwAAAAAAAGBlaCB2ZWN0b3IgY29uc3RydWN0b3IgaXRlcmF0b3InAAAAAAAAAABgZWggdmVjdG9yIGRlc3RydWN0b3IgaXRlcmF0b3InAGBlaCB2ZWN0b3IgdmJhc2UgY29uc3RydWN0b3IgaXRlcmF0b3InAABgY29weSBjb25zdHJ1Y3RvciBjbG9zdXJlJwAAAAAAAGB1ZHQgcmV0dXJuaW5nJwBgRUgAYFJUVEkAAAAAAAAAYGxvY2FsIHZmdGFibGUnAGBsb2NhbCB2ZnRhYmxlIGNvbnN0cnVjdG9yIGNsb3N1cmUnACBuZXdbXQAAAAAAACBkZWxldGVbXQAAAAAAAABgb21uaSBjYWxsc2lnJwAAYHBsYWNlbWVudCBkZWxldGUgY2xvc3VyZScAAAAAAABgcGxhY2VtZW50IGRlbGV0ZVtdIGNsb3N1cmUnAAAAAGBtYW5hZ2VkIHZlY3RvciBjb25zdHJ1Y3RvciBpdGVyYXRvcicAAABgbWFuYWdlZCB2ZWN0b3IgZGVzdHJ1Y3RvciBp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AAAAAAAAAAABAAAA6OkBAAAAAAAAAAAAAAAAAAAAAAACAAAA+OkBAAAAAAAAAAAAAAAAAAEAAAAGAAAAgOkBAAAAAAAAAAAAAQAAAAAAAAAAAAAAkDICAEjrAQBg6wEAAAAAAAAAAAAAAAAAAAAAAKAqAgACAAAAAAAAAP////8AAAAAQAAAANDrAQAAAAAAAAAAAAAAAACI6wEAsPABADDxAQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAwAAALDrAQAAAAAAAAAAAAEAAAAAAAAAAAAAAKAqAgDQ6wEA6OsBAAAAAAAAAAAAAAAAAAAAAABQNAIAAgAAAAAAAAD/////AAAAAEAAAACw7AEAAAAAAAAAAAAAAAAAaOwBAEDuAQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAACAAAAOOwBAAAAAAAAAAAA2DQCAAEAAAAAAAAA/////wAAAABAAAAAUOwBAAAAAAAAAAAAAAAAABDsAQBo7AEAQO4BAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAADAAAAkOwBAAAAAAAAAAAAAQAAAAAAAAAAAAAAUDQCALDsAQDI7AEAAAAAAAAAAAAAAAAAAAAAAAEAAAAAAAAAAAAAAAg1AgBo7wEA8OwBAAAAAAAAAAAAAAAAAAAAAAAINQIABQAAAAAAAAD/////AAAAAEAAAABo7wEAAAAAAAAAAAAAAAAAcO0BAEDuAQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAACAAAAQO0BAAAAAAAAAAAAQDUCAAEAAAAAAAAA/////wAAAABAAAAAWO0BAAAAAAAAAAAAAAAAALA1AgAAAAAAAAAAAP////8AAAAAQgAAAHjuAQAAAAAAAAAAAAAAAABwNQIAAQAAAAAAAAD/////AAAAAEAAAAAA7gEAAAAAAAAAAAAAAAAAwO0BAEDuAQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAACAAAA6O0BAAAAAAAAAAAAcDUCAAEAAAAIAAAA/////wAAAABAAAAAAO4BAAAAAAAAAAAAAAAAALA1AgAAAAAAAAAAAP////8AAAAAQAAAAHjuAQAAAAAAAAAAAAAAAABA7gEAAAAAAAAAAAAAAAAAAAAAAAAAAAABAAAAaO4BAAAAAAAAAAAAsDUCAAAAAAAIAAAA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AAAAAAAAAAAAAAAAAgAAAKAKAgCYEAIAAAAAAAAAAAAAAAAAAAAAABAlAAAAAAAAyAoCAAAAAAAAAAAAAAAAAAAAAAAAAAAACCsCAAAAAAD/////AAAAABgAAABQJQAAAAAAAAAAAAAAAAAAAgAAAAALAgCYEAIAAAAAAAAAAAAAAAAAAAAAABAlAAAAAAAAKAsCAAAAAAAAAAAAAAAAAAAAAAAAAAAAQCsCAAAAAAD/////AAAAABgAAABQJQAAAAAAAAAAAAAAAAAAAgAAAGALAgCYEAIAAAAAAAAAAAAAAAAAAAAAABAlAAAAAAAAiAsCAAAAAAAAAAAAAAAAAAAAAAAAAAAAeCsCAAAAAAD/////AAAAABgAAABQJQAAAAAAAAAAAAAAAAAAAgAAAMALAgCYEAIAAAAAAAAAAAAAAAAAAAAAABAlAAAAAAAA6AsCAAAAAAAAAAAAAAAAAAAAAAAAAAAAsCsCAAAAAAD/////AAAAABgAAABQJQAAAAAAAAAAAAAAAAAAAgAAACAMAgCYEAIAAAAAAAAAAAAAAAAAAAAAABAlAAAAAAAASAwCAAAAAAAAAAAAAAAAAAAAAAAAAAAA6CsCAAAAAAD/////AAAAABgAAABQJQAAAAAAAAAAAAAAAAAAAgAAAIAMAgCYEAIAAAAAAAAAAAAAAAAAAAAAABAlAAAAAAAAqAwCAAAAAAAAAAAAAAAAAAAAAAAAAAAAICwCAAAAAAD/////AAAAABgAAABQJQAAAAAAAAAAAAAAAAAAAgAAAOAMAgCYEAIAAAAAAAAAAAAAAAAAAAAAABAlAAAAAAAACA0CAAAAAAAAAAAAAAAAAAAAAAAAAAAAWCwCAAAAAAD/////AAAAABgAAABQJQAAAAAAAAAAAAAAAAAAAgAAAEANAgCYEAIAAAAAAAAAAAAAAAAAAAAAABAlAAAAAAAAaA0CAAAAAAAAAAAAAAAAAAAAAAAAAAAAmCwCAAAAAAD/////AAAAABgAAABQJQAAAAAAAAAAAAAAAAAAAgAAAKANAgCYEAIAAAAAAAAAAAAAAAAAAAAAABAlAAAAAAAAyA0CAAAAAAAAAAAAAAAAAAAAAAAAAAAA0CwCAAAAAAD/////AAAAABgAAABQJQAAAAAAAAAAAAAAAAAAAgAAAAAOAgCYEAIAAAAAAAAAAAAAAAAA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ZldpbmRvd1Byb2NXAABbAkxvYWRJY29uVwBZAkxvYWRDdXJzb3JXAMYBR2V0U3lzQ29sb3JCcnVzaAAA3wJSZWdpc3RlckNsYXNzRXhXAAAzAUdldENsaWVudFJlY3QARwFHZXREZXNrdG9wV2luZG93AAAeAUdldEFuY2VzdG9yAHYAQ3JlYXRlV2luZG93RXhXAFVTRVIzMi5kbGwAAPUEUnRsQ2FwdHVyZUNvbnRleHQA/QRSdGxMb29rdXBGdW5jdGlvbkVudHJ5AAAEBVJ0bFZpcnR1YWxVbndpbmQAAOYFVW5oYW5kbGVkRXhjZXB0aW9uRmlsdGVyAACkBVNldFVuaGFuZGxlZEV4Y2VwdGlvbkZpbHRlcgAyAkdldEN1cnJlbnRQcm9jZXNzAMQFVGVybWluYXRlUHJvY2VzcwAAqANJc1Byb2Nlc3NvckZlYXR1cmVQcmVzZW50AHAEUXVlcnlQZXJmb3JtYW5jZUNvdW50ZXIAMwJHZXRDdXJyZW50UHJvY2Vzc0lkADcCR2V0Q3VycmVudFRocmVhZElkAAAKA0dldFN5c3RlbVRpbWVBc0ZpbGVUaW1lAIoDSW5pdGlhbGl6ZVNMaXN0SGVhZACgA0lzRGVidWdnZXJQcmVzZW50APECR2V0U3RhcnR1cEluZm9XAAMFUnRsVW53aW5kRXgA/wRSdGxQY1RvRmlsZUhlYWRlcgCHBFJhaXNlRXhjZXB0aW9uAACQA0ludGVybG9ja2VkUHVzaEVudHJ5U0xpc3QAjgNJbnRlcmxvY2tlZEZsdXNoU0xpc3QAfQJHZXRMYXN0RXJyb3IAAGQFU2V0TGFzdEVycm9yAABFAUVuY29kZVBvaW50ZXIASQFFbnRlckNyaXRpY2FsU2VjdGlvbgAA4ANMZWF2ZUNyaXRpY2FsU2VjdGlvbgAAIwFEZWxldGVDcml0aWNhbFNlY3Rpb24AhgNJbml0aWFsaXplQ3JpdGljYWxTZWN0aW9uQW5kU3BpbkNvdW50ANYFVGxzQWxsb2MAANgFVGxzR2V0VmFsdWUA2QVUbHNTZXRWYWx1ZQDXBVRsc0ZyZWUAxQFGcmVlTGlicmFyeQDNAkdldFByb2NBZGRyZXNzAADmA0xvYWRMaWJyYXJ5RXhXAAB4AUV4aXRQcm9jZXNzAJQCR2V0TW9kdWxlSGFuZGxlRXhXAACRAkdldE1vZHVsZUZpbGVOYW1lVwAAcANIZWFwRnJlZQAAbANIZWFwQWxsb2MAjwFGaW5kQ2xvc2UAlQFGaW5kRmlyc3RGaWxlRXhXAACmAUZpbmROZXh0RmlsZVcArgNJc1ZhbGlkQ29kZVBhZ2UAzAFHZXRBQ1AAALYCR2V0T0VNQ1AAANsBR2V0Q1BJbmZvAPABR2V0Q29tbWFuZExpbmVBAPEBR2V0Q29tbWFuZExpbmVXABIETXVsdGlCeXRlVG9XaWRlQ2hhcgA3BldpZGVDaGFyVG9NdWx0aUJ5dGUAUwJHZXRFbnZpcm9ubWVudFN0cmluZ3NXAADEAUZyZWVFbnZpcm9ubWVudFN0cmluZ3NXALQBRmxzQWxsb2MAALYBRmxzR2V0VmFsdWUAtwFGbHNTZXRWYWx1ZQC1AUZsc0ZyZWUA1ANMQ01hcFN0cmluZ1cAANQCR2V0UHJvY2Vzc0hlYXAAAPMCR2V0U3RkSGFuZGxlAABqAkdldEZpbGVUeXBlAPgCR2V0U3RyaW5nVHlwZVcAAHUDSGVhcFNpemUAAHMDSGVhcFJlQWxsb2MAfwVTZXRTdGRIYW5kbGUAALkBRmx1c2hGaWxlQnVmZmVycwAASwZXcml0ZUZpbGUAGgJHZXRDb25zb2xlT3V0cHV0Q1AAABYCR2V0Q29uc29sZU1vZGUAAFUFU2V0RmlsZVBvaW50ZXJFeAAA2gBDcmVhdGVGaWxlVwCUAENsb3NlSGFuZGxlAEoGV3JpdGVDb25zb2xlVwDBAUZvcm1hdE1lc3NhZ2VXAADnA0xvYWRMaWJyYXJ5VwAASAVTZXRFdmVudAAAzgBDcmVhdGVFdmVudFcAABAGV2FpdEZvclNpbmdsZU9iamVjdAApAENvQ3JlYXRlRnJlZVRocmVhZGVkTWFyc2hhbGVyAG9sZTMyLmRsbABPTEVBVVQzMi5kbGwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAP//////////AQAAAAIAAAAAAAgAAAAAAAAAAAIAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAyot8tmSsAAAAAAAAAAAAA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AAAAAC4/QVVocmVzdWx0X2FjY2Vzc19kZW5pZWRAd2lucnRAQAAAAAAAAACoaQGAAQAAAAAAAAAAAAAALj9BVWhyZXN1bHRfd3JvbmdfdGhyZWFkQHdpbnJ0QEAAAAAAAAAAAKhpAYABAAAAAAAAAAAAAAAuP0FVaHJlc3VsdF9ub3RfaW1wbGVtZW50ZWRAd2lucnRAQAAAAAAAqGkBgAEAAAAAAAAAAAAAAC4/QVVocmVzdWx0X2ludmFsaWRfYXJndW1lbnRAd2lucnRAQAAAAACoaQGAAQAAAAAAAAAAAAAALj9BVWhyZXN1bHRfb3V0X29mX2JvdW5kc0B3aW5ydEBAAAAAAAAAAKhpAYABAAAAAAAAAAAAAAAuP0FVaHJlc3VsdF9ub19pbnRlcmZhY2VAd2lucnRAQAAAAAAAAAAAqGkBgAEAAAAAAAAAAAAAAC4/QVVocmVzdWx0X2NsYXNzX25vdF9hdmFpbGFibGVAd2lucnRAQACoaQGAAQAAAAAAAAAAAAAALj9BVWhyZXN1bHRfY2xhc3Nfbm90X3JlZ2lzdGVyZWRAd2lucnRAQAAAAAAAAAAAqGkBgAEAAAAAAAAAAAAAAC4/QVVocmVzdWx0X2NoYW5nZWRfc3RhdGVAd2lucnRAQAAAAAAAAACoaQGAAQAAAAAAAAAAAAAALj9BVWhyZXN1bHRfaWxsZWdhbF9tZXRob2RfY2FsbEB3aW5ydEBAAKhpAYABAAAAAAAAAAAAAAAuP0FVaHJlc3VsdF9pbGxlZ2FsX3N0YXRlX2NoYW5nZUB3aW5ydEBAAAAAAAAAAAAAAAAAAAAAAKhpAYABAAAAAAAAAAAAAAAuP0FVaHJlc3VsdF9pbGxlZ2FsX2RlbGVnYXRlX2Fzc2lnbm1lbnRAd2lucnRAQACoaQGAAQAAAAAAAAAAAAAALj9BVWhyZXN1bHRfY2FuY2VsZWRAd2lucnRAQAAAAACoaQGAAQAAAAAAAAAAAAAALj9BVmludmFsaWRfYXJndW1lbnRAc3RkQEAAAAAAAACoaQGAAQAAAAAAAAAAAAAALj9BVmxvZ2ljX2Vycm9yQHN0ZEBAAAAAqGkBgAEAAAAAAAAAAAAAAC4/QVZiYWRfYWxsb2NAc3RkQEAAAAAAAKhpAYABAAAAAAAAAAAAAAAuP0FWZXhjZXB0aW9uQHN0ZEBAAAAAAACoaQGAAQAAAAAAAAAAAAAALj9BVWhyZXN1bHRfZXJyb3JAd2lucnRAQAAAAAAAAAAAAAAAAAAAAKhpAYABAAAAAAAAAAAAAAAuP0FWdHlwZV9pbmZvQEAAqGkBgAEAAAAAAAAAAAAAAC4/QVU/JGltcGxlbWVudHNfZGVsZWdhdGVAVT8kQXN5bmNPcGVyYXRpb25Db21wbGV0ZWRIYW5kbGVyQFVXZWJUb2tlblJlcXVlc3RSZXN1bHRAQ29yZUBXZWJAQXV0aGVudGljYXRpb25AU2VjdXJpdHlAV2luZG93c0B3aW5ydEBAQEZvdW5kYXRpb25AV2luZG93c0B3aW5ydEBAVXNoYXJlZF90eXBlQD8xPz8/JHdhaXRfZm9yX2NvbXBsZXRlZEBVPyRJQXN5bmNPcGVyYXRpb25AVVdlYlRva2VuUmVxdWVzdFJlc3VsdEBDb3JlQFdlYkBBdXRoZW50aWNhdGlvbkBTZWN1cml0eUBXaW5kb3dzQHdpbnJ0QEBARm91bmRhdGlvbkBXaW5kb3dzQHdpbnJ0QEBAaW1wbEA0QFlBP0FfUEFFQlU/JElBc3luY09wZXJhdGlvbkBVV2ViVG9rZW5SZXF1ZXN0UmVzdWx0QENvcmVAV2ViQEF1dGhlbnRpY2F0aW9uQFNlY3VyaXR5QFdpbmRvd3NAd2lucnRAQEAyMzRASUBaQEBpbXBsQHdpbnJ0QEAAAAAAAACoaQGAAQAAAAAAAAAAAAAALj9BVXNoYXJlZF90eXBlQD8xPz8/JHdhaXRfZm9yX2NvbXBsZXRlZEBVPyRJQXN5bmNPcGVyYXRpb25AVVdlYlRva2VuUmVxdWVzdFJlc3VsdEBDb3JlQFdlYkBBdXRoZW50aWNhdGlvbkBTZWN1cml0eUBXaW5kb3dzQHdpbnJ0QEBARm91bmRhdGlvbkBXaW5kb3dzQHdpbnJ0QEBAaW1wbEB3aW5ydEBAWUE/QV9QQUVCVT8kSUFzeW5jT3BlcmF0aW9uQFVXZWJUb2tlblJlcXVlc3RSZXN1bHRAQ29yZUBXZWJAQXV0aGVudGljYXRpb25AU2VjdXJpdHlAV2luZG93c0B3aW5ydEBAQEZvdW5kYXRpb25AV2luZG93c0AyQElAWkAAAAAAAAAAAKhpAYABAAAAAAAAAAAAAAAuP0FVdHlwZUA/JGFiaUBVPyRBc3luY09wZXJhdGlvbkNvbXBsZXRlZEhhbmRsZXJAVVdlYlRva2VuUmVxdWVzdFJlc3VsdEBDb3JlQFdlYkBBdXRoZW50aWNhdGlvbkBTZWN1cml0eUBXaW5kb3dzQHdpbnJ0QEBARm91bmRhdGlvbkBXaW5kb3dzQHdpbnJ0QEBYQGltcGxAd2lucnRAQAAAAAAAAAAAAAAAAAAAAKhpAYABAAAAAAAAAAAAAAAuP0FVPyRkZWxlZ2F0ZUBVPyRBc3luY09wZXJhdGlvbkNvbXBsZXRlZEhhbmRsZXJAVVdlYlRva2VuUmVxdWVzdFJlc3VsdEBDb3JlQFdlYkBBdXRoZW50aWNhdGlvbkBTZWN1cml0eUBXaW5kb3dzQHdpbnJ0QEBARm91bmRhdGlvbkBXaW5kb3dzQHdpbnJ0QEBVc2hhcmVkX3R5cGVAPzE/Pz8kd2FpdF9mb3JfY29tcGxldGVkQFU/JElBc3luY09wZXJhdGlvbkBVV2ViVG9rZW5SZXF1ZXN0UmVzdWx0QENvcmVAV2ViQEF1dGhlbnRpY2F0aW9uQFNlY3VyaXR5QFdpbmRvd3NAd2lucnRAQEBGb3VuZGF0aW9uQFdpbmRvd3NAd2lucnRAQEBpbXBsQDRAWUE/QV9QQUVCVT8kSUFzeW5jT3BlcmF0aW9uQFVXZWJUb2tlblJlcXVlc3RSZXN1bHRAQ29yZUBXZWJAQXV0aGVudGljYXRpb25AU2VjdXJpdHlAV2luZG93c0B3aW5ydEBAQDIzNEBJQFpAQGltcGxAd2lucnRAQACoaQGAAQAAAAAAAAAAAAAALj9BVW1hcnNoYWxlckA/MT8/bWFrZV9tYXJzaGFsZXJAaW1wbEB3aW5ydEBAWUFIUEVBVXR5cGVAPyRhYmlAVUlVbmtub3duQEZvdW5kYXRpb25AV2luZG93c0B3aW5ydEBAWEAyM0BQRUFQRUFYQFpAAAAAAAAAqGkBgAEAAAAAAAAAAAAAAC4/QVVJTWFyc2hhbEBpbXBsQHdpbnJ0QEAAAAAAAAAAqGkBgAEAAAAAAAAAAAAAAC4/QVVlcnJvcl9pbmZvX2ZhbGxiYWNrQGltcGxAd2lucnRAQAAAAACoaQGAAQAAAAAAAAAAAAAALj9BVUlFcnJvckluZm9AaW1wbEB3aW5y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5P0BABiHAACuhwAAmPsBALCHAACdiAAAiP4BAKCIAAAoiQAAmPsBAOCJAACuigAANP0BALCKAABgiwAATP0BAJCLAACriwAAEP0BALCLAACxiwAAEP0BAMCLAADBiwAAEP0BAPyLAABCjAAAUPcBAESMAAB7jAAAUPcBAHyMAADLjQAAmAACAMyNAAARjgAAUPcBABSOAABajgAAUPcBAFyOAACijgAAUPcBAKSOAAD1jgAAdPgBAPiOAABZjwAAEPsBAKCPAADgjwAAiPYBAOCPAAAKkAAAiPYBABCQAAA2kAAAiPYBAECQAACHkAAAiPYBAIiQAADekAAAiPYBAOCQAABHkQAAkPcBAEiRAACykQAA9AACALSRAAAAkgAAEPsBAACSAABbkwAAwAACAGSTAAD/kwAABAECAACUAADMlAAA4AACAMyUAADqlAAAuAACAOyUAAAblQAAuAACAByVAABjlQAAiPYBAKyVAADzlQAAUPcBAPSVAAAWlgAAUPcBABiWAAA4lgAAiPYBADiWAABYlgAAiPYBAFiWAADzlgAAEPsBAByXAABalwAAdPgBAFyXAACTlwAAEAECAJyXAADdlwAAMAECAOCXAAAnmAAAdPgBADyYAAC9mQAAmPsBAMCZAAD4mQAAoAECAPiZAADDmgAAgAECAMSaAACJmwAAWAECAIybAAC+mwAAUPcBAMCbAADhmwAAiPYBAOSbAABTnAAAPAECAICcAABFngAAmAACAEieAAClngAAUPcBAKieAAAuoAAAxAECADCgAACjoAAAdPgBAKSgAACzoQAA6AECALShAAD1oQAAMAECAPihAAASogAAiPYBABSiAAAuogAAiPYBADCiAABzogAA2AECAHyiAAC3ogAAGAICALiiAADYowAA5P4BANijAAAbpAAAAAICAFikAAB7pAAAiPYBAICkAACQpAAAiPYBAJCkAADNpAAAUPcBANikAAAYpQAAUPcBABilAABzpQAAiPYBAIilAACcpQAAiPYBAJylAACspQAAiPYBAKylAADhpQAAiPYBAOSlAAD0pQAAiPYBAPSlAAAepgAAPAICACCmAAB/pgAAUPcBAJCmAAAGpwAAEPsBACCnAACdpwAAYAICAMynAAAUqAAArAICABSoAABCqAAAZAICAGSoAADdqgAAbAICAOCqAAAfqwAA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wK7IrtCu2K7gruiu8K74rgCvCK8QrxivIK8orzCvOK9Ar0ivUK9Yr2CvaK9wr3ivgK+Ir5CvmK+gr6ivsK+4r8CvyK/Qr9iv4K/or/CvAAAAgAEAlAAAAACgCKAQoBigIKAooDCgOKBAoEigUKBYoGCgaKBwoHiggKCIoJCgmKCgoKigsKC4oMCgyKDQoNig4KDooPCg+KAAoQihEKEYoSChKKEwoTihQKFIoVChWKEgpSilMKU4pYCliKWQpZiloKWopbCluKXApcil0KXYpeCl6KXwpfilAKYIphCmGKYgpgAAAJABAJwBAAB4o4ijmKOoo7ijyKPYo+ij+KMIpBikKKQ4pEikWKRopHikiKSYpKikuKTIpNik6KT4pAilGKUopTilSKVYpWileKWIpZilqKW4pcil2KXopfilCKYYpiimOKZIplimaKZ4poimmKaoprimyKbYpuim+KYIpxinKKc4p0inWKdop3iniKeYp6inuKfIp9in6Kf4pwioGKgoqDioSKhYqGioeKiIqJioqKi4qMio2KjoqPioCKkYqSipOKlIqVipaKl4qYipmKmoqbipyKnYqeip+KkIqhiqKKo4qkiqWKpoqniqiKqYqqiquKrIqtiq6Kr4qgirGKsoqzirSKtYq2ireKuIq5irqKu4q8ir2Kvoq/irCKwYrCisOKxIrFisaKx4rIismKyorLisyKzYrOis+KwIrRitKK04rUitWK1orXitiK2YraituK3Irdit6K34rQiuGK4orjiuSK5YrmiueK6IrpiuqK64rsiu2K7orviuCK8YryivOK9Ir1ivaK94r4ivmK+or7ivyK/Yr+iv+K8AAACgAQCIAAAACKAYoCigOKBIoFigaKB4oIigmKCooLigyKDYoOig+KAIoRihKKE4oUihWKFooXihiKGYoaihwK3QreCt8K0ArhCuIK4wrkCuUK5grnCugK6QrqCusK7ArtCu4K7wrgCvEK8grzCvQK9Qr2CvcK+Ar5CvoK+wr8Cv0K/gr/CvAAAAsAEAiAEAAACgEKAgoDCgQKBQoGCgcKCAoJCgoKCwoMCg0KDgoPCgAKEQoSChMKFAoVChYKFwoYChkKGgobChwKHQoeCh8KEAohCiIKIwokCiUKJgonCigKKQoqCisKLAotCi4KLwogCjEKMgozCjQKNQo2CjcKOAo5CjoKOwo8Cj0KPgo/CjAKQQpCCkMKRApFCkYKRwpICkkKSgpLCkwKTQpOCk8KQApRClIKUwpUClUKVgpXClgKWQpaClsKXApdCl4KXwpQCmEKYgpjCmQKZQpmCmcKaAppCmoKawpsCm0KbgpvCmAKcQpyCnMKdAp1CnYKdwp4CnkKegp7CnwKfQp+Cn8KcAqBCoIKgwqECoUKhgqHCogKiQqKCosKjAqNCo4KjwqACpEKkgqTCpQKlQqWCpcKmAqZCpoKmwqcCp0KngqfCpAKoQqiCqMKpAqlCqYKpwqoCqkKqgqrCqwKrQquCq8KoAqxCrIKswq0CrUKtgq3CrgKuQq6CrsKvAq9Cr4KvwqwDgAQBYAAAAiKGQoZihoKGooUCiSKJQoliiYKJoonCieKKAooii8KL4ogCjCKMQoxijIKMoozCjcKN4o4CjiKOQo5ij0KQ4pVClWKXgpfilAKYIphCmGKYAIAIAdAAAAMCgCKEooUihaKGIobih0KHYoeChGKIgoniogKiIqJComKigqKiosKi4qMCoyKjYqOCo6KjwqPioAKkIqRCpUKp4qqCq0KoIq0CreKuwq+irIKxYrJis0KwIrVCtkK3ArfCtGK5ArmiuoK7ArgAwAgAcAAAAkKDQoZCiUKTYpAilQKVwpbClCKYAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
'@

Export-ModuleMember -Function Test-ProcessElevated, Get-Privilege, Test-DebugPrivilege, Enable-DebugPrivilege, Disable-DebugPrivilege, Start-WamTrace, Stop-WamTrace, Start-OutlookTrace, Stop-OutlookTrace, Start-NetshTrace, Stop-NetshTrace, Start-PSR, Stop-PSR, Save-EventLog, Get-InstalledUpdate, Save-OfficeRegistry, Get-ProxySetting, Get-WinInetProxy, Get-WinHttpDefaultProxy, Get-ProxyAutoConfig, Save-OSConfiguration, Get-NLMConnectivity, Get-WSCAntivirus, Save-CachedAutodiscover, Remove-CachedAutodiscover, Save-CachedOutlookConfig, Remove-CachedOutlookConfig, Remove-IdentityCache, Start-LdapTrace, Stop-LdapTrace, Get-OfficeModuleInfo, Save-OfficeModuleInfo, Start-CAPITrace, Stop-CapiTrace, Start-FiddlerCap, Start-FiddlerEverywhereReporter, Start-Procmon, Stop-Procmon, Start-TcoTrace, Stop-TcoTrace, Get-ConnTimeout, Set-ConnTimeout, Remove-ConnTimeout, Get-OfficeInfo, Add-WerDumpKey, Remove-WerDumpKey, Start-WfpTrace, Stop-WfpTrace, Save-Dump, Save-HungDump, Save-MSIPC, Save-MIP, Enable-DrmExtendedLogging, Disable-DrmExtendedLogging, Get-DRMConfig, Get-EtwSession, Stop-EtwSession, Get-Token, Test-Autodiscover, Get-LogonUser, Get-JoinInformation, Get-OutlookProfile, Get-OutlookAddin, Get-ClickToRunConfiguration, Get-WebView2, Get-DeviceJoinStatus, Save-NetworkInfo, Download-TTD, Expand-TTDMsixBundle, Install-TTD, Uninstall-TTD, Start-TTDMonitor, Stop-TTDMonitor, Cleanup-TTD, Attach-TTD, Detach-TTD, Start-PerfTrace, Stop-PerfTrace, Start-Wpr, Stop-Wpr, Get-IMProvider, Get-MeteredNetworkCost, Save-PolicyNudge, Save-CLP, Save-DLP, Invoke-WamSignOut, Enable-PageHeap, Disable-PageHeap, Get-OfficeIdentityConfig, Get-OfficeIdentity, Get-OneAuthAccount, Remove-OneAuthAccount, Get-AlternateId, Get-UseOnlineContent, Get-AutodiscoverConfig, Get-SocialConnectorConfig, Get-ImageFileExecutionOptions, Start-Recording, Stop-Recording, Get-OutlookOption, Get-WordMailOption, Get-ImageInfo, Get-PresentationMode, Get-AnsiCodePage, Get-PrivacyPolicy, Save-GPResult, Get-AppContainerRegistryAcl, Get-StructuredQuerySchema, Get-NetFrameworkVersion, Get-MapiCorruptFiles, Save-MonarchLog, Save-MonarchSetupLog, Enable-WebView2DevTools, Disable-WebView2DevTools, Enable-WebView2Netlog, Disable-WebView2Netlog, Get-WebView2Flags, Add-WebView2Flags, Remove-WebView2Flags, Get-FileExtEditFlags, Get-ExperimentConfigs, Get-CloudSettings, Get-ProcessWithModule, Get-PickLogonProfile, Enable-PickLogonProfile, Disable-PickLogonProfile, Enable-AccountSetupV2, Disable-AccountSetupV2, Save-USOSharedLog, Receive-WinRTAsyncResult, Get-WebAccount, Get-WebAccountProvider, Get-TokenSilently, Invoke-WebAccountSignOut, Invoke-RequestToken, Collect-OutlookInfo, Save-WamInteropDll