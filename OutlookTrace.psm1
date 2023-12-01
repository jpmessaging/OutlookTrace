<#
.NOTES
Copyright (c) 2021 Ryusuke Fujita

This software is released under the MIT License.
http://opensource.org/licenses/mit-license.php

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
#>

$Version = 'v2023-11-30'
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

        // https://learn.microsoft.com/en-us/windows/win32/api/securitybaseapi/nf-securitybaseapi-gettokeninformation
        [DllImport("advapi32.dll", ExactSpelling = true, SetLastError = true)]
        public static extern bool GetTokenInformation(
            IntPtr TokenHandle,
            int TokenInformationClass,
            out TOKEN_ELEVATION TokenElevation,
            uint TokenInformationLength,
            out uint ReturnLength);

        // https://learn.microsoft.com/en-us/windows/win32/api/winnt/ns-winnt-token_elevation
        [StructLayout(LayoutKind.Sequential)]
        public struct TOKEN_ELEVATION
        {
            public bool TokenIsElevated;
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

        public override bool IsInvalid
        {
            get { return IsClosed || handle == IntPtr.Zero; }
        }

        override protected bool ReleaseHandle()
        {
            Marshal.FreeCoTaskMem(handle);
            return true;
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
}
'@

function Open-Log {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$Path,
        [switch]$AutoFlush
    )

    if ($Script:LogWriter) {
        Close-Log
    }

    # Open a file & add header
    try {
        [IO.StreamWriter]$Script:LogWriter = [IO.File]::AppendText($Path)

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
            Write-Log "Closing LogWriter."
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

    Write-Log "Setting up a Runspace Pool with an initialSessionState. MinRunspaces:$MinRunspaces, MaxRunspaces:$MaxRunspaces."
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

    Write-Log "RunspacePool ($($Script:RunspacePool.InstanceId.ToString())) is opened."
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
    Write-Log "RunspacePool ($id) is closed."
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
    Write-Error "Timeout."
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
        Write-Error -Message "Open-TaskRunspace must be called in advance."
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
        [int]$Id,
        [switch]$EnableDebugPrivilege
    )

    begin {
        # Enable Debug privilege if possible
        $debugPrivilegeEnabled = $false

        if ($EnableDebugPrivilege) {
            try {
                [System.Diagnostics.Process]::EnterDebugMode()
                $debugPrivilegeEnabled = $true
            }
            catch {
                Write-Log -Message "EnterDebugMode failed" -ErrorRecord $_ -Category Warning
            }
        }
    }

    process {
        $hProcess = $null
        $hToken = [IntPtr]::Zero

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

            # https://learn.microsoft.com/en-us/windows/win32/api/winnt/ne-winnt-token_information_class
            $TokenElevation = 20

            $length = 0;
            $elevation = New-Object Win32.Advapi32+TOKEN_ELEVATION

            if (-not [Win32.Advapi32]::GetTokenInformation(
                    $hToken,
                    $TokenElevation,
                    [ref]$elevation,
                    [System.Runtime.InteropServices.Marshal]::SizeOf([type]'Win32.Advapi32+TOKEN_ELEVATION'),
                    [ref]$length)) {
                $ex = New-Object System.ComponentModel.Win32Exception -ArgumentList ([System.Runtime.InteropServices.Marshal]::GetLastWin32Error())
                Write-Error -Message "GetTokenInformation failed for PID $Id. $($ex.Message) (NativeErrorCode:$($ex.NativeErrorCode))" -Exception $ex
                return
            }

            $elevation.TokenIsElevated
        }
        finally {
            if ($hToken) {
                $null = [Win32.Kernel32]::CloseHandle($hToken)
            }

            if ($hProcess) {
                $hProcess.Dispose()
            }
        }
    }

    end {
        if ($debugPrivilegeEnabled) {
            [System.Diagnostics.Process]::LeaveDebugMode()
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
            $Path = Resolve-Path -LiteralPath $Path
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
            $Destination = Resolve-Path -LiteralPath $Destination
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
            $Path = Resolve-Path -LiteralPath $Path
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
            $Destination = Resolve-Path -LiteralPath $Destination
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
            $Path = Resolve-Path $Path
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
            $Destination = Resolve-Path -LiteralPath $Destination
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
            Write-Error -Message "There are $($files.Count) files in $Path, but none can be opened."
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
        Write-Error "FromDateTime must be less than or equal to ToDateTime."
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

function Enable-EventLog {
    [CmdletBinding(PositionalBinding = $false)]
    param(
        [Parameter(Mandatory = $true, Position = 0)]
        [string]$EventName
    )

    $($null = wevtutil.exe set-log $EventName /enabled:true /retention:false /quiet:true) 2>&1 | ForEach-Object {
        Write-Error -ErrorRecord $_
    }
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

function Enable-WamEventLog {
    [CmdletBinding(PositionalBinding = $false)]
    param()

    Enable-EventLog 'Microsoft-Windows-WebAuth/Operational'
    Enable-EventLog 'Microsoft-Windows-WebAuthN/Operational'
    Enable-EventLog 'Microsoft-Windows-AAD/Operational'
    Enable-EventLog 'Microsoft-Windows-AAD/Analytic'
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
        [string]$FileName = 'wam.etl',
        [string]$SessionName = 'WamTrace',
        [ValidateSet('NewFile', 'Circular')]
        [string]$LogFileMode = 'NewFile',
        [ValidateRange(1, [int]::MaxValue)]
        [int]$MaxFileSizeMB = 256
    )

    if (-not (Test-Path $Path)) {
        $null = New-Item $Path -ItemType Directory -ErrorAction Stop
    }
    $Path = Resolve-Path $Path

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

    Write-Log "Starting a WAM trace."
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
        [string]$FileName = 'outlook.etl',
        [string]$SessionName = 'OutlookTrace',
        [ValidateSet('NewFile', 'Circular')]
        [string]$LogFileMode = 'NewFile',
        [ValidateRange(1, [int]::MaxValue)]
        [int]$MaxFileSizeMB = 256
    )

    if (-not (Test-Path $Path)) {
        $null = New-Item $Path -ItemType Directory -ErrorAction Stop
    }

    $Path = Resolve-Path $Path
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
        $FileName = 'nettrace-winhttp-webio.etl',
        [ValidateSet('None', 'Mini', 'Full')]
        $ReportMode = 'None'
    )

    if (-not (Test-Path $Path)) {
        $null = New-Item $Path -ItemType Directory -ErrorAction Stop
    }
    $Path = Resolve-Path $Path

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
        Write-Error "Cannot find $netshexe."
        return
    }

    Write-Log "Clearing dns cache"
    $null = & ipconfig.exe /flushdns

    Write-Log "Starting a netsh trace."
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

    Write-Log "ReportMode $ReportMode is configured."
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
            Write-Log "$SessionName was not found. Retrying after $retry seconds."
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

    Write-Log "ReportMode $reportMode is found."

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
        Write-Error "Cannot find $netshexe."
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
        Write-Error "psr.exe is not available."
        return
    }

    # Make sure psr isn't running already.
    $processes = @(Get-Process psr -ErrorAction SilentlyContinue)

    if ($processes.Count -gt 0) {
        Write-Error "PSR is already running. ($($processes.ID -join ','))"
        return
    }

    if (-not (Test-Path $Path -ErrorAction Stop)) {
        $null = New-Item -ItemType Directory $Path -ErrorAction Stop
    }

    $Path = Resolve-Path $Path

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

    $err = $($process = Start-Process 'psr' -ArgumentList $psrArgs -PassThru) 2>&1

    if (-not $process -or $process.HasExited) {
        Write-Error "PSR failed to start. $err" -Exception $err.Exception
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
    $err = $($stopInstance = Start-Process 'psr' -ArgumentList '/stop' -PassThru)

    if (-not $stopInstance) {
        Write-Error "Failed to run psr.exe /stop. $err" -Exception $err.Exception
        return
    }

    # Do not use Wait-Process here because it can fail with Access denied when running as non-admin
    $currentInstance.WaitForExit()
    Write-Log "PSR (PID:$($currentInstance.Id)) is stopped. ExitCode:$($currentInstance.ExitCode)"

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
    $Path = Resolve-Path $Path

    $logs = @(
        'Application'
        'System'
        (wevtutil el) -match "Microsoft-Windows-Windows Firewall With Advanced Security|AAD|Microsoft-Windows-Bits-Client|WebAuth|CAPI2|AppLocker"
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

<#
.SYNOPSIS
    Get-MicrosoftUpdate
.NOTES
    Deprecated. Use Get-InstalledUpdate instead.
#>
function Get-MicrosoftUpdate {
    [CmdletBinding()]
    param(
        [switch]$OfficeOnly,
        [switch]$AppliedOnly
    )

    # Constants
    # https://docs.microsoft.com/en-us/windows/desktop/api/msi/nf-msi-msienumpatchesexa
    $PatchState = @{
        1  = 'MSIPATCHSTATE_APPLIED'
        2  = 'MSIPATCHSTATE_SUPERSEDED'
        4  = 'MSIPATCHSTATE_OBSOLETED'
        8  = 'MSIPATCHSTATE_REGISTERED'
        15 = 'MSIPATCHSTATE_ALL'
    }

    $hklm = $productsKey = $null
    try {
        # Use .NET registry API (with [RegistryView]::Registry64) instead of PowerShell here to avoid registry redirection occurs on 32bit PowerShell on 64bit OS for HKLM\SOFTWARE.
        if ('Microsoft.Win32.RegistryView' -as [type]) {
            $hklm = [Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::LocalMachine, [Microsoft.Win32.RegistryView]::Registry64);
        }
        elseif (-not $env:PROCESSOR_ARCHITEW6432) {
            # RegistryView is not available, but it's OK because no WOW64.
            $hklm = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey([Microsoft.Win32.RegistryHive]::LocalMachine, [string]::Empty);
        }
        else {
            # This is the case where registry rediction takes place (32bit PowerShell on 64bit OS). Bail.
            Write-Error "32bit PowerShell is running on 64bit OS and .NET 4.0 is not used. Please run 64bit PowerShell."
            return
        }

        $productsKey = $hklm.OpenSubKey('SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData\S-1-5-18\Products')

        foreach ($productName in $productsKey.GetSubKeyNames()) {
            if ($null -eq $productName -or ($OfficeOnly -and $productName -notmatch 'F01FEC')) {
                continue
            }

            $productKey = $productsKey.OpenSubKey($productName)

            foreach ($subkeyName in $productKey.GetSubKeyNames()) {
                if ($subkeyName -ne 'Patches') {
                    continue
                }

                $patchesKey = $productKey.OpenSubKey($subkeyName)
                foreach ($patchName in $patchesKey.GetSubKeyNames()) {
                    $patchKey = $patchesKey.OpenSubKey($patchName)

                    $state = $patchKey.GetValue('State')

                    if ($AppliedOnly -and $PatchState[$state] -ne 'MSIPATCHSTATE_APPLIED') {
                        continue
                    }

                    $displayName = $patchKey.GetValue('DisplayName')
                    $moreInfoURL = $patchKey.GetValue('MoreInfoURL')
                    $installed = $patchKey.GetValue('Installed')

                    if (-not $displayName -and -not $moreInfoURL) {
                        continue
                    }

                    # extract KB number
                    $KB = $null
                    if ($moreInfoURL -match 'https?://support.microsoft.com/kb/(?<KB>\d+)') {
                        $KB = $Matches['KB']
                    }

                    [PSCustomObject]@{
                        DisplayName = $displayName
                        KB          = $KB
                        MoreInfoURL = $moreInfoURL
                        Installed   = $installed
                        PatchState  = $PatchState[$state]
                    }

                    $patchKey.Close()
                }

                $patchesKey.Close()
            }

            $productKey.Close()
        }
    }
    finally {
        if ($productsKey) {
            $productsKey.Close()
        }

        if ($hklm) {
            $hklm.Close()
        }
    }
}

<#
.SYNOPSIS
    Save-MicrosoftUpdate
.NOTES
    Deprecated. Use Get-InstalledUpdate instead.
#>
function Save-MicrosoftUpdate {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        $Path
    )

    if (-not (Test-Path $Path)) {
        $null = New-Item $Path -ItemType directory -ErrorAction Stop
    }

    $cmdletName = $PSCmdlet.MyInvocation.MyCommand.Name
    $name = $cmdletName.Substring($cmdletName.IndexOf('-') + 1)
    Get-MicrosoftUpdate | Export-Clixml -Depth 4 -Path $(Join-Path $Path -ChildPath "$name.xml")
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

    # Is is SID?
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
                $owner = Invoke-CimMethod -InputObject $win32Process -MethodName GetOwnerSid

                # Without admin privilege, owner info other than self can be empty.
                if (-not $owner.Sid) {
                    Write-Verbose "Cannot obtain the owner of explorer (PID $($win32Process.ProcessID)). Probably you are runnning without admin privilege."
                    return
                }

                if (-not $usersCache.ContainsKey($owner.Sid)) {
                    $usersCache.Add($owner.Sid, $null)
                    Resolve-User $owner.Sid
                }
            }
            finally {
                $win32Process.Dispose()
            }
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
        [string]$User,
        # Skip "Registry::" prefix
        [switch]$SkipRegistryPrefix
    )

    if ($User) {
        $resolvedUser = Resolve-User $User

        if (-not $resolvedUser) {
            return
        }

        $userRegRoot = "HKEY_USERS\$($resolvedUser.Sid)"

        if (-not ($userRegRoot -and (Test-Path "Registry::$userRegRoot"))) {
            Write-Error "Cannot find $userRegRoot."
            return
        }
    }
    else {
        Write-Log "User is empty. Use HKCU."
        $userRegRoot = 'HKCU'
    }

    if (-not $SkipRegistryPrefix) {
        $userRegRoot = "Registry::$userRegRoot"
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

        # HKLM
        'HKLM\SOFTWARE\Microsoft\Office'
        'HKLM\SOFTWARE\WOW6432Node\Microsoft\Office'
        'HKLM\SOFTWARE\Microsoft\MSIPC'
        'HKLM\SOFTWARE\WOW6432Node\Microsoft\MSIPC'
        'HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Internet Settings\WinHttp'
        'HKLM\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Internet Settings\WinHttp'
        'HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies'
        'HKLM\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols'

        # Policies
        'HKCU\Software\Policies'
        'HKCU\Software\Wow6432Node\Policies'
        'HKLM\SOFTWARE\Policies'
        'HKLM\SOFTWARE\WOW6432Node\Policies'
    )

    $userRegRoot = Get-UserRegistryRoot $User -SkipRegistryPrefix

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
        Write-Error "$regexe is not avaialble."
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
        @{ScriptBlock = { Get-CimInstance -Class Win32_ComputerSystem }; FileName = 'Win32_ComputerSystem.xml' }
        @{ScriptBlock = { Get-CimInstance -Class Win32_OperatingSystem }; FileName = 'Win32_OperatingSystem.xml' }
        @{ScriptBlock = { Get-CimInstance -Class Win32_Processor }; FileName = 'Win32_Processor.xml' }
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
        @{ScriptBlock = { Get-AppxPackage } }
        @{ScriptBlock = { Get-Service } }
        @{ScriptBlock = { Get-SmbMapping } }
        @{ScriptBlock = { Get-AnsiCodePage } }
        @{ScriptBlock = { cmdkey /list }; FileName = 'cmdkey.txt' }

        $userArg = @{ User = $User }
        @{ScriptBlock = { param($User) Get-WebView2 @PSBoundParameters }; ArgumentList = $userArg }
        @{ScriptBlock = { param($User) Get-WinInetProxy @PSBoundParameters }; ArgumentList = $userArg }
        @{ScriptBlock = { param($User) Get-ProxyAutoConfig @PSBoundParameters }; ArgumentList = $userArg }
        @{ScriptBlock = { param($User) Get-AppContainerRegistryAcl @PSBoundParameters }; ArgumentList = $userArg }
        @{ScriptBlock = { param($User) Get-StructuredQuerySchema @PSBoundParameters }; ArgumentList = $userArg }

        # These are just for troubleshooting.
        @{ScriptBlock = { Get-ChildItem 'Registry::HKEY_USERS' | Select-Object 'Name' }; FileName = 'Users.xml' }
        @{ScriptBlock = { whoami.exe /USER }; FileName = 'whoami.txt' }
    } | & {
        process {
            if ($CancellationToken.IsCancellationRequested) {
                return
            }

            Invoke-ScriptBlock @_ -Path $Path
        }
    }
}

function Save-OSConfigurationMT {
    [CmdletBinding()]
    param (
        [parameter(Mandatory = $true)]
        $Path
    )

    # If this command is run by itself (not from Collect-OutlookInfo), need to create a runspace pool.
    if (-not $Script:RunspacePool) {
        Open-TaskRunspace
        $runspaceOpened = $true
    }

    if (-not (Test-Path $Path)) {
        $null = New-Item $Path -ItemType directory -ErrorAction Stop
    }

    $tasks = @(
        Start-Task { param($path) Get-WmiObject -Class Win32_ComputerSystem | Export-Clixml -Path $path } -ArgumentList (Join-Path $Path "Win32_ComputerSystem.xml")
        Start-Task { param($path) Get-WmiObject -Class Win32_OperatingSystem | Export-Clixml -Path $path } -ArgumentList (Join-Path $Path "Win32_OperatingSystem.xml")
        Start-Task { param($path) Get-ProxySetting | Export-Clixml -Path $path } -ArgumentList (Join-Path $Path "ProxySetting.xml")
        Start-Task { param($path) Get-NLMConnectivity | Export-Clixml -Path $path } -ArgumentList (Join-Path $Path "NLMConnectivity.xml")
        Start-Task { param($path) Get-WSCAntivirus -ErrorAction SilentlyContinue | Export-Clixml -Path $path } -ArgumentList (Join-Path $Path "WSCAntivirus.xml")
        Start-Task { param($path) Get-InstalledUpdate -ErrorAction SilentlyContinue | Export-Clixml -Path $path } -ArgumentList (Join-Path $Path "InstalledUpdate.xml")
        Start-Task { param($path) Get-JoinInformation -ErrorAction SilentlyContinue | Export-Clixml -Path $path } -ArgumentList (Join-Path $Path "JoinInformation.xml")
        Start-Task { param($path) Get-DeviceJoinStatus -ErrorAction SilentlyContinue | Out-File -FilePath $path } -ArgumentList (Join-Path $Path "DeviceJoinStatus.txt")
    )

    Write-Verbose "waiting for tasks..."
    $tasks | Receive-Task -AutoRemoveTask

    if ($runspaceOpened) {
        Close-TaskRunspace
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

    # Wrap is an array in case a single object is passed as ArgumentList (otherwise, splatting does not work as expected)
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
    Write-Log "$scriptBlockName took $($elapsed.TotalMilliseconds) ms.$(if ($null -eq $result) {" It returned nothing."})"

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
        # 1. It the ScriptBlock has FunctionDefinitionAst, use its Name (that is, the function name)
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

function Save-NetworkInfoMT {
    [CmdletBinding()]
    param (
        [parameter(Mandatory = $true)]
        $Path
    )

    if (-not (Test-Path $Path)) {
        $null = New-Item $Path -ItemType directory -ErrorAction Stop
    }

    # Path must be resolved before it's used as an argument to Start-Task.
    $Path = Resolve-Path -LiteralPath $Path
    $PSDefaultParameterValues.Add('Start-Task:ArgumentList', $Path)

    # These are from C:\Windows\System32\gatherNetworkInfo.vbs with some extra.
    $tasks = @(
        Start-Task { param ($Path) Get-NetAdapter -IncludeHidden | Export-Clixml (Join-Path $Path 'NetAdapter.xml') }
        Start-Task { param ($Path) Get-NetAdapterAdvancedProperty | Export-Clixml (Join-Path $Path 'NetAdapterAdvancedProperty.xml') }
        Start-Task { param ($Path) Get-NetAdapterBinding -IncludeHidden | Export-Clixml (Join-Path $Path 'NetAdapterBinding.xml') }
        Start-Task { param ($Path) Get-NetIpConfiguration -Detailed | Export-Clixml (Join-Path $Path 'NetIpConfiguration.xml') }
        Start-Task { param ($Path) Get-DnsClientNrptPolicy | Export-Clixml (Join-Path $Path 'DnsClientNrptPolicy.xml') }
        Start-Task { param ($Path) Get-NetRoute | Export-Clixml (Join-Path $Path 'NetRoute.xml') }
        Start-Task { param ($Path) Get-NetIPaddress | Export-Clixml (Join-Path $Path 'NetIPaddress.xml') }
        Start-Task { param ($Path) Get-NetLbfoTeam | Export-Clixml (Join-Path $Path 'NetLbfoTeam.xml') }
        Start-Task { param ($Path) Get-NetIPInterface | Export-Clixml (Join-Path $Path 'NetIPInterface.xml') }
        Start-Task { param ($Path) Get-NetConnectionProfile | Export-Clixml (Join-Path $Path 'NetConnectionProfile.xml') }
        Start-Task { param ($Path) ipconfig /all | Out-File (Join-Path $Path 'ipconfig_all.txt') }
        Start-Task { param ($Path) netsh advfirewall monitor show currentprofile | Out-File (Join-Path $Path 'netsh advfirewall monitor show currentprofile.txt') }
        Start-Task { param ($Path) netsh advfirewall monitor show firewall | Out-File (Join-Path $Path 'netsh advfirewall monitor show firewall.txt') }
        Start-Task { param ($Path) netsh advfirewall firewall show rule name=all verbose | Out-File (Join-Path $Path 'netsh advfirewall firewall show rule name=all verbose.txt') }
        Start-Task { param ($Path) netsh advfirewall consec show rule name=all verbose | Out-File (Join-Path $Path 'netsh advfirewall consec show rule name=all verbose.txt') }
        Start-Task { param ($Path) netsh advfirewall monitor show firewall rule name=all | Out-File (Join-Path $Path 'netsh advfirewall monitor show firewall rule name=all.txt') }
        Start-Task { param ($Path) netsh advfirewall monitor show consec rule name=all | Out-File (Join-Path $Path 'netsh advfirewall monitor show consec rule name=all.txt') }
    )

    $PSDefaultParameterValues.Remove('Start-Task:ArgumentList')

    Write-Log "Waiting for tasks to complete."
    $tasks | Receive-Task -AutoRemoveTask
    Write-Log "All tasks are complete."
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
        Write-Error ("Win32 WinHttpGetIEProxyConfigForCurrentUser failed with 0x{0:x8}" -f [System.Runtime.InteropServices.Marshal]::GetLastWin32Error())
    }

    # If ProxySettingsPerUser is 0, then check HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Internet Settings\Connections, instead of the user's registry.
    $proxySettingsPerUser = Get-ItemProperty 'Registry::HKLM\SOFTWARE\Policies\Microsoft\Windows\CurrentVersion\Internet Settings' -Name 'ProxySettingsPerUser' -ErrorAction SilentlyContinue `
    | Select-Object -ExpandProperty 'ProxySettingsPerUser'

    if ($proxySettingsPerUser -eq 0) {
        $regRoot = 'Registry::HKLM'
    }
    else {
        $err = $($regRoot = Get-UserRegistryRoot -User $User) 2>&1
        if (-not $regRoot) {
            Write-Error "Cannot get user $User's registry root. $err"
            return
        }
    }

    # There might be multiple connections besides "DefaultConnectionSettings" if there are VPNs.
    $connectionsKey = Join-Path $regRoot 'SOFTWARE\Microsoft\Windows\CurrentVersion\Internet Settings\Connections'
    $connections = @(Get-Item $connectionsKey | Select-Object -ExpandProperty Property)

    # It's possible that there is no connection at all (maybe IE has never been started).
    # In this case, return the default configuration (This is what WinHttpGetIEProxyConfigForCurrentUser does anyway).
    if ($connections.Count -eq 0) {
        Write-Log "No connections are found under $connectionsKey. Returning a default setting."

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
            Write-Error "Get-ItemProperty failed for a connection, $connection. $_" -Exception $_.Exception
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
                        Write-Log "Skipped $pacUrl because it's already tried."
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
                    Write-Log "Skipped $pacUrl because it's already tried."
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
        Write-Error "NetGetJoinInformation failed with $sc." -Exception (New-Object ComponentModel.Win32Exception($sc))
        return
    }

    $name = [System.Runtime.InteropServices.Marshal]::PtrToStringUni($pName)
    $sc = [Win32.Netapi32]::NetApiBufferFree($pName)

    if ($sc -ne 0) {
        Write-Error "NetApiBufferFree failed with $sc." -Exception (New-Object ComponentModel.Win32Exception($sc))
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
                    Write-Log "Profile $profileName does not have mail accounts" -Category Warning
                }

                [PSCustomObject]@{
                    User           = $User
                    Name           = $profileName
                    Path           = $prof.Name
                    IsDefault      = $profileName -eq $defaultProfile
                    Accounts       = $mailAccounts | Select-Object -Property * -ExcludeProperty 'EmsmdbUid'
                    StoreProviders = $storeProviders
                    DataFiles      = $dataFiles
                    OfflineState   = $globalSection.OfflineState
                    CacheSyncMode  = $globalSection.CacheSyncMode
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

    if ($cachedModePolicy.enable) {
        $props.Enable = $cachedModePolicy.enable -eq 1
    }

    if ($cachedModePolicy.cachedexchangemode) {
        $props.SyncMode = [Win32.Mapi.CacheSyncMode]([int]::Parse($cachedModePolicy.cachedexchangemode))
    }

    $props.SyncWindow = Get-SyncWindow -Days $cachedModePolicy.syncwindowsettingdays -Months $cachedModePolicy.syncwindowsetting

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
        Eamil                = $Account.Email
        Pop3Server           = $Account.'POP3 Server'
        Pop3Port             = if ($Account.'Pop3 Port') { $Account.'Pop3 Port' } else { $Pop3DefaultPort }
        Pop3User             = $Account.'POP3 User'
        SmtpServer           = $Account.'SMTP Server'
        SmtpPort             = if ($Account.'SMTP Port') { $Account.'SMTP Port' } else { $SmtpDefaultPort }
        SmtpUser             = $Account.'SMTP User'
        SmtpUseAuth          = $Account.'SMTP Use Auth' -eq 1
        SmtpUseSPA           = $Account.'SMTP Use SPA' -eq 1
        SmtpSecureConnection = switch ($Account.'SMTP Secure Connection') { 0 { 'None'; break } 1 { 'SSL'; break } 2 { 'TLS'; break } 3 { 'Auto'; break } }
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
        Eamil                = $Account.Email
        ImapServer           = $Account.'IMAP Server'
        ImapPort             = if ($Account.'IMAP Port') { $Account.'IMAP Port' } else { $ImapDefaultPort }
        ImapUser             = $Account.'IMAP User'
        ImapUseSPA           = $Account.'IMAP Use SPA' -eq 1
        SmtpServer           = $Account.'SMTP Server'
        SmtpPort             = if ($Account.'SMTP Port') { $Account.'SMTP Port' } else { $SmtpDefaultPort }
        SmtpUser             = $Account.'SMTP User'
        SmtpUseAuth          = $Account.'SMTP Use Auth' -eq 1
        SmtpUseSPA           = $Account.'SMTP Use SPA' -eq 1
        SmtpSecureConnection = switch ($Account.'SMTP Secure Connection') { 0 { 'None'; break } 1 { 'SSL'; break } 2 { 'TLS'; break } 3 { 'Auto'; break } }
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
    )

    $globalSection = Join-Path $Profile.PSPath $KnownSections.Global | Get-ItemProperty -Name $properties -ErrorAction SilentlyContinue

    # It's possible that Global Section does not have any of above properties and thus null.
    $offlineState = 'Unknown'
    $cacheSyncMode = [Win32.Mapi.CacheSyncMode]::FullItems

    if ($globalSection) {
        if ($offlineStateBin = $globalSection.$($PropTags.PR_LAST_OFFLINESTATE_OFFLINE)) {
            [Win32.Mapi.OfflineState]$offlineState = [BitConverter]::ToInt32($offlineStateBin, 0) -band [Win32.Mapi.OfflineState]::Mask
        }

        if ($syncModeBin = $globalSection.$($PropTags.PR_CACHE_SYNC_MODE)) {
            $cacheSyncMode = [BitConverter]::ToInt32($syncModeBin, 0)
        }
    }

    [PSCustomObject]@{
        DisplayName   = 'Outlook Global Section'
        Uid           = $KnownSections.Global
        OfflineState  = $offlineState
        CacheSyncMode = $cacheSyncMode
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
                $props.DisplayName = [System.Text.Encoding]::Unicode.GetString($displayNameBin)
            }

            if ($resourceFlagsBin = $store.$($PropTags.PR_RESOURCE_FLAGS)) {
                $props.ResourceFlags = [Win32.Mapi.ResourceFlags][BitConverter]::ToUInt32($resourceFlagsBin, 0)
            }

            if ($alternateStoreTypeBin = $store.$($PropTags.PR_PROFILE_ALTERNATE_STORE_TYPE)) {
                $props.AlternateStoreType = [System.Text.Encoding]::Unicode.GetString($alternateStoreTypeBin)
            }

            if ($pstPath = $store.$($PropTags.PR_PROFILE_PST_PATH)) {
                $props.PstPath = [System.Text.Encoding]::Unicode.GetString($pstPath, 0, $pstPath.Length - 2)
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
                $props.UserSmtpEmailAddress = [System.Text.Encoding]::Unicode.GetString($userSmtpEmailAddressBin)
            }

            if ($tenantIdBin = $store.$($PropTags.PR_PROFILE_TENANT_ID)) {
                $props.TenantId = [System.Text.Encoding]::Unicode.GetString($tenantIdBin)
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

    $serviceUid = [BitConverter]::ToString($account.'Service UID').Replace('-', [String]::Empty)
    $service = Join-Path $profRoot $serviceUid | Get-ItemProperty -Name $PropTags.PR_EMSMDB_SECTION_UID -ErrorAction SilentlyContinue
    $emsmdbUid = [BitConverter]::ToString($service.$($PropTags.PR_EMSMDB_SECTION_UID)).Replace('-', '').ToLowerInvariant()

    # Get EMSMDB section properties
    $emsmdbProperties = @(
        $PropTags.PR_DISPLAY_NAME
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

    if ($displayNameBin = $emsmdb.$($PropTags.PR_DISPLAY_NAME)) {
        $props.DisplayName = [System.Text.Encoding]::Unicode.GetString($displayNameBin)
    }

    if ($identityUniqueIdBin = $emsmdb.$($PropTags.PR_EMSMDB_IDENTITY_UNIQUEID)) {
        $props.IdentityUniqueId = [System.Text.Encoding]::Unicode.GetString($identityUniqueIdBin)
    }

    if ($userFullNameBin = $emsmdb.$($PropTags.PR_PROFILE_USER_FULL_NAME)) {
        $props.UserFullName = [System.Text.Encoding]::Unicode.GetString($userFullNameBin)
    }

    if ($ostPath = $emsmdb.$($PropTags.PR_PROFILE_OFFLINE_STORE_PATH)) {
        $props.OstPath = [System.Text.Encoding]::Unicode.GetString($ostPath, 0, $ostPath.Length - 2)
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
        $officeInfo = Get-OfficeInfo

        $offsets = @(
            for ($i = 0; $i -lt $count; ++$i) {
                if ($officeInfo.Architecture -eq 'x86') {
                    $reader.ReadInt32()
                }
                else {
                    $reader.ReadInt64()
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
            [ValidateSet('Mail', 'Calendar', 'Tasks', 'Advanced', 'Power', 'Security', 'Setup', 'Search')]
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
            $option.Path = $Property.PSPath | ConvertFrom-PSPath

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
        New-Option -Name 'NewMailDesktopAlerts' -Description 'Display a Desktop Alert' -Category Mail -Value $true
        New-Option -Name 'NewMailDesktopAlertsDRMPreview' -Description 'Enable preview for Rights Protected messages' -Category Mail -Value $false
        New-Option -Name 'SaveSent' -Description 'Save copies of messages in the Sent Items folder' -Category Mail -Value $true
        New-Option -Name 'DelegateSentItemsStyle' -Description "When set to 1, items sent on behalf of a manager will now go to the manager's sent items box" -Category Mail -Value $false
        New-Option -Name 'ShowLegacySharingUX' -Description 'Turn off Calendar Sharing REST API and use Legacy UI' -Category Calendar -Value $false
        New-Option -Name 'OpenTasksWithToDoApp' -Description 'When opening from a reminder, open tasks with ToDo App' -Category Tasks -Value $false
        New-Option -Name 'Autodetect_CodePageOut' -Description 'Automatically select encoding for outgoing messages' -Category Advanced -Value $true
        New-Option -Name 'Default_CodePageOut' -Description 'Preferred encoding for outgoing messages' -Category Advanced -Value $null
        New-Option -Name 'HighCostMeteredNetworkBehavior' -Description 'Behavior on a high cost metered network' -Category Power -Value 'Default'
        New-Option -Name 'ConservativeMeteredNetworkBehavior' -Description 'Behavior on a conservative metered network' -Category Power -Value 'Default'
        New-Option -Name 'BatteryMode' -Description 'Battery mode' -Category Power -Value 'Default'
        New-Option -Name 'MarkInternalAsUnsafe' -Description 'Use Protected View for attachments received from internal senders' -Category Security -Value $false
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
        }
    }

    & {
        Join-Path $optionsPath 'Calendar'
        Join-Path $optionsPath 'Calendar' | ConvertTo-PolicyPath
    } | Get-ItemProperty -ErrorAction SilentlyContinue | & {
        process {
            $PSDefaultParameterValues['Set-Option:Property'] = $_
            Set-Option -Name 'ShowLegacySharingUX'
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
                param (
                    $regValue,
                    $regName
                )

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
                param (
                    $regValue
                )

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

    Get-CachedAutodiscoverLocation -User $User | & {
        # Copy Autodiscover XML files to Path
        param ([Parameter(ValueFromPipeline)]$cachePath)
        process {
            Write-Log "Searching $($cachePath.Name) $($cachePath.Path)"

            # Use recurse only for the path under LOCALAPPDATA.
            try {
                Get-ChildItem $cachePath.Path -Filter '*Autod*.xml' -Force -Recurse:$($cachePath.Name -eq 'UnderLocalAppData') -ErrorAction SilentlyContinue `
                | Copy-Item -Destination $Path -PassThru `
                | Remove-HiddenAttribute
            }
            catch {
                # Just in case Copy-Item throws a terminating error.
                Write-Error -ErrorRecord $_
            }
        }
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
    $sourcePath = Get-UserShellFolder -User $User -ShellFolderName $LocalAppData | Join-Path -ChildPath 'Microsoft\Outlook\'

    if (-not $sourcePath) {
        Write-Error "Cannot find $LocalAppData for $User."
        return
    }

    if (-not (Test-Path $Path)) {
        $null = New-Item $Path -ItemType Directory -ErrorAction Stop
    }

    $Path = Resolve-Path $Path

    # Get OutlookConfig & OPX json files and copy them to Path
    try {
        Get-ChildItem $sourcePath -Filter '*Config*.json' -Force -Recurse -ErrorAction SilentlyContinue `
        | Copy-Item -Destination $Path -PassThru `
        | Remove-HiddenAttribute
    }
    catch {
        # Just in case Copy-Item throws a terminating error.
        Write-Error -ErrorRecord $_
    }
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
        Write-Error "Cannot find LocalAppData for $User."
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
    | Get-ChildItem `
    | Remove-Item -Recurse -Force

    $localAppData = Get-UserShellFolder -User $User -ShellFolderName 'Local AppData'

    if (-not $localAppData) {
        return
    }

    Stop-Service 'TokenBroker'

    # Remove identity and token cache
    & {
        'Microsoft\OneAuth\accounts'

        # IdentityCache
        'Microsoft\IdentityCache'

        # Accounts
        'Packages\Microsoft.AAD.BrokerPlugin_cw5n1h2txyewy\ac\TokenBroker\Accounts'
        'Packages\Microsoft.Windows.CloudExperienceHost_cw5n1h2txyewy\AC\TokenBroker\Accounts'
    } `
    | Join-Path $localAppData -ChildPath { $_ } `
    | Where-Object { Test-Path $_ } `
    | Remove-Item -Recurse -Force

    # Restart TokenBroker service
    Write-Verbose "Restarting TokenBroker service"
    Start-Service 'TokenBroker'
}

function Start-LdapTrace {
    [CmdletBinding(PositionalBinding = $false)]
    param(
        [Parameter(Mandatory = $true, Position = 0, HelpMessage = "Directory for output file")]
        [string]$Path,
        [Parameter(Mandatory = $true, HelpMessage = "Process name to trace. e.g. Outlook.exe")]
        [string]$TargetProcess,
        [string]$FileName = 'ldap.etl',
        [string]$SessionName = 'LdapTrace',
        [ValidateSet('NewFile', 'Circular')]
        [string]$LogFileMode = 'NewFile',
        [ValidateRange(1, [int]::MaxValue)]
        [int]$MaxFileSizeMB = 256
    )

    if (-not (Test-Path $Path)) {
        $null = New-Item $Path -ItemType directory -ErrorAction Stop
    }

    $Path = Resolve-Path $Path

    # Process name must contain the extension such as "Outlook.exe", instead of "Outlook"
    if ([IO.Path]::GetExtension($TargetProcess) -ne 'exe') {
        $TargetProcess = [IO.Path]::GetFileNameWithoutExtension($TargetProcess) + ".exe"
    }

    # Create a registry key under HKLM\SYSTEM\CurrentControlSet\Services\ldap\tracing
    $keypath = "HKLM:\SYSTEM\CurrentControlSet\Services\ldap\tracing"
    if (-not (Test-Path $keypath)) {
        $null = New-Item (Split-Path $keypath) -Name 'tracing' -ErrorAction SilentlyContinue
    }

    $null = New-Item $keypath -Name $TargetProcess -ErrorAction SilentlyContinue
    $key = Get-Item (Join-Path $keypath -ChildPath $TargetProcess)

    if (!$key) {
        Write-Error "Failed to create the key under $keypath. Make sure to run as an administrator"
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
    Write-Log "Starting a LDAP trace"
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
        [Parameter(Mandatory = $true)]
        $TargetProcess
    )

    Write-Log "Stopping $SessionName"
    $null = Stop-EtwSession $SessionName

    # Remove a registry key under HKLM\SYSTEM\CurrentControlSet\Services\ldap\tracing (ignore any errors)

    # Process name must contain the extension such as "Outlook.exe", instead of "Outlook"
    if ([IO.Path]::GetExtension($TargetProcess) -ne 'exe') {
        $TargetProcess = [IO.Path]::GetFileNameWithoutExtension($TargetProcess) + ".exe"
    }

    $keypath = "HKLM:\SYSTEM\CurrentControlSet\Services\ldap\tracing\$TargetProcess"
    $null = Remove-Item $keypath -ErrorAction SilentlyContinue
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

            [PSCustomObject]@{
                Name        = $file.Name
                FullName    = $file.FullName
                #VersionInfo = $item.VersionInfo # too much info and FileVersionRaw is harder to find
                FileVersion = $fileVersion
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

    $Path = Resolve-Path $Path

    # If MS Office is not installed, bail.
    $officeInfo = Get-OfficeInfo -ErrorAction SilentlyContinue
    if (-not $officeInfo) {
        Write-Error "MS Office is not installed."
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
        Write-Log "Waiting for the job (ID:$($job.Id)) up to $TimeoutSecond seconds."
        if (Wait-Job -Job $job -Timeout $TimeoutSecond) {
            # https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.core/wait-job
            # > This cmdlet returns job objects that represent the completed jobs. If the wait ends because the value of the Timeout parameter is exceeded, Wait-Job does not return any objects.
            Write-Log "Job was completed."
        }
        else {
            Write-Log "Job did not complete. It will be stopped by event signal."
        }

        # Signal the event and close
        try {
            $null = $namedEvent.Set()
            $namedEvent.Close()
            Write-Log "Event (Handle:$($namedEvent.Handle)) was closed."
        }
        catch {
            Write-Error -ErrorRecord $_
        }

        # Let the job finish
        $null = Wait-Job -Job $job
        Stop-Job -Job $job
        # Receive-Job -Job $job
        Remove-Job -Job $job
        Write-Log "Job (ID:$($job.Id)) was removed."
    }
}

function Start-CapiTrace {
    [CmdletBinding(PositionalBinding = $false)]
    param(
        [Parameter(Mandatory = $true, Position = 0)]
        [string]$Path,
        [string]$FileName = 'capi.etl',
        [string]$SessionName = 'CapiTrace',
        [ValidateSet('NewFile', 'Circular')]
        [string]$LogFileMode = 'NewFile',
        [ValidateRange(1, [int]::MaxValue)]
        [int]$MaxFileSizeMB = 256
    )

    if (-not (Test-Path $Path)) {
        $null = New-Item $Path -ItemType directory -ErrorAction Stop
    }
    $Path = Resolve-Path $Path

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

    $Path = Resolve-Path $Path
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
                Write-Error "It seems there is no connectivity to Internet. Please download FiddlerCapSetup.exe from `"$fiddlerCapUrl`" and place it `"$Path`". Then run again."
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
                    $ErrorActionPreference = "Continue"
                    # Do not double-quote $fiddlerPath here like /D=`"$fiddlerPath`". FiddlerSetupCap.exe doesn't like it for some reason. It's ok to have spaces in the path.
                    Start-Process $fiddlerSetupFile -ArgumentList "/S /D=$fiddlerPath" -Wait -PassThru
                }) 2>&1

            if ($process.ExitCode -ne 0) {
                Write-Error "Failed to extract $fiddlerExe. $(if ($process.ExitCode) {"exit code = $($process.ExitCode)."}) $err"
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
                $ErrorActionPreference = "Continue"
                try {
                    Start-Process $fiddlerExe -PassThru
                }
                catch {
                    Write-Error -ErrorRecord $_
                }
            }) 2>&1

        if (-not $process -or $process.HasExited) {
            Write-Error "FiddlerCap failed to start or prematurely exited. $(if ($null -ne $process.ExitCode) {"exit code = $($process.ExitCode)."}) $err"
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

function Start-Procmon {
    [CmdletBinding()]
    param(
        [parameter(Mandatory = $true)]
        $Path,
        $PmlFileName = "Procmon.pml",
        $ProcmonSearchPath # Look for existing procmon.exe before downloading
    )

    # Explicitly check admin rights
    if (-not (Test-RunAsAdministrator)) {
        Write-Warning "Please run as administrator"
        return
    }

    if (-not (Test-Path $Path)) {
        $null = New-Item $Path -ItemType Directory -ErrorAction Stop
    }

    $Path = Resolve-Path $Path
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
            Write-Log "ProcmonSearchPath '$ProcmonSearchPath' is provided, but coulnd't find Procmon.exe or Procmon64.exe."
        }
    }

    $procmonZipDownloaded = $false

    if ($procmonFile -and (Test-Path $procmonFile)) {
        Write-Log "$procmonFile is found. Skip searching & downloading ProcessMonitor.zip."
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
            Write-Log "$procmonZipFile is found. Skip downloading."
        }
        else {
            # If 'ProcessMonitor.zip' isn't there, download it.
            # If it's not connected to internet, bail.
            $connectivity = Get-NLMConnectivity
            if (-not $connectivity.IsConnectedToInternet) {
                Write-Error "It seems there is no connectivity to Internet. Please download the ProcessMonitor from `"$procmonDownloadUrl`""
                return
            }

            Write-Log "Downloading procmon"
            Write-Progress -Activity "Downloading procmon from $procmonDownloadUrl" -Status "Please wait" -PercentComplete -1
            $webClient = $null
            try {
                $webClient = New-Object System.Net.WebClient
                $webClient.UseDefaultCredentials = $true
                $webClient.DownloadFile($procmonDownloadUrl, $procmonZipFile)
                $procmonZipDownloaded = $true
            }
            catch {
                Write-Error -Message "Failed to download procmon from $procmonDownloadUrl. $_" -Exception $_.Exception
                return
            }
            finally {
                if ($webClient) {
                    $webClient.Dispose()
                }

                Write-Progress -Activity "Downloading procmon from $procmonDownloadUrl" -Status "Done" -Completed
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
        Write-Error "Failed to find $procmonFile."
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
            $ErrorActionPreference = "Continue"
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
            Write-Error "procmon failed to start or prematurely exited. $(if ($process.ExitCode) {"exit code = $($process.ExitCode)."}) $err"
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
                $ErrorActionPreference = "Continue"
                try {
                    Start-Process $procmonFile -ArgumentList "/Terminate" -Wait -PassThru
                }
                catch {
                    Write-Error -ErrorRecord $_
                }
            }) 2>&1

        if ($process.ExitCode -ne 0) {
            Write-Error "procmon failed to stop. $(if ($process.ExitCode) {"exit code = $($process.ExitCode)."}) $err"
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

    Write-Log "Using $keypath."

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

    if (-not (Test-Path $Path)) {
        $null = New-Item $Path -ItemType Directory -ErrorAction Stop
    }
    $Path = Resolve-Path $Path

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
    foreach ($item in @(Get-ChildItem -Path "$env:TEMP\*" -Include "office.log", "*.exe.log")) {
        try {
            Copy-Item $item -Destination $Path
        }
        catch {
            Write-Error -ErrorRecord $_
        }
    }
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
        Write-Error "tttracer.exe is not available."
        return
    }

    # Make sure $Executable exists.
    if (-not (Test-Path $Executable)) {
        Write-Error "Cannot find $Executable."
        return
    }

    if (-not (Test-Path $Path)) {
        $null = New-Item $Path -ItemType Directory -ErrorAction Stop
    }

    $stdout = Join-Path $Path 'stdout.txt'
    $stderr = Join-Path $Path 'stderr.txt'

    if ($OnLaunch) {
        Write-Log "TTD monitoring $Executable."
        # trace file name must include a wildcard ("%") for OnLaunch recording
        $outPath = Join-Path $Path "$([IO.Path]::GetFileNameWithoutExtension($Executable))_$(Get-Date -Format "yyyyMMdd_HHmmss")_%.run"
        $process = Start-Process $tttracer -ArgumentList "-out `"$outPath`"", "-onLaunch `"$Executable`"", "-parent *" -PassThru -WindowStyle Hidden -RedirectStandardOutput $stdout -RedirectStandardError $stderr
    }
    else {
        Write-Log "TTD launching $Executable."
        $outPath = Join-Path $Path "$([IO.Path]::GetFileNameWithoutExtension($Executable))_$(Get-Date -Format "yyyyMMdd_HHmmss").run"
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
            Write-Error "Cannot find the new instance of $targetName."
            return
        }

        Write-Log "Target process $($targetProcess.Name) (PID:$($targetProcess.Id)) has started."

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
        Write-Error "tttracer.exe is not available."
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
        Write-Log "Killing tttracer (PID:$($tttracerProcess.Id)) running in OnLaunch mode."
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
        Write-Error "tttracer.exe is not available."
        return
    }

    if ($targetProcess = Get-Process -Id $ProcessID -ErrorAction SilentlyContinue) {
        $targetName = $targetProcess.Name
    }
    else {
        Write-Error "Cannot find a process with PID $ProcessID."
        return
    }

    if (-not (Test-Path $Path)) {
        $null = New-Item $Path -ItemType Directory -ErrorAction Stop
    }

    # Form the output file name.
    $outPath = Join-Path $Path "$($targetName)_$(Get-Date -Format "yyyyMMdd_HHmmss").run"

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

    $Path = Resolve-Path $Path

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
        # Invoke-WebRequest throws a terminating error if host name cannot be resolved.
        try {
            $err = $(Invoke-Command -ScriptBlock {
                    $ProgressPreference = "SilentlyContinue";
                    Invoke-WebRequest -Uri $downloadUrl -OutFile $appInstallerPath
                }) 2>&1

            Write-Log "TTD.appinstaller is successfully downloaded from $downloadUrl"
        }
        catch {
            $err = $_
        }

        if (-not (Test-Path $appInstallerPath)) {
            Write-Error -Message "Failed to download TTD.appinstaller from $downloadUrl. $err"
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

    try {
        $err = $(Invoke-Command -ScriptBlock {
                $ProgressPreference = "SilentlyContinue";
                Invoke-WebRequest -Uri $mainBundle.Uri -OutFile $msixBundlePath
            }) 2>&1

        Write-Log "$msixName is successfully downloaded from $($mainBundle.Uri)"
    }
    catch {
        $err = $_
    }

    if (-not (Test-Path $msixBundlePath)) {
        Write-Error "Failed to download $mainBundle.Uri. $err"
        return
    }

    [PSCustomObject]@{
        MsixBundlePath = $msixBundlePath
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
        # Output folder path
        $Path,
        # Name of executable (such as outlook.exe)
        [Parameter(Mandatory)]
        $ExecutableName,
        [Alias('CmdLineFilter')]
        [string]$CommandlineFilter,
        [switch]$ShowUI
    )

    if (-not (Test-Path $Path)) {
        $null = New-Item -Path $Path -ItemType Directory -ErrorAction Stop
    }

    $Path = Resolve-Path $Path

    # Make sure extension is ".exe"
    if (-not ([System.IO.Path]::GetExtension($ExecutableName))) {
        $ExecutableName = "$ExecutableName.exe"
    }

    $outPath = $Path.ToString()

    if ($outPath.IndexOf(' ') -gt 0) {
        $outPath = "`"$outPath`""
    }

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

        if (-not $ShowUI) {
            '-noUI'
        }
    )

    $stderr = Join-Path $Path 'stderr.txt'
    Write-Log "Invoking 'ttd.exe $($ttdArgs -join ' ')'"

    $process = Start-Process 'ttd.exe' -ArgumentList $ttdArgs -WindowStyle Hidden -RedirectStandardError $stderr -PassThru

    # Make sure ttd.exe started successfully
    Start-Sleep -Seconds 1

    if (-not $process -or $process.HasExited) {
        if ($process) {
            $process.Dispose()
        }

        if (Test-Path $stderr) {
            $errText = [IO.File]::ReadAllText($stderr)
        }

        Write-Error "ttd.exe failed to start. $errText"
        return
    }

    Write-Log "ttd.exe (PID:$($process.Id)) has successfully started"

    [PSCustomObject]@{
        Process       = $process
        StandardError = $stderr
    }
}

function Stop-TTDMonitor {
    [CmdletBinding(DefaultParameterSetName = 'InputObject')]
    param(
        [Parameter(ParameterSetName = 'InputObject', ValueFromPipeline)]
        # Output of Start-TTDMonitor
        $InputObject
    )

    if ($InputObject) {
        $ttdProcess = $InputObject.Process

        # Sanity check if the target process is ttd
        if ($ttdProcess.Name -ne 'TTD') {
            Write-Error "Given process is not TTD"
            return
        }
    }
    else {
        # Even when multiple proceses are traced simultaneouly, there should be only one ttd.exe
        $ttdProcess = Get-Process -Name 'ttd' -ErrorAction SilentlyContinue | Select-Object -First 1

        if (-not $ttdProcess) {
            Write-Error "Cannot find any running instance of ttd.exe"
            return
        }
    }

    # Stop current tracing, if any
    Write-Log "Invoking 'ttd.exe -stop all'"
    $null = & ttd.exe -stop all

    # Stop monitoring
    Write-Log "Stopping ttd.exe (PID:$($ttdProcess.Id))"
    Stop-Process -InputObject $ttdProcess
    $ttdProcess.WaitForExit()
    $ttdProcess.Dispose()

    if ($InputObject) {
        if ((Test-Path $InputObject.StandardError) -and -not (Get-Content $InputObject.StandardError)) {
            Remove-Item $InputObject.StandardError 2>&1 | Write-Log -Category Warning
        }
    }
}

function Cleanup-TTD {
    [CmdletBinding()]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseApprovedVerbs', '')]
    param(
    )

    # Clean up ProcLaunchMon
    Write-Log "Invoking 'ttd.exe -cleanup'"
    $null = & ttd.exe -cleanup

    if ($LASTEXITCODE -ne 0) {
        Write-Error "ttd.exe -cleanup failed with $LASTEXITCODE"
    }
}

function Attach-TTD {
    [CmdletBinding()]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseApprovedVerbs', '')]
    param(
        [Parameter(Mandatory)]
        # Output folder path
        $Path,
        [Parameter(Mandatory)]
        # Target Process ID
        [ValidateRange(4, [int]::MaxValue)]
        [int]$ProcessId,
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

    # Make sure ttd.exe is available.
    $ttd = Get-Command 'ttd.exe' -ErrorAction SilentlyContinue

    if (-not $ttd) {
        Write-Error "TTD is not available"
        return
    }

    if (-not (Test-Path $Path)) {
        $null = New-Item $Path -ItemType Directory -ErrorAction Stop
    }

    $Path = Resolve-Path $Path | Select-Object -ExpandProperty Path

    # If Path contains spaces, surround by double-quotes
    $outPath = $Path

    if ($outPath.IndexOf(' ') -gt 0) {
        $outPath = "`"$outPath`""
    }

    # Create a named event for onInitCompleteEvent parameter of ttd.exe
    $initCompleteEventName = [Guid]::NewGuid().ToString()
    $initCompleteEvent = New-Object System.Threading.EventWaitHandle -ArgumentList $false, ([System.Threading.EventResetMode]::ManualReset), $initCompleteEventName

    $ttdArgs = @(
        '-acceptEula'
        '-timestampFileName'
        '-out', $outPath
        '-attach', $ProcessId,
        '-onInitCompleteEvent', $initCompleteEventName

        if (-not $ShowUI) {
            '-noUI'
        }
    )

    $stderr = Join-Path $Path 'stderr.txt'
    Write-Log "Invoking ttd.exe $($ttdArgs -join ' ')"

    $process = Start-Process 'ttd.exe' -ArgumentList $ttdArgs -WindowStyle Hidden -RedirectStandardError $stderr -PassThru

    $attachStart = Get-Timestamp
    $checkInterval = [TimeSpan]::FromSeconds(1)

    # Check if ttd.exe successfully attached. And if so, wait until ttd.exe signals initCompleteEvent
    while ($true) {
        if (-not $process -or $process.HasExited) {
            $initCompleteEvent.Dispose()

            if ($process) {
                $process.Dispose()
            }

            if (Test-Path $stderr) {
                $errText = [IO.File]::ReadAllText($stderr)
            }

            Write-Error "ttd.exe failed to attach to the target (PID:$ProcessId). $errText"
            return
        }

        if ($initCompleteEvent.WaitOne($checkInterval)) {
            # initCompleteEvent is signaled.
            $initCompleteEvent.Dispose()
            break
        }

        Start-Sleep -Seconds 1
    }

    $attachElapsed = Get-Elapsed $attachStart
    Write-Log "TTD (PID:$($process.Id)) successfully attached the target (PID:$ProcessId). MainWindowTitle:$($process.MainWindowTitle); Attach Wait:$attachElapsed"

    [PSCustomObject]@{
        Process       = $process # ttd.exe process
        ProcessId     = $ProcessId # target process PID
        StandardError = $stderr
    }
}

function Detach-TTD {
    [CmdletBinding(DefaultParameterSetName = 'InputObject')]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseApprovedVerbs', '')]
    param(
        [Parameter(ParameterSetName = 'InputObject', Mandatory, ValueFromPipeline)]
        # Output of Attach-TTD
        $InputObject,
        # Target Process ID (Not ttd.exe's PID)
        [Parameter(ParameterSetName = 'ProcessId', Mandatory)]
        [int]$ProcessId
    )

    if ($InputObject) {
        $ProcessId = $InputObject.ProcessId
    }

    # The target process might not exist anymore (e.g., it might have crashed)
    $targetProcess = $null

    try {
        $targetProcess = Get-Process -Id $ProcessId -ErrorAction SilentlyContinue

        if (-not $targetProcess) {
            Write-Log "Target Process (PID:$ProcessId) does not exist anymore"
            return
        }

        Write-Log "Detaching TTD. Invoking 'ttd.exe -stop $ProcessId'"
        $null = & ttd.exe -stop $ProcessId

        if ($LASTEXITCODE -ne 0) {
            Write-Error "Failed to detach TTD from $ProcessId"
        }
    }
    finally {
        if ($targetProcess) {
            $targetProcess.Dispose()
        }

        if ($InputObject) {
            if ($InputObject.Process) {
                # Wait for ttd.exe to finish.
                $InputObject.Process.WaitForExit()
                $InputObject.Process.Dispose()
            }

            if ((Test-Path $InputObject.StandardError) -and -not (Get-Content $InputObject.StandardError)) {
                Remove-Item $InputObject.StandardError 2>&1 | Write-Log -Category Warning
            }
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
                Write-Error "32bit PowerShell 2.0 is running on 64bit OS. Please run 64bit PowerShell."
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
                            Version     = $subKey.GetValue('DisplayVersion')
                            Location    = $subKey.GetValue('InstallLocation')
                            DisplayName = $displayName
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

        $Path = (Resolve-Path $Path -ErrorAction Stop).Path

        # Create a key 'LocalDumps' under HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\Windows Error Reporting\LocalDumps, if it doesn't exist
        $werKey = 'HKLM:\SOFTWARE\Microsoft\Windows\Windows Error Reporting'

        if (-not (Test-Path (Join-Path $werKey 'LocalDumps'))) {
            $null = New-Item $werKey -Name 'LocalDumps' -ErrorAction Stop
        }

        $localDumpsKey = Join-Path $werKey 'LocalDumps'
        $beginBlockComplete = $true
    }

    process {
        # Bail if begin block failed.
        if (-not $Local:beginBlockComplete) {
            return
        }

        if ($TargetProcess) {
            if (-not $TargetProcess.EndsWith(".exe")) {
                Write-Log "TargetProcess '$TargetProcess' does not end with '.exe'. Adding '.exe'"
                $TargetProcess += '.exe'
            }

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

        Write-Log "Temporarily disabling dwwin."
        Disable-DWWin 2>&1 | Write-Log
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
            Write-Log "Cannot find $localDumpsKey."
            return
        }

        $beginBlockComplete = $true
    }

    process {
        if (-not $Local:beginBlockComplete) {
            return
        }

        if ($TargetProcess) {
            if (-not $TargetProcess.EndsWith(".exe")) {
                Write-Log "$TargetProcess does not end with '.exe'. Adding '.exe'"
                $TargetProcess += '.exe'
            }

            $targetKey = Join-Path $localDumpsKey $TargetProcess
            Write-Log "Removing $targetKey"
            Remove-Item $targetKey
        }
        else {
            Write-Log "Removing values of $localDumpsKey"
            Remove-ItemProperty $localDumpsKey -Name 'CustomDumpFlags', 'DumpType', 'DumpFolder'
        }

        if (-not (Test-Path (Join-Path $localDumpsKey '*'))) {
            Write-Log "Removing $localDumpsKey because it has no subkeys."
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

    $Path = Resolve-Path $Path

    if ($env:PROCESSOR_ARCHITEW6432) {
        $netshexe = Join-Path $env:SystemRoot 'SysNative\netsh.exe'
    }
    else {
        $netshexe = Join-Path $env:SystemRoot 'System32\netsh.exe'
    }

    # Dump some wfp show commands.
    $now = Get-Date -Format 'yyyyMMdd_HHmmss'

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
        param($Path, $Interval, $MaxDuration)

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
            $eventFilePath = Join-Path $Path "netevents_$($now.ToString('yyyyMMdd_HHmmss')).xml"
            $null = & $netshexe wfp show netevents file="$eventFilePath" timewindow=$($Interval.TotalSeconds)

            Start-Sleep -Seconds $Interval.TotalSeconds
        }
    } -ArgumentList $Path, $Interval, $MaxDuration

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
        Write-Error "Cannot find a process with PID $ProcessId."
        return
    }
    elseif (-not $process.SafeHandle) {
        # This scenario is possible for a system process.
        Write-Error "Cannot obtain the process SafeHandle of $($process.Name)."
        return
    }

    if (-not (Test-Path $Path)) {
        $null = New-Item $Path -ItemType Directory -ErrorAction Stop
    }
    $Path = Resolve-Path $Path

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
        $dumpFile = Join-Path $Path "$($process.Name)_PID$($ProcessId)_$(Get-Date -Format 'yyyy-MM-dd-HHmmss').dmp"
        $dumpFileStream = [System.IO.File]::Create($dumpFile)
        $writeDumpSuccess = $false

        $dumpType = [Win32.Dbghelp+MINIDUMP_TYPE]'MiniDumpWithTokenInformation, MiniDumpIgnoreInaccessibleMemory, MiniDumpWithThreadInfo, MiniDumpWithFullMemoryInfo, MiniDumpWithUnloadedModules, MiniDumpWithHandleData, MiniDumpWithFullMemory'

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
        [Threading.CancellationToken]$CancellationToken
    )

    if (-not ($process = Get-Process -Id $ProcessId -ErrorAction SilentlyContinue)) {
        Write-Error "Cannnot find a process with PID $ProcessId."
        return
    }

    $ERROR_TIMEOUT = 1460
    $savedDumpCount = 0
    $intervalSeconds = 1

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

            # Need to get the process object every time since MainWindowHandle can change during the life time of a process.
            if ($proc = Get-Process -Id $ProcessId -ErrorAction SilentlyContinue) {
                $hWnd = $proc.MainWindowHandle
                $proc.Dispose()

                # During start up and shut down, MainWindowHandle can be 0.
                if ($hWnd -eq 0) {
                    Start-Sleep -Seconds $intervalSeconds
                    continue
                }
            }
            else {
                Write-Error "Cannnot find a process with PID $ProcessId"
                return
            }

            $result = [IntPtr]::Zero
            if (-not ([Win32.User32]::SendMessageTimeoutW($hWnd, 0, [IntPtr]::Zero, [IntPtr]::Zero, 0, $Timeout.TotalMilliseconds, [ref]$result))) {
                $ec = [Runtime.InteropServices.Marshal]::GetLastWin32Error()

                # if error code is 0 or ERROR_TIMEOUT, timeout occurred.
                if ($ec -eq 0 -or $ec -eq $ERROR_TIMEOUT) {
                    Write-Log "Hung window detected with $name (PID:$ProcessId, hWnd:$hWnd). $($savedDumpCount+1)/$DumpCount" -Category Warning
                    $dumpResult = Save-Dump -Path $Path -ProcessId $ProcessId

                    if ($dumpResult) {
                        $savedDumpCount++
                        Write-Log "Dump file is saved as '$($dumpResult.DumpFile)'"
                    }

                    if ($savedDumpCount -ge $DumpCount) {
                        Write-Log "Dump count reached $DumpCount. Exiting"
                        return
                    }
                }
                else {
                    Write-Error ("SendMessageTimeoutW failed with 0x{0:x8}" -f $ec)
                }
            }

            Start-Sleep -Seconds $intervalSeconds
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

    # Copy "MSIPC" folder and its subfolders, except ".lock" files.
    if ($All) {
        try {
            if (-not (Test-Path $Path)) {
                $null = New-Item -ItemType Directory $Path -ErrorAction Stop
            }

            Join-Path $msipcPath '*' | Copy-Item -Exclude '*.lock' -Destination $Path -Recurse -Force
        }
        catch {
            Write-Error -ErrorRecord $_
        }

        return
    }

    # Just copy *.ipclog files
    $msipcLogPath = Join-Path $msipcPath '\Logs\*.ipclog'
    $destination = Join-Path $Path 'Logs'

    if (-not (Test-Path $msipcLogPath)) {
        Write-Log "There are no files matching $msipcLogPath"
        return
    }

    # Because copying only contents of a folder, need to create the destination folder first.
    if (-not (Test-Path $destination)) {
        $null = New-Item -ItemType Directory $destination -ErrorAction Stop
    }

    try {
        Copy-Item $msipcLogPath -Destination $destination
    }
    catch {
        Write-Error -ErrorRecord $_
    }
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
        Write-Error "Cannot find LocalAppData folder for User $User."
        return
    }

    $sourcePath = Join-Path $localAppdata -ChildPath 'Microsoft\Outlook\*'
    $fileNameFilter = 'PolicyNudge*'

    if (-not (Test-Path $sourcePath -Filter $fileNameFilter)) {
        Write-Log "There are no files matching '$fileNameFilter' in $(Split-Path $sourcePath)."
        return
    }

    if (-not (Test-Path $Path)) {
        $null = New-Item $Path -ItemType Directory -ErrorAction Stop
    }

    try {
        Copy-Item -Path $sourcePath -Filter $fileNameFilter -Destination $Path
    }
    catch {
        Write-Error -ErrorRecord $_
    }
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
        Write-Error "Cannot find LocalAppData folder for User $User."
        return
    }

    $sourcePath = Join-Path $localAppdata -ChildPath 'Microsoft\Office\CLP'

    if (-not (Test-Path $sourcePath)) {
        Write-Log "Cannot find $sourcePath."
        return
    }

    try {
        Copy-Item $sourcePath -Destination $Path -Recurse
    }
    catch {
        Write-Error -ErrorRecord $_
    }
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
        Write-Error "Cannot find LocalAppData folder for User $User."
        return
    }

    $sourcePath = Join-Path $localAppdata -ChildPath 'Microsoft\Office\DLP'

    if (-not (Test-Path $sourcePath)) {
        Write-Log "Cannot find $sourcePath."
        return
    }

    try {
        Copy-Item $sourcePath -Destination $Path -Recurse
    }
    catch {
        Write-Error -ErrorRecord $_
    }
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
        Write-Verbose "MSAL Loggin is enabled. Log file:$logFile"

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
            Write-Log "Skipping $url because it's invalid."
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
                $path = ConvertFrom-PSPath $clsidProp.PsPath
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
                    Write-Log "Skipping $($props.ProgId) because its LoadBehavior is null."
                    return
                }

                if ($cache.ContainsKey($props.ProgId)) {
                    Write-Log "Skipping $($props.ProgId) because it's already found."
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
    }
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
        Write-Log "$dsregcmd is not available."
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
        [string]$FileName = 'perf',
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
    $Path = Resolve-Path $Path

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
    )

    $configFile = Join-Path $Path "perf.config"
    # Note:Encoding must be Ascii here ('utf8' will write as UTF-8 with BOM, which does not work for logman)
    Set-Content -LiteralPath $configFile -Value $counters -Encoding Ascii

    $filePath = Join-Path $Path $FileName
    Write-Log "Staring PerfCounter. Mode:$LogFileMode, IntervalSecond:$IntervalSecond, MaxFileSizeMB:$MaxFileSizeMB, FilePath:$filePath"

    switch ($LogFileMode) {
        'NewFile' {
            #$stdout = & logman.exe create counter -n 'PerfCounter' -cf $configFile -si $IntervalSecond -max $MaxFileSizeMB -o $filePath -cnf 0 -ow -v 'nnnnnn' -f 'csv'
            $stdout = & logman.exe create counter -n 'PerfCounter' -cf $configFile -si $IntervalSecond -max $MaxFileSizeMB -o $filePath -ow -v 'mmddhhmm' -f 'bin' -cnf 0
            break
        }

        'Circular' {
            $stdout = & logman.exe create counter -n 'PerfCounter' -cf $configFile -si $IntervalSecond -max $MaxFileSizeMB -o $filePath -ow -v 'mmddhhmm' -f 'bincirc' # -cnf 0
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

    $outFileName = "Win32_Process_$(Get-Date -Format "yyyyMMdd_HHmmss").xml"

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

function Start-ProcessMonitoring {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        # Folder path to save process list
        [string]$Path,
        [Parameter(Mandatory = $true)]
        # Name of processes to monitor
        [string[]]$Name,
        [System.Threading.CancellationToken]$CancellationToken,
        [TimeSpan]$Interval = [TimeSpan]::FromSeconds(3)
    )

    $targets = $Name | & { process { @{ Name = $_; Ids = @{} } } }
    Write-Log "Start monitoring $($Name -join ', ')."

    while ($true) {
        $found = Get-Process -Name $Name -ErrorAction SilentlyContinue | & {
            process {
                # Update the found target's info.
                $name = $_.ProcessName
                $id = $_.Id
                $startTime = $_.StartTime
                $_.Dispose()

                foreach ($target in $targets) {
                    if ($name -like $target.Name -and (-not $target.Ids.ContainsKey($id) -or $target.Ids[$id] -ne $startTime) ) {
                        Write-Log "Found new $name with PID $id, StartTime $startTime."
                        $target.Ids.Add($id, $startTime)
                        $true
                        break
                    }
                }
            }
        }

        # $found can be a bool or an array of bool, and either case is ok.
        if ($found) {
            Save-Process -Path $Path -Name $Name
        }

        if ($CancellationToken.IsCancellationRequested) {
            Write-Log "Cancel request acknowledged"
            return
        }

        Start-Sleep -Seconds $Interval.TotalSeconds
    }
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
        # This is just for testing
        [Switch]$EnablePerfCheck
    )

    if (-not (Test-Path $Path)) {
        $null = New-Item $Path -ItemType Directory -ErrorAction Stop
    }

    $runAsAdmin = Test-RunAsAdministrator

    if (-not $runAsAdmin) {
        Write-Log "Not running as admin. IncludeUserName won't be used." -Category Warning
    }

    # PowerShell 4's Get-Process has -IncludeUserName parameter. Use it if available (must be Elevated); otherwise fall back to WMI.
    $includeUserNameAvailable = $runAsAdmin -and ((Get-Command -Name 'Get-Process').Parameters.ContainsKey('IncludeUserName'))

    # Using a Hash table is much faster than Where-Object.
    $hashSet = @{}

    while ($true) {
        $start = Get-Timestamp

        Get-CimInstance Win32_Process | & {
            param ([Parameter(ValueFromPipeline)]$win32Process)
            process {
                try {
                    $key = "$($win32Process.ProcessId),$($win32Process.CreationDate)"

                    if ($hashSet.ContainsKey($key)) {
                        return
                    }

                    # Check if process is elevated except for "System Idle Process" (PID 0) and "System" (PID 4))
                    $isElevated = $null

                    if ($win32Process.ProcessId -gt 4) {
                        try {
                            $err = $($isElevated = Test-ProcessElevated $win32Process.ProcessId) 2>&1
                        }
                        catch {
                            $err = $_
                        }

                        if ($err) {
                            if (-not $runAsAdmin -and $err.Exception.NativeErrorCode -eq 5) {
                                # If not running as admin, Test-ProcessElevated is expected to fail with Access Denied (5) for some processes.
                                # No need to log this error.
                            }
                            else {
                                $errMsg = "Test-ProcessElevated failed for $($win32Process.Name) (PID:$($win32Process.ProcessId))"

                                # Maybe the process is gone already. In this case, OpenProcess would fail with ERROR_INVALID_PARAMETER (87).
                                if ($proc = Get-Process -Id $win32Process.ProcessId -ErrorAction SilentlyContinue) {
                                    Write-Log $errMsg -ErrorRecord $err -Category Error
                                    $proc.Dispose()
                                }
                                else {
                                    Write-Log "$errMsg because the process has already exited" -ErrorRecord $err
                                }
                            }
                        }
                    }

                    $obj = @{
                        Name            = $win32Process.Name
                        Id              = $win32Process.ProcessId
                        CreationDate    = $win32Process.CreationDate
                        Path            = $win32Process.Path
                        Elevated        = $isElevated
                        CommandLine     = $win32Process.CommandLine
                        ParentProcessId = $win32Process.ParentProcessId
                    }

                    # For processes specified in NamePattern parameter, save its User & Environment Variables.
                    if ($win32Process.ProcessName -match $NamePattern) {
                        Write-Log "Found a new instance of $($win32Process.ProcessName) (PID:$($win32Process.ProcessId), Elevated:$isElevated)"

                        if ($includeUserNameAvailable) {
                            if ($proc = Get-Process -Id $win32Process.ProcessId -IncludeUserName -ErrorAction SilentlyContinue) {
                                $obj.Add('User', $proc.UserName)
                                $obj.Add('EnvironmentVariables', $proc.StartInfo.EnvironmentVariables)
                                $proc.Dispose()
                            }
                        }
                        else {
                            try {
                                # GetOwner() could fail if the process has exited. Not likely, but be defensive here.
                                $owner = Invoke-CimMethod -InputObject $win32Process -MethodName 'GetOwner'
                                $obj.Add('User', "$($owner.Domain)\$($owner.User)")
                            }
                            catch {
                                Write-Error "Invoke-CimMethod with GetOwner failed for $($win32Process.Name)" -Exception $_.Exception
                            }

                            if ($proc = Get-Process -Id $win32Process.ProcessId -ErrorAction SilentlyContinue) {
                                $obj.Add('EnvironmentVariables', $proc.StartInfo.EnvironmentVariables)
                                $proc.Dispose()
                            }
                        }
                    }

                    $hashSet.Add($key, [PSCustomObject]$obj)
                }
                finally {
                    $win32Process.Dispose()
                }
            }
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

    $hashSet.Values | Export-Clixml -Path (Join-Path $Path 'Win32_Process.xml')
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
        [System.Threading.EventWaitHandle]$IsStartedEvent,
        [bool]$Circular
    )

    if ($IsStartedEvent) {
        $null = $IsStartedEvent.Set()
    }

    while ($true) {
        $startResult = Start-PSR -Path $Path -FileName "PSR_$(Get-Date -f 'MMdd_HHmmss')"
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
                    $cutoff = [datetime]::Now.AddHours(-1)
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
        $Path,
        [Parameter(Mandatory)]
        $Name,
        [TimeSpan]$Timeout,
        [int]$DumpCount,
        $CancellationToken,
        [System.Threading.EventWaitHandle]$IsStartedEvent
    )

    if ($IsStartedEvent) {
        $null = $IsStartedEvent.Set()
    }

    # Key:Process Hash, Value:true/false for "need to log".
    $procCache = @{}

    # The target process may restart while being monitored. Keep monitoring until canceled (via hungDumpCts).
    while ($true) {
        # Wait for the target process ($Name) to come live.
        while ($true) {
            if ($CancellationToken.IsCancellationRequested) {
                return
            }

            # There could be multiple process instances.
            # Not really expected for Outlook, but monitor the one that started first, while skipping the ones that have been monitored already.
            $targetList = @(
                Get-Process -Name $Name -ErrorAction SilentlyContinue | & {
                    param(
                        [Parameter(ValueFromPipeline)]
                        [System.Diagnostics.Process]$process)
                    process {
                        if (-not $process.SafeHandle) {
                            Write-Log "Cannot obtain SafeHandle of $($process.Name) (PID:$($process.Id)). Skipping this instance"
                            return
                        }

                        $target = [PSCustomObject]@{
                            Id        = $process.Id
                            StartTime = [DateTime]::MinValue
                        }

                        # It's possible to get Access Denied on StartTime (If SafeHandle is available, StartTime should also be available, but be defensive just in case)
                        try {
                            $target.StartTime = $process.StartTime
                        }
                        catch {
                            Write-Error -Message "Failed to get StartTime of $($process.Name) (PID:$($process.Id))" -Exception $_.Exception
                        }

                        $process.Dispose()

                        # Do not use GetHashCode() for the Process itself because it returns a new value every time.
                        $hash = $target.Id.GetHashCode() -bxor $target.StartTime.GetHashCode()

                        if ($procCache.ContainsKey($hash)) {
                            if ($procCache[$hash]) {
                                Write-Log "This instance of $Name (PID:$($target.Id), StartTime:$($target.StartTime)) has been seen already. This instance will be not be monitored"
                                $procCache[$hash] = $false
                            }
                        }
                        else {
                            $procCache.Add($hash, $true)
                            $target
                        }
                    }
                }
            )

            # If there are multiple processes available, pick the one that started earliest.
            $targetPid = $targetList | Sort-Object StartTime | Select-Object -First 1 -ExpandProperty Id

            if ($targetList.Count -gt 1) {
                Write-Log "There are $($targetList.Count) instances of $Name found"
            }

            if ($targetPid) {
                break
            }

            Start-Sleep -Seconds 2
        }

        Write-Log "Found $Name (PID:$targetPid). Starting hung window monitoring."

        $argsTable = @{
            Path      = $Path
            ProcessId = $targetPid
        }

        if ($DumpCount) {
            $argsTable.DumpCount = $DumpCount
        }

        if ($Timeout) {
            $argsTable.Timeout = $Timeout
        }

        if ($CancellationToken) {
            $argsTable.CancellationToken = $CancellationToken
        }

        Save-HungDump @argsTable 2>&1 | Write-Log -Category Error -PassThru
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
        $message = "Skipped autoupdate because OutlookTrace seems be installed as a module."
    }
    elseif (-not (Get-NLMConnectivity).IsConnectedToInternet) {
        $message = "Skipped autoupdate because there's no connectivity to internet."
    }
    else {
        try {
            Write-Progress -Activity "AutoUpdate" -Status 'Checking if a newer version is available. Please wait' -PercentComplete -1
            $release = Invoke-RestMethod -Uri $GitHubUri -UseDefaultCredentials -ErrorAction Stop

            if ($Version -ge $release.name) {
                $message = "Skipped because the current script ($Version) is newer than or equal to GitHub's latest release ($($release.name))."
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

                Write-Verbose "Lastest script ($($release.name)) was successfully downloaded."
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
        [string]$Path
    )

    # wpr is available on Win10 and above
    if (-not (Get-Command 'wpr.exe' -ErrorAction SilentlyContinue)) {
        Write-Error "WPR is not available on this machine."
        return
    }

    if ($PSBoundParameters.ContainsKey('Path')) {
        if (-not (Test-Path $Path)) {
            $null = New-Item $Path -ItemType Directory -ErrorAction Stop
        }

        $Path = Resolve-Path $Path
    }

    Write-Log "Starting WPR trace"

    if ($Path) {
        # For some reason, if the path contains a space & is double-quoted & ends with a backslash, wpr fails with "Invalid temporary trace directory. Error code:0xc5586004"
        # Make sure to remove the last backslash.
        if ($Path.EndsWith('\')) {
            $Path = $Path.Remove($Path.Length - 1)
        }

        $errs = $($null = wpr.exe -start GeneralProfile -start CPU -start Network -filemode -RecordTempTo $Path) 2>&1
    }
    else {
        $errs = $($null = wpr.exe -start GeneralProfile -start CPU -start Network -filemode) 2>&1
    }

    $errorMsg = $errs | ForEach-Object { $msg = $_.Exception.Message.Trim(); if ($msg) { $msg } }

    if ($LASTEXITCODE -ne 0) {
        Write-Error "wpr failed to start. LASTEXITCODE:0x$('{0:x}' -f $LASTEXITCODE).`n$errorMsg"
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
        Write-Error "WPR is not available on this machine."
        return
    }

    if (-not (Test-Path $Path)) {
        $null = New-Item $Path -ItemType Directory -ErrorAction Stop
    }

    $Path = Resolve-Path $Path
    $filePath = Join-Path $Path $FileName

    Write-Log "Stopping WPR trace"

    # Here Start-Process is used to suppress progress bar written by wpr.exe.
    $process = Start-Process -FilePath 'wpr.exe' -ArgumentList "-stop `"$filePath`"", '-skipPdbGen' -WindowStyle Hidden -PassThru -Wait
    $exitCode = $process.ExitCode
    $process.Dispose()

    # If "Invalid command syntax", retry without -skipPdbGen because the option might not be available (e.g. W2019)
    if ($exitCode -eq 0xc5600602) {
        $process = Start-Process -FilePath 'wpr.exe' -ArgumentList "-stop `"$filePath`"" -WindowStyle Hidden -PassThru -Wait
        $exitCode = $process.ExitCode
        $process.Dispose()
    }

    if ($exitCode -ne 0) {
        Write-Error "wpr failed to stop. ExitCode:0x$('{0:x}' -f $exitCode)."
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
        default { Write-Error "Failed to get CLSID of DefaultIMApp '$defaultIMApp'."; return }
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
.NOTES
WinRT interop (AwaitAction, Await) is borrowed from:

    https://fleexlab.blogspot.com/2018/02/using-winrts-iasyncoperation-in.html
.LINK
WebAccount.SignOutAsync Method
https://docs.microsoft.com/en-us/uwp/api/windows.security.credentials.webaccount.signoutasync

#>
function Invoke-WamSignOut {
    [CmdletBinding(PositionalBinding = $false)]
    param(
        # ClientId
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

    function AwaitAction($WinRtAction) {
        # WindowsRuntimeSystemExtensions.AsTask() creates a Task from WinRT future.
        # https://devblogs.microsoft.com/dotnet/asynchronous-programming-for-windows-store-apps-net-is-up-to-the-task/
        $asTask = ([System.WindowsRuntimeSystemExtensions].GetMethods() | Where-Object { $_.Name -eq 'AsTask' -and $_.GetParameters().Count -eq 1 -and !$_.IsGenericMethod })[0]
        $netTask = $asTask.Invoke($null, @($WinRtAction))
        $null = $netTask.Wait(-1)
    }

    function Await($WinRtTask, $ResultType) {
        # https://fleexlab.blogspot.com/2018/02/using-winrts-iasyncoperation-in.html
        $asTaskGeneric = ([System.WindowsRuntimeSystemExtensions].GetMethods() | Where-Object { $_.Name -eq 'AsTask' -and $_.GetParameters().Count -eq 1 -and $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncOperation`1' })[0]
        $asTask = $asTaskGeneric.MakeGenericMethod($ResultType)
        $netTask = $asTask.Invoke($null, @($WinRtTask))
        $null = $netTask.Wait(-1)
        $netTask.Result
    }

    # https://docs.microsoft.com/en-us/uwp/api/windows.security.authentication.web.core.webauthenticationcoremanager.findaccountproviderasync?view=winrt-20348#Windows_Security_Authentication_Web_Core_WebAuthenticationCoreManager_FindAccountProviderAsync_System_String_
    $provider = Await ([Windows.Security.Authentication.Web.Core.WebAuthenticationCoreManager, Windows, ContentType = WindowsRuntime]::FindAccountProviderAsync('https://login.microsoft.com', 'organizations')) ([Windows.Security.Credentials.WebAccountProvider, Windows, ContentType = WindowsRuntime])

    # https://docs.microsoft.com/en-us/uwp/api/windows.security.authentication.web.core.webauthenticationcoremanager.findallaccountsasync?view=winrt-20348
    $accounts = Await ([Windows.Security.Authentication.Web.Core.WebAuthenticationCoreManager, Windows, ContentType = WindowsRuntime]::FindAllAccountsAsync($provider, $ClientId)) ([Windows.Security.Authentication.Web.Core.FindAllAccountsResult, Windows, ContentType = WindowsRuntime])

    $count = $accounts.Accounts | Measure-Object | Select-Object -ExpandProperty Count

    if ($count -eq 0) {
        Write-Log "No account found."
        return
    }

    Write-Log "$count account$(if ($count -gt 1) {'s'}) found."

    foreach ($account in $accounts.Accounts) {
        $accountId = "UserName:$($account.UserName), Id:$($account.Id)"
        $state = $account.State

        if ($state -ne [Windows.Security.Credentials.WebAccountState]::Connected) {
            Write-Log "Skipping $accountId because its State is $state"
            continue
        }

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
        AwaitAction ($account.SignOutAsync($ClientId))
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
        Write-Error "Cannot find $Path."
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
        Write-Log "Cannot find Office Identities."
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
        Write-Log "Found active identity $($activeIdentity.EmailAddress) based on active profile."
    }
    else {
        # If there is no active profile, then pick one with LiveId, OrgId, or ADAL
        $activeIdentity = $identities `
        | Where-Object { $_.SignedOut -ne 1 -and ($IdpMapping[$_.IdP] -eq 'LiveId' -or $IdpMapping[$_.IdP] -eq 'OrgId' -or $IdpMapping[$_.IdP] -eq 'ADAL') } `
        | Select-Object -First 1

        if ($activeIdentity) {
            Write-Log "Found active identity $($activeIdentity.EmailAddress) based on IdP $($IdpMapping[$activeIdentity.IdP])."
        }
        else {
            $activeIdentity = $identities | Where-Object { $_.SignedOut -ne 1 } | Select-Object -First 1

            if ($activeIdentity) {
                Write-Log "Found active identity $($activeIdentity.EmailAddress) based on not SignedOut."
            }
            else {
                Write-Log "There is no active identity."
            }
        }
    }

    foreach ($identity in $identities) {
        $connectedExperience = Get-ConnectedExperience $identity

        [PSCustomObject]@{
            Profile                    = $identity.PSChildName
            LastSwitchedTime           = $identity.LastSwitchedTime
            IsActive                   = $identity -eq $activeIdentity
            EmailAddress               = $identity.EmailAddress
            FriendlyName               = $identity.FriendlyName
            HomeTenantId               = $identity.HomeTenantId
            SigninName                 = $identity.SigninName
            IdP                        = $IdpMapping[$identity.IdP]
            # Persisted                  = $identity.Persisted -eq 1
            SignedOut                  = $identity.SignedOut -eq 1
            ConnectedExperienceEnabled = $connectedExperience.Enabled
        }
    }
}

# Get ConnectedExperience state from the Office's roaming settings.
function Get-ConnectedExperience {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        $Identity
    )

    if (-not $userRegRoot) {
        $userRegRoot = Get-UserRegistryRoot

        if (-not $userRegRoot) {
            return
        }
    }

    # Check Office's roaming settings.
    $profileName = $Identity.PSChildName
    $roamingSettingsPath = Join-Path $userRegRoot 'Software\Microsoft\Office\16.0\Common\Roaming\Identities' `
    | Join-Path -ChildPath $profileName | Join-Path -ChildPath 'Settings\1272\{00000000-0000-0000-0000-000000000000}'

    if (-not (Test-Path $roamingSettingsPath)) {
        Write-Log "Cannot find roaming settings for $profileName."
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
        Write-Log "There is no roaming data for $profileName."
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

    Join-Path $localAppdata 'Microsoft\OneAuth\accounts' `
    | Get-ChildItem -ErrorAction SilentlyContinue | & {
        process {
            try {
                Get-Content $_.FullName -Encoding UTF8 | ConvertFrom-Json
            }
            catch {
                Write-Error -Message "Failed to parse $($_.FullName)" -Exception $_.Exception
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

    if (-not (Test-Path $Path)) {
        $null = New-Item $Path -ItemType Directory -ErrorAction Stop
    }

    Join-Path $localAppdata 'Microsoft\OneAuth\accounts' `
    | Get-ChildItem -ErrorAction SilentlyContinue | & {
        process {
            try {
                Copy-Item -LiteralPath $_.FullName -Destination $Path
            }
            catch {
                Write-Error -Message "Failed to copy $($_.FullName)" -Exception $_.Exception
            }
        }
    }
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
        $domainZone = Join-Path $userRegRoot "Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Domains\$domainHint" `
        | Get-ChildItem `
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
                    Path  = $Property.PSPath | ConvertFrom-PSPath
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
    $webClient = $null

    if (-not (Test-Path $Path)) {
        $null = New-Item $Path -ItemType Directory -ErrorAction Stop
    }

    try {
        $webClient = New-Object System.Net.WebClient
        $webClient.UseDefaultCredentials = $true
        $webClient.DownloadFile($url, $zoomItZip)
    }
    catch {
        Write-Error -Message "Failed to download ZoomIt from $url" -Exception $_.Exception
    }
    finally {
        if ($webClient) {
            $webClient.Dispose()
        }
    }

    if (-not (Test-Path $zoomItZip)) {
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

        # The first 4 bytes of IMAGE_FILE_HEADER is the machine architecture
        # https://learn.microsoft.com/en-us/windows/win32/api/winnt/ns-winnt-image_file_header
        $machine = $reader.ReadUInt16()

        $arch =
        switch ($machine) {
            0x014c { 'x86'; break }
            0x0200 { 'IA64'; break }
            0x8664 { 'x64'; break }
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
        $startProcArgs['RedirectStandardOutput'] = $filePath
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
    Look for StructuredQuerySchema.bin under %LOCALAPPDATA%\Microsoft\Windows.
    e.g. "C:\Users\admin\AppData\Local\Microsoft\Windows\1041\StructuredQuerySchema.bin"
#>
function Get-StructuredQuerySchema {
    [CmdletBinding(PositionalBinding = $false)]
    param(
        $User = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
    )

    # Look for StructuredQuerySchema.bin under %LOCALAPPDATA%
    # e.g. "C:\Users\admin\AppData\Local\Microsoft\Windows\1041\StructuredQuerySchema.bin"
    $localAppdata = Get-UserShellFolder -User $User -ShellFolderName 'Local AppData'

    if (-not $localAppdata) {
        return
    }

    Join-Path $localAppdata 'Microsoft\Windows' `
    | Get-ChildItem -Filter 'StructuredQuerySchema.bin' -Recurse `
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

                [PSCustomObject]@{
                    User       = $User
                    LocaleName = $culture.Name
                    LocaleId   = $lcid
                    Path       = $fileInfo.FullName
                }
            }
            else {
                Write-Error "Failed to extract LCID from '$($fileInfo.FullName)'"
            }
        }
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
            if ([Console]::KeyAvailable) {
                [ConsoleKeyInfo]$keyInfo = [Console]::ReadKey($true)

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

            Start-Sleep -Seconds 1
        }

        [Console]::TreatControlCAsInput = $false
        Write-Host
    }
    else {
        # Read-Host is not used here because it'd block background tasks.
        # When using UI.ReadLine(), Ctrl+C cannot be detected.
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
        [ValidateSet('Outlook', 'Netsh', 'PSR', 'LDAP', 'CAPI', 'Configuration', 'Fiddler', 'TCO', 'Dump', 'CrashDump', 'HungDump', 'Procmon', 'WAM', 'WFP', 'TTD', 'Performance', 'WPR', 'Recording')]
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
        # Timespan used to detect a hung window when "HungDump" is requested in Component.
        [ValidateRange('00:00:01', '00:01:00')]
        [TimeSpan]$HungTimeout = [TimeSpan]::FromSeconds(5),
        # Max number of hung dump files to be saved per process instance
        [ValidateRange(1, 10)]
        [int]$MaxHungDumpCount = 3,
        # Name of the target process to monitor a hung window (This is just for testing)
        [string]$HungMonitorTarget = 'Outlook',
        # Names of the target processes for crash dumps. When not specified, all processes will be the targets.
        [string[]]$CrashDumpTargets,
        # Switch to enable full page heap for Outlook.exe (With page heap, Outlook will consume a lot of memory and slow down)
        [switch]$EnablePageHeap,
        # Switch to add Microsoft.AAD.BrokerPlugin to Loopback Exempt
        [switch]$EnableLoopbackExempt,
        # Skip script version check.
        [switch]$SkipVersionCheck,
        # Command line filter for TTD monitor
        [string]$TTDCommandlineFilter,
        # Switch to show TTD UI
        [switch]$TTDShowUI,
        # Switch to remove cached identites & authentication tokens
        [switch]$RemoveIdentityCache
    )

    $runAsAdmin = Test-RunAsAdministrator

    # Explicitly check admin rights depending on the request.
    if (-not $runAsAdmin -and (($Component -join ' ') -match 'Outlook|Netsh|CAPI|LDAP|WAM|WPR|WFP|CrashDump|TTD' -or $EnablePageHeap -or $EnableLoopbackExempt -or $RemoveIdentityCache)) {
        Write-Warning "Please run as administrator"
        return
    }

    # Enable Debug Privilege if running as admin
    $debugPrivilegeEnabled = $false

    if ($runAsAdmin) {
        try {
            [System.Diagnostics.Process]::EnterDebugMode()
            $debugPrivilegeEnabled = $true
        }
        catch {
            Write-Error -Message "Cannot enable Debug Privilege. $_" -Exception $_.Exception
            return
        }
    }

    if ($env:PROCESSOR_ARCHITEW6432) {
        Write-Error "32-bit PowerShell is running on 64-bit OS. Please use 64-bit PowerShell from C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe"
        return
    }

    if (-not $SkipVersionCheck -and -not (Test-ScriptExpiration)) {
        Write-Error "This script is too old. The script version is $Version and it has passed $($Script:ValidTimeSpan.Days) days.`nPlease download the latest version from https://github.com/jpmessaging/OutlookTrace.`nYou can skip this check by using -SkipVersionCheck switch."
        return
    }

    # MS Office must be installed to collect Outlook, TCO, or TTD.
    # This is just a fail fast. Start-OutlookTrace/TCOTrace fail anyway.
    if ($Component -contains 'Outlook' -or $Component -contains 'TCO' -or $Component -contains 'TTD') {
        $err = $($null = Get-OfficeInfo -ErrorAction Continue) 2>&1

        if ($err) {
            Write-Error "Component `"Outlook`" and/or `"TCO`" is specified, but installation of Microsoft Office is not found. $err"
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
            Write-Error "Cannot find any logon user."
            return
        }
        elseif ($logonUsers.Count -eq 1) {
            $targetUser = $logonUsers[0]
        }
        else {
            # Multiple logon users are found. 'User' parameter needs to be used.
            Write-Error "Found multiple logon users ($($logonUsers.Count) users$(if ($logonUsers.Count -le 3) { "; $($logonUsers.Name -join ',')" })). Please specify the target user by `"-User`" parameter."
            return
        }
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
    $Path = Resolve-Path -LiteralPath $Path
    $tempPath = Join-Path $Path -ChildPath $([Guid]::NewGuid().ToString())
    $null = New-Item $tempPath -ItemType directory -ErrorAction Stop

    # Start logging.
    Open-Log -Path (Join-Path $tempPath 'Log.txt') -AutoFlush:$AutoFlush -ErrorAction Stop
    Write-Log "Script Version:$Script:Version (Module Version $($MyInvocation.MyCommand.Module.Version.ToString())); PID:$pid"
    Write-Log "PSVersion:$($PSVersionTable.PSVersion); CLRVersion:$($PSVersionTable.CLRVersion)"
    Write-Log "PROCESSOR_ARCHITECTURE:$env:PROCESSOR_ARCHITECTURE; PROCESSOR_ARCHITEW6432:$env:PROCESSOR_ARCHITEW6432"
    Write-Log "Running as $($currentUser.Name) ($($currentUser.Sid)); RunningAsAdmin:$runAsAdmin"
    Write-Log "Target user:$($targetUser.Name) ($($targetUser.Sid))"
    Write-Log "AutoUpdate:$(if ($SkipAutoUpdate) { 'Skipped due to SkipAutoUpdate switch' } else { $autoUpdate.Message })"

    $invocation = Get-CommandExpression -Invocation $MyInvocation
    Write-Log "Invocation:$invocation"

    $ScriptInfo = [PSCustomObject]@{
        Version    = $Script:Version
        Invocation = $invocation
        RunAsUser  = "$($currentUser.Name) ($($currentUser.Sid))"
        TargetUser = "$($targetUser.Name) ($($targetUser.Sid))"
        Start      = Get-Date
        WaitStart  = $null
        WaitStop   = $null
        End        = $null
    }

    try {
        # Set thread culture to en-US for consitent logging.
        Set-ThreadCulture 'en-US' 2>&1 | Write-Log -Category Error

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

        # Enable PageHeap for outlook.exe
        if ($EnablePageHeap) {
            $pageHeapEnabled = Enable-PageHeap -ProcessName 'outlook.exe' -ErrorAction Stop
        }

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
            $processCaptureTask = Start-Task -Name 'ProcessCaptureTask' -ScriptBlock { param($Path, $NamePattern, $CancellationToken) Start-ProcessCapture @PSBoundParameters } `
                -ArgumentList @{
                Path              = $OSDir
                NamePattern       = 'outlook|fiddler|explorer|backgroundTaskHost'
                CancellationToken = $processCaptureTaskCts.Token
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
            Invoke-ScriptBlock { param($User) Get-OutlookProfile @PSBoundParameters }
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
            $PSDefaultParameterValues.Remove('Invoke-ScriptBlock:ArgumentList')
            $PSDefaultParameterValues.Remove('Invoke-ScriptBlock:Path')

            Write-Progress -PercentComplete 80

            Invoke-ScriptBlock { param($User, $Path) Save-CachedAutodiscover @PSBoundParameters } -ArgumentList @{ User = $targetUser; Path = Join-Path $OfficeDir 'Cached AutoDiscover' }
            Invoke-ScriptBlock { param($User, $Path) Save-CachedOutlookConfig @PSBoundParameters } -ArgumentList @{ User = $targetUser; Path = Join-Path $OfficeDir 'Cached OutlookConfig' }
            Invoke-ScriptBlock { param($User, $Path) Save-PolicyNudge @PSBoundParameters } -ArgumentList @{ User = $targetUser; Path = Join-Path $OfficeDir 'PolicyNudge' }
            Invoke-ScriptBlock { param($User, $Path) Save-DLP @PSBoundParameters } -ArgumentList @{ User = $targetUser; Path = Join-Path $OfficeDir 'DLP' }
            Invoke-ScriptBlock { param($User, $Path) Save-CLP @PSBoundParameters } -ArgumentList @{ User = $targetUser; Path = Join-Path $OfficeDir 'CLP' }

            Write-Progress -Completed
        }

        # Check if Microsoft.AAD.BrokerPlugin is available.
        $($brokerPlugin = Get-AADBrokerPlugin) 2>&1 | Write-Log -Category Warning

        # Add Microsoft.AAD.BrokerPlugin to Loopback Exempt list if that's appropriate. If it is already added, Add-LoopbackExempt does nothing.
        $loopbackExemptAdded = $false

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
                $null = Start-FiddlerCap -Path $Path -ErrorAction Stop
                Write-Warning "FiddlerCap has started. Please manually configure and start capture."
            }
            else {
                # If target user is different from current user, don't start FiddlerCap because it won't be able to capture (WinInet proxy needs to be configured for the target user).
                $fiddler = Start-FiddlerCap -Path $Path -ErrorAction Stop -CheckAvailabilityOnly
                Write-Warning "Let the user ($($targetUser.Name)) start $($fiddler.FiddlerPath)."
            }

            $fiddlerCapStarted = $true
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
            Start-OutlookTrace -Path (Join-Path $tempPath 'Outlook') -ErrorAction Stop
            $outlookTraceStarted = $true
        }

        if ($Component -contains 'PSR') {
            Write-Progress -Status 'Starting PSR'

            # Make sure psr isn't running already.
            $psrProcesses = @(Get-Process psr -ErrorAction SilentlyContinue)

            if ($psrProcesses.Count -gt 0) {
                Write-Error "PSR is already running (PID:$($psrProcesses.ID -join ',')).`nPlease stop PSR first and run again."
                return
            }

            $psrCts = New-Object System.Threading.CancellationTokenSource
            $psrStartedEvent = New-Object System.Threading.EventWaitHandle($false, [Threading.EventResetMode]::ManualReset)
            Write-Log "Starting PSRTask. PsrRecycleInterval:$PsrRecycleInterval"

            $psrTask = Start-Task -Name 'PSRTask' -ScriptBlock { param($Path, $CancellationToken, $WaitInterval, $IsStartedEvent, $Circular) Start-PsrMonitor @PSBoundParameters } `
                -ArgumentList @{
                Path              = Join-Path $tempPath 'PSR'
                CancellationToken = $psrCts.Token
                WaitInterval      = $PsrRecycleInterval
                IsStartedEvent    = $psrStartedEvent
                Circular          = $LogFileMode -eq 'Circular'
            }

            $null = $psrStartedEvent.WaitOne([System.Threading.Timeout]::InfiniteTimeSpan)
            $psrStartedEvent.Dispose()
            $psrStarted = $true
        }

        if ($Component -contains 'LDAP') {
            Start-LDAPTrace -Path (Join-Path $tempPath 'LDAP') -TargetProcess 'Outlook.exe'
            $ldapTraceStarted = $true
        }

        if ($Component -contains 'CAPI') {
            Enable-EventLog 'Microsoft-Windows-CAPI2/Operational'
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
            Start-Wpr -Path (Join-Path $tempPath 'WPR') -ErrorAction Stop
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
                Write-Host "Hit enter to save a process dump of Outlook. To quit, enter q: " -NoNewline
                $userInput = $host.UI.ReadLine()

                if ($userInput.ToLower().Trim() -eq 'q') {
                    break
                }

                if (-not ($process = Get-Process -Name 'Outlook' -ErrorAction SilentlyContinue)) {
                    Write-Host "Cannot find Outlook.exe. Please start Outlook." -ForegroundColor Yellow
                    continue
                }

                Write-Progress -Activity "Saving a process dump of Outlook." -Status "Please wait." -PercentComplete -1
                $dumpResult = Save-Dump -Path (Join-Path $tempPath 'Dump') -ProcessId $process.Id
                Write-Progress -Activity "Saving a process dump of Outlook." -Status "Done" -Completed
                Write-Log "Saved a dump file:$($dumpResult.DumpFile)"
                $process.Dispose()
            }
        }

        if ($Component -contains 'HungDump') {
            Write-Progress -Status 'Starting HungMonitorTask'
            $hungDumpCts = New-Object System.Threading.CancellationTokenSource
            $monitorStartedEvent = New-Object System.Threading.EventWaitHandle($false, [Threading.EventResetMode]::ManualReset)
            Write-Log "Starting HungMonitorTask. HungMonitorTarget:$HungMonitorTarget, HungTimeout:$HungTimeout"

            $hungMonitorTask = Start-Task -Name 'HungMonitorTask' -ScriptBlock { param($Path, $Timeout, $DumpCount, $CancellationToken, $Name, $IsStartedEvent) Start-HungMonitor @PSBoundParameters } `
                -ArgumentList @{
                Path              = Join-Path $tempPath 'HungDump'
                Timeout           = $HungTimeout
                DumpCount         = $MaxHungDumpCount
                CancellationToken = $hungDumpCts.Token
                Name              = $HungMonitorTarget
                IsStartedEvent    = $monitorStartedEvent
            }

            $null = $monitorStartedEvent.WaitOne([System.Threading.Timeout]::InfiniteTimeSpan)
            $monitorStartedEvent.Dispose()
            $hungDumpStarted = $true
        }

        if ($Component -contains 'TTD') {
            Write-Progress -Status 'Starting TTD'

            # First make sure TTD.exe is available; Otherwise locate or download.
            if (-not (Get-Command 'ttd.exe' -ErrorAction SilentlyContinue)) {
                Download-TTD -Path $Path -ErrorAction Stop | Install-TTD
            }

            # Log TTD version
            if ($TTDPackage = Get-AppxPackage -Name 'Microsoft.TimeTravelDebugging') {
                Write-Log "Using Microsoft.TimeTravelDebugging version $($TTDPackage.Version)"
            }

            $ttdArgs = @{
                Path   = Join-Path $tempPath 'TTD'
                ShowUI = $TTDShowUI
            }

            # If Outlook is already running, attach to it. Otherwise, start monitoring for outlook.exe.
            if ($outlookProcess = Get-Process -Name 'Outlook' -ErrorAction SilentlyContinue) {
                Write-Log "Attaching TTD to Outlook (PID:$($outlookProcess.Id))."
                Write-Progress -Status "Attaching TTD to Outlook (PID:$($outlookProcess.Id)). This might take a while. Please wait"

                $ttdArgs.ProcessId = $outlookProcess.Id
                $ttdProcess = Attach-TTD @ttdArgs -ErrorAction Stop
            }
            else {
                $ttdArgs.ExecutableName = 'outlook.exe'
                $ttdArgs.CommandlineFilter = $TTDCommandlineFilter

                $ttdProcess = Start-TTDMonitor @ttdArgs -ErrorAction Stop
            }

            $ttdStarted = $true
        }

        if ($Component -contains 'Recording') {
            $recording = Start-Recording -ZoomItDownloadPath (Join-Path $Path 'ZoomIt') -ZoomItSearchPath $Path -ErrorAction Stop
            $recordingStarted = $true
        }

        Write-Progress -Completed
        $waitStart = Get-Timestamp

        if ($netshTraceStarted -or $outlookTraceStarted -or $psrStarted -or $ldapTraceStarted -or $capiTraceStarted -or $tcoTraceStarted -or $fiddlerCapStarted -or $crashDumpStarted -or $procmonStared -or $wamTraceStarted -or $wfpStarted -or $ttdStarted -or $perfStarted -or $hungDumpStarted -or $wprStarted -or $recordingStarted) {
            Write-Log "Waiting for the user to stop"
            $ScriptInfo.WaitStart = Get-Date

            Write-Host 'Hit enter to stop: ' -NoNewline
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
        $ScriptInfo.WaitStop = Get-Date

        $PSDefaultParameterValues['Write-Progress:Activity'] = 'Stopping traces'

        if ($recordingStarted) {
            # This will show the user a Save As dialog
            Stop-Recording
            Write-Host "Please save the recording (Save As dialog should appear). Then hit enter to continue:" -ForegroundColor Yellow -NoNewline
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

        if ($ttdStarted) {
            Write-Progress -Status 'Stopping TTD trace'

            # Having ProcessId means it's attached.
            if ($ttdProcess.ProcessId) {
                $ttdProcess | Detach-TTD 2>&1 | Write-Log -Category Error -PassThru
            }
            else {
                $ttdProcess | Stop-TTDMonitor 2>&1 | Write-Log -Category Error -PassThru
            }

            Uninstall-TTD
        }

        if ($netshTraceStarted) {
            Write-Progress -Status 'Stopping Netsh trace'
            Stop-NetshTrace 2>&1 | Write-Log -Category Error -PassThru
        }

        if ($outlookTraceStarted) {
            Write-Progress -Status 'Stopping Outlook trace'
            Stop-OutlookTrace 2>&1 | Write-Log -Category Error -PassThru
        }

        if ($ldapTraceStarted) {
            Write-Progress -Status 'Stopping LDAP trace'
            Stop-LDAPTrace -TargetProcess 'Outlook.exe'
        }

        if ($capiTraceStarted) {
            Write-Progress -Status 'Stopping CAPI trace'
            Disable-EventLog 'Microsoft-Windows-CAPI2/Operational'
            Stop-CAPITrace
        }

        if ($tcoTraceStarted) {
            Stop-TcoTrace -Path (Join-Path $tempPath 'TCO')
        }

        if ($wamTraceStarted) {
            Write-Progress -Status 'Stopping WAM trace'
            Disable-WamEventLog -ErrorAction SilentlyContinue
            Stop-WamTrace
        }

        if ($procmonStared) {
            Write-Progress -Status 'Stopping Procmon'
            Stop-Procmon 2>&1 | Write-Log -Category Error -PassThru
        }

        if ($wfpStarted) {
            Write-Progress -Status 'Stopping WFP trace'
            Stop-WfpTrace $wfpJob
        }

        if ($perfStarted) {
            Stop-PerfTrace 2>&1 | Write-Log -Category Error -PassThru
        }

        if ($wprStarted) {
            Write-Progress -Status 'Stopping WPR trace'
            Stop-Wpr -Path (Join-Path $tempPath 'WPR') | Write-Log -Category Error -PassThru
        }

        if ($hungDumpStarted) {
            $hungDumpCts.Cancel()
            Receive-Task $hungMonitorTask -AutoRemoveTask 2>&1 | Write-Log -Category Error
            $hungDumpCts.Dispose()
        }

        if ($crashDumpStarted) {
            $CrashDumpTargets | Remove-WerDumpKey
        }

        if ($fiddlerCapStarted) {
            Write-Warning "Please stop FiddlerCap and save the capture manually."
        }

        if ($psrStarted) {
            Write-Progress -Status "Stopping PSR"
            $psrCts.Cancel()
            Receive-Task $psrTask -AutoRemoveTask 2>&1 | Write-Log -Category Error -PassThru
            $psrCts.Dispose()
        }

        # Wait for the tasks started earlier and save the event logs
        if ($Component -contains 'Configuration') {
            if ($processCaptureTask) {
                Write-Progress -Status 'Stopping Process capture task'
                $processCaptureTaskCts.Cancel()
                $processCaptureTask | Receive-Task -AutoRemoveTask 2>&1 | Write-Log -Category Error
                $processCaptureTaskCts.Dispose()
                Write-Log "$($processCaptureTask.Name) is complete"
            }

            if ($startSuccess) {
                Write-Progress -Status 'Saving Event logs'
                Save-EventLog -Path $EventDir 2>&1 | Write-Log -Category Error

                Write-Progress -Status 'Saving MSIPC logs'
                Invoke-ScriptBlock { param($User, $Path, $All) Save-MSIPC @PSBoundParameters } -ArgumentList @{ User = $targetUser; Path = $MSIPCDir; All = $true }

                Invoke-ScriptBlock { param($User) Get-OneAuthAccount @PSBoundParameters } -ArgumentList @{ User = $targetUser } -Path $OfficeDir
                Invoke-ScriptBlock { param($Path, $User) Save-OneAuthAccount @PSBoundParameters } -ArgumentList @{ Path = Join-Path $OfficeDir 'OneAuthAccount'; User = $targetUser }
            }

            if ($osConfigurationTask) {
                Write-Progress -Status 'Saving OS configuration'

                if (-not $startSuccess) {
                    Write-Log "Canceling osConfigurationTask because startSuccess is false"
                    $osConfigurationTaskCts.Cancel()
                }

                $osConfigurationTask | Receive-Task -AutoRemoveTask 2>&1 | Write-Log -Category Error
                $osConfigurationTaskCts.Dispose()
                Write-Log "$($osConfigurationTask.Name) is complete"
            }

            if ($officeRegistryTask) {
                Write-Progress -Status 'Saving Office Registry'
                $officeRegistryTask | Receive-Task -AutoRemoveTask 2>&1 | Write-Log -Category Error
                Write-Log "$($officeRegistryTask.Name) is complete"
            }

            if ($networkInfoTask) {
                Write-Progress -Status 'Saving Network info'

                if (-not $startSuccess) {
                    Write-Log "Canceling networkInfoTask because startSuccess is false"
                    $networkInfoTaskCts.Cancel()
                }

                $networkInfoTask | Receive-Task -AutoRemoveTask 2>&1 | Write-Log -Category Error
                $networkInfoTaskCts.Dispose()
                Write-Log "$($networkInfoTask.Name) is complete"
            }

            if ($officeModuleInfoTask) {
                Write-Progress -Status "Saving Office module info"

                if ($startSuccess) {
                    $timeout = [TimeSpan]::FromSeconds(30)

                    if (Wait-Task $officeModuleInfoTask -Timeout $timeout) {
                        Write-Log "$($officeModuleInfoTask.Name) is complete before timeout."
                    }
                    else {
                        Write-Log "$($officeModuleInfoTask.Name) timed out after $($timeout.TotalSeconds) seconds. Task will be canceled." -Category Warning
                        $officeModuleInfoTaskCts.Cancel()
                    }
                }
                else {
                    Write-Log "Canceling $($officeModuleInfoTask.Name) because startSuccess is false"
                    $officeModuleInfoTaskCts.Cancel()
                }

                $officeModuleInfoTask | Receive-Task -AutoRemoveTask 2>&1 | Write-Log -Category Error
                $officeModuleInfoTaskCts.Dispose()
                Write-Log "$($officeRegistryTask.Name) is complete"
            }

            if ($gpresultTask) {
                Write-Progress -Status 'Saving Group Policy'

                if ($startSuccess) {
                    $timeout = [TimeSpan]::FromSeconds(30)

                    if (Wait-Task -Task $gpresultTask -Timeout $timeout) {
                        Write-Log "$($gpresultTask.Name) is complete before timeout"
                    }
                    else {
                        Write-Log "$($gpresultTask.Name) timed out after $($timeout.TotalSeconds) seconds. Task will be canceled." -Category Warning
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

        if ($pageHeapEnabled) {
            Disable-PageHeap -ProcessName 'outlook.exe' 2>&1 | Write-Log -Category Error -PassThru
        }

        if ($loopbackExemptAdded) {
            Remove-LoopbackExempt $brokerPlugin.PackageFamilyName 2>&1 | Write-Log -Category Error -PassThru
        }

        if ($debugPrivilegeEnabled) {
            try {
                [System.Diagnostics.Process]::LeaveDebugMode()
            }
            catch {
                Write-Log -Message "System.Diagnostics.Process.LeaveDebugMode failed" -ErrorRecord $_
            }
        }

        Write-Progress -Completed
        Reset-ThreadCulture 2>&1 | Write-Log -Category Error
        Close-TaskRunspace 2>&1 | Write-Log -Category Error
        Close-Log

        $ScriptInfo.End = Get-Date
        $ScriptInfo | Export-CliXml (Join-Path $tempPath 'ScriptInfo.xml')
    }

    # Bail if something failed or user interruped with Ctrl+C.
    if (-not $startSuccess) {
        return
    }

    $archiveName = "Outlook_$($env:COMPUTERNAME)_$(Get-Date -Format "yyyyMMdd_HHmmss")"

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

$Script:ValidTimeSpan = [TimeSpan]'120.00:00:00'

Export-ModuleMember -Function Test-ProcessElevated, Start-WamTrace, Stop-WamTrace, Start-OutlookTrace, Stop-OutlookTrace, Start-NetshTrace, Stop-NetshTrace, Start-PSR, Stop-PSR, Save-EventLog, Get-InstalledUpdate, Save-OfficeRegistry, Get-ProxySetting, Get-WinInetProxy, Get-WinHttpDefaultProxy, Get-ProxyAutoConfig, Save-OSConfiguration, Get-NLMConnectivity, Get-WSCAntivirus, Save-CachedAutodiscover, Remove-CachedAutodiscover, Save-CachedOutlookConfig, Remove-CachedOutlookConfig, Remove-IdentityCache, Start-LdapTrace, Stop-LdapTrace, Get-OfficeModuleInfo, Save-OfficeModuleInfo, Start-CAPITrace, Stop-CapiTrace, Start-FiddlerCap, Start-Procmon, Stop-Procmon, Start-TcoTrace, Stop-TcoTrace, Get-OfficeInfo, Add-WerDumpKey, Remove-WerDumpKey, Start-WfpTrace, Stop-WfpTrace, Save-Dump, Save-HungDump, Save-MSIPC, Get-EtwSession, Stop-EtwSession, Get-Token, Test-Autodiscover, Get-LogonUser, Get-JoinInformation, Get-OutlookProfile, Get-OutlookAddin, Get-ClickToRunConfiguration, Get-WebView2, Get-DeviceJoinStatus, Save-NetworkInfo, Download-TTD, Install-TTD, Uninstall-TTD, Start-TTDMonitor, Stop-TTDMonitor, Cleanup-TTD, Attach-TTD, Detach-TTD, Start-PerfTrace, Stop-PerfTrace, Start-Wpr, Stop-Wpr, Get-IMProvider, Get-MeteredNetworkCost, Save-PolicyNudge, Save-CLP, Save-DLP, Invoke-WamSignOut, Enable-PageHeap, Disable-PageHeap, Get-OfficeIdentityConfig, Get-OfficeIdentity, Get-OneAuthAccount, Remove-OneAuthAccount, Get-AlternateId, Get-UseOnlineContent, Get-AutodiscoverConfig, Get-SocialConnectorConfig, Get-ImageFileExecutionOptions, Start-Recording, Stop-Recording, Get-OutlookOption, Get-WordMailOption, Get-ImageInfo, Get-PresentationMode, Get-AnsiCodePage, Get-PrivacyPolicy, Save-GPResult, Get-AppContainerRegistryAcl, Collect-OutlookInfo