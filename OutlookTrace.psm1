<#
.NOTES
Copyright (c) 2021 Ryusuke Fujita

This software is released under the MIT License.
http://opensource.org/licenses/mit-license.php

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
#>

$Version = 'v2022-01-10'
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

        [DllImport("kernel32.dll")]
        public static extern bool IsWow64Process(SafeHandle hProcess, out bool wow64Process);

        [DllImport("kernel32.dll", ExactSpelling = true)]
        public static extern void RtlZeroMemory(IntPtr dst, int length);
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

    public static class User32
    {
        [DllImport("user32.dll", SetLastError=true, CharSet=CharSet.Auto)]
        public static extern uint SendMessageTimeoutW(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam, uint fuFlags, uint uTimeout, out IntPtr lpdwResult);
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

    if ($Script:logWriter) {
        Close-Log
    }

    # Open a file & add header
    try {
        [IO.StreamWriter]$Script:logWriter = [IO.File]::AppendText($Path)
        if ($AutoFlush) {
            $Script:logWriter.AutoFlush = $true
        }
        $Script:logWriter.WriteLine("date-time,thread_relative_delta(ms),thread,function,info")
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
        [System.Management.Automation.ErrorRecord]$ErrorRecord
    )

    process {
        # If ErrorRecord is provided, use it.
        if ($ErrorRecord) {
            $Message += "$(if ($Message) {'; '})[ErrorRecord]InvocationInfo.MyCommand: $($ErrorRecord.InvocationInfo.MyCommand), Exception.Message: $($ErrorRecord.Exception.Message), InvocationInfo.Line: '$($ErrorRecord.InvocationInfo.Line.Trim())', ScriptStackTrace: $($ErrorRecord.ScriptStackTrace.Replace([Environment]::NewLine, ' '))"
        }

        # Ignore null or an empty string.
        if (-not $Message) {
            return
        }

        # If Open-Log is not called beforehand, just output to verbose.
        if (-not $Script:logWriter) {
            Write-Verbose $Message
            return
        }

        $currentTime = Get-Date
        $currentTimeFormatted = $currentTime.ToString('o')

        # Delta time is relative to thread.
        # Each thread has it's own copy of lastLogTime now.
        [TimeSpan]$delta = 0;
        if ($Script:lastLogTime) {
            $delta = $currentTime.Subtract($Script:lastLogTime)
        }

        # Format as CSV:
        $sb = New-Object System.Text.StringBuilder
        $sb.Append($currentTimeFormatted).Append(',') | Out-Null
        $sb.Append($delta.TotalMilliseconds).Append(',') | Out-Null
        $sb.Append([System.Threading.Thread]::CurrentThread.ManagedThreadId).Append(',') | Out-Null
        $sb.Append((Get-PSCallStack)[1].Command).Append(',') | Out-Null
        $sb.Append('"').Append($Message.Replace('"', "'")).Append('"') | Out-Null

        # Protect from concurrent write
        [System.Threading.Monitor]::Enter($Script:logWriter)
        try {
            $Script:logWriter.WriteLine($sb.ToString())
        }
        finally {
            [System.Threading.Monitor]::Exit($Script:logWriter)
        }

        $sb = $null
        $Script:lastLogTime = $currentTime
    }
}

function Close-Log {
    if ($Script:logWriter) {
        Write-Log "Closing logWriter."
        $Script:logWriter.Close()
        $Script:logWriter = $null
        $Script:lastLogTime = $null
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

    if (-not $Script:runspacePool) {
        Write-Log "Setting up a Runspace Pool with an initialSessionState. MinRunspaces: $MinRunspaces, MaxRunspaces: $MaxRunspaces."
        $initialSessionState = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault()

        # Add functions from this script module. This will find all the functions including non-exported ones.
        # Note: I just want to call "ImportPSModule". It works, but emits "WARNING: The names of some imported commands ...".
        # Just to avoid this, I'm manually adding each command.
        #   $initialSessionState.ImportPSModule($MyInvocation.MyCommand.Module.Path)
        if ($MyInvocation.MyCommand.Module) {
            Get-Command -Module $MyInvocation.MyCommand.Module | ForEach-Object {
                $initialSessionState.Commands.Add($(
                        New-Object System.Management.Automation.Runspaces.SessionStateFunctionEntry($_.Name, $_.ScriptBlock)
                    ))
            }
        }

        # Import extra modules.
        if ($Modules) {
            $initialSessionState.ImportPSModule($Modules)
        }

        # Import Script-scoped variable.
        if ($IncludeScriptVariables) {
            foreach ($_ in @(Get-Variable -Scope Script | Where-Object { $_.Options -notmatch 'Constant' -and $_.Value })) {
                $initialSessionState.Variables.Add($(
                        New-Object System.Management.Automation.Runspaces.SessionStateVariableEntry -ArgumentList $_.Name, $_.Value, <# description #>$null
                    ))
            }
        }

        # Import given variables
        foreach ($_ in $Variables) {
            $initialSessionState.Variables.Add($(
                    New-Object System.Management.Automation.Runspaces.SessionStateVariableEntry -ArgumentList $_.Name, $_.Value, <# description #>$null
                ))
        }

        $Script:runspacePool = [runspacefactory]::CreateRunspacePool($MinRunspaces, $MaxRunspaces, $initialSessionState, $Host)
        $Script:runspacePool.Open()

        Write-Log "RunspacePool ($($Script:runspacePool.InstanceId.ToString())) is opened."
    }
}

function Close-TaskRunspace {
    [CmdletBinding()]
    param()

    if (-not $Script:runspacePool) {
        return
    }

    $id = $Script:runspacePool.InstanceId.ToString()
    $Script:runspacePool.Close()
    $Script:runspacePool = $null
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
    [CmdletBinding()]
    param (
        # Command to execute.
        [Parameter(ParameterSetName = 'Command', Mandatory = $true, Position = 0)]
        [string]$Command,
        # Parameters (name and value) to the command.
        [Parameter(ParameterSetName = 'Command')]
        $Parameters,
        # ScriptBlock to execute.
        [Parameter(ParameterSetName = 'Script', Mandatory = $true, Position = 0)]
        [ScriptBlock]$ScriptBlock,
        # ArgumentList to ScriptBlock
        [Parameter(ParameterSetName = 'Script')]
        [object[]]$ArgumentList
    )

    if (-not $Script:runspacePool) {
        Write-Error -Message "Open-TaskRunspace must be called in advance."
        return
    }

    # Create a PowerShell instance and set paramters if any.
    [PowerShell]$ps = [PowerShell]::Create()
    $ps.RunspacePool = $Script:runspacePool

    switch -Wildcard ($PSCmdlet.ParameterSetName) {
        'Command' {
            $ps.AddCommand($Command) | Out-Null
            foreach ($key in $Parameters.Keys) {
                $ps.AddParameter($key, $Parameters[$key]) | Out-Null
            }
            break
        }

        'Script' {
            $ps.AddScript($ScriptBlock) | Out-Null
            foreach ($p in $ArgumentList) {
                $ps.AddArgument($p) | Out-Null
            }
            break
        }
    }

    # Start the command
    $ar = $ps.BeginInvoke()

    [PSCustomObject]@{
        AsyncResult  = $ar
        PowerShell   = $ps
        # These are for diagnostic purpose
        ScriptBlock  = $ScriptBlock
        ArgumentList = $ArgumentList
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

            try {
                # To support Ctrl+C, wake up once in while.
                while ($true) {
                    if ($ar.AsyncWaitHandle.WaitOne(2000)) {
                        break
                    }
                }
                $ps.EndInvoke($ar)
            }
            catch {
                Write-Error -Message "Task threw a terminating error.`nScriptBlock: $($t.ScriptBlock); ArgumentList: $($t.ArgumentList)`n$_" -Exception $_.Exception
            }

            if ($ps.HadErrors) {
                $ps.Streams.Error | ForEach-Object {
                    # Include the ErrorRecord's InvocationInfo so that it's easier to understand the origin of error.
                    Write-Error -Message "Task has a non-terminating error.`nScriptBlock: $($t.ScriptBlock); ArgumentList: $($t.ArgumentList);`n$($_.InvocationInfo.MyCommand): $($_.Exception.Message);`n$($_.InvocationInfo.PositionMessage)" -Exception $_.Exception

                    if ($TaskErrorVariable) {
                        # Scope 1 is the parent scope, but it's not necessarily the caller scope.
                        # If the caller is a function in this module, then scope 1 is the caller function.
                        # However, if it's called from outside of module, scope 1 is the module's script scope. Thus the caller does not get the error.
                        # Because this function is meant to be moudule-internal and should be called only within the moudle, Scope 1 is ok for now.
                        New-Variable -Name $TaskErrorVariable -Value $($ps.Streams.Error.ReadAll()) -Scope 1 -Force

                        # To see if it's called from within this moudle, maybe I can check the SessionState.
                        # if ($SessionState -eq $ExecutionContext.SessionState) {
                        #     New-Variable -Name $TaskErrorVariable -Value $($ps.Streams.Error.ReadAll()) -Scope 1 -Force
                        # }
                        # else {
                        #     $SessionState.PSVariable.Set($TaskErrorVariable,$ps.Streams.Error.ReadAll())
                        # }
                    }
                }
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

        if (Test-Path $Destination) {
            $Destination = Resolve-Path -LiteralPath $Destination
        }
        else {
            $Destination = New-Item $Destination -ItemType Directory -ErrorAction Stop | Select-Object -ExpandProperty FullName
        }

        # Apply filters if any.
        if ($Filter.Count) {
            $files = @(foreach ($f in $Filter) { Get-ChildItem -LiteralPath $Path -Filter $f -Recurse -Force | Where-Object { -not $_.PSIsContainer } })
        }
        else {
            $files = @(Get-ChildItem -LiteralPath $Path -Recurse -Force | Where-Object { -not $_.PSIsContainer })
        }

        if ($PSBoundParameters.ContainsKey('FromDateTime') -and $FromDateTime -ne [DateTime]::MinValue) {
            $files = @($files | Where-Object { $_.LastWriteTime -ge $FromDateTime })
        }

        if ($PSBoundParameters.ContainsKey('ToDateTime') -and $ToDateTime -ne ([DateTime]::MaxValue)) {
            $files = @($files | Where-Object { $_.LastWriteTime -le $ToDateTime })
        }

        # Remove duplicate by Fullname
        $files = @($files | Group-Object -Property 'FullName' | ForEach-Object { $_.Group | Select-Object -First 1 })

        # If there are no files after filters are applied, bail.
        if ($files.Count -eq 0) {
            Write-Error "There are no files after filters are applied. Server: $env:COMPUTERNAME, Path: $Path, Filter: $Filter, FromDateTime: $FromDateTime, ToDateTime: $ToDateTime"
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
            New-Item $zipFilePath -ItemType file | Out-Null

            $zipStream = New-Object System.IO.FileStream -ArgumentList $zipFilePath, ([IO.FileMode]::Open)
            $zipArchive = New-Object System.IO.Compression.ZipArchive -ArgumentList $zipStream, ([IO.Compression.ZipArchiveMode]::Create)
            $count = 0
            $prevProgress = 0

            foreach ($file in $files) {
                $progress = 100 * $count / $files.Count
                if ($progress -ge $prevProgress + 10) {
                    Write-Progress -Activity "Creating a zip file $zipFilePath" -Status "Please wait" -PercentComplete $progress
                    $prevProgress = $progress
                }

                $fileStream = $zipEntryStream = $null
                try {
                    $fileStream = New-Object System.IO.FileStream -ArgumentList $file.FullName, ([IO.FileMode]::Open), ([IO.FileAccess]::Read), ([IO.FileShare]::ReadWrite)
                    $zipEntry = $zipArchive.CreateEntry($file.FullName.Substring($Path.TrimEnd('\').Length + 1))
                    $zipEntryStream = $zipEntry.Open()
                    $fileStream.CopyTo($zipEntryStream)

                    ++$count
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

            Write-Progress -Activity "Creating a zip file $zipFilePath" -Status "Done" -Completed
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
                Write-Error "There are no files after filters are applied. Server: $env:COMPUTERNAME, Path: $Path, Filter: $Filter, FromDateTime: $FromDateTime, ToDateTime: $ToDateTime"
                return
            }

            # Copy filtered files to a temporary folder
            $tempPath = Join-Path $env:TEMP ([IO.Path]::GetRandomFileName().Substring(0, 8))
            New-Item $tempPath -ItemType Directory | Out-Null

            foreach ($file in $files) {
                $dest = $tempPath
                $subPath = $file.DirectoryName.SubString($Path.Length)
                if ($subPath) {
                    $dest = Join-Path $tempPath $subPath
                    if (-not (Test-Path -LiteralPath $dest)) {
                        New-Item -ItemType Directory -Path $dest | Out-Null
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

        Write-Verbose "targetPath: $targetPath"

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
        Set-Content $archivePath ("PK" + [char]5 + [char]6 + ("$([char]0)" * 18))
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

        [System.Runtime.Interopservices.Marshal]::FinalReleaseComObject($shellApp) | Out-Null

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
            Write-Error "There are no files after filters are applied. Server: $env:COMPUTERNAME, Path: $Path, Filter: $Filter, FromDateTime: $FromDateTime, ToDateTime: $ToDateTime"
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
            Write-Error "MakeCab.exe failed; exitCode: $LASTEXITCODE; stdout:`"$stdout`"; Error: $err"
            return
        }

        New-Object PSCustomObject -Property @{
            ArchivePath = $cabFilePath
            # Message = $stdout
        }
    }

    # Here's main body of Compress-Folder
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
    $PSBoundParameters.Keys | ForEach-Object { if ($compressCmd.Parameters.ContainsKey($_)) { $params.Add($_, $PSBoundParameters[$_]) } }
    & $compressCmd @params
}

function Enable-EventLog {
    [CmdletBinding(PositionalBinding = $false)]
    param(
        [Parameter(Mandatory = $true, Position = 0)]
        [string]$EventName
    )

    $(wevtutil.exe set-log $EventName /enabled:true /retention:false /quiet:true | Out-Null) 2>&1 | ForEach-Object {
        Write-Error -ErrorRecord $_
    }
}

function Disable-EventLog {
    [CmdletBinding(PositionalBinding = $false)]
    param(
        [Parameter(Mandatory = $true, Position = 0)]
        [string]$EventName
    )

    $(wevtutil.exe set-log $EventName /enabled:false /retention:false /quiet:true | Out-Null) 2>&1 | ForEach-Object {
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
        New-Item $Path -ItemType Directory -ErrorAction Stop | Out-Null
    }
    $Path = Resolve-Path $Path

    # Create a provider listing
    $providerFile = Join-Path $Path -ChildPath 'wam.prov'
    Set-Content $WamProviders -Path $providerFile -ErrorAction Stop

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
    Stop-EtwSession $SessionName | Out-Null
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
        New-Item $Path -ItemType Directory -ErrorAction Stop | Out-Null
    }

    $Path = Resolve-Path $Path
    $providerFile = Join-Path $Path -ChildPath 'Office.prov'
    $officeInfo = Get-OfficeInfo -ErrorAction Stop
    $major = $officeInfo.Version.Split('.')[0] -as [int]
    Write-Log "Creating a provider listing according to the version $major"

    switch ($major) {
        14 { Set-Content $Outlook2010Providers -Path $providerFile -ErrorAction Stop; break }
        15 { Set-Content $Outlook2013Providers -Path $providerFile -ErrorAction Stop; break }
        16 { Set-Content $Outlook2016Providers -Path $providerFile -ErrorAction Stop; break }
        default { throw "Couldn't find the version from $_" }
    }

    # Configure log file mode, filename, and max file size if ncessary.
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
        Write-Log "Starting an Outlook trace. SessionName:`"$SessionName`", traceFile:`"$traceFile`", logFileMode:`"$mode`", maxFileSize: `"$MaxFileSizeMB`""

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
    Stop-EtwSession $SessionName | Out-Null
}

function Start-NetshTrace {
    [CmdletBinding()]
    param (
        [parameter(Mandatory = $true)]
        $Path,
        $FileName = 'nettrace-winhttp-webio.etl',
        [ValidateSet('None', 'Mini', 'Full')]
        $RerpotMode = 'None'
    )

    if (-not (Test-Path $Path)) {
        New-Item $Path -ItemType Directory -ErrorAction Stop | Out-Null
    }
    $Path = Resolve-Path $Path

    # Use "InternetClient_dbg" for Win10
    $win32os = Get-WmiObject win32_operatingsystem
    $osMajor = $win32os.Version.Split(".")[0] -as [int]
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
    & ipconfig /flushdns | Out-Null

    Write-Log "Starting a netsh trace."
    $traceFile = Join-Path $Path -ChildPath $FileName
    $err = $($stdout = Invoke-Command {
            $ErrorActionPreference = 'Continue'
            & $netshexe trace start scenario=$scenario capture=yes tracefile="`"$traceFile`"" overwrite=yes maxSize=2048 # correlation=yes
        }) 2>&1

    if ($err -or $LASTEXITCODE -ne 0) {
        Write-Error "netsh failed.`nexit code: $LASTEXITCODE; stdout: $stdout; error: $err"
        return
    }

    # Even with "report=no" (by default), "HKEY_CURRENT_USER\System\CurrentControlSet\Control\NetTrace\Session\MiniReportEnabled" might be set to 1.
    # (This depends on Win10 version with a scenario. For InternetClient_dbg scenario, Win10 2004 and above does not generate mini report).
    # In order to suppress generating a minireport (i.e. C:\Windows\System32\gatherNetworkInfo.vbs), set MiniReportEnabled to 0 before netsh trace stop.
    # * You could set "report=disabled", but if you want the mini report specifically (not Full report), you need to manually configure the registry value.
    $netshRegPath = 'HKCU:\System\CurrentControlSet\Control\NetTrace\Session\'
    switch ($RerpotMode) {
        'None' { Set-ItemProperty -Path $netshRegPath -Name 'MiniReportEnabled' -Type DWord -Value 0; break }
        'Mini' { Set-ItemProperty -Path $netshRegPath -Name 'MiniReportEnabled' -Type DWord -Value 1; break }
        'Full' { Set-ItemProperty -Path $netshRegPath -Name 'ReportEnabled' -Type DWord -Value 1; break }
    }

    Write-Log "RerpotMode $RerpotMode is configured."
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
        Write-Log "Failed to stop netsh trace ($SessionName). exit code: $LASTEXITCODE; stdout: $stdout; error: $err"
        Write-Log "Stopping with Stop-EtwSession"
        Stop-EtwSession -SessionName $SessionName -ErrorAction SilentlyContinue
    }

    Write-Progress -Activity "Stopping netsh trace" -Status "Done" -Completed
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
        return [Win32.Advapi32]::StopTrace($SessionName)
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

    if (-not (Test-Path $Path -ErrorAction Stop)) {
        New-Item -ItemType Directory $Path -ErrorAction Stop | Out-Null
    }
    $Path = Resolve-Path $Path

    # File name must be ***.mht
    if ([IO.Path]::GetExtension($FileName) -ne ".mht") {
        $FileName = [IO.Path]::GetFileNameWithoutExtension($FileName) + '.mht'
    }

    # For Win7, maxsc is 100
    $maxScreenshotCount = 100

    $win32os = Get-WmiObject win32_operatingsystem
    $osMajor = $win32os.Version.Split(".")[0] -as [int]
    $osMinor = $win32os.Version.Split(".")[1] -as [int]

    if ($osMajor -gt 6 -or ($osMajor -eq 6 -and $osMinor -ge 3)) {
        $maxScreenshotCount = 300
    }

    if (-not (Get-Command 'psr.exe' -ErrorAction SilentlyContinue)) {
        Write-Error "psr.exe is not available."
        return
    }

    $outputFile = Join-Path $Path -ChildPath $FileName
    if ($ShowGUI) {
        & psr /start /maxsc $maxScreenshotCount /maxlogsize 10 /output $outputFile /exitonsave 1 /noarc 1
    }
    else {
        & psr /start /maxsc $maxScreenshotCount /maxlogsize 10 /output $outputFile /exitonsave 1 /gui 0 /noarc 1
    }

    # PSR doesn't return anything even on failure. Check if process is spawned.
    $process = Get-Process -Name psr -ErrorAction SilentlyContinue
    if (-not $process) {
        Write-Error "PSR failed to start"
        return
    }

    Write-Log "PSR started $(if ($ShowGUI) {'with UI'} else {'without UI'}). PID: $($process.Id), maxScreenshotCount: $maxScreenshotCount"
}

function Stop-PSR {
    [CmdletBinding()]
    param ()

    $currentInstance = Get-Process -Name psr -ErrorAction SilentlyContinue

    if (-not $currentInstance) {
        Write-Error 'There is no psr.exe process'
        return
    }

    Write-Log 'Stopping PSR'
    $stopInstance = Start-Process 'psr' -ArgumentList '/stop' -PassThru

    Wait-Process -InputObject $currentInstance

    # When there were no clicks, the instance of 'psr /stop' remains after the existing instance exits. This causes a hung.
    # The existing instance is supposed to signal an event and 'psr /stop' instance is waiting for this event to be signaled. But it seems this does not happen when there were no clicks.
    # So to avoid this, the following code manually signal the handle so that 'psr /stop' shuts down.
    try {
        $PSR_CLEANUP_COMPLETED = '{CD3E5009-5C9D-4E9B-B5B6-CAE1D8799AE3}'
        $h = [System.Threading.EventWaitHandle]::OpenExisting($PSR_CLEANUP_COMPLETED)
        $h.Set() | Out-Null
        Write-Log "PSR_CLEANUP_COMPLETED was manually signaled."
        Wait-Process -InputObject $stopInstance
    }
    catch {
        # ignore
    }
    finally {
        if ($stopInstance) {
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

    if (-not (Test-Path $Path -ErrorAction Stop)) {
        New-Item -ItemType directory $Path | Out-Null
    }
    $Path = Resolve-Path $Path

    $logs = @(
        'Application'
        'System'
        (wevtutil el) -match "Microsoft-Windows-Windows Firewall With Advanced Security|AAD|Microsoft-Windows-Bits-Client|WebAuth|CAPI2"
    )

    $tasks = @(
        foreach ($log in $logs) {
            $fileName = $log.Replace('/', '_') + '.evtx'
            $filePath = Join-Path $Path -ChildPath $fileName
            Write-Log "Saving $log to $filePath"
            Start-Task -ScriptBlock {
                param ($log, $filePath)
                wevtutil export-log $log $filePath /ow
                wevtutil archive-log $filePath
                # wevtutil archive-log $filePath /locale:en-US
                # wevtutil archive-log $filePath /locale:ja-JP
            } -ArgumentList $log, $filePath
        }
    )

    $tasks | Receive-Task -AutoRemoveTask
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
        New-Item $Path -ItemType directory -ErrorAction Stop | Out-Null
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
            $installedOn = New-Object string -ArgumentList (, $($installedOnRaw.ToCharArray() | Where-Object { $_ -lt 128 }))

            # https://docs.microsoft.com/en-us/windows/win32/shell/folder-getdetailsof
            [PSCustomObject]@{
                Name        = $item.Name
                Program     = $appUpdates.GetDetailsOf($item, 2)
                Version     = $appUpdates.GetDetailsOf($item, 3)
                Publisher   = $appUpdates.GetDetailsOf($item, 4)
                URL         = $appUpdates.GetDetailsOf($item, 7)
                InstalledOn = $installedOn
            }
            [System.Runtime.Interopservices.Marshal]::FinalReleaseComObject($item) | Out-Null
        }
    }
    finally {
        if ($appUpdates) {
            [System.Runtime.Interopservices.Marshal]::FinalReleaseComObject($appUpdates) | Out-Null
        }
        if ($shell) {
            [System.Runtime.Interopservices.Marshal]::FinalReleaseComObject($shell) | Out-Null
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

    # Note: WMI Win32_UserAccount can be very slow. I'm avoiding here.
    # Get-WmiObject -Class Win32_UserAccount -Filter "Name = '$userName'"

    # Is is SID?
    $sid = $account = $nul
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
        Write-Error "Cannot resolve $Identity."
        return
    }

    $resolved = [PSCustomObject]@{
        Name = $account.ToString()
        SID  = $sid.ToString()
    } | Add-Member -MemberType ScriptMethod -Name 'ToString' -Value { $this.Name } -Force -PassThru

    # Add to cache
    $Script:ResolveCache.Add($resolved.Name, $resolved)
    $Script:ResolveCache.Add($resolved.SID, $resolved)

    $resolved
}

function Get-LogonUser {
    [CmdletBinding()]
    param(
        [switch]$IgnoreCache
    )

    # If there's a cache, use it unless IgnoreCache is specified
    if (-not $IgnoreCache -and $Script:LogonUser) {
        Write-Log "Returning a cache."
        return $Script:LogonUser
    }

    # quser.exe might not be available.
    if (-not (Get-Command quser.exe -ErrorAction SilentlyContinue)) {
        Write-Error "quser.exe is not available."
        return
    }

    $err = $($quserResult = Invoke-Command {
            $ErrorActionPreference = 'Continue'
            & quser.exe
        }) 2>&1

    if ($err -or $LASTEXITCODE -ne 0) {
        Write-Error "quser failed. exit code:$LASTEXITCODE; stdout:`"$quserResult`"; error:`"$err`""
        return
    }

    $currentSession = $quserResult | Where-Object { $_.StartsWith('>') } | Select-Object -First 1
    if (-not $currentSession) {
        Write-Error "Cannot find current session with quser."
        return
    }

    Write-Log "Current session: $currentSession"
    $match = [Regex]::Match($currentSession, '^>(?<name>.+?)\s{2,}')
    $userName = $match.Groups['name'].Value

    $Script:LogonUser = Resolve-User $userName
    $Script:LogonUser
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
        $userRegRoot = "HKEY_USERS\$($resolvedUser.SID)"

        if (-not ($userRegRoot -and (Test-Path "Registry::$userRegRoot"))) {
            Write-Error "Cannot find $userRegRoot."
            return
            # Write-Log "Cannot find $userRegRoot. Falling back to HKCU"
            # $userRegRoot = 'HKCU'
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
    $userProfile = Get-ItemProperty "Registry::HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList\$($resolvedUser.SID)\"
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
        New-Item $Path -ItemType directory -ErrorAction Stop | Out-Null
    }

    $registryKeys = @(
        "HKCU\Software\Microsoft\Exchange"
        "HKCU\Software\Policies\Microsoft\Exchange"
        "HKCU\Software\Microsoft\Office"
        "HKCU\Software\Policies\Microsoft\Office"
        "HKCU\Software\Wow6432Node\Microsoft\Office"
        "HKCU\Software\Wow6432Node\Policies\Microsoft\Office"
        "HKCU\SOFTWARE\Classes\Local Settings\Software\Microsoft\MSIPC"
        "HKCU\Software\Policies"
        "HKCU\Software\IM Providers"
        "HKCU\Software\Microsoft\Windows\CurrentVersion\Notifications"
        "HKLM\Software\Microsoft\Office"
        "HKLM\Software\Policies\Microsoft\Office"
        "HKLM\Software\WOW6432Node\Microsoft\Office"
        "HKLM\Software\WOW6432Node\Policies\Microsoft\Office"

        # This is WinInet proxy settings and maybe out of place, but I wanted to collect for now.
        'HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings'
    )

    $userRegRoot = Get-UserRegistryRoot $User -SkipRegistryPrefix
    if ($userRegRoot) {
        $registryKeys = $registryKeys | ForEach-Object { $_.Replace("HKCU", $userRegRoot).TrimEnd('\') }
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
                & $regexe export $key $filePath | Out-Null
            }) 2>&1

        if ($LASTEXITCODE -ne 0) {
            Write-Error "$key is not exported. exit code = $LASTEXITCODE. $err"
        }
    }
}

function Save-OSConfiguration {
    [CmdletBinding()]
    param (
        [parameter(Mandatory = $true)]
        $Path
    )

    if (-not (Test-Path $Path)) {
        New-Item $Path -ItemType directory -ErrorAction Stop | Out-Null
    }

    $commands = @(
        @{ScriptBlock = { Get-WmiObject -Class Win32_ComputerSystem }; FileName = 'Win32_ComputerSystem.xml' }
        @{ScriptBlock = { Get-WmiObject -Class Win32_OperatingSystem }; FileName = 'Win32_OperatingSystem.xml' }
        @{ScriptBlock = { Get-WinHttpDefaultProxy } }
        @{ScriptBlock = { Get-NLMConnectivity } }
        @{ScriptBlock = { Get-MeteredNetworkCost } }
        @{ScriptBlock = { Get-WSCAntivirus } }
        @{ScriptBlock = { Get-InstalledUpdate } }
        @{ScriptBlock = { Get-JoinInformation } }
        @{ScriptBlock = { Get-DeviceJoinStatus }; FileName = 'DeviceJoinStatus.txt' }
        @{ScriptBlock = { Get-WebView2 } }
        # this is just for troubleshooting.
        @{ScriptBlock = { Get-ChildItem 'Registry::HKEY_USERS' | Select-Object 'Name' }; FileName = 'Users.xml' }
        @{ScriptBlock = { whoami.exe /USER }; FileName = 'whoami.txt' }
    )

    foreach ($command in $commands) {
        Run-Command @command -Path $Path
    }
}

function Save-OSConfigurationMT {
    [CmdletBinding()]
    param (
        [parameter(Mandatory = $true)]
        $Path
    )

    if (-not (Test-Path $Path)) {
        New-Item $Path -ItemType directory -ErrorAction Stop | Out-Null
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
}

function Save-NetworkInfo {
    [CmdletBinding()]
    param (
        [parameter(Mandatory = $true)]
        $Path
    )

    if (-not (Test-Path $Path)) {
        New-Item $Path -ItemType directory -ErrorAction Stop | Out-Null
    }

    # These are from C:\Windows\System32\gatherNetworkInfo.vbs with some extra.
    $commands = @(
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
        @{ScriptBlock = { ipconfig /all } }

        # Dump Windows Firewall config
        @{ScriptBlock = { netsh advfirewall monitor show currentprofile } }
        @{ScriptBlock = { netsh advfirewall monitor show firewall } }
        @{ScriptBlock = { netsh advfirewall monitor show consec } }
        @{ScriptBlock = { netsh advfirewall firewall show rule name=all verbose } }
        @{ScriptBlock = { netsh advfirewall consec show rule name=all verbose } }
        @{ScriptBlock = { netsh advfirewall monitor show firewall rule name=all } }
        @{ScriptBlock = { netsh advfirewall monitor show consec rule name=all } }
    )

    foreach ($command in $commands) {
        Run-Command @command -Path $Path
    }
}

<#
.DESCRIPTION
Run a given script block. If Path is given, save the result there.
If FileName is given, it's used for the file name for saving the result. If its extension is not ".xml", Out-File will be used. Otherwise Export-CliXml will be used.
If FileName is not give, the file name will be auto-decided. If the command is an application, then Out-File will be used. Otherwise Export-CliXml will be used.
#>
function Run-Command {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true, Position = 0)]
        [ScriptBlock]$ScriptBlock,
        [object[]]$ArgumentList,
        # Folder to save to
        $Path,
        # File name used for saving
        [string]$FileName
    )

    $result = $null
    $sw = [System.Diagnostics.Stopwatch]::StartNew()

    try {
        # To redirect error, call operator (&) is used, instead of $ScriptBlock.InvokeReturnAsIs().
        $err = $($result = & $ScriptBlock @ArgumentList) 2>&1
        foreach ($e in $err) {
            Write-Log "'$ScriptBlock' had a non-terminating error. $e" -ErrorRecord $e
        }
    }
    catch {
        Write-Log "'$ScriptBlock' threw a terminating error. $_" -ErrorRecord $_
    }

    $sw.Stop()
    Write-Log "'$ScriptBlock' took $($sw.ElapsedMilliseconds) ms. $(if ($null -eq $result) {"It returned nothing."})"

    if ($null -eq $result) {
        return
    }

    # If Path is given, save the result.
    if ($Path) {
        $sb = $ScriptBlock.ToString()
        $exportAsXml = $true

        if ($FileName) {
            if ([IO.Path]::GetExtension($FileName) -ne '.xml') {
                $exportAsXml = $false
            }
        }
        else {
            # Decide the filename & export method based on the command type
            $Command = ([RegEx]::Match($sb, '\w+-\w+')).Value
            if (-not $Command) {
                $Command = $sb
            }

            $Command = $Command.Trim()
            if ($Command.IndexOf(' ') -ge 0) {
                $commandName = $Command.SubString(0, $Command.IndexOf(' '))
            }
            else {
                $commandName = $Command
            }

            $cmd = Get-Command $commandName -ErrorAction SilentlyContinue
            if ($cmd.CommandType -eq 'Application') {
                # To be more strict, I could use [System.IO.Path]::GetInvalidFileNameChars(). But it's ok for now.
                $FileName = $Command.Replace('/', '-') + ".txt"
                $exportAsXml = $false
            }
            else {
                $FileName = $commandName.SubString($commandName.IndexOf('-') + 1) + ".xml"
            }
        }

        if (-not (Test-Path $Path)) {
            New-Item $Path -ItemType Directory -ErrorAction SilentlyContinue | Out-Null
        }

        if ($exportAsXml) {
            $result | Export-Clixml -Path (Join-Path $Path $FileName)
        }
        else {
            $result | Out-File -FilePath (Join-Path $Path $FileName)
        }
    }
    else {
        $result
    }
}

function Save-NetworkInfoMT {
    [CmdletBinding()]
    param (
        [parameter(Mandatory = $true)]
        $Path
    )

    if (-not (Test-Path $Path)) {
        New-Item $Path -ItemType directory -ErrorAction Stop | Out-Null
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
        $User = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name.Split('\')[1]
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
    $proxySettingsPerUser = Get-ItemProperty 'Registry::HKLM\SOFTWARE\Policies\Microsoft\Windows\CurrentVersion\Internet Settings' -Name 'ProxySettingsPerUser' -ErrorAction SilentlyContinue | Select-Object -ExpandProperty 'ProxySettingsPerUser'

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
        $props = [ordered]@{}
        $props['ProxySettingsPerUser'] = $proxySettingsPerUser
        $props['User'] = $User
        $props['Connection'] = 'DefaultConnectionSettings'
        $props['AutoDetect'] = $true
        $props['AutoConfigUrl'] = $null
        $props['Proxy'] = $null
        $props['ProxyBypass'] = $null
        $props['ActiveConnectionProxy'] = $currentUserActiveConnProxy
        [PSCustomObject]$props
        return
    }

    foreach ($connection in $connections) {
        # Skip SavedLegacySettings & WinHttpSettings (in HKLM)
        if ($connection -eq 'SavedLegacySettings' -or $connection -eq 'WinHttpSettings') {
            continue
        }

        $raw = $null
        $raw = Get-ItemProperty $connectionsKey -Name $connection -ErrorAction SilentlyContinue | Select-Object -ExpandProperty $connection

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

        # This data is temporarily.
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
    [Win32.Kernel32]::GlobalFree($Ptr) | Out-Null
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
            Write-Error -Message "Downloading a PAC file failed from $Url. $_" -Exception $_.Exception
        }
        finally {
            if ($response) { $response.Dispose() }
            if ($copied) { $copied.Dispose() }
        }

        $result
    }

    foreach ($proxy in @(Get-WinInetProxy -User $User)) {
        # If AutoDetect is on, detect URL with WPAD using WinHttpDetectAutoProxyConfigUrl.
        if ($proxy.AutoDetect) {
            [Win32.SafeGlobalFreeString]$wpadUrl = $null
            if ([Win32.WinHttp]::WinHttpDetectAutoProxyConfigUrl([Win32.WinHttp+AutoDetectType] 'WINHTTP_AUTO_DETECT_TYPE_DHCP, WINHTTP_AUTO_DETECT_TYPE_DNS_A', [ref]$wpadUrl)) {
                Get-PAC $wpadUrl.ToString() | ForEach-Object { $_.Add('WPAD', $true); $_.Add('User', $proxy.User); [PSCustomObject]$_ }
            }
            else {
                $ec = [System.Runtime.InteropServices.Marshal]::GetLastWin32Error()
                $winhttpEc = $ec -as [Win32.WinHttp+Error]

                if ($winhttpEc) {
                    Write-Error "WinHttpDetectAutoProxyConfigUrl failed with $winhttpEc ($($winhttpEc.value__))"
                }
                else {
                    Write-Error "WinHttpDetectAutoProxyConfigUrl failed with $ec"
                }

                return
            }
        }

        if ($proxy.AutoConfigUrl) {
            Get-PAC $proxy.AutoConfigUrl | ForEach-Object { $_.Add('WPAD', $false); $_.Add('User', $proxy.User); [PSCustomObject]$_ }
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
    Write-Log ("INetworkListManager::GetConnectivity: $connectivity (0x$("{0:x8}" -f $connectivity.value__))")

    $refCount = [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($nlm)
    Write-Log "NetworkListManager COM object's remaining ref count: $refCount"
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

function Get-OutlookProfile {
    [CmdletBinding()]
    param(
        [string]$User
    )

    if (-not $User) {
        $User = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
    }

    $userRegRoot = Get-UserRegistryRoot $User
    if (-not $userRegRoot) {
        return
    }

    # List Outlook "profiles" keys
    $profiles = New-Object System.Collections.Generic.List[object]
    $versionKeys = @(Get-ChildItem (Join-Path $userRegRoot 'Software\Microsoft\Office\') -ErrorAction SilentlyContinue | Where-Object { $_.Name -match '\d\d\.0' })
    $defaultProfile = $null

    foreach ($versionKey in $versionKeys) {
        Get-ChildItem (Join-Path $versionKey.PsPath '\Outlook\Profiles') -ErrorAction SilentlyContinue | ForEach-Object { $profiles.Add($_) }
        if (-not $defaultProfile) {
            $defaultProfile = Get-ItemProperty (Join-Path $versionKey.PsPath 'Outlook') -Name 'DefaultProfile' -ErrorAction SilentlyContinue | Select-Object -ExpandProperty 'DefaultProfile'
        }
        $versionKey.Close()
    }

    if (-not $profiles.Count) {
        Write-Log "There are no profiles for $User."
        return
    }

    $PR_PROFILE_CONFIG_FLAGS = '00036601'

    # Flag constants:
    # This corresponds [Use Cached Exchange Mode]
    $CONFIG_OST_CACHE_PRIVATE = 0x180
    # This corresponds to [Download Public Folder Favorites]
    $CONFIG_OST_CACHE_PUBLIC = 0x400
    # This corresponds to [Download shared folders]
    $CONFIG_OST_CACHE_DELEGATE_PIM = 0x800

    foreach ($profile in $profiles) {
        $flags = 0
        $CACHE_PRIVATE = $false
        $CACHE_PUBLIC = $false
        $CACHE_DELEGATE_PIM = $false

        $subkeys = Get-ChildItem $profile.PSPath -Recurse

        foreach ($subkey in $subkeys) {
            if ($subkey.Property | Where-Object { $_ -eq $PR_PROFILE_CONFIG_FLAGS }) {
                $bytes = $subkey.GetValue($PR_PROFILE_CONFIG_FLAGS)
                $flags = [BitConverter]::ToUInt32($bytes, 0)
                break
            }
        }

        # Close all the sub keys
        $subkeys | ForEach-Object { $_.Close() }

        if (($flags -band $CONFIG_OST_CACHE_PRIVATE) -ne 0) {
            $CACHE_PRIVATE = $true
        }

        if (($flags -band $CONFIG_OST_CACHE_PUBLIC) -ne 0) {
            $CACHE_PUBLIC = $true
        }

        if (($flags -band $CONFIG_OST_CACHE_DELEGATE_PIM) -ne 0) {
            $CACHE_DELEGATE_PIM = $true
        }

        [PSCustomObject]@{
            User                          = $User
            Profile                       = $profile.Name
            IsDefault                     = (Split-Path $profile.Name -Leaf) -eq $defaultProfile
            CachedMode                    = $CACHE_PRIVATE -or $CACHE_PUBLIC -or $CACHE_DELEGATE_PIM
            DownloadPublicFolderFavorites = $CACHE_PUBLIC
            DownloadSharedFolders         = $CACHE_DELEGATE_PIM
            PR_PROFILE_CONFIG_FLAGS       = $flags
        }

        $profile.Close()
    }
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
    $officeInfo = Get-OfficeInfo -ErrorAction Stop
    $ver = ($officeInfo.Version.Split('.')[0] -as [int]).ToString('00.0')

    foreach ($keyPath in @("SOFTWARE\Policies\Microsoft\Office\$ver\Outlook", "SOFTWARE\Microsoft\Office\$ver\Outlook")) {
        $forcePstPath = Get-ItemProperty $(Join-Path $userRegRoot $keyPath) -Name 'ForcePSTPath' -ErrorAction SilentlyContinue | Select-Object -ExpandProperty 'ForcePSTPath'
        if ($forcePstPath) {
            [PSCustomObject]@{
                Name = 'ForcePSTPath'
                Path = [System.Environment]::ExpandEnvironmentVariables($forcePstPath)
            }

            # If ForcePSTPath is found in the policy key, no need to check the rest.
            break
        }
    }
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

    if (-not (Test-Path $Path)) {
        New-Item $Path -ItemType Directory -ErrorAction Stop | Out-Null
    }

    foreach ($cachePath in @(Get-CachedAutodiscoverLocation -User $User)) {
        Write-Log "Searching $($cachePath.Name) $($cachePath.Path)"
        # Get Autodiscover XML files and copy them to Path
        try {
            if (Test-Path $cachePath.Path) {
                # Use recurse only for the path under LOCALAPPDATA.
                Get-ChildItem $cachePath.Path -Filter '*Autod*.xml' -Force -Recurse:$($cachePath.Name -eq 'UnderLocalAppData') | Copy-Item -Destination $Path
            }
        }
        catch {
            # Just in case Copy-Item throws a terminating error.
            Write-Error -ErrorRecord $_
        }
    }

    # Remove Hidden attribute
    foreach ($file in @(Get-ChildItem $Path -Force)) {
        if ((Get-ItemProperty $file.FullName).Attributes -band [IO.FileAttributes]::Hidden) {
            try {
                # Unfortunately, this does not work before PowerShellV5.
                (Get-ItemProperty $file.FullName).Attributes -= 'Hidden'
                continue
            }
            catch {
                # ignore error
            }

            # This could fail if attributes other than Archive, Hidden, Normal, ReadOnly, or System are set (such as NotContentIndexed)
            Set-ItemProperty $file.Fullname -Name Attributes -Value ((Get-ItemProperty $file.FullName).Attributes -bxor [IO.FileAttributes]::Hidden)
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

    Get-CachedAutodiscoverLocation | ForEach-Object {
        Get-ChildItem -LiteralPath $_.Path -Filter '*Autod*.xml' -Force -Recurse:$($_.Name -eq 'UnderLocalAppData') | ForEach-Object { Remove-Item $_.FullName -Force }
    }
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
        New-Item $Path -ItemType directory -ErrorAction Stop | Out-Null
    }

    $Path = Resolve-Path $Path

    # Process name must contain the extension such as "Outlook.exe", instead of "Outlook"
    if ([IO.Path]::GetExtension($TargetProcess) -ne 'exe') {
        $TargetProcess = [IO.Path]::GetFileNameWithoutExtension($TargetProcess) + ".exe"
    }

    # Create a registry key under HKLM\SYSTEM\CurrentControlSet\Services\ldap\tracing
    $keypath = "HKLM:\SYSTEM\CurrentControlSet\Services\ldap\tracing"
    if (-not (Test-Path $keypath)) {
        New-Item (Split-Path $keypath) -Name 'tracing' -ErrorAction SilentlyContinue | Out-Null
    }

    New-Item $keypath -Name $TargetProcess -ErrorAction SilentlyContinue | Out-Null
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
    Stop-EtwSession $SessionName | Out-Null

    # Remove a registry key under HKLM\SYSTEM\CurrentControlSet\Services\ldap\tracing (ignore any errors)

    # Process name must contain the extension such as "Outlook.exe", instead of "Outlook"
    if ([IO.Path]::GetExtension($TargetProcess) -ne 'exe') {
        $TargetProcess = [IO.Path]::GetFileNameWithoutExtension($TargetProcess) + ".exe"
    }

    $keypath = "HKLM:\SYSTEM\CurrentControlSet\Services\ldap\tracing\$TargetProcess"
    Remove-Item $keypath -ErrorAction SilentlyContinue | Out-Null
}

function Save-OfficeModuleInfo {
    [CmdletBinding()]
    param (
        [parameter(Mandatory = $true)]
        $Path,
        # filter items by their name using -match (e.g. 'outlook.exe','mso\d\d.*\.dll'). These are treated as "OR".
        [string[]]$Filters,
        [Threading.CancellationToken]$CancellationToken
    )

    if (-not (Test-Path $Path)) {
        New-Item -ItemType Directory $Path -ErrorAction Stop | Out-Null
    }

    # If MS Office is not installed, bail.
    $officeInfo = Get-OfficeInfo -ErrorAction SilentlyContinue
    if (-not $officeInfo) {
        Write-Error "MS Office is not installed."
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
    )

    Write-Log "officePaths are $officePaths"

    $sw = [System.Diagnostics.Stopwatch]::StartNew()

    # Get exe and dll
    # It's slightly faster to run gci twice with -Filter than running once with -Include *exe, *.dll
    $items = @(
        foreach ($officePath in $officePaths) {
            Get-ChildItem -Path $officePath -Filter *.exe -Recurse -ErrorAction SilentlyContinue
            Get-ChildItem -Path $officePath -Filter *.dll -Recurse -ErrorAction SilentlyContinue
        }
    )

    $listingFinished = $sw.Elapsed
    Write-Log "Listing $($items.Count) items took $($listingFinished.TotalMilliseconds) ms."

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

    $cmdletName = $PSCmdlet.MyInvocation.MyCommand.Name
    $name = $cmdletName.Substring($cmdletName.IndexOf('-') + 1)

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

            if ($CancellationToken.IsCancellationRequested) {
                Write-Log "Cancel request acknowledged."
                break
            }

        }) | Export-Clixml -Depth 4 -Path $(Join-Path $Path -ChildPath "$name.xml") #-Encoding UTF8

    $sw.Stop()
    Write-Log "Enumerating items took $(($sw.Elapsed - $listingFinished).TotalMilliseconds) ms."
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
        New-Item -ItemType Directory $Path -ErrorAction Stop | Out-Null
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

    Write-Log "Job (ID: $($job.Id)) has started. A Named Event (Handle: $($namedEvent.Handle), Name: '$eventName') is created"
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
        Write-Log "Waiting for the job (ID: $($job.Id)) up to $TimeoutSecond seconds."
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
            $namedEvent.Set() | Out-Null
            $namedEvent.Close()
            Write-Log "Event (Handle: $($namedEvent.Handle)) was closed."
        }
        catch {
            Write-Error -ErrorRecord $_
        }

        # Let the job finish
        Wait-Job -Job $job | Out-Null
        Stop-Job -Job $job
        # Receive-Job -Job $job
        Remove-Job -Job $job
        Write-Log "Job (ID: $($job.Id)) was removed."
    }
}

function Save-MSInfo32 {
    [CmdletBinding()]
    param(
        [parameter(Mandatory = $true)]
        $Path
    )

    if (-not (Test-Path $Path -ErrorAction Stop)) {
        New-Item -ItemType Directory $Path -ErrorAction Stop | Out-Null
    }

    $filePath = Join-Path $Path -ChildPath "$($env:COMPUTERNAME).nfo"

    $processName = "msinfo32.exe"
    $process = $null

    try {
        $process = Start-Process $processName -ArgumentList "/nfo $filePath" -Wait -PassThru
        if ($process.ExitCode -ne 0) {
            Write-Error "$processName failed. exit code = $($process.ExitCode)"
        }
    }
    catch {
        Write-Error -Message "Failed to start $processName.`n$_" -Exception $_.Exception
    }
    finally {
        if ($process) {
            $process.Dispose()
        }
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
        New-Item $Path -ItemType directory -ErrorAction Stop | Out-Null
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
        Write-Error "logman failed. exit code: $LASTEXITCODE; stdout: $logmanResult"
        return
    }

    # Note: Depending on the OS version, not all providers are available.
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
    Stop-EtwSession $SessionName | Out-Null
}

function Start-FiddlerCap {
    [CmdletBinding()]
    param(
        [parameter(Mandatory = $true)]
        $Path
    )

    if (-not (Test-Path $Path -ErrorAction Stop)) {
        New-Item -ItemType Directory $Path -ErrorAction Stop | Out-Null
    }

    $Path = Resolve-Path $Path
    $fiddlerPath = Join-Path $Path -ChildPath "FiddlerCap"
    $fiddlerExe = Join-Path $fiddlerPath -ChildPath 'FiddlerCap.exe'

    #  FiddlerCap is not available.
    if (-not (Test-Path $fiddlerExe)) {
        $fiddlerCapUrl = "https://telerik-fiddler.s3.amazonaws.com/fiddler/FiddlerCapSetup.exe"
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
    if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
        Write-Warning "Please run as administrator."
        return
    }

    if (-not (Test-Path $Path)) {
        New-Item $Path -ItemType Directory -ErrorAction Stop | Out-Null
    }

    $Path = Resolve-Path $Path
    $procmonFile = $null

    # Search procmon.exe or procmon64.exe under $ProcmonSearchPath (including subfolders).
    if ($ProcmonSearchPath -and (Test-Path $ProcmonSearchPath)) {
        $findResult = @(Get-ChildItem -Path $ProcmonSearchPath -Filter 'procmon*.exe' -Recurse)
        if ($findResult.Count -ge 1) {
            $procmonFile = $findResult[0].FullName
            if ($env:PROCESSOR_ARCHITECTURE -eq 'AMD64') {
                $procmon64 = $findResult | Where-Object { $_.Name -eq 'procmon64.exe' } | Select-Object -First 1
                if ($procmon64) {
                    $procmonFile = $procmon64.FullName
                }
            }
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
            New-Item $procmonFolderPath -ItemType Directory -ErrorAction Stop | Out-Null
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
        ProcmonProcessId     = $process.Id
        PMLFile              = $pmlFile
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
    Write-Progress -Activity "Stopping procmon" -Status "Please wait" -PercentComplete -1
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

    Write-Progress -Activity "Stopping procmon" -Status "Done" -Completed
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
        New-Item $keypath -ErrorAction Stop | Out-Null
    }

    Write-Log "Starting a TCO trace by setting up $keypath"
    New-ItemProperty $keypath -Name 'TCOTrace' -PropertyType DWORD -Value 7 -ErrorAction SilentlyContinue | Out-Null
    New-ItemProperty $keypath -Name 'MsoHttpVerbose' -PropertyType DWORD -Value 1 -ErrorAction SilentlyContinue | Out-Null

    # If failed, throw a terminating error
    Get-ItemProperty $keypath -Name 'TCOTrace' -ErrorAction Stop | Out-Null
    Get-ItemProperty $keypath -Name 'MsoHttpVerbose' -ErrorAction Stop | Out-Null
}

function Stop-TcoTrace {
    [CmdletBinding()]
    param(
        [parameter(Mandatory = $true)]
        $Path,
        [string]$User
    )

    if (-not (Test-Path $Path)) {
        New-Item $Path -ItemType Directory -ErrorAction Stop | Out-Null
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
    Remove-ItemProperty $keypath -Name 'TCOTrace' -ErrorAction SilentlyContinue | Out-Null
    Remove-ItemProperty $keypath -Name 'MsoHttpVerbose' -ErrorAction SilentlyContinue | Out-Null

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
function Start-TTD {
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
        New-Item $Path -ItemType Directory -ErrorAction Stop | Out-Null
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
        Write-Error "tttracer.exe failed to start. ExitCode: $($process.ExitCode). $(Get-Content $stderr)"
        $process.Dispose()
        return
    }

    $targetProcess = $null

    if (-not $OnLaunch) {
        # Find out the new process instantiated by tttracer.exe. This might take a bit.
        # The new process starts as a child process of tttracer.exe.
        $maxRetry = 3
        foreach ($i in 1..$maxRetry) {
            if ($newProcess = Get-WmiObject Win32_Process -Filter "Name='$targetName.exe' AND ParentProcessId='$($process.Id)'") {
                $targetProcess = Get-Process -Id $newProcess.ProcessId
                break
            }

            Start-Sleep -Seconds $i
        }

        if (-not $targetProcess) {
            Write-Error "Cannot find the new instance of $targetName."
            return
        }

        Write-Log "Target process $($targetProcess.Name) (PID: $($targetProcess.Id)) has started."

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

function Stop-TTD {
    [CmdletBinding()]
    param(
        # The returned object of Start-TTD
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
        Write-Error "Invalid input. tttracer PID: $($tttracerProcess.ID), target process PID: $($targetProcess.ID)"
        return
    }

    $stopTarget = 'all'
    if (Get-Process -Id $targetProcess.Id -ErrorAction SilentlyContinue) {
        $stopTarget = $targetProcess.ID
    }
    else {
        Write-Log "Target process $($targetProcess.Name) (PID: $($targetProcess.Id)) does not exist; ExitCode: $($targetProcess.ExitCode), ExitTime: $(if ($targetProcess.ExitTime) {$targetProcess.ExitTime.ToString('o')}), ElapsedTime: $($targetProcess.ExitTime - $targetProcess.StartTime)"
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
        Write-Error $("'tttracer -stop' failed. ExitCode: 0x{0:x}" -f $exitCode)
    }

    if ($onLaunch) {
        Write-Log "Killing tttracer (PID: $($tttracerProcess.Id)) running in OnLaunch mode."
        $tttracerProcess.Kill()
        $message += ";" + (& $tttracer -cleanup)
    }

    # Wait for the tttracer to exit.
    # Wait-Process writes a non-terminating error when the process has exited. Ignore this error.
    $(Wait-Process -InputObject $tttracerProcess -ErrorAction SilentlyContinue) 2>&1 | Out-Null

    [PSCustomObject]@{
        ExitCode = $exitCode  # This is the exit code of "tttracer -stop"
        Message  = $message # message of "tttracer -stop"
    }

    if ($AutoRemove) {
        $ttd.Dispose()
    }
}

function Attach-TTD {
    [CmdletBinding()]
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
        New-Item $Path -ItemType Directory -ErrorAction Stop | Out-Null
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
        Write-Error "tttracer.exe failed to attach. ExitCode: $exitCodeHex; Error: $err.`n$stderrContent"
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
        $keys = @(Get-ChildItem HKLM:\SOFTWARE\Microsoft\Office\ | Where-Object { [RegEx]::IsMatch($_.PSChildName, '\d\d\.0') -or $_.PSChildName -eq 'ClickToRun' })

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

    $outlookReg = Get-ItemProperty 'HKLM:\SOFTWARE\Clients\Mail\Microsoft Outlook' -ErrorAction SilentlyContinue
    if ($outlookReg) {
        $mapiDll = Get-ItemProperty $outlookReg.DLLPathEx -ErrorAction SilentlyContinue
    }

    $Script:OfficeInfoCache =
    [PSCustomObject]@{
        DisplayName     = $displayName
        Version         = $version
        InstallPath     = $installPath
        MapiDllFileInfo = $mapiDll
    }

    $Script:OfficeInfoCache
}

function Add-WerDumpKey {
    [CmdletBinding()]
    param (
        [parameter(Mandatory = $true)]
        [string]$TargetProcess, # Target Process (e.g. Outlook.exe)
        [parameter(Mandatory = $true)]
        $Path # Folder to save dump files
    )

    if (-not (Test-Path $Path)) {
        New-Item $Path -ItemType Directory -ErrorAction Stop | Out-Null
    }

    $Path = (Resolve-Path $Path -ErrorAction Stop).Path

    # Check if $TargetProcess ends with ".exe".
    if (-not $TargetProcess.EndsWith(".exe")) {
        Write-Log "$TargetProcess does not end with '.exe'.  Adding '.exe'"
        $TargetProcess += '.exe'
    }

    # Create a key 'LocalDumps' under HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\Windows Error Reporting\LocalDumps, if it doesn't exist
    $werKey = 'HKLM:\SOFTWARE\Microsoft\Windows\Windows Error Reporting'
    if (-not (Test-Path (Join-Path $werKey 'LocalDumps'))) {
        New-Item $werKey -Name 'LocalDumps' -ErrorAction Stop | Out-Null
    }

    # Create a ProcessName key under LocalDumps, if it doesn't exist
    $localDumpsKey = Join-Path $werKey 'LocalDumps'
    if (-not (Test-Path (Join-Path $localDumpsKey $TargetProcess))) {
        New-Item $localDumpsKey -Name $TargetProcess -ErrorAction Stop | Out-Null
    }

    # Create "CustomDumpFlags", "DumpType", and "DumpFolder" values in ProcessName key
    $ProcessKey = Join-Path $localDumpsKey $TargetProcess
    Write-Log "Setting up $ProcessKey with CustomDumpFlags:0x61826, DumpType:0, DumpFolder:$Path"
    # -Force will overwrite existing value
    # 0x61826 = MiniDumpWithTokenInformation | MiniDumpIgnoreInaccessibleMemory | MiniDumpWithThreadInfo (0x1000) | MiniDumpWithFullMemoryInfo (0x800) |MiniDumpWithUnloadedModules (0x20) | MiniDumpWithHandleData (4)| MiniDumpWithFullMemory (2)
    New-ItemProperty $ProcessKey -Name 'CustomDumpFlags' -Value 0x00061826 -Force -ErrorAction Stop | Out-Null
    New-ItemProperty $ProcessKey -Name 'DumpType' -Value 0 -PropertyType DWORD -Force -ErrorAction Stop | Out-Null
    New-ItemProperty $ProcessKey -Name 'DumpFolder' -Value $Path -PropertyType String -Force -ErrorAction Stop | Out-Null

    # Rename DW Installed keys to "_Installed" to temporarily disable
    $pcHealth = @(
        # For C2R
        'HKLM:\Software\Microsoft\Office\ClickToRun\REGISTRY\MACHINE\Software\Microsoft\PCHealth\ErrorReporting\DW\Installed'
        'HKLM:\Software\Microsoft\Office\ClickToRun\REGISTRY\MACHINE\Software\Wow6432Node\Microsoft\PCHealth\ErrorReporting\DW\Installed'

        # For MSI
        'HKLM:\Software\Microsoft\PCHealth\ErrorReporting\DW\Installed'
        'HKLM:\Software\Wow6432Node\Microsoft\PCHealth\ErrorReporting\DW\Installed'
    )

    foreach ($installedKey in $pcHealth) {
        if (Test-Path $installedKey) {
            Write-Log "Temporarily renaming $installedKey to `"_Installed`""
            Rename-Item $installedKey -NewName '_Installed'
        }
    }
}

function Remove-WerDumpKey {
    [CmdletBinding()]
    param (
        [parameter(Mandatory = $true)]
        [string]$TargetProcess # Target Process (e.g. Outlook.exe)
    )

    # Check if $TargetProcess ends with ".exe".
    if (-not $TargetProcess.EndsWith(".exe")) {
        Write-Log "$TargetProcess does not end with '.exe'.  Adding '.exe'"
        $TargetProcess += '.exe'
    }

    $werKey = 'HKLM:\SOFTWARE\Microsoft\Windows\Windows Error Reporting'
    $localDumpsKey = Join-Path $werKey 'LocalDumps'
    $ProcessKey = Join-Path $localDumpsKey $TargetProcess

    if (Test-Path $ProcessKey) {
        Write-Log "Removing $ProcessKey"
        Remove-Item $ProcessKey
    }
    else {
        Write-Error "$ProcessKey does not exist"
    }

    # Rename DW "_Installed" keys back to "Installed"
    $pcHealth = @(
        # For C2R
        'HKLM:\Software\Microsoft\Office\ClickToRun\REGISTRY\MACHINE\Software\Microsoft\PCHealth\ErrorReporting\DW\_Installed'
        'HKLM:\Software\Microsoft\Office\ClickToRun\REGISTRY\MACHINE\Software\Wow6432Node\Microsoft\PCHealth\ErrorReporting\DW\_Installed'

        # For MSI
        'HKLM:\Software\Microsoft\PCHealth\ErrorReporting\DW\_Installed'
        'HKLM:\Software\Wow6432Node\Microsoft\PCHealth\ErrorReporting\DW\_Installed'
    )

    foreach ($installedKey in $pcHealth) {
        if (Test-Path $installedKey) {
            Write-Log "Renaming $installedKey back to `"Installed`""
            Rename-Item $installedKey -NewName 'Installed'
        }
    }
}

function Start-WfpTrace {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        $Path,
        [Parameter(Mandatory = $true)]
        [int]$IntervalSeconds,
        [TimeSpan]$MaxDuration = [TimeSpan]::FromHours(1)  # Just for safety, make sure to stop after a period
    )

    if (-not (Test-Path $Path)) {
        New-Item -ItemType directory $Path -ErrorAction Stop | Out-Null
    }
    $Path = Resolve-Path $Path

    # Start WFP trace
    # TODO: This does not return the control sometimes. Figure out why.
    if ($env:PROCESSOR_ARCHITEW6432) {
        $netshexe = Join-Path $env:SystemRoot 'SysNative\netsh.exe'
    }
    else {
        $netshexe = Join-Path $env:SystemRoot 'System32\netsh.exe'
    }

    $filePath = Join-Path $Path 'wfp'
    Write-Log "Starting WFP trace"
    Start-Process $netshexe -ArgumentList "wfp capture start cab=OFF file=`"$filePath`"" -WindowStyle Hidden

    Write-Log "Starting a WFP job"
    $job = Start-Job -ScriptBlock {
        param($Path, $IntervalSeconds, $MaxDuration)

        $expiration = [DateTime]::Now.Add($MaxDuration)

        while ($true) {
            if ([DateTime]::Now -gt $expiration) {
                Write-Output "WfpTrace expired after $MaxDuration"
                break
            }

            # dump filters
            $filterFilePath = Join-Path $Path "filters_$(Get-Date -Format 'yyyyMMdd_HHmmss').xml"
            netsh wfp show filters file=$filterFilePath verbose=on | Out-Null

            # dump netevents
            $eventFilePath = Join-Path $Path "netevents_$(Get-Date -Format 'yyyyMMdd_HHmmss').xml"
            netsh wfp show netevents file="$eventFilePath" <#timewindow=$IntervalSeconds#> | Out-Null
            Start-Sleep -Seconds $IntervalSeconds
        }
    } -ArgumentList $Path, $IntervalSeconds, $MaxDuration

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

    & $netshexe wfp capture stop | Out-Null

    Write-Log "Stopping a WFP job"
    Stop-Job -Job $WfpJob
    Remove-Job -Job $WfpJob
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
    elseif (-not $process.Handle) {
        # This scenario is possible for a system process.
        Write-Error "Cannot obtain the process handle of $($process.Name)."
        return
    }

    if (-not (Test-Path $Path)) {
        New-Item $Path -ItemType Directory -ErrorAction Stop | Out-Null
    }
    $Path = Resolve-Path $Path

    $wow64 = $false
    if (-not $SkipWow64Check) {
        # Check if the target process is WOW6432 (i.e. 32bit on 64bit OS)
        [Win32.Kernel32]::IsWow64Process($process.SafeHandle, [ref]$wow64) | Out-Null
    }

    if ($wow64) {
        $ps32 = Join-Path $env:SystemRoot 'SysWOW64\WindowsPowerShell\v1.0\powershell.exe'
        $command = "& {Import-Module '$Script:MyModulePath' -DisableNameChecking; Save-Dump -Path '$Path' -ProcessId $ProcessId -SkipWow64Check}"
        Write-Log "Invoking $ps32 -c `"$command`""
        $errs = $($result = & $ps32 -NoProfile -WindowStyle Hidden -OutputFormat XML -c $command) 2>&1
        $result
        $errs | ForEach-Object { Write-Error -ErrorRecord $_ }
    }
    else {
        $dumpFile = Join-Path $Path "$($process.Name)_$(Get-Date -Format 'yyyy-MM-dd-HHmmss').dmp"
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
        # Target process (either Name or PID)
        [Parameter(Mandatory = $true, Position = 1)]
        [int]$ProcessId,
        [int]$TimeoutSecond = 5,
        [int]$DumpCount = 1,
        [Threading.CancellationToken]$CancellationToken
    )

    if (-not ($process = Get-Process -Id $ProcessId -ErrorAction SilentlyContinue)) {
        Write-Error "Cannnot find a process with PID $ProcessId."
        return
    }

    $hWnd = $process.MainWindowHandle
    $name = $process.Name

    $ERROR_INVALID_WINDOW_HANDLE = 1400
    $ERROR_TIMEOUT = 1460
    $savedDumpCount = 0

    # Start monitoring
    while ($true) {
        if ($CancellationToken.IsCancellationRequested) {
            Write-Log "Cancel request acknowledged."
            break
        }

        $result = [IntPtr]::Zero
        if (-not ([Win32.User32]::SendMessageTimeoutW($hWnd, 0, [IntPtr]::Zero, [IntPtr]::Zero, 0, $TimeoutSecond * 1000, [ref]$result))) {
            $ec = [Runtime.InteropServices.Marshal]::GetLastWin32Error()

            # if error code is 0 or ERROR_TIMEOUT, timeout occurred.
            if ($ec -eq 0 -or $ec -eq $ERROR_TIMEOUT) {
                # Write-Host "Hung window detected with $name (PID $ProcessId). $($savedDumpCount+1)/$DumpCount" -ForegroundColor Green
                Write-Log "Hung window detected with $name (PID $ProcessId). $($savedDumpCount+1)/$DumpCount"
                $dumpResult = Save-Dump -Path $Path -ProcessId $ProcessId
                $savedDumpCount++
                Write-Log "Saved dump file: $($dumpResult.DumpFile)"

                if ($savedDumpCount -ge $DumpCount) {
                    Write-Log "Dump count reached $DumpCount. Existing."
                    break
                }

                Start-Sleep -Seconds 1
                # To avoid too many dumps in a short time period, wait for one timeout period before starting the next monitoring cycle.
                #Start-Sleep -Seconds $TimeoutSecond
            }
            elseif ($ec -eq $ERROR_INVALID_WINDOW_HANDLE -and $process.HasExited) {
                Write-Log "$($process.Name) (PID $ProcessId) has exited."
                return
            }
            else {
                Write-Error ("SendMessageTimeoutW failed with 0x{0:x8}" -f $ec)
                return
            }
        }
        else {
            # Write-Verbose "SendMessageTimeoutW succeeded"
            Start-Sleep -Seconds 1
        }
    }

    $process.Dispose()
}

function Save-MSIPC {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        $Path,
        $User
    )

    # MSIPC info is in %LOCALAPPDATA%\Microsoft\MSIPC
    if ($localAppdata = Get-UserShellFolder -User $User -ShellFolderName 'Local AppData') {
        $msipcPath = Join-Path $localAppdata 'Microsoft\MSIPC'
    }
    else {
        return
    }

    if (-not (Test-Path $msipcPath)) {
        Write-Error "$msipcPath does not exist"
        return
    }

    if (-not (Test-Path $Path -ErrorAction Stop)) {
        New-Item -ItemType Directory $Path -ErrorAction Stop | Out-Null
    }

    try {
        # Just copy *.ipclog files
        Copy-Item (Join-Path $msipcPath 'Logs\*') -Destination $Path -Exclude '*.lock'
    }
    catch {
        Write-Error -ErrorRecord $_
    }
}

<#
Save DLP policy files
#>
function Save-DLP {
    [CmdletBinding(PositionalBinding = $false)]
    param (
        [Parameter(Mandatory = $true)]
        # Destination folder path to save to
        $Path,
        [string]
        $User
    )

    # Get the path to %LOCALAPPDATA%\Microsoft\Outlook.
    $sourcePath = $null

    if ($localAppdata = Get-UserShellFolder -User $User -ShellFolderName 'Local AppData') {
        $sourcePath = Join-Path $localAppdata -ChildPath 'Microsoft\Outlook'
    }

    $files = @(Get-ChildItem $sourcePath -Filter 'PolicyNudge*')

    if ($files.Count -eq 0) {
        return
    }

    if (-not (Test-Path $Path)) {
        New-Item $Path -ItemType Directory -ErrorAction Stop | Out-Null
    }

    $Path = Resolve-Path $Path

    try {
        Get-ChildItem $sourcePath -Filter 'PolicyNudge*' | Copy-Item -Destination $Path
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
    Register-ObjectEvent -InputObject $proxy -EventName Logging -Action $Callback -MessageData $ArgumentList | Out-Null

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
        $builder.WithRedirectUri($RedirectUri) | Out-Null
    }
    else {
        # WithDefaultRedirectUri() makes the redirect_uri "https://login.microsoftonline.com/common/oauth2/nativeclient".
        # Without it, redirect_uri would be "urn:ietf:wg:oauth:2.0:oob".
        $builder.WithDefaultRedirectUri() | Out-Null
    }

    $writer = $null

    if ($EnableLogging) {
        $logFile = Join-Path (Split-Path $PSCommandPath) 'msal.log'
        [IO.StreamWriter]$writer = [IO.File]::AppendText($logFile)
        Write-Verbose "MSAL Loggin is enabled. Log file: $logFile"

        # Add a CSV header line
        $writer.WriteLine("datetime,level,containsPii,message");

        $builder.WithLogging(
            # Microsoft.Identity.Client.LogCallback
            (New-LogCallback {
                param([Microsoft.Identity.Client.LogLevel]$level, [string]$message, [bool]$containsPii)
                $writer = $Event.MessageData[0]
                $writer.WriteLine("$((Get-Date).ToString('o')),$level,$containsPii,`"$message`"")
            } -ArgumentList $writer),
            [Microsoft.Identity.Client.LogLevel]::Verbose,
            $true, # enablePiiLogging
            $false # enableDefaultPla`tformLogging
        ) | Out-Null
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

    # By default, MSAL asks for scopes: openid, profile, and offline_access.
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
        [switch]$ForceBasicAuth
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
                Write-Log "Received a redirect URL: $redirectUrl"
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
        [string]$ProgID,
        [string]$User
    )

    [uint32]$S_OK = 0

    [Guid]$CLSID = [Guid]::Empty
    [uint32]$hr = [Win32.Ole32]::CLSIDFromProgID($ProgID, [ref]$CLSID)

    if ($hr -ne $S_OK) {
        Write-Verbose -Message $("CLSIDFromProgID for `"$ProgID`" failed with 0x{0:x}. Trying ClickToRun registry." -f $hr)

        $userRegRoot = Get-UserRegistryRoot -User $User

        $locations = @(
            # ClickToRun Registry & the user's Classes
            "Registry::HKLM\SOFTWARE\Microsoft\Office\ClickToRun\REGISTRY\MACHINE\Software\Classes\"
            (Join-Path $userRegRoot "SOFTWARE\Classes\")
        )

        foreach ($loc in $locations) {
            $clsidProp = Get-ItemProperty (Join-Path $loc "$ProgID\CLSID") -ErrorAction SilentlyContinue
            $curVerProp = Get-ItemProperty (Join-Path $loc "$ProgID\CurVer") -ErrorAction SilentlyContinue

            if ($clsidProp) {
                $CLSID = $clsidProp.'(default)'
                break
            }
            elseif ($curVerProp) {
                $curProgID = $curVerProp.'(default)'
                $clsidProp = Get-ItemProperty (Join-Path $loc "$curProgID\CLSID") -ErrorAction SilentlyContinue
                $CLSID = $clsidProp.'(default)'
                break
            }
        }

        if ($CLSID -eq [Guid]::Empty) {
            Write-Error -Message $("CLSIDFromProgID for `"$ProgID`" failed with 0x{0:x}. Also, it was not found in the ClickToRun & user registry" -f $hr)
            return
        }
    }

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

    # Get keys under "Addins"
    $addinKeys = @(
        @(
            'Registry::HKLM\SOFTWARE\Microsoft\Office\Outlook\Addins'
            'Registry::HKLM\SOFTWARE\WOW6432Node\Microsoft\Office\Outlook\Addins'
            'Registry::HKLM\SOFTWARE\Microsoft\Office\ClickToRun\REGISTRY\MACHINE\Software\Microsoft\Office\Outlook\Addins'
            'Registry::HKLM\SOFTWARE\Microsoft\Office\ClickToRun\REGISTRY\MACHINE\Software\Wow6432Node\Microsoft\Office\Outlook\AddIns'
            Join-Path $userRegRoot 'Software\Microsoft\Office\Outlook\Addins'
        ) |
        ForEach-Object {
            Get-ChildItem $_ -ErrorAction SilentlyContinue
        }
    )

    $LoadBehavior = @{
        0  = 'None'
        1  = 'NoneLoaded'
        2  = 'StartupUnloaded'
        3  = 'Startup'
        8  = 'LoadOnDemandUnloaded'
        9  = 'LoadOnDemand'
        16 = 'LoadAtNextStartupOnly'
    }

    $cache = @{}

    foreach ($addin in $addinKeys) {
        $props = @{}
        $props['Path'] = $addin.Name
        $props['ProgID'] = $addin.PSChildName

        # First check LoadBehavior and if it's missing, ignore this entry
        $loadBehaviorValue = $addin.GetValue('LoadBehavior')

        if ($loadBehaviorValue) {
            $props['LoadBehavior'] = $LoadBehavior[$loadBehaviorValue]
        }
        else {
            Write-Log "Skipping $($props['ProgID']) because its LoadBehavior is null."
            continue
        }

        if ($cache.ContainsKey($props['ProgID'])) {
            Write-Log "Skipping $($props['ProgID']) because it's already found."
            continue
        }
        else {
            $cache.Add($props['ProgID'], $null)
        }

        # Try to get CLSID.
        $($clsid = ConvertTo-CLSID $props['ProgID'] -User $User -ErrorAction Continue) 2>&1 | Write-Log

        if ($clsid) {
            $props['CLSID'] = $clsid.String

            # Check InprocServer32, LocalServer32, RemoteServer32
            foreach ($comType in @('InprocServer32', 'LocalServer32', 'RemoteServer32')) {
                $comSpec = Get-ItemProperty "Registry::HKEY_CLASSES_ROOT\CLSID\$($props['CLSID'])\$comType" -ErrorAction SilentlyContinue

                if (-not $comSpec) {
                    $comSpec = Get-ItemProperty "Registry::HKEY_CLASSES_ROOT\WOW6432Node\CLSID\\$($props['CLSID'])\$comType" -ErrorAction SilentlyContinue
                }

                if (-not $comSpec) {
                    $comSpec = Get-ItemProperty "Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\ClickToRun\REGISTRY\MACHINE\Software\Classes\CLSID\$($props['CLSID'])\$comType" -ErrorAction SilentlyContinue
                }

                if (-not $comSpec) {
                    $comSpec = Get-ItemProperty "Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\ClickToRun\REGISTRY\MACHINE\Software\Classes\Wow6432Node\CLSID\$($props['CLSID'])\$comType" -ErrorAction SilentlyContinue
                }

                if ($comSpec) {
                    $props['Location'] = $comSpec.'(default)'
                    $props['ThreadingModel'] = $comSpec.ThreadingModel
                    $props['CodeBase'] = $comSpec.CodeBase
                    break
                }
            }
        }
        elseif ($manifest = $addin.GetValue('Manifest')) {
            # A managed addin does not have CLSID. Check "Manifest" instead.
            $props['Location'] = $manifest
            Write-Log "Manifest is found. This is a VSTO addin."
        }
        else {
            # If both CLSID & Manifest are missing, ignore this entry.
            continue
        }

        # ToDo: text might get garbled in DBCS environment.
        $props['Description'] = $addin.GetValue('Description')
        $props['FriendlyName'] = $addin.GetValue('FriendlyName')

        [PSCustomObject]$props
    }

    # Close all the keys
    $addinKeys | ForEach-Object { $_.Close() }
}

function Get-ClickToRunConfiguration {
    [CmdletBinding()]
    param()

    Get-ItemProperty Registry::HKLM\SOFTWARE\Microsoft\Office\ClickToRun\Configuration
}

function Get-WebView2 {
    [CmdletBinding(PositionalBinding = $false)]
    param (
    )

    @(
        'Registry::HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Microsoft\EdgeUpdate\Clients\{F3017226-FE2A-4295-8BDF-00C3A9A7E4C5}'
        'Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\EdgeUpdate\Clients\{F3017226-FE2A-4295-8BDF-00C3A9A7E4C5}'
    ) | ForEach-Object { Get-ItemProperty $_ -ErrorAction SilentlyContinue }
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
        New-Item -ItemType Directory -Path $Path -ErrorAction Stop | Out-Null
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
        New-Item $Path -ItemType Directory -ErrorAction Stop | Out-Null
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
    Out-File -FilePath $configFile -InputObject $counters -Force -Encoding "ascii"

    $filePath = Join-Path $Path $FileName
    Write-Log "Staring PerfCounter. Mode: $LogFileMode, IntervalSecond: $IntervalSecond, MaxFileSizeMB: $MaxFileSizeMB, FilePath: $filePath"

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
Get processes and its user (only for Outlook.exe).
PowerShell 4's Get-Process has -IncludeUserName, but I'm using WMI here for now.
#>
function Save-Process {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        $Path
    )

    if (-not (Test-Path $Path)) {
        New-Item -ItemType Directory -Path $Path -ErrorAction Stop | Out-Null
    }

    Write-Log "Saving Win32_Process"
    Get-WmiObject -Class Win32_Process | ForEach-Object {
        if ($_.ProcessName -eq 'Outlook.exe') {
            $owner = $_.GetOwner()
            $_ | Add-Member -MemberType NoteProperty -Name 'User' -Value "$($owner.Domain)\$($owner.User)"
        }
        $_
    } | Export-Clixml -Path (Join-Path $Path "Win32_Process_$(Get-Date -Format "yyyyMMdd_HHmmss").xml")
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
                $message = "Skipped because the current script ($Version) is newer than GitHub's latest release ($($release.name))."
            }
            else {
                Write-Verbose "Downloading the latest script."
                $response = Invoke-Command {
                    # Suppress progress on Invoke-WebRequest.
                    $ProgressPreference = "SilentlyContinue"
                    Invoke-WebRequest -Uri $release.assets.browser_download_url -UseDefaultCredentials -UseBasicParsing
                }

                # Rename the current script and replace with the latest one.
                $newName = [IO.Path]::GetFileNameWithoutExtension($PSCommandPath) + "_" + $Version + [IO.Path]::GetExtension($PSCommandPath)
                if (Test-Path (Join-Path ([IO.Path]::GetDirectoryName($PSCommandPath)) $newName)) {
                    $newName = [IO.Path]::GetFileNameWithoutExtension($PSCommandPath) + "_" + $Version + [IO.Path]::GetRandomFileName() + [IO.Path]::GetExtension($PSCommandPath)
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
        Success = $autoUpdateSuccess
        Message = $message
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
            New-Item $Path -ItemType Directory -ErrorAction Stop | Out-Null
        }
        $Path = Resolve-Path $Path
    }

    if ($Path) {
        # For some reason, if the path contains a space & is double-quoted & ends with a backslash, wpr fails with "Invalid temporary trace directory. Error code: 0xc5586004"
        # Make sure to remove the last backslash.
        if ($Path.EndsWith('\')) {
            $Path = $Path.Remove($Path.Length - 1)
        }

        $errs = $(wpr.exe -start GeneralProfile -start CPU -start Network -filemode -RecordTempTo $Path | Out-Null) 2>&1
    }
    else {
        $errs = $(wpr.exe -start GeneralProfile -start CPU -start Network -filemode | Out-Null) 2>&1
    }

    $errorMsg = $errs | ForEach-Object { $msg = $_.Exception.Message.Trim(); if ($msg) { $msg } }

    if ($LASTEXITCODE -ne 0) {
        Write-Error "wpr failed to start. LASTEXITCODE: 0x$('{0:x}' -f $LASTEXITCODE).`n$errorMsg"
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
        New-Item $Path -ItemType Directory -ErrorAction Stop | Out-Null
    }
    $Path = Resolve-Path $Path

    $filePath = Join-Path $Path $FileName
    $(wpr.exe -stop $filePath -skipPdbGen) 2>&1 | Out-Null

    # If "Invalid command syntax", retry without -skipPdbGen because the option might not be avaiable (e.g. W2019)
    if ($LASTEXITCODE -eq 0xc5600602) {
        $(wpr.exe -stop $filePath) 2>&1 | Out-Null
    }

    if ($LASTEXITCODE -ne 0) {
        Write-Error "wpr failed to stop. LASTEXITCODE: 0x$('{0:x}' -f $LASTEXITCODE)."
    }
}

function Get-IMProvider {
    [CmdletBinding(PositionalBinding = $false)]
    param(
        $User
    )

    $root = Get-UserRegistryRoot $User
    $defaultIMApp = Get-ItemProperty (Join-Path $root 'SOFTWARE\IM Providers') -Name 'DefaultIMApp' -ErrorAction SilentlyContinue | Select-Object -ExpandProperty 'DefaultIMApp'
    if (-not $defaultIMApp) {
        Write-Error "Failed to get DefaultIMApp."
        return
    }

    [Guid]$clsid = [Guid]::Empty

    switch ($defaultIMApp) {
        'Teams' { $clsid = '00425F68-FFC1-445F-8EDF-EF78B84BA1C7'; break }
        'Lync' { $clsid = 'A0651028-BA7A-4D71-877F-12E0175A5806'; break }
    }

    if ($clsid -eq [Guid]::Empty) {
        Write-Error "Failed to get CLSID of DefaultIMApp $defaultIMApp."
        return
    }

    $isRunning = $false
    $process = Get-Process -Name $defaultIMApp -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($process) {
        $isRunning = $true
        $process.Dispose()
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
            Write-Error -Message "Failed to create an instance of $defaultIMApp (CLSID: {$clsid}).`n$($_.Exception.Message)" -Exception $_.Exception
        }
        elseif ($pIUCOfficeIntegration -eq [IntPtr]::Zero) {
            Write-Error -Message "Failed to obtain IUCOfficeIntegration interface.`n$($_.Exception.Message)" -Exception $_.Exception
        }
        else {
            Write-Error -ErrorRecord $_
        }
    }
    finally {
        if ($punk -ne [IntPtr]::Zero) {
            [System.Runtime.InteropServices.Marshal]::Release($punk) | Out-Null
        }

        if ($pIUCOfficeIntegration -ne [IntPtr]::Zero) {
            [System.Runtime.InteropServices.Marshal]::Release($pIUCOfficeIntegration) | Out-Null
        }

        if ($imProvider) {
            [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($imProvider) | Out-Null
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
        $netTask.Wait(-1) | Out-Null
    }

    function Await($WinRtTask, $ResultType) {
        # https://fleexlab.blogspot.com/2018/02/using-winrts-iasyncoperation-in.html
        $asTaskGeneric = ([System.WindowsRuntimeSystemExtensions].GetMethods() | Where-Object { $_.Name -eq 'AsTask' -and $_.GetParameters().Count -eq 1 -and $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncOperation`1' })[0]
        $asTask = $asTaskGeneric.MakeGenericMethod($ResultType)
        $netTask = $asTask.Invoke($null, @($WinRtTask))
        $netTask.Wait(-1) | Out-Null
        $netTask.Result
    }

    # https://docs.microsoft.com/en-us/uwp/api/windows.security.authentication.web.core.webauthenticationcoremanager.findaccountproviderasync?view=winrt-20348#Windows_Security_Authentication_Web_Core_WebAuthenticationCoreManager_FindAccountProviderAsync_System_String_
    $provider = Await ([Windows.Security.Authentication.Web.Core.WebAuthenticationCoreManager, Windows, ContentType = WindowsRuntime]::FindAccountProviderAsync('https://login.microsoft.com', 'organizations')) ([Windows.Security.Credentials.WebAccountProvider, Windows, ContentType = WindowsRuntime])

    # https://docs.microsoft.com/en-us/uwp/api/windows.security.authentication.web.core.webauthenticationcoremanager.findallaccountsasync?view=winrt-20348
    $accounts = Await ([Windows.Security.Authentication.Web.Core.WebAuthenticationCoreManager, Windows, ContentType = WindowsRuntime]::FindAllAccountsAsync($provider, $ClientId)) ([Windows.Security.Authentication.Web.Core.FindAllAccountsResult, Windows, ContentType = WindowsRuntime])

    $count = 0
    $accounts.Accounts | ForEach-Object { $count++ }

    if ($count -eq 0) {
        Write-Log "No account found."
        return
    }

    Write-Log "$count account$(if ($count -gt 1) {'s'}) found."

    foreach ($account in $accounts.Accounts) {
        $accountId = "UserName: $($account.UserName), Id: $($account.Id)"
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
    param (
        # Folder to place collected data
        [Parameter(Mandatory = $true, Position = 0)]
        $Path,
        # What to collect
        [Parameter(Mandatory = $true)]
        [ValidateSet('Outlook', 'Netsh', 'PSR', 'LDAP', 'CAPI', 'Configuration', 'Fiddler', 'TCO', 'Dump', 'CrashDump', 'HungDump', 'Procmon', 'WAM', 'WFP', 'TTD', 'Performance', 'WPR')]
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
        [Alias('SkipZip')]
        [switch]$SkipArchive,
        # AutoFlush log file.
        [switch]$AutoFlush,
        # Skip running autoupdate of this script.
        [switch]$SkipAutoUpdate,
        # PSR recycle interval in minutes.
        [int]$PsrRecycleIntervalMin = 10,
        # Target user whose configuration is collected. By default, it's the logon user (Note: Not necessarily the current user running the script).
        [string]$User,
        # Number of seconds used to detect a hung window when "HungDump" is requested in Component.
        [ValidateRange(1, [int]::MaxValue)]
        [int]$HungTimeoutSecond = 5,
        [string]$HungMonitorTarget = 'Outlook', # This is just for testing.
        [switch]$WamSignOut
    )

    # Explicitly check admin rights depending on the request.
    if ($Component -contains 'Outlook' -or $Component -contains 'Netsh' -or $Component -contains 'CAPI' -or $Component -contains 'LDAP' -or $Component -contains 'WAM' -or $Component -contains 'WPR') {
        if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
            Write-Warning "Please run as administrator."
            return
        }
    }

    if ($env:PROCESSOR_ARCHITEW6432 -and $PSVersionTable.PSVersion.Major -eq 2) {
        Write-Error "32bit PowerShell 2.0 is running on 64bit OS. Please use 64bit PowerShell."
        return
    }

    # MS Office must be installed to collect Outlook, TCO, or TTD.
    # This is just a fail fast. Start-OutlookTrace/TCOTrace fail anyway.
    if ($Component -contains 'Outlook' -or $Component -contains 'TCO' -or $Component -contains 'TTD') {
        $err = $(Get-OfficeInfo -ErrorAction Continue | Out-Null) 2>&1
        if ($err) {
            Write-Error "Component `"Outlook`" and/or `"TCO`" is specified, but installation of Microsoft Office is not found. $err"
            return
        }
    }

    if ($Component -contains 'TTD' -and -not (Get-Command 'tttracer.exe' -ErrorAction SilentlyContinue)) {
        $os = Get-WmiObject -Class 'Win32_OperatingSystem'
        Write-Error "tttracer is not available on this machine. $($os.Caption) ($($os.Version))."
        return
    }

    # If User is given, use it as the target user; Otherwise, use the logon user.
    if ($PSBoundParameters.ContainsKey('User')) {
        $targetUser = Resolve-User $User
        if (-not $targetUser) {
            return
        }
    }
    else {
        # Get logon user & save the error (cannot use Write-Log yet).
        $logonUserError = $($targetUser = Get-LogonUser) 2>&1
        # If Get-LogonUser fails for some reason (e.g. Access Denied), fall back to current user
        if (-not $targetUser) {
            $targetUser = Resolve-User ([System.Security.Principal.WindowsIdentity]::GetCurrent().Name)
        }
    }

    if (-not (Test-Path $Path -ErrorAction Stop)) {
        New-Item -ItemType Directory $Path -ErrorAction Stop | Out-Null
    }

    if (-not $SkipAutoUpdate) {
        $autoUpdate = Invoke-AutoUpdate
        if ($autoUpdate.Success) {
            $updatedSelf = Get-Command $MyInvocation.MyCommand.Name

            # Get the list of current parameters that's also available in the updated cmdlet
            $params = @{}
            $PSBoundParameters.Keys | ForEach-Object { if ($updatedSelf.Parameters.ContainsKey($_)) { $params.Add($_, $PSBoundParameters[$_]) } }

            if ($updatedSelf.Parameters.ContainsKey('SkipAutoUpdate')) {
                $params.Add('SkipAutoUpdate', $true)
            }

            & $updatedSelf @params
            return
        }
    }

    # Create a temporary folder to store data.
    $Path = Resolve-Path -LiteralPath $Path
    $tempPath = Join-Path $Path -ChildPath $([Guid]::NewGuid().ToString())
    New-Item $tempPath -ItemType directory -ErrorAction Stop | Out-Null

    # Start logging.
    Open-Log -Path (Join-Path $tempPath 'Log.txt') -AutoFlush:$AutoFlush -ErrorAction Stop
    Write-Log "Script Version: $Script:Version (Module Version $($MyInvocation.MyCommand.Module.Version.ToString()))"
    Write-Log "PSVersion: $($PSVersionTable.PSVersion); CLRVersion: $($PSVersionTable.CLRVersion)"
    Write-Log "PROCESSOR_ARCHITECTURE: $env:PROCESSOR_ARCHITECTURE; PROCESSOR_ARCHITEW6432: $env:PROCESSOR_ARCHITEW6432"
    Write-Log "Running as $([System.Security.Principal.WindowsIdentity]::GetCurrent().Name)"
    Write-Log "AutoUpdate: $(if ($SkipAutoUpdate) { 'Skipped due to SkipAutoUpdate switch' } else { $autoUpdate.Message })"
    Write-Log "Target user: $($targetUser.Name) ($($targetUser.SID))"
    Write-Log $logonUserError

    $sb = New-Object System.Text.StringBuilder
    foreach ($paramName in $PSBoundParameters.Keys) {
        $var = Get-Variable $paramName -ErrorAction SilentlyContinue
        if ($var) {
            $sb.Append("$($var.Name):$($var.Value -join ', '); ") | Out-Null
        }
    }
    Write-Log "Parameters $($sb.ToString())"

    # To use Start-Task, make sure to open runspaces first and close it when finished.
    # Currently MaxRunspaces is 7 or more because there are 7 tasks at most. 3 of them, outlookMonitorTask, psrTask, and hungMonitorTask are long running.
    Open-TaskRunspace -IncludeScriptVariables -MinRunspaces ([int]$env:NUMBER_OF_PROCESSORS) -MaxRunspaces ([math]::Max(7, (2 * [int]$env:NUMBER_OF_PROCESSORS)))
    # Open-TaskRunspace -Variables (Get-Variable 'logWriter')

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

    # Sign out of all WAM accounts.
    if ($WamSignOut) {
        Invoke-WamSignOut -Force 2>&1 | Write-Log
    }

    Write-Log "Starting traces"
    try {
        if ($Component -contains 'Configuration') {
            # Sub directories
            $ConfigDir = Join-Path $tempPath 'Configuration'
            $OSDir = Join-Path $ConfigDir 'OS'
            $OfficeDir = Join-Path $ConfigDir 'Office'
            $RegistryDir = Join-Path $ConfigDir 'Registry'
            $NetworkDir = Join-Path $ConfigDir 'Network'
            $MSIPCDir = Join-Path $ConfigDir 'MSIPC'
            $EventDir = Join-Path $ConfigDir 'EventLog'

            $activity = "Saving configuration"
            $status = "Please wait"
            Write-Progress -Activity $activity -Status $status -PercentComplete 0

            # First start tasks that might take a while.

            # MSInfo32 takes a long time. Currently disabled.
            # $msinfo32Task = Start-Task -Command 'Save-MSInfo32' -Parameters @{Path = $OSDir}

            Write-Log "Starting officeModuleInfoTask."
            $cts = New-Object System.Threading.CancellationTokenSource
            $officeModuleInfoTask = Start-Task { param($path, $token) Save-OfficeModuleInfo -Path $path -CancellationToken $token } -ArgumentList $OfficeDir, $cts.Token

            Write-Log "Starting networkInfoTask."
            $networkInfoTask = Start-Task { param($path) Save-NetworkInfo -Path $path } -ArgumentList $NetworkDir

            Write-Progress -Activity $activity -Status $status -PercentComplete 20

            Write-Log "Starting officeRegistryTask."
            $officeRegistryTask = Start-Task { param($path, $user) Save-OfficeRegistry -Path $path -User $user } -ArgumentList $RegistryDir, $targetUser

            Write-Log "Starting oSConfigurationTask."
            $oSConfigurationTask = Start-Task { param($path) Save-OSConfiguration -Path $path } -ArgumentList $OSDir
            Run-Command { param($user) Get-WinInetProxy -User $user } -ArgumentList $targetUser -Path $OSDir
            Run-Command { param($user) Get-ProxyAutoConfig -User $user } -ArgumentList $targetUser -Path $OSDir

            Write-Progress -Activity $activity -Status $status -PercentComplete 40
            Run-Command { Get-OfficeInfo } -Path $OfficeDir
            Run-Command { param($user) Get-OutlookProfile -User $user } -ArgumentList $targetUser -Path $OfficeDir
            Run-Command { param($user) Get-OutlookAddin -User $user } -ArgumentList $targetUser -Path $OfficeDir
            Run-Command { Get-ClickToRunConfiguration } -Path $OfficeDir
            Run-Command { param($user) Get-IMProvider -User $user } -ArgumentList $targetUser -Path $OfficeDir

            Write-Progress -Activity $activity -Status $status -PercentComplete 60
            Run-Command { param($user, $OfficeDir) Save-CachedAutodiscover -User $user -Path $(Join-Path $OfficeDir 'Cached AutoDiscover') } -ArgumentList $targetUser, $OfficeDir
            Run-Command { param($user, $OfficeDir) Save-DLP -User $user -Path $(Join-Path $OfficeDir 'DLP') } -ArgumentList $targetUser, $OfficeDir

            Write-Progress -Activity $activity -Status $status -PercentComplete 80
            Run-Command { param($OSDir) Save-Process -Path $OSDir } -ArgumentList $OSDir

            if ($targetUser) {
                $targetUser | Export-Clixml -Path (Join-Path $OSDir 'User.xml')
            }

            # The user might start & stop Outlook while tracing. In order to capture Outlook's instance, run a task to check Outlook.exe periodically until it finds an instance.
            Write-Log "Starting outlookMonitorTask."
            $outlookMonitorTask = Start-Task {
                param($OSDir)
                while ($true) {
                    if ($p = Get-Process -Name 'Outlook') {
                        Write-Log "outlookMonitorTask found Outlook's process (PID: $($p.Id))"
                        Save-Process -Path $OSDir
                        $p.Dispose()
                        return
                    }
                    else {
                        Start-Sleep -Seconds 3
                    }
                }
            } -ArgumentList $OSDir

            Write-Progress -Activity $activity -Status 'Done' -Completed
        }

        if ($Component -contains 'Fiddler') {
            Start-FiddlerCap -Path $Path -ErrorAction Stop | Out-Null
            $fiddlerCapStarted = $true

            Write-Warning "FiddlerCap has started. Please manually configure and start capture."
        }

        if ($Component -contains 'Netsh') {
            Write-Progress -Activity "Starting Netsh trace" -Status "Please wait" -PercentComplete -1
            # When netsh trace is run for the first time, it does not capture packets (even with "capture=yes").
            # To workaround, netsh is started and stopped immediately.
            $tempNetshName = 'netsh_test'
            Start-NetshTrace -Path (Join-Path $tempPath $tempNetshName) -FileName "$tempNetshName.etl" -RerpotMode 'None' -ErrorAction SilentlyContinue
            Stop-NetshTrace -ErrorAction SilentlyContinue
            Remove-Item (Join-Path $tempPath $tempNetshName) -Recurse -Force -ErrorAction SilentlyContinue

            Start-NetshTrace -Path (Join-Path $tempPath 'Netsh') -RerpotMode $NetshReportMode
            Write-Progress -Activity "Starting Netsh trace" -Completed
            $netshTraceStarted = $true
        }

        if ($Component -contains 'Outlook') {
            # Stop a lingering session if any.
            Stop-OutlookTrace -ErrorAction SilentlyContinue
            Start-OutlookTrace -Path (Join-Path $tempPath 'Outlook')
            $outlookTraceStarted = $true
        }

        if ($Component -contains 'PSR') {
            # Start PSR as a task and restart after some time until canceled.
            # This task creates PSR_***.mht in $psrPath. When LogFileMode is 'Circular', only files writen within the last 1 hour will be kept.
            $psrCts = New-Object System.Threading.CancellationTokenSource
            $psrPath = Join-Path $tempPath 'PSR'
            $psrStartedEvent = New-Object System.Threading.EventWaitHandle($false, [Threading.EventResetMode]::ManualReset)

            Write-Log "Starting a PSR task. PsrRecycleIntervalMin: $PsrRecycleIntervalMin"

            $psrTask = Start-Task -ScriptBlock {
                param(
                    [string]$path,
                    [System.Threading.CancellationToken]$cancelToken,
                    [int]$waitDurationMS,
                    $psrStartedEvent,
                    [bool]$circular
                )

                while ($true) {
                    Start-PSR -Path $path -FileName "PSR_$(Get-Date -f 'MMdd_HHmmss')"
                    $psrStartedEvent.Set() | Out-Null
                    $canceled = $cancelToken.WaitHandle.WaitOne($waitDurationMS)
                    Stop-PSR
                    if ($canceled) {
                        Write-Log "PSR task cancellation is acknowledged."
                        break
                    }
                    if ($circular) {
                        $removedCount = 0
                        $cutoff = [datetime]::Now.AddHours(-1)
                        Get-ChildItem $path -Filter '*.mht' | Where-Object { $_.LastWriteTime -lt $cutoff } | ForEach-Object { Remove-Item $_.FullName; $removedCount++ }
                        if ($removedCount) {
                            Write-Log "$removedCount mht files were removed because they were older than 1 hour"
                        }
                    }
                }
            } -ArgumentList $psrPath, $psrCts.Token, ($PsrRecycleIntervalMin * 60 * 1000), $psrStartedEvent, ($LogFileMode -eq 'Circular')

            $psrStartedEvent.WaitOne([System.Threading.Timeout]::InfiniteTimeSpan) | Out-Null
            $psrStarted = $true
        }

        if ($Component -contains 'LDAP') {
            Start-LDAPTrace -Path (Join-Path $tempPath 'LDAP') -TargetProcess 'Outlook.exe'
            $ldapTraceStarted = $true
        }

        if ($Component -contains 'CAPI') {
            Enable-EventLog 'Microsoft-Windows-CAPI2/Operational'
            Start-CAPITrace -Path (Join-Path $tempPath 'CAPI')
            $capiTraceStarted = $true
        }

        if ($Component -contains 'TCO') {
            Start-TCOTrace
            $tcoTraceStarted = $true
        }

        if ($Component -contains 'WAM') {
            Enable-WamEventLog -ErrorAction SilentlyContinue
            Stop-WamTrace -ErrorAction SilentlyContinue
            Start-WamTrace -Path (Join-Path $tempPath 'WAM')
            $wamTraceStarted = $true
        }

        if ($Component -contains 'Procmon') {
            $procmonResult = Start-Procmon -Path (Join-Path $tempPath 'Procmon') -ProcmonSearchPath $Path -ErrorAction Stop
            $procmonStared = $true
        }

        if ($Component -contains 'WFP') {
            $wfpJob = Start-WfpTrace -Path (Join-Path $tempPath 'WFP') -IntervalSeconds 15
            $wfpStarted = $true
        }

        if ($Component -contains 'Performance') {
            Write-Progress -Activity "Starting performance trace" -Status "Please wait" -PercentComplete -1
            Start-PerfTrace -Path (Join-Path $tempPath 'Performance')
            Write-Progress -Activity "Starting performance trace" -Completed
            $perfStarted = $true
        }

        if ($Component -contains 'WPR') {
            Start-Wpr -Path (Join-Path $tempPath 'WPR') -ErrorAction Stop
            $wprStarted = $true
        }

        if ($Component -contains 'CrashDump') {
            Add-WerDumpKey -Path (Join-Path $tempPath 'WerDump') -TargetProcess 'Outlook.exe'
            $crashDumpStarted = $true
        }

        if ($Component -contains 'Dump') {
            # Ask a user when she/he wants to save a dump file
            while ($true) {
                $userInput = Read-Host "Hit enter to save a process dump of Outlook. To quit, enter q"

                if ($userInput.ToLower() -eq 'q') {
                    break
                }

                if (-not ($process = Get-Process -Name 'Outlook' -ErrorAction SilentlyContinue)) {
                    Write-Host "Cannot find Outlook.exe. Please start Outlook." -ForegroundColor Yellow
                    continue
                }

                Write-Progress -Activity "Saving a process dump of Outlook." -Status "Please wait." -PercentComplete -1
                $dumpResult = Save-Dump -Path (Join-Path $tempPath 'Dump') -ProcessId $process.Id
                Write-Progress -Activity "Saving a process dump of Outlook." -Status "Done" -Completed
                Write-Log "Saved dump file: $($dumpResult.DumpFile)"
            }
        }

        if ($Component -contains 'HungDump') {
            $hungDumpCts = New-Object System.Threading.CancellationTokenSource
            $monitorStartedEvent = New-Object System.Threading.EventWaitHandle($false, [Threading.EventResetMode]::ManualReset)
            Write-Log "Starting hungMonitorTask. HungTimeoutSecond: $HungTimeoutSecond."

            # Save at most 10 dump files for now.
            $hungMonitorTask = Start-Task -ScriptBlock {
                param($path, $timeout, $dumpCount, $cancelToken, $name, $monitorStartedEvent)

                $monitorStartedEvent.Set() | Out-Null

                # Wait for Outlook to come live.
                while ($true) {
                    if ($cancelToken.IsCancellationRequested) {
                        return
                    }

                    $outlookProc = Get-Process -Name $name -ErrorAction SilentlyContinue

                    if ($outlookProc) {
                        break
                    }
                    Start-Sleep -Seconds 2
                }

                $id = $outlookProc.Id
                $outlookProc.Dispose()

                Write-Log "hungMonitorTask has found $name (PID $id). Starting hung window monitoring."
                Save-HungDump -Path $path -ProcessId $id -DumpCount $dumpCount -CancellationToken $cancelToken
            } -ArgumentList (Join-Path $tempPath 'HungDump'), $HungTimeoutSecond, 10, $hungDumpCts.Token, $HungMonitorTarget, $monitorStartedEvent

            $monitorStartedEvent.WaitOne([System.Threading.Timeout]::InfiniteTimeSpan) | Out-Null
            $hungDumpStarted = $true
        }

        if ($Component -contains 'TTD') {
            # If Outlook is already running, attach to it. Otherwise, start TTD with OnLaunch option and ask the user to start Outlook.
            if ($outlookProcess = Get-Process -Name 'Outlook' -ErrorAction SilentlyContinue) {
                Write-Log "TTD attaching to Outlook (PID $($outlookProcess.Id))."
                $ttd = Attach-TTD -Path (Join-Path $tempPath 'TTD')  -ProcessID  $outlookProcess.Id -ErrorAction Stop
            }
            else {
                $outlookExe = $null
                $officeInfo = Get-OfficeInfo
                $executables = @(Get-ChildItem -Path $officeInfo.InstallPath -Filter 'Outlook.exe' -File -Recurse)

                if ($executables.Count -eq 1) {
                    $outlookExe = $executables[0]
                }
                else {
                    # For ClickToRun, there might be more than one Outlook.exe; downloaded Outlook.exe under the Office installation path.
                    # e.g. C:\Program Files\Microsoft Office\Updates\Download\PackageFiles\7ABA93E3-58C2-4BEE-AB49-3438C9F29D70\root\Office16\Outlook.exe
                    # Pick the one without 'PackageFiles' in the path.
                    $outlookExe = $executables | Where-Object { $_.FullName -notlike '*PackageFiles*' } | Select-Object -First 1
                }

                # Start monitoring launch of Outlook
                $ttd = Start-TTD -Path (Join-Path $tempPath 'TTD') -Executable $outlookExe.FullName -OnLaunch -ErrorAction Stop
                Write-Host "Please start Outlook now." -ForegroundColor Green

                while (-not ($outlookProcess = Get-Process -Name 'Outlook' -ErrorAction SilentlyContinue)) {
                    Start-Sleep -Seconds 3
                }

                Write-Log "Outlook.exe (PID: $($outlookProcess.Id)) detected"
                $outlookProcess.EnableRaisingEvents = $true
                $ttd.TargetProcess = $outlookProcess

                Write-Host "Outlook has started (PID: $($ttd.TargetProcess.Id)). It might take some time for Outlook to appear." -ForegroundColor Green
            }

            $ttdStarted = $true
        }

        if ($netshTraceStarted -or $outlookTraceStarted -or $psrStarted -or $ldapTraceStarted -or $capiTraceStarted -or $tcoTraceStarted -or $fiddlerCapStarted -or $crashDumpStarted -or $procmonStared -or $wamTraceStarted -or $wfpStarted -or $ttdStarted -or $perfStarted -or $hungDumpStarted -or $wprStarted) {
            Write-Log "Waiting for the user to stop"
            Read-Host 'Hit enter to stop'

            # To allow Write-Host from another runspace, don't block the host by Read-Host here.
            # Write-Host "Hit enter to stop:"
            # while (-not $Host.UI.RawUI.KeyAvailable -or $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown").VirtualKeyCode -ne 13) {
            #     Start-Sleep 1
            # }
        }
    }
    catch {
        # Log & save the exception so that I can analyze later. Then rethrow.
        Write-Log "Exception occured. $_"
        $_ | Export-CliXml (Join-Path $tempPath 'Exception.xml')
        throw
    }
    finally {
        Write-Log "Stopping traces"
        Write-Progress -Activity 'Stopping traces' -Status "Please wait." -PercentComplete -1

        if ($ttdStarted) {
            Write-Progress -Activity 'Stopping TTD trace' -Status "Please wait." -PercentComplete -1
            $($stopResult = Stop-TTD $ttd) 2>&1 | Write-Log
            Write-Log "Stop-TTD Message: $($stopResult.Message)"

            # Outlook might be holding the TTD file.
            # Tell the user to stop Outlook and wait for the process to shutdown.
            if (-not $ttd.TargetProcess.HasExited) {
                Write-Log "Waiting for the user to shutdown Outlook."
                Write-Host "TTD Tracing is stopped. Please shutdown Outlook" -ForegroundColor Green
                Write-Progress -Activity 'Stopping traces' -Status "Please shutdown Outlook." -PercentComplete -1

                # Wait for Outlook to be stopped. Nudge the user once in a while.
                while ($true) {
                    $timeout = $(Wait-Process -InputObject $ttd.TargetProcess -Timeout 30 -ErrorAction Continue) 2>&1
                    if ($timeout) {
                        Write-Host "Please shutdown Outlook." -ForegroundColor Yellow
                    }
                    else {
                        Write-Host "Outlook is closed. Moving on." -ForegroundColor Green
                        break
                    }
                }
            }

            Write-Log "Outlook ExitCode: $($ttd.TargetProcess.ExitCode), ExitTime: $(if ($ttd.TargetProcess.ExitTime) {$ttd.TargetProcess.ExitTime.ToString('o')}), ElapsedTime: $($ttd.TargetProcess.ExitTime - $ttd.TargetProcess.StartTime)"
            $ttd.Dispose()
            Write-Progress -Activity 'Stopping TTD trace' -Status "Done" -Completed
        }

        if ($netshTraceStarted) {
            Stop-NetshTrace
        }

        if ($outlookTraceStarted) {
            Stop-OutlookTrace
        }

        if ($ldapTraceStarted) {
            Stop-LDAPTrace -TargetProcess 'Outlook.exe'
        }

        if ($capiTraceStarted) {
            Disable-EventLog 'Microsoft-Windows-CAPI2/Operational'
            Stop-CAPITrace
        }

        if ($tcoTraceStarted) {
            Stop-TcoTrace -Path (Join-Path $tempPath 'TCO')
        }

        if ($wamTraceStarted) {
            Disable-WamEventLog -ErrorAction SilentlyContinue
            Stop-WamTrace
        }

        if ($procmonStared) {
            Stop-Procmon
            # Remove procmon
            # if ($procmonResult -and $procmonResult.ProcmonZipDownloaded) {
            #     Remove-Item $procmonResult.ProcmonFolderPath -Force -Recurse
            # }
        }

        if ($wfpStarted) {
            Stop-WfpTrace $wfpJob
        }

        if ($perfStarted) {
            Stop-PerfTrace
        }

        if ($wprStarted) {
            Write-Progress -Activity 'Stopping WPR' -Status "Please wait." -PercentComplete -1
            Stop-Wpr -Path (Join-Path $tempPath 'WPR')
            Write-Progress -Activity 'Stopping WPR' -Completed
        }

        if ($hungDumpStarted) {
            $hungDumpCts.Cancel()
            $(Receive-Task $hungMonitorTask -AutoRemoveTask) 2>&1 | Write-Log
        }

        if ($crashDumpStarted) {
            Remove-WerDumpKey -TargetProcess 'Outlook.exe'
        }

        if ($fiddlerCapStarted) {
            Write-Warning "Please stop FiddlerCap and save the capture manually."
        }

        if ($psrStarted) {
            $psrCts.Cancel()
            $(Receive-Task $psrTask -AutoRemoveTask) 2>&1 | Write-Log
        }

        Write-Progress -Activity 'Stopping traces' -Status "Please wait." -Completed

        # Wait for the tasks started earlier and save the event logs
        if ($Component -contains 'Configuration') {
            if ($outlookMonitorTask) {
                # This task just tries to save Outlook process's info. No need to wait or receive.
                $(Remove-Task $outlookMonitorTask) 2>&1 | Write-Log
            }

            Write-Progress -Activity 'Saving event logs.' -Status 'Please wait.' -PercentComplete -1
            $(Save-EventLog -Path $EventDir) 2>&1 | Write-Log
            Run-Command { param($user, $MSIPCDir) Save-MSIPC -Path $MSIPCDir -User $user } -ArgumentList $targetUser, $MSIPCDir
            Write-Progress -Activity 'Saving event logs.' -Status 'Please wait.' -Completed

            if ($oSConfigurationTask) {
                Write-Progress -Activity 'Saving OS configuration' -Status "Please wait." -PercentComplete -1
                $($oSConfigurationTask | Receive-Task -AutoRemoveTask) 2>&1 | Write-Log
                Write-Progress -Activity 'Saving OS configuration' -Status "Please wait." -Completed
                Write-Log "oSConfigurationTask is complete."
            }

            if ($officeRegistryTask) {
                Write-Progress -Activity 'Saving Office Registry' -Status "Please wait." -PercentComplete -1
                $($officeRegistryTask | Receive-Task -AutoRemoveTask) 2>&1 | Write-Log
                Write-Progress -Activity 'Saving Office Registry' -Status "Please wait." -Completed
                Write-Log "officeRegistryTask is complete."
            }

            if ($networkInfoTask) {
                Write-Progress -Activity 'Saving network info' -Status "Please wait." -PercentComplete -1
                $($networkInfoTask | Receive-Task -AutoRemoveTask) 2>&1 | Write-Log
                Write-Progress -Activity 'Saving network info' -Status "Please wait." -Completed
                Write-Log "networkInfoTask is complete."
            }

            if ($officeModuleInfoTask) {
                [TimeSpan]$timeout = [TimeSpan]::FromSeconds(30)
                Write-Progress -Activity 'Saving Office module info' -Status "Please wait up to $timeout" -PercentComplete -1

                if (Wait-Task $officeModuleInfoTask -Timeout $timeout) {
                    Write-Log "officeModuleInfoTask is complete before timeout."
                }
                else {
                    Write-Log "officeModuleInfoTask timed out after $($timeout.TotalSeconds) seconds. Task will be canceled."
                    $cts.Cancel()
                }

                $($officeModuleInfoTask | Receive-Task -AutoRemoveTask) 2>&1 | Write-Log
                Write-Progress -Activity 'Saving Office module info' -Status 'Please wait.' -Completed
                Write-Log "officeRegistryTask is complete."
            }

            if ($msinfo32Task) {
                Write-Progress -Activity 'Saving MSInfo32' -Status 'Please wait.' -PercentComplete -1
                $($msinfo32Task | Receive-Task -AutoRemoveTask) 2>&1 | Write-Log
                Write-Progress -Activity 'Saving MSInfo32' -Status 'Please wait.' -Completed
            }

            # Save process list again after traces
            if ($Component.Count -gt 1) {
                Run-Command { param($OSDir) Save-Process -Path $OSDir } -ArgumentList $OSDir
            }
        }

        Close-TaskRunspace
        Close-Log
    }

    $archiveName = "Outlook_$($env:COMPUTERNAME)_$(Get-Date -Format "yyyyMMdd_HHmmss")"

    if ($SkipArchive) {
        Rename-Item -LiteralPath $tempPath -NewName $archiveName
        return
    }

    $archive = Compress-Folder -Path $tempPath -Destination $Path -ArchiveType $ArchiveType -ErrorAction Stop
    Rename-Item $archive.ArchivePath -NewName "$archiveName$([IO.Path]::GetExtension($archive.ArchivePath))"

    if (Test-Path $tempPath) {
        # Removing temp files might take a while. Do it in a background.
        $job = Start-Job -ScriptBlock {
            Remove-Item $using:tempPath -Recurse -Force
        }
        Write-Verbose "Temporary folder `"$tempPath`" will be removed by a background job (Job ID: $($job.Id))"
    }

    Write-Host "The collected data is `"$(Join-Path $Path "$archiveName$([IO.Path]::GetExtension($archive.ArchivePath))")`"" -ForegroundColor Green
    Invoke-Item $Path
}

# Configure Export-Clixml & Out-File to use UTF8 by default.
if ($PSDefaultParameterValues -ne $null -and -not $PSDefaultParameterValues.Contains("Export-CliXml:Encoding")) {
    $PSDefaultParameterValues.Add("Export-Clixml:Encoding", 'UTF8')
}

if ($PSDefaultParameterValues -ne $null -and -not $PSDefaultParameterValues.Contains("Out-File:Encoding")) {
    $PSDefaultParameterValues.Add("Out-File:Encoding", 'utf8')
}

# Add type for Win32 interop
if (-not ('Win32.Kernel32' -as [type])) {
    Add-Type -TypeDefinition $Win32Interop
}

# Save this module path ("...\OutlookTrace.psm1") so that functions can easily find it when running in other runspaces.
$Script:MyModulePath = $PSCommandPath

Export-ModuleMember -Function Start-WamTrace, Stop-WamTrace, Start-OutlookTrace, Stop-OutlookTrace, Start-NetshTrace, Stop-NetshTrace, Start-PSR, Stop-PSR, Save-EventLog, Get-InstalledUpdate, Save-OfficeRegistry, Get-ProxySetting, Get-WinInetProxy, Get-WinHttpDefaultProxy, Get-ProxyAutoConfig, Save-OSConfiguration, Get-NLMConnectivity, Get-WSCAntivirus, Save-CachedAutodiscover, Remove-CachedAutodiscover, Start-LdapTrace, Stop-LdapTrace, Save-OfficeModuleInfo, Save-MSInfo32, Start-CAPITrace, Stop-CapiTrace, Start-FiddlerCap, Start-Procmon, Stop-Procmon, Start-TcoTrace, Stop-TcoTrace, Get-OfficeInfo, Add-WerDumpKey, Remove-WerDumpKey, Start-WfpTrace, Stop-WfpTrace, Save-Dump, Save-HungDump, Save-MSIPC, Get-EtwSession, Stop-EtwSession, Get-Token, Test-Autodiscover, Get-LogonUser, Get-JoinInformation, Get-OutlookProfile, Get-OutlookAddin, Get-ClickToRunConfiguration, Get-WebView2, Get-DeviceJoinStatus, Save-NetworkInfo, Start-TTD, Stop-TTD, Attach-TTD, Start-PerfTrace, Stop-PerfTrace, Start-Wpr, Stop-Wpr, Get-IMProvider, Get-MeteredNetworkCost, Save-DLP, Invoke-WamSignOut, Collect-OutlookInfo