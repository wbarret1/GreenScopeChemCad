using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
// using System.Threading.Tasks;

namespace GreenScopeChemCad
{

    class VBServerWrapper : IDisposable
    {
        //CHEMCAD.ICHEMCADVBServer p_IServer;
        IDispatch p_IDispatch;
        StreamComponent[] m_Components;
        int[] m_ComponentIDs;
        string[] m_ComponentNames;

        const int LOCALE_SYSTEM_DEFAULT = 2048;
        const int DISPATCH_METHOD = 0x1;
        const int DISPATCH_PROPERTYGET = 0x2;
        const int DISPATCH_PROPERTYPUT = 0x4;
        const int SizeOfNativeVariant = 16;
        const int DISPID_PROPERTYPUT = -3;

        //[System.Runtime.InteropServices.DllImport("ole32.Dll")]
        //static public extern uint CoCreateInstance(ref Guid clsid,
        //   [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.IUnknown)] object inner,
        //   uint context,
        //   ref Guid uuid,
        //   [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.IUnknown)] out object rReturnedComObject);

        [System.Security.Permissions.EnvironmentPermission(System.Security.Permissions.SecurityAction.Assert)]
        public VBServerWrapper()
        {
            //AppDomain domain = AppDomain.CurrentDomain;
            //Guid chemCadIID = new Guid("2B03C5E1-25B6-11D4-BBD3-0050DACD255C");
            //Guid iDispatchIID = new Guid("00020400-0000-0000-C000-000000000046");
            //Guid IID_IUnknown = new Guid("00000000-0000-0000-C000-000000000046");
            //object obj = null;
            //uint CLSCTX_LOCAL_SERVER = 0x00000004;
            //uint hr = CoCreateInstance(ref chemCadIID, null, CLSCTX_LOCAL_SERVER, ref IID_IUnknown, out obj);
        }

        ~VBServerWrapper()
        {
            Dispose(false);
        }

        // Flag: Has Dispose already been called? 
        bool disposed = false;

        // Public implementation of Dispose pattern callable by consumers. 
        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        // Protected implementation of Dispose pattern. 
        protected virtual void Dispose(bool disposing)
        {
            if (disposed)
                return;

            if (disposing)
            {
                // Free any other managed objects here. 
                //
            }

            // Free any unmanaged objects here. 
            //
            if (p_IDispatch != null)
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(p_IDispatch);
            disposed = true;
        }

        public StreamComponent[] Components
        {
            get
            {
                return m_Components;
            }

        }

        public string[] ComponentNames
        {
            get
            {
                return m_ComponentNames;
            }

        }

        public int[] ComponentIDs
        {
            get
            {
                return m_ComponentIDs;
            }
        }

        public bool LoadJob(string bstrJobPath)
        {
            try
            {
                Type chemCadType = Type.GetTypeFromProgID("CHEMCAD.VBServer");
                p_IDispatch = (IDispatch)Activator.CreateInstance(chemCadType);
            }
            catch (Exception ex)
            {
                return false;
            }
            Guid IID_NULL = new Guid("00000000-0000-0000-0000-000000000000");
            string[] rgsNames = new string[1] { "LoadJob" };
            int[] rgDispId = new int[1] { 0 };

            int hrRet = p_IDispatch.GetIDsOfNames
            (
                ref IID_NULL,
                rgsNames,
                1,
                LOCALE_SYSTEM_DEFAULT,
                rgDispId
            );
            object varResult = null;

            if (hrRet == 0)
            {
                System.Runtime.InteropServices.ComTypes.EXCEPINFO ExcepInfo = new System.Runtime.InteropServices.ComTypes.EXCEPINFO();
                UInt32 pArgErr = 0;

                IntPtr pString0 = System.Runtime.InteropServices.Marshal.StringToBSTR(bstrJobPath);
                System.Runtime.InteropServices.GCHandle handle0 = System.Runtime.InteropServices.GCHandle.Alloc(pString0, System.Runtime.InteropServices.GCHandleType.Pinned);
                Variant[] v = new Variant[1];
                v[0].vt = 0x4008;
                v[0].data01 = handle0.AddrOfPinnedObject();
                System.Runtime.InteropServices.GCHandle rgvarg = System.Runtime.InteropServices.GCHandle.Alloc(v, System.Runtime.InteropServices.GCHandleType.Pinned);

                var dispParams = new System.Runtime.InteropServices.ComTypes.DISPPARAMS()
                {
                    cArgs = 1,
                    cNamedArgs = 0,
                    rgdispidNamedArgs = IntPtr.Zero,
                    rgvarg = rgvarg.AddrOfPinnedObject()
                };

                hrRet = p_IDispatch.Invoke
                (
                    rgDispId[0],
                    ref IID_NULL,
                    LOCALE_SYSTEM_DEFAULT,
                    DISPATCH_METHOD,
                    ref dispParams,
                    ref varResult,
                    ref ExcepInfo,
                    out pArgErr
                );
                rgvarg.Free();
            }

            if (hrRet == 0)
            {
                Flowsheet sheet = this.GetFlowsheet();
                if (sheet == null)
                {
                    // There is no flowsheet available. Maybe there is no available ChemCAD license...
                    // Let the user know and tell them to check.
                    System.Windows.Forms.MessageBox.Show("The ChemCAD Flowsheet did not load. Make sure there is a license available. You may need to close ChemCAD on the local machine.", "No Flowsheet!!");
                    this.CloseSimulation();
                    return false;
                }
                StreamInfo p_StreamInfo = this.GetStreamInfo();
                int[] ids = sheet.AllStreamIDs;
                int numComps = p_StreamInfo.NumberOfComponents;
                m_Components = new StreamComponent[numComps];
                m_ComponentIDs = new int[numComps];
                m_ComponentNames = new string[numComps];
                for (int i = 0; i < numComps; i++)
                {
                    m_Components[i] = new StreamComponent(ids[0], i, this);
                    m_ComponentIDs[i] = p_StreamInfo.GetComponentIDByPos(i);
                    m_ComponentNames[i] = p_StreamInfo.GetComponentNameByPos(i);
                }
                return (bool)varResult;
            }
            return false;
        }

        public string GetJobAt(int jobIndex)
        {
            Guid IID_NULL = new Guid("00000000-0000-0000-0000-000000000000");
            string[] rgsNames = new string[1] { "GetJobAt" };
            int[] rgDispId = new int[1] { 0 };

            int hrRet = p_IDispatch.GetIDsOfNames
            (
                ref IID_NULL,
                rgsNames,
                1,
                LOCALE_SYSTEM_DEFAULT,
                rgDispId
            );
            object varResult = null;

            if (hrRet == 0)
            {
                System.Runtime.InteropServices.ComTypes.EXCEPINFO ExcepInfo = new System.Runtime.InteropServices.ComTypes.EXCEPINFO();
                UInt32 pArgErr = 0;

                Variant[] v = new Variant[1];
                v[0].vt = 0x0002;
                v[0].iVal = (short)jobIndex;
                System.Runtime.InteropServices.GCHandle rgvarg = System.Runtime.InteropServices.GCHandle.Alloc(v, System.Runtime.InteropServices.GCHandleType.Pinned);

                var dispParams = new System.Runtime.InteropServices.ComTypes.DISPPARAMS()
                {
                    cArgs = 1,
                    cNamedArgs = 0,
                    rgdispidNamedArgs = IntPtr.Zero,
                    rgvarg = rgvarg.AddrOfPinnedObject()
                };

                hrRet = p_IDispatch.Invoke
                (
                    rgDispId[0],
                    ref IID_NULL,
                    LOCALE_SYSTEM_DEFAULT,
                    DISPATCH_METHOD,
                    ref dispParams,
                    ref varResult,
                    ref ExcepInfo,
                    out pArgErr
                );
                rgvarg.Free();
            }

            if (hrRet == 0)
            {
                return (string)varResult;
            }
            return string.Empty;
        }

        public string GetWorkDir()
        {
            Guid IID_NULL = new Guid("00000000-0000-0000-0000-000000000000");
            string[] rgsNames = new string[1] { "GetWorkDir" };
            int[] rgDispId = new int[1] { 0 };

            int hrRet = p_IDispatch.GetIDsOfNames
            (
                ref IID_NULL,
                rgsNames,
                1,
                LOCALE_SYSTEM_DEFAULT,
                rgDispId
            );
            object varResult = null;

            if (hrRet == 0)
            {
                System.Runtime.InteropServices.ComTypes.EXCEPINFO ExcepInfo = new System.Runtime.InteropServices.ComTypes.EXCEPINFO();
                UInt32 pArgErr = 0;

                Variant[] v = new Variant[0];
                System.Runtime.InteropServices.GCHandle rgvarg = System.Runtime.InteropServices.GCHandle.Alloc(v, System.Runtime.InteropServices.GCHandleType.Pinned);

                var dispParams = new System.Runtime.InteropServices.ComTypes.DISPPARAMS()
                {
                    cArgs = 0,
                    cNamedArgs = 0,
                    rgdispidNamedArgs = IntPtr.Zero,
                    rgvarg = rgvarg.AddrOfPinnedObject()
                };

                hrRet = p_IDispatch.Invoke
                (
                    rgDispId[0],
                    ref IID_NULL,
                    LOCALE_SYSTEM_DEFAULT,
                    DISPATCH_METHOD,
                    ref dispParams,
                    ref varResult,
                    ref ExcepInfo,
                    out pArgErr
                );
                rgvarg.Free();
            }

            if (hrRet == 0)
            {
                return (string)varResult;
            }
            return string.Empty;
        }

        public bool TakeProcessSnapShot(string bstrJobPath)
        {
            Guid IID_NULL = new Guid("00000000-0000-0000-0000-000000000000");
            string[] rgsNames = new string[1] { "TakeProcessSnapShot" };
            int[] rgDispId = new int[1] { 0 };

            int hrRet = p_IDispatch.GetIDsOfNames
            (
                ref IID_NULL,
                rgsNames,
                1,
                LOCALE_SYSTEM_DEFAULT,
                rgDispId
            );
            object varResult = null;

            if (hrRet == 0)
            {
                System.Runtime.InteropServices.ComTypes.EXCEPINFO ExcepInfo = new System.Runtime.InteropServices.ComTypes.EXCEPINFO();
                UInt32 pArgErr = 0;

                IntPtr pString0 = System.Runtime.InteropServices.Marshal.StringToBSTR(bstrJobPath);
                System.Runtime.InteropServices.GCHandle handle0 = System.Runtime.InteropServices.GCHandle.Alloc(pString0, System.Runtime.InteropServices.GCHandleType.Pinned);
                Variant[] v = new Variant[1];
                v[0].vt = 0x4008;
                v[0].data01 = handle0.AddrOfPinnedObject();
                System.Runtime.InteropServices.GCHandle rgvarg = System.Runtime.InteropServices.GCHandle.Alloc(v, System.Runtime.InteropServices.GCHandleType.Pinned);

                var dispParams = new System.Runtime.InteropServices.ComTypes.DISPPARAMS()
                {
                    cArgs = 1,
                    cNamedArgs = 0,
                    rgdispidNamedArgs = IntPtr.Zero,
                    rgvarg = rgvarg.AddrOfPinnedObject()
                };

                hrRet = p_IDispatch.Invoke
                (
                    rgDispId[0],
                    ref IID_NULL,
                    LOCALE_SYSTEM_DEFAULT,
                    DISPATCH_METHOD,
                    ref dispParams,
                    ref varResult,
                    ref ExcepInfo,
                    out pArgErr
                );
                rgvarg.Free();
            }

            if (hrRet == 0)
            {
                //return (bool)varResult;
            }
            return false;
        }

        public bool RunJob()
        {
            Guid IID_NULL = new Guid("00000000-0000-0000-0000-000000000000");
            string[] rgsNames = new string[1] { "RunJob" };
            int[] rgDispId = new int[1] { 0 };

            int hrRet = p_IDispatch.GetIDsOfNames
            (
                ref IID_NULL,
                rgsNames,
                1,
                LOCALE_SYSTEM_DEFAULT,
                rgDispId
            );
            object varResult = null;

            if (hrRet == 0)
            {
                System.Runtime.InteropServices.ComTypes.EXCEPINFO ExcepInfo = new System.Runtime.InteropServices.ComTypes.EXCEPINFO();
                UInt32 pArgErr = 0;

                Variant[] v = new Variant[0];
                System.Runtime.InteropServices.GCHandle rgvarg = System.Runtime.InteropServices.GCHandle.Alloc(v, System.Runtime.InteropServices.GCHandleType.Pinned);

                var dispParams = new System.Runtime.InteropServices.ComTypes.DISPPARAMS()
                {
                    cArgs = 0,
                    cNamedArgs = 0,
                    rgdispidNamedArgs = IntPtr.Zero,
                    rgvarg = rgvarg.AddrOfPinnedObject()
                };

                hrRet = p_IDispatch.Invoke
                (
                    rgDispId[0],
                    ref IID_NULL,
                    LOCALE_SYSTEM_DEFAULT,
                    DISPATCH_METHOD,
                    ref dispParams,
                    ref varResult,
                    ref ExcepInfo,
                    out pArgErr
                );
                rgvarg.Free();
            }

            if (hrRet == 0)
            {
                return (bool)varResult;
            }
            return false;
        }

        //public sbyte RunJob()
        //{
        //    return p_IServer.RunJob();
        //}

        public Flowsheet GetFlowsheet()
        {
            Guid IID_NULL = new Guid("00000000-0000-0000-0000-000000000000");
            string[] rgsNames = new string[1] { "GetFlowsheet" };
            int[] rgDispId = new int[1] { 0 };

            int hrRet = p_IDispatch.GetIDsOfNames
            (
                ref IID_NULL,
                rgsNames,
                1,
                LOCALE_SYSTEM_DEFAULT,
                rgDispId
            );
            object varResult = null;

            if (hrRet == 0)
            {
                System.Runtime.InteropServices.ComTypes.EXCEPINFO ExcepInfo = new System.Runtime.InteropServices.ComTypes.EXCEPINFO();
                UInt32 pArgErr = 0;

                Variant[] v = new Variant[0];
                System.Runtime.InteropServices.GCHandle rgvarg = System.Runtime.InteropServices.GCHandle.Alloc(v, System.Runtime.InteropServices.GCHandleType.Pinned);

                var dispParams = new System.Runtime.InteropServices.ComTypes.DISPPARAMS()
                {
                    cArgs = 0,
                    cNamedArgs = 0,
                    rgdispidNamedArgs = IntPtr.Zero,
                    rgvarg = rgvarg.AddrOfPinnedObject()
                };

                hrRet = p_IDispatch.Invoke
                (
                    rgDispId[0],
                    ref IID_NULL,
                    LOCALE_SYSTEM_DEFAULT,
                    DISPATCH_METHOD,
                    ref dispParams,
                    ref varResult,
                    ref ExcepInfo,
                    out pArgErr
                );
                rgvarg.Free();
            }

            if (hrRet == 0)
            {
                return new Flowsheet((IDispatch)varResult);
            }
            return null;
        }

        public UnitOperationInfo GetUnitOpInfo()
        {
            Guid IID_NULL = new Guid("00000000-0000-0000-0000-000000000000");
            string[] rgsNames = new string[1] { "GetUnitOpInfo" };
            int[] rgDispId = new int[1] { 0 };

            int hrRet = p_IDispatch.GetIDsOfNames
            (
                ref IID_NULL,
                rgsNames,
                1,
                LOCALE_SYSTEM_DEFAULT,
                rgDispId
            );
            object varResult = null;

            if (hrRet == 0)
            {
                System.Runtime.InteropServices.ComTypes.EXCEPINFO ExcepInfo = new System.Runtime.InteropServices.ComTypes.EXCEPINFO();
                UInt32 pArgErr = 0;

                Variant[] v = new Variant[0];
                System.Runtime.InteropServices.GCHandle rgvarg = System.Runtime.InteropServices.GCHandle.Alloc(v, System.Runtime.InteropServices.GCHandleType.Pinned);

                var dispParams = new System.Runtime.InteropServices.ComTypes.DISPPARAMS()
                {
                    cArgs = 0,
                    cNamedArgs = 0,
                    rgdispidNamedArgs = IntPtr.Zero,
                    rgvarg = rgvarg.AddrOfPinnedObject()
                };

                hrRet = p_IDispatch.Invoke
                (
                    rgDispId[0],
                    ref IID_NULL,
                    LOCALE_SYSTEM_DEFAULT,
                    DISPATCH_METHOD,
                    ref dispParams,
                    ref varResult,
                    ref ExcepInfo,
                    out pArgErr
                );
                rgvarg.Free();
            }

            if (hrRet == 0)
            {
                return new UnitOperationInfo((IDispatch)varResult);
            }
            return null;
      }

        public UnitOpSpecUnitConversion GetUnitOpSpecUnitConversion()
        {
            Guid IID_NULL = new Guid("00000000-0000-0000-0000-000000000000");
            string[] rgsNames = new string[1] { "GetUnitOpSpecUnitConversion" };
            int[] rgDispId = new int[1] { 0 };

            int hrRet = p_IDispatch.GetIDsOfNames
            (
                ref IID_NULL,
                rgsNames,
                1,
                LOCALE_SYSTEM_DEFAULT,
                rgDispId
            );
            object varResult = null;

            if (hrRet == 0)
            {
                System.Runtime.InteropServices.ComTypes.EXCEPINFO ExcepInfo = new System.Runtime.InteropServices.ComTypes.EXCEPINFO();
                UInt32 pArgErr = 0;

                Variant[] v = new Variant[0];
                System.Runtime.InteropServices.GCHandle rgvarg = System.Runtime.InteropServices.GCHandle.Alloc(v, System.Runtime.InteropServices.GCHandleType.Pinned);

                var dispParams = new System.Runtime.InteropServices.ComTypes.DISPPARAMS()
                {
                    cArgs = 0,
                    cNamedArgs = 0,
                    rgdispidNamedArgs = IntPtr.Zero,
                    rgvarg = rgvarg.AddrOfPinnedObject()
                };

                hrRet = p_IDispatch.Invoke
                (
                    rgDispId[0],
                    ref IID_NULL,
                    LOCALE_SYSTEM_DEFAULT,
                    DISPATCH_METHOD,
                    ref dispParams,
                    ref varResult,
                    ref ExcepInfo,
                    out pArgErr
                );
                rgvarg.Free();
            }

            if (hrRet == 0)
            {
                return new UnitOpSpecUnitConversion((IDispatch)varResult);
            }
            return null;
        }

        public StreamInfo GetStreamInfo()
        {
            Guid IID_NULL = new Guid("00000000-0000-0000-0000-000000000000");
            string[] rgsNames = new string[1] { "GetStreamInfo" };
            int[] rgDispId = new int[1] { 0 };

            int hrRet = p_IDispatch.GetIDsOfNames
            (
                ref IID_NULL,
                rgsNames,
                1,
                LOCALE_SYSTEM_DEFAULT,
                rgDispId
            );
            object varResult = null;

            if (hrRet == 0)
            {
                System.Runtime.InteropServices.ComTypes.EXCEPINFO ExcepInfo = new System.Runtime.InteropServices.ComTypes.EXCEPINFO();
                UInt32 pArgErr = 0;

                Variant[] v = new Variant[0];
                System.Runtime.InteropServices.GCHandle rgvarg = System.Runtime.InteropServices.GCHandle.Alloc(v, System.Runtime.InteropServices.GCHandleType.Pinned);

                var dispParams = new System.Runtime.InteropServices.ComTypes.DISPPARAMS()
                {
                    cArgs = 0,
                    cNamedArgs = 0,
                    rgdispidNamedArgs = IntPtr.Zero,
                    rgvarg = rgvarg.AddrOfPinnedObject()
                };

                hrRet = p_IDispatch.Invoke
                (
                    rgDispId[0],
                    ref IID_NULL,
                    LOCALE_SYSTEM_DEFAULT,
                    DISPATCH_METHOD,
                    ref dispParams,
                    ref varResult,
                    ref ExcepInfo,
                    out pArgErr
                );
                rgvarg.Free();
            }

            if (hrRet == 0)
            {
                return new StreamInfo((IDispatch)varResult);
            }
            return null;
        }

        //public object GetStreamUnitConversion()
        //{
        //    return p_IServer.GetStreamUnitConversion();
        //}

        public StreamProperty GetStreamProperty()
        {
            Guid IID_NULL = new Guid("00000000-0000-0000-0000-000000000000");
            string[] rgsNames = new string[1] { "GetStreamProperty" };
            int[] rgDispId = new int[1] { 0 };

            int hrRet = p_IDispatch.GetIDsOfNames
            (
                ref IID_NULL,
                rgsNames,
                1,
                LOCALE_SYSTEM_DEFAULT,
                rgDispId
            );
            object varResult = null;

            if (hrRet == 0)
            {
                System.Runtime.InteropServices.ComTypes.EXCEPINFO ExcepInfo = new System.Runtime.InteropServices.ComTypes.EXCEPINFO();
                UInt32 pArgErr = 0;

                Variant[] v = new Variant[0];
                System.Runtime.InteropServices.GCHandle rgvarg = System.Runtime.InteropServices.GCHandle.Alloc(v, System.Runtime.InteropServices.GCHandleType.Pinned);

                var dispParams = new System.Runtime.InteropServices.ComTypes.DISPPARAMS()
                {
                    cArgs = 0,
                    cNamedArgs = 0,
                    rgdispidNamedArgs = IntPtr.Zero,
                    rgvarg = rgvarg.AddrOfPinnedObject()
                };

                hrRet = p_IDispatch.Invoke
                (
                    rgDispId[0],
                    ref IID_NULL,
                    LOCALE_SYSTEM_DEFAULT,
                    DISPATCH_METHOD,
                    ref dispParams,
                    ref varResult,
                    ref ExcepInfo,
                    out pArgErr
                );
                rgvarg.Free();
            }

            if (hrRet == 0)
            {
                return new StreamProperty((IDispatch)varResult);
            }
            return null;
        }

        //public object GetKValues()
        //{
        //    return p_IServer.GetKValues();
        //}

        //public object GetEnthalpy()
        //{
        //    return p_IServer.GetEnthalpy();
        //}

        //public object GetFlash()
        //{
        //    return p_IServer.GetFlash();
        //}

        //public object GetEngUnitConversion()
        //{
        //    return p_IServer.GetEngUnitConversion();
        //}

        public CompPPData GetCompPPData()
        {
            Guid IID_NULL = new Guid("00000000-0000-0000-0000-000000000000");
            string[] rgsNames = new string[1] { "GetCompPPData" };
            int[] rgDispId = new int[1] { 0 };

            int hrRet = p_IDispatch.GetIDsOfNames
            (
                ref IID_NULL,
                rgsNames,
                1,
                LOCALE_SYSTEM_DEFAULT,
                rgDispId
            );
            object varResult = null;

            if (hrRet == 0)
            {
                System.Runtime.InteropServices.ComTypes.EXCEPINFO ExcepInfo = new System.Runtime.InteropServices.ComTypes.EXCEPINFO();
                UInt32 pArgErr = 0;

                Variant[] v = new Variant[0];
                System.Runtime.InteropServices.GCHandle rgvarg = System.Runtime.InteropServices.GCHandle.Alloc(v, System.Runtime.InteropServices.GCHandleType.Pinned);

                var dispParams = new System.Runtime.InteropServices.ComTypes.DISPPARAMS()
                {
                    cArgs = 0,
                    cNamedArgs = 0,
                    rgdispidNamedArgs = IntPtr.Zero,
                    rgvarg = rgvarg.AddrOfPinnedObject()
                };

                hrRet = p_IDispatch.Invoke
                (
                    rgDispId[0],
                    ref IID_NULL,
                    LOCALE_SYSTEM_DEFAULT,
                    DISPATCH_METHOD,
                    ref dispParams,
                    ref varResult,
                    ref ExcepInfo,
                    out pArgErr
                );
                rgvarg.Free();
            }

            if (hrRet == 0)
            {
                return new CompPPData((IDispatch)varResult);
            }
            return null;
        }

        public object GetFlash()
        {
            Guid IID_NULL = new Guid("00000000-0000-0000-0000-000000000000");
            string[] rgsNames = new string[1] { "GetFlash" };
            int[] rgDispId = new int[1] { 0 };

            int hrRet = p_IDispatch.GetIDsOfNames
            (
                ref IID_NULL,
                rgsNames,
                1,
                LOCALE_SYSTEM_DEFAULT,
                rgDispId
            );
            object varResult = null;

            if (hrRet == 0)
            {
                System.Runtime.InteropServices.ComTypes.EXCEPINFO ExcepInfo = new System.Runtime.InteropServices.ComTypes.EXCEPINFO();
                UInt32 pArgErr = 0;

                Variant[] v = new Variant[0];
                System.Runtime.InteropServices.GCHandle rgvarg = System.Runtime.InteropServices.GCHandle.Alloc(v, System.Runtime.InteropServices.GCHandleType.Pinned);

                var dispParams = new System.Runtime.InteropServices.ComTypes.DISPPARAMS()
                {
                    cArgs = 0,
                    cNamedArgs = 0,
                    rgdispidNamedArgs = IntPtr.Zero,
                    rgvarg = rgvarg.AddrOfPinnedObject()
                };

                hrRet = p_IDispatch.Invoke
                (
                    rgDispId[0],
                    ref IID_NULL,
                    LOCALE_SYSTEM_DEFAULT,
                    DISPATCH_METHOD,
                    ref dispParams,
                    ref varResult,
                    ref ExcepInfo,
                    out pArgErr
                );
                rgvarg.Free();
            }

            if (hrRet == 0)
            {
                return (IDispatch)varResult;
            }
            return null;
        }

        //public string GetWorkDir()
        //{
        //    return p_IServer.GetWorkDir();
        //}

        //public sbyte SwitchWorkDir(string bstrNewWorkDir)
        //{
        //    return p_IServer.SwitchWorkDir(bstrNewWorkDir);
        //}

        //public short GetNoOfJobsInWorkDir()
        //{
        //    return p_IServer.GetNoOfJobsInWorkDir();
        //}

        //public string GetJobAt(short jobIndex)
        //{
        //    return p_IServer.GetJobAt(jobIndex);
        //}

        //public short GetNoOfCasesInJob(string jobName)
        //{
        //    return p_IServer.GetNoOfCasesInJob(jobName);
        //}

        //public string GetCaseAt(ref string jobName, short caseIndex)
        //{
        //    return p_IServer.GetCaseAt(jobName, caseIndex);
        //}

        //public short SSRunAllUnits()
        //{
        //    return p_IServer.SSRunAllUnits();
        //}

        //public short SSRunSelectedUnits(object unitIDs)
        //{
        //    return p_IServer.SSRunSelectedUnits(unitIDs);
        //}

        //public short DynamicRunAllSteps()
        //{
        //    return p_IServer.DynamicRunAllSteps();
        //}

        //public short DynamicRunStep()
        //{
        //    return p_IServer.DynamicRunStep();
        //}

        //public void DynamicRestoreToInitialState()
        //{
        //    p_IServer.DynamicRestoreToInitialState();
        //}

        //public float GetDynamicTimeInMinute()
        //{
        //    return p_IServer.GetDynamicTimeInMinute();
        //}

        //public float GetDynamicTimeStepInMinute()
        //{
        //    return p_IServer.GetDynamicTimeStepInMinute();
        //}

        //public short GetSimulationMode()
        //{
        //    return p_IServer.GetSimulationMode();
        //}

        //public void TakeProcessSnapShot(string pathName)
        //{
        //    p_IServer.TakeProcessSnapShot(pathName);
        //}

        //public void LoadProcessSnapShot(string pathName)
        //{
        //    p_IServer.LoadProcessSnapShot(pathName);
        //}

        //public void PauseSimulation()
        //{
        //    p_IServer.PauseSimulation();
        //}

        //public void SetCurrentAsInitialState()
        //{
        //    p_IServer.SetCurrentAsInitialState();
        //}

        //public float OTSGetTimeScale()
        //{
        //    return p_IServer.OTSGetTimeScale();
        //}

        //public short OTSSetTimeScale(float timeScale)
        //{
        //    return p_IServer.OTSSetTimeScale(timeScale);
        //}

        //public string ShowRunTimeMessages()
        //{
        //    return p_IServer.ShowRunTimeMessages();
        //}

        //public short SSRunOptimization(short runMode)
        //{
        //    return p_IServer.SSRunOptimization(runMode);
        //}

        //public short GetAppVersion()
        //{
        //    return p_IServer.GetAppVersion();
        //}

        //public short EditUnitOpPar(short unitOpID)
        //{
        //    return p_IServer.EditUnitOpPar(unitOpID);
        //}

        //public short SSRunOptimizationFile(string filePath, short runMode)
        //{
        //    return p_IServer.SSRunOptimizationFile(filePath, runMode);
        //}

        //public short SSGetOptimizationVariables(string bstrOptFile, ref short nVars, object fVarIniVal, object fVarLB, object fVarUB, ref short nCons, object fConsLB, object fConsUB)
        //{
        //    return p_IServer.SSGetOptimizationVariables(bstrOptFile, nVars, fVarIniVal, fVarLB, fVarUB, nCons, fConsLB, fConsUB);
        //}

        //public short SSPutOptimizationVariables(string bstrOptFile, short nVars, object fVarIniVal, object fVarLB, object fVarUB, short nCons, object fConsLB, object fConsUB)
        //{
        //    return p_IServer.SSPutOptimizationVariables(bstrOptFile, nVars, fVarIniVal, fVarLB, fVarUB, nCons, fConsLB, fConsUB);
        //}

        //public short SaveSimulation()
        //{
        //    return p_IServer.SaveSimulation();
        //}

        public void CloseSimulation()
        {
            Guid IID_NULL = new Guid("00000000-0000-0000-0000-000000000000");
            string[] rgsNames = new string[1] { "CloseSimulation" };
            int[] rgDispId = new int[1] { 0 };

            int hrRet = p_IDispatch.GetIDsOfNames
            (
                ref IID_NULL,
                rgsNames,
                1,
                LOCALE_SYSTEM_DEFAULT,
                rgDispId
            );
            object varResult = null;

            if (hrRet == 0)
            {
                System.Runtime.InteropServices.ComTypes.EXCEPINFO ExcepInfo = new System.Runtime.InteropServices.ComTypes.EXCEPINFO();
                UInt32 pArgErr = 0;

                Variant[] v = new Variant[0];
                System.Runtime.InteropServices.GCHandle rgvarg = System.Runtime.InteropServices.GCHandle.Alloc(v, System.Runtime.InteropServices.GCHandleType.Pinned);

                var dispParams = new System.Runtime.InteropServices.ComTypes.DISPPARAMS()
                {
                    cArgs = 0,
                    cNamedArgs = 0,
                    rgdispidNamedArgs = IntPtr.Zero,
                    rgvarg = rgvarg.AddrOfPinnedObject()
                };

                hrRet = p_IDispatch.Invoke
                (
                    rgDispId[0],
                    ref IID_NULL,
                    LOCALE_SYSTEM_DEFAULT,
                    DISPATCH_METHOD,
                    ref dispParams,
                    ref varResult,
                    ref ExcepInfo,
                    out pArgErr
                );
                rgvarg.Free();
            }

            if (hrRet == 0)
            {
            }
        }

        //public short SSRunNamedSequence(object SequenceName)
        //{
        //    return p_IServer.SSRunNamedSequence(SequenceName);
        //}

        //public short LoadSim(object sim_name, object read_only)
        //{
        //    return p_IServer.LoadSim(sim_name, read_only);
        //}

        //public long pid()
        //{
        //    return p_IServer.pid();
        //}


    }
}
