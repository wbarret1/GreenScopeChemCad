using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
//// using System.Threading.Tasks;

namespace GreenScopeChemCad
{
    class StreamInfo
    {
        //CHEMCAD.IStreamInfo p_IStreamInfo;
        IDispatch p_IDispatch;

        const int LOCALE_SYSTEM_DEFAULT = 2048;
        const int DISPATCH_METHOD = 0x1;
        const int DISPATCH_PROPERTYGET = 0x2;
        const int DISPATCH_PROPERTYPUT = 0x4;
        const int SizeOfNativeVariant = 16;
        const int DISPID_PROPERTYPUT = -3;

        public StreamInfo(object streamInfo)
        {
            //p_IStreamInfo = (CHEMCAD.IStreamInfo)streamInfo;
            p_IDispatch = (IDispatch)streamInfo;
        }

        //~StreamInfo()
        //{
        //    //if (p_IFlowsheet != null)
        //    //    System.Runtime.InteropServices.Marshal.FinalReleaseComObject(p_IFlowsheet);
        //}

        public int NumberOfComponents
        {
            get
            {
                Guid IID_NULL = new Guid("00000000-0000-0000-0000-000000000000");
                string[] rgsNames = new string[1] { "GetNoOfComponents" };
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
                int retVal = 0;

                if (hrRet == 0)
                {
                    retVal = (short)varResult;
                }
                return retVal;
            }
        }

        public String GetStreamLabelByID(int streamID)
        {
            Guid IID_NULL = new Guid("00000000-0000-0000-0000-000000000000");
            string[] rgsNames = new string[1] { "GetStreamLabelByID" };
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
                v[0].iVal = (short)streamID;
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
            string retVal = String.Empty;

            if (hrRet == 0)
            {
                retVal = (string)varResult;
            }
            return retVal;
        }

        //void GetCurUserUnitString(VARIANT* tempUnit, VARIANT* presUnit, VARIANT* enthUnit, VARIANT* tMoleRateUnit, VARIANT* tMassRateUnit, VARIANT* tStdLVolRateUnit, VARIANT* tStdVVolRateUnit, VARIANT* compUnit, VARIANT* compCate);

        public void GetCurUserUnitString(ref string temperatureUnit, ref string pressureUnit, ref string enthalpyUnit, ref string TotalMoleRateUnit, ref string TotalMassRateUnit, ref string TotalStdLVolRateUnit, ref string TotalStdVVolRateUnit, ref string ComponentFlowRateUnit, ref string compCate)
        {
            Guid IID_NULL = new Guid("00000000-0000-0000-0000-000000000000");
            string[] rgsNames = new string[1] { "GetCurUserUnitString" };
            int[] rgDispId = new int[1] { 0 };
            int[] retVal = { 0 };

            int hrRet = p_IDispatch.GetIDsOfNames
            (
                ref IID_NULL,
                rgsNames,
                1,
                LOCALE_SYSTEM_DEFAULT,
                rgDispId
            );
            object varResult = null;
            IntPtr pString0 = System.Runtime.InteropServices.Marshal.StringToBSTR(String.Empty);
            System.Runtime.InteropServices.GCHandle handle0 = System.Runtime.InteropServices.GCHandle.Alloc(pString0, System.Runtime.InteropServices.GCHandleType.Pinned);
            IntPtr pString1 = System.Runtime.InteropServices.Marshal.StringToBSTR(String.Empty);
            System.Runtime.InteropServices.GCHandle handle1 = System.Runtime.InteropServices.GCHandle.Alloc(pString1, System.Runtime.InteropServices.GCHandleType.Pinned);
            IntPtr pString2 = System.Runtime.InteropServices.Marshal.StringToBSTR(String.Empty);
            System.Runtime.InteropServices.GCHandle handle2 = System.Runtime.InteropServices.GCHandle.Alloc(pString2, System.Runtime.InteropServices.GCHandleType.Pinned);
            IntPtr pString3 = System.Runtime.InteropServices.Marshal.StringToBSTR(String.Empty);
            System.Runtime.InteropServices.GCHandle handle3 = System.Runtime.InteropServices.GCHandle.Alloc(pString3, System.Runtime.InteropServices.GCHandleType.Pinned);
            IntPtr pString4 = System.Runtime.InteropServices.Marshal.StringToBSTR(String.Empty);
            System.Runtime.InteropServices.GCHandle handle4 = System.Runtime.InteropServices.GCHandle.Alloc(pString4, System.Runtime.InteropServices.GCHandleType.Pinned);
            IntPtr pString5 = System.Runtime.InteropServices.Marshal.StringToBSTR(String.Empty);
            System.Runtime.InteropServices.GCHandle handle5 = System.Runtime.InteropServices.GCHandle.Alloc(pString5, System.Runtime.InteropServices.GCHandleType.Pinned);
            IntPtr pString6 = System.Runtime.InteropServices.Marshal.StringToBSTR(String.Empty);
            System.Runtime.InteropServices.GCHandle handle6 = System.Runtime.InteropServices.GCHandle.Alloc(pString6, System.Runtime.InteropServices.GCHandleType.Pinned);
            IntPtr pString7 = System.Runtime.InteropServices.Marshal.StringToBSTR(String.Empty);
            System.Runtime.InteropServices.GCHandle handle7 = System.Runtime.InteropServices.GCHandle.Alloc(pString7, System.Runtime.InteropServices.GCHandleType.Pinned);
            IntPtr pString8 = System.Runtime.InteropServices.Marshal.StringToBSTR(String.Empty);
            System.Runtime.InteropServices.GCHandle handle8 = System.Runtime.InteropServices.GCHandle.Alloc(pString8, System.Runtime.InteropServices.GCHandleType.Pinned);
            IntPtr pString9 = System.Runtime.InteropServices.Marshal.StringToBSTR(String.Empty);
            System.Runtime.InteropServices.GCHandle handle9 = System.Runtime.InteropServices.GCHandle.Alloc(pString9, System.Runtime.InteropServices.GCHandleType.Pinned);

            if (hrRet == 0)
            {
                System.Runtime.InteropServices.ComTypes.EXCEPINFO ExcepInfo = new System.Runtime.InteropServices.ComTypes.EXCEPINFO();
                UInt32 pArgErr = 0;
                //void GetCurUserUnitString(VARIANT* tempUnit, VARIANT* presUnit, VARIANT* enthUnit, VARIANT* tMoleRateUnit, VARIANT* tMassRateUnit, VARIANT* tStdLVolRateUnit, VARIANT* tStdVVolRateUnit, VARIANT* compUnit, VARIANT* compCate);

                Variant[] v = new Variant[9];
                v[0].vt = 0x4008;
                v[0].data01 = handle0.AddrOfPinnedObject();
                v[1].vt = 0x4008;
                v[1].data01 = handle1.AddrOfPinnedObject();
                v[2].vt = 0x4008;
                v[2].data01 = handle2.AddrOfPinnedObject();
                v[3].vt = 0x4008;
                v[3].data01 = handle3.AddrOfPinnedObject();
                v[4].vt = 0x4008;
                v[4].data01 = handle4.AddrOfPinnedObject();
                v[5].vt = 0x4008;
                v[5].data01 = handle5.AddrOfPinnedObject();
                v[6].vt = 0x4008;
                v[6].data01 = handle6.AddrOfPinnedObject();
                v[7].vt = 0x4008;
                v[7].data01 = handle7.AddrOfPinnedObject();
                v[8].vt = 0x4008;
                v[8].data01 = handle8.AddrOfPinnedObject();
                System.Runtime.InteropServices.GCHandle rgvarg = System.Runtime.InteropServices.GCHandle.Alloc(v, System.Runtime.InteropServices.GCHandleType.Pinned);

                var dispParams = new System.Runtime.InteropServices.ComTypes.DISPPARAMS()
                {
                    cArgs = 9,
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
                //public void GetCurUserUnitString(ref string temperatureUnit, ref string pressureUnit, ref string enthalpyUnit, ref string TotalMoleRateUnit, ref string TotalMassRateUnit, ref string TotalStdLVolRateUnit, ref string TotalStdVVolRateUnit, ref string ComponentFlowRateUnit, ref string compCate)
                compCate = System.Runtime.InteropServices.Marshal.PtrToStringBSTR((IntPtr)handle0.Target);
                ComponentFlowRateUnit = System.Runtime.InteropServices.Marshal.PtrToStringBSTR((IntPtr)handle1.Target);
                TotalStdVVolRateUnit = System.Runtime.InteropServices.Marshal.PtrToStringBSTR((IntPtr)handle2.Target);
                TotalStdLVolRateUnit = System.Runtime.InteropServices.Marshal.PtrToStringBSTR((IntPtr)handle3.Target);
                TotalMassRateUnit = System.Runtime.InteropServices.Marshal.PtrToStringBSTR((IntPtr)handle4.Target);
                TotalMoleRateUnit = System.Runtime.InteropServices.Marshal.PtrToStringBSTR((IntPtr)handle5.Target);
                enthalpyUnit = System.Runtime.InteropServices.Marshal.PtrToStringBSTR((IntPtr)handle6.Target);
                pressureUnit = System.Runtime.InteropServices.Marshal.PtrToStringBSTR((IntPtr)handle7.Target);
                temperatureUnit = System.Runtime.InteropServices.Marshal.PtrToStringBSTR((IntPtr)handle8.Target);
            }
            handle0.Free();
            handle1.Free();
            handle2.Free();
            handle3.Free();
            handle4.Free();
            handle5.Free();
            handle6.Free();
            handle7.Free();
            handle8.Free();
        }

        public void GetStreamInCurUserUnitByID(int streamID, ref double temperature, ref double pressure, ref double enthalpy, ref double moleVapFrac, ref double tMoleRate, ref double tMassRate, ref double tStdLVolRate, ref double tStdVVolRate, ref double[] compRate)
        {
            Guid IID_NULL = new Guid("00000000-0000-0000-0000-000000000000");
            string[] rgsNames = new string[1] { "GetStreamInCurUserUnitByID" };// GetStreamInCurUserUnitByID
            int[] rgDispId = new int[1] { 0 };
            int[] retVal = { 0 };

            int hrRet = p_IDispatch.GetIDsOfNames
            (
                ref IID_NULL,
                rgsNames,
                1,
                LOCALE_SYSTEM_DEFAULT,
                rgDispId
            );
            object varResult = null;

            SafeArrayBounds bounds = new SafeArrayBounds();
            bounds.cElements = 201;
            bounds.lBound = 0;
            IntPtr psa = NativeMethods.SafeArrayCreate((ushort)(VARENUM.VT_R4), 1, ref bounds);
            System.Runtime.InteropServices.GCHandle handle0 = System.Runtime.InteropServices.GCHandle.Alloc(psa, System.Runtime.InteropServices.GCHandleType.Pinned);
            float f_tStdVVolRate = 0.0F;
            System.Runtime.InteropServices.GCHandle handle1 = System.Runtime.InteropServices.GCHandle.Alloc(f_tStdVVolRate, System.Runtime.InteropServices.GCHandleType.Pinned);
            float f_tStdLVolRate = 0.0F;
            System.Runtime.InteropServices.GCHandle handle2 = System.Runtime.InteropServices.GCHandle.Alloc(f_tStdLVolRate, System.Runtime.InteropServices.GCHandleType.Pinned);
            float f_tMassRate = 0.0F;
            System.Runtime.InteropServices.GCHandle handle3 = System.Runtime.InteropServices.GCHandle.Alloc(f_tMassRate, System.Runtime.InteropServices.GCHandleType.Pinned);
            float f_tMoleRate = 0.0F;
            System.Runtime.InteropServices.GCHandle handle4 = System.Runtime.InteropServices.GCHandle.Alloc(f_tMoleRate, System.Runtime.InteropServices.GCHandleType.Pinned);
            float fMVF = 0.0F;
            System.Runtime.InteropServices.GCHandle handle5 = System.Runtime.InteropServices.GCHandle.Alloc(fMVF, System.Runtime.InteropServices.GCHandleType.Pinned);
            float fEnthalpy = 0.0F;
            System.Runtime.InteropServices.GCHandle handle6 = System.Runtime.InteropServices.GCHandle.Alloc(fEnthalpy, System.Runtime.InteropServices.GCHandleType.Pinned);
            float fPress = 0.0F;
            System.Runtime.InteropServices.GCHandle handle7 = System.Runtime.InteropServices.GCHandle.Alloc(fPress, System.Runtime.InteropServices.GCHandleType.Pinned);
            float fTemp = 0.0F;
            System.Runtime.InteropServices.GCHandle handle8 = System.Runtime.InteropServices.GCHandle.Alloc(fTemp, System.Runtime.InteropServices.GCHandleType.Pinned);

            if (hrRet == 0)
            {
                System.Runtime.InteropServices.ComTypes.EXCEPINFO ExcepInfo = new System.Runtime.InteropServices.ComTypes.EXCEPINFO();
                UInt32 pArgErr = 0;
                //short GetStreamInCurUserUnitByID(short streamID, 
                // float* temp,
                //float* pres, 
                //float* enth, 
                //float* mvf, 
                //float* tMoleRate, 
                // float* tMassRate, 
                // float* tStdLVolRate, 
                //float* tStdVVolRate, 
                // VARIANT compRate);


                Variant[] v = new Variant[10];
                v[0].vt = 0x6004;
                v[0].data01 = handle0.AddrOfPinnedObject();
                v[1].vt = 0x4004;
                v[1].data01 = handle1.AddrOfPinnedObject();
                v[2].vt = 0x4004;
                v[2].data01 = handle2.AddrOfPinnedObject();
                v[3].vt = 0x4004;
                v[3].data01 = handle3.AddrOfPinnedObject();
                v[4].vt = 0x4004;
                v[4].data01 = handle4.AddrOfPinnedObject();
                v[5].vt = 0x4004;
                v[5].data01 = handle5.AddrOfPinnedObject();
                v[6].vt = 0x4004;
                v[6].data01 = handle6.AddrOfPinnedObject();
                v[7].vt = 0x4004;
                v[7].data01 = handle7.AddrOfPinnedObject();
                v[8].vt = 0x4004;
                v[8].data01 = handle8.AddrOfPinnedObject();
                v[9].vt = 0x0002;
                v[9].iVal = (short)streamID;
                System.Runtime.InteropServices.GCHandle rgvarg = System.Runtime.InteropServices.GCHandle.Alloc(v, System.Runtime.InteropServices.GCHandleType.Pinned);

                var dispParams = new System.Runtime.InteropServices.ComTypes.DISPPARAMS()
                {
                    cArgs = 10,
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
                temperature = (float)handle8.Target;
                pressure = (float)handle7.Target;
                enthalpy = (float)handle6.Target;
                moleVapFrac = (float)handle5.Target;
                tMoleRate = (float)handle4.Target;
                tMassRate = (float)handle3.Target;
                tStdLVolRate = (float)handle2.Target;
                tStdVVolRate = (float)handle1.Target;
                int numComponents = this.NumberOfComponents;
                compRate = new double[numComponents];
                float val = 0.0F;
                System.Runtime.InteropServices.GCHandle pVal = System.Runtime.InteropServices.GCHandle.Alloc(val, System.Runtime.InteropServices.GCHandleType.Pinned);
                for (long i = 0; i < numComponents; i++)
                {
                    long longVal = i;
                    hrRet = NativeMethods.SafeArrayGetElement(psa, ref longVal, pVal.AddrOfPinnedObject());
                    compRate[i] = (float)pVal.Target;
                }
                pVal.Free();
            }
            NativeMethods.SafeArrayDestroy(psa);
            handle0.Free();
            handle1.Free();
            handle2.Free();
            handle3.Free();
            handle4.Free();
            handle5.Free();
            handle6.Free();
            handle7.Free();
            handle8.Free();
        }

        public void GetStreamByID(int streamID, ref double tempR, ref double pressPsia, ref double moleVapFrac, ref double enth_BTUHR, ref double[] compFlowLbMolHr)
        {
            Guid IID_NULL = new Guid("00000000-0000-0000-0000-000000000000");
            string[] rgsNames = new string[1] { "GetStreamByID" };
            int[] rgDispId = new int[1] { 0 };
            int[] retVal = { 0 };

            int hrRet = p_IDispatch.GetIDsOfNames
            (
                ref IID_NULL,
                rgsNames,
                1,
                LOCALE_SYSTEM_DEFAULT,
                rgDispId
            );
            object varResult = null;

            SafeArrayBounds bounds = new SafeArrayBounds();
            bounds.cElements = 201;
            bounds.lBound = 0;
            IntPtr psa = NativeMethods.SafeArrayCreate((ushort)(VARENUM.VT_R4), 1, ref bounds);
            System.Runtime.InteropServices.GCHandle handle0 = System.Runtime.InteropServices.GCHandle.Alloc(psa, System.Runtime.InteropServices.GCHandleType.Pinned);
            float fEnthalpy = 0.0F;
            System.Runtime.InteropServices.GCHandle handle1 = System.Runtime.InteropServices.GCHandle.Alloc(fEnthalpy, System.Runtime.InteropServices.GCHandleType.Pinned);
            float fMVF = 0.0F;
            System.Runtime.InteropServices.GCHandle handle2 = System.Runtime.InteropServices.GCHandle.Alloc(fMVF, System.Runtime.InteropServices.GCHandleType.Pinned);
            float fPress = 0.0F;
            System.Runtime.InteropServices.GCHandle handle3 = System.Runtime.InteropServices.GCHandle.Alloc(fPress, System.Runtime.InteropServices.GCHandleType.Pinned);
            float fTemp = 0.0F;
            System.Runtime.InteropServices.GCHandle handle4 = System.Runtime.InteropServices.GCHandle.Alloc(fTemp, System.Runtime.InteropServices.GCHandleType.Pinned);

            if (hrRet == 0)
            {
                System.Runtime.InteropServices.ComTypes.EXCEPINFO ExcepInfo = new System.Runtime.InteropServices.ComTypes.EXCEPINFO();
                UInt32 pArgErr = 0;

                Variant[] v = new Variant[6];
                v[0].vt = 0x6004;
                v[0].data01 = handle0.AddrOfPinnedObject();
                v[1].vt = 0x4004;
                v[1].data01 = handle1.AddrOfPinnedObject();
                v[2].vt = 0x4004;
                v[2].data01 = handle2.AddrOfPinnedObject();
                v[3].vt = 0x4004;
                v[3].data01 = handle3.AddrOfPinnedObject();
                v[4].vt = 0x4004;
                v[4].data01 = handle4.AddrOfPinnedObject();
                v[5].vt = 0x0002;
                v[5].iVal = (short)streamID;
                System.Runtime.InteropServices.GCHandle rgvarg = System.Runtime.InteropServices.GCHandle.Alloc(v, System.Runtime.InteropServices.GCHandleType.Pinned);

                var dispParams = new System.Runtime.InteropServices.ComTypes.DISPPARAMS()
                {
                    cArgs = 6,
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
                tempR = (float)handle4.Target;
                pressPsia = (float)handle3.Target;
                moleVapFrac = (float)handle2.Target;
                enth_BTUHR = (float)handle1.Target;
                int numComponents = (short)varResult;
                compFlowLbMolHr = new double[numComponents];
                float val = 0.0F;
                System.Runtime.InteropServices.GCHandle pVal = System.Runtime.InteropServices.GCHandle.Alloc(val, System.Runtime.InteropServices.GCHandleType.Pinned);
                for (long i = 0; i < numComponents; i++)
                {
                    long longVal = i;
                    hrRet = NativeMethods.SafeArrayGetElement(psa, ref longVal, pVal.AddrOfPinnedObject());
                    compFlowLbMolHr[i] = (float)pVal.Target;
                }
                pVal.Free();
            }
            NativeMethods.SafeArrayDestroy(psa);
            handle0.Free();
            handle1.Free();
            handle2.Free();
            handle3.Free();
            handle4.Free();
        }

        public void GetStreamCost(int streamID, ref int costType, ref double cost)
        {
            Guid IID_NULL = new Guid("00000000-0000-0000-0000-000000000000");
            string[] rgsNames = new string[1] { "GetStreamCost" };
            int[] rgDispId = new int[1] { 0 };
            int[] retVal = { 0 };

            int hrRet = p_IDispatch.GetIDsOfNames
            (
                ref IID_NULL,
                rgsNames,
                1,
                LOCALE_SYSTEM_DEFAULT,
                rgDispId
            );
            object varResult = null;

            float fVal = 0.0F;
            System.Runtime.InteropServices.GCHandle handle0 = System.Runtime.InteropServices.GCHandle.Alloc(fVal, System.Runtime.InteropServices.GCHandleType.Pinned);
            short iVal = 0;
            System.Runtime.InteropServices.GCHandle handle1 = System.Runtime.InteropServices.GCHandle.Alloc(iVal, System.Runtime.InteropServices.GCHandleType.Pinned);

            if (hrRet == 0)
            {
                System.Runtime.InteropServices.ComTypes.EXCEPINFO ExcepInfo = new System.Runtime.InteropServices.ComTypes.EXCEPINFO();
                UInt32 pArgErr = 0;

                Variant[] v = new Variant[3];
                v[0].vt = 0x4004;
                v[0].data01 = handle0.AddrOfPinnedObject();
                v[1].vt = 0x4002;
                v[1].data01 = handle1.AddrOfPinnedObject();
                v[2].vt = 0x0002;
                v[2].iVal = (short)streamID;
                System.Runtime.InteropServices.GCHandle rgvarg = System.Runtime.InteropServices.GCHandle.Alloc(v, System.Runtime.InteropServices.GCHandleType.Pinned);

                var dispParams = new System.Runtime.InteropServices.ComTypes.DISPPARAMS()
                {
                    cArgs = 3,
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
                cost = (float)handle0.Target;
                costType = (short)handle1.Target;
            }
            handle0.Free();
            handle1.Free();
        }

        public string GetComponentNameByPos(int position)
        {
            Guid IID_NULL = new Guid("00000000-0000-0000-0000-000000000000");
            string[] rgsNames = new string[1] { "GetComponentNameByPosBaseOne" };
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
                v[0].iVal = (short)(position + 1);
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
            string retVal = String.Empty;

            if (hrRet == 0)
            {
                retVal = (string)varResult;
            }
            return retVal;
        }

        public int GetComponentIDByPos(int position)
        {
            Guid IID_NULL = new Guid("00000000-0000-0000-0000-000000000000");
            string[] rgsNames = new string[1] { "GetComponentIDByPosBaseOne" };
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
                v[0].iVal = (short)(position + 1);
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
            int retVal = 0;

            if (hrRet == 0)
            {
                retVal = (short)varResult;
            }
            return retVal;
        }

        //vbserver_stream_status get_stream_current_engineering_units(short streamID, float* temp, float* pres, float* enth, float* mvf, total_flow_spec_type total_spec, float* total_rate, component_flow_spec_type comp_spec, VARIANT* comp_rate);
        public ServerStreamStatus getStreamCurrentEngineeringUnits(int streamID, ref double temperature, ref double pressure, ref double enthalpy, ref double vaporMoleFraction, TotalFlowSpecification  totalFlowSpecification, ref double totalFlowRate, ComponentFlowSpecType componentSpecification, ref double[] componentFlowRate)
        {
            Guid IID_NULL = new Guid("00000000-0000-0000-0000-000000000000");
            string[] rgsNames = new string[1] { "get_stream_current_engineering_units"};
            int[] rgDispId = new int[1] { 0 };
            ServerStreamStatus retVal = ServerStreamStatus.STREAM_STATUS_PARAMETER_INVALID;

            int hrRet = p_IDispatch.GetIDsOfNames
            (
                ref IID_NULL,
                rgsNames,
                1,
                LOCALE_SYSTEM_DEFAULT,
                rgDispId
            );
            object varResult = null;
            SafeArrayBounds bounds = new SafeArrayBounds();
            bounds.cElements = 201;
            bounds.lBound = 0;
            IntPtr psa = NativeMethods.SafeArrayCreate((ushort)(VARENUM.VT_R4), 1, ref bounds);
            System.Runtime.InteropServices.GCHandle handle0 = System.Runtime.InteropServices.GCHandle.Alloc(psa, System.Runtime.InteropServices.GCHandleType.Pinned);
            float fTotalFlowRate = 0.0F;
            System.Runtime.InteropServices.GCHandle handle1 = System.Runtime.InteropServices.GCHandle.Alloc(fTotalFlowRate, System.Runtime.InteropServices.GCHandleType.Pinned);
            float fMVF = 0.0F;
            System.Runtime.InteropServices.GCHandle handle2 = System.Runtime.InteropServices.GCHandle.Alloc(fMVF, System.Runtime.InteropServices.GCHandleType.Pinned);
            float fEnthalpy = 0.0F;
            System.Runtime.InteropServices.GCHandle handle3 = System.Runtime.InteropServices.GCHandle.Alloc(fEnthalpy, System.Runtime.InteropServices.GCHandleType.Pinned);
            float fPress = 0.0F;
            System.Runtime.InteropServices.GCHandle handle4 = System.Runtime.InteropServices.GCHandle.Alloc(fPress, System.Runtime.InteropServices.GCHandleType.Pinned);
            float fTemp = 0.0F;
            System.Runtime.InteropServices.GCHandle handle5 = System.Runtime.InteropServices.GCHandle.Alloc(fTemp, System.Runtime.InteropServices.GCHandleType.Pinned);

            if (hrRet == 0)
            {
                System.Runtime.InteropServices.ComTypes.EXCEPINFO ExcepInfo = new System.Runtime.InteropServices.ComTypes.EXCEPINFO();
                UInt32 pArgErr = 0;

                Variant[] v = new Variant[9];
                // CompRate
                v[0].vt = 0x6004;
                v[0].data01 = handle0.AddrOfPinnedObject();
                // CompFlowSpec
                v[1].vt = 0x0002;
                v[1].iVal = (short)componentSpecification;
                // Total Flow
                v[2].vt = 0x4004;
                v[2].data01 = handle1.AddrOfPinnedObject();
                // Total Flow Spec
                v[3].vt = 0x0002;
                v[3].iVal = (short)totalFlowSpecification;
                // mvf
                v[4].vt = 0x4004;
                v[4].data01 = handle2.AddrOfPinnedObject();
                // enth
                v[5].vt = 0x4004;
                v[5].data01 = handle3.AddrOfPinnedObject();
                // press
                v[6].vt = 0x4004;
                v[6].data01 = handle4.AddrOfPinnedObject();
                // temp
                v[7].vt = 0x4004;
                v[7].data01 = handle5.AddrOfPinnedObject();
                //streamID
                v[8].vt = 0x0002;
                v[8].iVal = (short)streamID;
                System.Runtime.InteropServices.GCHandle rgvarg = System.Runtime.InteropServices.GCHandle.Alloc(v, System.Runtime.InteropServices.GCHandleType.Pinned);

                var dispParams = new System.Runtime.InteropServices.ComTypes.DISPPARAMS()
                {
                    cArgs = 9,
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
            }
            if (hrRet == 0)
            {
                temperature = (float)handle5.Target;
                pressure = (float)handle4.Target;
                enthalpy = (float)handle3.Target;
                vaporMoleFraction = (float)handle2.Target;
                totalFlowRate = (float)handle1.Target;
                int numComponents = this.NumberOfComponents;
                componentFlowRate = new double[numComponents];
                float val = 0.0F;
                System.Runtime.InteropServices.GCHandle pVal = System.Runtime.InteropServices.GCHandle.Alloc(val, System.Runtime.InteropServices.GCHandleType.Pinned);
                for (long i = 0; i < numComponents; i++)
                {
                    long longVal = i;
                    hrRet = NativeMethods.SafeArrayGetElement(psa, ref longVal, pVal.AddrOfPinnedObject());
                    componentFlowRate[i] = (float)pVal.Target;
                }
                pVal.Free();
            }
            NativeMethods.SafeArrayDestroy(psa);
            handle0.Free();
            handle1.Free();
            handle2.Free();
            handle3.Free();
            handle4.Free();
            return ServerStreamStatus.STREAM_STATUS_NO_ERROR;
        }


        //vbserver_stream_status get_current_engineering_unit_labels(VARIANT* tempUnit, VARIANT* presUnit, VARIANT* enthUnit, total_flow_spec_type total_flow_type, VARIANT* total_flow_unit, component_flow_spec_type comp_rate_type, VARIANT* compUnit);
        public void getCurrentEngineeringUnitLabels(ref string temperatureUnit, ref string pressureUnit, ref string enthalpyUnit, TotalFlowSpecification totalFlowSpecification, ref string TotalFlowRateUnit, ComponentFlowSpecType componentSpecification, ref string ComponentFlowRateUnit)
        {
            Guid IID_NULL = new Guid("00000000-0000-0000-0000-000000000000");
            string[] rgsNames = new string[1] { "get_current_engineering_unit_labels" };
            int[] rgDispId = new int[1] { 0 };
            int[] retVal = { 0 };

            int hrRet = p_IDispatch.GetIDsOfNames
            (
                ref IID_NULL,
                rgsNames,
                1,
                LOCALE_SYSTEM_DEFAULT,
                rgDispId
            );
            object varResult = null;
            IntPtr pString0 = System.Runtime.InteropServices.Marshal.StringToBSTR(String.Empty);
            System.Runtime.InteropServices.GCHandle handle0 = System.Runtime.InteropServices.GCHandle.Alloc(pString0, System.Runtime.InteropServices.GCHandleType.Pinned);
            IntPtr pString1 = System.Runtime.InteropServices.Marshal.StringToBSTR(String.Empty);
            System.Runtime.InteropServices.GCHandle handle1 = System.Runtime.InteropServices.GCHandle.Alloc(pString1, System.Runtime.InteropServices.GCHandleType.Pinned);
            IntPtr pString2 = System.Runtime.InteropServices.Marshal.StringToBSTR(String.Empty);
            System.Runtime.InteropServices.GCHandle handle2 = System.Runtime.InteropServices.GCHandle.Alloc(pString2, System.Runtime.InteropServices.GCHandleType.Pinned);
            IntPtr pString3 = System.Runtime.InteropServices.Marshal.StringToBSTR(String.Empty);
            System.Runtime.InteropServices.GCHandle handle3 = System.Runtime.InteropServices.GCHandle.Alloc(pString3, System.Runtime.InteropServices.GCHandleType.Pinned);
            IntPtr pString4 = System.Runtime.InteropServices.Marshal.StringToBSTR(String.Empty);
            System.Runtime.InteropServices.GCHandle handle4 = System.Runtime.InteropServices.GCHandle.Alloc(pString4, System.Runtime.InteropServices.GCHandleType.Pinned);

            if (hrRet == 0)
            {
                System.Runtime.InteropServices.ComTypes.EXCEPINFO ExcepInfo = new System.Runtime.InteropServices.ComTypes.EXCEPINFO();
                UInt32 pArgErr = 0;
                //void GetCurUserUnitString(VARIANT* tempUnit, VARIANT* presUnit, VARIANT* enthUnit, VARIANT* tMoleRateUnit, VARIANT* tMassRateUnit, VARIANT* tStdLVolRateUnit, VARIANT* tStdVVolRateUnit, VARIANT* compUnit, VARIANT* compCate);

                Variant[] v = new Variant[7];
                //compUnit
                v[0].vt = 0x4008;
                v[0].data01 = handle0.AddrOfPinnedObject();
                //comp_rate_type
                v[1].vt = 0x0002;
                v[1].iVal = (short)componentSpecification;
                //total_flow_unit
                v[2].vt = 0x4008;
                v[2].data01 = handle1.AddrOfPinnedObject();
                //total_flow_type
                v[3].vt = 0x0002;
                v[3].iVal = (short)totalFlowSpecification;
                //enthUnit
                v[4].vt = 0x4008;
                v[4].data01 = handle2.AddrOfPinnedObject();
                //presUnit
                v[5].vt = 0x4008;
                v[5].data01 = handle3.AddrOfPinnedObject();
                //tempUnit
                v[6].vt = 0x4008;
                v[6].data01 = handle4.AddrOfPinnedObject();
                System.Runtime.InteropServices.GCHandle rgvarg = System.Runtime.InteropServices.GCHandle.Alloc(v, System.Runtime.InteropServices.GCHandleType.Pinned);

                var dispParams = new System.Runtime.InteropServices.ComTypes.DISPPARAMS()
                {
                    cArgs = 7,
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
                //public void GetCurUserUnitString(ref string temperatureUnit, ref string pressureUnit, ref string enthalpyUnit, ref string TotalMoleRateUnit, ref string TotalMassRateUnit, ref string TotalStdLVolRateUnit, ref string TotalStdVVolRateUnit, ref string ComponentFlowRateUnit, ref string compCate)
                ComponentFlowRateUnit = System.Runtime.InteropServices.Marshal.PtrToStringBSTR((IntPtr)handle0.Target);
                TotalFlowRateUnit = System.Runtime.InteropServices.Marshal.PtrToStringBSTR((IntPtr)handle1.Target);
                enthalpyUnit = System.Runtime.InteropServices.Marshal.PtrToStringBSTR((IntPtr)handle2.Target);
                pressureUnit = System.Runtime.InteropServices.Marshal.PtrToStringBSTR((IntPtr)handle3.Target);
                temperatureUnit = System.Runtime.InteropServices.Marshal.PtrToStringBSTR((IntPtr)handle4.Target);
            }
            handle0.Free();
            handle1.Free();
            handle2.Free();
            handle3.Free();
            handle4.Free();
        }




    }
}
