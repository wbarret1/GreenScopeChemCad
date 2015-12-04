using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
// using System.Threading.Tasks;

namespace GreenScopeChemCad
{
    class UnitOpSpecUnitConversion
    {
        //CHEMCAD.IUnitOpSpecUnitConversion p_IUnitOpSpecUnitConversion;
        IDispatch p_IDispatch;

        const int LOCALE_SYSTEM_DEFAULT = 2048;
        const int DISPATCH_METHOD = 0x1;
        const int DISPATCH_PROPERTYGET = 0x2;
        const int DISPATCH_PROPERTYPUT = 0x4;
        const int SizeOfNativeVariant = 16;
        const int DISPID_PROPERTYPUT = -3;

        public UnitOpSpecUnitConversion(object unitOpInfo)
        {
           // p_IUnitOpSpecUnitConversion = (CHEMCAD.IUnitOpSpecUnitConversion)unitOpInfo;
            p_IDispatch = (IDispatch)unitOpInfo;
        }

        //~UnitOpSpecUnitConversion()
        //{
        //    //if (p_IFlowsheet != null)
        //    //    System.Runtime.InteropServices.Marshal.FinalReleaseComObject(p_IFlowsheet);
        //}

        public int NumberOfParameters(int unitOpId)
        {
            Guid IID_NULL = new Guid("00000000-0000-0000-0000-000000000000");
            string[] rgsNames = new string[1] { "GetNoOfParameters" };
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
                v[0].iVal = (short)(unitOpId);
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
                if ((short)varResult == 0) return (int)((short)varResult);
            }
            return 250;
        }

        public void GetCurUserUnitString(int unitOpId, int paramID, ref string paramUnit, ref string paramName)
        {
            Guid IID_NULL = new Guid("00000000-0000-0000-0000-000000000000");
            string[] rgsNames = new string[1] { "GetCurUserUnitString" };
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
            IntPtr pString0 = System.Runtime.InteropServices.Marshal.StringToBSTR(String.Empty);
            System.Runtime.InteropServices.GCHandle handle0 = System.Runtime.InteropServices.GCHandle.Alloc(pString0, System.Runtime.InteropServices.GCHandleType.Pinned);
            IntPtr pString1 = System.Runtime.InteropServices.Marshal.StringToBSTR(String.Empty);
            System.Runtime.InteropServices.GCHandle handle1 = System.Runtime.InteropServices.GCHandle.Alloc(pString0, System.Runtime.InteropServices.GCHandleType.Pinned);

            if (hrRet == 0)
            {
                System.Runtime.InteropServices.ComTypes.EXCEPINFO ExcepInfo = new System.Runtime.InteropServices.ComTypes.EXCEPINFO();
                UInt32 pArgErr = 0;

                Variant[] v = new Variant[4];
                v[0].vt = 0x4008;
                v[0].data01 = handle0.AddrOfPinnedObject();
                v[1].vt = 0x4008;
                v[1].data01 = handle0.AddrOfPinnedObject();
                v[2].vt = 0x0002;
                v[2].iVal = (short)(paramID);
                v[3].vt = 0x0002;
                v[3].iVal = (short)(unitOpId);
                System.Runtime.InteropServices.GCHandle rgvarg = System.Runtime.InteropServices.GCHandle.Alloc(v, System.Runtime.InteropServices.GCHandleType.Pinned);

                var dispParams = new System.Runtime.InteropServices.ComTypes.DISPPARAMS()
                {
                    cArgs = 4,
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
                paramUnit = System.Runtime.InteropServices.Marshal.PtrToStringAuto((IntPtr)handle0.Target);
                paramName = System.Runtime.InteropServices.Marshal.PtrToStringAuto((IntPtr)handle1.Target);
            }
        }

        public int GetUnitOpSpecArrayDimension()
        {
            Guid IID_NULL = new Guid("00000000-0000-0000-0000-000000000000");
            string[] rgsNames = new string[1] { "GetUnitOpSpecArrayDimension" };
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

        public void FromInternalUnitsToCurUserUnits(double[] specInInternal, ref double[] specInCurUser)
        {
            Guid IID_NULL = new Guid("00000000-0000-0000-0000-000000000000");
            string[] rgsNames = new string[1] { "FromInternalUnitsToCurUserUnits" };
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
            SafeArrayBounds bounds = new SafeArrayBounds();
            bounds.cElements = 251;
            bounds.lBound = 0;
            IntPtr psa0 = NativeMethods.SafeArrayCreate((ushort)(VARENUM.VT_R4), 1, ref bounds);
            System.Runtime.InteropServices.GCHandle handle0 = System.Runtime.InteropServices.GCHandle.Alloc(psa0, System.Runtime.InteropServices.GCHandleType.Pinned);

            float val = 0.0F;
            System.Runtime.InteropServices.GCHandle pVal = System.Runtime.InteropServices.GCHandle.Alloc(val, System.Runtime.InteropServices.GCHandleType.Pinned);
            for (long i = 0; i < specInInternal.Length; i++)
            {
                long longVal = i;
                val = (float)specInInternal[i];
                hrRet = NativeMethods.SafeArrayPutElement(psa0, ref longVal, pVal.AddrOfPinnedObject());
            }
            pVal.Free();


            IntPtr psa1 = NativeMethods.SafeArrayCreate((ushort)(VARENUM.VT_R4), 1, ref bounds);
            System.Runtime.InteropServices.GCHandle handle1 = System.Runtime.InteropServices.GCHandle.Alloc(psa1, System.Runtime.InteropServices.GCHandleType.Pinned);

            if (hrRet == 0)
            {
                System.Runtime.InteropServices.ComTypes.EXCEPINFO ExcepInfo = new System.Runtime.InteropServices.ComTypes.EXCEPINFO();
                UInt32 pArgErr = 0;

                Variant[] v = new Variant[2];
                v[0].vt = 0x6004;
                v[0].data01 = handle0.AddrOfPinnedObject();
                v[1].vt = 0x4004;
                v[1].data01 = handle1.AddrOfPinnedObject();
                System.Runtime.InteropServices.GCHandle rgvarg = System.Runtime.InteropServices.GCHandle.Alloc(v, System.Runtime.InteropServices.GCHandleType.Pinned);

                var dispParams = new System.Runtime.InteropServices.ComTypes.DISPPARAMS()
                {
                    cArgs = 2,
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
                specInCurUser = new double[251];
                float val1 = 0.0F;
                pVal = System.Runtime.InteropServices.GCHandle.Alloc(val1, System.Runtime.InteropServices.GCHandleType.Pinned);
                for (int i = 0; i < 250 + 1; i++)
                {
                    long longVal = i;
                    hrRet = NativeMethods.SafeArrayGetElement(psa1, ref longVal, pVal.AddrOfPinnedObject());
                    specInCurUser[i] = (float)pVal.Target;
                }
                pVal.Free();
            }
            handle0.Free();
            handle1.Free();
        }


        public void FromCurUserUnitsToInternalUnits(double[] specInInternal, ref double[] specInCurUser)
        {
            Guid IID_NULL = new Guid("00000000-0000-0000-0000-000000000000");
            string[] rgsNames = new string[1] { "FromCurUserUnitsToInternalUnits" };
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
            SafeArrayBounds bounds = new SafeArrayBounds();
            bounds.cElements = 251;
            bounds.lBound = 0;
            IntPtr psa0 = NativeMethods.SafeArrayCreate((ushort)(VARENUM.VT_R4), 1, ref bounds);
            System.Runtime.InteropServices.GCHandle handle0 = System.Runtime.InteropServices.GCHandle.Alloc(psa0, System.Runtime.InteropServices.GCHandleType.Pinned);

            float val = 0.0F;
            System.Runtime.InteropServices.GCHandle pVal = System.Runtime.InteropServices.GCHandle.Alloc(val, System.Runtime.InteropServices.GCHandleType.Pinned);
            for (long i = 0; i < specInInternal.Length; i++)
            {
                long longVal = i;
                val = (float)specInInternal[i];
                hrRet = NativeMethods.SafeArrayPutElement(psa0, ref longVal, pVal.AddrOfPinnedObject());
            }
            pVal.Free();


            IntPtr psa1 = NativeMethods.SafeArrayCreate((ushort)(VARENUM.VT_R4), 1, ref bounds);
            System.Runtime.InteropServices.GCHandle handle1 = System.Runtime.InteropServices.GCHandle.Alloc(psa1, System.Runtime.InteropServices.GCHandleType.Pinned);

            if (hrRet == 0)
            {
                System.Runtime.InteropServices.ComTypes.EXCEPINFO ExcepInfo = new System.Runtime.InteropServices.ComTypes.EXCEPINFO();
                UInt32 pArgErr = 0;

                Variant[] v = new Variant[2];
                v[0].vt = 0x6004;
                v[0].data01 = handle0.AddrOfPinnedObject();
                v[1].vt = 0x4004;
                v[1].data01 = handle1.AddrOfPinnedObject();
                System.Runtime.InteropServices.GCHandle rgvarg = System.Runtime.InteropServices.GCHandle.Alloc(v, System.Runtime.InteropServices.GCHandleType.Pinned);

                var dispParams = new System.Runtime.InteropServices.ComTypes.DISPPARAMS()
                {
                    cArgs = 2,
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
                specInCurUser = new double[251];
                float val1 = 0.0F;
                pVal = System.Runtime.InteropServices.GCHandle.Alloc(val1, System.Runtime.InteropServices.GCHandleType.Pinned);
                for (int i = 0; i < 250 + 1; i++)
                {
                    long longVal = i;
                    hrRet = NativeMethods.SafeArrayGetElement(psa1, ref longVal, pVal.AddrOfPinnedObject());
                    specInCurUser[i] = (float)pVal.Target;
                }
                pVal.Free();
            }
            handle0.Free();
            handle1.Free();
        }
    }
}
