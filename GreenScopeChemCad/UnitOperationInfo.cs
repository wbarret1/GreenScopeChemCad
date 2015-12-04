using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
// using System.Threading.Tasks;

namespace GreenScopeChemCad
{
    class UnitOperationInfo
    {
        //CHEMCAD.IUnitOpInfo p_IUnitOpInfo;
        IDispatch p_IDispatch;

        const int LOCALE_SYSTEM_DEFAULT = 2048;
        const int DISPATCH_METHOD = 0x1;
        const int DISPATCH_PROPERTYGET = 0x2;
        const int DISPATCH_PROPERTYPUT = 0x4;
        const int SizeOfNativeVariant = 16;
        const int DISPID_PROPERTYPUT = -3;

        public UnitOperationInfo(object unitOpInfo)
        {
            //p_IUnitOpInfo = (CHEMCAD.IUnitOpInfo)unitOpInfo;
            p_IDispatch = (IDispatch)unitOpInfo;
        }

        //~UnitOperationInfo()
        //{
        //    //if (p_IFlowsheet != null)
        //    //    System.Runtime.InteropServices.Marshal.FinalReleaseComObject(p_IFlowsheet);
        //}

        public int UnitOpSpecArrayDiemsions()
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

        public double[] GetUnitOpSpec(int unitOpId)
        {
            Guid IID_NULL = new Guid("00000000-0000-0000-0000-000000000000");
            string[] rgsNames = new string[1] { "GetUnitOpSpecByID" };
            int[] rgDispId = new int[1] { 0 };
            double[] retVal = { 0 };

            int hrRet = p_IDispatch.GetIDsOfNames
            (
                ref IID_NULL,
                rgsNames,
                1,
                LOCALE_SYSTEM_DEFAULT,
                rgDispId
            );

            if (hrRet == 0)
            {
                object varResult = null;
                System.Runtime.InteropServices.ComTypes.EXCEPINFO ExcepInfo = new System.Runtime.InteropServices.ComTypes.EXCEPINFO();
                UInt32 pArgErr = 0;

                SafeArrayBounds bounds = new SafeArrayBounds();
                bounds.cElements = 251;
                bounds.lBound = 0;
                IntPtr psa = NativeMethods.SafeArrayCreate((ushort)(VARENUM.VT_R4), 1, ref bounds);
                System.Runtime.InteropServices.GCHandle handle0 = System.Runtime.InteropServices.GCHandle.Alloc(psa, System.Runtime.InteropServices.GCHandleType.Pinned);

                Variant[] v = new Variant[3];
                v[0].vt = 0x6004;
                v[0].data01 = handle0.AddrOfPinnedObject();
                v[1].vt = 0x0002;
                v[1].iVal = (short)unitOpId;
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

                if (hrRet == 0)
                {

                    short numParams = (short)varResult;
                    retVal = new double[numParams + 1];
                    float val = 0.0F;
                    System.Runtime.InteropServices.GCHandle pVal = System.Runtime.InteropServices.GCHandle.Alloc(val, System.Runtime.InteropServices.GCHandleType.Pinned);
                    for (int i = 0; i < numParams + 1; i++)
                    {
                        long longVal = i;
                        hrRet = NativeMethods.SafeArrayGetElement(psa, ref longVal, pVal.AddrOfPinnedObject());
                        retVal[i] = (float)pVal.Target;
                    }
                    pVal.Free();
                }
                NativeMethods.SafeArrayDestroy(psa);
                rgvarg.Free();
                handle0.Free();
            }
            return retVal;
        }

        public string GetUnitOpLabel(int unitOpId)
        {
            Guid IID_NULL = new Guid("00000000-0000-0000-0000-000000000000");
            string[] rgsNames = new string[1] { "GetUnitOpLabelByID" };
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
                v[0].iVal = (short)unitOpId;
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

        public string GetUnitOpCategory(int unitOpId)
        {
            Guid IID_NULL = new Guid("00000000-0000-0000-0000-000000000000");
            string[] rgsNames = new string[1] { "GetUnitOpCategoryByID" };
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

            if (hrRet == 0)
            {
                System.Runtime.InteropServices.ComTypes.EXCEPINFO ExcepInfo = new System.Runtime.InteropServices.ComTypes.EXCEPINFO();
                UInt32 pArgErr = 0;

                Variant[] v = new Variant[2];
                v[0].vt = 0x4008;
                v[0].data01 = handle0.AddrOfPinnedObject();
                v[1].vt = 0x0002;
                v[1].iVal = (short)unitOpId;
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
            string retVal = String.Empty;

            if (hrRet == 0)
            {
                retVal = System.Runtime.InteropServices.Marshal.PtrToStringBSTR((IntPtr)handle0.Target);
            }
            handle0.Free();
            return retVal;
        }

        //public string FromInternalUnitsToCurUserUnits(double[] specInInternal, double[] specInCurUser)
        //{
        //    Guid IID_NULL = new Guid("00000000-0000-0000-0000-000000000000");
        //    string[] rgsNames = new string[1] { "FromInternalUnitsToCurUserUnits" };
        //    int[] rgDispId = new int[1] { 0 };

        //    int hrRet = p_IDispatch.GetIDsOfNames
        //    (
        //        ref IID_NULL,
        //        rgsNames,
        //        1,
        //        LOCALE_SYSTEM_DEFAULT,
        //        rgDispId
        //    );
        //    object varResult = null;
            
        //    System.Runtime.InteropServices.GCHandle handle0 = System.Runtime.InteropServices.GCHandle.Alloc(pString0, System.Runtime.InteropServices.GCHandleType.Pinned);

        //    if (hrRet == 0)
        //    {
        //        System.Runtime.InteropServices.ComTypes.EXCEPINFO ExcepInfo = new System.Runtime.InteropServices.ComTypes.EXCEPINFO();
        //        UInt32 pArgErr = 0;

        //        Variant[] v = new Variant[2];
        //        v[0].vt = 0x4008;
        //        v[0].data01 = handle0.AddrOfPinnedObject();
        //        v[1].vt = 0x0002;
        //        v[1].iVal = (short)unitOpId;
        //        System.Runtime.InteropServices.GCHandle rgvarg = System.Runtime.InteropServices.GCHandle.Alloc(v, System.Runtime.InteropServices.GCHandleType.Pinned);

        //        var dispParams = new System.Runtime.InteropServices.ComTypes.DISPPARAMS()
        //        {
        //            cArgs = 2,
        //            cNamedArgs = 0,
        //            rgdispidNamedArgs = IntPtr.Zero,
        //            rgvarg = rgvarg.AddrOfPinnedObject()
        //        };

        //        hrRet = p_IDispatch.Invoke
        //        (
        //            rgDispId[0],
        //            ref IID_NULL,
        //            LOCALE_SYSTEM_DEFAULT,
        //            DISPATCH_METHOD,
        //            ref dispParams,
        //            ref varResult,
        //            ref ExcepInfo,
        //            out pArgErr
        //        );
        //        rgvarg.Free();
        //    }
        //    string retVal = String.Empty;

        //    if (hrRet == 0)
        //    {
        //        retVal = System.Runtime.InteropServices.Marshal.PtrToStringBSTR((IntPtr)handle0.Target);
        //    }
        //    handle0.Free();
        //    return retVal;
        //}
    }
}
