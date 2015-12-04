using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
//// using System.Threading.Tasks;

namespace GreenScopeChemCad
{

    // The COM IDispatch interface must be included in the C# source codes.
    [System.Runtime.InteropServices.ComImport()]
    [System.Runtime.InteropServices.Guid("00020400-0000-0000-C000-000000000046")]
    [System.Runtime.InteropServices.InterfaceType(System.Runtime.InteropServices.ComInterfaceType.InterfaceIsIUnknown)]
    public interface IDispatch
    {
        [System.Runtime.InteropServices.PreserveSig]
        int GetTypeInfoCount(out int Count);

        [System.Runtime.InteropServices.PreserveSig]
        int GetTypeInfo
            (
                [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.U4)] int iTInfo,
                [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.U4)] int lcid,
                out System.Runtime.InteropServices.ComTypes.ITypeInfo typeInfo
            );

        [System.Runtime.InteropServices.PreserveSig]
        int GetIDsOfNames
            (
                ref Guid riid,
                [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.LPArray, ArraySubType = System.Runtime.InteropServices.UnmanagedType.LPWStr)]
                string[] rgsNames,
                int cNames,
                int lcid,
                [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.LPArray)] int[] rgDispId
            );

        [System.Runtime.InteropServices.PreserveSig]
        int Invoke
            (
                int dispIdMember,
                ref Guid riid,
                uint lcid,
                ushort wFlags,
                ref System.Runtime.InteropServices.ComTypes.DISPPARAMS pDispParams,
                ref object pVarResult,
                ref System.Runtime.InteropServices.ComTypes.EXCEPINFO pExcepInfo,
                out UInt32 pArgErr
            );
    }

    enum VARENUM : ushort
    {
        VT_EMPTY = 0,
        VT_NULL = 1,
        VT_I2 = 2,
        VT_I4 = 3,
        VT_R4 = 4,
        VT_R8 = 5,
        VT_CY = 6,
        VT_DATE = 7,
        VT_BSTR = 8,
        VT_DISPATCH = 9,
        VT_ERROR = 10,
        VT_BOOL = 11,
        VT_VARIANT = 12,
        VT_UNKNOWN = 13,
        VT_DECIMAL = 14,
        VT_I1 = 16,
        VT_UI1 = 17,
        VT_UI2 = 18,
        VT_UI4 = 19,
        VT_I8 = 20,
        VT_UI8 = 21,
        VT_INT = 22,
        VT_UINT = 23,
        VT_VOID = 24,
        VT_HRESULT = 25,
        VT_PTR = 26,
        VT_SAFEARRAY = 27,
        VT_CARRAY = 28,
        VT_USERDEFINED = 29,
        VT_LPSTR = 30,
        VT_LPWSTR = 31,
        VT_RECORD = 36,
        VT_INT_PTR = 37,
        VT_UINT_PTR = 38,
        VT_FILETIME = 64,
        VT_BLOB = 65,
        VT_STREAM = 66,
        VT_STORAGE = 67,
        VT_STREAMED_OBJECT = 68,
        VT_STORED_OBJECT = 69,
        VT_BLOB_OBJECT = 70,
        VT_CF = 71,
        VT_CLSID = 72,
        VT_VERSIONED_STREAM = 73,
        VT_BSTR_BLOB = 0xfff,
        VT_VECTOR = 0x1000,
        VT_ARRAY = 0x2000,
        VT_BYREF = 0x4000

    };

    [System.Runtime.InteropServices.StructLayout(System.Runtime.InteropServices.LayoutKind.Explicit, Size = 16)]
    public struct Variant
    {
        [System.Runtime.InteropServices.FieldOffset(0)]
        public ushort vt;
        [System.Runtime.InteropServices.FieldOffset(2)]
        public ushort wReserved1;
        [System.Runtime.InteropServices.FieldOffset(4)]
        public ushort wReserved2;
        [System.Runtime.InteropServices.FieldOffset(6)]
        public ushort wReserved3;
        [System.Runtime.InteropServices.FieldOffset(8)]
        public short iVal;
        [System.Runtime.InteropServices.FieldOffset(8)]
        public float fltVal;
        [System.Runtime.InteropServices.FieldOffset(8)]
        public IntPtr data01;
        [System.Runtime.InteropServices.FieldOffset(12)]
        public IntPtr data02;
    }

    public struct SafeArrayBounds
    {
        public uint cElements;
        public uint lBound;
    }

    public class Flowsheet
    {
        const int LOCALE_SYSTEM_DEFAULT = 2048;
        const int DISPATCH_METHOD = 0x1;
        const int DISPATCH_PROPERTYGET = 0x2;
        const int DISPATCH_PROPERTYPUT = 0x4;
        const int SizeOfNativeVariant = 16;
        const int DISPID_PROPERTYPUT = -3;
        //CHEMCAD.IFlowsheet p_IFlowsheet;
        IDispatch p_IDispatch;

        public Flowsheet(object sheet)
        {
            //p_IFlowsheet = (CHEMCAD.IFlowsheet)sheet;
            p_IDispatch = (IDispatch)sheet;
        }

        //~Flowsheet()
        //{
        //    if (p_IFlowsheet != null)
        //        System.Runtime.InteropServices.Marshal.FinalReleaseComObject(p_IFlowsheet);
        //}

        public int NumberofStreams{
            get
            {
                Guid IID_NULL = new Guid("00000000-0000-0000-0000-000000000000");
                string[] rgsNames = new string[1] { "GetNoOfStreams" };
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

        public int NumberOfUnitOps
        {
            get
            {
                Guid IID_NULL = new Guid("00000000-0000-0000-0000-000000000000");
                string[] rgsNames = new string[1] { "GetNoOfUnitOps" };
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

        public int[] AllStreamIDs
        {
            get
            {
                Guid IID_NULL = new Guid("00000000-0000-0000-0000-000000000000");
                string[] rgsNames = new string[1] { "GetAllStreamIDs" };
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
                IntPtr psa = IntPtr.Zero;
                short numStreams = 0;

                if (hrRet == 0)
                {
                    System.Runtime.InteropServices.ComTypes.EXCEPINFO ExcepInfo = new System.Runtime.InteropServices.ComTypes.EXCEPINFO();
                    UInt32 pArgErr = 0;

                    SafeArrayBounds bounds = new SafeArrayBounds();
                    bounds.cElements = 1001;
                    bounds.lBound = 0;
                    psa = NativeMethods.SafeArrayCreate((int)(VARENUM.VT_I2), 1, ref bounds);
                    System.Runtime.InteropServices.GCHandle handle = System.Runtime.InteropServices.GCHandle.Alloc(psa, System.Runtime.InteropServices.GCHandleType.Pinned);

                    Variant[] v = new Variant[1];
                    v[0].vt = 0x6002;
                    v[0].data01 = handle.AddrOfPinnedObject();
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
                    numStreams = (short)varResult;
                    rgvarg.Free();
                    handle.Free();
                }

                if (hrRet == 0)
                {

                    retVal = new int[numStreams];
                    short val = 0;
                    System.Runtime.InteropServices.GCHandle pVal = System.Runtime.InteropServices.GCHandle.Alloc(val, System.Runtime.InteropServices.GCHandleType.Pinned);
                    for (int i = 0; i < numStreams; i++)
                    {
                        long longVal = i + 1;
                        hrRet = NativeMethods.SafeArrayGetElement(psa, ref longVal, pVal.AddrOfPinnedObject());
                        retVal[i] = (short)pVal.Target;
                    }
                    pVal.Free();
                }
                NativeMethods.SafeArrayDestroy(psa);
                return retVal;
            }
        }

        public int[] AllUnitOpIDs
        {
            get
            {
                Guid IID_NULL = new Guid("00000000-0000-0000-0000-000000000000");
                string[] rgsNames = new string[1] { "GetAllUnitOpIDs" };
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
                IntPtr psa = IntPtr.Zero;
                short numStreams = 0;

                if (hrRet == 0)
                {
                    System.Runtime.InteropServices.ComTypes.EXCEPINFO ExcepInfo = new System.Runtime.InteropServices.ComTypes.EXCEPINFO();
                    UInt32 pArgErr = 0;

                    SafeArrayBounds bounds = new SafeArrayBounds();
                    bounds.cElements = 1001;
                    bounds.lBound = 0;
                    psa = NativeMethods.SafeArrayCreate((int)(VARENUM.VT_I2), 1, ref bounds);
                    System.Runtime.InteropServices.GCHandle handle = System.Runtime.InteropServices.GCHandle.Alloc(psa, System.Runtime.InteropServices.GCHandleType.Pinned);

                    Variant[] v = new Variant[1];
                    v[0].vt = 0x6002;
                    v[0].data01 = handle.AddrOfPinnedObject();
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
                    numStreams = (short)varResult;
                    rgvarg.Free();
                    handle.Free();
                }

                if (hrRet == 0)
                {
                    retVal = new int[numStreams];

                    short val = 0;
                    System.Runtime.InteropServices.GCHandle pVal = System.Runtime.InteropServices.GCHandle.Alloc(val, System.Runtime.InteropServices.GCHandleType.Pinned);
                    for (int i = 0; i < numStreams; i++)
                    {
                        long longVal = i + 1;
                        hrRet = NativeMethods.SafeArrayGetElement(psa, ref longVal, pVal.AddrOfPinnedObject());
                        retVal[i] = (short)pVal.Target;
                    }
                    pVal.Free();
                }
                NativeMethods.SafeArrayDestroy(psa);
                return retVal;
            }
        }

        public void GetStreamCountsToUnitOp(int unitOpID, ref int nInlets, ref int nOutlets)
        {
            Guid IID_NULL = new Guid("00000000-0000-0000-0000-000000000000");
            string[] rgsNames = new string[1] { "GetStreamCountsToUnitOp" };
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

            if (hrRet == 0)
            {
                object varResult = null;
                System.Runtime.InteropServices.ComTypes.EXCEPINFO ExcepInfo = new System.Runtime.InteropServices.ComTypes.EXCEPINFO();
                UInt32 pArgErr = 0;

                SafeArrayBounds bounds = new SafeArrayBounds();
                bounds.cElements = 1001;
                bounds.lBound = 0;
                short outlets = 0;
                System.Runtime.InteropServices.GCHandle handle0 = System.Runtime.InteropServices.GCHandle.Alloc(outlets, System.Runtime.InteropServices.GCHandleType.Pinned);
                short inlets = 0;
                System.Runtime.InteropServices.GCHandle handle1 = System.Runtime.InteropServices.GCHandle.Alloc(inlets, System.Runtime.InteropServices.GCHandleType.Pinned);

                Variant[] v = new Variant[3];
                v[0].vt = 0x4002;
                v[0].data01 = handle0.AddrOfPinnedObject();
                v[0].vt = 0x4002;
                v[0].data01 = handle1.AddrOfPinnedObject();
                v[1].vt = 0x0002;
                v[1].iVal = (short)unitOpID;
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

                if (hrRet == 0)
                {
                    nOutlets = (short)handle0.Target;
                    nInlets = (short)handle1.Target;
                }
                rgvarg.Free();
                handle0.Free();
                handle1.Free();
            }
        }

        public int[] GetStreamIDsToUnitOp(int unitOpID)
        {
            Guid IID_NULL = new Guid("00000000-0000-0000-0000-000000000000");
            string[] rgsNames = new string[1] { "GetStreamIDsToUnitOp" };
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

            if (hrRet == 0)
            {
                object varResult = null;
                System.Runtime.InteropServices.ComTypes.EXCEPINFO ExcepInfo = new System.Runtime.InteropServices.ComTypes.EXCEPINFO();
                UInt32 pArgErr = 0;

                SafeArrayBounds bounds = new SafeArrayBounds();
                bounds.cElements = 1001;
                bounds.lBound = 0;
                IntPtr psa = NativeMethods.SafeArrayCreate((int)(VARENUM.VT_I2), 1, ref bounds);
                System.Runtime.InteropServices.GCHandle handle0 = System.Runtime.InteropServices.GCHandle.Alloc(psa, System.Runtime.InteropServices.GCHandleType.Pinned);

                Variant[] v = new Variant[2];
                v[0].vt = 0x6002;
                v[0].data01 = handle0.AddrOfPinnedObject();
                v[1].vt = 0x0002;
                v[1].iVal = (short)unitOpID;
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

                    short numStreams = (short)varResult;
                    retVal = new int[numStreams];
                    short val = 0;
                    System.Runtime.InteropServices.GCHandle pVal = System.Runtime.InteropServices.GCHandle.Alloc(val, System.Runtime.InteropServices.GCHandleType.Pinned);
                    for (int i = 0; i < numStreams; i++)
                    {
                        long longVal = i + 1;
                        hrRet = NativeMethods.SafeArrayGetElement(psa, ref longVal, pVal.AddrOfPinnedObject());
                        retVal[i] = (short)pVal.Target;
                    }
                    pVal.Free();
                }
                NativeMethods.SafeArrayDestroy(psa);
                rgvarg.Free();
                handle0.Free();
            }
            return retVal;
        }

        public void GetSourceAndTargetForStream(int streamID, ref int sourceID, ref int targetID)
        {
            Guid IID_NULL = new Guid("00000000-0000-0000-0000-000000000000");
            string[] rgsNames = new string[1] { "GetSourceAndTargetForStream" };
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

            if (hrRet == 0)
            {
                object varResult = null;
                System.Runtime.InteropServices.ComTypes.EXCEPINFO ExcepInfo = new System.Runtime.InteropServices.ComTypes.EXCEPINFO();
                UInt32 pArgErr = 0;

                SafeArrayBounds bounds = new SafeArrayBounds();
                bounds.cElements = 1001;
                bounds.lBound = 0;
                short target = 0;
                System.Runtime.InteropServices.GCHandle handle0 = System.Runtime.InteropServices.GCHandle.Alloc(target, System.Runtime.InteropServices.GCHandleType.Pinned);
                short source = 0;
                System.Runtime.InteropServices.GCHandle handle1 = System.Runtime.InteropServices.GCHandle.Alloc(source, System.Runtime.InteropServices.GCHandleType.Pinned);

                Variant[] v = new Variant[3];
                v[0].vt = 0x4002;
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

                if (hrRet == 0)
                {
                    targetID = (short)handle0.Target;
                    sourceID = (short)handle1.Target;
                }
                rgvarg.Free();
                handle0.Free();
                handle1.Free();
            }
        }

        public int NumberOfFeedStreams
        {
            get
            {
                Guid IID_NULL = new Guid("00000000-0000-0000-0000-000000000000");
                string[] rgsNames = new string[1] { "GetNoOfFeedStreams" };
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

        public int NumberOfProductStreams
        {
            get
            {
                Guid IID_NULL = new Guid("00000000-0000-0000-0000-000000000000");
                string[] rgsNames = new string[1] { "GetNoOfProductStreams" };
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

        public int[] FeedStreamIDs
        {
            get
            {
                Guid IID_NULL = new Guid("00000000-0000-0000-0000-000000000000");
                string[] rgsNames = new string[1] { "GetFeedStreamIDs" };
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
                IntPtr psa = IntPtr.Zero;
                short numStreams = 0;

                if (hrRet == 0)
                {
                    System.Runtime.InteropServices.ComTypes.EXCEPINFO ExcepInfo = new System.Runtime.InteropServices.ComTypes.EXCEPINFO();
                    UInt32 pArgErr = 0;

                    SafeArrayBounds bounds = new SafeArrayBounds();
                    bounds.cElements = 1001;
                    bounds.lBound = 0;
                    psa = NativeMethods.SafeArrayCreate((int)(VARENUM.VT_I2), 1, ref bounds);
                    System.Runtime.InteropServices.GCHandle handle = System.Runtime.InteropServices.GCHandle.Alloc(psa, System.Runtime.InteropServices.GCHandleType.Pinned);

                    Variant[] v = new Variant[1];
                    v[0].vt = 0x6002;
                    v[0].data01 = handle.AddrOfPinnedObject();
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
                    numStreams = (short)varResult;
                    rgvarg.Free();
                    handle.Free();
                }

                if (hrRet == 0)
                {

                    retVal = new int[numStreams];
                    short val = 0;
                    System.Runtime.InteropServices.GCHandle pVal = System.Runtime.InteropServices.GCHandle.Alloc(val, System.Runtime.InteropServices.GCHandleType.Pinned);
                    for (int i = 0; i < numStreams; i++)
                    {
                        long longVal = i + 1;
                        hrRet = NativeMethods.SafeArrayGetElement(psa, ref longVal, pVal.AddrOfPinnedObject());
                        retVal[i] = (short)pVal.Target;
                    }
                    pVal.Free();
                }
                NativeMethods.SafeArrayDestroy(psa);
                return retVal;
            }
        }

        public int[] ProductStreamIDs
        {
            get
            {

                Guid IID_NULL = new Guid("00000000-0000-0000-0000-000000000000");
                string[] rgsNames = new string[1] { "GetProductStreamIDs" };
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
                IntPtr psa = IntPtr.Zero;
                short numStreams = 0;

                if (hrRet == 0)
                {
                    System.Runtime.InteropServices.ComTypes.EXCEPINFO ExcepInfo = new System.Runtime.InteropServices.ComTypes.EXCEPINFO();
                    UInt32 pArgErr = 0;

                    SafeArrayBounds bounds = new SafeArrayBounds();
                    bounds.cElements = 1001;
                    bounds.lBound = 0;
                    psa = NativeMethods.SafeArrayCreate((int)(VARENUM.VT_I2), 1, ref bounds);
                    System.Runtime.InteropServices.GCHandle handle = System.Runtime.InteropServices.GCHandle.Alloc(psa, System.Runtime.InteropServices.GCHandleType.Pinned);

                    Variant[] v = new Variant[1];
                    v[0].vt = 0x6002;
                    v[0].data01 = handle.AddrOfPinnedObject();
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
                    numStreams = (short)varResult;
                    rgvarg.Free();
                    handle.Free();
                }

                if (hrRet == 0)
                {
                    retVal = new int[numStreams];

                    short val = 0;
                    System.Runtime.InteropServices.GCHandle pVal = System.Runtime.InteropServices.GCHandle.Alloc(val, System.Runtime.InteropServices.GCHandleType.Pinned);
                    for (int i = 0; i < numStreams; i++)
                    {
                        long longVal = i + 1;
                        hrRet = NativeMethods.SafeArrayGetElement(psa, ref longVal, pVal.AddrOfPinnedObject());
                        retVal[i] = (short)pVal.Target;
                    }
                    pVal.Free();
                }
                NativeMethods.SafeArrayDestroy(psa);
                return retVal;
            }
        }

        public int[] GetInletStreamIDsToUnitOp(int unitOpID)
        {
            Guid IID_NULL = new Guid("00000000-0000-0000-0000-000000000000");
            string[] rgsNames = new string[1] { "GetInletStreamIDsToUnitOp" };
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

            if (hrRet == 0)
            {
                object varResult = null;
                System.Runtime.InteropServices.ComTypes.EXCEPINFO ExcepInfo = new System.Runtime.InteropServices.ComTypes.EXCEPINFO();
                UInt32 pArgErr = 0;

                SafeArrayBounds bounds = new SafeArrayBounds();
                bounds.cElements = 1001;
                bounds.lBound = 0;
                IntPtr psa = NativeMethods.SafeArrayCreate((int)(VARENUM.VT_I2), 1, ref bounds);
                System.Runtime.InteropServices.GCHandle handle0 = System.Runtime.InteropServices.GCHandle.Alloc(psa, System.Runtime.InteropServices.GCHandleType.Pinned);

                Variant[] v = new Variant[2];
                v[0].vt = 0x6002;
                v[0].data01 = handle0.AddrOfPinnedObject();
                v[1].vt = 0x0002;
                v[1].iVal = (short)unitOpID;
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
                    short numStreams = (short)varResult;
                    retVal = new int[numStreams];

                    short val = 0;
                    System.Runtime.InteropServices.GCHandle pVal = System.Runtime.InteropServices.GCHandle.Alloc(val, System.Runtime.InteropServices.GCHandleType.Pinned);
                    for (int i = 0; i < numStreams; i++)
                    {
                        long longVal = i + 1;
                        hrRet = NativeMethods.SafeArrayGetElement(psa, ref longVal, pVal.AddrOfPinnedObject());
                        retVal[i] = (short)pVal.Target;
                    }
                    pVal.Free();
                }
                NativeMethods.SafeArrayDestroy(psa);
                rgvarg.Free();
                handle0.Free();
            }
            return retVal;
        }

        public int[] GetOutletStreamIDsToUnitOp(int unitOpID)
        {
            Guid IID_NULL = new Guid("00000000-0000-0000-0000-000000000000");
            string[] rgsNames = new string[1] { "GetOutletStreamIDsToUnitOp" };
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

            if (hrRet == 0)
            {
                object varResult = null;
                System.Runtime.InteropServices.ComTypes.EXCEPINFO ExcepInfo = new System.Runtime.InteropServices.ComTypes.EXCEPINFO();
                UInt32 pArgErr = 0;

                SafeArrayBounds bounds = new SafeArrayBounds();
                bounds.cElements = 1001;
                bounds.lBound = 0;
                IntPtr psa = NativeMethods.SafeArrayCreate((int)(VARENUM.VT_I2), 1, ref bounds);
                System.Runtime.InteropServices.GCHandle handle0 = System.Runtime.InteropServices.GCHandle.Alloc(psa, System.Runtime.InteropServices.GCHandleType.Pinned);

                Variant[] v = new Variant[2];
                v[0].vt = 0x6002;
                v[0].data01 = handle0.AddrOfPinnedObject();
                v[1].vt = 0x0002;
                v[1].iVal = (short)unitOpID;
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
                    short numStreams = (short)varResult;
                    retVal = new int[numStreams];

                    short val = 0;
                    System.Runtime.InteropServices.GCHandle pVal = System.Runtime.InteropServices.GCHandle.Alloc(val, System.Runtime.InteropServices.GCHandleType.Pinned);
                    for (int i = 0; i < numStreams; i++)
                    {
                        long longVal = i + 1;
                        hrRet = NativeMethods.SafeArrayGetElement(psa, ref longVal, pVal.AddrOfPinnedObject());
                        retVal[i] = (short)pVal.Target;
                    }
                    pVal.Free();
                }
                NativeMethods.SafeArrayDestroy(psa);
                rgvarg.Free();
                handle0.Free();
            }
            return retVal;
        }

        public void GetDynamicTime(ref float dynTime, ref String timeUnit)
        {
            Guid IID_NULL = new Guid("00000000-0000-0000-0000-000000000000");
            string[] rgsNames = new string[1] { "GetDynamicTime" };
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

            if (hrRet == 0)
            {
                object varResult = null;
                System.Runtime.InteropServices.ComTypes.EXCEPINFO ExcepInfo = new System.Runtime.InteropServices.ComTypes.EXCEPINFO();
                UInt32 pArgErr = 0;

                SafeArrayBounds bounds = new SafeArrayBounds();
                bounds.cElements = 1001;
                bounds.lBound = 0;
                IntPtr pString0 = System.Runtime.InteropServices.Marshal.StringToBSTR(String.Empty);
                System.Runtime.InteropServices.GCHandle handle0 = System.Runtime.InteropServices.GCHandle.Alloc(pString0, System.Runtime.InteropServices.GCHandleType.Pinned);
                float time = 0;
                System.Runtime.InteropServices.GCHandle handle1 = System.Runtime.InteropServices.GCHandle.Alloc(time, System.Runtime.InteropServices.GCHandleType.Pinned);

                Variant[] v = new Variant[3];
                v[0].vt = 0x4008;
                v[0].data01 = handle0.AddrOfPinnedObject();
                v[0].vt = 0x4004;
                v[0].data01 = handle1.AddrOfPinnedObject();
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

                if (hrRet == 0)
                {
                    timeUnit = System.Runtime.InteropServices.Marshal.PtrToStringBSTR((IntPtr)handle0.Target);
                    dynTime = (float)handle1.Target;
                }
                rgvarg.Free();
                handle0.Free();
                handle1.Free();
            }
        }

        public int NumberOfCutStreams
        {
            get
            {
                Guid IID_NULL = new Guid("00000000-0000-0000-0000-000000000000");
                string[] rgsNames = new string[1] { "GetNoOfCutStreams" };
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

        public int[] CutStreamsIDs
        {
            get
            {
                Guid IID_NULL = new Guid("00000000-0000-0000-0000-000000000000");
                string[] rgsNames = new string[1] { "GetCutStreamsIDs" };
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
                IntPtr psa = IntPtr.Zero;

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
                    short[] streams = (short[])varResult;
                    retVal = new int[streams.Length];
                    for (int i = 0; i < streams.Length; i++ )
                    {
                        retVal[i] = streams[i];
                    }
                }
                return retVal;
            }
        }
    }
}
