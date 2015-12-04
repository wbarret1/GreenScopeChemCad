using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
// using System.Threading.Tasks;

namespace GreenScopeChemCad
{
    class StreamProperty
    {
         //CHEMCAD.IStreamProperty p_IStreamProperty;
        IDispatch p_IDispatch;

        const int LOCALE_SYSTEM_DEFAULT = 2048;
        const int DISPATCH_METHOD = 0x1;
        const int DISPATCH_PROPERTYGET = 0x2;
        const int DISPATCH_PROPERTYPUT = 0x4;
        const int SizeOfNativeVariant = 16;
        const int DISPID_PROPERTYPUT = -3;

        public StreamProperty(object streamProp)
        {
            //p_IStreamProperty = (CHEMCAD.IStreamProperty)streamProp;
            p_IDispatch = (IDispatch)streamProp;
        }

        //~StreamProperty()
        //{
        //    //if (p_IFlowsheet != null)
        //    //    System.Runtime.InteropServices.Marshal.FinalReleaseComObject(p_IFlowsheet);
        //}

        public void GetStreamPropertiesInUserUnits(int streamID, ref double[] propertyValues, ref string[] propertyUnits)
        {
            Guid IID_NULL = new Guid("00000000-0000-0000-0000-000000000000");
            string[] rgsNames = new string[1] { "GetStreamPropertiesInUserUnits" };
            int[] rgDispId = new int[1] { 0 };
            int[] retVal = { 0 };
            object varResult = null;

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
                System.Runtime.InteropServices.ComTypes.EXCEPINFO ExcepInfo = new System.Runtime.InteropServices.ComTypes.EXCEPINFO();
                UInt32 pArgErr = 0;

                SafeArrayBounds bounds = new SafeArrayBounds();
                bounds.cElements = 61;
                bounds.lBound = 1;
                IntPtr psa0 = NativeMethods.SafeArrayCreate((ushort)(VARENUM.VT_BSTR), 1, ref bounds);
                System.Runtime.InteropServices.GCHandle handle0 = System.Runtime.InteropServices.GCHandle.Alloc(psa0, System.Runtime.InteropServices.GCHandleType.Pinned);

                IntPtr psa1 = NativeMethods.SafeArrayCreate((ushort)(VARENUM.VT_R4), 1, ref bounds);
                System.Runtime.InteropServices.GCHandle handle1 = System.Runtime.InteropServices.GCHandle.Alloc(psa1, System.Runtime.InteropServices.GCHandleType.Pinned);

                Variant[] v = new Variant[3];
                v[0].vt = 0x6008;
                v[0].data01 = handle0.AddrOfPinnedObject();
                v[1].vt = 0x6004;
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

                    short numProps = (short)varResult;
                    propertyValues = new double[numProps + 1];
                    propertyUnits = new string[numProps + 1];
                    float val = 0.0F;
                    //object obj = null;
                    IntPtr pString = System.Runtime.InteropServices.Marshal.StringToBSTR(String.Empty);
                    System.Runtime.InteropServices.GCHandle pVal = System.Runtime.InteropServices.GCHandle.Alloc(val, System.Runtime.InteropServices.GCHandleType.Pinned);
                    System.Runtime.InteropServices.GCHandle pVal1 = System.Runtime.InteropServices.GCHandle.Alloc(pString, System.Runtime.InteropServices.GCHandleType.Pinned);
                    for (int i = 0; i < numProps + 1; i++)
                    {
                        long longVal = i;
                        hrRet = NativeMethods.SafeArrayGetElement(psa1, ref longVal, pVal.AddrOfPinnedObject());
                        propertyValues[i] = (float)pVal.Target;
                        hrRet = NativeMethods.SafeArrayGetElement(psa0, ref longVal, pVal1.AddrOfPinnedObject());
                        propertyUnits[i] = System.Runtime.InteropServices.Marshal.PtrToStringBSTR((IntPtr)pVal1.Target);
                    }
                    pVal.Free();
                    pVal1.Free();
                }
                NativeMethods.SafeArrayDestroy(psa0);
                NativeMethods.SafeArrayDestroy(psa1);
                rgvarg.Free();
                handle0.Free();
            }
        }

   }
}
