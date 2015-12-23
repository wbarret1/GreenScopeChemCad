using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
//// using System.Threading.Tasks;

namespace GreenScopeChemCad
{
    enum ServerStreamStatus
    {
        STREAM_STATUS_NO_ERROR = 0,
        STREAM_STATUS_NO_MEM = 1,
        STREAM_STATUS_PARAMETER_TYPE_MISMATCH = 2,
        STREAM_STATUS_PARAMETER_INVALID = 3,
        STREAM_STATUS_FLASH_CALC_FAILED = 4
    };

    enum FlashSpecification
    {
        FLASH_SPEC_NOP = 0,
        FLASH_SPEC_TP = 1,
        FLASH_SPEC_VP_FLASH = 2,
        FLASH_SPEC_VT_FLASH = 3,
        FLASH_SPEC_HP_FLASH = 4
    };

    enum TotalFlowSpecification
    {
        SPEC_MOLE_RATE = 0,
        SPEC_MASS_RATE = 1,
        SPEC_LIQ_VOL_RATE = 2,
        SPEC_VAP_VOL_RATE = 3
    };

    enum ComponentFlowSpecType
    {
        SPEC_COMP_MOLE_RATE = 0,
        SPEC_COMP_MASS_RATE = 1,
        SPEC_COMP_LIQ_VOL = 2,
        SPEC_COMP_VAP_VOL = 3,
        SPEC_COMP_MOLE_FRAC = 4,
        SPEC_COMP_MASS_FRAC = 5,
        SPEC_COMP_LIQ_VOL_FRAC = 6,
        SPEC_COMP_VAP_VOL_FRAC = 7,
        SPEC_COMP_MOLE_PERCENT = 8,
        SPEC_COMP_MASS_PERCENT = 9,
        SPEC_COMP_LIQ_VOL_PERCENT = 10,
        SPEC_COMP_VAP_VOL_PERCENT = 11,
        SPEC_COMP_MOLE_PPM = 12,
        SPEC_COMP_MASS_PPM = 13,
        SPEC_COMP_LIQ_VOL_PPM = 14,
        SPEC_COMP_VAP_VOL_PPM = 15,
        SPEC_COMP_CONCENTRATION = 16,
        SPEC_COMP_K_VALUES = 17,
        SPEC_COMP_ACTIVITY_COEFF = 18,
        SPEC_COMP_FUGACITY_COEFF = 19
    };

    enum TotalStreamPropertyDescription
    {
        Temperature = 1,
        Pressure = 2,
        MoleVaporFraction = 3,
        Enthalpy = 4,
        TotalMoleRate = 5,
        TotalMassRate = 6,
        TotalStdLiquidVolumeRate = 7,
        TotalStdVaporVolumeRate = 8,
        TotalActualVolumeRate = 9,
        TotalActualDensity = 10,
        TotalMW = 11,
        GrossHValue = 12,
        NetHValue = 13,
        ReidVaporPpressure = 14,
        UOPK = 15,
        VABP = 16,
        MeABP = 17,
        FlashPoint = 18,
        PourPoint = 19,
        TotalEntropy = 20,
        MassVaporFraction = 21,
        PHValue = 22,
        VaporMoleRate = 26,
        VaporMassRate = 27,
        VaporEnthalpy = 28,
        VaporEntropy = 29,
        VaporMW = 30,
        VaporActualDensity = 31,
        VaporActualVolumeRate = 32,
        VapStdLiquidVolumeRate = 33,
        VapStdVaporVolumeRate = 34,
        VaporCP = 35,
        VaporZ = 36,
        VaporViscosity = 37,
        VaporThermalConductivity = 38,
        CpOverCv = 39,
        LiquidMoleRate = 41,
        LiquidMassRate = 42,
        LiquidEnthalpy = 43,
        LiquidEntropy = 44,
        LiquidMW = 45,
        LiquidActualDensity = 46,
        LiquidActualVolumeRate = 47,
        LiquidStdLiquidVolumeRate = 48,
        LiquidStdVaporVolumeRate = 49,
        LiquidCP = 50,
        LiquidZ = 51,
        LiquidViscosity = 52,
        LiquidThermalConductivity = 53,
        LiquidSurfaceTension = 54,
        LiquidAndSolidActualDensity = 55,
        LiquidAndSolidActualVolume = 56,
        LiquidLatentHeat = 57,
        SolidMoleRate = 60,
        SolidMassRate = 61,
        SolidMW = 62,
        SolidEnthalpy = 63,
        SolidCP = 64,
        SolidActualVolume = 65,
        SolidDensity = 66,
        SolidStdVaporVolumeRate = 67,
        SolidStdLiquidVolumeRate = 68
    };

    enum ComponentStreamPropertyDescription
    {
        MassFlowRateIthComponent = 200,
        StdLiquidVolumeFlowTateIthComponent = 400,
        MoleFractionIthComponent = 600,
        MassFractionIthcomponent = 800,
        StdLiquidVolumeFractionIthComponent = 1000
    };

    [System.Runtime.Serialization.DataContract]
    class Stream
    {
        StreamInfo p_StreamInfo;
        StreamProperty p_StreamProperty;
        Flowsheet p_Flowsheet;
        IDispatch p_Flash;

        [System.Runtime.Serialization.DataMember]
        int m_StreamID;

        [System.Runtime.Serialization.DataMember]
        string m_StreamName;

        [System.Runtime.Serialization.DataMember]
        StreamComponent[] m_Components;
        int[] m_ComponentIDs;
        string[] m_ComponentNames;

        [System.Runtime.Serialization.DataMember]
        int m_SourceUnitOperation;

        [System.Runtime.Serialization.DataMember]
        int m_TargetUnitOperation;

        [System.Runtime.Serialization.DataMember]
        double m_Temperature = 0;

        [System.Runtime.Serialization.DataMember]
        double m_Pressure = 0;

        [System.Runtime.Serialization.DataMember]
        double m_MoleVapFrac = 0;

        [System.Runtime.Serialization.DataMember]
        double m_Enthalpy = 0;

        [System.Runtime.Serialization.DataMember]
        string m_TemperatureUnit = string.Empty;

        [System.Runtime.Serialization.DataMember]
        string m_PressureUnit = string.Empty;

        [System.Runtime.Serialization.DataMember]
        string m_EnthalpyUnit = string.Empty;

        [System.Runtime.Serialization.DataMember]
        double m_TotalMassFlowRate = 0;

        [System.Runtime.Serialization.DataMember]
        string m_TotalMassFlowRateUnit = string.Empty;

        [System.Runtime.Serialization.DataMember]
        double m_TotalMoleFlowRate = 0;

        [System.Runtime.Serialization.DataMember]
        string m_TotalMoleFlowRateUnit = string.Empty;

        [System.Runtime.Serialization.DataMember]
        double m_LiquidVolumetricFlowRate = 0;

        [System.Runtime.Serialization.DataMember]
        string m_TotalLiquidVolumeFlowRateUnit = string.Empty;

        [System.Runtime.Serialization.DataMember]
        double m_VaporVolumetricFlowRate = 0;

        [System.Runtime.Serialization.DataMember]
        string m_TotalVaporVolumeFlowRateUnit = string.Empty;

        [System.Runtime.Serialization.DataMember]
        string m_ComponentMassFlowRateUnit = string.Empty;

        [System.Runtime.Serialization.DataMember]
        string m_ComponentMoleFlowRateUnit = string.Empty;

        [System.Runtime.Serialization.DataMember]
        string m_ComponentVolumeFlowRateUnit = string.Empty;

        [System.Runtime.Serialization.DataMember]
        double[] m_CompMassFraction = null;

        [System.Runtime.Serialization.DataMember]
        double[] m_CompMoleFraction = null;

        [System.Runtime.Serialization.DataMember]
        double[] m_CompMoleFlow = null;

        [System.Runtime.Serialization.DataMember]
        double[] m_CompMassFlowRate = null;

        [System.Runtime.Serialization.DataMember]
        int m_CostType = 0;
        double m_Cost = 0;
        double[] m_StreamPropertyValues = null;
        string[] m_StreamPropertyUnits = null;
        double m_TempR;
        double m_PressPsia;
        double m_Enth_BTUHR;
        //double m_MoleVapFrac;
        double[] m_CompFlowLbMolHr;
        double[] m_VaporCompFlowLbMolHr;
        double[] m_LiquidCompFlowLbMolHr;

        const int LOCALE_SYSTEM_DEFAULT = 2048;
        const int DISPATCH_METHOD = 0x1;
        const int DISPATCH_PROPERTYGET = 0x2;
        const int DISPATCH_PROPERTYPUT = 0x4;
        const int SizeOfNativeVariant = 16;
        const int DISPID_PROPERTYPUT = -3;


        public Stream(int StreamID, VBServerWrapper vbServer)
        {
            m_StreamID = StreamID;
            p_StreamInfo = vbServer.GetStreamInfo();
            p_Flowsheet = vbServer.GetFlowsheet();
            p_StreamProperty = vbServer.GetStreamProperty();
            m_StreamName = p_StreamInfo.GetStreamLabelByID(m_StreamID);
            p_Flowsheet.GetSourceAndTargetForStream(m_StreamID, ref m_SourceUnitOperation, ref m_TargetUnitOperation);
            double[] compFlows = null;
            string compCate = null;
            p_StreamInfo.GetCurUserUnitString(ref m_TemperatureUnit, ref m_PressureUnit, ref m_EnthalpyUnit, ref m_TotalMoleFlowRateUnit, ref m_TotalMassFlowRateUnit, ref m_TotalLiquidVolumeFlowRateUnit, ref m_TotalVaporVolumeFlowRateUnit, ref m_ComponentMassFlowRateUnit, ref compCate);
            p_StreamInfo.GetStreamInCurUserUnitByID(m_StreamID, ref m_Temperature, ref m_Pressure, ref m_Enthalpy, ref m_MoleVapFrac, ref m_TotalMoleFlowRate, ref m_TotalMassFlowRate, ref m_LiquidVolumetricFlowRate, ref m_VaporVolumetricFlowRate, ref m_CompMoleFraction);
            double tempR = 0.0;
            double pressPsia = 0.0;
            double mvf = 0.0;
            double enth_BTUHR = 0.0;
            p_StreamInfo.GetStreamByID(m_StreamID, ref tempR, ref pressPsia, ref mvf, ref enth_BTUHR, ref m_CompMoleFlow);
            m_CompMoleFlow = (double[])compFlows;
            p_StreamInfo.getCurrentEngineeringUnitLabels(ref m_TemperatureUnit, ref m_PressureUnit, ref m_EnthalpyUnit, TotalFlowSpecification.SPEC_MASS_RATE, ref m_TotalMassFlowRateUnit, ComponentFlowSpecType.SPEC_COMP_MASS_RATE, ref m_ComponentMassFlowRateUnit);
            p_StreamInfo.getStreamCurrentEngineeringUnits(m_StreamID, ref m_Temperature, ref m_Pressure, ref m_Enthalpy, ref m_MoleVapFrac, TotalFlowSpecification.SPEC_MASS_RATE, ref m_TotalMassFlowRate, ComponentFlowSpecType.SPEC_COMP_MASS_RATE, ref compFlows);
            m_CompMassFlowRate = (double[])compFlows;
            p_StreamInfo.getCurrentEngineeringUnitLabels(ref m_TemperatureUnit, ref m_PressureUnit, ref m_EnthalpyUnit, TotalFlowSpecification.SPEC_MASS_RATE, ref m_TotalMassFlowRateUnit, ComponentFlowSpecType.SPEC_COMP_MOLE_RATE, ref m_ComponentMoleFlowRateUnit);
            p_StreamInfo.getStreamCurrentEngineeringUnits(m_StreamID, ref m_Temperature, ref m_Pressure, ref m_Enthalpy, ref m_MoleVapFrac, TotalFlowSpecification.SPEC_MOLE_RATE, ref m_TotalMoleFlowRate, ComponentFlowSpecType.SPEC_COMP_MASS_FRAC, ref compFlows);
            m_CompMassFraction = (double[])compFlows;
            p_StreamInfo.getCurrentEngineeringUnitLabels(ref m_TemperatureUnit, ref m_PressureUnit, ref m_EnthalpyUnit, TotalFlowSpecification.SPEC_LIQ_VOL_RATE, ref m_TotalLiquidVolumeFlowRateUnit, ComponentFlowSpecType.SPEC_COMP_VAP_VOL, ref m_ComponentVolumeFlowRateUnit);
            p_StreamInfo.getStreamCurrentEngineeringUnits(m_StreamID, ref m_Temperature, ref m_Pressure, ref m_Enthalpy, ref m_MoleVapFrac, TotalFlowSpecification.SPEC_LIQ_VOL_RATE, ref m_LiquidVolumetricFlowRate, ComponentFlowSpecType.SPEC_COMP_MOLE_RATE, ref compFlows);
            m_CompMoleFlow = (double[])compFlows;
            p_StreamInfo.getCurrentEngineeringUnitLabels(ref m_TemperatureUnit, ref m_PressureUnit, ref m_EnthalpyUnit, TotalFlowSpecification.SPEC_VAP_VOL_RATE, ref m_TotalVaporVolumeFlowRateUnit, ComponentFlowSpecType.SPEC_COMP_MASS_RATE, ref m_ComponentMassFlowRateUnit);
            p_StreamInfo.getStreamCurrentEngineeringUnits(m_StreamID, ref m_Temperature, ref m_Pressure, ref m_Enthalpy, ref m_MoleVapFrac, TotalFlowSpecification.SPEC_VAP_VOL_RATE, ref m_VaporVolumetricFlowRate, ComponentFlowSpecType.SPEC_COMP_MOLE_FRAC, ref compFlows);
            m_CompMoleFraction = (double[])compFlows;
            p_StreamInfo.GetStreamCost(m_StreamID, ref m_CostType, ref m_Cost);
            p_StreamProperty.GetStreamPropertiesInUserUnits(m_StreamID, ref m_StreamPropertyValues, ref m_StreamPropertyUnits);
            m_Components = vbServer.Components;
            m_ComponentNames = vbServer.ComponentNames;
            m_ComponentIDs = vbServer.ComponentIDs;
        }

        void defineFeedStream(double tempR, double pressPsia, double enth_BTUHR, double[] CompFlowLbMolHr)
        {
            Guid IID_NULL = new Guid("00000000-0000-0000-0000-000000000000");
            string[] rgsNames = new string[1] { "DefineFeedStream" };
            int[] rgDispId = new int[1] { 0 };
            int[] retVal = { 0 };

            int hrRet = p_Flash.GetIDsOfNames
            (
                ref IID_NULL,
                rgsNames,
                1,
                LOCALE_SYSTEM_DEFAULT,
                rgDispId
            );
            object varResult = null;

            SafeArrayBounds bounds = new SafeArrayBounds();
            bounds.cElements = (uint)CompFlowLbMolHr.Length;
            bounds.lBound = 1;
            IntPtr psa = NativeMethods.SafeArrayCreate((ushort)(VARENUM.VT_R4), 1, ref bounds);
            float val = 0.0F;
            System.Runtime.InteropServices.GCHandle pVal = System.Runtime.InteropServices.GCHandle.Alloc(val, System.Runtime.InteropServices.GCHandleType.Pinned);
            for (long i = 0; i < CompFlowLbMolHr.Length; i++)
            {
                long longVal = i;
                val = (float)CompFlowLbMolHr[i];
                hrRet = NativeMethods.SafeArrayPutElement(psa, ref longVal, pVal.AddrOfPinnedObject());
            }
            pVal.Free();
            System.Runtime.InteropServices.GCHandle handle0 = System.Runtime.InteropServices.GCHandle.Alloc(psa, System.Runtime.InteropServices.GCHandleType.Pinned);

            if (hrRet == 0)
            {
                System.Runtime.InteropServices.ComTypes.EXCEPINFO ExcepInfo = new System.Runtime.InteropServices.ComTypes.EXCEPINFO();
                UInt32 pArgErr = 0;

                Variant[] v = new Variant[4];
                v[0].vt = 0x4004;
                v[0].data01 = handle0.AddrOfPinnedObject();
                v[1].vt = (ushort)(VARENUM.VT_R4);
                v[1].fltVal = (float)enth_BTUHR;
                v[2].vt = (ushort)(VARENUM.VT_R4);
                v[2].fltVal = (float)pressPsia;
                v[3].vt = (ushort)(VARENUM.VT_R4);
                v[3].fltVal = (float)tempR;
                System.Runtime.InteropServices.GCHandle rgvarg = System.Runtime.InteropServices.GCHandle.Alloc(v, System.Runtime.InteropServices.GCHandleType.Pinned);

                var dispParams = new System.Runtime.InteropServices.ComTypes.DISPPARAMS()
                {
                    cArgs = 4,
                    cNamedArgs = 0,
                    rgdispidNamedArgs = IntPtr.Zero,
                    rgvarg = rgvarg.AddrOfPinnedObject()
                };

                hrRet = p_Flash.Invoke
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
            NativeMethods.SafeArrayDestroy(psa);
            handle0.Free();
        }

        void CalculateHPFlash(double pressPsia, double enth_BTUHR)
        {
            Guid IID_NULL = new Guid("00000000-0000-0000-0000-000000000000");
            string[] rgsNames = new string[1] { "CalculateHPFlash" };
            int[] rgDispId = new int[1] { 0 };
            int[] retVal = { 0 };

            int hrRet = p_Flash.GetIDsOfNames
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

                Variant[] v = new Variant[2];
                v[0].vt = (ushort)(VARENUM.VT_R4);
                v[0].fltVal = (float)pressPsia;
                v[1].vt = (ushort)(VARENUM.VT_R4);
                v[1].fltVal = (float)enth_BTUHR;
                System.Runtime.InteropServices.GCHandle rgvarg = System.Runtime.InteropServices.GCHandle.Alloc(v, System.Runtime.InteropServices.GCHandleType.Pinned);

                var dispParams = new System.Runtime.InteropServices.ComTypes.DISPPARAMS()
                {
                    cArgs = 2,
                    cNamedArgs = 0,
                    rgdispidNamedArgs = IntPtr.Zero,
                    rgvarg = rgvarg.AddrOfPinnedObject()
                };

                hrRet = p_Flash.Invoke
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

        public void GetVaporStream(ref double tempR, ref double pressPsia, ref double enth_BTUHR, ref double rateLbMolHr, ref double[] compFlowLbMolHr)
        {
            Guid IID_NULL = new Guid("00000000-0000-0000-0000-000000000000");
            string[] rgsNames = new string[1] { "GetVaporStream" };
            int[] rgDispId = new int[1] { 0 };
            int[] retVal = { 0 };

            int hrRet = p_Flash.GetIDsOfNames
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
            float fRateLbMolHr = 0.0F;
            System.Runtime.InteropServices.GCHandle handle1 = System.Runtime.InteropServices.GCHandle.Alloc(fRateLbMolHr, System.Runtime.InteropServices.GCHandleType.Pinned);
            float fEnthalpy = 0.0F;
            System.Runtime.InteropServices.GCHandle handle2 = System.Runtime.InteropServices.GCHandle.Alloc(fEnthalpy, System.Runtime.InteropServices.GCHandleType.Pinned);
            float fPress = 0.0F;
            System.Runtime.InteropServices.GCHandle handle3 = System.Runtime.InteropServices.GCHandle.Alloc(fPress, System.Runtime.InteropServices.GCHandleType.Pinned);
            float fTemp = 0.0F;
            System.Runtime.InteropServices.GCHandle handle4 = System.Runtime.InteropServices.GCHandle.Alloc(fTemp, System.Runtime.InteropServices.GCHandleType.Pinned);

            if (hrRet == 0)
            {
                System.Runtime.InteropServices.ComTypes.EXCEPINFO ExcepInfo = new System.Runtime.InteropServices.ComTypes.EXCEPINFO();
                UInt32 pArgErr = 0;

                Variant[] v = new Variant[5];
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
                System.Runtime.InteropServices.GCHandle rgvarg = System.Runtime.InteropServices.GCHandle.Alloc(v, System.Runtime.InteropServices.GCHandleType.Pinned);

                var dispParams = new System.Runtime.InteropServices.ComTypes.DISPPARAMS()
                {
                    cArgs = 5,
                    cNamedArgs = 0,
                    rgdispidNamedArgs = IntPtr.Zero,
                    rgvarg = rgvarg.AddrOfPinnedObject()
                };

                hrRet = p_Flash.Invoke
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
                enth_BTUHR = (float)handle2.Target;
                rateLbMolHr = (float)handle1.Target;
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

        public void GetLiquidStream(ref double tempR, ref double pressPsia, ref double enth_BTUHR, ref double rateLbMolHr, ref double[] compFlowLbMolHr)
        {
            Guid IID_NULL = new Guid("00000000-0000-0000-0000-000000000000");
            string[] rgsNames = new string[1] { "GetLiquidStream" };
            int[] rgDispId = new int[1] { 0 };
            int[] retVal = { 0 };

            int hrRet = p_Flash.GetIDsOfNames
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
            float fRateLbMolHr = 0.0F;
            System.Runtime.InteropServices.GCHandle handle1 = System.Runtime.InteropServices.GCHandle.Alloc(fRateLbMolHr, System.Runtime.InteropServices.GCHandleType.Pinned);
            float fEnthalpy = 0.0F;
            System.Runtime.InteropServices.GCHandle handle2 = System.Runtime.InteropServices.GCHandle.Alloc(fEnthalpy, System.Runtime.InteropServices.GCHandleType.Pinned);
            float fPress = 0.0F;
            System.Runtime.InteropServices.GCHandle handle3 = System.Runtime.InteropServices.GCHandle.Alloc(fPress, System.Runtime.InteropServices.GCHandleType.Pinned);
            float fTemp = 0.0F;
            System.Runtime.InteropServices.GCHandle handle4 = System.Runtime.InteropServices.GCHandle.Alloc(fTemp, System.Runtime.InteropServices.GCHandleType.Pinned);

            if (hrRet == 0)
            {
                System.Runtime.InteropServices.ComTypes.EXCEPINFO ExcepInfo = new System.Runtime.InteropServices.ComTypes.EXCEPINFO();
                UInt32 pArgErr = 0;

                Variant[] v = new Variant[5];
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
                System.Runtime.InteropServices.GCHandle rgvarg = System.Runtime.InteropServices.GCHandle.Alloc(v, System.Runtime.InteropServices.GCHandleType.Pinned);

                var dispParams = new System.Runtime.InteropServices.ComTypes.DISPPARAMS()
                {
                    cArgs = 5,
                    cNamedArgs = 0,
                    rgdispidNamedArgs = IntPtr.Zero,
                    rgvarg = rgvarg.AddrOfPinnedObject()
                };

                hrRet = p_Flash.Invoke
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
                enth_BTUHR = (float)handle2.Target;
                rateLbMolHr = (float)handle1.Target;
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

        public int StreamID
        {
            get
            {
                return m_StreamID;
            }
        }

        public string StreamName
        {
            get
            {
                return m_StreamName;
            }
        }

        public int[] ComponentIDs
        {
            get
            {
                return m_ComponentIDs;
            }
        }

        public string[] ComponentNames
        {
            get
            {
                return m_ComponentNames;
            }
        }

        public int NumberOfComponents
        {
            get
            {
                return this.m_Components.Length;
            }
        }

        public string casNumber(int index)
        {
            return this.m_Components[index].CASNumber;
        }

        public string ERPG2(int index)
        {
            return this.m_Components[index].ERPG2;
        }

        public string ERPG3(int index)
        {
            return this.m_Components[index].ERPG3;
        }

        public string IDLH(int index)
        {
            return this.m_Components[index].IDLHvalue;
        }

        public string MAK(int index)
        {
            return this.m_Components[index].MAK;
        }

        public bool IsHazarous(int index)
        {
            return this.m_Components[index].Hazarous;
        }

        public bool OnTRIList(int index)
        {
            return this.m_Components[index].IsOnTRIList;
        }

        public bool IsPBTList(int index)
        {
            return this.m_Components[index].IsPBT;
        }

        public string EC_Class(int index)
        {
            return this.m_Components[index].EC_Class;
        }

        public string R_Phrase(int index)
        {
            return this.m_Components[index].RPhrase;
        }

        public string NFPA_Flammable(int index)
        {
            return this.m_Components[index].NFPAFlammability;
        }

        public string NFPA_Reactive(int index)
        {
            return this.m_Components[index].NFPAReactivity;
        }

        public string MolecularFormula(int index)
        {
            return this.m_Components[index].MolecularFormula;
        }

        public double MolecularWeight(int index)
        {
            return this.m_Components[index].MolecularWeight;
        }

        public string FlashPoint(int index)
        {
            return this.m_Components[index].FlashPoint;
        }

        public string HeatOfCombustion(int index)
        {
            return this.m_Components[index].heatOfCombustion;
        }

        public string HeatOfVaporization(int index)
        {
            return this.m_Components[index].HeatOfVaporization;
        }

        public string Density(int index)
        {
            return this.m_Components[index].Density;
        }

        public string VaporPressure(int index)
        {
            return this.m_Components[index].VaporPressure;
        }

        public int NumberOfCarbonAtoms(int index)
        {
            return this.m_Components[index].CarbonAtoms;
        }

        public int NumberOfHydrogenAtoms(int index)
        {
            return this.m_Components[index].HydrogenAtoms;
        }

        public int NumberOfNitrogenAtoms(int index)
        {
            return this.m_Components[index].NitrogenAtoms;
        }

        public int NumberOfChlorineAtoms(int index)
        {
            return this.m_Components[index].ChlorineAtoms;
        }

        public int NumberOfSodiumAtoms(int index)
        {
            return this.m_Components[index].SodiumAtoms;
        }

        public int NumberOfOxygenAtoms(int index)
        {
            return this.m_Components[index].OxygenAtoms;
        }

        public int NumberOfPhosphorousAtoms(int index)
        {
            return this.m_Components[index].Phosphoroustoms;
        }

        public int NumberOfSulfurAtoms(int index)
        {
            return this.m_Components[index].SulfurAtoms;
        }

        public double AccentricFactor(int index)
        {
            return this.m_Components[index].AccentricFactor;
        }

        public double CriticalTemperature(int index)
        {
            return this.m_Components[index].CriticalTemperature;
        }

        public double CriticalPressure(int index)
        {
            return this.m_Components[index].CriticalPressure;
        }

        public double boilingPoint(int index)
        {
            return this.m_Components[index].boilingPoint;
        }

        public string meltingPoint(int index)
        {
            return this.m_Components[index].MeltingPoint;
        }

        public double IdealGasHeatOfFormation(int index)
        {
            return this.m_Components[index].IdealGasHeatOfFormation;
        }

        public double IdealGasGibbsFreeEnergyOfFormation(int index)
        {
            return this.m_Components[index].IdealGasGibbsFreeEnergyOfFormation;
        }

        public int SourceUnitOperation
        {
            get
            {
                return m_SourceUnitOperation;
            }
        }

        public int TargetUnitOperation
        {
            get
            {
                return m_TargetUnitOperation;
            }
        }

        public double TemperatureC
        {
            get
            {
                if (m_TemperatureUnit == "C") return m_Temperature;
                if (m_TemperatureUnit == "K") return m_Temperature - 273.15;
                if (m_TemperatureUnit == "R") return (m_Temperature - 491.67) * 5 / 9;
                if (m_TemperatureUnit == "F") return (m_Temperature - 32) * 5 / 9; // Temperature in Farenheit.
                throw new System.ArgumentException("Unit not found in list");
            }
        }

        public double Temperature
        {
            get
            {
                return m_Temperature;
            }
        }

        public string TemperatureUnit
        {
            get
            {
                return m_TemperatureUnit;
            }
        }

        public double PressureKPa
        {
            get
            {
                if (m_PressureUnit == "atm") return m_Pressure * 1.01325e+02;
                if (m_PressureUnit == "psia") return m_Pressure * 6.89476;
                if (m_PressureUnit == "psig") return m_Pressure * 6.89476 + 1.01325e+02;
                if (m_PressureUnit == "torr") return m_Pressure * 1.33322e-01;
                if (m_PressureUnit == "mmHg") return m_Pressure * 1.33322e-01;
                if (m_PressureUnit == "Pa") return m_Pressure / 1000;
                if (m_PressureUnit == "kPa") return m_Pressure;
                if (m_PressureUnit == "MPa G") return m_Pressure * 1000 + 1.01325e+02;
                if (m_PressureUnit == "bar") return m_Pressure * 1e+02;
                if (m_PressureUnit == "bar G") return m_Pressure * 1e+02 + 1.01325e+02;
                if (m_PressureUnit == "mbar") return m_Pressure * 0.1;
                if (m_PressureUnit == "kg/cm2") return m_Pressure * 98.0665;
                if (m_PressureUnit == "kg/cm2 G") return m_Pressure * 98.0665 + 1.01325e+02;
                if (m_PressureUnit == "in-water") return m_Pressure * 2.49089e-01;
                if (m_PressureUnit == "mm-water") return m_Pressure * 9.80665e-03;
                throw new System.ArgumentException("Unit not found in list");
            }
        }

        public double Pressure
        {
            get
            {
                return m_Pressure;
            }
        }

        public string PressureUnit
        {
            get
            {
                return m_PressureUnit;
            }
        }

        public double MoleVaporFraction
        {
            get
            {
                return m_MoleVapFrac;
            }
        }

        public double EnthalpyMJHR
        {
            get
            {
                if (0 == String.Compare(m_EnthalpyUnit, "Btu/h", true)) return m_Enthalpy * 1055.05598654593 / 1000000;
                if (0 == String.Compare(m_EnthalpyUnit, "Btu/hr", true)) return m_Enthalpy * 1055.05598654593 / 1000000;
                if (0 == String.Compare(m_EnthalpyUnit, "kBtu/h", true)) return m_Enthalpy * 1055.05598654593 / 1000;
                if (0 == String.Compare(m_EnthalpyUnit, "kBtu/hr", true)) return m_Enthalpy * 1055.05598654593 / 1000;
                if (0 == String.Compare(m_EnthalpyUnit, "mmBtu/h", true)) return m_Enthalpy * 1055.05598654593;
                if (0 == String.Compare(m_EnthalpyUnit, "mmBtu/hr", true)) return m_Enthalpy * 1055.05598654593;
                if (0 == String.Compare(m_EnthalpyUnit, "J/h", true)) return m_Enthalpy * 1e-06;
                if (0 == String.Compare(m_EnthalpyUnit, "J/hr", true)) return m_Enthalpy * 1e-06;
                if (0 == String.Compare(m_EnthalpyUnit, "kJ/sec", true)) return m_Enthalpy * 3.6;
                if (0 == String.Compare(m_EnthalpyUnit, "kJ/h", true)) return m_Enthalpy * 1e-03;
                if (0 == String.Compare(m_EnthalpyUnit, "kJ/hr", true)) return m_Enthalpy * 1e-03;
                if (0 == String.Compare(m_EnthalpyUnit, "mJ/h", true)) return m_Enthalpy;
                if (0 == String.Compare(m_EnthalpyUnit, "mJ/hr", true)) return m_Enthalpy;
                if (0 == String.Compare(m_EnthalpyUnit, "cal/h", true)) return m_Enthalpy * 4.1868 / 1000000;
                if (0 == String.Compare(m_EnthalpyUnit, "cal/hr", true)) return m_Enthalpy * 4.1868 / 1000000;
                if (0 == String.Compare(m_EnthalpyUnit, "kcal/h", true)) return m_Enthalpy * 4.1868 / 1000;
                if (0 == String.Compare(m_EnthalpyUnit, "kcal/hr", true)) return m_Enthalpy * 4.1868 / 1000;
                if (0 == String.Compare(m_EnthalpyUnit, "mcal/h", true)) return m_Enthalpy * 4.1868;
                if (0 == String.Compare(m_EnthalpyUnit, "mcal/hr", true)) return m_Enthalpy * 4.1868;
                //if (0 == String.Compare(m_EnthalpyUnit, "hp", true)) return m_Enthalpy * 7.457e+02 * 3600;
                //if (0 == String.Compare(m_EnthalpyUnit, "W", true)) return m_Enthalpy * 3600;
                //if (0 == String.Compare(m_EnthalpyUnit, "kw", true)) return m_Enthalpy * 3.6;
                //if (0 == String.Compare(m_EnthalpyUnit, "mw", true)) return m_Enthalpy * 0.0036;
                throw new System.ArgumentException("Unit not found in list");
            }
        }

        public double Enthalpy
        {
            get
            {                
                return m_Enthalpy;
            }
        }

        public string EnthalpyUnit
        {
            get
            {
                return m_EnthalpyUnit;
            }
        }

        public double Entropy
        {
            get
            {
                return m_StreamPropertyValues[20];
            }
        }

        public double EntropyMJKHR
        {
            get
            {
                if (m_StreamPropertyUnits[20] == "MMBtu/C/h") return m_StreamPropertyValues[20] * 1054.5;
                if (m_StreamPropertyUnits[20] == "MJ/C/h") return m_StreamPropertyValues[20];
                if (m_StreamPropertyUnits[20] == "kJ/C/sec") return m_StreamPropertyValues[20]/1000*3600;
                throw new System.ArgumentException("Unit not found in list");
            }
        }

        public string EntropyUnit
        {
            get
            {
                return m_StreamPropertyUnits[20];
            }
        }

        public double TotalMassFlowRate
        {
            get
            {
                return m_TotalMassFlowRate;
            }
        }

        public string TotalMassFlowRateUnit
        {
            get
            {
                return m_TotalMassFlowRateUnit;
            }
        }

        public double TotalMassFlowRateKGH
        {
            get
            {
                double factor = 1.0;
                if (0 == String.Compare(m_TotalMassFlowRateUnit, "lb/h", true)) factor = 0.45359237;
                else if (0 == String.Compare(m_TotalMassFlowRateUnit, "lb/min", true)) factor = 0.45359237 * 60;
                else if (0 == String.Compare(m_TotalMassFlowRateUnit, "lb/day", true)) factor = 0.45359237 / 24;
                else if (0 == String.Compare(m_TotalMassFlowRateUnit, "lb/sec", true)) factor = 0.45359237 * 3600;
                else if (0 == String.Compare(m_TotalMassFlowRateUnit, "lb/hr", true)) factor = 0.45359237;
                else if (0 == String.Compare(m_TotalMassFlowRateUnit, "lb/s", true)) factor = 0.45359237 * 3600;
                //else if (0 == String.Compare(m_ComponentMassFlowRateUnit, "lb/batch", true)) factor = 0.45359237;
                else if (0 == String.Compare(m_TotalMassFlowRateUnit, "kg/h", true)) factor = 1;
                else if (0 == String.Compare(m_TotalMassFlowRateUnit, "kg/min", true)) factor = 1 * 60;
                else if (0 == String.Compare(m_TotalMassFlowRateUnit, "kg/day", true)) factor = 1 / 24;
                else if (0 == String.Compare(m_TotalMassFlowRateUnit, "kg/sec", true)) factor = 1 * 3600;
                else if (0 == String.Compare(m_TotalMassFlowRateUnit, "kg/hr", true)) factor = 1;
                else if (0 == String.Compare(m_TotalMassFlowRateUnit, "kg/s", true)) factor = 1 * 3600;
                //else if (0 == String.Compare(m_ComponentMassFlowRateUnit, "kg/batch", true)) factor = 1;
                else if (0 == String.Compare(m_TotalMassFlowRateUnit, "g/h", true)) factor = 0.001;
                else if (0 == String.Compare(m_TotalMassFlowRateUnit, "g/min", true)) factor = 0.001 * 60;
                else if (0 == String.Compare(m_TotalMassFlowRateUnit, "g/day", true)) factor = 0.001 / 24;
                else if (0 == String.Compare(m_TotalMassFlowRateUnit, "g/sec", true)) factor = 0.001 * 3600;
                else if (0 == String.Compare(m_TotalMassFlowRateUnit, "g/hr", true)) factor = 0.001;
                else if (0 == String.Compare(m_TotalMassFlowRateUnit, "g/s", true)) factor = 0.001 * 3600;
                else if (0 == String.Compare(m_TotalMassFlowRateUnit, "g/batch", true)) factor = 0.001;
                //else if (0 == String.Compare(m_ComponentMassFlowRateUnit, "g", true)) factor = 0.001;
                else throw new System.ArgumentException("Unit not found in list");
                return m_TotalMassFlowRate * factor;
            }
        }

        public double TotalMoleFlowRateMolHr
        {
            get
            {
                double factor = 1.0;
                if (0 == String.Compare(m_TotalMassFlowRateUnit, "lbmol/h", true)) factor = 453.59237;
                else if (0 == String.Compare(m_TotalMassFlowRateUnit, "lbmol/min", true)) factor = 453.59237 * 60;
                else if (0 == String.Compare(m_TotalMassFlowRateUnit, "lbmol/day", true)) factor = 453.59237 / 24;
                else if (0 == String.Compare(m_TotalMassFlowRateUnit, "lbmol/sec", true)) factor = 453.59237 * 3600;
                else if (0 == String.Compare(m_TotalMassFlowRateUnit, "lbmol/hr", true)) factor = 453.59237;
                else if (0 == String.Compare(m_TotalMassFlowRateUnit, "lbmol/s", true)) factor = 453.59237 * 3600;
                //else if (0 == String.Compare(m_ComponentMassFlowRateUnit, "lb/batch", true)) factor = 0.45359237;
                else if (0 == String.Compare(m_TotalMassFlowRateUnit, "kmole/h", true)) factor = 0.001;
                else if (0 == String.Compare(m_TotalMassFlowRateUnit, "kmole/min", true)) factor = 0.001 * 60;
                else if (0 == String.Compare(m_TotalMassFlowRateUnit, "kmole/day", true)) factor = 0.001 / 24;
                else if (0 == String.Compare(m_TotalMassFlowRateUnit, "kmole/sec", true)) factor = 0.001 * 3600;
                else if (0 == String.Compare(m_TotalMassFlowRateUnit, "kmole/hr", true)) factor = 0.001;
                else if (0 == String.Compare(m_TotalMassFlowRateUnit, "kmole/s", true)) factor = 0.001 * 3600;
                //else if (0 == String.Compare(m_ComponentMassFlowRateUnit, "kg/batch", true)) factor = 1;
                else if (0 == String.Compare(m_TotalMassFlowRateUnit, "mol/h", true)) factor = 1;
                else if (0 == String.Compare(m_TotalMassFlowRateUnit, "mol/min", true)) factor = 1 * 60;
                else if (0 == String.Compare(m_TotalMassFlowRateUnit, "mol/day", true)) factor = 1 / 24;
                else if (0 == String.Compare(m_TotalMassFlowRateUnit, "mol/sec", true)) factor = 1 * 3600;
                else if (0 == String.Compare(m_TotalMassFlowRateUnit, "mol/hr", true)) factor = 1;
                else if (0 == String.Compare(m_TotalMassFlowRateUnit, "mol/s", true)) factor = 1 * 3600;
                else if (0 == String.Compare(m_TotalMassFlowRateUnit, "mol/batch", true)) factor = 1;
                //else if (0 == String.Compare(m_ComponentMassFlowRateUnit, "g", true)) factor = 0.001;
                else throw new System.ArgumentException("Unit not found in list");
                return m_TotalMoleFlowRate * factor;
            }
        }

        public double TotalMoleFlowRate
        {
            get
            {
                return m_TotalMoleFlowRate;
            }
        }

        public string TotalMoleFlowRateUnit
        {
            get
            {
                return m_TotalMoleFlowRateUnit;
            }
        }

        public double LiquidVolumetricFlowRate
        {
            get
            {
                return m_LiquidVolumetricFlowRate;
            }
        }

        public double LiquidVolumetricFlowRateM3HR
        {
            get
            {
                if (0 == String.Compare(m_TotalLiquidVolumeFlowRateUnit, "ft3/hr", true)) return m_LiquidVolumetricFlowRate * 0.028316846592;
                if (0 == String.Compare(m_TotalLiquidVolumeFlowRateUnit, "ft3/day", true)) return m_LiquidVolumetricFlowRate * 0.028316846592 / 24;
                if (0 == String.Compare(m_TotalLiquidVolumeFlowRateUnit, "ft3/min", true)) return m_LiquidVolumetricFlowRate * 0.028316846592 * 60;
                if (0 == String.Compare(m_TotalLiquidVolumeFlowRateUnit, "gph", true)) return m_LiquidVolumetricFlowRate * 0.00440488377086;
                if (0 == String.Compare(m_TotalLiquidVolumeFlowRateUnit, "gpd", true)) return m_LiquidVolumetricFlowRate * 0.00440488377086 / 24;
                if (0 == String.Compare(m_TotalLiquidVolumeFlowRateUnit, "bbl/hr", true)) return m_LiquidVolumetricFlowRate * 0.158987294928;
                if (0 == String.Compare(m_TotalLiquidVolumeFlowRateUnit, "bbl/day", true)) return m_LiquidVolumetricFlowRate * 0.158987294928 / 24;
                if (0 == String.Compare(m_TotalLiquidVolumeFlowRateUnit, "bbl/min", true)) return m_LiquidVolumetricFlowRate * 0.158987294928 * 60;
                if (0 == String.Compare(m_TotalLiquidVolumeFlowRateUnit, "m3/h", true)) return m_LiquidVolumetricFlowRate;
                if (0 == String.Compare(m_TotalLiquidVolumeFlowRateUnit, "m3/min", true)) return m_LiquidVolumetricFlowRate * 60;
                if (0 == String.Compare(m_TotalLiquidVolumeFlowRateUnit, "m3/day", true)) return m_LiquidVolumetricFlowRate / 24;
                if (0 == String.Compare(m_TotalLiquidVolumeFlowRateUnit, "liter/day", true)) return m_LiquidVolumetricFlowRate / 1000 / 24;
                if (0 == String.Compare(m_TotalLiquidVolumeFlowRateUnit, "liter/hr", true)) return m_LiquidVolumetricFlowRate / 1000;
                if (0 == String.Compare(m_TotalLiquidVolumeFlowRateUnit, "liter/min", true)) return m_LiquidVolumetricFlowRate / 1000 * 60;
                if (0 == String.Compare(m_TotalLiquidVolumeFlowRateUnit, "cc/sec", true)) return m_LiquidVolumetricFlowRate / 1000 * 3600;
                if (0 == String.Compare(m_TotalLiquidVolumeFlowRateUnit, "Imp gph", true)) return m_LiquidVolumetricFlowRate * 0.00454609;
                if (0 == String.Compare(m_TotalLiquidVolumeFlowRateUnit, "Imp gpd", true)) return m_LiquidVolumetricFlowRate * 0.00454609 / 24;
                if (0 == String.Compare(m_TotalLiquidVolumeFlowRateUnit, "imp gpm", true)) return m_LiquidVolumetricFlowRate * 0.00454609 * 60;
                throw new System.ArgumentException("Unit not found in list");
            }
        }

        public string LiquidVolumetricFlowRateUnit
        {
            get
            {
                return m_TotalLiquidVolumeFlowRateUnit;
            }
        }

        public double VaporVolumetricFlowRateM3Hr
        {
            get
            {
                if (0 == String.Compare(m_TotalVaporVolumeFlowRateUnit, "ft3/hr", true)) return m_VaporVolumetricFlowRate * 0.028316846592;
                if (0 == String.Compare(m_TotalVaporVolumeFlowRateUnit, "MMft3/day", true)) return m_VaporVolumetricFlowRate * 0.028316846592 / 1000000 / 24;
                if (0 == String.Compare(m_TotalVaporVolumeFlowRateUnit, "ft3/min", true)) return m_VaporVolumetricFlowRate * 0.028316846592 * 60;
                if (0 == String.Compare(m_TotalVaporVolumeFlowRateUnit, "m3/h", true)) return m_VaporVolumetricFlowRate;
                if (0 == String.Compare(m_TotalVaporVolumeFlowRateUnit, "m3/min", true)) return m_VaporVolumetricFlowRate * 60;
                if (0 == String.Compare(m_TotalVaporVolumeFlowRateUnit, "m3/day", true)) return m_VaporVolumetricFlowRate / 24;
                if (0 == String.Compare(m_TotalVaporVolumeFlowRateUnit, "liter/hr", true)) return m_VaporVolumetricFlowRate / 1000;
                if (0 == String.Compare(m_TotalVaporVolumeFlowRateUnit, "liter/day", true)) return m_VaporVolumetricFlowRate / 1000 / 24;
                if (0 == String.Compare(m_TotalVaporVolumeFlowRateUnit, "liter/min", true)) return m_VaporVolumetricFlowRate / 1000 * 60;
                if (0 == String.Compare(m_TotalVaporVolumeFlowRateUnit, "cc/sec", true)) return m_VaporVolumetricFlowRate / 1000 * 3600;
                throw new System.ArgumentException("Unit not found in list");
            }
        }

        public double VaporVolumetricFlowRate
        {
            get
            {
                return m_VaporVolumetricFlowRate;
            }
        }

        public string VaporVolumetricFlowRateUnit
        {
            get
            {
                return m_TotalVaporVolumeFlowRateUnit;
            }
        }

        public double[] ComponentMassFlowRates
        {
            get
            {
                return m_CompMassFlowRate;
            }
        }

        public double[] ComponentMassFlowRatesKGH
        {
            get
            {
                int test1 = String.Compare(m_ComponentMassFlowRateUnit, "g/sec", true);
                int test2 = String.Compare(m_ComponentMassFlowRateUnit, "lb/min", true);
                double factor = 1.0;
                if (0 == String.Compare(m_ComponentMassFlowRateUnit, "lb/h", true)) factor = 0.45359237;
                else if (0 == String.Compare(m_ComponentMassFlowRateUnit, "lb/min", true)) factor = 0.45359237 * 60;
                else if (0 == String.Compare(m_ComponentMassFlowRateUnit, "lb/day", true)) factor = 0.45359237 /24;
                else if (0 == String.Compare(m_ComponentMassFlowRateUnit, "lb/sec", true)) factor = 0.45359237 *3600;
                else if (0 == String.Compare(m_ComponentMassFlowRateUnit, "lb/hr", true)) factor = 0.45359237;
                else if (0 == String.Compare(m_ComponentMassFlowRateUnit, "lb/s", true)) factor = 0.45359237 *3600;
                //else if (0 == String.Compare(m_ComponentMassFlowRateUnit, "lb/batch", true)) factor = 0.45359237;
                else if (0 == String.Compare(m_ComponentMassFlowRateUnit, "kg/h", true)) factor = 1;
                else if (0 == String.Compare(m_ComponentMassFlowRateUnit, "kg/min", true)) factor = 1 * 60;
                else if (0 == String.Compare(m_ComponentMassFlowRateUnit, "kg/day", true)) factor = 1 / 24;
                else if (0 == String.Compare(m_ComponentMassFlowRateUnit, "kg/sec", true)) factor = 1 * 3600;
                else if (0 == String.Compare(m_ComponentMassFlowRateUnit, "kg/hr", true)) factor = 1;
                else if (0 == String.Compare(m_ComponentMassFlowRateUnit, "kg/s", true)) factor = 1 * 3600;
                //else if (0 == String.Compare(m_ComponentMassFlowRateUnit, "kg/batch", true)) factor = 1;
                else if (0 == String.Compare(m_ComponentMassFlowRateUnit, "g/h", true)) factor = 0.001;
                else if (0 == String.Compare(m_ComponentMassFlowRateUnit, "g/min", true)) factor = 0.001 * 60;
                else if (0 == String.Compare(m_ComponentMassFlowRateUnit, "g/day", true)) factor = 0.001 / 24;
                else if (0 == String.Compare(m_ComponentMassFlowRateUnit, "g/sec", true)) factor = 0.001 *3600;
                else if (0 == String.Compare(m_ComponentMassFlowRateUnit, "g/hr", true)) factor = 0.001;
                else if (0 == String.Compare(m_ComponentMassFlowRateUnit, "g/s", true)) factor = 0.001 * 3600;
                else if (0 == String.Compare(m_ComponentMassFlowRateUnit, "g/batch", true)) factor = 0.001;
                //else if (0 == String.Compare(m_ComponentMassFlowRateUnit, "g", true)) factor = 0.001;
                else throw new System.ArgumentException("Unit not found in list");
                double[] retVal = new double[m_CompMassFlowRate.Length];
                for (int i = 0; i < m_CompMassFlowRate.Length; i++)
                {
                    retVal[i] = m_CompMassFlowRate[i] * factor;
                }
                return retVal;
            }
        }

        public string ComponentMassFlowRatesUnit
        {
            get
            {
                return m_ComponentMassFlowRateUnit;
            }
        }

        public string ComponentMoleFlowRatesUnit
        {
            get
            {
                return m_ComponentMoleFlowRateUnit;
            }
        }

        public double[] ComponentMoleFlowRates
        {
            get
            {
                return m_CompMoleFlow;
            }
        }

        public double[] ComponentMoleFractions
        {
            get
            {
                return m_CompMoleFraction;
            }
        }

        public double Cost
        {
            get
            {
                return m_Cost;
            }
        }
    }
}
