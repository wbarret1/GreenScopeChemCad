using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
// using System.Threading.Tasks;

namespace GreenScopeChemCad
{
    [System.Runtime.Serialization.DataContract]
    class UnitOperation
    {
        [System.Runtime.Serialization.DataMember]
        int m_UnitOpId;
        UnitOperationInfo p_UnitOpInfo;
        StreamInfo p_StreamInfo;
        UnitOpSpecUnitConversion p_UnitOpSpecUnitConversion;
        Flowsheet p_Flowsheet;
        int[] m_InletStreamIds;
        int[] m_OutletStreamIds;
        [System.Runtime.Serialization.DataMember]
        string m_Label;
        [System.Runtime.Serialization.DataMember]
        string m_Category;
        int m_UnitOpSpecArrayDimensions;
        int m_UnitOpSpecArrayDimensions2;
        int m_NumParams;
        double[] m_UnitOpSpec;
        double[] m_UnitOpSpec2;
        string[] m_ParamNames;
        string[] m_ParamUnits;
        int numParams;

        double[] inletTempR;
        double[] inletPressPSIA;
        double[] inletMvf;
        double[] inletEnthBTU_HR;
        double[][] inletCompFlowLbmol_hr;
        double[] outletTempR;
        double[] outletPressPSIA;
        double[] outletMvf;
        double[] outletEnthBTU_HR;
        double[][] outletCompFlowLbmol_hr;


        public UnitOperation(int unitOpID, object vbServer)
        {
            m_UnitOpId = unitOpID;
            p_UnitOpInfo = ((VBServerWrapper)vbServer).GetUnitOpInfo();
            p_StreamInfo = ((VBServerWrapper)vbServer).GetStreamInfo();
            p_Flowsheet = ((VBServerWrapper)vbServer).GetFlowsheet();
            p_UnitOpSpecUnitConversion = ((VBServerWrapper)vbServer).GetUnitOpSpecUnitConversion();
            m_InletStreamIds = p_Flowsheet.GetInletStreamIDsToUnitOp(unitOpID);
            m_OutletStreamIds = p_Flowsheet.GetOutletStreamIDsToUnitOp(unitOpID);
            m_Label = p_UnitOpInfo.GetUnitOpLabel(unitOpID);
            m_Category = p_UnitOpInfo.GetUnitOpCategory(unitOpID);
            m_UnitOpSpec = p_UnitOpInfo.GetUnitOpSpec(unitOpID);
            m_UnitOpSpecArrayDimensions = p_UnitOpInfo.UnitOpSpecArrayDiemsions();
            m_UnitOpSpecArrayDimensions2 = p_UnitOpSpecUnitConversion.GetUnitOpSpecArrayDimension();
            m_NumParams = p_UnitOpSpecUnitConversion.NumberOfParameters(m_UnitOpId);
            p_UnitOpSpecUnitConversion.FromCurUserUnitsToInternalUnits(m_UnitOpSpec, ref m_UnitOpSpec2);
            numParams = p_UnitOpSpecUnitConversion.NumberOfParameters(m_UnitOpId);
            m_ParamNames = new string[m_UnitOpSpecArrayDimensions];
            m_ParamUnits = new string[m_UnitOpSpecArrayDimensions];
            for (int i = 0; i < m_UnitOpSpecArrayDimensions; i++)
            {
                p_UnitOpSpecUnitConversion.GetCurUserUnitString(m_UnitOpId, i, ref m_ParamNames[i], ref m_ParamUnits[i]);
            }
            inletTempR = new double[m_InletStreamIds.Length];
            inletPressPSIA = new double[m_InletStreamIds.Length];
            inletMvf = new double[m_InletStreamIds.Length];
            inletEnthBTU_HR = new double[m_InletStreamIds.Length];
            inletCompFlowLbmol_hr = new double[m_InletStreamIds.Length][];
            for (int i = 0; i < m_InletStreamIds.Length; i++)
            {
                p_StreamInfo.GetStreamByID(m_InletStreamIds[i], ref inletTempR[i], ref inletPressPSIA[i], ref inletMvf[i], ref inletEnthBTU_HR[i], ref inletCompFlowLbmol_hr[i]);
            }
            outletTempR = new double[m_OutletStreamIds.Length];
            outletPressPSIA = new double[m_OutletStreamIds.Length];
            outletMvf = new double[m_OutletStreamIds.Length];
            outletEnthBTU_HR = new double[m_OutletStreamIds.Length];
            outletCompFlowLbmol_hr = new double[m_OutletStreamIds.Length][];
            for (int i = 0; i < m_OutletStreamIds.Length; i++)
            {
                p_StreamInfo.GetStreamByID(m_OutletStreamIds[i], ref outletTempR[i], ref outletPressPSIA[i], ref outletMvf[i], ref outletEnthBTU_HR[i], ref outletCompFlowLbmol_hr[i]);
            }
        }

        //~UnitOperation()
        //{
        //    //if (p_IFlowsheet != null)
        //    //    System.Runtime.InteropServices.Marshal.FinalReleaseComObject(p_IFlowsheet);
        //}

        public int UnitOpId
        {
            get
            {
                return m_UnitOpId;
            }
        }

        public string Label
        {
            get
            {
                return m_Label;
            }
        }

        public string Category
        {
            get
            {
                return m_Category;
            }
        }

        public double[] Specification
        {
            get
            {
                return m_UnitOpSpec;
            }
        }

        public double HeatAdded
        {
            get
            {
                if (m_Category == "PUMP") return 0.0;
                else if (m_Category == "COMP") return 0.0;
                else if (m_Category == "EXPN") return 0.0;
                else if (m_Category == "MIXE") return 0.0;
                else if (m_Category == "PIPE") return (1055.05598654593 / 1000000) * (m_UnitOpSpec[109]);// BTU
                else if (m_Category == "BATC") return (1055.05598654593 / 1000000) * (m_UnitOpSpec[60] - m_UnitOpSpec[61] + m_UnitOpSpec[80] - m_UnitOpSpec[81] + m_UnitOpSpec[100] - m_UnitOpSpec[101] + m_UnitOpSpec[120] - m_UnitOpSpec[121] + m_UnitOpSpec[140] - m_UnitOpSpec[141]);// BTU
                else if (m_Category == "SCDS") return (1055.05598654593 / 1000000) * (m_UnitOpSpec[38] + m_UnitOpSpec[39]);//MMBTU/hr
                else if (m_Category == "SHOR") return (1055.05598654593 / 1000000) * (m_UnitOpSpec[13] + m_UnitOpSpec[14]);//MMBTU/hr
                else if (m_Category == "TOWR") return (1055.05598654593 / 1000000) * (m_UnitOpSpec[40] + m_UnitOpSpec[41]);//MMBTU/hr
                else if (m_Category == "TPLS") return (1055.05598654593 / 1000000) * (m_UnitOpSpec[3] + m_UnitOpSpec[4] + m_UnitOpSpec[5] + m_UnitOpSpec[6] + m_UnitOpSpec[7] + m_UnitOpSpec[8] + m_UnitOpSpec[9] + m_UnitOpSpec[10] + m_UnitOpSpec[11] + m_UnitOpSpec[12] + m_UnitOpSpec[13] + m_UnitOpSpec[14]);//MMBTU/hr
                else if (m_Category == "FLAS") return (1055.05598654593 / 1000000) * (m_UnitOpSpec[5]);//MMBTU/hr
                else if (m_Category == "FIRE") return (1055.05598654593 / 1000000) * m_UnitOpSpec[7];//MMBTU/hr
                else if (m_Category == "HTXR") return (1055.05598654593 / 1000000) * m_UnitOpSpec[21];//MMBTU/hr
                else if (m_Category == "LNGH") return (1055.05598654593 / 1000000) * m_UnitOpSpec[23];//MMBTU/hr
                else if (m_Category == "EXTR") return 0.0;
                else if (m_Category == "VESL") return (1055.05598654593 / 1000000) * m_UnitOpSpec[5];
                else if (m_Category == "CSEP") return (1055.05598654593 / 1000000) * m_UnitOpSpec[41]; //MMBTU/hr
                else if (m_Category == "BREA") return (1055.05598654593 / 1000000) * m_UnitOpSpec[8];//MMBTU/hr
                else if (m_Category == "EREA") return (1055.05598654593 / 1000000) * m_UnitOpSpec[6];//MMBTU/hr
                else if (m_Category == "GIBS") return (1055.05598654593 / 1000000) * m_UnitOpSpec[5];//MMBTU/hr
                else if (m_Category == "KREA") return (1055.05598654593 / 1000000) * m_UnitOpSpec[8];//MMBTU/hr
                else if (m_Category == "POLY") return 0.0;
                else if (m_Category == "REAC") return (1055.05598654593 / 1000000) * m_UnitOpSpec[4];//MMBTU/hr
                else return 0.0;
            }
        }
        public double CondenserDuty
        {
            get
            {
                if (m_Category == "SCDS") return (1055.05598654593 / 1000000) * (m_UnitOpSpec[38]);//MMBTU/hr
                else if (m_Category == "SHOR") return (1055.05598654593 / 1000000) * (m_UnitOpSpec[13]);//MMBTU/hr
                else if (m_Category == "TOWR") return (1055.05598654593 / 1000000) * (m_UnitOpSpec[40]);//MMBTU/hr
                //else if (m_Category == "TPLS") return (1055.05598654593 / 1000000) * (m_UnitOpSpec[3] + m_UnitOpSpec[4] + m_UnitOpSpec[5] + m_UnitOpSpec[6] + m_UnitOpSpec[7] + m_UnitOpSpec[8] + m_UnitOpSpec[9] + m_UnitOpSpec[10] + m_UnitOpSpec[11] + m_UnitOpSpec[12] + m_UnitOpSpec[13] + m_UnitOpSpec[14]);//MMBTU/hr
                else return 0.0;
            }
        }
        public double ReboilerDuty
        {
            get
            {
                if (m_Category == "SCDS") return (1055.05598654593 / 1000000) * (m_UnitOpSpec[39]);//MMBTU/hr
                else if (m_Category == "SHOR") return (1055.05598654593 / 1000000) * (m_UnitOpSpec[14]);//MMBTU/hr
                else if (m_Category == "TOWR") return (1055.05598654593 / 1000000) * (m_UnitOpSpec[41]);//MMBTU/hr
                //else if (m_Category == "TPLS") return (1055.05598654593 / 1000000) * (m_UnitOpSpec[3] + m_UnitOpSpec[4] + m_UnitOpSpec[5] + m_UnitOpSpec[6] + m_UnitOpSpec[7] + m_UnitOpSpec[8] + m_UnitOpSpec[9] + m_UnitOpSpec[10] + m_UnitOpSpec[11] + m_UnitOpSpec[12] + m_UnitOpSpec[13] + m_UnitOpSpec[14]);//MMBTU/hr
                else return 0.0;
            }
        }

        public double HeatOfReaction // MJ/kg
        {
            get
            {
                if (m_Category == "BREA") return (1055.05598654593 / 1000000) * m_UnitOpSpec[51];//MMBTU/hr
                else if (m_Category == "EREA") return (1055.05598654593 / 1000000) * m_UnitOpSpec[33];//MMBTU/hr
                else if (m_Category == "GIBS") return (1055.05598654593 / 1000000) * m_UnitOpSpec[8];//MMBTU/hr
                else if (m_Category == "KREA") return (1055.05598654593 / 1000000) * m_UnitOpSpec[41];//MMBTU/hr
                else if (m_Category == "POLY") return 0.0;
                else if (m_Category == "REAC")
                {
                    int keyComponent = (int)m_UnitOpSpec[5];
                    double keyFlow = 0;
                    foreach (double[] compFlowRateLbmol_hr in inletCompFlowLbmol_hr)
                        keyFlow = keyFlow + compFlowRateLbmol_hr[keyComponent];
                    foreach (double[] compFlowRateLbmol_hr in outletCompFlowLbmol_hr)
                        keyFlow = keyFlow - compFlowRateLbmol_hr[keyComponent];
                    return (1055.05598654593 / 1000000) * keyFlow * m_UnitOpSpec[9];//BTU/lbmol
                }
                else return 0.0;
            }
        }

        public double Power
        {
            get
            {
                if (m_Category == "PUMP") return (0.00105505598654593) * m_UnitOpSpec[5];//BTU/hr
                else if (m_Category == "COMP") return (0.00105505598654593) * m_UnitOpSpec[6];//BTU/hr
                else if (m_Category == "EXPN") return (0.00105505598654593) * m_UnitOpSpec[6];//BTU/hr
                else if (m_Category == "MIXE") return 0.0;
                else if (m_Category == "BATC") return 0.0;
                else if (m_Category == "SCDS") return 0.0;
                else if (m_Category == "SHOR") return 0.0;//MMBTU/hr
                else if (m_Category == "TOWR") return 0.0;//MMBTU/hr
                else if (m_Category == "FLAS") return 0.0;//MMBTU/hr
                else if (m_Category == "TPLS") return 0.0;//(1055.05598654593 / 1000000) * (m_UnitOpSpec[16] + m_UnitOpSpec[17] + m_UnitOpSpec[18] + m_UnitOpSpec[19] + m_UnitOpSpec[20] + m_UnitOpSpec[21] + m_UnitOpSpec[22] + m_UnitOpSpec[23] + m_UnitOpSpec[24] + m_UnitOpSpec[25] + m_UnitOpSpec[26]);//MMBTU/hr
                else if (m_Category == "FIRE") return 0.0;//MMBTU/hr
                else if (m_Category == "HTXR") return 0.0;//MMBTU/hr
                else if (m_Category == "LNGH") return 0.0;//MMBTU/hr
                else if (m_Category == "EXTR") return 0.0;
                else if (m_Category == "VESL") return 0.0; //MMBTU/hr
                else if (m_Category == "CSEP") return 0.0; //MMBTU/hr
                else if (m_Category == "BREA") return 0.0;//MMBTU/hr
                else if (m_Category == "EREA") return 0.0;//MMBTU/hr
                else if (m_Category == "GIBS") return 0.0;//MMBTU/hr
                else if (m_Category == "KREA") return 0.0;//MMBTU/hr
                else if (m_Category == "POLY") return 0.0;
                else if (m_Category == "REAC") return 0.0;//MMBTU/hr
                else return 0.0;
            }
        }

        public double TotalPurchaseCost
        {
            get
            {
                if (m_Category == "PUMP") return m_UnitOpSpec[22];
                else if (m_Category == "COMP") return m_UnitOpSpec[22];
                else if (m_Category == "EXPN") return m_UnitOpSpec[22];
                else if (m_Category == "MIXE") return 0.0;
                else if (m_Category == "BATC") return 0.0;
                else if (m_Category == "SCDS") return m_UnitOpSpec[75];
                else if (m_Category == "SHOR") return 0.0;
                else if (m_Category == "TOWR") return m_UnitOpSpec[102];
                else if (m_Category == "TPLS") return 0.0;
                else if (m_Category == "FIRE") return m_UnitOpSpec[17];
                else if (m_Category == "HTXR") return m_UnitOpSpec[57];
                else if (m_Category == "LNGH") return 0.0;
                else if (m_Category == "EXTR") return 0.0;
                else if (m_Category == "CSEP") return 0.0;
                else if (m_Category == "BREA") return 0.0;
                else if (m_Category == "EREA") return 0.0;
                else if (m_Category == "GIBS") return 0.0;
                else if (m_Category == "KREA") return 0.0;
                else if (m_Category == "POLY") return 0.0;
                else if (m_Category == "REAC") return 0.0;
                else return 0.0;
            }
        }

        public double TotalInstalledCost
        {
            get
            {
                if (m_Category == "PUMP") return m_UnitOpSpec[23];
                else if (m_Category == "COMP") return m_UnitOpSpec[23];
                else if (m_Category == "EXPN") return m_UnitOpSpec[23];
                else if (m_Category == "MIXE") return 0.0;
                else if (m_Category == "BATC") return 0.0;
                else if (m_Category == "SCDS") return m_UnitOpSpec[76];
                else if (m_Category == "SHOR") return 0.0;
                else if (m_Category == "TOWR") return m_UnitOpSpec[103];
                else if (m_Category == "TPLS") return 0.0;
                else if (m_Category == "FIRE") return m_UnitOpSpec[18];
                else if (m_Category == "HTXR") return m_UnitOpSpec[56];
                else if (m_Category == "LNGH") return 0.0;
                else if (m_Category == "EXTR") return 0.0;
                else if (m_Category == "CSEP") return 0.0;
                else if (m_Category == "BREA") return 0.0;
                else if (m_Category == "EREA") return 0.0;
                else if (m_Category == "GIBS") return 0.0;
                else if (m_Category == "KREA") return 0.0;
                else if (m_Category == "POLY") return 0.0;
                else if (m_Category == "REAC") return 0.0;
                else return 0.0;
            }
        }

        public double ReactionStoicCoeff(int component)
        {
            if (m_Category == "REAC") return m_UnitOpSpec[50 + component];
            else return 0.0;
        }

    }
}
