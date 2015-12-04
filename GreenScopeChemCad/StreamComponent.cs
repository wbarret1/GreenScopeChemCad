using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
//// using System.Threading.Tasks;

namespace GreenScopeChemCad
{
    [System.Runtime.Serialization.DataContract]
    class StreamComponent
    {
        int m_StreamID;
        [System.Runtime.Serialization.DataMember]
        int m_ComponentPosition;
        [System.Runtime.Serialization.DataMember]
        int m_ComponentID;
        StreamInfo p_StreamInfo;
        Flowsheet flowsheet;
        CompPPData p_PropertyPackage;
        [System.Runtime.Serialization.DataMember]
        string m_ComponentName;
        [System.Runtime.Serialization.DataMember]
        string m_CASNumber;
        string m_MolecularFormula;
        [System.Runtime.Serialization.DataMember]
        double m_MolecularWeight;
        double m_CriticalTemperature;
        double m_CriticalPressure;
        double m_AcentricFactor;
        [System.Runtime.Serialization.DataMember]
        double m_BoilingPoint;
        double specificGravityAt60F;
        double m_IdealGasHeatOfFormation;
        double m_IdealGasGibbsFreeEnergyOfFormation;
        //double m_meltingPoint;
        //double m_FlashPoint;
        //double m_heatOfCombustion;
        //double m_idealGasHeatCapacity;
        //double m_liquidDensity;
        //double m_heatOfFormation;
        //double m_entropyOfFormation;


        public StreamComponent(int StreamID, int ComponentPosition, object vbServer)
        {
            m_StreamID = StreamID;
            m_ComponentPosition = ComponentPosition;
            p_StreamInfo = ((VBServerWrapper)vbServer).GetStreamInfo();
            m_ComponentName = p_StreamInfo.GetComponentNameByPos(m_ComponentPosition);
            m_ComponentID = p_StreamInfo.GetComponentIDByPos(m_ComponentPosition);
            m_CASNumber = NISTChemicalList.casNumber(m_ComponentName);
            m_MolecularFormula = NISTChemicalList.molecularFormula(m_ComponentName);
            flowsheet = ((VBServerWrapper)vbServer).GetFlowsheet();
            p_PropertyPackage = ((VBServerWrapper)vbServer).GetCompPPData();
            m_MolecularWeight = p_PropertyPackage.GetDataInInternalUnit(m_ComponentPosition, 1);
            m_CriticalTemperature = p_PropertyPackage.GetDataInInternalUnit(m_ComponentPosition, 2) * 5 / 9;
            m_CriticalPressure = p_PropertyPackage.GetDataInInternalUnit(m_ComponentPosition, 3) * 6.89476;
            m_AcentricFactor = p_PropertyPackage.GetDataInInternalUnit(m_ComponentPosition, 4);
            m_BoilingPoint = p_PropertyPackage.GetDataInInternalUnit(m_ComponentPosition, 5) * 5 / 9;
            specificGravityAt60F = p_PropertyPackage.GetDataInInternalUnit(m_ComponentPosition, 6);
            m_IdealGasHeatOfFormation = p_PropertyPackage.GetDataInInternalUnit(m_ComponentPosition, 7) * 2.32601 / 1000;
            m_IdealGasGibbsFreeEnergyOfFormation = p_PropertyPackage.GetDataInInternalUnit(m_ComponentPosition, 8) * 2.32601 / 1000;
            p_PropertyPackage.SaveUserCompData(m_ComponentPosition);
            double value = p_PropertyPackage.GetDataInInternalUnit(m_ComponentPosition, 6);
            double weight = 0;
            string formula = null;
            //m_meltingPoint = data[0];
            //m_FlashPoint = 0;// p_PropertyPackage.GetDataInInternalUnit(m_ComponentPosition  + 1, 8);
            //m_heatOfCombustion = p_PropertyPackage.GetData(m_ComponentPosition + 1, 48);
            //m_idealGasHeatCapacity = p_PropertyPackage.GetDataInInternalUnit(m_ComponentPosition + 1, 71);
            //m_liquidDensity = p_PropertyPackage.GetDataInInternalUnit(m_ComponentPosition + 1, 66);
            //m_entropyOfFormation = 0;// p_PropertyPackage.GetDataInInternalUnit(m_ComponentPosition, 8);
        }

        //~StreamComponent()
        //{
        //    //if (p_StreamInfo != null)
        //    //    System.Runtime.InteropServices.Marshal.ReleaseComObject(p_StreamInfo);
        //}

        public string ComponentName
        {
            get
            {
                return m_ComponentName;
            }
        }

        public string CASNumber
        {
            get
            {
                return m_CASNumber;
            }
        }

        private void GetCompoundData(string coumpoundName, ref double molecularWeight, ref String formula)
        {
            string url = "http://pubchem.ncbi.nlm.nih.gov/rest/pug/compound/name/";
            url = url + coumpoundName;
            url = url + "/cids/TXT";
            System.Net.HttpWebRequest request = (System.Net.HttpWebRequest)System.Net.WebRequest.Create(url);
            System.Net.WebResponse response = request.GetResponse();
            System.IO.Stream receiveStream = response.GetResponseStream();
            Encoding encode = System.Text.Encoding.GetEncoding("utf-8");
            // Pipes the stream to a higher level stream reader with the required encoding format. 
            System.IO.StreamReader readStream = new System.IO.StreamReader(receiveStream, encode);
            string cid = readStream.ReadLine();
            //readStream.
            response.Close();
            url = "http://pubchem.ncbi.nlm.nih.gov/rest/pug/compound/cid/";
            url = url + cid;
            url = url + "/property/MolecularWeight/TXT";
            request = (System.Net.HttpWebRequest)System.Net.WebRequest.Create(url);
            response = request.GetResponse();
            receiveStream = response.GetResponseStream();
            encode = System.Text.Encoding.GetEncoding("utf-8");
            // Pipes the stream to a higher level stream reader with the required encoding format. 
            readStream = new System.IO.StreamReader(receiveStream, encode);
            string result = readStream.ReadToEnd();
            //molecularWeight = Convert.ToDouble(readStream.ReadLine());
            response.Close();
        }

        public string MolecularFormula
        {
            get
            {
                return m_MolecularFormula;
            }
        }

        public double MolecularWeight
        {
            get
            {
                return m_MolecularWeight;
            }
        }

        public double CriticalTemperature
        {
            get
            {
                return m_CriticalTemperature;
            }
        }

        public double CriticalPressure
        {
            get
            {
                return m_CriticalPressure;
            }
        }

        public double boilingPoint
        {
            get
            {
                return m_BoilingPoint;
            }
        }

        public double AccentricFactor
        {
            get
            {
                return m_AcentricFactor;
            }
        }

        public double IdealGasHeatOfFormation
        {
            get
            {
                return m_IdealGasHeatOfFormation;
            }
        }

        public double IdealGasGibbsFreeEnergyOfFormation
        {
            get
            {
                return m_IdealGasGibbsFreeEnergyOfFormation;
            }
        }

        //public double meltingPoint
        //{
        //    get
        //    {
        //        return m_meltingPoint;
        //    }
        //}

        //public double FlashPoint
        //{
        //    get
        //    {
        //        return m_FlashPoint;
        //    }
        //}

        //public double heatOfCombustion
        //{
        //    get
        //    {
        //        return m_heatOfCombustion;
        //    }
        //}

        //public double idealGasHeatCapacity
        //{
        //    get
        //    {
        //        return m_idealGasHeatCapacity;
        //    }
        //}

        //public double heatOfVaporization
        //{
        //    get
        //    {
        //        return m_heatOfVaporization;
        //    }
        //}

        //public double liquidDensity
        //{
        //    get
        //    {
        //        return m_liquidDensity;
        //    }
        //}

        //public double heatOfFormation
        //{
        //    get
        //    {
        //        return 0.0;
        //    }
        //}

        //public double entropyOfFormation
        //{
        //    get
        //    {
        //        return 0.0;
        //    }
        //}
    }
}
