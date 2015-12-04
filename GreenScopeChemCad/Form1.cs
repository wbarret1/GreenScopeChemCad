using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
//// using System.Threading.Tasks;
using System.Windows.Forms;

namespace GreenScopeChemCad
{

    public partial class Form1 : Form
    {
        string chemCadFileName;
        string excelFileName;
        string processReferenceTemperatureUnit = "Celsius";
        string referenceTemperatureUnit = "Celsius";
        string referencePressureUnit = "kPa";
        double processReferenceTemperature = 60;
        double referenceTemperature = 25;
        double referencePressure = 101.325;
        DataTable feedStreamsTable;
        int[] feedStreamIds = new int[0];
        string[] feedStreamRenewables = new string[0];
        int[] productStreamIds = new int[0];
        string[] productStreamProductOrWastes = new string[0];
        string[] productStreamEcoProducts = new string[0];
        string[] productStreamPollutedNonPolluteds = new string[0];
        string[] productStreamRenewables = new string[0];
        DataTable productStreamsTable;
        int mainGlobalReaction = 0;
        int mainGlobalProduct = 0;
        int mainGlobalProductStream = 0;
        double[] stoichiometry = new double[0];
        //ChemicalSpecies m_NISTChemicals;

        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            chemCadFileName = Environment.GetFolderPath(System.Environment.SpecialFolder.Personal);
            openFileDialog1.ShowDialog();
            chemCadFileName = String.Copy(openFileDialog1.FileName);
            excelFileName = System.IO.Path.ChangeExtension(chemCadFileName, "xlsm");
            this.label2.Text = String.Concat("Excel File Name:  ", excelFileName);
            textBox1.Text = chemCadFileName;
            if (System.IO.File.Exists(excelFileName))
            {
                DocumentFormat.OpenXml.Packaging.SpreadsheetDocument spreadsheet = DocumentFormat.OpenXml.Packaging.SpreadsheetDocument.Open(excelFileName, true);
                this.GetReferenceConditionsFromSpreadsheet(spreadsheet);
                this.GetFeedStreamRenewableFromSpreadsheet(spreadsheet, ref feedStreamIds, ref feedStreamRenewables);
                this.GetProductStreamInformationFromSpreadsheet(spreadsheet, ref productStreamIds, ref productStreamProductOrWastes, ref productStreamEcoProducts, ref productStreamPollutedNonPolluteds, ref productStreamRenewables);
                this.GetReactionInformationFromSpreadsheet(spreadsheet, ref mainGlobalReaction, ref mainGlobalProduct, ref mainGlobalProductStream, ref stoichiometry);
                spreadsheet.Close();
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            //excelFileName = Environment.GetFolderPath(System.Environment.SpecialFolder.Personal);
            saveFileDialog1.FileName = excelFileName;
            saveFileDialog1.ShowDialog();
            excelFileName = String.Copy(saveFileDialog1.FileName);
            textBox2.Text = excelFileName;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            this.progressBar1.Value = 0;
            Cursor oldCursor = this.Cursor;
            this.Cursor = Cursors.WaitCursor;
            if (string.IsNullOrEmpty(chemCadFileName))
            {
                System.Windows.Forms.MessageBox.Show("Please enter a CHEMCAD file to be exported.");
                this.Cursor = oldCursor;
                return;
            }

            DocumentFormat.OpenXml.Packaging.SpreadsheetDocument spreadsheet;
            DocumentFormat.OpenXml.Packaging.SpreadsheetDocument greenScopeTemplate;


            if (System.IO.File.Exists(excelFileName))
            {
                string message = "Excel File Exists. Do you want to replace it?.";
                string caption = "Overwrite Excel File.";
                MessageBoxButtons buttons = MessageBoxButtons.YesNo;
                System.Windows.Forms.DialogResult result = System.Windows.Forms.MessageBox.Show(message, caption, buttons);
                if (result != System.Windows.Forms.DialogResult.Yes)
                {
                    this.Cursor = oldCursor;
                    return;
                }
                spreadsheet = DocumentFormat.OpenXml.Packaging.SpreadsheetDocument.Open(excelFileName, true);
            }
            else
            {
                spreadsheet = DocumentFormat.OpenXml.Packaging.SpreadsheetDocument.Create(excelFileName, DocumentFormat.OpenXml.SpreadsheetDocumentType.MacroEnabledWorkbook);
                greenScopeTemplate = DocumentFormat.OpenXml.Packaging.SpreadsheetDocument.Open(new System.IO.MemoryStream(Properties.Resources.GRNS_LCI_v1_2_05_15), true);

                //Make sure it's clear
                spreadsheet.DeleteParts<DocumentFormat.OpenXml.Packaging.OpenXmlPart>(spreadsheet.GetPartsOfType<DocumentFormat.OpenXml.Packaging.OpenXmlPart>());

                //Copy all parts into the new book
                foreach (DocumentFormat.OpenXml.Packaging.OpenXmlPart part in greenScopeTemplate.GetPartsOfType<DocumentFormat.OpenXml.Packaging.OpenXmlPart>())
                {
                    DocumentFormat.OpenXml.Packaging.OpenXmlPart newPart = spreadsheet.AddPart<DocumentFormat.OpenXml.Packaging.OpenXmlPart>(part);
                }

                //Close template
                greenScopeTemplate.Close();
            }

            this.progressBar1.Value = 10;

            //this.ClearComponentSpreadsheet(spreadsheet);
            //this.ClearUnitOpSpreadsheet(spreadsheet);

            VBServerWrapper server = new VBServerWrapper();
            if (!server.LoadJob(chemCadFileName))
            {
                System.Windows.Forms.MessageBox.Show("The desired simulation did not load properly.");
                server.CloseSimulation();
                this.Cursor = oldCursor;
                return;
            }

            this.UpdateSpreadsheetChangeLog(spreadsheet, chemCadFileName, excelFileName);
            this.UpdateReferenceConditionsInSpreadsheet(spreadsheet);
            this.progressBar1.Value = 20;
            Flowsheet flowsheet = server.GetFlowsheet();
            int numStreams = flowsheet.NumberofStreams;
            int[] streamIDS = (int[])flowsheet.AllStreamIDs;
            int numUOs = flowsheet.NumberOfUnitOps;
            int[] unitOpIds = (int[])flowsheet.AllUnitOpIDs;
            int numProductStreams = flowsheet.NumberOfProductStreams;
            int[] productStreamIDS = (int[])flowsheet.ProductStreamIDs;
            int numFeedStreams = flowsheet.NumberOfFeedStreams;
            int[] feedStreamIDs = (int[])flowsheet.FeedStreamIDs;
            int numCutStreams = flowsheet.NumberOfCutStreams;
            int[] cutStreamIDs = (int[])flowsheet.CutStreamsIDs;
            string[] streamNames = new string[numStreams];
            Stream[] allStreams = new Stream[numStreams];
            for (int i = 0; i < numStreams; i++)
            {
                allStreams[i] = new Stream(streamIDS[i], server);
                streamNames[i] = allStreams[i].StreamName;
            }
            Stream[] feedStreams = new Stream[numFeedStreams];
            for (int i = 0; i < numFeedStreams; i++)
            {
                for (int j = 0; j < numStreams; j++)
                {
                    if (feedStreamIDs[i] == allStreams[j].StreamID)
                    {
                        feedStreams[i] = allStreams[j];
                    }
                }
            }
            Stream[] productStreams = new Stream[numProductStreams];
            for (int i = 0; i < numProductStreams; i++)
            {
                for (int j = 0; j < numStreams; j++)
                {
                    if (productStreamIDS[i] == allStreams[j].StreamID)
                    {
                        productStreams[i] = allStreams[j];
                    }
                }
            }

            this.progressBar1.Value = 40;
            DataTable componentTable = this.CreateComponentDataTable(allStreams, spreadsheet);
            DataTable allStreamsTable = this.CreateAllStreamsDataTable("AllStreams", allStreams, componentTable);
            feedStreamsTable = this.CreateFeedStreamsDataTable("FeedStreams", feedStreams, null);
            productStreamsTable = this.CreateProductStreamsDataTable("ProductStreams", productStreams, null);

            AddInputStreamsToSpreadsheet(spreadsheet, feedStreams);
            AddOutputStreamsToSpreadsheet(spreadsheet, productStreams);

            //DataTable inletStreams;
            UnitOperation[] unitOps = new UnitOperation[flowsheet.NumberOfUnitOps];
            List<UnitOperation> pumps = new List<UnitOperation>();
            List<UnitOperation> mixers = new List<UnitOperation>();
            List<UnitOperation> distillationColumns = new List<UnitOperation>();
            List<UnitOperation> heatExchangers = new List<UnitOperation>();
            List<UnitOperation> extractors = new List<UnitOperation>();
            List<UnitOperation> componentSeparators = new List<UnitOperation>();
            List<UnitOperation> reactors = new List<UnitOperation>();
            List<UnitOperation> other = new List<UnitOperation>();
            for (int i = 0; i < flowsheet.NumberOfUnitOps; i++)
            {
                unitOps[i] = new UnitOperation(unitOpIds[i], server);
                if (unitOps[i].Category == "PUMP") pumps.Add(unitOps[i]);
                else if (unitOps[i].Category == "COMP") pumps.Add(unitOps[i]);
                else if (unitOps[i].Category == "EXPN") pumps.Add(unitOps[i]);
                else if (unitOps[i].Category == "MIXE") mixers.Add(unitOps[i]);
                else if (unitOps[i].Category == "BATC") distillationColumns.Add(unitOps[i]);
                else if (unitOps[i].Category == "SCDS") distillationColumns.Add(unitOps[i]);
                else if (unitOps[i].Category == "SHOR") distillationColumns.Add(unitOps[i]);
                else if (unitOps[i].Category == "TOWR") distillationColumns.Add(unitOps[i]);
                else if (unitOps[i].Category == "TPLS") distillationColumns.Add(unitOps[i]);
                else if (unitOps[i].Category == "FIRE") heatExchangers.Add(unitOps[i]);
                else if (unitOps[i].Category == "HTXR") heatExchangers.Add(unitOps[i]);
                else if (unitOps[i].Category == "LNGH") heatExchangers.Add(unitOps[i]);
                else if (unitOps[i].Category == "EXTR") extractors.Add(unitOps[i]);
                else if (unitOps[i].Category == "CSEP") componentSeparators.Add(unitOps[i]);
                else if (unitOps[i].Category == "BREA") reactors.Add(unitOps[i]);
                else if (unitOps[i].Category == "EREA") reactors.Add(unitOps[i]);
                else if (unitOps[i].Category == "GIBS") reactors.Add(unitOps[i]);
                else if (unitOps[i].Category == "KREA") reactors.Add(unitOps[i]);
                else if (unitOps[i].Category == "POLY") reactors.Add(unitOps[i]);
                else if (unitOps[i].Category == "REAC") reactors.Add(unitOps[i]);
                else other.Add(unitOps[i]);
            }

            this.AddReactionsToSpreadsheet(spreadsheet, reactors, allStreams[0]);

            this.progressBar1.Value = 70;
            server.CloseSimulation();
            server.Dispose();

            this.progressBar1.Value = 80;
            DataTable unitOpTable = this.CreateUnitOperationDataTable(unitOps);
            DataTable reactionsTable = this.CreateReactionsTable(allStreams, reactors.ToArray<UnitOperation>());
            this.AddMPumpUnitOpsToSpreadsheet(spreadsheet, pumps.ToArray<UnitOperation>());
            this.AddMixerUnitOpsToSpreadsheet(spreadsheet, mixers.ToArray<UnitOperation>());
            this.AddDistillationUnitOpsToSpreadsheet(spreadsheet, distillationColumns.ToArray<UnitOperation>());
            this.AddHeatExchangerUnitOpsToSpreadsheet(spreadsheet, heatExchangers.ToArray<UnitOperation>());
            this.AddExtractorUnitOpsToSpreadsheet(spreadsheet, extractors.ToArray<UnitOperation>());
            this.AddComponentSeparatorUnitOpsToSpreadsheet(spreadsheet, componentSeparators.ToArray<UnitOperation>());
            this.AddReactorUnitOpsToSpreadsheet(spreadsheet, reactors.ToArray<UnitOperation>());
            this.AddOtherUnitOpsToSpreadsheet(spreadsheet, other.ToArray<UnitOperation>());

            float time = 0;
            string unit = String.Empty;
            flowsheet.GetDynamicTime(ref time, ref unit);

            this.dataGridView1.DataSource = componentTable;
            this.dataGridView2.DataSource = allStreamsTable;
            this.dataGridView3.DataSource = feedStreamsTable;
            this.dataGridView4.DataSource = productStreamsTable;
            this.dataGridView5.DataSource = unitOpTable;
            this.dataGridView6.DataSource = reactionsTable;

            System.IO.FileStream writer = new System.IO.FileStream("jsonTest.txt", System.IO.FileMode.Create);
            System.Runtime.Serialization.Json.DataContractJsonSerializer unitOpSerializer =
                new System.Runtime.Serialization.Json.DataContractJsonSerializer(typeof(UnitOperation[]));
            //foreach (UnitOperation unitOp in unitOps)
            //{
            unitOpSerializer.WriteObject(writer, unitOps);
            //}
            System.Runtime.Serialization.Json.DataContractJsonSerializer streamSerializer =
               new System.Runtime.Serialization.Json.DataContractJsonSerializer(typeof(Stream[]));
            //foreach (Stream stream in allStreams)
            //{
            streamSerializer.WriteObject(writer, allStreams);
            //}
            writer.Close();


            //Perform 'save as'
            spreadsheet.WorkbookPart.Workbook.CalculationProperties.ForceFullCalculation = true;
            spreadsheet.WorkbookPart.Workbook.Save();
            spreadsheet.Close();
            this.progressBar1.Value = 100;
            this.Cursor = oldCursor;
        }

        private void button4_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        DataTable CreateComponentDataTable(Stream[] streams, DocumentFormat.OpenXml.Packaging.SpreadsheetDocument spreadsheet)
        {
            DataTable componentTable = new DataTable("Components");
            DataRow row; // Create new DataColumn, set DataType,  
            componentTable.Columns.Add(new DataColumn("ComponentId", typeof(System.Int32)));
            componentTable.Columns.Add(new DataColumn("ComponentName", typeof(System.String)));
            componentTable.Columns.Add(new DataColumn("CAS Number", typeof(System.String)));
            componentTable.Columns.Add(new DataColumn("Formula", typeof(System.String)));
            componentTable.Columns.Add(new DataColumn("MolecularWeight", typeof(System.Double)));
            componentTable.Columns.Add(new DataColumn("criticalTemperature", typeof(System.Double)));
            componentTable.Columns.Add(new DataColumn("criticalPressure", typeof(System.Double)));
            componentTable.Columns.Add(new DataColumn("accentricFactor", typeof(System.Double)));
            componentTable.Columns.Add(new DataColumn("boilingPoint", typeof(System.Double)));
            componentTable.Columns.Add(new DataColumn("idealGasHeatofFormation", typeof(System.Double)));
            componentTable.Columns.Add(new DataColumn("idealGasGibbsFreeEnergyOfFormation", typeof(System.Double)));

            int numComponents = streams[0].NumberOfComponents;
            int[] compIds = new int[numComponents];
            string[] compNames = new string[numComponents];
            for (int i = 0; i < numComponents; i++)
            {
                row = componentTable.NewRow();
                row["ComponentId"] = streams[0].ComponentIDs[i];
                row["ComponentName"] = streams[0].ComponentNames[i];
                row["CAS Number"] = streams[0].casNumber(i);
                row["Formula"] = streams[0].MolecularFormula(i);
                row["MolecularWeight"] = streams[0].MolecularWeight(i);
                row["criticalTemperature"] = streams[0].CriticalTemperature(i);
                row["criticalPressure"] = streams[0].CriticalPressure(i);
                row["accentricFactor"] = streams[0].AccentricFactor(i);
                row["boilingPoint"] = streams[0].boilingPoint(i);
                row["idealGasHeatofFormation"] = streams[0].IdealGasHeatOfFormation(i);
                row["idealGasGibbsFreeEnergyOfFormation"] = streams[0].IdealGasGibbsFreeEnergyOfFormation(i);
                componentTable.Rows.Add(row);
            }
            AddComponentsToSpreadsheet(spreadsheet, streams[0]);


            foreach (Stream stream in streams)
            {
                componentTable.Columns.Add(new DataColumn(stream.StreamID.ToString() + " TotalFlowRate", typeof(System.Double)));
                componentTable.Columns.Add(new DataColumn(stream.StreamID.ToString() + " LiquidFlowRate", typeof(System.Double)));
                componentTable.Columns.Add(new DataColumn(stream.StreamID.ToString() + " VaporFlowRate", typeof(System.Double)));
                string[] streamCompNames = stream.ComponentNames;
                double[] streamFlows = stream.ComponentMassFlowRates;
                double[] streamLiquidFlows = stream.ComponentMoleFlowRates;
                double[] streamVaporFlows = stream.ComponentMoleFractions;
                for (int i = 0; i < stream.NumberOfComponents; i++)
                {
                    int compId = stream.ComponentIDs[i];
                    foreach (DataRow currrentRow in componentTable.Rows)
                    {
                        if ((int)currrentRow["ComponentId"] == compId)
                        {
                            //currrentRow[stream.StreamID.ToString() + " TotalFlowRate"] = streamFlows[i];
                            //currrentRow[stream.StreamID.ToString() + " LiquidFlowRate"] = streamLiquidFlows[i];
                            //currrentRow[stream.StreamID.ToString() + " VaporFlowRate"] = streamVaporFlows[i];
                        }
                    }
                }
            }
            return componentTable;
        }

        DataTable CreateReactionsTable(Stream[] streams, UnitOperation[] unitOps)
        {
            DataTable componentTable = new DataTable("Reactions");
            DataRow row; // Create new DataColumn, set DataType,  
            componentTable.Columns.Add(new DataColumn("ComponentId", typeof(System.Int32)));
            componentTable.Columns.Add(new DataColumn("ComponentName", typeof(System.String)));

            int numComponents = streams[0].NumberOfComponents;
            int[] compIds = new int[numComponents];
            string[] compNames = new string[numComponents];
            for (int i = 0; i < numComponents; i++)
            {
                row = componentTable.NewRow();
                row["ComponentId"] = streams[0].ComponentIDs[i];
                row["ComponentName"] = streams[0].ComponentNames[i];
                componentTable.Rows.Add(row);
            }

            int totalNumReactions = 0;
            foreach (UnitOperation unitOp in unitOps)
            {
                if (unitOp.Category == "BREA")
                {
                    int numReactions = (int)unitOp.Specification[20];
                    for (int i = 0; i < numReactions; i++ )
                        componentTable.Columns.Add(new DataColumn(unitOp.Label + "Reaction " + (i+1).ToString(), typeof(System.Int32)));
                    totalNumReactions = totalNumReactions + numReactions;
                }
                else if (unitOp.Category == "EREA")
                {
                    int numReactions = (int)unitOp.Specification[9];
                    totalNumReactions = totalNumReactions + numReactions;

                }
                else if (unitOp.Category == "GIBS")
                {
                    int numReactions = 1;// (int)unitOp.Specification[20];
                    totalNumReactions = totalNumReactions + numReactions;
                }
                else if (unitOp.Category == "KREA")
                {
                    int numReactions = (int)unitOp.Specification[20];
                    totalNumReactions = totalNumReactions + numReactions;

                }
                else if (unitOp.Category == "POLY")
                {
                    int numReactions = 0;// (int)unitOp.Specification[20];
                    totalNumReactions = totalNumReactions + numReactions;

                }
                else if (unitOp.Category == "REAC")
                {
                    componentTable.Columns.Add(new DataColumn("Reactor Id " + unitOp.UnitOpId.ToString(), typeof(System.Double)));
                    for (int i = 0; i < numComponents; i++)
                    {
                        componentTable.Rows[i]["Reactor Id " + unitOp.UnitOpId.ToString()] = unitOp.ReactionStoicCoeff(i);
                    }
                }
            }
            return componentTable;
        }

        DataTable CreateAllStreamsDataTable(string tableName, Stream[] streams, DataTable componentTable)
        {
            DataTable streamsTable = new DataTable(tableName);
            DataRow row;
            streamsTable.Columns.Add(new DataColumn("StreamId", typeof(System.Int32)));
            streamsTable.Columns.Add(new DataColumn("StreamName", typeof(System.String)));
            streamsTable.Columns.Add(new DataColumn("SourceUnit", typeof(System.Int32)));
            streamsTable.Columns.Add(new DataColumn("TargetUnit", typeof(System.Int32)));
            streamsTable.Columns.Add(new DataColumn("Temperature", typeof(System.Double)));
            streamsTable.Columns.Add(new DataColumn("TemperatureUnit", typeof(System.String)));
            streamsTable.Columns.Add(new DataColumn("Pressure", typeof(System.Double)));
            streamsTable.Columns.Add(new DataColumn("PressureUnit", typeof(System.String)));
            streamsTable.Columns.Add(new DataColumn("MoleVaporFraction", typeof(System.Double)));
            streamsTable.Columns.Add(new DataColumn("Enthaply", typeof(System.Double)));
            streamsTable.Columns.Add(new DataColumn("EnthaplyUnit", typeof(System.String)));
            streamsTable.Columns.Add(new DataColumn("Entropy", typeof(System.Double)));
            streamsTable.Columns.Add(new DataColumn("EntropyUnit", typeof(System.String)));
            streamsTable.Columns.Add(new DataColumn("TotalMassFlowRate", typeof(System.Double)));
            streamsTable.Columns.Add(new DataColumn("TotalMassFlowRateUnit", typeof(System.String)));
            streamsTable.Columns.Add(new DataColumn("TotalMoleFlowRate", typeof(System.Double)));
            streamsTable.Columns.Add(new DataColumn("TotalMoleFlowRateUnit", typeof(System.String)));
            streamsTable.Columns.Add(new DataColumn("LiquidVolumetricFlowRate", typeof(System.Double)));
            streamsTable.Columns.Add(new DataColumn("LiquidVolumetricFlowRateUnit", typeof(System.String)));
            streamsTable.Columns.Add(new DataColumn("VaporVolumetricFlowRate", typeof(System.Double)));
            streamsTable.Columns.Add(new DataColumn("VaporVolumetricFlowRateUnit", typeof(System.String)));
            streamsTable.Columns.Add(new DataColumn("Cost", typeof(System.Double)));
            foreach (Stream stream in streams)
            {
                row = streamsTable.NewRow();
                row["StreamId"] = stream.StreamID;
                row["StreamName"] = stream.StreamName;
                row["SourceUnit"] = stream.SourceUnitOperation;
                row["TargetUnit"] = stream.TargetUnitOperation;
                row["Temperature"] = stream.Temperature;
                row["TemperatureUnit"] = stream.TemperatureUnit;
                row["Pressure"] = stream.Pressure;
                row["PressureUnit"] = stream.PressureUnit;
                row["MoleVaporFraction"] = stream.MoleVaporFraction;
                row["Enthaply"] = stream.Enthalpy;
                row["EnthaplyUnit"] = stream.EnthalpyUnit;
                row["Entropy"] = stream.Entropy;
                row["EntropyUnit"] = stream.EntropyUnit;
                row["TotalMassFlowRate"] = stream.TotalMassFlowRate;
                row["TotalMassFlowRateUnit"] = stream.TotalMassFlowRateUnit;
                row["TotalMoleFlowRate"] = stream.TotalMoleFlowRate;
                row["TotalMoleFlowRateUnit"] = stream.TotalMoleFlowRateUnit;
                row["LiquidVolumetricFlowRate"] = stream.LiquidVolumetricFlowRate;
                row["LiquidVolumetricFlowRateUnit"] = stream.LiquidVolumetricFlowRateUnit;
                row["VaporVolumetricFlowRate"] = stream.VaporVolumetricFlowRate;
                row["VaporVolumetricFlowRateUnit"] = stream.VaporVolumetricFlowRateUnit;
                row["Cost"] = stream.Cost;
                streamsTable.Rows.Add(row);
            }
            return streamsTable;
        }

        DataTable CreateFeedStreamsDataTable(string tableName, Stream[] streams, DataTable componentTable)
        {
            dataGridView3.AutoGenerateColumns = false;
            DataTable streamsTable = new DataTable(tableName);
            streamsTable.Columns.Add(new DataColumn("StreamId", typeof(System.String)));
            DataGridViewTextBoxColumn streamId = new DataGridViewTextBoxColumn();
            streamId.HeaderText = "StreamID";
            streamId.DataPropertyName = "StreamId";
            streamsTable.Columns.Add(new DataColumn("StreamName", typeof(System.String)));
            DataGridViewTextBoxColumn streamName = new DataGridViewTextBoxColumn();
            streamName.HeaderText = "StreamName";
            streamName.DataPropertyName = "StreamName";
            streamsTable.Columns.Add(new DataColumn("Renewable", typeof(System.String)));
            DataGridViewComboBoxColumn renewable = new DataGridViewComboBoxColumn();
            List<string> renewableOptions = new List<string>() { "", "yes", "no" };
            renewable.DataSource = renewableOptions;
            renewable.HeaderText = "Renewable";
            renewable.DataPropertyName = "Renewable";
            streamsTable.Columns.Add(new DataColumn("SourceUnit", typeof(System.Int32)));
            DataGridViewTextBoxColumn sourceUnit = new DataGridViewTextBoxColumn();
            sourceUnit.HeaderText = "SourceUnit";
            sourceUnit.DataPropertyName = "SourceUnit";
            streamsTable.Columns.Add(new DataColumn("TargetUnit", typeof(System.Int32)));
            DataGridViewTextBoxColumn targetUnit = new DataGridViewTextBoxColumn();
            targetUnit.HeaderText = "TargetUnit";
            targetUnit.DataPropertyName = "TargetUnit";
            streamsTable.Columns.Add(new DataColumn("Temperature", typeof(System.Double)));
            DataGridViewTextBoxColumn temperature = new DataGridViewTextBoxColumn();
            temperature.HeaderText = "Temperature";
            temperature.DataPropertyName = "Temperature";
            streamsTable.Columns.Add(new DataColumn("TemperatureUnit", typeof(System.String)));
            DataGridViewTextBoxColumn temperatureUnit = new DataGridViewTextBoxColumn();
            temperatureUnit.HeaderText = "TemperatureUnit";
            temperatureUnit.DataPropertyName = "TemperatureUnit";
            streamsTable.Columns.Add(new DataColumn("Pressure", typeof(System.Double)));
            DataGridViewTextBoxColumn pressure = new DataGridViewTextBoxColumn();
            pressure.HeaderText = "Pressure";
            pressure.DataPropertyName = "Pressure";
            streamsTable.Columns.Add(new DataColumn("PressureUnit", typeof(System.String)));
            DataGridViewTextBoxColumn pressureUnit = new DataGridViewTextBoxColumn();
            pressureUnit.HeaderText = "PressureUnit";
            pressureUnit.DataPropertyName = "PressureUnit";
            streamsTable.Columns.Add(new DataColumn("MoleVaporFraction", typeof(System.Double)));
            DataGridViewTextBoxColumn moleVaporFraction = new DataGridViewTextBoxColumn();
            moleVaporFraction.HeaderText = "MoleVaporFraction";
            moleVaporFraction.DataPropertyName = "MoleVaporFraction";
            streamsTable.Columns.Add(new DataColumn("Enthaply", typeof(System.Double)));
            DataGridViewTextBoxColumn enthaply = new DataGridViewTextBoxColumn();
            enthaply.HeaderText = "Enthaply";
            enthaply.DataPropertyName = "Enthaply";
            streamsTable.Columns.Add(new DataColumn("EnthaplyUnit", typeof(System.String)));
            DataGridViewTextBoxColumn enthaplyUnit = new DataGridViewTextBoxColumn();
            enthaplyUnit.HeaderText = "EnthaplyUnit";
            enthaplyUnit.DataPropertyName = "EnthaplyUnit";
            streamsTable.Columns.Add(new DataColumn("Entropy", typeof(System.Double)));
            DataGridViewTextBoxColumn entropy = new DataGridViewTextBoxColumn();
            entropy.HeaderText = "Entropy";
            entropy.DataPropertyName = "Entropy";
            streamsTable.Columns.Add(new DataColumn("EntropyUnit", typeof(System.String)));
            DataGridViewTextBoxColumn entropyUnit = new DataGridViewTextBoxColumn();
            entropyUnit.HeaderText = "EntropyUnit";
            entropyUnit.DataPropertyName = "EntropyUnit";
            streamsTable.Columns.Add(new DataColumn("TotalMassFlowRate", typeof(System.Double)));
            DataGridViewTextBoxColumn totalMassFlowRate = new DataGridViewTextBoxColumn();
            totalMassFlowRate.HeaderText = "TotalMassFlowRate";
            totalMassFlowRate.DataPropertyName = "TotalMassFlowRate";
            streamsTable.Columns.Add(new DataColumn("TotalMassFlowRateUnit", typeof(System.String)));
            DataGridViewTextBoxColumn totalMassFlowRateUnit = new DataGridViewTextBoxColumn();
            totalMassFlowRateUnit.HeaderText = "TotalMassFlowRateUnit";
            totalMassFlowRateUnit.DataPropertyName = "TotalMassFlowRateUnit";
            streamsTable.Columns.Add(new DataColumn("TotalMoleFlowRate", typeof(System.Double)));
            DataGridViewTextBoxColumn totalMoleFlowRate = new DataGridViewTextBoxColumn();
            totalMoleFlowRate.HeaderText = "TotalMoleFlowRate";
            totalMoleFlowRate.DataPropertyName = "TotalMoleFlowRate";
            streamsTable.Columns.Add(new DataColumn("TotalMoleFlowRateUnit", typeof(System.String)));
            DataGridViewTextBoxColumn totalMoleFlowRateUnit = new DataGridViewTextBoxColumn();
            totalMoleFlowRateUnit.HeaderText = "TotalMoleFlowRateUnit";
            totalMoleFlowRateUnit.DataPropertyName = "TotalMoleFlowRateUnit";
            streamsTable.Columns.Add(new DataColumn("LiquidVolumetricFlowRate", typeof(System.Double)));
            DataGridViewTextBoxColumn liquidVolumetricFlowRate = new DataGridViewTextBoxColumn();
            liquidVolumetricFlowRate.HeaderText = "LiquidVolumetricFlowRate";
            liquidVolumetricFlowRate.DataPropertyName = "LiquidVolumetricFlowRate";
            streamsTable.Columns.Add(new DataColumn("LiquidVolumetricFlowRateUnit", typeof(System.String)));
            DataGridViewTextBoxColumn liquidVolumetricFlowRateUnit = new DataGridViewTextBoxColumn();
            liquidVolumetricFlowRateUnit.HeaderText = "LiquidVolumetricFlowRateUnit";
            liquidVolumetricFlowRateUnit.DataPropertyName = "LiquidVolumetricFlowRateUnit";
            streamsTable.Columns.Add(new DataColumn("VaporVolumetricFlowRate", typeof(System.Double)));
            DataGridViewTextBoxColumn vaporVolumetricFlowRate = new DataGridViewTextBoxColumn();
            vaporVolumetricFlowRate.HeaderText = "VaporVolumetricFlowRate";
            vaporVolumetricFlowRate.DataPropertyName = "VaporVolumetricFlowRate";
            streamsTable.Columns.Add(new DataColumn("VaporVolumetricFlowRateUnit", typeof(System.String)));
            DataGridViewTextBoxColumn vaporVolumetricFlowRateUnit = new DataGridViewTextBoxColumn();
            vaporVolumetricFlowRateUnit.HeaderText = "VaporVolumetricFlowRateUnit";
            vaporVolumetricFlowRateUnit.DataPropertyName = "VaporVolumetricFlowRateUnit";
            streamsTable.Columns.Add(new DataColumn("Cost", typeof(System.Double)));
            DataGridViewTextBoxColumn cost = new DataGridViewTextBoxColumn();
            cost.HeaderText = "Cost";
            cost.DataPropertyName = "Cost";
            dataGridView3.Columns.AddRange(streamId, streamName, renewable, sourceUnit, targetUnit, temperature, temperatureUnit, pressure, pressureUnit);
            dataGridView3.Columns.AddRange(moleVaporFraction, enthaply, enthaplyUnit, entropy, entropyUnit, totalMassFlowRate, totalMassFlowRateUnit);
            dataGridView3.Columns.AddRange(totalMoleFlowRate, totalMoleFlowRateUnit, liquidVolumetricFlowRate, liquidVolumetricFlowRateUnit);
            dataGridView3.Columns.AddRange(vaporVolumetricFlowRate, vaporVolumetricFlowRateUnit, cost);

            DataRow row;
            foreach (Stream stream in streams)
            {
                row = streamsTable.NewRow();
                row["StreamId"] = stream.StreamID;
                row["StreamName"] = stream.StreamName;
                string renewableStream = "no";
                for (int i = 0; i < feedStreamIds.Length; i++)
                {
                    if (feedStreamIds[i] == stream.StreamID)
                    {
                        renewableStream = feedStreamRenewables[i];
                    }
                }
                row["Renewable"] = renewableStream;
                row["SourceUnit"] = stream.SourceUnitOperation;
                row["TargetUnit"] = stream.TargetUnitOperation;
                row["Temperature"] = stream.Temperature;
                row["TemperatureUnit"] = stream.TemperatureUnit;
                row["Pressure"] = stream.Pressure;
                row["PressureUnit"] = stream.PressureUnit;
                row["MoleVaporFraction"] = stream.MoleVaporFraction;
                row["Enthaply"] = stream.Enthalpy;
                row["EnthaplyUnit"] = stream.EnthalpyUnit;
                row["Entropy"] = stream.Entropy;
                row["EntropyUnit"] = stream.EntropyUnit;
                row["TotalMassFlowRate"] = stream.TotalMassFlowRate;
                row["TotalMassFlowRateUnit"] = stream.TotalMassFlowRateUnit;
                row["TotalMoleFlowRate"] = stream.TotalMoleFlowRate;
                row["TotalMoleFlowRateUnit"] = stream.TotalMoleFlowRateUnit;
                row["LiquidVolumetricFlowRate"] = stream.LiquidVolumetricFlowRate;
                row["LiquidVolumetricFlowRateUnit"] = stream.LiquidVolumetricFlowRateUnit;
                row["VaporVolumetricFlowRate"] = stream.VaporVolumetricFlowRate;
                row["VaporVolumetricFlowRateUnit"] = stream.VaporVolumetricFlowRateUnit;
                row["Cost"] = stream.Cost;
                streamsTable.Rows.Add(row);
            }
            return streamsTable;
        }

        DataTable CreateProductStreamsDataTable(string tableName, Stream[] streams, DataTable componentTable)
        {
            dataGridView4.AutoGenerateColumns = false;
            DataTable streamsTable = new DataTable(tableName);
            streamsTable.Columns.Add(new DataColumn("StreamId", typeof(System.String)));
            DataGridViewTextBoxColumn streamId = new DataGridViewTextBoxColumn();
            streamId.HeaderText = "StreamID";
            streamId.DataPropertyName = "StreamId";
            streamsTable.Columns.Add(new DataColumn("StreamName", typeof(System.String)));
            DataGridViewTextBoxColumn streamName = new DataGridViewTextBoxColumn();
            streamName.HeaderText = "StreamName";
            streamName.DataPropertyName = "StreamName";
            streamsTable.Columns.Add(new DataColumn("ProductOrWaste", typeof(System.String)));
            DataGridViewComboBoxColumn productOrWaste = new DataGridViewComboBoxColumn();
            List<string> productOrWasteOptions = new List<string>() { "", "N/A", "yes", "no" };
            productOrWaste.DataSource = productOrWasteOptions;
            productOrWaste.HeaderText = "ProductOrWaste";
            productOrWaste.DataPropertyName = "ProductOrWaste";
            streamsTable.Columns.Add(new DataColumn("EcoProduct", typeof(System.String)));
            DataGridViewComboBoxColumn ecoProduct = new DataGridViewComboBoxColumn();
            List<string> ecoProductOptions = new List<string>() { "", "N/A", "yes", "no" };
            ecoProduct.DataSource = ecoProductOptions;
            ecoProduct.HeaderText = "EcoProduct";
            ecoProduct.DataPropertyName = "EcoProduct";
            streamsTable.Columns.Add(new DataColumn("Polluted/NonPolluted", typeof(System.String)));
            DataGridViewComboBoxColumn pollutedOrNonPolluted = new DataGridViewComboBoxColumn();
            List<string> pollutedOrNonPollutedOptions = new List<string>() { "", "N/A", "yes", "no" };
            pollutedOrNonPolluted.DataSource = pollutedOrNonPollutedOptions;
            pollutedOrNonPolluted.HeaderText = "Polluted/NonPolluted";
            pollutedOrNonPolluted.DataPropertyName = "Polluted/NonPolluted";
            streamsTable.Columns.Add(new DataColumn("Renewable", typeof(System.String)));
            DataGridViewComboBoxColumn renewable = new DataGridViewComboBoxColumn();
            List<string> renewableOptions = new List<string>() { "", "yes", "no" };
            renewable.DataSource = renewableOptions;
            renewable.HeaderText = "Renewable";
            renewable.DataPropertyName = "Renewable";
            streamsTable.Columns.Add(new DataColumn("SourceUnit", typeof(System.Int32)));
            DataGridViewTextBoxColumn sourceUnit = new DataGridViewTextBoxColumn();
            sourceUnit.HeaderText = "SourceUnit";
            sourceUnit.DataPropertyName = "SourceUnit";
            streamsTable.Columns.Add(new DataColumn("TargetUnit", typeof(System.Int32)));
            DataGridViewTextBoxColumn targetUnit = new DataGridViewTextBoxColumn();
            targetUnit.HeaderText = "TargetUnit";
            targetUnit.DataPropertyName = "TargetUnit";
            streamsTable.Columns.Add(new DataColumn("Temperature", typeof(System.Double)));
            DataGridViewTextBoxColumn temperature = new DataGridViewTextBoxColumn();
            temperature.HeaderText = "Temperature";
            temperature.DataPropertyName = "Temperature";
            streamsTable.Columns.Add(new DataColumn("TemperatureUnit", typeof(System.String)));
            DataGridViewTextBoxColumn temperatureUnit = new DataGridViewTextBoxColumn();
            temperatureUnit.HeaderText = "TemperatureUnit";
            temperatureUnit.DataPropertyName = "TemperatureUnit";
            streamsTable.Columns.Add(new DataColumn("Pressure", typeof(System.Double)));
            DataGridViewTextBoxColumn pressure = new DataGridViewTextBoxColumn();
            pressure.HeaderText = "Pressure";
            pressure.DataPropertyName = "Pressure";
            streamsTable.Columns.Add(new DataColumn("PressureUnit", typeof(System.String)));
            DataGridViewTextBoxColumn pressureUnit = new DataGridViewTextBoxColumn();
            pressureUnit.HeaderText = "PressureUnit";
            pressureUnit.DataPropertyName = "PressureUnit";
            streamsTable.Columns.Add(new DataColumn("MoleVaporFraction", typeof(System.Double)));
            DataGridViewTextBoxColumn moleVaporFraction = new DataGridViewTextBoxColumn();
            moleVaporFraction.HeaderText = "MoleVaporFraction";
            moleVaporFraction.DataPropertyName = "MoleVaporFraction";
            streamsTable.Columns.Add(new DataColumn("Enthaply", typeof(System.Double)));
            DataGridViewTextBoxColumn enthaply = new DataGridViewTextBoxColumn();
            enthaply.HeaderText = "Enthaply";
            enthaply.DataPropertyName = "Enthaply";
            streamsTable.Columns.Add(new DataColumn("EnthaplyUnit", typeof(System.String)));
            DataGridViewTextBoxColumn enthaplyUnit = new DataGridViewTextBoxColumn();
            enthaplyUnit.HeaderText = "EnthaplyUnit";
            enthaplyUnit.DataPropertyName = "EnthaplyUnit";
            streamsTable.Columns.Add(new DataColumn("Entropy", typeof(System.Double)));
            DataGridViewTextBoxColumn entropy = new DataGridViewTextBoxColumn();
            entropy.HeaderText = "Entropy";
            entropy.DataPropertyName = "Entropy";
            streamsTable.Columns.Add(new DataColumn("EntropyUnit", typeof(System.String)));
            DataGridViewTextBoxColumn entropyUnit = new DataGridViewTextBoxColumn();
            entropyUnit.HeaderText = "EntropyUnit";
            entropyUnit.DataPropertyName = "EntropyUnit";
            streamsTable.Columns.Add(new DataColumn("TotalMassFlowRate", typeof(System.Double)));
            DataGridViewTextBoxColumn totalMassFlowRate = new DataGridViewTextBoxColumn();
            totalMassFlowRate.HeaderText = "TotalMassFlowRate";
            totalMassFlowRate.DataPropertyName = "TotalMassFlowRate";
            streamsTable.Columns.Add(new DataColumn("TotalMassFlowRateUnit", typeof(System.String)));
            DataGridViewTextBoxColumn totalMassFlowRateUnit = new DataGridViewTextBoxColumn();
            totalMassFlowRateUnit.HeaderText = "TotalMassFlowRateUnit";
            totalMassFlowRateUnit.DataPropertyName = "TotalMassFlowRateUnit";
            streamsTable.Columns.Add(new DataColumn("TotalMoleFlowRate", typeof(System.Double)));
            DataGridViewTextBoxColumn totalMoleFlowRate = new DataGridViewTextBoxColumn();
            totalMoleFlowRate.HeaderText = "TotalMoleFlowRate";
            totalMoleFlowRate.DataPropertyName = "TotalMoleFlowRate";
            streamsTable.Columns.Add(new DataColumn("TotalMoleFlowRateUnit", typeof(System.String)));
            DataGridViewTextBoxColumn totalMoleFlowRateUnit = new DataGridViewTextBoxColumn();
            totalMoleFlowRateUnit.HeaderText = "TotalMoleFlowRateUnit";
            totalMoleFlowRateUnit.DataPropertyName = "TotalMoleFlowRateUnit";
            streamsTable.Columns.Add(new DataColumn("LiquidVolumetricFlowRate", typeof(System.Double)));
            DataGridViewTextBoxColumn liquidVolumetricFlowRate = new DataGridViewTextBoxColumn();
            liquidVolumetricFlowRate.HeaderText = "LiquidVolumetricFlowRate";
            liquidVolumetricFlowRate.DataPropertyName = "LiquidVolumetricFlowRate";
            streamsTable.Columns.Add(new DataColumn("LiquidVolumetricFlowRateUnit", typeof(System.String)));
            DataGridViewTextBoxColumn liquidVolumetricFlowRateUnit = new DataGridViewTextBoxColumn();
            liquidVolumetricFlowRateUnit.HeaderText = "LiquidVolumetricFlowRateUnit";
            liquidVolumetricFlowRateUnit.DataPropertyName = "LiquidVolumetricFlowRateUnit";
            streamsTable.Columns.Add(new DataColumn("VaporVolumetricFlowRate", typeof(System.Double)));
            DataGridViewTextBoxColumn vaporVolumetricFlowRate = new DataGridViewTextBoxColumn();
            vaporVolumetricFlowRate.HeaderText = "VaporVolumetricFlowRate";
            vaporVolumetricFlowRate.DataPropertyName = "VaporVolumetricFlowRate";
            streamsTable.Columns.Add(new DataColumn("VaporVolumetricFlowRateUnit", typeof(System.String)));
            DataGridViewTextBoxColumn vaporVolumetricFlowRateUnit = new DataGridViewTextBoxColumn();
            vaporVolumetricFlowRateUnit.HeaderText = "VaporVolumetricFlowRateUnit";
            vaporVolumetricFlowRateUnit.DataPropertyName = "VaporVolumetricFlowRateUnit";
            streamsTable.Columns.Add(new DataColumn("Cost", typeof(System.Double)));
            DataGridViewTextBoxColumn cost = new DataGridViewTextBoxColumn();
            cost.HeaderText = "Cost";
            cost.DataPropertyName = "Cost";
            dataGridView4.Columns.AddRange(streamId, streamName, productOrWaste, ecoProduct, pollutedOrNonPolluted, renewable, sourceUnit, targetUnit);
            dataGridView4.Columns.AddRange(temperature, temperatureUnit, pressure, pressureUnit, moleVaporFraction, enthaply, enthaplyUnit, entropy);
            dataGridView4.Columns.AddRange(entropyUnit, totalMassFlowRate, totalMassFlowRateUnit, totalMoleFlowRate, totalMoleFlowRateUnit);
            dataGridView4.Columns.AddRange(liquidVolumetricFlowRate, liquidVolumetricFlowRateUnit, vaporVolumetricFlowRate, vaporVolumetricFlowRateUnit);
            dataGridView4.Columns.AddRange(cost);

            DataRow row;
            foreach (Stream stream in streams)
            {
                row = streamsTable.NewRow();
                row["StreamId"] = stream.StreamID;
                row["StreamName"] = stream.StreamName;
                string productOrWasteValue = "";
                string ecoProductValue = "";
                string pollutedOrNonPollutedValue = "";
                string renewableValue = "no";
                for (int i = 0; i < productStreamIds.Length; i++)
                {
                    if (productStreamIds[i] == stream.StreamID)
                    {
                        productOrWasteValue = productStreamProductOrWastes[i];
                        ecoProductValue = productStreamEcoProducts[i];
                        pollutedOrNonPollutedValue = productStreamPollutedNonPolluteds[i];
                        renewableValue = productStreamRenewables[i];
                    }
                }
                row["ProductOrWaste"] = productOrWasteValue;
                row["EcoProduct"] = ecoProductValue;
                row["Polluted/NonPolluted"] = pollutedOrNonPollutedValue;
                row["Renewable"] = renewableValue;
                row["SourceUnit"] = stream.SourceUnitOperation;
                row["TargetUnit"] = stream.TargetUnitOperation;
                row["Temperature"] = stream.Temperature;
                row["TemperatureUnit"] = stream.TemperatureUnit;
                row["Pressure"] = stream.Pressure;
                row["PressureUnit"] = stream.PressureUnit;
                row["MoleVaporFraction"] = stream.MoleVaporFraction;
                row["Enthaply"] = stream.Enthalpy;
                row["EnthaplyUnit"] = stream.EnthalpyUnit;
                row["Entropy"] = stream.Entropy;
                row["EntropyUnit"] = stream.EntropyUnit;
                row["TotalMassFlowRate"] = stream.TotalMassFlowRate;
                row["TotalMassFlowRateUnit"] = stream.TotalMassFlowRateUnit;
                row["TotalMoleFlowRate"] = stream.TotalMoleFlowRate;
                row["TotalMoleFlowRateUnit"] = stream.TotalMoleFlowRateUnit;
                row["LiquidVolumetricFlowRate"] = stream.LiquidVolumetricFlowRate;
                row["LiquidVolumetricFlowRateUnit"] = stream.LiquidVolumetricFlowRateUnit;
                row["VaporVolumetricFlowRate"] = stream.VaporVolumetricFlowRate;
                row["VaporVolumetricFlowRateUnit"] = stream.VaporVolumetricFlowRateUnit;
                row["Cost"] = stream.Cost;
                streamsTable.Rows.Add(row);
            }
            return streamsTable;
        }

        DataTable CreateUnitOperationDataTable(UnitOperation[] unitOps)
        {
            DataTable unitOpTable = new DataTable("UnitOps");
            DataRow row; // Create new DataColumn, set DataType,  
            unitOpTable.Columns.Add(new DataColumn("UnitOpId", typeof(System.Int32)));
            unitOpTable.Columns.Add(new DataColumn("UnitOpLabel", typeof(System.String)));
            unitOpTable.Columns.Add(new DataColumn("Category", typeof(System.String)));
            unitOpTable.Columns.Add(new DataColumn("HeatAdded", typeof(System.Double)));
            unitOpTable.Columns.Add(new DataColumn("Power", typeof(System.Double)));
            unitOpTable.Columns.Add(new DataColumn("TotalPurchaseCost", typeof(System.Double)));
            unitOpTable.Columns.Add(new DataColumn("TotalInstalledCost", typeof(System.Double)));

            for (int i = 0; i < 250; i++)
            {
                unitOpTable.Columns.Add(new DataColumn("Spec" + i.ToString(), typeof(System.Double)));
            }

            int nUnitOps = unitOps.Length;
            for (int i = 0; i < nUnitOps; i++)
            {
                //StreamComponent components = new StreamComponent(streamIDS[0], i, wrapper);
                row = unitOpTable.NewRow();
                row["UnitOpId"] = unitOps[i].UnitOpId;
                row["UnitOpLabel"] = unitOps[i].Label;
                row["Category"] = unitOps[i].Category;
                row["HeatAdded"] = unitOps[i].HeatAdded;
                row["Power"] = unitOps[i].Power;
                row["TotalPurchaseCost"] = unitOps[i].TotalPurchaseCost;
                row["TotalInstalledCost"] = unitOps[i].TotalInstalledCost;
                for (int j = 0; j < 250; j++)
                {
                    row["Spec" + j.ToString()] = unitOps[i].Specification[j];
                }
                unitOpTable.Rows.Add(row);
            }
            return unitOpTable;
        }

        // Given a worksheet, a column name, and a row index, gets the cell at the specified column and row.
        private DocumentFormat.OpenXml.Spreadsheet.Cell GetSpreadsheetCell(DocumentFormat.OpenXml.Spreadsheet.Worksheet worksheet, string columnName, uint rowIndex)
        {
            IEnumerable<DocumentFormat.OpenXml.Spreadsheet.Row> rows = worksheet.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.SheetData>().Elements<DocumentFormat.OpenXml.Spreadsheet.Row>().Where(r => r.RowIndex == rowIndex);
            if (rows.Count() == 0)
            {
                // A cell does not exist at the specified row.
                return null;
            }

            IEnumerable<DocumentFormat.OpenXml.Spreadsheet.Cell> cells = rows.First().Elements<DocumentFormat.OpenXml.Spreadsheet.Cell>().Where(c => string.Compare(c.CellReference.Value, columnName + rowIndex, true) == 0);
            if (cells.Count() == 0)
            {
                // A cell does not exist at the specified column, in the specified row.
                return null;
            }

            return cells.First();
        }

        // Given a worksheet, a column name, and a row index, gets the cell at the specified column and row.
        private void SetSpreadsheetCellValue(DocumentFormat.OpenXml.Spreadsheet.Worksheet worksheet, string columnName, int rowIndex, string value)
        {
            IEnumerable<DocumentFormat.OpenXml.Spreadsheet.Row> rows = worksheet.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.SheetData>().Elements<DocumentFormat.OpenXml.Spreadsheet.Row>().Where(r => r.RowIndex == rowIndex);
            if (rows.Count() == 0)
            {
                // A cell does not exist at the specified row.
                return;
            }

            IEnumerable<DocumentFormat.OpenXml.Spreadsheet.Cell> cells = rows.First().Elements<DocumentFormat.OpenXml.Spreadsheet.Cell>().Where(c => string.Compare(c.CellReference.Value, columnName + rowIndex, true) == 0);
            if (cells.Count() == 0)
            {
                // A cell does not exist at the specified column, in the specified row.
                return;
            }

            DocumentFormat.OpenXml.Spreadsheet.Cell cell = cells.First();
            if (cell != null)
            {
                // The specified cell does not exist.
                cell.CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue(value);
                cell.DataType = new DocumentFormat.OpenXml.EnumValue<DocumentFormat.OpenXml.Spreadsheet.CellValues>(DocumentFormat.OpenXml.Spreadsheet.CellValues.String);
            }
        }

        // Given a worksheet, a column name, and a row index, gets the cell at the specified column and row.
        //private void SetSpreadsheetCellValueNumeric(DocumentFormat.OpenXml.Spreadsheet.Worksheet worksheet, string columnName, int rowIndex, string value)
        //{
        //    IEnumerable<DocumentFormat.OpenXml.Spreadsheet.Row> rows = worksheet.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.SheetData>().Elements<DocumentFormat.OpenXml.Spreadsheet.Row>().Where(r => r.RowIndex == rowIndex);
        //    if (rows.Count() == 0)
        //    {
        //        // A cell does not exist at the specified row.
        //        return;
        //    }

        //    IEnumerable<DocumentFormat.OpenXml.Spreadsheet.Cell> cells = rows.First().Elements<DocumentFormat.OpenXml.Spreadsheet.Cell>().Where(c => string.Compare(c.CellReference.Value, columnName + rowIndex, true) == 0);
        //    if (cells.Count() == 0)
        //    {
        //        // A cell does not exist at the specified column, in the specified row.
        //        return;
        //    }

        //    DocumentFormat.OpenXml.Spreadsheet.Cell cell = cells.First();
        //    if (cell != null)
        //    {
        //        // The specified cell does not exist.
        //        cell.CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue(value);
        //        cell.DataType = new DocumentFormat.OpenXml.EnumValue<DocumentFormat.OpenXml.Spreadsheet.CellValues>(DocumentFormat.OpenXml.Spreadsheet.CellValues.String);
        //    }
        //}

        // Given a worksheet, a column name, and a row index, gets the cell at the specified column and row.
        private void SetSpreadsheetCellValue(DocumentFormat.OpenXml.Spreadsheet.Worksheet worksheet, string columnName, int rowIndex, int value)
        {
            IEnumerable<DocumentFormat.OpenXml.Spreadsheet.Row> rows = worksheet.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.SheetData>().Elements<DocumentFormat.OpenXml.Spreadsheet.Row>().Where(r => r.RowIndex == rowIndex);
            if (rows.Count() == 0)
            {
                // A cell does not exist at the specified row.
                return;
            }

            IEnumerable<DocumentFormat.OpenXml.Spreadsheet.Cell> cells = rows.First().Elements<DocumentFormat.OpenXml.Spreadsheet.Cell>().Where(c => string.Compare(c.CellReference.Value, columnName + rowIndex, true) == 0);
            if (cells.Count() == 0)
            {
                // A cell does not exist at the specified column, in the specified row.
                return;
            }

            DocumentFormat.OpenXml.Spreadsheet.Cell cell = cells.First();
            if (cell != null)
            {
                // The specified cell does not exist.
                cell.CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue(value.ToString());
                cell.DataType = new DocumentFormat.OpenXml.EnumValue<DocumentFormat.OpenXml.Spreadsheet.CellValues>(DocumentFormat.OpenXml.Spreadsheet.CellValues.Number);
            }
        }

        // Given a worksheet, a column name, and a row index, gets the cell at the specified column and row.
        private void SetSpreadsheetCellValue(DocumentFormat.OpenXml.Spreadsheet.Worksheet worksheet, string columnName, int rowIndex, double value)
        {
            IEnumerable<DocumentFormat.OpenXml.Spreadsheet.Row> rows = worksheet.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.SheetData>().Elements<DocumentFormat.OpenXml.Spreadsheet.Row>().Where(r => r.RowIndex == rowIndex);
            if (rows.Count() == 0)
            {
                // A cell does not exist at the specified row.
                return;
            }

            IEnumerable<DocumentFormat.OpenXml.Spreadsheet.Cell> cells = rows.First().Elements<DocumentFormat.OpenXml.Spreadsheet.Cell>().Where(c => string.Compare(c.CellReference.Value, columnName + rowIndex, true) == 0);
            if (cells.Count() == 0)
            {
                // A cell does not exist at the specified column, in the specified row.
                return;
            }

            DocumentFormat.OpenXml.Spreadsheet.Cell cell = cells.First();
            if (cell != null)
            {
                // The specified cell does not exist.
                cell.CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue(value.ToString());
                cell.DataType = new DocumentFormat.OpenXml.EnumValue<DocumentFormat.OpenXml.Spreadsheet.CellValues>(DocumentFormat.OpenXml.Spreadsheet.CellValues.Number);
            }
        }

        void UpdateSpreadsheetChangeLog(DocumentFormat.OpenXml.Packaging.SpreadsheetDocument spreadsheet, string chemcadFileName, string excelFileName)
        {
            string message = String.Empty;
            IEnumerable<DocumentFormat.OpenXml.Spreadsheet.Sheet> sheets = spreadsheet.WorkbookPart.Workbook.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.Sheets>().Elements<DocumentFormat.OpenXml.Spreadsheet.Sheet>().Where(s => s.Name == "Change Log");
            if (sheets.Count() == 0)
            {
                // The specified worksheet does not exist.
                return;
            }
            string relationshipId = sheets.First().Id.Value;
            DocumentFormat.OpenXml.Packaging.WorksheetPart worksheetPart = (DocumentFormat.OpenXml.Packaging.WorksheetPart)spreadsheet.WorkbookPart.GetPartById(relationshipId);
            DocumentFormat.OpenXml.Spreadsheet.Cell currentCell = this.GetSpreadsheetCell(worksheetPart.Worksheet, "A", (uint)1);
            bool newFile = false;
            if (currentCell == null)
            {
                message = "The spreadsheet was created on " + message + DateTime.Now.ToString();
                currentCell = this.InsertCellInWorksheet("A", (uint)1, worksheetPart);
                currentCell.CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue(message);
                currentCell.DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.String;
                newFile = true;
            }
            if (!newFile)
            {
                uint currentUpdate = 1;
                currentCell = this.GetSpreadsheetCell(worksheetPart.Worksheet, "A", ++currentUpdate);
                while (currentCell != null)
                {
                    currentCell = this.GetSpreadsheetCell(worksheetPart.Worksheet, "A", ++currentUpdate);
                }
                message = "The spreadsheet was updated on " + message + DateTime.Now.ToString();
                currentCell = this.InsertCellInWorksheet("A", currentUpdate, worksheetPart);
                currentCell.CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue(message);
                currentCell.DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.String;
            }
            sheets = spreadsheet.WorkbookPart.Workbook.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.Sheets>().Elements<DocumentFormat.OpenXml.Spreadsheet.Sheet>().Where(s => s.Name == "I. Stream & Compound Data");
            if (sheets.Count() == 0)
            {
                // The specified worksheet does not exist.
                return;
            }
            relationshipId = sheets.First().Id.Value;
            worksheetPart = (DocumentFormat.OpenXml.Packaging.WorksheetPart)spreadsheet.WorkbookPart.GetPartById(relationshipId);
            //if (newFile) currentCell = this.InsertCellInWorksheet("C", (uint)13, worksheetPart);
            currentCell = this.InsertCellInWorksheet("C", (uint)13, worksheetPart);
            currentCell.CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue(message);
            currentCell.DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.String;
        }

        // Given a column name, a row index, and a WorksheetPart, inserts a cell into the worksheet. 
        // If the cell already exists, returns it. 
        private DocumentFormat.OpenXml.Spreadsheet.Cell InsertCellInWorksheet(string columnName, uint rowIndex, DocumentFormat.OpenXml.Packaging.WorksheetPart worksheetPart)
        {
            DocumentFormat.OpenXml.Spreadsheet.Worksheet worksheet = worksheetPart.Worksheet;
            DocumentFormat.OpenXml.Spreadsheet.SheetData sheetData = worksheet.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.SheetData>();
            string cellReference = columnName + rowIndex;

            // If the worksheet does not contain a row with the specified row index, insert one.
            DocumentFormat.OpenXml.Spreadsheet.Row row;
            if (sheetData.Elements<DocumentFormat.OpenXml.Spreadsheet.Row>().Where(r => r.RowIndex == rowIndex).Count() != 0)
            {
                row = sheetData.Elements<DocumentFormat.OpenXml.Spreadsheet.Row>().Where(r => r.RowIndex == rowIndex).First();
            }
            else
            {
                row = new DocumentFormat.OpenXml.Spreadsheet.Row() { RowIndex = rowIndex };
                sheetData.Append(row);
            }

            // If there is not a cell with the specified column name, insert one.  
            if (row.Elements<DocumentFormat.OpenXml.Spreadsheet.Cell>().Where(c => c.CellReference.Value == columnName + rowIndex).Count() > 0)
            {
                return row.Elements<DocumentFormat.OpenXml.Spreadsheet.Cell>().Where(c => c.CellReference.Value == cellReference).First();
            }
            else
            {
                // Cells must be in sequential order according to CellReference. Determine where to insert the new cell.
                DocumentFormat.OpenXml.Spreadsheet.Cell refCell = null;
                foreach (DocumentFormat.OpenXml.Spreadsheet.Cell cell in row.Elements<DocumentFormat.OpenXml.Spreadsheet.Cell>())
                {
                    if (string.Compare(cell.CellReference.Value, cellReference, true) > 0)
                    {
                        refCell = cell;
                        break;
                    }
                }

                DocumentFormat.OpenXml.Spreadsheet.Cell newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = cellReference };
                row.InsertBefore(newCell, refCell);

                worksheet.Save();
                return newCell;
            }
        }

        private void AddComponentsToSpreadsheet(DocumentFormat.OpenXml.Packaging.SpreadsheetDocument spreadsheet, Stream stream)
        {
            IEnumerable<DocumentFormat.OpenXml.Spreadsheet.Sheet> sheets = spreadsheet.WorkbookPart.Workbook.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.Sheets>().Elements<DocumentFormat.OpenXml.Spreadsheet.Sheet>().Where(s => s.Name == "I. Stream & Compound Data");
            if (sheets.Count() == 0)
            {
                // The specified worksheet does not exist.
                return;
            }
            string relationshipId = sheets.First().Id.Value;
            DocumentFormat.OpenXml.Packaging.WorksheetPart worksheetPart = (DocumentFormat.OpenXml.Packaging.WorksheetPart)spreadsheet.WorkbookPart.GetPartById(relationshipId);
            for (int i = 0; i < stream.NumberOfComponents; i++)
            {
                SetSpreadsheetCellValue(worksheetPart.Worksheet, "D", i + 32, stream.ComponentNames[i]);
                SetSpreadsheetCellValue(worksheetPart.Worksheet, "E", i + 32, stream.MolecularFormula(i));
                SetSpreadsheetCellValue(worksheetPart.Worksheet, "F", i + 32, stream.MolecularWeight(i));
                SetSpreadsheetCellValue(worksheetPart.Worksheet, "G", i + 32, stream.casNumber(i));
                SetSpreadsheetCellValue(worksheetPart.Worksheet, "AY", i + 262, stream.boilingPoint(i));
            }
        }

        private void GetReferenceConditionsFromSpreadsheet(DocumentFormat.OpenXml.Packaging.SpreadsheetDocument spreadsheet)
        {
            IEnumerable<DocumentFormat.OpenXml.Spreadsheet.Sheet> sheets = spreadsheet.WorkbookPart.Workbook.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.Sheets>().Elements<DocumentFormat.OpenXml.Spreadsheet.Sheet>().Where(s => s.Name == "I. Stream & Compound Data");
            if (sheets.Count() == 0)
            {
                // The specified worksheet does not exist.
                return;
            }
            string relationshipId = sheets.First().Id.Value;
            DocumentFormat.OpenXml.Packaging.WorksheetPart worksheetPart = (DocumentFormat.OpenXml.Packaging.WorksheetPart)spreadsheet.WorkbookPart.GetPartById(relationshipId);
            this.numericUpDown1.Value = Convert.ToDecimal(this.GetSpreadsheetCell(worksheetPart.Worksheet, "F", (uint)19).CellValue.Text);
            this.numericUpDown2.Value = Convert.ToDecimal(this.GetSpreadsheetCell(worksheetPart.Worksheet, "F", (uint)20).CellValue.Text);
            this.numericUpDown3.Value = Convert.ToDecimal(this.GetSpreadsheetCell(worksheetPart.Worksheet, "F", (uint)21).CellValue.Text);
        }

        private void UpdateReferenceConditionsInSpreadsheet(DocumentFormat.OpenXml.Packaging.SpreadsheetDocument spreadsheet)
        {
            IEnumerable<DocumentFormat.OpenXml.Spreadsheet.Sheet> sheets = spreadsheet.WorkbookPart.Workbook.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.Sheets>().Elements<DocumentFormat.OpenXml.Spreadsheet.Sheet>().Where(s => s.Name == "I. Stream & Compound Data");
            if (sheets.Count() == 0)
            {
                // The specified worksheet does not exist.
                return;
            }
            string relationshipId = sheets.First().Id.Value;
            DocumentFormat.OpenXml.Packaging.WorksheetPart worksheetPart = (DocumentFormat.OpenXml.Packaging.WorksheetPart)spreadsheet.WorkbookPart.GetPartById(relationshipId);
            if (processReferenceTemperatureUnit == "Kelvin") processReferenceTemperature = processReferenceTemperature - 273.15;
            if (processReferenceTemperatureUnit == "Rankine") processReferenceTemperature = (processReferenceTemperature - 491.67) * 5 / 9;
            if (processReferenceTemperatureUnit == "Farenheit") processReferenceTemperature = (processReferenceTemperature - 32) * 5 / 9; // Temperature in Farenheit.
            SetSpreadsheetCellValue(worksheetPart.Worksheet, "F", 19, processReferenceTemperature);
            if (referenceTemperatureUnit == "Kelvin") referenceTemperature = referenceTemperature - 273.15;
            if (referenceTemperatureUnit == "Rankine") referenceTemperature = (referenceTemperature - 491.67) * 5 / 9;
            if (referenceTemperatureUnit == "Farenheit") referenceTemperature = (referenceTemperature - 32) * 5 / 9; // Temperature in Farenheit.
            SetSpreadsheetCellValue(worksheetPart.Worksheet, "F", 20, referenceTemperature);
            if (referencePressureUnit == "atm") referencePressure = referencePressure * 1.01325e+02;
            if (referencePressureUnit == "psia") referencePressure = referencePressure * 6.89476;
            if (referencePressureUnit == "psig") referencePressure = referencePressure * 6.89476 + 1.01325e+02;
            if (referencePressureUnit == "torr") referencePressure = referencePressure * 1.33322e-01;
            if (referencePressureUnit == "mmHg") referencePressure = referencePressure * 1.33322e-01;
            if (referencePressureUnit == "Pa") referencePressure = referencePressure / 1000;
            if (referencePressureUnit == "MPa G") referencePressure = referencePressure * 1000 + 1.01325e+02;
            if (referencePressureUnit == "bar") referencePressure = referencePressure * 1e+02;
            if (referencePressureUnit == "bar G") referencePressure = referencePressure * 1e+02 + 1.01325e+02;
            if (referencePressureUnit == "mbar") referencePressure = referencePressure * 0.1;
            if (referencePressureUnit == "kg/cm2") referencePressure = referencePressure * 98.0665;
            if (referencePressureUnit == "kg/cm2 G") referencePressure = referencePressure * 98.0665 + 1.01325e+02;
            if (referencePressureUnit == "in-water") referencePressure = referencePressure * 2.49089e-01;
            if (referencePressureUnit == "mm-water") referencePressure = referencePressure * 9.80665e-03;
            SetSpreadsheetCellValue(worksheetPart.Worksheet, "F", 21, referencePressure);
        }

        private void GetFeedStreamRenewableFromSpreadsheet(DocumentFormat.OpenXml.Packaging.SpreadsheetDocument spreadsheet, ref int[] streams, ref string[] renewables)
        {
            IEnumerable<DocumentFormat.OpenXml.Spreadsheet.Sheet> sheets = spreadsheet.WorkbookPart.Workbook.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.Sheets>().Elements<DocumentFormat.OpenXml.Spreadsheet.Sheet>().Where(s => s.Name == "I. Stream & Compound Data");
            if (sheets.Count() == 0)
            {
                // The specified worksheet does not exist.
                return;
            }
            string relationshipId = sheets.First().Id.Value;
            DocumentFormat.OpenXml.Packaging.WorksheetPart worksheetPart = (DocumentFormat.OpenXml.Packaging.WorksheetPart)spreadsheet.WorkbookPart.GetPartById(relationshipId);
            string[] streamColumns = { "I", "L", "O", "R", "U", "X", "AA", "AD", "AG", "AJ", "AM", "AP", "AS", "AV", "AY", "BB", "BE", "BH", "BK", "BN", "BQ", "BT", "BW", "BZ", "CC", "CF", "CI", "CL", "CO", "CR", "CU", "CX", "DA", "DD", "DG", "DJ", "DM", "DP", "DS", "DV" };
            List<int> streamList = new List<int>();
            List<string> renewableList = new List<string>();
            foreach (string column in streamColumns)
            {
                DocumentFormat.OpenXml.Spreadsheet.Cell streamIdCell = this.GetSpreadsheetCell(worksheetPart.Worksheet, column, (uint)28);
                DocumentFormat.OpenXml.Spreadsheet.Cell renewableCell = this.GetSpreadsheetCell(worksheetPart.Worksheet, column, (uint)30);
                if (streamIdCell.CellValue != null)
                {
                    streamList.Add(Convert.ToInt32(streamIdCell.CellValue.Text));
                    if (renewableCell.CellValue != null)
                    {
                        renewableList.Add(renewableCell.CellValue.Text);
                    }
                    else renewableList.Add("no");
                }
            }
            streams = streamList.ToArray<int>();
            renewables = renewableList.ToArray<string>();
        }

        private void GetProductStreamInformationFromSpreadsheet(DocumentFormat.OpenXml.Packaging.SpreadsheetDocument spreadsheet, ref int[] streams, ref string[] productOrWastes, ref string[] ecoProducts, ref string[] pollutedNotPolluted, ref string[] renewables)
        {
            IEnumerable<DocumentFormat.OpenXml.Spreadsheet.Sheet> sheets = spreadsheet.WorkbookPart.Workbook.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.Sheets>().Elements<DocumentFormat.OpenXml.Spreadsheet.Sheet>().Where(s => s.Name == "I. Stream & Compound Data");
            if (sheets.Count() == 0)
            {
                // The specified worksheet does not exist.
                return;
            }
            string relationshipId = sheets.First().Id.Value;
            DocumentFormat.OpenXml.Packaging.WorksheetPart worksheetPart = (DocumentFormat.OpenXml.Packaging.WorksheetPart)spreadsheet.WorkbookPart.GetPartById(relationshipId);
            string[] streamColumns = { "I", "L", "O", "R", "U", "X", "AA", "AD", "AG", "AJ", "AM", "AP", "AS", "AV", "AY", "BB", "BE", "BH", "BK", "BN", "BQ", "BT", "BW", "BZ", "CC", "CF", "CI", "CL", "CO", "CR", "CU", "CX", "DA", "DD", "DG", "DJ", "DM", "DP", "DS", "DV" };
            List<int> streamList = new List<int>();
            List<string> productOrWasteList = new List<string>();
            List<string> ecoProductList = new List<string>();
            List<string> pollutedOrNonpollutedProductList = new List<string>();
            List<string> renewableList = new List<string>();
            for (uint i = 0; i < streamColumns.Length; i++)
            {
                DocumentFormat.OpenXml.Spreadsheet.Cell streamIdCell = this.GetSpreadsheetCell(worksheetPart.Worksheet, streamColumns[i], (uint)178);
                DocumentFormat.OpenXml.Spreadsheet.Cell productOrWasteCell = this.GetSpreadsheetCell(worksheetPart.Worksheet, streamColumns[i], (uint)179);
                DocumentFormat.OpenXml.Spreadsheet.Cell ecoProductCell = this.GetSpreadsheetCell(worksheetPart.Worksheet, streamColumns[i], (uint)181);
                DocumentFormat.OpenXml.Spreadsheet.Cell pollutedORNotPollutedCell = this.GetSpreadsheetCell(worksheetPart.Worksheet, streamColumns[i], (uint)182);
                DocumentFormat.OpenXml.Spreadsheet.Cell renewableCell = this.GetSpreadsheetCell(worksheetPart.Worksheet, streamColumns[i], (uint)183);
                if (streamIdCell.CellValue != null)
                {
                    streamList.Add(Convert.ToInt32(streamIdCell.CellValue.Text));
                    if (productOrWasteCell.CellValue != null)
                    {
                        if (Convert.ToInt32(productOrWasteCell.CellValue.Text) == 0)
                            productOrWasteList.Add("Waste");
                        if (Convert.ToInt32(productOrWasteCell.CellValue.Text) == 1)
                            productOrWasteList.Add("Product");
                    }
                    else productOrWasteList.Add("N/A");
                    if (ecoProductCell.CellValue != null)
                    {
                        if (Convert.ToInt32(ecoProductCell.CellValue.Text) == 0)
                            ecoProductList.Add("No");
                        if (Convert.ToInt32(ecoProductCell.CellValue.Text) == 1)
                            ecoProductList.Add("Yes");
                        //else ecoProductList.Add("N/A");
                    }
                    else ecoProductList.Add("N/A");
                    if (pollutedORNotPollutedCell.CellValue != null)
                    {
                        if (Convert.ToInt32(pollutedORNotPollutedCell.CellValue.Text) == 0)
                            pollutedOrNonpollutedProductList.Add("Polluted");
                        if (Convert.ToInt32(pollutedORNotPollutedCell.CellValue.Text) == 1)
                            pollutedOrNonpollutedProductList.Add("NonPolluted");
                        //else pollutedOrNonpollutedProductList.Add("N/A");
                    }
                    else pollutedOrNonpollutedProductList.Add("N/A");
                    if (renewableCell.CellValue != null)
                    {
                        renewableList.Add(renewableCell.CellValue.Text);
                    }
                    else renewableList.Add("no");
                }
            }
            streams = streamList.ToArray<int>();
            productOrWastes = productOrWasteList.ToArray<string>();
            ecoProducts = ecoProductList.ToArray<string>();
            pollutedNotPolluted = pollutedOrNonpollutedProductList.ToArray<string>();
            renewables = renewableList.ToArray<string>();
        }

        private void GetReactionInformationFromSpreadsheet(DocumentFormat.OpenXml.Packaging.SpreadsheetDocument spreadsheet, ref int mainReaction, ref int mainProduct, ref int mainProductStream, ref double[] reactionStoich)
        {
            IEnumerable<DocumentFormat.OpenXml.Spreadsheet.Sheet> sheets = spreadsheet.WorkbookPart.Workbook.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.Sheets>().Elements<DocumentFormat.OpenXml.Spreadsheet.Sheet>().Where(s => s.Name == "I. Stream & Compound Data");
            if (sheets.Count() == 0)
            {
                // The specified worksheet does not exist.
                return;
            }
            string relationshipId = sheets.First().Id.Value;
            DocumentFormat.OpenXml.Packaging.WorksheetPart worksheetPart = (DocumentFormat.OpenXml.Packaging.WorksheetPart)spreadsheet.WorkbookPart.GetPartById(relationshipId);
            DocumentFormat.OpenXml.Spreadsheet.Cell tempCell = this.GetSpreadsheetCell(worksheetPart.Worksheet, "G", (uint)104);
            if (tempCell.CellValue != null)
            {
                mainReaction = Convert.ToInt32(tempCell.CellValue.Text);
            }
            else mainReaction = 0;
            tempCell = this.GetSpreadsheetCell(worksheetPart.Worksheet, "G", (uint)105);
            if (tempCell.CellValue != null)
            {
                mainProduct = Convert.ToInt32(tempCell.CellValue.Text);
            }
            else mainProduct = 0;
            tempCell = this.GetSpreadsheetCell(worksheetPart.Worksheet, "G", (uint)106);
            if (tempCell.CellValue != null)
            {
                mainProductStream = Convert.ToInt32(tempCell.CellValue.Text);
            }
            mainProductStream = 0;
            string[] stoichColumns = { "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z", "AA" };
            List<double> stoicCoeff = new List<double>();
            for (int i = 0; i < stoichColumns.Length; i++)
            {
                for (uint j = 0; j < 50; j++)
                {
                    DocumentFormat.OpenXml.Spreadsheet.Cell reacStoicCell = this.GetSpreadsheetCell(worksheetPart.Worksheet, stoichColumns[i], j + 111);
                    if (reacStoicCell.CellValue != null)
                    {
                        stoicCoeff.Add(Convert.ToDouble(reacStoicCell.CellValue.Text));
                    }
                }
            }
            reactionStoich = stoicCoeff.ToArray<double>();
        }

        private void AddInputStreamsToSpreadsheet(DocumentFormat.OpenXml.Packaging.SpreadsheetDocument spreadsheet, Stream[] streams)
        {
            IEnumerable<DocumentFormat.OpenXml.Spreadsheet.Sheet> sheets = spreadsheet.WorkbookPart.Workbook.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.Sheets>().Elements<DocumentFormat.OpenXml.Spreadsheet.Sheet>().Where(s => s.Name == "I. Stream & Compound Data");
            if (sheets.Count() == 0)
            {
                // The specified worksheet does not exist.
                return;
            }
            string relationshipId = sheets.First().Id.Value;
            DocumentFormat.OpenXml.Packaging.WorksheetPart worksheetPart = (DocumentFormat.OpenXml.Packaging.WorksheetPart)spreadsheet.WorkbookPart.GetPartById(relationshipId);
            string[] componentColumns = { "H", "K", "N", "Q", "T", "W", "Z", "AC", "AF", "AI", "AL", "AO", "AR", "AU", "AX", "BA", "BD", "BG", "BJ", "BM", "BP", "BS", "BV", "BY", "CB", "CE", "CH", "CK", "CN", "CQ", "CT", "CW", "CZ", "DC", "DF", "DI", "DL", "DO", "DR", "DU" };
            string[] streamColumns = { "I", "L", "O", "R", "U", "X", "AA", "AD", "AG", "AJ", "AM", "AP", "AS", "AV", "AY", "BB", "BE", "BH", "BK", "BN", "BQ", "BT", "BW", "BZ", "CC", "CF", "CI", "CL", "CO", "CR", "CU", "CX", "DA", "DD", "DG", "DJ", "DM", "DP", "DS", "DV" };
            for (uint i = 0; i < streams.Length; i++)
            {
                SetSpreadsheetCellValue(worksheetPart.Worksheet, streamColumns[i], 28, streams[i].StreamID);
                SetSpreadsheetCellValue(worksheetPart.Worksheet, streamColumns[i], 29, streams[i].StreamName);
                for (int j = 0; j < streams[i].ComponentMassFlowRatesKGH.Length; j++)
                {
                    SetSpreadsheetCellValue(worksheetPart.Worksheet, componentColumns[i], j + 32, streams[i].ComponentMassFlowRatesKGH[j]);
                }
                SetSpreadsheetCellValue(worksheetPart.Worksheet, componentColumns[i], 83, streams[i].TemperatureC);
                SetSpreadsheetCellValue(worksheetPart.Worksheet, componentColumns[i], 84, streams[i].PressureKPa);
                SetSpreadsheetCellValue(worksheetPart.Worksheet, componentColumns[i], 85, streams[i].MoleVaporFraction);
                SetSpreadsheetCellValue(worksheetPart.Worksheet, componentColumns[i], 86, streams[i].EnthalpyMJHR);
                SetSpreadsheetCellValue(worksheetPart.Worksheet, componentColumns[i], 87, streams[i].EntropyMJKHR);
                SetSpreadsheetCellValue(worksheetPart.Worksheet, componentColumns[i], 91, streams[i].Cost);
            }
        }
        private void AddReactionsToSpreadsheet(DocumentFormat.OpenXml.Packaging.SpreadsheetDocument spreadsheet, List<UnitOperation> reactors, Stream stream)
        {
            IEnumerable<DocumentFormat.OpenXml.Spreadsheet.Sheet> sheets = spreadsheet.WorkbookPart.Workbook.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.Sheets>().Elements<DocumentFormat.OpenXml.Spreadsheet.Sheet>().Where(s => s.Name == "I. Stream & Compound Data");
            if (sheets.Count() == 0)
            {
                // The specified worksheet does not exist.
                return;
            }
            string relationshipId = sheets.First().Id.Value;
            DocumentFormat.OpenXml.Packaging.WorksheetPart worksheetPart = (DocumentFormat.OpenXml.Packaging.WorksheetPart)spreadsheet.WorkbookPart.GetPartById(relationshipId);
            string[] reactionColumns = { "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z", "AA"};
            for (int i = 0; i < reactors.Count; i++)
            {
                int startCell = 111;
                for (int j = 0; j < stream.NumberOfComponents; j++)
                {
                    SetSpreadsheetCellValue(worksheetPart.Worksheet, reactionColumns[i], startCell + j, reactors[i].ReactionStoicCoeff(j));
                }
            }
        }

        private void AddOutputStreamsToSpreadsheet(DocumentFormat.OpenXml.Packaging.SpreadsheetDocument spreadsheet, Stream[] streams)
        {
            IEnumerable<DocumentFormat.OpenXml.Spreadsheet.Sheet> sheets = spreadsheet.WorkbookPart.Workbook.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.Sheets>().Elements<DocumentFormat.OpenXml.Spreadsheet.Sheet>().Where(s => s.Name == "I. Stream & Compound Data");
            if (sheets.Count() == 0)
            {
                // The specified worksheet does not exist.
                return;
            }
            string relationshipId = sheets.First().Id.Value;
            DocumentFormat.OpenXml.Packaging.WorksheetPart worksheetPart = (DocumentFormat.OpenXml.Packaging.WorksheetPart)spreadsheet.WorkbookPart.GetPartById(relationshipId);
            string[] componentColumns = { "H", "K", "N", "Q", "T", "W", "Z", "AC", "AF", "AI", "AL", "AO", "AR", "AU", "AX", "BA", "BD", "BG", "BJ", "BM", "BP", "BS", "BV", "BY", "CB", "CE", "CH", "CK", "CN", "CQ", "CT", "CW", "CZ", "DC", "DF", "DI", "DL", "DO", "DR", "DU" };
            string[] streamColumns = { "I", "L", "O", "R", "U", "X", "AA", "AD", "AG", "AJ", "AM", "AP", "AS", "AV", "AY", "BB", "BE", "BH", "BK", "BN", "BQ", "BT", "BW", "BZ", "CC", "CF", "CI", "CL", "CO", "CR", "CU", "CX", "DA", "DD", "DG", "DJ", "DM", "DP", "DS", "DV" };
            for (uint i = 0; i < streams.Length; i++)
            {
                SetSpreadsheetCellValue(worksheetPart.Worksheet, streamColumns[i], 178, streams[i].StreamID);
                SetSpreadsheetCellValue(worksheetPart.Worksheet, streamColumns[i], 180, streams[i].StreamName);
                for (int j = 0; j < streams[i].ComponentMassFlowRatesKGH.Length; j++)
                {
                    SetSpreadsheetCellValue(worksheetPart.Worksheet, componentColumns[i], j + 186, streams[i].ComponentMassFlowRatesKGH[j]);
                }
                SetSpreadsheetCellValue(worksheetPart.Worksheet, componentColumns[i], 237, streams[i].TemperatureC);
                SetSpreadsheetCellValue(worksheetPart.Worksheet, componentColumns[i], 238, streams[i].PressureKPa);
                SetSpreadsheetCellValue(worksheetPart.Worksheet, componentColumns[i], 239, streams[i].MoleVaporFraction);
                SetSpreadsheetCellValue(worksheetPart.Worksheet, componentColumns[i], 240, streams[i].EnthalpyMJHR);
                SetSpreadsheetCellValue(worksheetPart.Worksheet, componentColumns[i], 241, streams[i].EntropyMJKHR);
                SetSpreadsheetCellValue(worksheetPart.Worksheet, componentColumns[i], 243, streams[i].LiquidVolumetricFlowRateM3HR);
                SetSpreadsheetCellValue(worksheetPart.Worksheet, componentColumns[i], 246, streams[i].Cost);
            }
        }


        //private void ClearComponentSpreadsheet(DocumentFormat.OpenXml.Packaging.SpreadsheetDocument spreadsheet)
        //{
        //    IEnumerable<DocumentFormat.OpenXml.Spreadsheet.Sheet> sheets = spreadsheet.WorkbookPart.Workbook.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.Sheets>().Elements<DocumentFormat.OpenXml.Spreadsheet.Sheet>().Where(s => s.Name == "I. Stream & Compound Data");
        //    if (sheets.Count() == 0)
        //    {
        //        // The specified worksheet does not exist.
        //        return;
        //    }
        //    string relationshipId = sheets.First().Id.Value;
        //    DocumentFormat.OpenXml.Packaging.WorksheetPart worksheetPart = (DocumentFormat.OpenXml.Packaging.WorksheetPart)spreadsheet.WorkbookPart.GetPartById(relationshipId);

        //    string[] compoundColumns = { "D", "E", "G" };
        //    string[] componentColumns = { "H", "K", "N", "Q", "T", "W", "Z", "AC", "AF", "AI", "AL", "AO", "AR", "AU", "AX", "BA", "BD", "BG", "BJ", "BM", "BP", "BS", "BV", "BY", "CB", "CE", "CH", "CK", "CN", "CQ", "CT", "CW", "CZ", "DC", "DF", "DI", "DL", "DO", "DR", "DU" };
        //    string[] streamColumns = { "I", "L", "O", "R", "U", "X", "AA", "AD", "AG", "AJ", "AM", "AP", "AS", "AV", "AY", "BB", "BE", "BH", "BK", "BN", "BQ", "BT", "BW", "BZ", "CC", "CF", "CI", "CL", "CO", "CR", "CU", "CX", "DA", "DD", "DG", "DJ", "DM", "DP", "DS", "DV" };
        //    string[] reactionColumns = { "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z", "AA" };
        //    string[] propertyColumns = { "K", "R", "S", "T", "U", "V", "X", "Y", "Z", "AA", "AB", "AC", "AD", "AE", "AF", "AG", "AI", "AK", "AM", "AN", "AO", "AP", "AR", "AT", "AV", "AW", "AX", "AY", "AZ", "BA", "BB", "BC", "BD", "BE", "BF", "BG", "BH", "BI", "BJ", "BK", "BL", "BM", "BN", "BO", "BQ", "BS", "BU", "BW", "BY", "CA", "CB", "CD", "CF", "CH", "CI", "CJ", "CK", "CL", "CM", "CN", "CO", "CQ", "CS" };
        //    foreach (string colName in compoundColumns)
        //    {
        //        for (int rowIndex = 32; rowIndex < 82; rowIndex++)
        //        {
        //            SetSpreadsheetCellValue(worksheetPart.Worksheet, colName, rowIndex, String.Empty);
        //        }
        //    }
        //    for (int rowIndex = 32; rowIndex < 82; rowIndex++)
        //    {
        //        SetSpreadsheetCellValue(worksheetPart.Worksheet, "F", rowIndex, -1);
        //    }
        //    foreach (string colName in componentColumns)
        //    {
        //        for (int rowIndex = 32; rowIndex < 82; rowIndex++)
        //        {
        //            SetSpreadsheetCellValue(worksheetPart.Worksheet, colName, rowIndex, String.Empty);
        //        }
        //        for (int rowIndex = 83; rowIndex < 93; rowIndex++)
        //        {
        //            SetSpreadsheetCellValue(worksheetPart.Worksheet, colName, rowIndex, String.Empty);
        //        }
        //        for (int rowIndex = 186; rowIndex < 236; rowIndex++)
        //        {
        //            SetSpreadsheetCellValue(worksheetPart.Worksheet, colName, rowIndex, String.Empty);
        //        }
        //        for (int rowIndex = 237; rowIndex < 248; rowIndex++)
        //        {
        //            SetSpreadsheetCellValue(worksheetPart.Worksheet, colName, rowIndex, String.Empty);
        //        }
        //    }
        //    foreach (string colName in streamColumns)
        //    {
        //        for (int rowIndex = 27; rowIndex < 31; rowIndex++)
        //        {
        //            SetSpreadsheetCellValue(worksheetPart.Worksheet, colName, rowIndex, String.Empty);
        //        }
        //        for (int rowIndex = 178; rowIndex < 185; rowIndex++)
        //        {
        //            SetSpreadsheetCellValue(worksheetPart.Worksheet, colName, rowIndex, String.Empty);
        //        }
        //    }
        //    foreach (string colName in reactionColumns)
        //    {
        //        SetSpreadsheetCellValue(worksheetPart.Worksheet, colName, 108, String.Empty);
        //        SetSpreadsheetCellValue(worksheetPart.Worksheet, colName, 109, String.Empty);
        //        for (int rowIndex = 111; rowIndex < 162; rowIndex++)
        //        {
        //            SetSpreadsheetCellValue(worksheetPart.Worksheet, colName, rowIndex, String.Empty);
        //        }
        //        for (int rowIndex = 163; rowIndex < 167; rowIndex++)
        //        {
        //            SetSpreadsheetCellValue(worksheetPart.Worksheet, colName, rowIndex, String.Empty);
        //        }
        //    }
        //    foreach (string colName in propertyColumns)
        //    {
        //        for (int rowIndex = 262; rowIndex < 312; rowIndex++)
        //        {
        //            SetSpreadsheetCellValue(worksheetPart.Worksheet, colName, rowIndex, String.Empty);
        //        }
        //    }
        //    worksheetPart.Worksheet.Save();
        //}

        //private void ClearUnitOpSpreadsheet(DocumentFormat.OpenXml.Packaging.SpreadsheetDocument spreadsheet)
        //{
        //    IEnumerable<DocumentFormat.OpenXml.Spreadsheet.Sheet> sheets = spreadsheet.WorkbookPart.Workbook.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.Sheets>().Elements<DocumentFormat.OpenXml.Spreadsheet.Sheet>().Where(s => s.Name == "II. Equipment & Cost Data");
        //    if (sheets.Count() == 0)
        //    {
        //        // The specified worksheet does not exist.
        //        return;
        //    }
        //    string relationshipId = sheets.First().Id.Value;
        //    DocumentFormat.OpenXml.Packaging.WorksheetPart worksheetPart = (DocumentFormat.OpenXml.Packaging.WorksheetPart)spreadsheet.WorkbookPart.GetPartById(relationshipId);

        //    string[] unitOpColumns = { "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "W", "X" };
        //    uint[] mixerRows = { 10, 11, 12, 13, 14, 15, 16, 17, 18 };
        //    uint[] pumpRows = { 26, 26, 27, 28, 29, 30, 31, 32, 33 };
        //    uint[] distillationRows = { 40, 41, 42, 43, 44, 45, 46, 47, 48, 49 };
        //    uint[] heatExchangerRows = { 56, 57, 58, 59, 60, 61, 62, 63, 64 };
        //    uint[] extractorRows = { 71, 72, 73, 74, 75, 76, 77, 78, 79 };
        //    uint[] componentSeparatorRows = { 86, 87, 88, 89, 90, 91, 92, 93, 94 };
        //    uint[] reactorRows = { 101, 102, 103, 104, 105, 106, 107, 108, 109, 110 };
        //    uint[] otherEquipemtRows = { 117, 118, 119, 120, 121, 122, 123, 124, 125 };
        //    foreach (string colName in unitOpColumns)
        //    {
        //        foreach (int rowIndex in mixerRows)
        //        {
        //            SetSpreadsheetCellValue(worksheetPart.Worksheet, colName, rowIndex, String.Empty);
        //        }
        //        foreach (int rowIndex in pumpRows)
        //        {
        //            SetSpreadsheetCellValue(worksheetPart.Worksheet, colName, rowIndex, String.Empty);
        //        }
        //        foreach (int rowIndex in distillationRows)
        //        {
        //            SetSpreadsheetCellValue(worksheetPart.Worksheet, colName, rowIndex, String.Empty);
        //        }
        //        foreach (int rowIndex in heatExchangerRows)
        //        {
        //            SetSpreadsheetCellValue(worksheetPart.Worksheet, colName, rowIndex, String.Empty);
        //        }
        //        foreach (int rowIndex in extractorRows)
        //        {
        //            SetSpreadsheetCellValue(worksheetPart.Worksheet, colName, rowIndex, String.Empty);
        //        }
        //        foreach (int rowIndex in componentSeparatorRows)
        //        {
        //            SetSpreadsheetCellValue(worksheetPart.Worksheet, colName, rowIndex, String.Empty);
        //        }
        //        foreach (int rowIndex in reactorRows)
        //        {
        //            SetSpreadsheetCellValue(worksheetPart.Worksheet, colName, rowIndex, String.Empty);
        //        }
        //        foreach (int rowIndex in otherEquipemtRows)
        //        {
        //            SetSpreadsheetCellValue(worksheetPart.Worksheet, colName, rowIndex, String.Empty);
        //        }
        //    }
        //    worksheetPart.Worksheet.Save();
        //}

        private void AddUnitOpsToSpreadsheet(DocumentFormat.OpenXml.Packaging.SpreadsheetDocument spreadsheet, UnitOperation[] unitOps, int[] unitOpRows)
        {
            IEnumerable<DocumentFormat.OpenXml.Spreadsheet.Sheet> sheets = spreadsheet.WorkbookPart.Workbook.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.Sheets>().Elements<DocumentFormat.OpenXml.Spreadsheet.Sheet>().Where(s => s.Name == "II. Equipment & Cost Data");
            if (sheets.Count() == 0)
            {
                // The specified worksheet does not exist.
                return;
            }
            string relationshipId = sheets.First().Id.Value;
            DocumentFormat.OpenXml.Packaging.WorksheetPart worksheetPart = (DocumentFormat.OpenXml.Packaging.WorksheetPart)spreadsheet.WorkbookPart.GetPartById(relationshipId);

            string[] unitOpColumns = { "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "W", "X" };

            for (int i = 0; i < unitOps.Length; i++)
            {

                SetSpreadsheetCellValue(worksheetPart.Worksheet, unitOpColumns[i], unitOpRows[1], unitOps[i].UnitOpId.ToString());
                SetSpreadsheetCellValue(worksheetPart.Worksheet, unitOpColumns[i], unitOpRows[2], unitOps[i].Label);
                SetSpreadsheetCellValue(worksheetPart.Worksheet, unitOpColumns[i], unitOpRows[3], String.Empty);
                SetSpreadsheetCellValue(worksheetPart.Worksheet, unitOpColumns[i], unitOpRows[4], unitOps[i].HeatAdded);
                SetSpreadsheetCellValue(worksheetPart.Worksheet, unitOpColumns[i], unitOpRows[5], unitOps[i].Power);
                SetSpreadsheetCellValue(worksheetPart.Worksheet, unitOpColumns[i], unitOpRows[6], unitOps[i].TotalPurchaseCost);
                SetSpreadsheetCellValue(worksheetPart.Worksheet, unitOpColumns[i], unitOpRows[7], unitOps[i].TotalInstalledCost);
                SetSpreadsheetCellValue(worksheetPart.Worksheet, unitOpColumns[i], unitOpRows[8], String.Empty);
            }
            worksheetPart.Worksheet.Save();
        }

        private void AddMixerUnitOpsToSpreadsheet(DocumentFormat.OpenXml.Packaging.SpreadsheetDocument spreadsheet, UnitOperation[] mixers)
        {
            int[] mixerRows = { 10, 11, 12, 13, 14, 15, 16, 17, 18 };
            AddUnitOpsToSpreadsheet(spreadsheet, mixers, mixerRows);
        }

        private void AddMPumpUnitOpsToSpreadsheet(DocumentFormat.OpenXml.Packaging.SpreadsheetDocument spreadsheet, UnitOperation[] pumps)
        {
            int[] pumpRows = { 26, 26, 27, 28, 29, 30, 31, 32, 33 };
            AddUnitOpsToSpreadsheet(spreadsheet, pumps, pumpRows);
        }

        private void AddDistillationUnitOpsToSpreadsheet(DocumentFormat.OpenXml.Packaging.SpreadsheetDocument spreadsheet, UnitOperation[] distillationColumns)
        {
            int[] distillationRows = { 40, 41, 42, 43, 44, 45, 46, 47, 48, 49 };
            AddUnitOpsToSpreadsheet(spreadsheet, distillationColumns, distillationRows);
        }

        private void AddHeatExchangerUnitOpsToSpreadsheet(DocumentFormat.OpenXml.Packaging.SpreadsheetDocument spreadsheet, UnitOperation[] heatExchangers)
        {
            int[] heatExchangerRows = { 56, 57, 58, 59, 60, 61, 62, 63, 64 };
            AddUnitOpsToSpreadsheet(spreadsheet, heatExchangers, heatExchangerRows);
        }

        private void AddExtractorUnitOpsToSpreadsheet(DocumentFormat.OpenXml.Packaging.SpreadsheetDocument spreadsheet, UnitOperation[] extractors)
        {
            int[] extractorRows = { 71, 72, 73, 74, 75, 76, 77, 78, 79 };
            AddUnitOpsToSpreadsheet(spreadsheet, extractors, extractorRows);
        }

        private void AddReactorUnitOpsToSpreadsheet(DocumentFormat.OpenXml.Packaging.SpreadsheetDocument spreadsheet, UnitOperation[] reactors)
        {
            int[] reactorRows = { 101, 102, 103, 104, 105, 106, 107, 108, 109, 110 };
            AddUnitOpsToSpreadsheet(spreadsheet, reactors, reactorRows);
        }

        private void AddComponentSeparatorUnitOpsToSpreadsheet(DocumentFormat.OpenXml.Packaging.SpreadsheetDocument spreadsheet, UnitOperation[] componentSeparators)
        {
            int[] componentSeparatorRows = { 86, 87, 88, 89, 90, 91, 92, 93, 94 };
            AddUnitOpsToSpreadsheet(spreadsheet, componentSeparators, componentSeparatorRows);
        }

        private void AddOtherUnitOpsToSpreadsheet(DocumentFormat.OpenXml.Packaging.SpreadsheetDocument spreadsheet, UnitOperation[] otherUnitOps)
        {
            int[] otherEquipemtRows = { 117, 118, 119, 120, 121, 122, 123, 124, 125 };
            AddUnitOpsToSpreadsheet(spreadsheet, otherUnitOps, otherEquipemtRows);
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            ComboBox comboBox = (ComboBox)sender;
            string tempUnit = (string)comboBox.SelectedItem;
            if (processReferenceTemperatureUnit == "Celsius")
            {
                if (tempUnit == "Celsius") { }
                else if (tempUnit == "Kelvin") processReferenceTemperature = processReferenceTemperature + 273.15;
                else if (tempUnit == "Rankine") processReferenceTemperature = (processReferenceTemperature + 273.15) * 9 / 5;
                else processReferenceTemperature = processReferenceTemperature * 9 / 5 + 32; // Temperature in Farenheit.
            }
            else if (processReferenceTemperatureUnit == "Kelvin")
            {
                if (tempUnit == "Celsius") processReferenceTemperature = processReferenceTemperature - 273.15;
                else if (tempUnit == "Kelvin") { }
                else if (tempUnit == "Rankine") processReferenceTemperature = processReferenceTemperature * 9 / 5;
                else processReferenceTemperature = (processReferenceTemperature * 9 / 5) - 491.67; // Temperature in Farenheit.
            }
            else if (processReferenceTemperatureUnit == "Rankine")
            {
                if (tempUnit == "Celsius") processReferenceTemperature = processReferenceTemperature * 5 / 9 - 273.15;
                else if (tempUnit == "Kelvin") processReferenceTemperature = processReferenceTemperature * 5 / 9;
                else if (tempUnit == "Rankine") { }
                else processReferenceTemperature = processReferenceTemperature - 491.67; // Temperature in Farenheit.
            }
            else
            {
                if (tempUnit == "Celsius") processReferenceTemperature = (processReferenceTemperature - 32) * 5 / 9;
                else if (tempUnit == "Kelvin") processReferenceTemperature = (processReferenceTemperature - 32) * 5 / 9 + 273.15;
                else if (tempUnit == "Rankine") processReferenceTemperature = processReferenceTemperature + 491.67; // Temperature in Farenheit.
                else { } // Temperature in Farenheit.
            }
            processReferenceTemperatureUnit = tempUnit;
            this.numericUpDown1.Value = (decimal)processReferenceTemperature;
        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            ComboBox comboBox = (ComboBox)sender;
            string tempUnit = (string)comboBox.SelectedItem;
            if (referenceTemperatureUnit == "Celsius")
            {
                if (tempUnit == "Celsius") { }
                else if (tempUnit == "Kelvin") referenceTemperature = referenceTemperature + 273.15;
                else if (tempUnit == "Rankine") referenceTemperature = (referenceTemperature + 273.15) * 9 / 5;
                else referenceTemperature = referenceTemperature * 9 / 5 + 32; // Temperature in Farenheit.
            }
            else if (referenceTemperatureUnit == "Kelvin")
            {
                if (tempUnit == "Celsius") referenceTemperature = referenceTemperature - 273.15;
                else if (tempUnit == "Kelvin") { }
                else if (tempUnit == "Rankine") referenceTemperature = referenceTemperature * 9 / 5;
                else referenceTemperature = (referenceTemperature * 9 / 5) - 491.67; // Temperature in Farenheit.
            }
            else if (referenceTemperatureUnit == "Rankine")
            {
                if (tempUnit == "Celsius") referenceTemperature = referenceTemperature * 5 / 9 - 273.15;
                else if (tempUnit == "Kelvin") referenceTemperature = referenceTemperature * 5 / 9;
                else if (tempUnit == "Rankine") { }
                else referenceTemperature = referenceTemperature - 459.67; // Temperature in Farenheit.
            }
            else
            {
                if (tempUnit == "Celsius") referenceTemperature = (referenceTemperature - 32) * 5 / 9;
                else if (tempUnit == "Kelvin") referenceTemperature = (referenceTemperature - 32) * 5 / 9 + 273.15;
                else if (tempUnit == "Rankine") referenceTemperature = referenceTemperature + 459.67; // Temperature in Farenheit.
                else { } // Temperature in Farenheit.
            }
            referenceTemperatureUnit = tempUnit;
            this.numericUpDown2.Value = (decimal)referenceTemperature;
        }

        private void comboBox3_SelectedIndexChanged(object sender, EventArgs e)
        {
            ComboBox comboBox = (ComboBox)sender;
            string pressUnit = (string)comboBox.SelectedItem;
            double newRefPress = 0;

            // convert to kPa
            if (referencePressureUnit == "atm") newRefPress = referencePressure * 1.01325e+02;
            if (referencePressureUnit == "psia") newRefPress = referencePressure * 6.89476;
            if (referencePressureUnit == "psig") newRefPress = referencePressure * 6.89476 + 1.01325e+02;
            if (referencePressureUnit == "torr") newRefPress = referencePressure * 1.33322e-01;
            if (referencePressureUnit == "mmHg") newRefPress = referencePressure * 1.33322e-01;
            if (referencePressureUnit == "Pa") newRefPress = referencePressure / 1000;
            if (referencePressureUnit == "kPa") newRefPress = referencePressure;
            if (referencePressureUnit == "MPa G") newRefPress = referencePressure * 1000 + 1.01325e+02;
            if (referencePressureUnit == "bar") newRefPress = referencePressure * 1e+02;
            if (referencePressureUnit == "bar G") newRefPress = referencePressure * 1e+02 + 1.01325e+02;
            if (referencePressureUnit == "mbar") newRefPress = referencePressure * 0.1;
            if (referencePressureUnit == "kg/cm2") newRefPress = referencePressure * 98.0665;
            if (referencePressureUnit == "kg/cm2 G") newRefPress = referencePressure * 98.0665 + 1.01325e+02;
            if (referencePressureUnit == "in-water") newRefPress = referencePressure * 2.49089e-01;
            if (referencePressureUnit == "mm-water") newRefPress = referencePressure * 9.80665e-03;

            // convert to desired unit
            if (pressUnit == "atm") referencePressure = newRefPress / 1.01325e+02;
            if (pressUnit == "psia") referencePressure = newRefPress / 6.89476;
            if (pressUnit == "psig") referencePressure = (newRefPress - 1.01325e+02) / 6.89476;
            if (pressUnit == "torr") referencePressure = newRefPress / 1.33322e-01;
            if (pressUnit == "mmHg") referencePressure = newRefPress / 1.33322e-01;
            if (pressUnit == "Pa") referencePressure = newRefPress * 1000;
            if (pressUnit == "kPa") referencePressure = newRefPress;
            if (pressUnit == "MPa G") referencePressure = (newRefPress - 1.01325e+02) / 1000;
            if (pressUnit == "bar") referencePressure = newRefPress / 1e+02;
            if (pressUnit == "bar G") referencePressure = (newRefPress - 1.01325e+02) / 1e+02;
            if (pressUnit == "mbar") referencePressure = newRefPress / 0.1;
            if (pressUnit == "kg/cm2") referencePressure = newRefPress / 98.0665;
            if (pressUnit == "kg/cm2 G") referencePressure = (newRefPress - 1.01325e+02) / 98.0665;
            if (pressUnit == "in-water") referencePressure = newRefPress / 2.49089e-01;
            if (pressUnit == "mm-water") referencePressure = newRefPress / 9.80665e-03;
            referencePressureUnit = pressUnit;
            this.numericUpDown3.Value = (decimal)referencePressure;
        }

        private void numericUpDown1_ValueChanged(object sender, EventArgs e)
        {
            processReferenceTemperature = (double)this.numericUpDown1.Value;
        }

        private void numericUpDown2_ValueChanged(object sender, EventArgs e)
        {
            referenceTemperature = (double)this.numericUpDown2.Value;
        }

        private void numericUpDown3_ValueChanged(object sender, EventArgs e)
        {
            referencePressure = (double)this.numericUpDown3.Value;
        }


    }
}
