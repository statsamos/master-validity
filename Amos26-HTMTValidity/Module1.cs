using System;
using System.Collections;
using System.Collections.Generic;
#region Imports
using System.Diagnostics;
using System.Xml;
using Amos;
using Amos26_HTMTValidity;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;


namespace Amos26_HTMTValidity
{

    // This plugin was written July 2020 by Matthew James for James Gaskin
    // Updated July 2020
    // This plugin was updated 2022 by Joseph Steed

    [System.ComponentModel.Composition.Export(typeof(IPlugin))]
    public class CustomCode : IPlugin
    {

        private ArrayList HigherOrderVariables = new ArrayList();

        public string Name()
        {
            return "Validity and Reliability Test w/ HTMT Analysis";
        }

        public string Description()
        {
            return "Calculates CR, AVE, ASV, MSV and outputs them in a table with the correlations. Includes interpretations from the results. See statwiki.kolobkreations.com for more information.";
        }

        // Struct to hold the estimates for a latent variable that you are testing.
        public struct Estimates
        {

            public double CR;
            public double AVE;
            public double LCR;
            public double UCR;
            public double LAVE;
            public double UAVE;
            public double MSV;
            public double MaxR;
            public double SQRT;

        }

        public int MainSub()
        {

            // Check if the model is a causal rather than measurment model by looping through paths.
            foreach (PDElement variable in pd.PDElements)
            {
                if (variable.IsPath())
                {
                    // Model contains paths from latent to latent or observed to observed.
                    if (variable.Variable1.IsObservedVariable() & variable.Variable2.IsObservedVariable() | variable.Variable1.IsObservedVariable() & variable.Variable2.IsLatentVariable())
                    {
                        Interaction.MsgBox("The master validity plugin cannot be used for causal models.");
                        return default;
                    }
                    if (variable.Variable1.IsLatentVariable() & variable.Variable2.IsLatentVariable())
                    {
                        foreach (PDElement subvariable in pd.PDElements)
                        {
                            if (subvariable.IsPath())
                            {
                                if ((variable.Variable1.NameOrCaption ?? "") == (subvariable.Variable1.NameOrCaption ?? "") & subvariable.Variable2.IsObservedVariable())
                                {
                                    Interaction.MsgBox("The master validity plugin cannot be used for causal models.");
                                    return default;
                                }
                            }
                        }
                    }
                }
            }

            // Ensure that the output includes standardized estimates.
            pd.GetCheckBox("AnalysisPropertiesForm", "StandardizedCheck").Checked = true;
            pd.GetCheckBox("AnalysisPropertiesForm", "SampleMomCheck").Checked = true;
            pd.GetCheckBox("AnalysisPropertiesForm", "MeansInterceptsCheck").Checked = true;
            pd.GetCheckBox("AnalysisPropertiesForm", "CorestCheck").Checked = true;
            pd.GetCheckBox("AnalysisPropertiesForm", "CovestCheck").Checked = true;

            // Enable Boostrap
            pd.GetCheckBox("AnalysisPropertiesForm", "DoBootstrapCheck").Checked = true;
            pd.GetCheckBox("AnalysisPropertiesForm", "ConfidencePCCheck").Checked = true;
            pd.GetTextBox("AnalysisPropertiesForm", "BootstrapText").Text = "500";
            pd.GetTextBox("AnalysisPropertiesForm", "ConfidencePCText").Text = "95";

            // Method to run the model.
            pd.AnalyzeCalculateEstimates();

            // If error, some variable in the model is unlinked or "floating." Ensure that all variables in model are linked. For example, "E1" (Error Term 1) is not linked to an observed variable.
            if (Util.GetNodeCount(Util.GetXML("body/div/div[@ntype='models']/div[@ntype='model'][position() = 1]/div[@ntype='group'][position() = 1]/div[@ntype='estimates']/div[@ntype='scalars']/div[@nodecaption='Standardized Regression Weights:']/table/tbody")) == 0)
            {
                Interaction.MsgBox("Invalid Model: Cannot Compute Standardized Regression Weights. All variables need to have at least one relationship to another variable.");
                return default;
            }

            // Grab Information Related To Higher Order Models
            HigherOrderVariables = Util.GetHigherOrderLatent();

            // Produces the matrix of estimates and correlations.
            var estimateMatrix = GetMatrix();
            // If you want to test explicit values with a dataset, you can change the values here, as shown below
            //estimateMatrix[1, 5] = 0.1d;
            //estimateMatrix[6, 10] = 0.1d;

            // Sub procedure to create the output file.
            CreateOutput(estimateMatrix);
            return default;

        }

        // Use the estimates matrix to create an html file of output.
        public void CreateOutput(double[,] estimateMatrix)
        {

            // Get the output tables and the count of correlations
            var tableCorrelation = Util.GetXML("body/div/div[@ntype='models']/div[@ntype='model'][position() = 1]/div[@ntype='group'][position() = 1]/div[@ntype='estimates']/div[@ntype='scalars']/div[@nodecaption='Correlations:']/table/tbody");
            var tableCovariance = Util.GetXML("body/div/div[@ntype='models']/div[@ntype='model'][position() = 1]/div[@ntype='group'][position() = 1]/div[@ntype='estimates']/div[@ntype='scalars']/div[@nodecaption='Covariances:']/table/tbody");
            int numCorrelation = Util.GetNodeCount(tableCorrelation);

            // Get a list of the latent variables
            var latentVariables = GetLatent();
            // Temp testing of anything in estimate matrix to force things to occur:


            // Check the matrix for validity concerns
            var validityConcerns = checkValidity(estimateMatrix);

            // Variables
            bool bMalhotra = true;
            int iMalhotra = 0;
            bool missingCorrelation = false;

            // Delete the output file if it exists
            if (System.IO.File.Exists("MasterValidity.html"))
            {
                System.IO.File.Delete("MasterValidity.html");
            }

            // Start the debugger for the html output
            var debug = new AmosDebug.AmosDebug();
            // File the output will be written to
            var resultWriter = new TextWriterTraceListener("MasterValidity.html");
            Trace.Listeners.Add(resultWriter);

            debug.PrintX("<html><body><h1>Model Validity Measures</h1><hr/><h3>Validity Analysis[edit-added]</h3>");

            // Optional debug before printing arrays, etc.
            // debug.PrintX("<p>")
            // debug.PrintX(estimateMatrix(0, 0))
            debug.PrintX(estimateMatrix[6, 10]);
            // debug.PrintX(estimateMatrix(0, 0))
            // debug.PrintX(estimateMatrix(0, 0))
            // debug.PrintX(estimateMatrix(0, 0))
            // debug.PrintX(estimateMatrix(0, 0))
            // debug.PrintX("</p>")
            // Start table of Validity Analysis
            debug.PrintX("<table><tr><td></td><th>CR</th><th>AVE</th>");

            // Handle zero correlations
            if (numCorrelation > 0)
            {
                debug.PrintX("<th>MSV</th><th>MaxR(H)</th>");
                foreach (var latent in latentVariables)
                    debug.PrintX("<th>" + latent + "[latent]</th>");
            }
            else
            {
                debug.PrintX("<th>MaxR(H)</th>");
            }
            debug.PrintX("</tr>");

            // Loop through the matrix to output variables, estimates, and correlations
            for (int y = 0, loopTo = latentVariables.Count - 1; y <= loopTo; y++)
            {
                debug.PrintX("<tr><td><span style='font-weight:bold;'>" + latentVariables[y] + "[ - y]</span></td>");
                for (int x = 0, loopTo1 = latentVariables.Count + 3; x <= loopTo1; x++)
                {
                    string significance = "";
                    switch (x)
                    {
                        case 0: // This is the first column, the CR column
                            {
                                if (estimateMatrix[y, x] < 0.7d | estimateMatrix[y, x] < estimateMatrix[y, x + 1])
                                {
                                    debug.PrintX("<td style='color:red'>");
                                    bMalhotra = false;
                                }
                                else
                                {
                                    debug.PrintX("<td>");
                                }

                                break;
                            }
                        case 1: // This is the second column, the AVE column
                            {
                                if (estimateMatrix[y, x] < 0.5d | estimateMatrix[y, x] < estimateMatrix[y, x + 1] & numCorrelation > 1)
                                {
                                    debug.PrintX("<td style='color:red'>");
                                    if (bMalhotra == true)
                                    {
                                        iMalhotra += 1;
                                    }
                                }
                                else
                                {
                                    debug.PrintX("<td>");

                                }

                                break;
                            }
                        case var @case when @case == y + 4: // These are all the cases for the latent variable columns, which line up with the x and y by being four apart.
                            {
                                // We are assuming there are no errors. 
                                bool red = false;
                                for (int i = 1, loopTo2 = numCorrelation; i <= loopTo2; i++)
                                {
                                    if (Conversions.ToBoolean(Operators.OrObject(Operators.ConditionalCompareObjectEqual(latentVariables[y], Util.MatrixName(tableCorrelation, i, 0L), false), Operators.ConditionalCompareObjectEqual(latentVariables[y], Util.MatrixName(tableCorrelation, i, 2L), false))))
                                    {
                                        if (Math.Sqrt(estimateMatrix[y, x]) < Util.MatrixElement(tableCorrelation, i, 3L))
                                        {
                                            red = true;
                                        }
                                    }
                                }
                                if (red == true & numCorrelation > 0)
                                {
                                    debug.PrintX("<td><span style='font-weight:bold; color:red;'>"); // This is the diagonal correlation!
                                }
                                else if (numCorrelation > 0)
                                {
                                    debug.PrintX("<td><span style='font-weight:bold; font-family:arial;'>debug3 x" + x.ToString() + " y" + y.ToString()); // This is the diagonal correlation!
                                }

                                break;
                            }
                        case 2: // This is the third column, MSV. There is no red checks on here, so nothing needed
                            {
                                debug.PrintX("<td>");
                                break;
                            }
                        case 3: // This is the fourth column, MaxR(H), there is no red checks here, so nothing needed
                            {
                                if (numCorrelation > 0)
                                {
                                    debug.PrintX("<td>");
                                }

                                break;
                            }

                        default:
                            {
                                // Interpretations of significance
                                if (numCorrelation > 0)
                                {
                                    debug.PrintX("<td>");
                                    for (int i = 1, loopTo3 = numCorrelation; i <= loopTo3; i++)
                                    {
                                        if (Conversions.ToBoolean(Operators.OrObject(Operators.AndObject(Operators.ConditionalCompareObjectEqual(Util.MatrixName(tableCovariance, i, 0L), latentVariables[x - 4], false), Operators.ConditionalCompareObjectEqual(Util.MatrixName(tableCovariance, i, 2L), latentVariables[y], false)), Operators.AndObject(Operators.ConditionalCompareObjectEqual(Util.MatrixName(tableCovariance, i, 0L), latentVariables[y], false), Operators.ConditionalCompareObjectEqual(Util.MatrixName(tableCovariance, i, 2L), latentVariables[x - 4], false)))))
                                        {
                                            if (Util.MatrixName(tableCovariance, i, 6L) == "***")
                                            {
                                                significance = "***";
                                            }
                                            else if (Util.MatrixElement(tableCovariance, i, 6L) < 0.01d)
                                            {
                                                significance = "**";
                                            }
                                            else if (Util.MatrixElement(tableCovariance, i, 6L) < 0.05d)
                                            {
                                                significance = "*";
                                            }
                                            else if (Util.MatrixElement(tableCovariance, i, 6L) < 0.1d)
                                            {
                                                significance = "&#8224;";
                                            }
                                            else
                                            {
                                                significance = "";
                                            }
                                        }
                                    }
                                }
                                break;
                            }
                    }

                    if (x <= 2)
                    {
                        debug.PrintX(estimateMatrix[y, x].ToString("#0.000") + "[x is ltet 2]</span></td>");
                    }
                    else if (x > 2 & estimateMatrix[y, x] != 0d & numCorrelation > 0 & y + 4 >= x)
                    {
                        debug.PrintX(estimateMatrix[y, x].ToString("#0.000") + significance + "[x is mt 2]</span></td>");
                    }
                    else if (x > 2 & numCorrelation > 0 & estimateMatrix[y, x] == 0d & y + 4 > x)
                    {
                        debug.PrintX("----</td>");
                        missingCorrelation = true;
                    }
                    else
                    {
                        debug.PrintX("</td>");
                    }
                }
                debug.PrintX("</tr>");
            }
            debug.PrintX("</table><br>");

            if (missingCorrelation == true)
            {
                debug.PrintX(" &#8258 Correlation is not specified in the model.<br>");
            }

            if (numCorrelation == 0)
            {
                debug.PrintX("You only had one latent variable so there is no correlation matrix or MSV.<br>");
            }
            else if (validityConcerns.Count == 0)
            {
                debug.PrintX("No validity concerns here.<br>");
            }

            // Print the validity concerns
            foreach (var message in validityConcerns)
                debug.PrintX(message+ "<br>");

            // Print confidence intervals
            debug.PrintX("<h3>Validity Analysis - Confidence Intervals</h3><table><tr><td></td><th>CR</th><th>AVE</th><th>Lower 95% CR</th><th>Upper 95% CR</th><th>Lower 95% AVE</th><th>Upper 95% AVE</th>");

            for (int y = 0, loopTo4 = latentVariables.Count - 1; y <= loopTo4; y++)
            {
                debug.PrintX("</tr><tr><td><span style='font-weight:bold;'>" + latentVariables[y] + "[debug1]</span></td>");
                for (int x = 0; x <= 5; x++)
                    debug.PrintX("<td>" + estimateMatrix[y + latentVariables.Count, x].ToString("#0.000") + "[debug2]</span></td>");
            }

            debug.PrintX("</tr></table><br>");

            // HTMT Analysis
            // Variables needed for output
            const double constLiberal = 0.9d;
            const double constStrict = 0.85d;
            var liberalWarnings = new ArrayList();
            var strictWarnings = new ArrayList();
            latentVariables = GetLatent();
            var htmtCalculations = GetHTMTCalculations();

            // Write the beginning of the document
            debug.PrintX("<h3>HTMT Analysis</h3><table><tr><th></th>");
            foreach (var latent in latentVariables) // Headers (latent variables) for table columns
                debug.PrintX("<th>" + latent + "</th>");
            debug.PrintX("</tr>");

            var lastvar = default(int);
            for (int i = 0, loopTo5 = latentVariables.Count - 1; i <= loopTo5; i++) // Headers (latent variable) for table rows
            {
                debug.PrintX("<tr><th>" + latentVariables[i] + "</th>");
                for (int j = 0, loopTo6 = i; j <= loopTo6; j++)
                {
                    if (i == j) // Skip this. We are not interested in HTMT of same latent variables
                    {
                        debug.PrintX("<td class=\"black\"></td>");
                    }
                    else if (htmtCalculations[Conversions.ToString(Operators.AddObject(latentVariables[j], latentVariables[i]))] >= constLiberal) // Check to see if HTMT is above our liberal threshold
                    {
                        debug.PrintX("<td class=\"liberal\">" + htmtCalculations[Conversions.ToString(Operators.AddObject(latentVariables[j], latentVariables[i]))].ToString("#0.000") + "</td>");
                        liberalWarnings.Add(latentVariables[j]); // Add first latent variable to warning list
                        liberalWarnings.Add(latentVariables[i]); // Add second latent variable to warning list
                    }
                    else if (htmtCalculations[Conversions.ToString(Operators.AddObject(latentVariables[j], latentVariables[i]))] >= constStrict) // Check to see if HTMT is below our liberal threshold but above our strict threshold
                    {
                        debug.PrintX("<td class=\"strict\">" + htmtCalculations[Conversions.ToString(Operators.AddObject(latentVariables[j], latentVariables[i]))].ToString("#0.000") + "</td>");
                        strictWarnings.Add(latentVariables[j]); // Add first latent variable to warning list
                        strictWarnings.Add(latentVariables[i]); // Add second latent variable to warning list
                    }
                    else
                    {
                        debug.PrintX("<td>" + htmtCalculations[Conversions.ToString(Operators.AddObject(latentVariables[j], latentVariables[i]))].ToString("#0.000") + "</td>");
                    } // Passes threshold, print normally.
                    lastvar = j;
                }
                for (int k = lastvar, loopTo7 = latentVariables.Count - 2; k <= loopTo7; k++) // Fill rest of matrix with blanks
                    debug.PrintX("<td></td>");
                debug.PrintX("</tr>");
            }
            debug.PrintX("</table><br>Thresholds are 0.850 for strict and 0.900 for liberal discriminant validity.<br>");

            // Write Warnings. Note, we step 2 because AI add
            debug.PrintX("<h3>HTMT Warnings</h3>");
            if (liberalWarnings.Count > 0 | strictWarnings.Count > 0)
            {
                for (int x = 0, loopTo8 = strictWarnings.Count - 1; x <= loopTo8; x += 2) // Strict Warnings
                    debug.PrintX(strictWarnings[x] + " and " + strictWarnings[x + 1] + " are statistically indistinguishable.");
                for (int x = 0, loopTo9 = liberalWarnings.Count - 1; x <= loopTo9; x += 2) // Liberal Warnings
                    debug.PrintX(liberalWarnings[x] + " and " + liberalWarnings[x + 1] + " are nearly indistinguishable.");
            }
            else
            {
                debug.PrintX("There are no warnings for this HTMT analysis.");
            }

            // Print references. Malhotra only prints if the described condition exists.
            debug.PrintX("<h3>References</h3>Significance of Correlations:<br>&#8224; p < 0.100<br>* p < 0.050<br>** p < 0.010<br>*** p < 0.001<br>");
            debug.PrintX("<br>Thresholds From:<br>Fornell, C., & Larcker, D. F. (1981). Evaluating structural equation models with unobservable variables and measurement error. Journal of marketing research, 39-50.");
            if (iMalhotra > 0)
            {
                debug.PrintX("<br><p><sup>1 </sup>Malhotra N. K., Dash S. argue that AVE is often too strict, and reliability can be established through CR alone.<br>Malhotra N. K., Dash S. (2011). Marketing Research an Applied Orientation. London: Pearson Publishing.");
            }
            debug.PrintX("<br>Henseler, J., C. M. Ringle, and M. Sarstedt (2015). A New Criterion for Assessing Discriminant Validity in Variance-based Structural Equation Modeling, Journal of the Academy of Marketing Science, 43 (1), 115-135.");
            debug.PrintX("<hr/><p>--If you would like to cite this tool directly, please use the following:");
            debug.PrintX("Gaskin, J., James, M., Lim, J, and Steed, J. (2022), \"Master Validity Tool\", AMOS Plugin. <a href=\"http://statwiki.gaskination.com\">Gaskination's StatWiki</a>.</p>");

            // Write Style And close
            debug.PrintX("<style>h1{margin-left:60px;}table{border:1px solid black;border-collapse:collapse;}td{border:1px solid black;text-align:center;padding:5px;}th{text-weight:bold;padding:10px;border: 1px solid black;}.liberal{background-color: #ec826b;}.strict{background-color: #ffe876b5;}.black{background-color: black;}</style>");
            debug.PrintX("</body></html>");

            // Take down our debugging, release file, open html
            Trace.Flush();
            Trace.Listeners.Remove(resultWriter);
            resultWriter.Close();
            resultWriter.Dispose();
            Process.Start("MasterValidity.html");

        }

        // Returns a dictionary with all unique combinations of htmt calculations
        public Dictionary<string, double> GetHTMTCalculations()
        {

            var smpCorrelationOutput = Util.GetXML("body/div/div[@ntype='groups']/div[@ntype='group'][position() = 1]/div[@ntype='samplemoments']/div[@ntype='ppml']/table/tbody");
            string rowHeader;
            string colHeader;
            double corrValue;
            var Sample = default(double[,]);
            var smpCorrelation = new Dictionary<string, double>();

            // Prepare structure equation modeling object
            var Sem = new AmosEngineLib.AmosEngine();
            Sem.NeedEstimates(AmosEngineLib.AmosEngine.TMatrixID.SampleCorrelations);
            pd.SpecifyModel(Sem);
            Sem.FitModel();
            Sem.GetEstimates(AmosEngineLib.AmosEngine.TMatrixID.SampleCorrelations, ref Sample); // Stores sample correlation estimates into the sample multi-dimensional array

            // Store Sample Moment's Sample Correlations Matrix In Dictionary (THIS PART IS IN TERMS OF INDICATORS)
            for (int i = 1, loopTo = Information.UBound(Sample, 1) + 1; i <= loopTo; i++)
            {
                for (int j = 1, loopTo1 = i; j <= loopTo1; j++)
                {
                    // rowHeader = Left Side of Matrix | colHeader = Top Side of Matridx | corrValue = Body of Matrix
                    rowHeader = Util.MatrixName(smpCorrelationOutput, i, 0L);
                    colHeader = Util.MatrixName(smpCorrelationOutput, j, 0L);
                    corrValue = Sample[i - 1, j - 1]; // Sample(,) Array Is Zero-Based

                    // Populate Dictionary With FULL Matrix (Normal + Inverted)
                    smpCorrelation.Add(rowHeader + colHeader, corrValue); 
                    if (colHeader + rowHeader != rowHeader + colHeader)
                    {
                        smpCorrelation.Add(colHeader + rowHeader, corrValue);
                    }
                }
            }; 
            Sem.Dispose(); // Close SEM object

            // Calculate HTMT Correlations For Each Unique Pair Of Latent Variables (THIS GIVES US A DICTIONARY OF HTMT CORRELATIONS IN TERMS OF LATENT VARIABLES)
            var latentVariables = GetLatent();
            var htmtCalcuations = new Dictionary<string, double>();
            var pairCalculations = new Dictionary<string, double>();
            ArrayList indicators1;
            ArrayList indicators2;
            var htmtSum = default(double);
            var htmtCount = default(int);
            double avgHTMT;

            // Loop Through Each Latent Variable
            foreach (var latent1 in latentVariables)
            {
                indicators1 = GetIndicators(Conversions.ToString(latent1)); // Grabs Indicators of the First Latent Variable
                                                                            // Loop Through Each Latent Variable With Each Latent Variable To Get Pairings
                foreach (var latent2 in latentVariables)
                {
                    indicators2 = GetIndicators(Conversions.ToString(latent2)); // Grabs Indicators of the Second Latent Variable
                    foreach (var indicator1 in indicators1)
                    {
                        foreach (var indicator2 in indicators2)
                        {
                            if (Conversions.ToBoolean(Operators.ConditionalCompareObjectNotEqual(indicator1, indicator2, false)))
                            {
                                htmtSum = htmtSum + smpCorrelation[Conversions.ToString(Operators.AddObject(indicator1, indicator2))]; // Wrapped in Try-Catch Because We Only Stored Unique Pairings. Sums All Indicator Pairs Together Then Takes Average
                                htmtCount = htmtCount + 1;
                            }
                        }
                    }
                    if (htmtCount > 0)
                    {
                        avgHTMT = htmtSum / htmtCount; // Calculation of Average Indicator of Latent Pairings
                        pairCalculations.Add(Conversions.ToString(Operators.AddObject(latent1, latent2)), avgHTMT); // Adds To Dictionary
                        htmtSum = 0d;
                        htmtCount = 0;
                        avgHTMT = 0d;
                    }
                }
            }

            // Takes Correlation Values From Above And Actually Calculates HTMT Value
            foreach (var latent1 in latentVariables)
            {
                foreach (var latent2 in latentVariables)
                {
                    double A = pairCalculations[Conversions.ToString(Operators.AddObject(latent1, latent2))]; // Heterotrait Correlation
                    double B = Math.Sqrt(pairCalculations[Conversions.ToString(Operators.AddObject(latent1, latent1))]); // Monotrait A Correlation Square Root
                    double C = Math.Sqrt(pairCalculations[Conversions.ToString(Operators.AddObject(latent2, latent2))]); // Monotrait B Correlation Square Root
                    double htmt = Math.Abs(A / (B * C));
                    htmtCalcuations.Add(Conversions.ToString(Operators.AddObject(latent1, latent2)), htmt); // Adds To Dictionary
                }
            }

            return htmtCalcuations;
        }

        // Creates an arraylist of recommendations to improve the model.
        public ArrayList checkValidity(double[,] estimateMatrix)
        {

            var validityMessages = new ArrayList(); // The list of messages.
            var latentVariables = GetLatent(); // The list of latent variables
                                               // Xml tables used to check estimates and number of rows.
            var tableCorrelation = Util.GetXML("body/div/div[@ntype='models']/div[@ntype='model'][position() = 1]/div[@ntype='group'][position() = 1]/div[@ntype='estimates']/div[@ntype='scalars']/div[@nodecaption='Correlations:']/table/tbody");
            var tableRegression = Util.GetXML("body/div/div[@ntype='models']/div[@ntype='model'][position() = 1]/div[@ntype='group'][position() = 1]/div[@ntype='estimates']/div[@ntype='scalars']/div[@nodecaption='Standardized Regression Weights:']/table/tbody");
            int numCorrelation = Util.GetNodeCount(tableCorrelation);
            int numRegression = Util.GetNodeCount(tableRegression);

            // Variables
            int numIndicator = 0; // Temp variable to check if more than two indicators on a latent.
            bool bMalhotra = true; // Malhotra is only used for certain conditions
            string tempMessage = ""; // Temp variable to check if the message already exists.
            string sMalhotra = "";

            // Based only on the variables in the correlations table
            for (int y = 0, loopTo = latentVariables.Count - 1; y <= loopTo; y++)
            {
                double indicatorVal = 2d;
                double indicatorTest = 0d;
                string indicatorName = "";

                for (int i = 1, loopTo1 = numRegression; i <= loopTo1; i++)
                {
                    if (Conversions.ToBoolean(Operators.ConditionalCompareObjectEqual(latentVariables[y], Util.MatrixName(tableRegression, i, 2L), false)))
                    {
                        numIndicator += 1;
                    }
                }

                if (numIndicator > 2) // Check if latent is connected to at least two indicators.
                {
                    for (int i = 1, loopTo2 = numRegression; i <= loopTo2; i++)
                    {
                        if (Conversions.ToBoolean(Operators.ConditionalCompareObjectEqual(latentVariables[y], Util.MatrixName(tableRegression, i, 2L), false)))
                        {
                            indicatorTest = Util.MatrixElement(tableRegression, i, 3L);
                            if (indicatorTest < indicatorVal) // Look for the lowest indicator.
                            {
                                indicatorName = Util.MatrixName(tableRegression, i, 0L);
                                indicatorVal = indicatorTest;
                            }
                        }
                    }
                }

                for (int x = 0; x <= 3; x++)
                {
                    if (x == 0 & estimateMatrix[y, x] < 0.7d)
                    {
                        bMalhotra = false;
                        if (numIndicator > 2) // Find the LOWEST indicator
                        {
                            tempMessage = Conversions.ToString("Reliability: the CR for " + latentVariables[y] + " is less than 0.70. Try removing " + indicatorName + " to improve CR.");
                            if (!validityMessages.Contains(tempMessage))
                            {
                                validityMessages.Add(tempMessage);
                            }
                        }
                        else
                        {
                            tempMessage = Conversions.ToString("Reliability: the CR for " + latentVariables[y] + " is less than 0.70. No way to improve CR because you only have two indicators for that variable. Removing one indicator will make this not latent.");
                            if (!validityMessages.Contains(tempMessage))
                            {
                                validityMessages.Add(tempMessage);
                            }
                        }
                        if (estimateMatrix[y, x] < estimateMatrix[y, x + 1])
                        {
                            tempMessage = Conversions.ToString("Convergent Validity: the CR for " + latentVariables[y] + " is less than the AVE.");
                            if (!validityMessages.Contains(tempMessage))
                            {
                                validityMessages.Add(tempMessage);
                            }
                        }
                    }
                    else if (x == 1)
                    {
                        if (estimateMatrix[y, x] < 0.5d)
                        {
                            if (bMalhotra == true)
                            {
                                sMalhotra = "<sup>1 </sup>";
                            }
                            if (numIndicator > 2)
                            {
                                tempMessage = Conversions.ToString(sMalhotra + "Convergent Validity: the AVE for " + latentVariables[y] + " is less than 0.50. Try removing " + indicatorName) + " to improve AVE.";
                                if (!validityMessages.Contains(tempMessage))
                                {
                                    validityMessages.Add(tempMessage);
                                }
                            }
                            else
                            {
                                tempMessage = Conversions.ToString(sMalhotra + "Convergent Validity: the AVE for " + latentVariables[y] + " is less than 0.50. No way to improve AVE because you only have two indicators for that variable. Removing one indicator will make this not latent.");
                                if (!validityMessages.Contains(tempMessage))
                                {
                                    validityMessages.Add(tempMessage);
                                }
                            }

                            if (Math.Abs(estimateMatrix[y, x]) < Math.Abs(estimateMatrix[y, x + 1]) & numCorrelation > 1) // Check if AVE is less than MSV.
                            {
                                tempMessage = Conversions.ToString("Discriminant Validity: the AVE for " + latentVariables[y] + " is less than the MSV.");
                                if (!validityMessages.Contains(tempMessage))
                                {
                                    validityMessages.Add(tempMessage);
                                }
                            }
                        }
                    }

                    // Only check latent variables that are correlated
                    for (int i = 1, loopTo3 = numCorrelation; i <= loopTo3; i++)
                    {
                        if (Conversions.ToBoolean(Operators.ConditionalCompareObjectEqual(latentVariables[y], Util.MatrixName(tableCorrelation, i, 0L), false)))
                        {
                            if (Math.Sqrt(Math.Abs(estimateMatrix[y, x])) < Math.Abs(Util.MatrixElement(tableCorrelation, i, 3L))) // Check if the square root of AVE is less than the correlation.
                            {
                                tempMessage = Conversions.ToString("Discriminant Validity: the square root of the AVE for " + latentVariables[y] + " is less than its correlation with " + Util.MatrixName(tableCorrelation, i, 2L)) + ".";
                                if (!validityMessages.Contains(tempMessage))
                                {
                                    validityMessages.Add(tempMessage);
                                }
                            }
                        }
                        else if (Conversions.ToBoolean(Operators.ConditionalCompareObjectEqual(latentVariables[y], Util.MatrixName(tableCorrelation, i, 2L), false)))
                        {
                            if (Math.Sqrt(Math.Abs(estimateMatrix[y, x])) < Math.Abs(Util.MatrixElement(tableCorrelation, i, 3L)))
                            {
                                tempMessage = Conversions.ToString("Discriminant Validity: the square root of the AVE for " + latentVariables[y] + " is less than its correlation with " + Util.MatrixName(tableCorrelation, i, 0L) + ".");
                                if (!validityMessages.Contains(tempMessage))
                                {
                                    validityMessages.Add(tempMessage);
                                }
                            }
                        }
                    }
                }
            }

            return validityMessages;

        }

        // Finds all indicators given a single latent variable
        public ArrayList GetIndicators(string latent)
        {

            var indicatorVariables = new ArrayList();
            var tableIndicators = Util.GetXML("body/div/div[@ntype='models']/div[@ntype='model'][position() = 1]/div[@ntype='group'][position() = 1]/div[@ntype='estimates']/div[@ntype='scalars']/div[@nodecaption='Standardized Regression Weights:']/table/tbody");

            // Loops through variables in the model and stores the latent variable names in an array.
            int numIndicators = Util.GetNodeCount(tableIndicators);

            // Higher Order: Grabs Information About Latent Variable Related To Other Latent Variables
            var RelatedLatentVariables = new ArrayList();
            RelatedLatentVariables.Add(latent);

            if (HigherOrderVariables.Count > 0)
            {
                foreach (PDElement variable in pd.PDElements)
                {
                    if (variable.IsPath() && (variable.Variable1.NameOrCaption ?? "") == (latent ?? ""))
                    {
                        RelatedLatentVariables.Add(variable.Variable2.NameOrCaption);
                    }
                }
            }

            for (int x = 1, loopTo = numIndicators; x <= loopTo; x++)
            {
                if (RelatedLatentVariables.Contains(Util.MatrixName(tableIndicators, x, 2L)) && !HigherOrderVariables.Contains(Util.MatrixName(tableIndicators, x, 2L)))
                {
                    // MsgBox(latent + ": " + Util.MatrixName(tableIndicators, x, 0))
                    indicatorVariables.Add(Util.MatrixName(tableIndicators, x, 0L));
                }

            }

            return indicatorVariables;

        }

        // Get the CR, AVE, MSV, MaxR, and Sqrt for a latent variable.
        public Estimates GetEstimates(string latent)
        {

            // Store the standardized regression output table into a matrix.
            var tableRegression = Util.GetXML("body/div/div[@ntype='models']/div[@ntype='model'][position() = 1]/div[@ntype='group'][position() = 1]/div[@ntype='estimates']/div[@ntype='scalars']/div[@nodecaption='Standardized Regression Weights:']/table/tbody");
            var tableRegressionCI = Util.GetXML("body/div/div[@ntype='models']/div[@ntype='model'][position() = 1]/div[@ntype='group'][position() = 1]/div[@ntype='bootstrap']/div[@ntype='bootstrapconfidence']/div[@ntype='percentile']/div[@ntype='scalars']/div[@nodecaption='Standardized Regression Weights:']/table/tbody");

            // Count of elements in the regression table.
            int numRegression = Util.GetNodeCount(tableRegression);

            Estimates estimates;
            double MaxR = 0d; // Sum of squared SRW over one minus SRW squared
            double SRW = 0d; // Standardized Regression Weight
            double SSI = 0d; // Sum of the square SRW values for a variable
            double SL2 = 0d; // Squared SRW value
            double SL3 = 0d; // Sum of the SRW values for a variable
            double SRWError = 0d; // Sum of SRW errors
            double numSRW = 0d; // Count of SRW values

            // Lower Values
            double LSRW = 0d; // Standardized Regression Weight
            double LSSI = 0d; // Sum of the square SRW values for a variable
            double LSL2 = 0d; // Squared SRW value
            double LSL3 = 0d; // Sum of the SRW values for a variable
            double LSRWError = 0d; // Sum of SRW errors

            // Upper Values
            double USRW = 0d; // Standardized Regression Weight
            double USSI = 0d; // Sum of the square SRW values for a variable
            double USL2 = 0d; // Squared SRW value
            double USL3 = 0d; // Sum of the SRW values for a variable
            double USRWError = 0d; // Sum of SRW errors

            // Higher Order: Grabs Information About Latent Variable Related To Other Latent Variables
            var RelatedLatentVariables = new ArrayList();
            RelatedLatentVariables.Add(latent);

            // Loop through regression matrix to calculate estimates.
            for (int i = 1, loopTo = numRegression; i <= loopTo; i++)
            {
                if (RelatedLatentVariables.Contains(Util.MatrixName(tableRegression, i, 2L)))
                {
                    // Normal
                    SRW = Util.MatrixElement(tableRegression, i, 3L);
                    SL2 = Math.Pow(SRW, 2d);
                    SL3 = SL3 + SRW;
                    SSI = SSI + SL2;
                    SRWError = SRWError + (1d - SL2);
                    MaxR = MaxR + SL2 / (1d - SL2);
                    numSRW += 1d;
                    // Lower
                    LSRW = Util.MatrixElement(tableRegressionCI, i, 4L);
                    LSL2 = Math.Pow(LSRW, 2d);
                    LSL3 = LSL3 + LSRW;
                    LSSI = LSSI + LSL2;
                    LSRWError = LSRWError + (1d - LSL2);
                    // Upper
                    USRW = Util.MatrixElement(tableRegressionCI, i, 5L);
                    USL2 = Math.Pow(USRW, 2d);
                    USL3 = USL3 + USRW;
                    USSI = USSI + USL2;
                    USRWError = USRWError + (1d - USL2);

                    // MsgBox(Util.MatrixName(tableRegression, i, 0) + " " + Util.MatrixName(tableRegression, i, 2) + " " + SRW.ToString("#.###" + " " + LSRW.ToString("#.###") + " " + USRW.ToString("#.###")))
                }
            }

            // Output estimates to struct.
            SL3 = Math.Pow(SL3, 2d);
            estimates.CR = SL3 / (SL3 + SRWError);
            estimates.AVE = SSI / numSRW;
            estimates.MSV = GetMSV(latent);
            estimates.MaxR = 1d / (1d + 1d / MaxR);
            estimates.SQRT = Math.Sqrt(SSI / numSRW);

            // Lower + Upper
            LSL3 = Math.Pow(LSL3, 2d);
            USL3 = Math.Pow(USL3, 2d);
            estimates.LCR = LSL3 / (LSL3 + LSRWError);
            estimates.UCR = USL3 / (USL3 + USRWError);
            estimates.LAVE = LSSI / numSRW;
            estimates.UAVE = USSI / numSRW;

            return estimates;

        }

        // Matrix that holds the estimates for each latent variable.
        double[,] GetMatrix()
        {

            // Holds the correlation matrix from the output table.
            var tableCorrelation = Util.GetXML("body/div/div[@ntype='models']/div[@ntype='model'][position() = 1]/div[@ntype='group'][position() = 1]/div[@ntype='estimates']/div[@ntype='scalars']/div[@nodecaption='Correlations:']/table/tbody");
            // Count of elements in the correlation table
            int numCorrelation = Util.GetNodeCount(tableCorrelation);

            // Get an array with the names of the latent variables.
            var latentVariables = GetLatent();
            int numLatentVariables = latentVariables.Count;

            // Declare a two dimensional array to hole the estimates.
            var estimateMatrix = new double[numLatentVariables * 2 - 1 + 1, latentVariables.Count + 3 + 1];

            // Catch if there is only one latent in the model.
            CheckSingleLatent(latentVariables.Count);

            foreach (var latent in latentVariables)
            {

                var estimates = GetEstimates(Conversions.ToString(latent));
                int matrixColumn = 1;
                int matrixRow = latentVariables.IndexOf(latent);

                // First four rows are estimates
                estimateMatrix[matrixRow, 0] = estimates.CR;
                estimateMatrix[matrixRow, matrixColumn] = estimates.AVE;
                matrixColumn += 1;
                if (numCorrelation >= 1)
                {
                    estimateMatrix[matrixRow, matrixColumn] = estimates.MSV;
                    matrixColumn += 1;
                }
                estimateMatrix[matrixRow, matrixColumn] = estimates.MaxR;

                // Get the Correlation table and put it into an aligned matrix.
                if (numCorrelation != 0)
                {
                    for (int i = 1, loopTo = numCorrelation; i <= loopTo; i++)
                    {
                        if (Conversions.ToBoolean(Operators.ConditionalCompareObjectEqual(latent, Util.MatrixName(tableCorrelation, i, 2L), false)))
                        {
                            for (int index = 0, loopTo1 = latentVariables.Count - 1; index <= loopTo1; index++)
                            {
                                if (Conversions.ToBoolean(Operators.ConditionalCompareObjectEqual(latentVariables[index], Util.MatrixName(tableCorrelation, i, 0L), false)))
                                {
                                    // Catches an output table exception where the variable is only listed on one side.
                                    if (index + 4 > latentVariables.IndexOf(latent) + 4 & Util.MatrixElement(tableCorrelation, i, 3L) > 0d)
                                    {
                                        estimateMatrix[index, matrixRow + 4] = Util.MatrixElement(tableCorrelation, i, 3L);
                                        estimateMatrix[matrixRow, index + 4] = Util.MatrixElement(tableCorrelation, i, 3L);
                                    }
                                    else
                                    {
                                        estimateMatrix[index, matrixRow + 4] = Util.MatrixElement(tableCorrelation, i, 3L);
                                        estimateMatrix[matrixRow, index + 4] = Util.MatrixElement(tableCorrelation, i, 3L);
                                    }
                                }
                            }
                            if (Conversions.ToBoolean(Operators.ConditionalCompareObjectEqual(latent, latentVariables[matrixRow], false)))
                            {
                                estimateMatrix[matrixRow, matrixRow + 4] = estimates.SQRT;
                            }
                        }
                        else if (Conversions.ToBoolean(Operators.ConditionalCompareObjectEqual(latent, latentVariables[matrixRow], false)))
                        {
                            estimateMatrix[matrixRow, matrixRow + 4] = estimates.SQRT;
                        }
                    }
                }

                // Store CI Numbers
                estimateMatrix[matrixRow + numLatentVariables, 0] = estimates.CR;
                estimateMatrix[matrixRow + numLatentVariables, 1] = estimates.AVE;
                estimateMatrix[matrixRow + numLatentVariables, 2] = estimates.LCR;
                estimateMatrix[matrixRow + numLatentVariables, 3] = estimates.UCR;
                estimateMatrix[matrixRow + numLatentVariables, 4] = estimates.LAVE;
                estimateMatrix[matrixRow + numLatentVariables, 5] = estimates.UAVE;
            }

            return estimateMatrix;

        }

        #region Helper Functions

        // Check if the model has a single latent factor.
        public void CheckSingleLatent(int iLatent)
        {

            // Assign the variance estimates to a matrix
            var tableVariance = Util.GetXML("body/div/div[@ntype='models']/div[@ntype='model'][position() = 1]/div[@ntype='group'][position() = 1]/div[@ntype='estimates']/div[@ntype='scalars']/div[@nodecaption='Variances:']/table/tbody");

            // Loop through variance table with the number of latents
            for (int i = 1, loopTo = iLatent + 1; i <= loopTo; i++)
            {
                try
                {
                    Util.MatrixName(tableVariance, i, 3L); // Checks if there is no variance for an unobserved variable.
                }
                catch (Exception ex)
                {
                    for (int k = 1; k <= iLatent + 1; k++)
                    {
                        try
                        {
                            if (Util.MatrixName(tableVariance, k, 3L) != default) // Checks if the variable is "unidentified"
                            {
                                Interaction.MsgBox(Util.MatrixName(tableVariance, k, 0L) + @" is causing an error. Either:
                            1. It only has one indicator and is therefore, not latent. The CFA is only for latent variables, so don’t include " + Util.MatrixName(tableVariance, k, 0L) + @" in the CFA.
                            2. You are missing a constraint on an indicator.");
                            }
                        }
                        catch (NullReferenceException exc)
                        {
                            continue;
                        }
                    }
                }
            }

        }

        // Create an arraylist of the names of the latent variables.
        public ArrayList GetLatent()
        {

            var latentVariables = new ArrayList();

            // Loops through variables in the model and stores the latent variable names in an array.
            foreach (PDElement variable in pd.PDElements)
            {
                if (variable.IsLatentVariable() & !variable.IsEndogenousVariable())
                {
                    latentVariables.Add(variable.NameOrCaption);
                }
            }

            return latentVariables;

        }

        // Calculate the MSV for a correlation.
        public double GetMSV(string latent)
        {

            // Store correlation table in a matrix
            var tableCorrelation = Util.GetXML("body/div/div[@ntype='models']/div[@ntype='model'][position() = 1]/div[@ntype='group'][position() = 1]/div[@ntype='estimates']/div[@ntype='scalars']/div[@nodecaption='Correlations:']/table/tbody");
            // Count of elements in correlation table
            int numCorrelation = Util.GetNodeCount(tableCorrelation);

            // Variables
            double MSV = 0d;
            double testMSV = 0d;

            // Takes the max correlation as the MSV
            for (int i = 1, loopTo = numCorrelation; i <= loopTo; i++)
            {
                if ((latent ?? "") == (Util.MatrixName(tableCorrelation, i, 0L) ?? "") | (latent ?? "") == (Util.MatrixName(tableCorrelation, i, 2L) ?? ""))
                {
                    testMSV = Math.Pow(Util.MatrixElement(tableCorrelation, i, 3L), 2d);
                    if (testMSV > MSV)
                    {
                        MSV = testMSV;
                    }
                }
            }

            return MSV;

        }

        

        #endregion

    }
}

#endregion
