#Region "Imports"
Imports System.Diagnostics
Imports AmosEngineLib.AmosEngine.TMatrixID
Imports System.Xml
Imports Amos
Imports System.Collections.Generic
Imports System.Linq

'This plugin was written July 2020 by Matthew James for James Gaskin
'Updated July 2020
<System.ComponentModel.Composition.Export(GetType(Amos.IPlugin))>
Public Class CustomCode
    Implements IPlugin

    Dim HigherOrderVariables As New ArrayList

    Public Function Name() As String Implements IPlugin.Name
        Return "Validity and Reliability Test w/ HTMT Analysis"
    End Function

    Public Function Description() As String Implements IPlugin.Description
        Return "Calculates CR, AVE, ASV, MSV and outputs them in a table with the correlations. Includes interpretations from the results. See statwiki.kolobkreations.com for more information."
    End Function

    'Struct to hold the estimates for a latent variable that you are testing.
    Structure Estimates

        Public CR As Double
        Public AVE As Double
        Public LCR As Double
        Public UCR As Double
        Public LAVE As Double
        Public UAVE As Double
        Public MSV As Double
        Public MaxR As Double
        Public SQRT As Double

    End Structure

    Public Function Mainsub() As Integer Implements IPlugin.MainSub

        'Check if the model is a causal rather than measurment model by looping through paths.
        For Each variable As PDElement In pd.PDElements
            If variable.IsPath Then
                'Model contains paths from latent to latent or observed to observed.
                If (variable.Variable1.IsObservedVariable And variable.Variable2.IsObservedVariable) Or (variable.Variable1.IsObservedVariable And variable.Variable2.IsLatentVariable) Then
                    MsgBox("The master validity plugin cannot be used for causal models.")
                    Exit Function
                End If
                If variable.Variable1.IsLatentVariable And variable.Variable2.IsLatentVariable Then
                    For Each subvariable As PDElement In pd.PDElements
                        If subvariable.IsPath Then
                            If (variable.Variable1.NameOrCaption = subvariable.Variable1.NameOrCaption) And subvariable.Variable2.IsObservedVariable Then
                                MsgBox("The master validity plugin cannot be used for causal models.")
                                Exit Function
                            End If
                        End If
                    Next
                End If
            End If
        Next

        'Ensure that the output includes standardized estimates.
        Amos.pd.GetCheckBox("AnalysisPropertiesForm", "StandardizedCheck").Checked = True
        Amos.pd.GetCheckBox("AnalysisPropertiesForm", "SampleMomCheck").Checked = True
        Amos.pd.GetCheckBox("AnalysisPropertiesForm", "MeansInterceptsCheck").Checked = True
        Amos.pd.GetCheckBox("AnalysisPropertiesForm", "CorestCheck").Checked = True
        Amos.pd.GetCheckBox("AnalysisPropertiesForm", "CovestCheck").Checked = True

        'Enable Boostrap
        Amos.pd.GetCheckBox("AnalysisPropertiesForm", "DoBootstrapCheck").Checked = True
        Amos.pd.GetCheckBox("AnalysisPropertiesForm", "ConfidencePCCheck").Checked = True
        Amos.pd.GetTextBox("AnalysisPropertiesForm", "BootstrapText").Text = "500"
        Amos.pd.GetTextBox("AnalysisPropertiesForm", "ConfidencePCText").Text = "95"

        'Method to run the model.
        Amos.pd.AnalyzeCalculateEstimates()

        'If error, some variable in the model is unlinked or "floating." Ensure that all variables in model are linked. For example, "E1" (Error Term 1) is not linked to an observed variable.
        If GetNodeCount(GetXML("body/div/div[@ntype='models']/div[@ntype='model'][position() = 1]/div[@ntype='group'][position() = 1]/div[@ntype='estimates']/div[@ntype='scalars']/div[@nodecaption='Standardized Regression Weights:']/table/tbody")) = 0 Then
            MsgBox("Invalid Model: Cannot Compute Standardized Regression Weights. All variables need to have at least one relationship to another variable.")
            Exit Function
        End If

        'Grab Information Related To Higher Order Models
        HigherOrderVariables = GetHigherOrderLatent()

        'Produces the matrix of estimates and correlations.
        Dim estimateMatrix As Double(,) = GetMatrix()

        'Sub procedure to create the output file.
        CreateOutput(estimateMatrix)

    End Function

    'Use the estimates matrix to create an html file of output.
    Sub CreateOutput(estimateMatrix As Double(,))

        'Get the output tables and the count of correlations
        Dim tableCorrelation As XmlElement = GetXML("body/div/div[@ntype='models']/div[@ntype='model'][position() = 1]/div[@ntype='group'][position() = 1]/div[@ntype='estimates']/div[@ntype='scalars']/div[@nodecaption='Correlations:']/table/tbody")
        Dim tableCovariance As XmlElement = GetXML("body/div/div[@ntype='models']/div[@ntype='model'][position() = 1]/div[@ntype='group'][position() = 1]/div[@ntype='estimates']/div[@ntype='scalars']/div[@nodecaption='Covariances:']/table/tbody")
        Dim numCorrelation As Integer = GetNodeCount(tableCorrelation)

        'Get a list of the latent variables
        Dim latentVariables As ArrayList = GetLatent()

        'Check the matrix for validity concerns
        Dim validityConcerns As ArrayList = checkValidity(estimateMatrix)

        'Variables
        Dim bMalhotra As Boolean = True
        Dim iMalhotra As Integer = 0
        Dim missingCorrelation As Boolean = False

        'Delete the output file if it exists
        If (System.IO.File.Exists("MasterValidity.html")) Then
            System.IO.File.Delete("MasterValidity.html")
        End If

        'Start the debugger for the html output
        Dim debug As New AmosDebug.AmosDebug
        'File the output will be written to
        Dim resultWriter As New TextWriterTraceListener("MasterValidity.html")
        Trace.Listeners.Add(resultWriter)

        debug.PrintX("<html><body><h1>Model Validity Measures</h1><hr/><h3>Validity Analysis</h3><table><tr><td></td><th>CR</th><th>AVE</th>")

        'Handle zero correlations
        If numCorrelation > 0 Then
            debug.PrintX("<th>MSV</th><th>MaxR(H)</th>")
            For Each latent In latentVariables
                debug.PrintX("<th>" + latent + "</th>")
            Next
        Else
            debug.PrintX("<th>MaxR(H)</th>")
        End If
        debug.PrintX("</tr>")

        'Loop through the matrix to output variables, estimates, and correlations
        For y = 0 To latentVariables.Count - 1
            debug.PrintX("<tr><td><span style='font-weight:bold;'>" + latentVariables(y) + "</span></td>")
            For x = 0 To latentVariables.Count + 3
                Dim significance As String = ""
                Select Case x
                    Case 0
                        If estimateMatrix(y, x) < 0.7 Or estimateMatrix(y, x) < estimateMatrix(y, (x + 1)) Then
                            debug.PrintX("<td style='color:red'>")
                            bMalhotra = False
                        Else
                            debug.PrintX("<td>")
                        End If
                    Case 1
                        If estimateMatrix(y, x) < 0.5 Or (estimateMatrix(y, x) < estimateMatrix(y, (x + 1)) And numCorrelation > 1) Then
                            debug.PrintX("<td style='color:red'>")
                            If bMalhotra = True Then
                                iMalhotra += 1
                            End If
                        Else
                            debug.PrintX("<td>")
                        End If
                    Case y + 4
                        Dim red As Boolean = False
                        For i = 1 To numCorrelation
                            If latentVariables(y) = MatrixName(tableCorrelation, i, 0) Or latentVariables(y) = MatrixName(tableCorrelation, i, 2) Then
                                If Math.Sqrt(estimateMatrix(y, x)) < MatrixElement(tableCorrelation, i, 3) Then
                                    red = True
                                End If
                            End If
                        Next
                        If red = True And numCorrelation > 0 Then
                            debug.PrintX("<td><span style='font-weight:bold; color:red;'>")
                        ElseIf numCorrelation > 0 Then
                            debug.PrintX("<td><span style='font-weight:bold;'>")
                        End If
                    Case 2
                        debug.PrintX("<td>")
                    Case 3
                        If numCorrelation > 0 Then
                            debug.PrintX("<td>")
                        End If
                    Case Else
                        'Interpretations of significance
                        If numCorrelation > 0 Then
                            debug.PrintX("<td>")
                            For i = 1 To numCorrelation
                                If (MatrixName(tableCovariance, i, 0) = latentVariables(x - 4) And MatrixName(tableCovariance, i, 2) = latentVariables(y)) Or ((MatrixName(tableCovariance, i, 0) = latentVariables(y)) And MatrixName(tableCovariance, i, 2) = latentVariables(x - 4)) Then
                                    If MatrixName(tableCovariance, i, 6) = "***" Then
                                        significance = "***"
                                    ElseIf MatrixElement(tableCovariance, i, 6) < 0.01 Then
                                        significance = "**"
                                    ElseIf MatrixElement(tableCovariance, i, 6) < 0.05 Then
                                        significance = "*"
                                    ElseIf MatrixElement(tableCovariance, i, 6) < 0.1 Then
                                        significance = "&#8224;"
                                    Else
                                        significance = ""
                                    End If
                                End If
                            Next
                        End If
                        Exit Select
                End Select

                If x <= 2 Then
                    debug.PrintX(estimateMatrix(y, x).ToString("#0.000") + "</span></td>")
                ElseIf x > 2 And estimateMatrix(y, x) <> 0 And numCorrelation > 0 And y + 4 >= x Then
                    debug.PrintX(estimateMatrix(y, x).ToString("#0.000") + significance + "</span></td>")
                ElseIf x > 2 And numCorrelation > 0 And estimateMatrix(y, x) = 0 And y + 4 > x Then
                    debug.PrintX("&#8258</td>")
                    missingCorrelation = True
                Else
                    debug.PrintX("</td>")
                End If
            Next
            debug.PrintX("</tr>")
        Next
        debug.PrintX("</table><br>")

        If missingCorrelation = True Then
            debug.PrintX(" &#8258 Correlation is not specified in the model.<br>")
        End If

        If numCorrelation = 0 Then
            debug.PrintX("You only had one latent variable so there is no correlation matrix or MSV.<br>")
        ElseIf validityConcerns.Count = 0 Then
            debug.PrintX("No validity concerns here.<br>")
        End If

        'Print the validity concerns
        For Each message In validityConcerns
            debug.PrintX(message & "<br>")
        Next

        'Print confidence intervals
        debug.PrintX("<h3>Validity Analysis - Confidence Intervals</h3><table><tr><td></td><th>CR</th><th>AVE</th><th>Lower 95% CR</th><th>Upper 95% CR</th><th>Lower 95% AVE</th><th>Upper 95% AVE</th>")

        For y = 0 To latentVariables.Count - 1
            debug.PrintX("</tr><tr><td><span style='font-weight:bold;'>" + latentVariables(y) + "</span></td>")
            For x = 0 To 5
                debug.PrintX("<td>" + estimateMatrix(y + latentVariables.Count, x).ToString("#0.000") + "</span></td>")
            Next
        Next

        debug.PrintX("</tr></table><br>")

        'HTMT Analysis
        'Variables needed for output
        Const constLiberal As Double = 0.9
        Const constStrict As Double = 0.85
        Dim liberalWarnings As New ArrayList
        Dim strictWarnings As New ArrayList
        latentVariables = GetLatent()
        Dim htmtCalculations As Dictionary(Of String, Double) = GetHTMTCalculations()

        'Write the beginning of the document
        debug.PrintX("<h3>HTMT Analysis</h3><table><tr><th></th>")
        For Each latent In latentVariables 'Headers (latent variables) for table columns
            debug.PrintX("<th>" + latent + "</th>")
        Next
        debug.PrintX("</tr>")

        Dim lastvar As Integer
        For i = 0 To latentVariables.Count - 1 'Headers (latent variable) for table rows
            debug.PrintX("<tr><th>" + latentVariables(i) + "</th>")
            For j = 0 To i
                If i = j Then 'Skip this. We are not interested in HTMT of same latent variables
                    debug.PrintX("<td class=""black""></td>")
                ElseIf htmtCalculations(latentVariables(j) + latentVariables(i)) >= constLiberal Then 'Check to see if HTMT is above our liberal threshold
                    debug.PrintX("<td class=""liberal"">" + htmtCalculations(latentVariables(j) + latentVariables(i)).ToString("#0.000") + "</td>")
                    liberalWarnings.Add(latentVariables(j)) 'Add first latent variable to warning list
                    liberalWarnings.Add(latentVariables(i)) 'Add second latent variable to warning list
                ElseIf htmtCalculations(latentVariables(j) + latentVariables(i)) >= constStrict Then 'Check to see if HTMT is below our liberal threshold but above our strict threshold
                    debug.PrintX("<td class=""strict"">" + htmtCalculations(latentVariables(j) + latentVariables(i)).ToString("#0.000") + "</td>")
                    strictWarnings.Add(latentVariables(j)) 'Add first latent variable to warning list
                    strictWarnings.Add(latentVariables(i)) 'Add second latent variable to warning list
                Else
                    debug.PrintX("<td>" + htmtCalculations(latentVariables(j) + latentVariables(i)).ToString("#0.000") + "</td>") 'Passes threshold, print normally.
                End If
                lastvar = j
            Next
            For k = lastvar To latentVariables.Count - 2 'Fill rest of matrix with blanks
                debug.PrintX("<td></td>")
            Next
            debug.PrintX("</tr>")
        Next
        debug.PrintX("</table><br>Thresholds are 0.850 for strict and 0.900 for liberal discriminant validity.<br>")

        'Write Warnings. Note, we step 2 because AI add
        debug.PrintX("<h3>HTMT Warnings</h3>")
        If liberalWarnings.Count > 0 Or strictWarnings.Count > 0 Then
            For x = 0 To strictWarnings.Count - 1 Step 2 'Strict Warnings
                debug.PrintX(strictWarnings(x) + " and " + strictWarnings(x + 1) + " are statistically indistinguishable.")
            Next
            For x = 0 To liberalWarnings.Count - 1 Step 2 'Liberal Warnings
                debug.PrintX(liberalWarnings(x) + " and " + liberalWarnings(x + 1) + " are nearly indistinguishable.")
            Next
        Else
            debug.PrintX("There are no warnings for this HTMT analysis.")
        End If

        'Print references. Malhotra only prints if the described condition exists.
        debug.PrintX("<h3>References</h3>Significance of Correlations:<br>&#8224; p < 0.100<br>* p < 0.050<br>** p < 0.010<br>*** p < 0.001<br>")
        debug.PrintX("<br>Thresholds From:<br>Fornell, C., & Larcker, D. F. (1981). Evaluating structural equation models with unobservable variables and measurement error. Journal of marketing research, 39-50.")
        If iMalhotra > 0 Then
            debug.PrintX("<br><p><sup>1 </sup>Malhotra N. K., Dash S. argue that AVE is often too strict, and reliability can be established through CR alone.<br>Malhotra N. K., Dash S. (2011). Marketing Research an Applied Orientation. London: Pearson Publishing.")
        End If
        debug.PrintX("<br>Henseler, J., C. M. Ringle, and M. Sarstedt (2015). A New Criterion for Assessing Discriminant Validity in Variance-based Structural Equation Modeling, Journal of the Academy of Marketing Science, 43 (1), 115-135.")
        debug.PrintX("<hr/><p>--If you would like to cite this tool directly, please use the following:")
        debug.PrintX("Gaskin, J., James, M., and Lim, J. (2019), ""Master Validity Tool"", AMOS Plugin. <a href=""http://statwiki.kolobkreations.com"">Gaskination's StatWiki</a>.</p>")

        'Write Style And close
        debug.PrintX("<style>h1{margin-left:60px;}table{border:1px solid black;border-collapse:collapse;}td{border:1px solid black;text-align:center;padding:5px;}th{text-weight:bold;padding:10px;border: 1px solid black;}.liberal{background-color: #ec826b;}.strict{background-color: #ffe876b5;}.black{background-color: black;}</style>")
        debug.PrintX("</body></html>")

        'Take down our debugging, release file, open html
        Trace.Flush()
        Trace.Listeners.Remove(resultWriter)
        resultWriter.Close()
        resultWriter.Dispose()
        Process.Start("MasterValidity.html")

    End Sub

    'Returns a dictionary with all unique combinations of htmt calculations
    Function GetHTMTCalculations() As Dictionary(Of String, Double)

        Dim smpCorrelationOutput As XmlElement = GetXML("body/div/div[@ntype='groups']/div[@ntype='group'][position() = 1]/div[@ntype='samplemoments']/div[@ntype='ppml']/table/tbody")
        Dim rowHeader As String
        Dim colHeader As String
        Dim corrValue As Double
        Dim Sample(,) As Double
        Dim smpCorrelation As New Dictionary(Of String, Double)

        'Prepare structure equation modeling object
        Dim Sem As New AmosEngineLib.AmosEngine
        Sem.NeedEstimates(SampleCorrelations)
        Amos.pd.SpecifyModel(Sem)
        Sem.FitModel()
        Sem.GetEstimates(SampleCorrelations, Sample) 'Stores sample correlation estimates into the sample multi-dimensional array

        'Store Sample Moment's Sample Correlations Matrix In Dictionary (THIS PART IS IN TERMS OF INDICATORS)
        For i = 1 To UBound(Sample, 1) + 1
            For j = 1 To i
                'rowHeader = Left Side of Matrix | colHeader = Top Side of Matridx | corrValue = Body of Matrix
                rowHeader = MatrixName(smpCorrelationOutput, i, 0)
                colHeader = MatrixName(smpCorrelationOutput, j, 0)
                corrValue = Sample(i - 1, j - 1) 'Sample(,) Array Is Zero-Based

                'Populate Dictionary With FULL Matrix (Normal + Inverted)
                smpCorrelation.Add(rowHeader + colHeader, corrValue)
                On Error Resume Next 'ONLY ERROR THAT OCCURS IS WHEN WE ARE ADDING THE SAME INDICATOR PAIR (I.E., COMPUSE_1 & COMPUSE_1)
                smpCorrelation.Add(colHeader + rowHeader, corrValue)
            Next
        Next

        On Error GoTo 0 'Reset Error Handling

        Sem.Dispose() 'Close SEM object

        'Calculate HTMT Correlations For Each Unique Pair Of Latent Variables (THIS GIVES US A DICTIONARY OF HTMT CORRELATIONS IN TERMS OF LATENT VARIABLES)
        Dim latentVariables As ArrayList = GetLatent()
        Dim htmtCalcuations As New Dictionary(Of String, Double)
        Dim pairCalculations As New Dictionary(Of String, Double)
        Dim indicators1 As ArrayList
        Dim indicators2 As ArrayList
        Dim htmtSum As Double
        Dim htmtCount As Integer
        Dim avgHTMT As Double

        'Loop Through Each Latent Variable
        For Each latent1 In latentVariables
            indicators1 = GetIndicators(latent1) 'Grabs Indicators of the First Latent Variable
            'Loop Through Each Latent Variable With Each Latent Variable To Get Pairings
            For Each latent2 In latentVariables
                indicators2 = GetIndicators(latent2) 'Grabs Indicators of the Second Latent Variable
                For Each indicator1 In indicators1
                    For Each indicator2 In indicators2
                        If indicator1 <> indicator2 Then
                            htmtSum = htmtSum + smpCorrelation(indicator1 + indicator2) 'Wrapped in Try-Catch Because We Only Stored Unique Pairings. Sums All Indicator Pairs Together Then Takes Average
                            htmtCount = htmtCount + 1
                        End If
                    Next
                Next
                If (htmtCount > 0) Then
                    avgHTMT = htmtSum / htmtCount 'Calculation of Average Indicator of Latent Pairings
                    pairCalculations.Add(latent1 + latent2, avgHTMT) 'Adds To Dictionary
                    htmtSum = 0
                    htmtCount = 0
                    avgHTMT = 0
                End If
            Next
        Next

        'Takes Correlation Values From Above And Actually Calculates HTMT Value
        For Each latent1 In latentVariables
            For Each latent2 In latentVariables
                Dim A As Double = pairCalculations(latent1 + latent2) 'Heterotrait Correlation
                Dim B As Double = Math.Sqrt(pairCalculations(latent1 + latent1)) 'Monotrait A Correlation Square Root
                Dim C As Double = Math.Sqrt(pairCalculations(latent2 + latent2)) 'Monotrait B Correlation Square Root
                Dim htmt As Double = Math.Abs(A / (B * C))
                htmtCalcuations.Add(latent1 + latent2, htmt) 'Adds To Dictionary
            Next
        Next

        Return htmtCalcuations
    End Function

    'Creates an arraylist of recommendations to improve the model.
    Function checkValidity(estimateMatrix As Double(,)) As ArrayList

        Dim validityMessages As New ArrayList 'The list of messages.
        Dim latentVariables As ArrayList = GetLatent() 'The list of latent variables
        'Xml tables used to check estimates and number of rows.
        Dim tableCorrelation As XmlElement = GetXML("body/div/div[@ntype='models']/div[@ntype='model'][position() = 1]/div[@ntype='group'][position() = 1]/div[@ntype='estimates']/div[@ntype='scalars']/div[@nodecaption='Correlations:']/table/tbody")
        Dim tableRegression As XmlElement = GetXML("body/div/div[@ntype='models']/div[@ntype='model'][position() = 1]/div[@ntype='group'][position() = 1]/div[@ntype='estimates']/div[@ntype='scalars']/div[@nodecaption='Standardized Regression Weights:']/table/tbody")
        Dim numCorrelation As Integer = GetNodeCount(tableCorrelation)
        Dim numRegression As Integer = GetNodeCount(tableRegression)

        'Variables
        Dim numIndicator As Integer = 0 'Temp variable to check if more than two indicators on a latent.
        Dim bMalhotra As Boolean = True 'Malhotra is only used for certain conditions
        Dim tempMessage As String = "" 'Temp variable to check if the message already exists.
        Dim sMalhotra As String = ""

        'Based only on the variables in the correlations table
        For y = 0 To latentVariables.Count - 1
            Dim indicatorVal As Double = 2
            Dim indicatorTest As Double = 0
            Dim indicatorName As String = ""

            For i = 1 To numRegression
                If latentVariables(y) = MatrixName(tableRegression, i, 2) Then
                    numIndicator += 1
                End If
            Next

            If numIndicator > 2 Then 'Check if latent is connected to at least two indicators.
                For i = 1 To numRegression
                    If latentVariables(y) = MatrixName(tableRegression, i, 2) Then
                        indicatorTest = MatrixElement(tableRegression, i, 3)
                        If indicatorTest < indicatorVal Then 'Look for the lowest indicator.
                            indicatorName = MatrixName(tableRegression, i, 0)
                            indicatorVal = indicatorTest
                        End If
                    End If
                Next
            End If

            For x = 0 To 3
                If (x = 0 And estimateMatrix(y, x) < 0.7) Then
                    bMalhotra = False
                    If numIndicator > 2 Then 'Find the LOWEST indicator
                        tempMessage = "Reliability: the CR for " & latentVariables(y) & " is less than 0.70. Try removing " & indicatorName & " to improve CR."
                        If Not validityMessages.Contains(tempMessage) Then
                            validityMessages.Add(tempMessage)
                        End If
                    Else
                        tempMessage = "Reliability: the CR for " & latentVariables(y) & " is less than 0.70. No way to improve CR because you only have two indicators for that variable. Removing one indicator will make this not latent."
                        If Not validityMessages.Contains(tempMessage) Then
                            validityMessages.Add(tempMessage)
                        End If
                    End If
                    If estimateMatrix(y, x) < estimateMatrix(y, (x + 1)) Then
                        tempMessage = "Convergent Validity: the CR for " & latentVariables(y) & " is less than the AVE."
                        If Not validityMessages.Contains(tempMessage) Then
                            validityMessages.Add(tempMessage)
                        End If
                    End If

                ElseIf x = 1 Then
                    If estimateMatrix(y, x) < 0.5 Then
                        If bMalhotra = True Then
                            sMalhotra = "<sup>1 </sup>"
                        End If
                        If numIndicator > 2 Then
                            tempMessage = sMalhotra & "Convergent Validity: the AVE for " & latentVariables(y) & " is less than 0.50. Try removing " & indicatorName & " to improve AVE."
                            If Not validityMessages.Contains(tempMessage) Then
                                validityMessages.Add(tempMessage)
                            End If
                        Else
                            tempMessage = sMalhotra & "Convergent Validity: the AVE for " & latentVariables(y) & " is less than 0.50. No way to improve AVE because you only have two indicators for that variable. Removing one indicator will make this not latent."
                            If Not validityMessages.Contains(tempMessage) Then
                                validityMessages.Add(tempMessage)
                            End If
                        End If

                        If estimateMatrix(y, x) < estimateMatrix(y, x + 1) And numCorrelation > 1 Then 'Check if AVE is less than MSV.
                            tempMessage = "Discriminant Validity: the AVE for " & latentVariables(y) & " is less than the MSV."
                            If Not validityMessages.Contains(tempMessage) Then
                                validityMessages.Add(tempMessage)
                            End If
                        End If
                    End If
                End If

                'Only check latent variables that are correlated
                For i = 1 To numCorrelation
                    If latentVariables(y) = MatrixName(tableCorrelation, i, 0) Then
                        If Math.Sqrt(estimateMatrix(y, x)) < MatrixElement(tableCorrelation, i, 3) Then 'Check if the square root of AVE is less than the correlation.
                            tempMessage = "Discriminant Validity: the square root of the AVE for " & latentVariables(y) & " is less than its correlation with " & MatrixName(tableCorrelation, i, 2) & "."
                            If Not validityMessages.Contains(tempMessage) Then
                                validityMessages.Add(tempMessage)
                            End If
                        End If
                    ElseIf latentVariables(y) = MatrixName(tableCorrelation, i, 2) Then
                        If Math.Sqrt(estimateMatrix(y, x)) < MatrixElement(tableCorrelation, i, 3) Then
                            tempMessage = "Discriminant Validity: the square root of the AVE for " & latentVariables(y) & " is less than its correlation with " & MatrixName(tableCorrelation, i, 0) & "."
                            If Not validityMessages.Contains(tempMessage) Then
                                validityMessages.Add(tempMessage)
                            End If
                        End If
                    End If
                Next
            Next
        Next

        Return validityMessages

    End Function

    'Finds all indicators given a single latent variable
    Function GetIndicators(latent As String) As ArrayList

        Dim indicatorVariables As New ArrayList
        Dim tableIndicators As XmlElement = GetXML("body/div/div[@ntype='models']/div[@ntype='model'][position() = 1]/div[@ntype='group'][position() = 1]/div[@ntype='estimates']/div[@ntype='scalars']/div[@nodecaption='Standardized Regression Weights:']/table/tbody")

        'Loops through variables in the model and stores the latent variable names in an array.
        Dim numIndicators As Integer = GetNodeCount(tableIndicators)

        'Higher Order: Grabs Information About Latent Variable Related To Other Latent Variables
        Dim RelatedLatentVariables As New ArrayList
        RelatedLatentVariables.Add(latent)

        If (HigherOrderVariables.Count > 0) Then
            For Each variable As PDElement In pd.PDElements
                If variable.IsPath AndAlso variable.Variable1.NameOrCaption = latent Then
                    RelatedLatentVariables.Add(variable.Variable2.NameOrCaption)
                End If
            Next
        End If

        For x = 1 To numIndicators
            If RelatedLatentVariables.Contains(MatrixName(tableIndicators, x, 2)) AndAlso Not HigherOrderVariables.Contains(MatrixName(tableIndicators, x, 2)) Then
                'MsgBox(latent + ": " + MatrixName(tableIndicators, x, 0))
                indicatorVariables.Add(MatrixName(tableIndicators, x, 0))
            End If

        Next

        Return indicatorVariables

    End Function

    'Get the CR, AVE, MSV, MaxR, and Sqrt for a latent variable.
    Function GetEstimates(latent As String) As Estimates

        'Store the standardized regression output table into a matrix.
        Dim tableRegression As XmlElement = GetXML("body/div/div[@ntype='models']/div[@ntype='model'][position() = 1]/div[@ntype='group'][position() = 1]/div[@ntype='estimates']/div[@ntype='scalars']/div[@nodecaption='Standardized Regression Weights:']/table/tbody")
        Dim tableRegressionCI As XmlElement = GetXML("body/div/div[@ntype='models']/div[@ntype='model'][position() = 1]/div[@ntype='group'][position() = 1]/div[@ntype='bootstrap']/div[@ntype='bootstrapconfidence']/div[@ntype='percentile']/div[@ntype='scalars']/div[@nodecaption='Standardized Regression Weights:']/table/tbody")

        'Count of elements in the regression table.
        Dim numRegression As Integer = GetNodeCount(tableRegression)

        Dim estimates As Estimates
        Dim MaxR As Double = 0 'Sum of squared SRW over one minus SRW squared
        Dim SRW As Double = 0 'Standardized Regression Weight
        Dim SSI As Double = 0 'Sum of the square SRW values for a variable
        Dim SL2 As Double = 0 'Squared SRW value
        Dim SL3 As Double = 0 'Sum of the SRW values for a variable
        Dim SRWError As Double = 0 'Sum of SRW errors
        Dim numSRW As Double = 0 'Count of SRW values

        'Lower Values
        Dim LSRW As Double = 0 'Standardized Regression Weight
        Dim LSSI As Double = 0 'Sum of the square SRW values for a variable
        Dim LSL2 As Double = 0 'Squared SRW value
        Dim LSL3 As Double = 0 'Sum of the SRW values for a variable
        Dim LSRWError As Double = 0 'Sum of SRW errors

        'Upper Values
        Dim USRW As Double = 0 'Standardized Regression Weight
        Dim USSI As Double = 0 'Sum of the square SRW values for a variable
        Dim USL2 As Double = 0 'Squared SRW value
        Dim USL3 As Double = 0 'Sum of the SRW values for a variable
        Dim USRWError As Double = 0 'Sum of SRW errors

        'Higher Order: Grabs Information About Latent Variable Related To Other Latent Variables
        Dim RelatedLatentVariables As New ArrayList
        RelatedLatentVariables.Add(latent)

        'Loop through regression matrix to calculate estimates.
        For i = 1 To numRegression
            If RelatedLatentVariables.Contains(MatrixName(tableRegression, i, 2)) Then
                'Normal
                SRW = MatrixElement(tableRegression, i, 3)
                SL2 = Math.Pow(SRW, 2)
                SL3 = SL3 + SRW
                SSI = SSI + SL2
                SRWError = SRWError + (1 - SL2)
                MaxR = MaxR + (SL2 / (1 - SL2))
                numSRW += 1
                'Lower
                LSRW = MatrixElement(tableRegressionCI, i, 4)
                LSL2 = Math.Pow(LSRW, 2)
                LSL3 = LSL3 + LSRW
                LSSI = LSSI + LSL2
                LSRWError = LSRWError + (1 - LSL2)
                'Upper
                USRW = MatrixElement(tableRegressionCI, i, 5)
                USL2 = Math.Pow(USRW, 2)
                USL3 = USL3 + USRW
                USSI = USSI + USL2
                USRWError = USRWError + (1 - USL2)

                'MsgBox(MatrixName(tableRegression, i, 0) + " " + MatrixName(tableRegression, i, 2) + " " + SRW.ToString("#.###" + " " + LSRW.ToString("#.###") + " " + USRW.ToString("#.###")))
            End If
        Next

        'Output estimates to struct.
        SL3 = Math.Pow(SL3, 2)
        estimates.CR = SL3 / (SL3 + SRWError)
        estimates.AVE = SSI / numSRW
        estimates.MSV = GetMSV(latent)
        estimates.MaxR = 1 / (1 + (1 / MaxR))
        estimates.SQRT = Math.Sqrt(SSI / numSRW)

        'Lower + Upper
        LSL3 = Math.Pow(LSL3, 2)
        USL3 = Math.Pow(USL3, 2)
        estimates.LCR = LSL3 / (LSL3 + LSRWError)
        estimates.UCR = USL3 / (USL3 + USRWError)
        estimates.LAVE = LSSI / numSRW
        estimates.UAVE = USSI / numSRW

        Return estimates

    End Function

    'Matrix that holds the estimates for each latent variable.
    Function GetMatrix() As Double(,)

        'Holds the correlation matrix from the output table.
        Dim tableCorrelation As XmlElement = GetXML("body/div/div[@ntype='models']/div[@ntype='model'][position() = 1]/div[@ntype='group'][position() = 1]/div[@ntype='estimates']/div[@ntype='scalars']/div[@nodecaption='Correlations:']/table/tbody")
        'Count of elements in the correlation table
        Dim numCorrelation As Integer = GetNodeCount(tableCorrelation)

        'Get an array with the names of the latent variables.
        Dim latentVariables As ArrayList = GetLatent()
        Dim numLatentVariables As Integer = latentVariables.Count

        'Declare a two dimensional array to hole the estimates.
        Dim estimateMatrix((numLatentVariables * 2 - 1), (latentVariables.Count + 3)) As Double

        'Catch if there is only one latent in the model.
        CheckSingleLatent(latentVariables.Count)

        For Each latent In latentVariables

            Dim estimates As Estimates = GetEstimates(latent)
            Dim matrixColumn As Integer = 1
            Dim matrixRow As Integer = latentVariables.IndexOf(latent)

            'First four rows are estimates
            estimateMatrix(matrixRow, 0) = estimates.CR
            estimateMatrix(matrixRow, matrixColumn) = estimates.AVE
            matrixColumn += 1
            If numCorrelation >= 1 Then
                estimateMatrix(matrixRow, matrixColumn) = estimates.MSV
                matrixColumn += 1
            End If
            estimateMatrix(matrixRow, matrixColumn) = estimates.MaxR

            'Get the Correlation table and put it into an aligned matrix.
            If numCorrelation <> 0 Then
                For i = 1 To numCorrelation
                    If latent = MatrixName(tableCorrelation, i, 2) Then
                        For index = 0 To latentVariables.Count - 1
                            If latentVariables(index) = MatrixName(tableCorrelation, i, 0) Then
                                'Catches an output table exception where the variable is only listed on one side.
                                If (index + 4) > latentVariables.IndexOf(latent) + 4 And MatrixElement(tableCorrelation, i, 3) > 0 Then
                                    estimateMatrix(index, matrixRow + 4) = MatrixElement(tableCorrelation, i, 3)
                                    estimateMatrix(matrixRow, index + 4) = MatrixElement(tableCorrelation, i, 3)
                                Else
                                    estimateMatrix(index, matrixRow + 4) = MatrixElement(tableCorrelation, i, 3)
                                    estimateMatrix(matrixRow, index + 4) = MatrixElement(tableCorrelation, i, 3)
                                End If
                            End If
                        Next
                        If latent = latentVariables(matrixRow) Then
                            estimateMatrix(matrixRow, matrixRow + 4) = estimates.SQRT
                        End If
                    ElseIf latent = latentVariables(matrixRow) Then
                        estimateMatrix(matrixRow, matrixRow + 4) = estimates.SQRT
                    End If
                Next
            End If

            'Store CI Numbers
            estimateMatrix(matrixRow + numLatentVariables, 0) = estimates.CR
            estimateMatrix(matrixRow + numLatentVariables, 1) = estimates.AVE
            estimateMatrix(matrixRow + numLatentVariables, 2) = estimates.LCR
            estimateMatrix(matrixRow + numLatentVariables, 3) = estimates.UCR
            estimateMatrix(matrixRow + numLatentVariables, 4) = estimates.LAVE
            estimateMatrix(matrixRow + numLatentVariables, 5) = estimates.UAVE
        Next

        Return estimateMatrix

    End Function

#Region "Helper Functions"

    'Check if the model has a single latent factor.
    Sub CheckSingleLatent(iLatent As Integer)

        'Assign the variance estimates to a matrix
        Dim tableVariance As XmlElement = GetXML("body/div/div[@ntype='models']/div[@ntype='model'][position() = 1]/div[@ntype='group'][position() = 1]/div[@ntype='estimates']/div[@ntype='scalars']/div[@nodecaption='Variances:']/table/tbody")

        'Loop through variance table with the number of latents
        For i = 1 To iLatent + 1
            Try
                MatrixName(tableVariance, i, 3) 'Checks if there is no variance for an unobserved variable.
            Catch ex As Exception
                For k = 1 To iLatent + 1
                    Try
                        If Not MatrixName(tableVariance, k, 3) = Nothing Then 'Checks if the variable is "unidentified"
                            MsgBox(MatrixName(tableVariance, k, 0) + " is causing an error. Either:
                            1. It only has one indicator and is therefore, not latent. The CFA is only for latent variables, so don’t include " + MatrixName(tableVariance, k, 0) + " in the CFA.
                            2. You are missing a constraint on an indicator.")
                        End If
                    Catch exc As NullReferenceException
                        Continue For
                    End Try
                Next
            End Try
        Next

    End Sub

    'Create an arraylist of the names of the latent variables.
    Function GetLatent() As ArrayList

        Dim latentVariables As New ArrayList

        'Loops through variables in the model and stores the latent variable names in an array.
        For Each variable As PDElement In pd.PDElements
            If variable.IsLatentVariable And Not variable.IsEndogenousVariable Then
                latentVariables.Add(variable.NameOrCaption)
            End If
        Next

        Return latentVariables

    End Function

    'Calculate the MSV for a correlation.
    Function GetMSV(latent As String) As Double

        'Store correlation table in a matrix
        Dim tableCorrelation As XmlElement = GetXML("body/div/div[@ntype='models']/div[@ntype='model'][position() = 1]/div[@ntype='group'][position() = 1]/div[@ntype='estimates']/div[@ntype='scalars']/div[@nodecaption='Correlations:']/table/tbody")
        'Count of elements in correlation table
        Dim numCorrelation As Integer = GetNodeCount(tableCorrelation)

        'Variables
        Dim MSV As Double = 0
        Dim testMSV As Double = 0

        'Takes the max correlation as the MSV
        For i = 1 To numCorrelation
            If latent = MatrixName(tableCorrelation, i, 0) Or latent = MatrixName(tableCorrelation, i, 2) Then
                testMSV = Math.Pow(MatrixElement(tableCorrelation, i, 3), 2)
                If testMSV > MSV Then
                    MSV = testMSV
                End If
            End If
        Next

        Return MSV

    End Function

    'Get the number of rows in an xml table.
    Function GetNodeCount(table As XmlElement) As Integer

        Dim nodeCount As Integer = 0

        'Handles a model with zero correlations
        Try
            nodeCount = table.ChildNodes.Count
        Catch ex As NullReferenceException
            nodeCount = 0
        End Try

        Return nodeCount

    End Function

    'Use an output table path to get the xml version of the table.
    Function GetXML(path As String) As XmlElement

        'Gets the xpath expression for an output table.
        Dim doc As Xml.XmlDocument = New Xml.XmlDocument()
        doc.Load(Amos.pd.ProjectName & ".AmosOutput")
        Dim nsmgr As XmlNamespaceManager = New XmlNamespaceManager(doc.NameTable)
        Dim eRoot As Xml.XmlElement = doc.DocumentElement

        Return eRoot.SelectSingleNode(path, nsmgr)

    End Function

    'Get a string element from an xml table.
    Function MatrixName(eTableBody As XmlElement, row As Long, column As Long) As String

        Dim e As XmlElement

        Try
            e = eTableBody.ChildNodes(row - 1).ChildNodes(column) 'This means that the rows are not 0 based.
            MatrixName = e.InnerText
        Catch ex As Exception
            MatrixName = ""
        End Try

    End Function

    'Get a number from an xml table
    Function MatrixElement(eTableBody As XmlElement, row As Long, column As Long) As Double

        Dim e As XmlElement

        Try
            e = eTableBody.ChildNodes(row - 1).ChildNodes(column) 'This means that the rows are not 0 based.
            MatrixElement = CDbl(e.GetAttribute("x"))
        Catch ex As Exception
            MatrixElement = 0
        End Try

    End Function

    Function GetHigherOrderLatent() As ArrayList

        Dim HigherOrderVariables As New ArrayList

        For Each variable As PDElement In pd.PDElements
            If variable.IsPath AndAlso variable.Variable1.IsExogenousVariable AndAlso variable.Variable2.IsLatentVariable Then
                HigherOrderVariables.Add(variable.Variable1.NameOrCaption)
            End If
        Next

        GetHigherOrderLatent = HigherOrderVariables

    End Function

    Function test()
        Dim tableRegressionCI As XmlElement = GetXML("body/div/div[@ntype='models']/div[@ntype='model'][position() = 1]/div[@ntype='group'][position() = 1]/div[@ntype='bootstrap']/div[@ntype='bootstrapconfidence']/div[@ntype='percentile']/div[@ntype='matrices']/div[@ntype='matrix'][position() = 2]/div[@ntype='ppml'][position() = 1]/table/tbody")

        Dim rowHeader As String
        Dim colHeader As String
        Dim corrValue As Double
        Dim smpCorrelation As New Dictionary(Of String, Double)

        'Store Sample Moment's Sample Correlations Matrix In Dictionary (THIS PART IS IN TERMS OF INDICATORS)
        For i = 1 To GetNodeCount(tableRegressionCI)
            For j = 1 To i
                'rowHeader = Left Side of Matrix | colHeader = Top Side of Matridx | corrValue = Body of Matrix
                rowHeader = MatrixName(tableRegressionCI, i, 0)
                colHeader = MatrixName(tableRegressionCI, j, 0)
                corrValue = MatrixElement(tableRegressionCI, i, j) 'Sample(,) Array Is Zero-Based

                MsgBox(rowHeader + " " + colHeader + " " + corrValue.ToString("#.###"))
                On Error Resume Next 'ONLY ERROR THAT OCCURS IS WHEN WE ARE ADDING THE SAME INDICATOR PAIR (I.E., COMPUSE_1 & COMPUSE_1)
                smpCorrelation.Add(colHeader + rowHeader, corrValue)
            Next
        Next
    End Function

#End Region

End Class

#End Region
