using Amos;
using Microsoft.VisualBasic;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using Microsoft.VisualBasic.CompilerServices;

namespace Amos26_HTMTValidity
{
    internal class Util
    {

        // Use an output table path to get the xml version of the table.
        public static XmlElement GetXML(string path)
        {

            // Gets the xpath expression for an output table.
            var doc = new XmlDocument();
            doc.Load(pd.ProjectName + ".AmosOutput");
            var nsmgr = new XmlNamespaceManager(doc.NameTable);
            var eRoot = doc.DocumentElement;

            return (XmlElement)eRoot.SelectSingleNode(path, nsmgr);

        }
        // Get the number of rows in an xml table.
        public static int GetNodeCount(XmlElement table)
        {

            int nodeCount = 0;

            // Handles a model with zero correlations
            try
            {
                nodeCount = table.ChildNodes.Count;
            }
            catch (NullReferenceException ex)
            {
                nodeCount = 0;
            }

            return nodeCount;

        }


        // Get a string element from an xml table.
        public static string MatrixName(XmlElement eTableBody, long row, long column)
        {
            string MatrixNameRet = default;

            XmlElement e;

            try
            {
                e = (XmlElement)eTableBody.ChildNodes[(int)(row - 1L)].ChildNodes[(int)column]; // This means that the rows are not 0 based.
                MatrixNameRet = e.InnerText;
            }
            catch (Exception ex)
            {
                MatrixNameRet = "";
            }

            return MatrixNameRet;

        }

        // Get a number from an xml table
        public static double MatrixElement(XmlElement eTableBody, long row, long column)
        {
            double MatrixElementRet = default;

            XmlElement e;

            try
            {
                e = (XmlElement)eTableBody.ChildNodes[(int)(row - 1L)].ChildNodes[(int)column]; // This means that the rows are not 0 based.
                MatrixElementRet = Conversions.ToDouble(e.GetAttribute("x"));
            }
            catch (Exception ex)
            {
                MatrixElementRet = 0d;
            }

            return MatrixElementRet;

        }

        public static ArrayList GetHigherOrderLatent()
        {
            ArrayList GetHigherOrderLatentRet = default;

            var HigherOrderVariables = new ArrayList();

            foreach (PDElement variable in pd.PDElements)
            {
                if (variable.IsPath() && variable.Variable1.IsExogenousVariable() && variable.Variable2.IsLatentVariable())
                {
                    HigherOrderVariables.Add(variable.Variable1.NameOrCaption);
                }
            }

            GetHigherOrderLatentRet = HigherOrderVariables;
            return GetHigherOrderLatentRet;

        }

        public static object test()
        {
            var tableRegressionCI = Util.GetXML("body/div/div[@ntype='models']/div[@ntype='model'][position() = 1]/div[@ntype='group'][position() = 1]/div[@ntype='bootstrap']/div[@ntype='bootstrapconfidence']/div[@ntype='percentile']/div[@ntype='matrices']/div[@ntype='matrix'][position() = 2]/div[@ntype='ppml'][position() = 1]/table/tbody");

            string rowHeader;
            string colHeader;
            double corrValue;
            var smpCorrelation = new Dictionary<string, double>();

            // Store Sample Moment's Sample Correlations Matrix In Dictionary (THIS PART IS IN TERMS OF INDICATORS)
            for (int i = 1, loopTo = GetNodeCount(tableRegressionCI); i <= loopTo; i++)
            {
                for (int j = 1, loopTo1 = i; j <= loopTo1; j++)
                {
                    // rowHeader = Left Side of Matrix | colHeader = Top Side of Matridx | corrValue = Body of Matrix
                    rowHeader = MatrixName(tableRegressionCI, i, 0L);
                    colHeader = MatrixName(tableRegressionCI, j, 0L);
                    corrValue = MatrixElement(tableRegressionCI, i, j); // Sample(,) Array Is Zero-Based

                    Interaction.MsgBox(rowHeader + " " + colHeader + " " + corrValue.ToString("#.###"));
                    // Check to make sure we haven't added the same ones, etc...
                    if (colHeader + rowHeader != rowHeader + colHeader)
                    {
                        smpCorrelation.Add(colHeader + rowHeader, corrValue);
                    }
                }
            }

            return default;
        }
    }
}
