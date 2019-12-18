using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Demo
{
    // WalkingConnections.cs
    // <copyright>Copyright (c) Microsoft Corporation.  All rights reserved.
    // </copyright>
    // <summary>This file contains the implementation of the WalkingConnections
    // class.</summary>  

    /// <summary>This class demonstrates how to parse and display information
    /// about the connectors on a page.</summary>
    public class WalkingConnections
    {

        /// <summary>This method is the class constructor.</summary>
        public WalkingConnections()
        {

            // No initialization is required.
        }

        /// <summary>This method walks the connections of the page passed in and
        /// displays information about the shapes that are connected and the
        /// connectors.</summary>
        /// <param name="pageToWalk">Page for which connections will be walked
        /// </param>
        public void ShowPageConnections(
            Microsoft.Office.Interop.Visio.Page pageToWalk)
        {

            const string DEBUG_MESSAGE =
                "{0} of '{1}' is connected to {2} of '{3}' with Text: '{4}'";

            Microsoft.Office.Interop.Visio.Shape shapeFrom;
            Microsoft.Office.Interop.Visio.Shape shapeTo;
            Microsoft.Office.Interop.Visio.Connects connections;
            Microsoft.Office.Interop.Visio.VisFromParts visioFromData;
            Microsoft.Office.Interop.Visio.VisToParts visioToData;
            int fromIntegerData;
            int toIntegerData;
            string fromStringData;
            string toStringData;


            if (pageToWalk == null)
            {
                return;
            }


            try
            {

                // Get the Connects collection for the page.
                connections = pageToWalk.Connects;

                // Loop through the objects in the Connects collection.
                foreach (Microsoft.Office.Interop.Visio.Connect connection
                    in connections)
                {

                    // Get the From information of the Connect object.
                    shapeFrom = connection.FromSheet;
                    fromIntegerData = connection.FromPart;

                    // Converts the VisFromParts value to a string.
                    // Use fromIntegerData to determine the type of connection.					
                    if (fromIntegerData < (int)Microsoft.Office.Interop.Visio.
                        VisFromParts.visControlPoint)
                    {

                        visioFromData =
                            (Microsoft.Office.Interop.Visio.VisFromParts)
                            fromIntegerData;
                        fromStringData = visioFromData.ToString();
                    }
                    else
                    {

                        // Control points are numbered starting with 1
                        // in the UI but with 0 in the API.
                        fromStringData = "visControlPoint"
                            + Convert.ToString(fromIntegerData
                            - (int)Microsoft.Office.Interop.Visio.
                            VisFromParts.visControlPoint + 1,
                            System.Globalization.CultureInfo.InvariantCulture);
                    }

                    // Get the To information of the Connect object.
                    shapeTo = connection.ToSheet;
                    toIntegerData = connection.ToPart;

                    // Converts the VisToParts value to a string.
                    // Use toIntegerData to determine the type of shape to
                    // which the connector is connected.
                    if (toIntegerData < (int)Microsoft.Office.Interop.Visio.
                        VisToParts.visConnectionPoint)
                    {

                        visioToData =
                            (Microsoft.Office.Interop.Visio.VisToParts)
                            toIntegerData;
                        toStringData = visioToData.ToString();
                    }
                    else
                    {

                        // Connection points are numbered starting with 1
                        // in the UI but with 0 in the API
                        toStringData = "visConnectionPoint"
                            + Convert.ToString(toIntegerData
                            - (int)Microsoft.Office.Interop.Visio.
                            VisToParts.visConnectionPoint + 1,
                            System.Globalization.CultureInfo.InvariantCulture);
                    }

                    // Display the information in the Debug window.
                    System.Diagnostics.Debug.WriteLine(
                        String.Format(System.Globalization.CultureInfo.CurrentCulture,
                        DEBUG_MESSAGE,
                        fromStringData, shapeFrom.NameU,
                        toStringData, shapeTo.NameU, shapeFrom.Text));
                }
            }
            catch (Exception err)
            {
                System.Diagnostics.Debug.WriteLine(err.Message);
                throw;
            }
        }
    }




}
