using Microsoft.Office.Interop.Visio;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Demo
{
    public class ChangeShape
    {
        /// DemoChangeShape
        /// <summary>This procedure updates a basic flowchart diagram to BPMN.</summary>
        /// <param name="targetPage">Reference to Visio Page object</param>
        public void DemoChangeShape(Microsoft.Office.Interop.Visio.Page targetPage)
        {
            Shapes targetPageShapes;
            Document bpmnStencil;

            Master startMaster;
            Master endMaster;
            Master processMaster;
            Master decisionMaster;
            Master sequenceFlowMaster;

            try
            {
                // Proceed with the active page and the shapes on this page
                targetPageShapes = targetPage.Shapes;

                // Open the BPMN Metric stencil
                bpmnStencil = targetPage.Application.Documents.OpenEx(@"C:\Temp\test1.VSdX",
                    (short)Microsoft.Office.Interop.Visio.VisOpenSaveArgs.visOpenHidden +
                    (short)Microsoft.Office.Interop.Visio.VisOpenSaveArgs.visOpenDocked);

                // Getting masters from stencil
                startMaster = bpmnStencil.Masters.get_ItemU("Start Event");
                endMaster = bpmnStencil.Masters.get_ItemU("End Event");
                processMaster = bpmnStencil.Masters.get_ItemU("Task");
                decisionMaster = bpmnStencil.Masters.get_ItemU("Gateway");
                sequenceFlowMaster = bpmnStencil.Masters.get_ItemU("Sequence Flow");

                // For each shape on the page, check to see if it is either
                // a Dynamic connector, Process, Decision or Start/End shape.
                // For these shapes, we will use the Shape.ReplaceShape method to
                // replace them with the corresponding BPMN shapes.

                foreach (Microsoft.Office.Interop.Visio.Shape currentShape in targetPageShapes)
                {
                    if (currentShape.Master == null)
                        continue;

                    string masterNameU = currentShape.Master.NameU;

                    if (string.Equals(masterNameU, "Dynamic connector", StringComparison.InvariantCultureIgnoreCase))
                    {
                        // currentShape.ReplaceShape(sequenceFlowMaster);
                    }
                    else if (string.Equals(masterNameU, "Process", StringComparison.InvariantCultureIgnoreCase))
                    {
                        // currentShape.ReplaceShape(processMaster);
                    }
                    else if (string.Equals(masterNameU, "Decision", StringComparison.InvariantCultureIgnoreCase))
                    {
                        // currentShape.ReplaceShape(decisionMaster);
                    }
                    else if (string.Equals(masterNameU, "Start/End", StringComparison.InvariantCultureIgnoreCase))
                    {
                        System.Array incomingShapesArray = currentShape.ConnectedShapes(
                            Microsoft.Office.Interop.Visio.VisConnectedShapesFlags.visConnectedShapesIncomingNodes, "");

                        if (incomingShapesArray.Length == 0)
                        {
                            //currentShape.ReplaceShape(startMaster);
                        }
                        else
                        {
                            System.Array outgoingShapesArray = currentShape.ConnectedShapes(
                                Microsoft.Office.Interop.Visio.VisConnectedShapesFlags.visConnectedShapesOutgoingNodes, "");

                            if (outgoingShapesArray.Length == 0)
                            {
                                // currentShape.ReplaceShape(endMaster);
                            }
                        }
                    }
                }
            }
            catch (Exception error)
            {
                System.Diagnostics.Debug.WriteLine(error.Message);
                throw;
            }
        }
    }

}
