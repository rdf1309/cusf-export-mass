using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System.Runtime.InteropServices;
using System;
using System.Diagnostics;

namespace AddBodiesMassPropertyCSharp.csproj
{

    partial class SolidWorksMacro
    {

        public void Main()
        {

            ModelDoc2 swModelDoc2;
            Component2 swComponent;
            SelectionMgr swSelMgr;
            object bodyInfo = null;
            object[] bodies = null;
            double val = 0;
            double[] parms = null;
            MassProperty swMass;
            bool boolstatus = false;
            int errors = 0;
            int warnings = 0;

            swModelDoc2 = (ModelDoc2)swApp.OpenDoc6("C:\\cued-fs\users\General\rdf33\windows-home\Downloads\v7 - Entire Assembly-20220227T111712Z-001.zip\v7 - Entire Assembly\WHITE_DWARF.SLDASM", (int)swDocumentTypes_e.swDocASSEMBLY, (int)swOpenDocOptions_e.swOpenDocOptions_Silent, "", ref errors, ref warnings);

            swModel = (ModelDoc2)swApp.OpenDoc6(assemblyFile, (int)swDocumentTypes_e.swDocASSEMBLY, (int)swOpenDocOptions_e.swOpenDocOptions_Silent, "", ref errors, ref warnings);
            swModelDocExt = (ModelDocExtension)swModel.Extension;
            swSelMgr = (SelectionMgr)swModel.SelectionManager;
            SelectAllinDocument(swModel, swModelDocExt, swSelMgr);
            swComponent = (Component2)swSelMgr.GetSelectedObject6(1, 0);
            bodies = (object[])swComponent.GetBodies3((int)swBodyType_e.swAllBodies, out (object)bodyInfo);

            swMass = (MassProperty)swModelDoc2.Extension.CreateMassProperty();

            // Convert .NET objects to IDispatch by using DispatchWrapper
            // for call to IMassProperty::AddBodies, whose input objects
            // must be marshaled to Dispatch object arrays
            bodyArray = (DispatchWrapper[])ObjectArrayToDispatchWrapperArray(bodies);

            boolstatus = swMass.AddBodies((bodyArray));
            swMass.UseSystemUnits = false;
            val = swMass.Mass;
            Debug.Print(" Mass - " + val);
            val = swMass.Volume;
            Debug.Print(" Volume - " + val);
            val = swMass.Density;
            Debug.Print(" Density - " + val);
            val = swMass.SurfaceArea;
            Debug.Print(" Surface area - " + val);
            parms = (double[])swMass.CenterOfMass;


            Debug.Print(" Center of mass - X: " + parms[0] + " ,Y: " + parms[1] + " ,and Z: " + parms[2]);
        }

        public DispatchWrapper[] ObjectArrayToDispatchWrapperArray(object[] SwObjects)
        {
            int arraySize = 0;
            arraySize = SwObjects.GetUpperBound(0);
            DispatchWrapper[] dispwrap = new DispatchWrapper[arraySize + 1];
            int arrayIndex = 0;
            for (arrayIndex = 0; arrayIndex <= arraySize; arrayIndex++)
            {
                dispwrap[arrayIndex] = new DispatchWrapper(SwObjects[arrayIndex]);
            }
            return dispwrap;
        }
        public SldWorks swApp;
        public DispatchWrapper[] bodyArray;


    }
}

