using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Forms;

namespace SaveAllBodiesAsSTEP
{
    public partial class SaveAllBodiesAsSTEP
    {
        // The SldWorks swApp variable is pre-assigned for you.
        public SldWorks swApp;
        public ModelDoc2 swModel;
        private const int Visible = (int)swComponentVisibilityState_e.swComponentVisible;
        private const int NonVisible = (int)swComponentVisibilityState_e.swComponentHidden;
        private const int swThisConfiguration = (int)swInConfigurationOpts_e.swThisConfiguration;
        private const int swSolidBody = (int)swBodyType_e.swSolidBody;

        public void Main()
        {
            swModel = (ModelDoc2)swApp.ActiveDoc;
            var fullFilePath = swModel.GetPathName();
            var fullDirectoryPath = Path.GetDirectoryName(fullFilePath);

            if (IsAssembly_SLDASM(fullFilePath)) // ASSEMBLY
            {
                AssemblyDoc swAssembly = (AssemblyDoc)swModel;
                var visibleComponents = GetAllVisibleComponents(swAssembly);
                var visibleComponentsNames = visibleComponents.Select(c => c.Name2);

                TemporaryHideAllForSaving(visibleComponents);
                foreach (var component in visibleComponents)
                {
                    Debug.Print("Name of component: " + component.Name2);
                    ShowComponent(component);
                    SaveComponent(component, fullDirectoryPath);
                    HideComponent(component);
                }

                RestoreVisibility(visibleComponents);
            }

            if (IsPart_SLDPRT(fullFilePath)) // PART
            {
                PartDoc swPart = (PartDoc)swModel;
                var visibleBodies = ((Object[])swPart.GetBodies2(swSolidBody, true))
                    .Select(b => (Body2)b)
                    .ToList();

                var visibleBodiesNames = visibleBodies.Select(b => b.Name);

                foreach (var body in visibleBodies)
                {
                    Debug.Print("Name of body: " + body.Name);
                    SaveBody(body, fullDirectoryPath);
                }




                //var visibleComponentsNames = new List<string>();
                //visibleBodies.tol.ForEach(c => visibleComponentsNames.Add(c.Name2));
                ;

            }


            return;
        }







        /// <summary>
        /// ///////////////////////////////////
        /// ///////////////////////////////////
        /// ///////////////////////////////////
        /// ///////////////////////////////////
        /// ///////////////////////////////////
        /// ///////////////////////////////////
        /// ///////////////////////////////////
        /// ///////////////////////////////////
        /// ///////////////////////////////////
        /// ///////////////////////////////////
        /// ///////////////////////////////////
        /// ///////////////////////////////////
        /// ///////////////////////////////////
        /// ///////////////////////////////////
        /// ///////////////////////////////////
        /// ///////////////////////////////////
        /// ///////////////////////////////////
        /// ///////////////////////////////////
        /// ///////////////////////////////////
        /// ///////////////////////////////////
        /// </summary>
        /// <param name="component"></param>

        private void ShowComponent(Component2 component)
        {
            //component.SetVisibility(Visible, swThisConfiguration, null);
            component.Visible = Visible;
        }

        private void HideComponent(Component2 component)
        {
            //component.SetVisibility(NonVisible, swThisConfiguration, null);
            component.Visible = NonVisible;
        }

        private void TemporaryHideAllForSaving(List<Component2> visibleComponents)
        {
            visibleComponents.ForEach(c => HideComponent(c));
        }

        private void RestoreVisibility(List<Component2> visibleComponents)
        {
            visibleComponents.ForEach(c => ShowComponent(c));
        }



        private bool SaveComponent(Component2 component, string fullDirectoryPath)
        {
            var path = Path.Combine(fullDirectoryPath, component.Name2 + ".STEP");
            //result = swModel.SaveAs2(path, 0, true, false);
            //swModel.SaveAs3(path, 0, 2);
            return swModel.SaveAs(path);  //(path, 0, 2);
        }

        private bool SaveBody(Body2 body, string fullDirectoryPath)
        {
            var path = Path.Combine(fullDirectoryPath, body.Name + ".STEP");
            //result = swModel.SaveAs2(path, 0, true, false);
            //swModel.SaveAs3(path, 0, 2);
            return swModel.SaveAs(path);  //(path, 0, 2);
        }

        private List<Component2> GetAllVisibleComponents(AssemblyDoc swAssembly)
        {
            //no use GetVisibleComponentsInView becuase it gets all in tree
            //and havent switch toplevelonly
            //but i need only toplevel
            var allComponents = ((Object[])swAssembly.GetComponents(ToplevelOnly: true))
                .Select(c => (Component2)c)
                .Where(c => c.Visible == Visible)
                .ToList()
                ;

            return allComponents;
        }

        private bool IsAssembly_SLDASM(string fullFilePath)
        {
            return fullFilePath.ToUpper().EndsWith("SLDASM".ToUpper());
        }

        private bool IsPart_SLDPRT(string fullFilePath)
        {
            return fullFilePath.ToUpper().EndsWith("SLDPRT".ToUpper());
        }





























        /// <summary>
        /// //////////////////////////////////////
        /// //////////////////////////////////////
        /// //////////////////////////////////////
        /// //////////////////////////////////////
        /// //////////////////////////////////////
        /// //////////////////////////////////////
        /// //////////////////////////////////////
        /// //////////////////////////////////////
        /// //////////////////////////////////////
        /// //////////////////////////////////////
        /// //////////////////////////////////////
        /// //////////////////////////////////////
        /// //////////////////////////////////////
        /// //////////////////////////////////////
        /// //////////////////////////////////////
        /// //////////////////////////////////////
        /// //////////////////////////////////////
        /// //////////////////////////////////////
        /// //////////////////////////////////////
        /// </summary>

        public void Getallmates()
        {
            ModelDoc2 swModel;
            Component2 swComponent;
            AssemblyDoc swAssembly;
            object[] Components = null;
            object[] Mates = null;
            Mate2 swMate;
            MateInPlace swMateInPlace;
            int numMateEntities = 0;
            int typeOfMate = 0;
            int i = 0;

            swModel = (ModelDoc2)swApp.ActiveDoc;
            swAssembly = (AssemblyDoc)swModel;
            Components = (Object[])swAssembly.GetComponents(false);
            foreach (Object SingleComponent in Components)
            {
                swComponent = (Component2)SingleComponent;
                Debug.Print("Name of component: " + swComponent.Name2);
                Mates = (Object[])swComponent.GetMates();
                if ((Mates != null))
                {
                    foreach (Object SingleMate in Mates)
                    {
                        if (SingleMate is SolidWorks.Interop.sldworks.Mate2)
                        {
                            swMate = (Mate2)SingleMate;
                            typeOfMate = swMate.Type;
                            switch (typeOfMate)
                            {
                                case 0:
                                    Debug.Print(" Mate type: Coincident");
                                    break;
                                case 1:
                                    Debug.Print(" Mate type: Concentric");
                                    break;
                                case 2:
                                    Debug.Print(" Mate type: Perpendicular");
                                    break;
                                case 3:
                                    Debug.Print(" Mate type: Parallel");
                                    break;
                                case 4:
                                    Debug.Print(" Mate type: Tangent");
                                    break;
                                case 5:
                                    Debug.Print(" Mate type: Distance");
                                    break;
                                case 6:
                                    Debug.Print(" Mate type: Angle");
                                    break;
                                case 7:
                                    Debug.Print(" Mate type: Unknown");
                                    break;
                                case 8:
                                    Debug.Print(" Mate type: Symmetric");
                                    break;
                                case 9:
                                    Debug.Print(" Mate type: CAM follower");
                                    break;
                                case 10:
                                    Debug.Print(" Mate type: Gear");
                                    break;
                                case 11:
                                    Debug.Print(" Mate type: Width");
                                    break;
                                case 12:
                                    Debug.Print(" Mate type: Lock to sketch");
                                    break;
                                case 13:
                                    Debug.Print(" Mate type: Rack pinion");
                                    break;
                                case 14:
                                    Debug.Print(" Mate type: Max mates");
                                    break;
                                case 15:
                                    Debug.Print(" Mate type: Path");
                                    break;
                                case 16:
                                    Debug.Print(" Mate type: Lock");
                                    break;
                                case 17:
                                    Debug.Print(" Mate type: Screw");
                                    break;
                                case 18:
                                    Debug.Print(" Mate type: Linear coupler");
                                    break;
                                case 19:
                                    Debug.Print(" Mate type: Universal joint");
                                    break;
                                case 20:
                                    Debug.Print(" Mate type: Coordinate");
                                    break;
                                case 21:
                                    Debug.Print(" Mate type: Slot");
                                    break;
                                case 22:
                                    Debug.Print(" Mate type: Hinge");
                                    // 
                                    break;
                                    // Add new mate types introduced after SOLIDWORKS 2010 FCS here 
                            }
                        }
                        if (SingleMate is SolidWorks.Interop.sldworks.MateInPlace)
                        {
                            swMateInPlace = (MateInPlace)SingleMate;
                            numMateEntities = swMateInPlace.GetMateEntityCount();
                            for (i = 0; i <= numMateEntities - 1; i++)
                            {
                                Debug.Print(" Mate component name: " + swMateInPlace.get_MateComponentName(i));
                                Debug.Print(" Type of Inplace mate entity: " + swMateInPlace.get_MateEntityType(i));
                            }
                            Debug.Print("");
                        }
                    }
                }
                Debug.Print("");
            }
        }



        public void Main2()
        {

            ModelDoc2 swDoc = null;
            PartDoc swPart = null;
            DrawingDoc swDrawing = null;
            AssemblyDoc swAssembly = null;
            bool boolstatus = false;
            int longstatus = 0;
            int longwarnings = 0;
            swDoc = ((ModelDoc2)(swApp.ActiveDoc));
            boolstatus = swDoc.Extension.SelectByRay(-0.024179604907430985D, -0.0051289340948983408D, 0.048000000625847861D, -0.8236958650258458D, -0.189895394735217D, -0.53428911742396534D, 0.00020962654891361061D, 2, false, 0, 0);
            // 
            // Save As
            longstatus = swDoc.SaveAs3("D:\\macrotest-skasuj\\TEST =ASM TTGO T-Beam V1.2 Case TEST  antenna inside.STEP", 0, 2);


            return;
        }

    }
}

