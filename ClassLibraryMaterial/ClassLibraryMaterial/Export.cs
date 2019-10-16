using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using SldWorks;
using SwConst;
using System.Xml.Linq;

namespace MaterialDefinitionLibrary
{
    public static class Export
    {
        #region Export material

        public static void WriteMTLToFile(List<Material> listMaterials, string swModelName, string mtlFilePath)
        {
            // Name of MTL-Lib
            string mtlPathName = mtlFilePath + swModelName + ".mtl";

            // write each MTL-Material into MTL-Lib
            try
            {
                using (StreamWriter sw = new StreamWriter(mtlPathName))
                {

                    // Write header
                    sw.WriteLine("# MTL File: '{0}.MTL'", swModelName);
                    sw.WriteLine("# Material Count: {0}", listMaterials.Count);
                    sw.WriteLine("");

                    // Add default material
                    Material defaultMaterial = new Material();
                    listMaterials.Add(defaultMaterial);

                    // Compare materials in listMaterials, and remove the doubles
                    List<Material> listMaterialsNoDoubles = listMaterials.Distinct().ToList<Material>();

                    //Loop through <List>(Material)
                    foreach (Material material in listMaterialsNoDoubles)
                    {
                        string materialInfo = material.AppearanceToMTL();
                        // Change commas to dots
                        materialInfo = materialInfo.Replace(",", ".");
                        sw.WriteLine(material.AppearanceToMTL());
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Exception: " + ex.Message);
                throw;
            }
        }

        #endregion

        #region Export geometry

        public static void SaveToSTL(ModelDoc2 swModel, string filePath)
        {
            swModel.Extension.SaveAs(@"C:\Users\Maxs\Desktop\" + swModel.GetTitle() + ".stl",
                (int)swSaveAsVersion_e.swSaveAsCurrentVersion, (int)swSaveAsOptions_e.swSaveAsOptions_Silent, 1, 1, 1);
        }

        #endregion

        #region Export Material to XML

        public static void WriteMaterialsIntoXML(string xmlPath, List<Material> listMaterials, DataTable listMaterialPaths, string xlConnectionString)
        {
            XDocument doc = XDocument.Load(xmlPath);

            foreach (var part in doc.Root.Element("Parts").Elements())
            {
                foreach (var attribute in part.Element("Attributes").Elements())
                {
                    // matching material and geometry in XML
                    foreach (Material mat in listMaterials)
                    {
                        if (attribute.Element("Value").Value == mat.PartName)
                        {
                            // 
                            foreach (DataRow row in listMaterialPaths.Rows)
                            {
                                if ((string)row.ItemArray[1] == mat.Name)
                                {
                                    for (int countItems = 2; countItems < row.ItemArray.Length; countItems++)
                                    {
                                        string elementType = GetElementType(row.ItemArray[countItems]);
                                        string fileType = GetFileType(row.ItemArray[countItems]);
                                        XElement fileMaterial = new XElement(elementType);
                                        fileMaterial.Add(new XElement("Name", mat.Name));
                                        fileMaterial.Add(new XElement("Path", row.ItemArray[countItems]));
                                        fileMaterial.Add(new XElement("Type", fileType));
                                        part.Element("Files").Add(fileMaterial);
                                    }
                                }
                            }
                        }
                    }
                }
            }
            doc.Save(xmlPath);
        }

        private static string GetElementType(object filePath)
        {
            string fileElement = "file";

            string fileExtension = ((string)filePath).Split('.')[1];
            fileExtension = fileExtension.ToUpper();

            return fileElement += fileExtension;
        }

        private static string GetFileType(object filePath)
        {
            string fileElement;

            string fileExtension = ((string)filePath).Split('.')[1];

            return fileElement = "." + fileExtension; ;
        }

        #endregion
    }
}
