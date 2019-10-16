using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Windows;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using SldWorks;

namespace MaterialDefinitionLibrary
{
    public static class Convert
    {
        #region Get list of Materials

        public static List<Material> GetListOfMaterialsInDocument(ModelDoc2 swModel)
        {
            string nameModel = swModel.GetTitle();
            List<Material> listMaterials = new List<Material>();
            try
            {
                // Check if ASM or PRT
                if (swModel.GetType() == 2)
                {
                    ConfigurationManager swConfMgr;
                    Configuration swConf;
                    Component2 rootComponent;

                    swConfMgr = (ConfigurationManager)swModel.ConfigurationManager;
                    swConf = (Configuration)swConfMgr.ActiveConfiguration;
                    rootComponent = (Component2)swConf.GetRootComponent3(true);

                    listMaterials = new List<Material>();

                    Material mainASMMaterial = new Material(swModel, nameModel);

                    if (mainASMMaterial.Name == "Default Material")
                    {
                        LoopThroughAssembly(rootComponent, listMaterials, nameModel, mainASMMaterial, false);
                    }
                    else
                    {
                        LoopThroughAssembly(rootComponent, listMaterials, nameModel, mainASMMaterial, true);
                    }
                }
                else if (swModel.GetType() == 1)
                {
                    Material mainASMMaterial = new Material(swModel, nameModel);
                    listMaterials.Add(mainASMMaterial);
                }

                return listMaterials;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Open Part or Assembly\n" + ex.Message);
                return null;
            }

        }

        public static void LoopThroughAssembly(Component2 rootComponent, List<Material> listMaterials, string nameParent, Material parentMaterial, bool childrenHaveSameMaterial = false)
        {
            object[] arrayChildComponent;
            Component2 childComponent;

            arrayChildComponent = (object[])rootComponent.GetChildren();
            for (long i = 0; i < arrayChildComponent.Length; i++)
            {
                childComponent = (Component2)arrayChildComponent[i];
                ModelDoc2 childCompModel = childComponent.GetModelDoc2();
                string nameChild = nameParent;
                //nameChild += " - " + childCompModel.GetTitle();
                nameChild = nameParent + "_" + childComponent.Name2;
                nameChild = nameChild.Replace("/", "_");

                if (childrenHaveSameMaterial)
                {
                    Material sameMaterial = parentMaterial.CopyMaterial();
                    sameMaterial.PartName = nameChild;

                    // If child is ASM, recursively loop through it
                    if (childCompModel.GetType() == 2)
                    {
                        LoopThroughAssembly(childComponent, listMaterials, nameParent, sameMaterial, childrenHaveSameMaterial);
                        goto EndForLoop;
                    }
                    else
                    {
                        listMaterials.Add(sameMaterial);
                        goto EndForLoop;
                    }
                }

                double[] matModel = (double[])childCompModel.MaterialPropertyValues;
                double[] matComp = (double[])childComponent.MaterialPropertyValues;

                if (matComp != null)
                {
                    Material compMaterial = new Material(childComponent, nameChild);

                    // If child is ASM, recursively loop through it
                    if (childCompModel.GetType() == 2)
                    {
                        childrenHaveSameMaterial = true;
                        LoopThroughAssembly(childComponent, listMaterials, nameParent, compMaterial, childrenHaveSameMaterial);
                        goto EndForLoop;
                    }
                    else
                    {
                        listMaterials.Add(compMaterial);
                        goto EndForLoop;
                    }
                }
                else if (matModel != null)
                {
                    Material modelMaterial = new Material(childCompModel, nameChild);

                    if (modelMaterial.Name != "Default Material")
                    {
                        // If child is ASM, recursively loop through it
                        if (childCompModel.GetType() == 2)
                        {
                            childrenHaveSameMaterial = true;
                            LoopThroughAssembly(childComponent, listMaterials, nameParent, modelMaterial, childrenHaveSameMaterial);
                            goto EndForLoop;
                        }
                        else
                        {
                            listMaterials.Add(modelMaterial);
                            goto EndForLoop;
                        }
                    }
                    else
                    {
                        // If child is ASM, recursively loop through it
                        if (childCompModel.GetType() == 2)
                        {
                            LoopThroughAssembly(childComponent, listMaterials, nameParent, modelMaterial, childrenHaveSameMaterial);
                        }
                        else
                        {
                            listMaterials.Add(modelMaterial);
                        }
                    }
                }
                else
                {
                    Material defaultMaterial = new Material(childCompModel, "Default Material");
                    listMaterials.Add(defaultMaterial);

                    // If child is ASM, recursively loop through it
                    if (childCompModel.GetType() == 2)
                    {
                        LoopThroughAssembly(childComponent, listMaterials, nameParent, defaultMaterial, childrenHaveSameMaterial);
                    }
                }

                EndForLoop:
                continue;
            }
        }

        #endregion

        #region Combine OBJ and MTL

        public static void WriteMTLIntoOBJ(List<Material> allMaterials, string mtlLibPath, string objPath)
        {
            string line;
            string geometryPartName = "";
            List<string> allLinesOBJ = new List<string>();

            int lineCounter = 0;
            // Read the file and display it line by line.  
            using (StreamReader sr = new StreamReader(objPath))
            {
                while ((line = sr.ReadLine()) != null)
                {
                    // Get the first two and seven characters of line
                    string firstTwo = line != null ? string.Join("", line.Take(2)) : null;
                    string firstSeven = line != null ? string.Join("", line.Take(7)) : null;

                    if (firstSeven == "mtllib ")
                    {
                        mtlLibPath = mtlLibPath.Replace(@"\", "/");
                        line = "mtllib " + mtlLibPath.Split('/').Last();
                        allLinesOBJ.Add(line);
                        continue;
                    }
                    else if (firstTwo == "o ")
                    {
                        // change the OBJ name of the geometry to fit the name of the material
                        geometryPartName = line.Substring(2);

                        // replace underscores
                        geometryPartName = geometryPartName.Replace("_-_", "_");

                        allLinesOBJ.Add(line);
                        continue;
                    }
                    else if (firstSeven == "usemtl ")
                    {
                        bool foundMaterial = false;
                        // change line
                        // loop through listOfAllMaterials
                        foreach (Material mat in allMaterials)
                        {
                            Debug.Print(geometryPartName + "         " + mat.PartName);

                            if (geometryPartName == mat.PartName)
                            {
                                Debug.Print(geometryPartName + "         " + mat.PartName);
                                foundMaterial = true;
                                line = "usemtl " + mat.Name;
                                allLinesOBJ.Add(line);
                                break;
                            }
                        }

                        if (!foundMaterial)
                        {
                            line = "usemtl " + "Default Material";
                            allLinesOBJ.Add(line);
                        }
                    }
                    else
                    {
                        allLinesOBJ.Add(line);
                    }
                    lineCounter += 1;
                }
            }

            // Write OBJ into OBJ file
            using (StreamWriter sw = new StreamWriter(objPath))
            {
                foreach (var stringLine in allLinesOBJ)
                {
                    sw.WriteLine(stringLine);
                }
            }
        }

        #endregion

        #region Convert Materials with mapping-table

        public static bool MaterialAlreadyInExcel(Material material, string xlConnectionString)
        {
            string thisMaterialName = material.Name;
            int rValue = (int)(material.R * 255);
            int gValue = (int)(material.G * 255);
            int bValue = (int)(material.B * 255);
            string thisMaterialRGBValue = rValue.ToString() + " " + gValue.ToString() + " " + bValue.ToString();

            DataSet xlDataSet = new DataSet();

            // Connect to Excel via oleDB
            using (OleDbConnection xlConn = new OleDbConnection(xlConnectionString))
            {
                xlConn.Open();
                OleDbDataAdapter objDA = new System.Data.OleDb.OleDbDataAdapter("select * from [mappingTable$]", xlConn);
                objDA.Fill(xlDataSet);
            }

            // Loop through all materials
            // Check if Materialname and RGB-value are the same with one material
            foreach (DataTable xlTable in xlDataSet.Tables)
            {
                foreach (DataRow dataRow in xlTable.Rows)
                {
                    string xlRGBValue = dataRow[xlTable.Columns["RGB-Wert"]].ToString();
                    string xlMaterialName = dataRow[xlTable.Columns["SolidWorks-Name"]].ToString();

                    // Check if RGB-value in DataSet from Excel
                    int isTheSameRGBValue = string.Compare(thisMaterialRGBValue, xlRGBValue);

                    // Check if material name in DataSet from Excel
                    int isTheSameName = string.Compare(thisMaterialName, xlMaterialName);

                    if (isTheSameRGBValue == 0 && isTheSameName == 0)
                    {
                        return true;
                    }
                }
            }
            return false;
        }

        public static bool WriteRGBAndNameIntoExcel(Material material, string xlConnectionStringRead, string xlConnectionStringWrite)
        {
            int lastRow = 0;
            if (!MaterialAlreadyInExcel(material, xlConnectionStringRead))
            {

                // Find out lastRow
                using (OleDbConnection xlConnectionRead = new OleDbConnection(xlConnectionStringRead))
                {
                    try
                    {
                        xlConnectionRead.Open();
                        OleDbCommand readData = new OleDbCommand
                        {
                            Connection = xlConnectionRead,
                            CommandText = "select count(*) from [mappingTable$]"
                        };
                        lastRow = (int)readData.ExecuteScalar();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                        return false;
                    }
                    finally
                    {
                        xlConnectionRead.Close();
                        xlConnectionRead.Dispose();
                    }
                }

                // Write material into Excel table
                using (OleDbConnection xlConnectionWrite = new OleDbConnection(xlConnectionStringWrite))
                {
                    try
                    {
                        // Get name and RGB-value of material
                        string thisMaterialName = material.Name;
                        int rValue = (int)(material.R * 255);
                        int gValue = (int)(material.G * 255);
                        int bValue = (int)(material.B * 255);
                        string thisMaterialRGBValue = rValue.ToString() + " " + gValue.ToString() + " " + bValue.ToString();

                        xlConnectionWrite.Open();
                        // Write data into first and second column of lastRow + 1
                        OleDbCommand writeRGBValue = new OleDbCommand("INSERT INTO [mappingTable$] ([RGB-Wert],[SolidWorks-Name]) VALUES ('" + thisMaterialRGBValue + "', '" + thisMaterialName + "');", xlConnectionWrite);
                        writeRGBValue.ExecuteNonQuery();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                        return false;
                    }
                    finally
                    {
                        xlConnectionWrite.Close();
                        xlConnectionWrite.Dispose();
                    }
                }
                return true;
            }
            else
            {
                return false;
            }
        }

        public static DataTable GetMaterialPathsFromExcel(List<Material> listMaterials, string xlConnectionString)
        {
            DataTable allMaterialPaths = new DataTable();

            string selectionString = "";
            foreach (Material mat in listMaterials)
            {
                selectionString += "'" + mat.Name + "', ";
            }

            // Connect to XL and get a DataTable filled with all the rows that correspond to the materials in listMaterials, if existing
            // Connect to Excel via oleDB
            using (OleDbConnection xlConnectionRead = new OleDbConnection(xlConnectionString))
            {
                try
                {
                    // Get MDL-FilePath
                    xlConnectionRead.Open();
                    OleDbDataAdapter objDA = new System.Data.OleDb.OleDbDataAdapter("select * from [mappingTable$] where [SolidWorks-Name] in (" + selectionString + ")", xlConnectionRead);
                    objDA.Fill(allMaterialPaths);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
                finally
                {
                    xlConnectionRead.Close();
                    xlConnectionRead.Dispose();
                }
            }

            return allMaterialPaths;
        }

        #endregion
    }
}
