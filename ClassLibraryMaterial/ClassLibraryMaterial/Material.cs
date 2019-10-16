using SldWorks;
using System;
using System.Linq;

namespace MaterialDefinitionLibrary
{
    public class Material : IEquatable<Material>
    {
        #region "Properties"

        public double R { get; set; }
        public double G { set; get; }
        public double B { get; set; }
        public double Ambient { get; set; }
        public double Diffuse { get; set; }
        public double Specular { get; set; }
        public double Shininess { get; set; }
        public double Transparency { get; set; }
        public double Emission { get; set; }
        public double Tf { get; set; }

        public string Name { get; set; }
        public string PartName { get; set; }

        #endregion

        #region "Constructors"

        public Material(string partName = "null")
        {
            double defaultRGBValue = 0.7529;
            double[] defaultMaterial = { defaultRGBValue, defaultRGBValue, defaultRGBValue, 1, 1, 1, 0.3125, 0, 0 };

            SetMaterialProperties(defaultMaterial);
            PartName = partName;
            Name = "Default Material";
        }

        public Material(ModelDoc2 model, string partName)
        {
            if (model == null)
            {
                throw new ArgumentNullException();
            }

            SetMaterialProperties((double[])model.MaterialPropertyValues);
            PartName = partName;

            // Check if material is default material SW assignes
            double defaultRGBValue = 0.7529;
            double[] defaultMaterial1 = { defaultRGBValue, defaultRGBValue, defaultRGBValue, 1, 1, 1, 0.3125, 0, 0 };
            double[] defaultMaterial2 = { 0.7961, 0, 8235, 0.9372, 1, 1, 1, 0.3125, 0, 0 };
            double[] material = { this.R, this.G, this.B, this.Ambient, this.Diffuse, this.Specular, this.Shininess, this.Transparency, this.Emission };

            if (defaultMaterial1.SequenceEqual(material))
            {
                Name = "Default Material";
            }
            else if (defaultMaterial2.SequenceEqual(material))
            {
                Name = "Default Material";
            }
            else
            {
                Name = GetHashCode().ToString();
            }
        }

        public Material(Component2 component, string partName)
        {
            if (component == null)
            {
                throw new ArgumentNullException();
            }

            SetMaterialProperties((double[])component.MaterialPropertyValues);
            PartName = partName;

            // Check if material is default material SW assignes
            double defaultRGBValue = 0.7529;

            double[] defaultMaterial1 = { defaultRGBValue, defaultRGBValue, defaultRGBValue, 1, 1, 1, 0.3125, 0, 0 };
            double[] defaultMaterial2 = { 0.7961, 0, 8235, 0.9372, 1, 1, 1, 0.3125, 0, 0 };
            double[] material = { this.R, this.G, this.B, this.Ambient, this.Diffuse, this.Specular, this.Shininess, this.Transparency, this.Emission };

            if (defaultMaterial1.SequenceEqual(material))
            {
                Name = "Default Material";
            }
            else if (defaultMaterial2.SequenceEqual(material))
            {
                Name = "Default Material";
            }
            else
            {
                Name = GetHashCode().ToString();
            }
            
        }

        public Material(double[] materialProperties, string partName)
        {
            if (materialProperties == null)
            {
                throw new ArgumentNullException();
            }

            SetMaterialProperties(materialProperties);
            PartName = partName;

            // Check if material is default material SW assignes
            double defaultRGBValue = 0.7529;

            double[] defaultMaterial1 = { defaultRGBValue, defaultRGBValue, defaultRGBValue, 1, 1, 1, 0.3125, 0, 0 };
            double[] defaultMaterial2 = { 0.7961, 0, 8235, 0.9372, 1, 1, 1, 0.3125, 0, 0 };
            double[] material = { this.R, this.G, this.B, this.Ambient, this.Diffuse, this.Specular, this.Shininess, this.Transparency, this.Emission };

            if (defaultMaterial1.SequenceEqual(material))
            {
                Name = "Default Material";
            }
            else if (defaultMaterial2.SequenceEqual(material))
            {
                Name = "Default Material";
            }
            else
            {
                Name = GetHashCode().ToString();
            }
        }

        #endregion

        #region "Functions"

        private void SetMaterialProperties(double[] materialProperties)
        {
            try
            {
                this.R = Math.Round(materialProperties[0], 4);
                this.G = Math.Round(materialProperties[1], 4);
                this.B = Math.Round(materialProperties[2], 4);
                this.Ambient = materialProperties[3];
                this.Diffuse = materialProperties[4];
                this.Specular = materialProperties[5];
                this.Shininess = materialProperties[6];
                this.Transparency = materialProperties[7];
                this.Emission = materialProperties[8];



                return;
            }
            catch
            {
                throw new ArgumentException();
            }
        }

        // Write properties into a string
        public string AppearanceToMTL()
        {
            string mtlString = "";
            mtlString += "newmtl " + Name + "\r\n";
            mtlString += "Ka " + R * Ambient + " " + G * Ambient + " " + B * Ambient + "\r\n";
            mtlString += "Kd " + R * Diffuse + " " + G * Diffuse + " " + B * Diffuse + "\r\n";
            mtlString += "Ks " + R * Specular + " " + G * Specular + " " + B * Specular + "\r\n";

            if (Transparency > 0)
            {
                mtlString += "Tf " + Tf + "\r\n";
                mtlString += "illum 7" + "\r\n";
            }
            else
            {
                mtlString += "Tf 1.0" + "\r\n";
                mtlString += "illum 2" + "\r\n";
            }

            mtlString += "Ns " + (1000 * (1 - Shininess)) + "\r\n";

            if (Transparency == 0)
            {
                if (Name.Length == 0)
                    mtlString += "Ni 1.0" + "\r\n";
                else
                    mtlString += "Ni 1.3" + "\r\n";
            }
            else
            {
                mtlString += "Ni 1.5" + "\r\n";
            }

            return mtlString;
        }

        // Generate a random number out of material properties
        public override int GetHashCode()
        {
            return Math.Abs((Name + "," + R.ToString() + "," + G.ToString() + "," + B.ToString() + "," + Ambient.ToString() + "," +
                Diffuse.ToString() + "," + Specular.ToString() + "," + Shininess.ToString() + "," + Transparency.ToString() + "," + Emission.ToString()).GetHashCode());
        }

        public Material CopyMaterial()
        {
            Material material = new Material();

            material.R = this.R;
            material.G = this.G;
            material.B = this.B;
            material.Ambient = this.Ambient;
            material.Diffuse = this.Diffuse;
            material.Specular = this.Specular;
            material.Shininess = this.Shininess;
            material.Transparency = this.Transparency;
            material.Emission = this.Emission;
            material.Tf = this.Tf;

            material.Name = this.Name;
            material.PartName = "";

            return material;
        }

        // Override Equals()-method, to enable comparing of materials with Distinct()-method
        public bool Equals(Material material)
        {
            if (Name == material.Name)
                return true;

            return false;
        }

        #endregion
    }
}
