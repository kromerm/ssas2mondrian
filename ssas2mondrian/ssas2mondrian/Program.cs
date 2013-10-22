﻿using System;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.AnalysisServices;
using Microsoft.AnalysisServices.Hosting;

/* Convert simple base SSAS model to Mondrian XML schemas
 * Built October 2013 by Mark Kromer
 * mark_kromer@mail.com
 */

public class ssas2mondrian
{
    // Keep track of SQL Server (sschema) schema for database source
    public static String sschema;

    // Split out the schema and table name from source
    private static string[] splitField(String mystring)
    {
        string[] mysplit = mystring.Split('.');
        string s9 = mysplit[0];        
        int i9 = s9.IndexOf('_');
        if (i9 > 0)
        {
            mysplit[0] = s9.Substring(i9 + 1);            
            sschema = s9.Split('_')[0];
        }
        
        return mysplit;
    }

    // Use this to convert the SSAS Hierarchy types to Mondrian Level Types
    private static String convertLevelType(String sin)
    {
        string leveltype = sin;            
        
        if (sin.Equals("HalfYears"))
        {
            leveltype = "TimeHalfYears";
        }
        else
            if (sin.Equals("HalfYearOfYear"))
            {
                leveltype = "TimeHalfYear";
            }
            else
                if (sin.Contains("Week"))
                {
                    leveltype = "TimeWeeks";
                }
                else
                    if (sin.Contains("Day"))
                    {
                        leveltype = "TimeDays";
                    }
                    else
                    if (sin.Contains("Year"))
                    {
                        leveltype = "TimeYears";
                    }
                    else leveltype = "Regular";

        return leveltype;
    }

    private static String convertDataType (String sin)
    {
        String dt1 = sin;

        if (sin.ToLower().Contains("int"))
        {
            dt1 = "Integer";
        }

        switch (sin)
        {
            case "Currency":
            case "Double":
                dt1 = "Numeric";
                break;
            case "WChar":
                dt1 = "String";
                break;         
        }
            
        return dt1;
    }

    private static void outputhelp()
    {
        Console.WriteLine("required: ssas2mondrian /Sservername /Ccubename /Ddatabasename");
        Console.WriteLine("optional: /M    <--- include many-to-many dims");
        Console.WriteLine("optional: /P    <--- pause at the end of the execution");
        Console.WriteLine("optional: /N    <--- custom name for resulting Mondrian schema");
        Console.WriteLine("optional: /A    <--- include schema attribute in XML for SQL Server source database tables");
        Console.WriteLine("optional: /?    <--- show help");
        Console.WriteLine();
        Console.WriteLine("Sample 1: Set just required values");
        Console.WriteLine("          ssas2mondrian /Slocalhost /C\"Adventure works sales\" /DAdventureWorks");
        Console.WriteLine("Sample 2: Set name of Mondrian schema");
        Console.WriteLine("          ssas2mondrian /Slocalhost /C\"Adventure works sales\" /DAdventureWorks /Nmyschema");
        Console.WriteLine("Sample 3: Include M:M Dims and output the SQL Server database source schema in Table defs");
        Console.WriteLine("          ssas2mondrian /Slocalhost /C\"Adventure works sales\" /DAdventureWorks /M /A");
    }

    static void Main(string[] args)
    {

        // Hold the Calculated Member info here
        String cms = null;

        //Connection vars
        String ConnStr=null;
        String OLAPServerName=null;
        String OLAPDB=null;
        String OLAPCube=null;
        bool shallipause=false; // do u want me to wait at the end of the run?
        bool includeschema = false; // shall i include the SQL Server schema names in the Schema XML syntax output?
        bool boolM2M = true; // Include Many-to-many dimensions?
        String schemaname = null;

        // process command line switches
        try
        {
            foreach (string arg in args)
            {
                switch (arg.Substring(0, 2).ToUpper())
                {
                    case "/S": //Server name REQUIRED
                        OLAPServerName = arg.Substring(2);
                        break;
                    case "/C": //Cube Name REQUIRED
                        OLAPCube = arg.Substring(2);
                        break;
                    case "/D": // source SSAS Database name where the cube is located REQUIRED
                        OLAPDB = arg.Substring(2);
                        break;
                    case "/M": // include M:M dims OPTIONAL
                        boolM2M = true;
                        break;
                    case "/N": // what do you want to call the resulting Mondrian schema? OPTIONAL
                        schemaname = arg.Substring(2);
                        break;
                    case "/P": // do you want to pause at the end of the migration tool run? OPTIONAL
                        shallipause = true;
                        break;
                    case "/A": // shall I include the SQL Server schema from the source DB? OPTIONAL
                        includeschema = true;
                        break;
                    case "/?":
                    case "/H":
                    case "/HELP":
                        outputhelp();
                        break;
                    default:
                        break;
                }
            }
        }
        catch (Exception e)
        {
            Console.WriteLine("Sorry, the command line is very simplistic for this small program and requires no spaces between switch and value.");
            outputhelp();
        }

        // Set the output Mondrian schema name here
        if (schemaname == null) schemaname = OLAPDB;
        Server OLAPServer = new Server();
        ConnStr = "Provider=MSOLAP;Data Source=" + OLAPServerName + ";";            

        try
        {   
            OLAPServer.Connect(ConnStr);
        }
        catch (Exception e)
        {
            Console.WriteLine("Problem connecting to SSAS");
            Console.WriteLine(e.Message);
            Console.WriteLine();
            outputhelp();
            return;
        }

        // For Virtual Cubes, keep track of dims, meas and cubes
        List<String> Cubes = new List<String>();        
        Dictionary<String, List<String>> VCubeDims = new Dictionary<String, List<String>>();
        Dictionary<String, List<String>> VCubeMeas = new Dictionary<String, List<String>>();
        
        Console.WriteLine("<Schema name=\"" + schemaname + "\">");

        // Database 
        foreach (Database OLAPDatabase in OLAPServer.Databases)
        {

            if (OLAPDB == "" || OLAPDatabase.Name.ToString() == OLAPDB)
            {
                // Cube 
                foreach (Cube OLAPCubex in OLAPDatabase.Cubes)
                {
                    if (OLAPCube == "" || OLAPCubex.Name == OLAPCube)
                    {

                        List<String> dnames = new List <String>();
                        List<String> d2 = new List<String>(); // List of dimension names to be shared throughout XML
                        Dictionary<String,String> fks = new Dictionary<String,String>(); // List of FKs for dimensions
                        List<String> ignoredim = new List<String>();

                        // Check for incompatible dimension types
                        foreach (MeasureGroup OLAPMeasureGroup in OLAPCubex.MeasureGroups) 
                        {
                            foreach (MeasureGroupDimension mgdim in OLAPMeasureGroup.Dimensions)
                            {
                                if (boolM2M && mgdim is ManyToManyMeasureGroupDimension) ignoredim.Add(mgdim.CubeDimension.Name);
                                if (mgdim is DataMiningMeasureGroupDimension) ignoredim.Add(mgdim.CubeDimension.Name);
                            }
                        }
                        
                        foreach (CubeDimension OLAPDimension in OLAPCubex.Dimensions)
                        {

                            // Check to see if we should skip this dim
                            if (ignoredim.Contains(OLAPDimension.Name)) continue;
                            
                            // We can't handle composite keys in Mondrian, so I'm just taking the first key column
                            string[] myspl = splitField(OLAPDimension.Attributes[0].Attribute.KeyColumns[0].ToString());                            

                            String dname ="visible=\""+OLAPDimension.Visible+"\" highCardinality=\"false\" name=\"" + OLAPDimension.Name + "\">";
                            dnames.Add(dname);
                            fks.Add(dname,myspl[1]);
                            d2.Add(OLAPDimension.Name);
                            Console.WriteLine("<Dimension " + dname);
    
                            // Turn attributes into default hierarchy
                            Console.WriteLine("<Hierarchy name=\"Default\" visible=\"true\" hasAll=\"true\" allMemberName=\"All\" primaryKey=\""+myspl[1]+"\">");

                            if (includeschema)
                            {
                                Console.WriteLine("<Table name=\"" + myspl[0] + "\" schema=\""+sschema+"\"></Table>");
                            }
                            else
                            {
                                Console.WriteLine("<Table name=\"" + myspl[0] + "\"></Table>");
                            }
                                
                            foreach (CubeAttribute OLAPDimAttribute in OLAPDimension.Attributes)
                            {

                                String mystring =OLAPDimAttribute.Attribute.KeyColumns[0].ToString();                                
                                string [] mysplit=mystring.Split('.');

                                String ss1=convertDataType(OLAPDimAttribute.Attribute.KeyColumns[0].DataType.ToString());
                                String ss2 = convertLevelType(OLAPDimAttribute.Attribute.Type.ToString());
                                string[] splitNameColumn = OLAPDimAttribute.Attribute.NameColumn.ToString().Split('.');                                

                                Console.WriteLine("<Level name=\"" + OLAPDimAttribute.Attribute.Name + "\" visible=\"" + OLAPDimAttribute.AttributeHierarchyVisible +                                    
                                    "\" column=\"" + mysplit[1] + "\" type=\"" + ss1 +
                                    "\" uniqueMembers=\"true\" levelType=\"" + ss2 + "\" hideMemberIf=\"Never\">");
                                Console.WriteLine("</Level>");
                                    
                            }
                            Console.WriteLine("</Hierarchy>");


                            //Dimension Hierarchy 
                            foreach (CubeHierarchy OLAPDimHierarchy in OLAPDimension.Hierarchies)
                            {                                
                                string [] mysplit = splitField(OLAPDimHierarchy.Hierarchy.Levels[0].SourceAttribute.KeyColumns[0].ToString());

                                Console.WriteLine("<Hierarchy name=\""+ OLAPDimHierarchy.Hierarchy.Name  +"\" visible=\""+ OLAPDimHierarchy.Visible+
                                    "\" hasAll=\"true\" allMemberName=\""+OLAPDimHierarchy.Hierarchy.AllMemberName+
                                    "\" primaryKey=\""+mysplit[1]+"\">");

                                if (includeschema)
                                {
                                    Console.WriteLine("<Table name=\"" + mysplit[0] + "\" schema=\""+sschema+"\"></Table>");
                                }
                                else
                                {
                                    Console.WriteLine("<Table name=\"" + mysplit[0] + "\"></Table>");
                                }
                                                                       

                                foreach (Level OLAPDimHierachyLevel in OLAPDimHierarchy.Hierarchy.Levels)
                                {
                                    string[] ksplit = splitField(OLAPDimHierachyLevel.SourceAttribute.KeyColumns[0].ToString());
                                        
                                    String leveltype = null;
                                    leveltype=convertLevelType(OLAPDimHierachyLevel.SourceAttribute.Type.ToString());

                                    string dt1 = convertDataType(OLAPDimHierachyLevel.SourceAttribute.KeyColumns[0].DataType.ToString());
                                            
                                    Console.WriteLine("<Level name=\"" + OLAPDimHierachyLevel.Name + "\" visible=\"" + OLAPDimHierarchy.Visible + "\" column=\""+
                                        ksplit[1]+"\" type=\""+dt1+"\" uniqueMembers=\"true\" levelType=\""+leveltype+"\" hideMemberIf=\"Never\">");
                                    Console.WriteLine("</Level>");
                                }
                                Console.WriteLine("</Hierarchy>");
                            }
                            Console.WriteLine("</Dimension>");
                        } // NEXT OLAPDimension

                            //Measure Group 
                        int i2 = 0;
                            
                        foreach (MeasureGroup OLAPMeasureGroup in OLAPCubex.MeasureGroups)
                        {                               
                            string[] mysplit = splitField(OLAPMeasureGroup.Measures[0].Source.Source.ToString());
                            Cubes.Add(OLAPMeasureGroup.Name); // for VCube
                            Console.WriteLine("<Cube name=\"" + OLAPMeasureGroup.Name + "\" visible=\"true\" cache=\"true\" enabled=\"true\">");

                            if (includeschema)
                            {
                                if (OLAPMeasureGroup.Measures[0].Source.Source.ToString().Contains("_"))
                                {
                                    sschema = ((ColumnBinding)OLAPMeasureGroup.Measures[0].Source.Source).TableID.Split('_')[0];
                                }
                                Console.WriteLine("<Table name=\"" + mysplit[0] + "\" schema=\"" + sschema + "\"></Table>");
                            }
                            else
                            {
                                Console.WriteLine("<Table name=\"" + mysplit[0] + "\"></Table>");
                            }

                            int ii = 0;
                            foreach (String d in dnames) {
                                Console.WriteLine("<DimensionUsage source=\""+d2[ii]+"\" foreignKey=\""+fks[d]+"\" "+d+"</DimensionUsage>");
                                if (!VCubeDims.ContainsKey(OLAPMeasureGroup.Name))
                                {
                                    List<String> t = new List<String> {d2[ii]};
                                    VCubeDims.Add(OLAPMeasureGroup.Name, t);
                                }
                                else
                                {
                                    VCubeDims[OLAPMeasureGroup.Name].Add(d2[ii]);
                                }
                                ii++;
                            }
                            // Measures 
                            int i3 = 0;
                            //cms = null;
                            foreach (Measure OLAPMeasure in OLAPMeasureGroup.Measures)
                            {
                                String mtl = OLAPMeasure.AggregateFunction.ToString();
                                if (mtl != null) mtl = mtl.ToLower();                                

                                // Convert the aggregation types
                                switch (mtl)
                                {
                                    case "distinctcount":
                                        mtl = "distinct count";
                                        break;
                                    case "none":
                                    case "byaccount":
                                        mtl = "sum";
                                        break;
                                    case "averageofchildren":
                                        mtl = "avg";
                                        break;
                                    case "lastnonempty":
                                        mtl = "sum";
                                        break;
                                }
                                    
                                string[] mn = splitField(OLAPMeasure.Source.Source.ToString());

                                bool isl = false;
                                String omeasure = null;

                                // Check to see if there is a measure expression in the measure definition
                                // If so, we can make it a calculate measure
                                try {
                                    omeasure = OLAPMeasure.MeasureExpression;
                                    if (omeasure.Length > 0) isl = true;
                                }
                                catch (Exception e) {}
                                                                
                                if (isl)
                                {
                                    // This measure has a measure expression, so let's add it to the string of calculated measures (cms)
                                    cms += "<CalculatedMember name=\"" + OLAPMeasure.Name + "\" dimension=\"Measures\" visible=\"" + OLAPMeasure.Visible + "\">" + System.Environment.NewLine +
                                    "  <Formula>    <![CDATA[" + omeasure + "]]> </Formula>" + System.Environment.NewLine +
                                    "  <CalculatedMemberProperty name=\"FORMAT_STRING\" value=\"" + OLAPMeasure.FormatString + "\">  </CalculatedMemberProperty>" + System.Environment.NewLine +
                                    "</CalculatedMember>" + System.Environment.NewLine;                                    
                                }
                                else
                                if (mn.Count() > 1)
                                {
                                    // This measure does not have an expression, so just spit out the measure def without making it a calculated measure
                                    Console.WriteLine("<Measure name=\"" + OLAPMeasure.Name + "\" column=\"" + mn[1] + "\" formatString=\"" + OLAPMeasure.FormatString + "\" aggregator=\"" + mtl + "\">");
                                    Console.WriteLine("</Measure>");
                                    if (!VCubeMeas.ContainsKey(OLAPMeasureGroup.Name))
                                    {
                                        List<String> t = new List<String> { OLAPMeasure.Name };
                                        VCubeMeas.Add(OLAPMeasureGroup.Name, t);
                                    }
                                    else
                                    {
                                        VCubeMeas[OLAPMeasureGroup.Name].Add(OLAPMeasure.Name);
                                    }                                    
                                }
                            }

                            i2++;
                            i3++;

                            // Output the Calculated Members for Measures here
                            if (cms != null)
                            {
                                Console.WriteLine(cms);
                            }                                

                            Console.WriteLine("</Cube>");
                        } // MeasureGroups
                    } // Cube
                }
            }
        }


        // Wrap it into a single Virtual Cube
        Console.WriteLine("<VirtualCube enabled=\"true\" name=\"" + schemaname + "\">");

        // Iterate through cubes for the purposes of outputing the virtual cube(s)
        foreach (String cube in Cubes)
        {
            try
            {
                foreach (String vcd in VCubeDims[cube])
                {
                    Console.WriteLine("<VirtualCubeDimension cubeName=\"" + cube + "\" name=\""+vcd+"\"></VirtualCubeDimension>");
                }
            }
            catch (Exception e) { };
        }

        foreach (String cube in Cubes)
        {
            try
            {
                foreach (String vcm in VCubeMeas[cube])
                {
                    Console.WriteLine("<VirtualCubeMeasure cubeName=\"" + cube + "\" name=\"[Measures].[" + vcm + "]\" visible=\"true\"></VirtualCubeMeasure>");
                }
            }
            catch (Exception e) { };
        }

        if (cms != null) Console.WriteLine(cms);
        Console.WriteLine("</VirtualCube>");
        Console.WriteLine("</Schema>");
        if (shallipause) Console.ReadKey();
    }
}
