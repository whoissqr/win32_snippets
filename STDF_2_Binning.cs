using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Data;
using System.Data.SqlClient;
using System.Configuration;
using System.Diagnostics;

using LinqToStdf;
using LinqToStdf.Records.V4;

namespace STDF_HB_SB
{
    class Program
    {
        public static void writelog(string msg){
            string file = "C:\\temp\\STDF\\t1t\\dump.txt";
            StreamWriter swlog = new StreamWriter(file, true);
            swlog.Write(msg);
            swlog.Close();
        }

        static void Main(string[] args)
        {
           Stopwatch watch = new Stopwatch();
           watch.Start();

           string stdf_file = @"C:\temp\STDF\20130628033632.stdf";
           var stdf = new StdfFile(stdf_file);
           stdf.EnableCaching = false;  //this is faster and requires only 10mb of mem
           
           #region lot level info from MIR
           /*
           Mir mir = stdf.GetMir();

           string testprogramname = mir.AuxiliaryFile;
           string checksum = mir.DateCode;
           string maskset = mir.DesignRevision;
           string family = mir.FamilyId;
           string facility = mir.FacilityId;
           string testerid = mir.NodeName;
           string lotid = mir.LotId;
           string ftc = mir.LotId;
           string testtemp = mir.TestTemperature;
           string speedgrade = mir.FlowId;

           if (mir.LotId.Contains('-'))
           {
               lotid = mir.LotId.Split('-').First();
               ftc = mir.LotId.Split('-').Last();
           }

           string operatorid = mir.OperatorName;
           string packageid = mir.PackageType;
           string deviceid = mir.PartType;
           string lbid = mir.RomCode;
           string handler = mir.SerialNumber;
           DateTime lotstarttime = (DateTime)mir.StartTime;
           string scd = mir.TestCode;
           string testgrade = mir.SublotId;
           string testgroup = mir.SublotId;
           if (mir.SublotId.Contains('_'))
           {
               testgrade = mir.SublotId.Split('_').First();
               testgroup = mir.SublotId.Split('_').Last();
           }
           */
           #endregion 

           var sdr = stdf.GetRecords().OfExactType<Sdr>(); 
           int siteCNT = sdr.First().Sites.Count();

           bool between_pir_prr = false;
           List<Gdr> gdrList = new List<Gdr>();

           //for a quadsite program, the stdf structure is like following:
           //PIR0-PIR1-PIR2-PIR3 GDR0 GDR1 GDR2 GDR3...... PRR0-PRR1-PRR2-PRR3
           //GDR0 contains the wafer,x,y of site 0
           //GDR1 contains the wafer,x,y of site 1
           //............
           //PRR0 contains the site,hb,sb of site 0
           //............
           //note that GDR does not always appear in sequence, so we need to use 'site=%n' to search within the list
           //here, we uses linq for fast search.
           foreach (var rec in stdf.GetRecords())
           {
               if(rec is Pir) {
                    between_pir_prr = true;
                    gdrList.Clear();              
               }
               else if (rec is Gdr)
               {
                   if (!between_pir_prr) continue;
                   Gdr gdr = (Gdr) rec;
                   if (gdr.GenericData[0].ToString() == "Efuse") {
                       gdrList.Add(gdr);
                   }                
               }
               else if (rec is Prr)
               {
                   between_pir_prr = false;
                   Prr prr = (Prr)rec;

                   var gdr = from g in gdrList
                             where g.GenericData[1].ToString().Split('=').Last().Trim() == prr.SiteNumber.ToString()
                             select new
                             {
                                 waferNum = g.GenericData[3].ToString().Split('=').Last().Trim(),
                                 xcor = g.GenericData[4].ToString().Split('=').Last().Trim(),
                                 ycor = g.GenericData[5].ToString().Split('=').Last().Trim()
                             };
                   if (gdr.Count() == 0)
                   {
                       writelog("No S/N found; [S" + prr.SiteNumber + "]: HB " + prr.HardBin.ToString() + " SB " + prr.SoftBin.ToString() + "\n");
                       continue;  //no efuse detected for this dut.
                   }
                   writelog("Wafer " + gdr.First().waferNum + ", " + "(" + gdr.First().xcor + "," + gdr.First().ycor + ")");
                   writelog(" [S" + prr.SiteNumber + "]: HB " + prr.HardBin.ToString() + " SB " + prr.SoftBin.ToString() + "\n");
               }               
               else
               {
                   //we don't care about other nodes.
               }
           }
           watch.Stop();
           Console.WriteLine("time taken by method a is " + watch.ElapsedMilliseconds + "ms.");           
        }
    }
}
