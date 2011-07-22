namespace Excel.NAnt.Tasks
{
    using System;
    using global::NAnt.Core;
    using global::NAnt.Core.Attributes;
    using global::NAnt.Core.Types;
    using XL=Microsoft.Office.Interop.Excel;
    using System.IO;
    using System.Collections.Generic;


    [TaskName("CreateExcelFile")]
    public class CreateExcelFileTask : Task
    {

        protected override void ExecuteTask()
        {
            XL.Application app = new XL.Application();
            try
            {
                string savePath = resolveFilePath(this.OutputFile);
                if (File.Exists(savePath))
                    throw new BuildException(string.Format("The file '{0}' already exists.", savePath));

                XL.Workbook w = app.Workbooks.Add(XL.XlWBATemplate.xlWBATWorksheet);
                
                Console.WriteLine("Setting References ...{0} of them.", references.FileNames.Count);

                foreach (var path in getReferences())
                {
                    Console.WriteLine("Adding reference to {0}.", path);
                    w.VBProject.References.AddFromFile(path);
                }


                Console.WriteLine("Adding modules...{0} of them.",modules.FileNames.Count);

                foreach (var path in modules.FileNames)
                {
                    Console.WriteLine("Importing {0}.", path);
                    w.VBProject.VBComponents.Import(path);
                }

                bool firstWorksheetModified = false;
                foreach (Worksheet sheet in this.Worksheets)
                {
                    Console.WriteLine("Creating Worksheet {0}.", sheet.SheetName);
                    XL.Worksheet newSheet;
                    if (!firstWorksheetModified)
                    {
                        newSheet = w.Worksheets[1] as XL.Worksheet;
                        firstWorksheetModified = true;
                    }
                    else
                        newSheet = w.Worksheets.Add() as XL.Worksheet;

                    newSheet.Name = sheet.SheetName;
                }
                w.SaveAs(savePath);
                Console.WriteLine("File saved to: " + w.FullName);
            }
            finally
            {
                app.Quit();
            }
        }

        private string[] getReferences()
        {
            List<string> output = new List<string>();
            foreach (var path in references.FileNames)
            {
                output.Add(
                    ValidateFilePath(path)
                );
            }
            return output.ToArray();
        }

        private string ValidateFilePath(string s)
        {
            string filePath=resolveFilePath(s);

            if(!File.Exists(filePath))
                throw new BuildException(string.Format("File '{0}' cannot be found.",filePath));
            return filePath;
        }

        private string resolveFilePath(string fileName)
        {
            string output;
            if (Path.IsPathRooted(fileName))
                output = fileName;
            else
                output = Path.Combine(this.Project.BaseDirectory, this.OutputFile);
            return output;
        }

        [TaskAttribute("outputFile", Required = true)]
        [StringValidator(AllowEmpty = false)]
        public string OutputFile { get; set; }

        

        private FileSet modules=new FileSet();

        [BuildElement("modules")]
        [StringValidator(AllowEmpty=true)]
        public FileSet Modules
        {
            get { return modules; }
            set { modules = value; }
        }

        private FileSet references=new FileSet();

        [BuildElement("references")]
        [StringValidator(AllowEmpty = true)]
        public FileSet References
        {
            get { return references; }
            set { references = value; }
        }

        private Worksheets worksheets=new Worksheets();

        [BuildElementCollection("worksheets","worksheet",Required=false)]
        public Worksheets Worksheets
        {
            get { return worksheets; }
            set { worksheets = value; }
        }
        
        
    }
}
