using System;
using System.Linq;
using System.ComponentModel;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

using SwissAcademic.Citavi;
using SwissAcademic.Citavi.Metadata;
using SwissAcademic.Citavi.Shell;
using SwissAcademic.Collections;
using SwissAcademic.Citavi.Citations;

using System.Net;
using System.Text;
using System.Text.RegularExpressions;
// Implementation of macro editor is preliminary and experimental.
// The Citavi object model is subject to change in future version.
// 将选中的文献信息（作者、年份、DOI、影响因子等）格式化并导出为YAML格式的文本。
public static class CitaviMacro
{
	public static void Main()
	{
		//Get the active project
		Project project = Program.ActiveProjectShell.Project;
		
		//Get the active ("primary") MainForm
		MainForm mainForm = Program.ActiveProjectShell.PrimaryMainForm;
		//reference to active Project
        Project activeProject = Program.ActiveProjectShell.Project;
		List<Reference> references = mainForm.GetSelectedReferences();
		
		
		foreach (Reference reference in references)
		{
			
			
			//	Cite
			string newShortTitle = null;
			newShortTitle = NewShortTitle(reference);
			// DebugMacro.WriteLine(newShortTitle);
			DebugMacro.WriteLine("> Cite:: "+newShortTitle);
			
			// 路径打印：
	        List<Location> refLocation = reference.Locations.ToList();
			string pathString = "";
            foreach (Location location in refLocation)
            {
                if (reference.Locations == null) continue;
				if (string.IsNullOrEmpty(location.Address.ToString())) continue;
				if (location.LocationType != LocationType.ElectronicAddress) continue;
				string filePath = location.Address.ToString();
				//reference.Locations.Remove
				// && filePath.StartsWith(@"\\"))
				if (filePath.EndsWith(".pdf") || filePath.EndsWith(".xlsx") || filePath.EndsWith(".docx") || filePath.EndsWith(".doc") || filePath.EndsWith(".xls"))
				{
					string path = location.Address.Resolve().LocalPath;
					path = path.Replace(" ","%20");
					string thisString = "["+filePath+"]"+"(file:///"+ path +");";
					pathString = pathString+thisString;

				}
            }
			DebugMacro.WriteLine("> File:: "+pathString);
			// Style高亮设置
			string color3 = "rgb(248, 238, 205)";
			string color35 = "rgb(219, 237, 219)";
			string color510 = "rgb(232, 222, 238)";
			string color10 = "rgb(255, 226, 221)";
			// 使用Parse方法进行转换，如果转换失败会抛出异常
			float IF = float.Parse(reference.CustomField2.Replace("IF: ",""));
			string styleString = "";
			if (IF < 3){
				styleString = "> Style:: <span style=\"background-color: "+ color3+ ";\">IF "+ IF+"</span>";
			}else if(IF >=3 && IF <5){
				styleString = "> Style:: <span style=\"background-color: "+ color35+ ";\">IF "+ IF+"</span>";
			}else if(IF >=5 && IF <10){
				styleString = "> Style:: <span style=\"background-color: "+ color510+ ";\">IF "+ IF+"</span>";
			}else if(IF >=10){
				styleString = "> Style:: <span style=\"background-color: "+ color10+ ";\">IF "+ IF+"</span>";
			}
			// 分区
			string partmy = reference.CustomField4;
			if (partmy.Contains("4区")){
				styleString = styleString + " <span style=\"background-color: "+ color3+ ";\">"+ partmy+"</span>";
			}else if(partmy.Contains("3区")){
				styleString = styleString + " <span style=\"background-color: "+ color35+ ";\">"+ partmy+"</span>";
			}else if(partmy.Contains("2区")){
				styleString = styleString + " <span style=\"background-color: "+ color510+ ";\">"+ partmy+"</span>";
			}else if(partmy.Contains("1区")){
				styleString = styleString + " <span style=\"background-color: "+ color10+ ";\">"+ partmy+"</span>";
			}
			DebugMacro.WriteLine(styleString);
			
			// 作者
			string result= "";
            List<Person> refPersons = reference.Authors.ToList();
            List<string> nameString5 = new List<string>(); // 创建一个空的 List<string>
            foreach (Person person in refPersons)
            {
                if (reference.Authors == null) continue;
                nameString5.Add(person.FullName);
            }
            // 将 List 转换为数组后连接起来
			string[] strings = nameString5.ToArray();
			result = string.Join("; ", strings);
			DebugMacro.WriteLine("> Author:: "+result.Replace(",",""));
			
			// 其他
			DebugMacro.WriteLine("> Year:: "+reference.Year);
			DebugMacro.WriteLine("> DOI:: "+reference.Doi);
			DebugMacro.WriteLine("> IF:: =="+reference.CustomField2.Replace("IF: ","")+"==");
			DebugMacro.WriteLine("> 中科院分区升级版:: =="+reference.CustomField4 +"==");
			DebugMacro.WriteLine("> 5年影响因子:: "+reference.CustomField5.Replace("5年影响因子: ",""));
			DebugMacro.WriteLine("> JCR分区:: "+reference.CustomField3);
			DebugMacro.WriteLine("> Journal:: "+reference.Periodical.ToString());
			DebugMacro.WriteLine("> journalAbbreviation:: "+reference.Periodical.UserAbbreviation1.ToString());
			DebugMacro.WriteLine("> ShortNote:: "+reference.CustomField1);
			
			DebugMacro.WriteLine("> titleTranslation:: "+reference.TranslatedTitle);
			DebugMacro.WriteLine("> abstractTranslation:: "+reference.TableOfContents.Text);
			

		}
	
	}
	
	// 此函数是用来打印相应cite引用的，需要设置35行的styleName
	public static string NewShortTitle(Reference reference)
	{
		try
		{
			if (reference == null) return null;

			string newShortTitle = string.Empty;

			Project project = Program.ActiveProjectShell.Project;

			string styleName = "CitaviDefaultCitationStyle.ccs"; // 在这里修改你的参考文献样式
			
			// 这是用户style
			//string folder = project.Addresses.GetFolderPath(CitaviFolder.CustomCitationStyles).ToString(); 
			//DebugMacro.WriteLine(project.Addresses.GetFolderPath(CitaviFolder.CustomCitationStyles).ToString());
			//string fullPath = folder + @"\" + styleName;
			
			// 这是citavi系统style文件夹
			string folder = project.Addresses.GetFolderPath(CitaviFolder.CitationStyles).ToString(); 
			string fullPath = folder + @"\" + styleName;
			
			if (!System.IO.File.Exists(fullPath))
			{
				MessageBox.Show("没有Style");
				return null;
			}
			Uri uri = new Uri(fullPath);
			CitationStyle citationStyle = CitationStyle.Load(uri.AbsoluteUri); // 根据路径加载style

			List<Reference> references = new List<Reference>();
			references.Add(reference);
			CitationManager citationManager = new CitationManager(Program.ActiveProjectShell.Project.Engine, citationStyle, references);
			if (citationManager == null) return null;
			BibliographyCitation bibliographyCitation = citationManager.BibliographyCitations.FirstOrDefault();
			if (bibliographyCitation == null) return null;
			List<ITextUnit> textUnits = bibliographyCitation.GetTextUnits();
			if (textUnits == null) return null;

			var output = new TextUnitCollection();

			foreach (ITextUnit textUnit in textUnits)
			{
				output.Add(textUnit);
			}

			newShortTitle = output.ToString();
			return newShortTitle;
		}
		catch { return null; }
	}

}