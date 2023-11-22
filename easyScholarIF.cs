using System;
using System.Net;
using Newtonsoft.Json;
using System.Text.RegularExpressions;

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

// Implementation of macro editor is preliminary and experimental.
// The Citavi object model is subject to change in future version.

public static class CitaviMacro
{
	public static void Main()
	{
		//Get the active project
		Project project = Program.ActiveProjectShell.Project;
		
		//Get the active ("primary") MainForm
		MainForm mainForm = Program.ActiveProjectShell.PrimaryMainForm;
		
		List<Reference> references = Program.ActiveProjectShell.PrimaryMainForm.GetSelectedReferences();
		
		// easyscholar配置
		string secretKey = "bd6104d14938446ba7f7767bf719e162";
		string apiUrl = "https://www.easyscholar.cc/open/getPublicationRank";
		
		foreach (Reference reference in references)
		{	
			//MessageBox.Show(reference.ReferenceType.ToString());
			if (reference.ReferenceType.ToString() != "Journal Article")continue;
			//if(string.IsNullOrEmpty(response)){continue;};
			//MessageBox.Show((reference.Periodical == null).ToString()) ;
			if (reference.Periodical == null)continue;
			//MessageBox.Show(reference.Periodical);
			string publicationName = reference.Periodical.Name; //"Journal of Cancer";	
			string url = apiUrl + "?secretKey=" + secretKey + "&publicationName=" + Uri.EscapeDataString(publicationName);
			using (WebClient client = new WebClient())
			{
				client.Encoding = System.Text.Encoding.UTF8;
				string response = client.DownloadString(url);
				if(string.IsNullOrEmpty(response)){continue;};
				// dynamic result = JsonConvert.DeserializeObject(response);
	            // 使用正则表达式提取数据
	            string sciif = ExtractValue(response, "\"sciif\":\"(.*?)\"");
	            string sci = ExtractValue(response, "\"sci\":\"(.*?)\"");
	            string sciUp = ExtractValue(response, "\"sciUp\":\"(.*?)\"");

	            // 打印输出
	            //MessageBox.Show("sciif: " + sciif);
	            //MessageBox.Show("sci: " + sci);
	            //MessageBox.Show("sciUp: " + sciUp);
				if(!string.IsNullOrEmpty(sciif)){
					reference.CustomField2 = "IF: "+ sciif;
				}
				
				if(!string.IsNullOrEmpty(sci)){
					reference.CustomField3 = sci;
				}
				if(!string.IsNullOrEmpty(sciUp)){
					reference.CustomField4 = sciUp;
				}

			}
            // 延迟 300 ms后发送 POST 请求
            Thread.Sleep(300);
		
		}
		MessageBox.Show("期刊信息更新完成!");

	}
	
    static string ExtractValue(string input, string pattern)
    {
        Match match = Regex.Match(input, pattern);
        if (match.Success)
        {
            return match.Groups[1].Value;
        }
        else
        {
            return "";
        }
    }
}