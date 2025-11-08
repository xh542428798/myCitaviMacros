// autoref "D:\Program Files (x86)\Citavi 6\bin\System.Net.Http.dll"

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

using System.Net.Http; 

// 请在tools->reference里添加 System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Net;
using System.Text.RegularExpressions;


// Implementation of macro editor is preliminary and experimental.
// The Citavi object model is subject to change in future version.

public static class CitaviMacro
{
	public static void Main()
	{
		//ProjectReferenceCollection references = Program.ActiveProjectShell.Project.References;		
		List<Reference> references = Program.ActiveProjectShell.PrimaryMainForm.GetSelectedReferences();
		int count = 0;
		int totalCount = references.Count;
		foreach (Reference reference in references)
		{	
			string q = reference.Title;
			//DebugMacro.WriteLine(q);
			if (string.IsNullOrEmpty( q ) )
			{
				MessageBox.Show("错误");
				continue;
			}
			string q_trans = translateglm(q);
			// 将TranslatedTitle转成CitationKey
			// reference.TranslatedTitle = q_trans;
			if (q_trans!="nodata")
			{
				reference.CustomField6 = q_trans; //把标题翻译放入CF6
				// reference.Subtitle = q_trans;
				// MessageBox.Show(q_trans);
			}

			//DebugMacro.WriteLine(q_trans);
			count++;
		    DebugMacro.WriteLine(count + " / " + totalCount);
		}
	    MessageBox.Show("已完成选中条目的标题翻译，保存在CustomField6中");
	}
	
	static string translateglm(string q)
	{
		string url = "http://10.241.0.114:1555/v1/chat/completions";

        using (HttpClient client = new HttpClient())
        {
            var request = new HttpRequestMessage(HttpMethod.Post, url);
            
            // 设置请求头
            request.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));

            // 设置请求体
		    
		    string jsonTemplate = "{{\"model\":\"gpt-3.5-turbo\",\"messages\":[{{\"role\":\"user\",\"content\":\"请将以下内容翻译成中文：{0}\"}}]}}";
		    string escapedQ = EscapeForJson(q);
		    string jsonContent = string.Format(jsonTemplate, escapedQ);
            //var jsonContent = "{\"model\":\"gpt-3.5-turbo\",\"messages\":[{\"role\":\"user\",\"content\":\"请将以下内容翻译成英文：肝内多房囊性占位\"}]}";
            request.Content = new StringContent(jsonContent, Encoding.UTF8, "application/json");

            try
            {
                HttpResponseMessage response = client.SendAsync(request).Result;
                response.EnsureSuccessStatusCode();
                string responseBody = response.Content.ReadAsStringAsync().Result;
				//DebugMacro.WriteLine(responseBody);
                // 解析并提取所需内容
				//string jsonResponse = ExtractValue(responseBody, "\"content\":\"(.*?)\"");
				string pattern = "\"content\":\"((?:\\\\\"|[^\"])*)\"";

		        Match match = Regex.Match(responseBody, pattern);

		        if (match.Success)
		        {
		            string content = match.Groups[1].Value;
		            content = content.Replace("\\\"", "\""); // 处理转义引号
		            //DebugMacro.WriteLine("Extracted Content: " + content);
					return(content);
		        }
		        else
		        {
		            //DebugMacro.WriteLine("No match found.");
					return("");
		        }
            }
            catch (HttpRequestException e)
            {
                DebugMacro.WriteLine("Request error: " + e.Message);
				return("");
            }
        }
	
	}
	
	static string EscapeForJson(string str)
    {
        StringBuilder sb = new StringBuilder();
        foreach (char c in str)
        {
            switch (c)
            {
                case '\"':
                    sb.Append("\\\"");
                    break;
                case '\\':
                    sb.Append("\\\\");
                    break;
                case '\b':
                    sb.Append("\\b");
                    break;
                case '\f':
                    sb.Append("\\f");
                    break;
                case '\n':
                    sb.Append("\\n");
                    break;
                case '\r':
                    sb.Append("\\r");
                    break;
                case '\t':
                    sb.Append("\\t");
                    break;
                default:
                    if (c < 32 || c > 126)
                    {
                        sb.AppendFormat("\\u{0:X4}", (int)c);
                    }
                    else
                    {
                        sb.Append(c);
                    }
                    break;
            }
        }
        return sb.ToString();
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