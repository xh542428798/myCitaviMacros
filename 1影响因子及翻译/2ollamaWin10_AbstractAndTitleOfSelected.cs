// autoref "System.Net.Http.dll"

// 请在tools->reference里添加 System.Net.Http;
// win10安装Ollama，启动Ollama程序后，程序自动运行了openai api，就可以像调用openai v1一样调用这个软件了;
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
using System.Net.Http.Headers;
// 请在tools->reference里添加 System.Net.Http;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;


// Implementation of macro editor is preliminary and experimental.
// The Citavi object model is subject to change in future version.

// 请在tools->reference里添加 System.Net.Http;
public static class CitaviMacro
{
	public static void Main()
	{
		//ProjectReferenceCollection references = Program.ActiveProjectShell.Project.References;		
		List<Reference> references = Program.ActiveProjectShell.PrimaryMainForm.GetSelectedReferences();
		
		int count = 0;
		int totalCount = references.Count;
		
		// 非并行
		foreach (Reference reference in references)
		{
			// 摘要翻译
			string q = reference.Abstract.Text;
			// 翻译
			string q_trans = translateglm(q);
			if ( ! (string.IsNullOrEmpty( q_trans ) ) )
			{
			  reference.TableOfContents.Text = q_trans;
			}

			// 标题翻译
			string q_title = reference.Title;
			//if (not (string.IsNullOrEmpty( q ) ) )
			//{
			//	continue;
			//}
			string q_title_trans = translateglm(q_title);
			if (! (string.IsNullOrEmpty( q_title_trans ) ) )
			{
			  //将TranslatedTitle转成CitationKey
			  //reference.TranslatedTitle = q_title_trans;
			  reference.CustomField6 = q_title_trans; 
			}

			//把标题翻译放入CF6
		    count++;
		    DebugMacro.WriteLine(count + " / " + totalCount);
		}

	    MessageBox.Show("已完成选中条目的摘要和标题翻译，保存在TableOfContents、CustomField6中");
	}
	
	static string translateglm(string q)
	{
		string url = "http://127.0.0.1:11434/v1/chat/completions";

        using (HttpClient client = new HttpClient())
        {
            var request = new HttpRequestMessage(HttpMethod.Post, url);
            
            // 设置请求头
            request.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));

            // 设置请求体
		    // 注意这里要修改模型名称gemma3:12b、qwen3:8b
		    string jsonTemplate = "{{\"model\":\"gemma3:12b\",\"messages\":[{{\"role\":\"user\",\"content\":\"无需深入思考或长篇解释，请将以下内容翻译成中文,只返回翻译后的内容：{0}/nothink\"}}]}}";
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
        			//content = Regex.Replace(content, @"\\u003c(think|\/think)\\u003e.*?\\u003c/think\\u003e", "", RegexOptions.IgnoreCase); // 去除 <think> 和 </think> 标签  
					content = Regex.Replace(content,@"\\u003c(think|\/think)\\u003e.*?\\u003c/think\\u003e[\s\S]{4}", "", RegexOptions.IgnoreCase); // 去除 <think> 和 </think> 标签  
					
					//DebugMacro.WriteLine("Extracted Content: " + content);
					content = Regex.Replace(content, @"[\r\n]", "").Replace(@"[\n]", "");
					//DebugMacro.WriteLine(content);  // 去除换行符  // 去除换行符 \n\n
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

    /// <summary>
    ///   替换部分字符串
    /// </summary>
    /// <param name="sPassed">需要替换的字符串</param>
    /// <returns></returns>
    public static string ReplaceString(string JsonString)
    {
        if (JsonString == null) { return JsonString; }
        if (JsonString.Contains("\\"))
        {
            JsonString = JsonString.Replace("\\", "\\\\");
        }
        if (JsonString.Contains("\'"))
        {
            JsonString = JsonString.Replace("\'", "\\\'");
        }
        if (JsonString.Contains("\""))
        {
            JsonString = JsonString.Replace("\"", "\\\\\\\"");
        }
        //去掉字符串的回车换行符
        JsonString = Regex.Replace(JsonString, @"[\n\r]", "");
        JsonString = JsonString.Trim();
        return JsonString;
    }

}