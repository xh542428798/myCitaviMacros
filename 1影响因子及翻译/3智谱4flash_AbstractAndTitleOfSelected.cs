// autoref "Newtonsoft.Json.dll"

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

using Newtonsoft.Json; //Newtonsoft.Json
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
		String TranText;
		// 非并行
		foreach (Reference reference in references)
		{
			// 标题翻译
			string q_title = reference.Title;
	        TranText =  TranslateAsync(q_title);
			if (! (string.IsNullOrEmpty( TranText ) ) )
			{
			  //将TranslatedTitle转成CitationKey
			  //reference.TranslatedTitle = q_title_trans;
			  reference.CustomField6 = TranText;  //把标题翻译放入CF6
			}
			//DebugMacro.WriteLine(TranText);
		
			// 摘要翻译
			string q_Abstract = reference.Abstract.Text;

	        TranText =  TranslateAsync(q_Abstract);
			if ( ! (string.IsNullOrEmpty( TranText ) ) )
			{
			  reference.TableOfContents.Text = TranText;
			}
			//DebugMacro.WriteLine(TranText);
			
		    count++;
		    DebugMacro.WriteLine(count + " / " + totalCount);
		}


		
	    DebugMacro.WriteLine("已完成选中条目的摘要和标题翻译，保存在TableOfContents、CustomField6中");
	}
	
    /// <summary> 执行翻译任务的核心逻辑 </summary>
    static String TranslateAsync(String selectedText)
    {
		
        string apiKey = ""; // 确保此路径正确！
        //string selectedText = "MedicalNet is released under the MIT License (refer to the LICENSE file for detailso).";
		string myContent = string.Format("无需深入思考或长篇解释，请将以下内容翻译成中文，只返回翻译后的内容：{0}" , selectedText);

		// 3. 构建请求
        string apiUrl = "https://open.bigmodel.cn/api/paas/v4/chat/completions";
		
		// 定义一个静态的HttpClient，以提高性能
		using (HttpClient client = new HttpClient())
        {
			try
	        {
	            // 设置请求头
	            client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", apiKey);
	            
	            // 构建请求体，严格遵循您提供的Python示例结构
	            var requestBody = new
	            {
	                model = "glm-4-flash-250414", // 填写您要调用的模型名称
	                messages = new[]
	                {
						new { role = "system", content = "你是一个医学英语专家，正在进行医学文献翻译"},
	                    new { role = "user", content = myContent}
	                }
	            };

	            string jsonContent = JsonConvert.SerializeObject(requestBody);
	            var content = new StringContent(jsonContent, Encoding.UTF8, "application/json");

	            // 4. 发送POST请求
				HttpResponseMessage response = client.PostAsync(apiUrl, content).Result;
	            
                // 6. 解析响应
                response.EnsureSuccessStatusCode();
                string responseBody = response.Content.ReadAsStringAsync().Result;
				// 'responseBody' 是您从API收到的那个长长的JSON字符串
				// 1. 反序列化为我们定义的 ZhipuApiResponse 类
				ZhipuApiResponse responseObj = JsonConvert.DeserializeObject<ZhipuApiResponse>(responseBody);
				// 按照我们分析的路径 choices[0].message.content 来获取内容
				string translatedText = responseObj.Choices[0].Message.Content;
				return(translatedText);
	        }
	        catch (HttpRequestException httpEx)
	        {
	            // 捕获HTTP请求相关的错误
	            DebugMacro.WriteLine("网络请求失败: {httpEx.Message}\n\n请检查网络连接或API Key是否有效。");
				return("");
	        }
	        catch (Exception ex)
	        {
	            // 捕获其他所有错误
	            DebugMacro.WriteLine("翻译过程中发生错误: {ex.Message}");
				return("");
	        }
		
		}

    }

	// --- ZhipuApiResponse responseObj 类定义 ---
	// 1. 代表整个JSON响应的最外层对象
	public class ZhipuApiResponse
	{
	    // [JsonProperty("...")] 是一个特性，它告诉JSON解析器：
	    // C#里的这个属性名（如 Choices）对应JSON里的哪个键名（如 "choices"）。
	    // 这样即使命名风格不同（如C#用大写开头，JSON用小写），也能正确匹配。
	    [JsonProperty("choices")]
	    public List<Choice> Choices { get; set; }

	    // 我们只关心我们需要的字段，其他字段（如 created, id）可以忽略，解析器会自动跳过它们。
	}

	// 2. 代表 "choices" 数组里的每一个元素
	public class Choice
	{
	    [JsonProperty("message")]
	    public Message Message { get; set; }
	}

	// 3. 代表 "message" 对象
	public class Message
	{
	    [JsonProperty("content")]
	    public string Content { get; set; }

	    // 我们也可以把 "role" 定义上，虽然翻译用不到，但能让结构更完整
	    [JsonProperty("role")]
	    public string Role { get; set; }
	}
	
}