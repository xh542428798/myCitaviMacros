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


using System.Net;
using System.Text;
using System.Text.RegularExpressions;


// Implementation of macro editor is preliminary and experimental.
// The Citavi object model is subject to change in future version.

public static class CitaviMacro
{
	public static void Main()
	{
		//ProjectReferenceCollection references = Program.ActiveProjectShell.Project.References;		
		List<Reference> references = Program.ActiveProjectShell.PrimaryMainForm.GetSelectedReferences();
		
		foreach (Reference reference in references)
		{	
			string q = reference.Title;
			// if (not  (string.IsNullOrEmpty( q ) ) )
			// {
				// MessageBox.Show("错误");
			//	continue;
			//}
			string q_trans = Google(q);
			// 将TranslatedTitle转成CitationKey
			reference.Subtitle = q_trans;
			// MessageBox.Show(q_trans);
			
		}
	    MessageBox.Show("已完成选中条目的标题翻译，保存在Subtitle中");
	}
	
	
	
	
    public static string Google(string q)
    {
        string url = "https://translate.google.com/_/TranslateWebserverUi/data/batchexecute?rpcids=MkEWBc&bl=boq_translate-webserver_20210927.13_p0&soc-app=1&soc-platform=1&soc-device=1&rt=c";
        // string q = "Expression of immune checkpoint regulators, cytotoxic T-lymphocyte antigen-4, and programmed death-ligand 1 in epstein-barr virus-associated nasopharyngeal carcinoma";
        string from = "auto";
        string to = "zh";
        var from_data = "f.req=" + System.Net.WebUtility.UrlEncode( //System.Web.HttpUtility.UrlEncode(
            string.Format("[[[\"MkEWBc\",\"[[\\\"{0}\\\",\\\"{1}\\\",\\\"{2}\\\",true],[null]]\", null, \"generic\"]]]",
            ReplaceString(q), from, to)).Replace("+", "%20");
        byte[] postData = Encoding.UTF8.GetBytes(from_data);
        WebClient client = new WebClient();
        client.Encoding = Encoding.UTF8;
        client.Headers.Add("Content-Type", "application/x-www-form-urlencoded; charset=UTF-8");
        client.Headers.Add("ContentLength", postData.Length.ToString());
        client.Headers.Add("accept", "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9");
        client.Headers.Add("User-Agent", "Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/93.0.4577.82 Safari/537.36");
        byte[] responseData = client.UploadData(url, "POST", postData);
        string content = Encoding.UTF8.GetString(responseData);
		
        //MessageBox.Show(MatchResult(content));
		string q_trans = MatchResult(content);
		//return content;
		return q_trans;
    }

    /// <summary>
    /// 匹配翻译结果
    /// </summary>
    /// <returns></returns>
    public static string MatchResult(string content)
    {
        string result = "";
        string pattern  = @",\[\[\\\""(.*?)\\\"",";
		//string pattern = @"(?<=""zh"",true]],""zh"",true]],\[.*?]]),"; // 这是Gpt写的
        MatchCollection matchcollection = Regex.Matches(content, pattern);
        if (matchcollection.Count > 0)
        {
            List<string> list = new List<string>();
            foreach (Match match in matchcollection.Cast<Match>().Skip(1))
            {
				// MessageBox.Show(match.Groups[1].Value);
                list.Add(match.Groups[1].Value);
            }
            // 或者一行完成转换和声明
            List<string> newList = list.Distinct().ToList();

            //result = string.Join(" ", newList.GetRange(1, newList.Count - 1));
			result = string.Join(" ", newList);
            if (result.LastIndexOf(@"\""]]]],\""") > 0)
            {
                result = result.Substring(0, result.LastIndexOf(@"\""]]]],"));
            }
        }
        return result;
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