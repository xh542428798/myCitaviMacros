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

// Implementation of macro editor is preliminary and experimental.
// The Citavi object model is subject to change in future version.

public static class CitaviMacro
{
	// 注意要在133、134行设置Knowlwdge合并的组别!!!!!
	public static void Main()
	{
		//Get the active project
		Project project = Program.ActiveProjectShell.Project;
		
		//Get the active ("primary") MainForm
		MainForm mainForm = Program.ActiveProjectShell.PrimaryMainForm;
		
		//if this macro should ALWAYS affect all titles in active project, choose:
		//ProjectReferenceCollection references = project.References;		

		//if this macro should affect just filtered rows in the active MainForm, choose:
		List<Reference> references = Program.ActiveProjectShell.PrimaryMainForm.GetSelectedReferences();
		Dictionary<string, Category> categoryDictionary = new Dictionary<string, Category>();
		foreach (Reference currentReference in references)
		{
			// 获取作者、时间、Title
			Person author = currentReference.Authors[0];
			string year = currentReference.Year;
			string IF = currentReference.CustomField1;
			string Qpart = currentReference.CustomField2;
			// 获取Title并提取前10个单词
			string originalTitle = currentReference.Title;
			string[] words = originalTitle.Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
			string result;
			if (words.Length >= 10)
			{// 取前10个单词，首字母大写，其余小写
				result = string.Join(" ", words.Take(10).Select(word => word.First().ToString().ToUpper() + word.Substring(1).ToLower())); 
			}
			else
			{// 所有单词，首字母大写，其余小写
				result = string.Join(" ", words.Select(word => word.First().ToString().ToUpper() + word.Substring(1).ToLower())); 
			}
			string citationkey = author.LastName.ToString() +year+"_"+result+"_"+IF+Qpart;
		

			Category category = project.Categories.Add(citationkey);
			currentReference.Categories.Add(category);
		}
		
		
		//****************************************************************************************************************
		// ADD REFERENCE CATEGORIES, KEYWORDS AND GROUPS TO KNOWLEDGE ITEMS AND VICE VERSA
		// 2.0 -- 2017-03-16
		// 
		//
		// EDIT HERE
		//reference to active Project
		Project activeProject = Program.ActiveProjectShell.Project;
		if (activeProject == null) return;
		//List<Reference> references = Program.ActiveProjectShell.PrimaryMainForm.GetSelectedReferences();
		// 
        // choose direction
        int direction = 1; // 1 for reference -> knowledge item, 2 for knowledge item -> reference

        // set one or several of the following to true
		bool setCategories = true;  

		// DO NOT EDIT BELOW THIS LINE
		// ****************************************************************************************************************

		if (Program.ProjectShells.Count == 0) return;		//no project open
		//if (IsBackupAvailable() == false) return;			//user wants to backup his/her project first


        // wrong direction
        if (direction != 1 && direction != 2)
        {
            DebugMacro.WriteLine("Direction not set correctly, please change line 29 in the code!");
            return;
        }


		//counters

		int categoryCounter = 0;
        int keywordCounter = 0;
        int groupCounter = 0;
        int errorCounter = 0;

		foreach (Reference reference in references)
		{
			foreach (KnowledgeItem knowledgeItem in reference.Quotations)
			{
				if (knowledgeItem.Reference == null) continue; // ignore Ideas
				//MessageBox.Show(knowledgeItem.QuotationType.ToString());
				if (knowledgeItem.QuotationType == QuotationType.Highlight) continue; // ignore highlights

				if (setCategories)
				{
					try
                    {
                        List<Category> kiCategories = knowledgeItem.Categories.ToList();
                        List<Category> refCategories = knowledgeItem.Reference.Categories.ToList();
                        
                        List<Category> mergedCategories = kiCategories.Union(refCategories).ToList();
                        mergedCategories.Sort();

                        switch (direction)
                        {
                            case 1:
                                knowledgeItem.Categories.Clear();
                                knowledgeItem.Categories.AddRange(mergedCategories);
                                categoryCounter++;
                                break;

                            case 2:
                                reference.Categories.Clear();
                                reference.Categories.AddRange(mergedCategories);
                                categoryCounter++;
                                break;
                        }
                                
                       
                    }
                    catch (Exception e)
                    {
                        string errorString = String.Format("An error occurred with knowledge item '{0}' in reference {1}:\n  {2}", knowledgeItem.CoreStatement, reference.ShortTitle, e.Message);
                        DebugMacro.WriteLine(errorString);
                        errorCounter++;
                    }
                   
				}
			
			}		
		}		

		// Message upon completion
		string message = String.Empty;
        
        switch (direction)
        {
            case 1:
                message = "On {0} knowledge items categories have been changed,\n on {1} knowledge items keywords have been changed,\n on {2} knowledge items groups have been changed.\n {3} errors occurred.";
		        break;
            case 2:
                message = "On {0} references categories have been changed,\n on {1} references keywords have been changed,\n on {2} references groups have been changed.\n {3} errors occurred.";
                break;
        }
              
        message = string.Format(message, categoryCounter.ToString(), keywordCounter.ToString(), groupCounter.ToString(), errorCounter.ToString());       
         
		MessageBox.Show(message, "Macro", MessageBoxButtons.OK, MessageBoxIcon.Information);
	}
}