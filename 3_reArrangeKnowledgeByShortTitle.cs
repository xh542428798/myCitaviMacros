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
	public static void Main()
	{
		//Get the active project
		Project project = Program.ActiveProjectShell.Project;
		
		//Get the active ("primary") MainForm
		MainForm mainForm = Program.ActiveProjectShell.PrimaryMainForm;
		
		//if this macro should ALWAYS affect all titles in active project, choose:
		//先排序,然后添加subheading	
        var category = mainForm.GetSelectedKnowledgeOrganizerCategory();

        var knowledgeItems = category.KnowledgeItems.ToList();

        var firstKnowledgeItem = knowledgeItems.First();
        for (int i = 1; i < knowledgeItems.Count; i++)
        {
            category.KnowledgeItems.Move(knowledgeItems[i], firstKnowledgeItem);
            firstKnowledgeItem = knowledgeItems[i];
        }
	//if this macro should affect just filtered rows in the active MainForm, choose:
	//List<Reference> references = mainForm.GetFilteredReferences();	
    //var category = mainForm.GetSelectedKnowledgeOrganizerCategory();
		
		CreateSubheadings(knowledgeItems, category, false);
	}

	
    static void CreateSubheadings(List<KnowledgeItem> knowledgeItems, Category category, bool overwriteSubheadings)
    {
        var mainForm = Program.ActiveProjectShell.PrimaryMainForm;
        var projectShell = Program.ActiveProjectShell;
        var project = projectShell.Project;

        var categoryKnowledgeItems = category.KnowledgeItems;
        var subheadings = knowledgeItems.Where(item => item.KnowledgeItemType == KnowledgeItemType.Subheading).ToList();

        Reference currentReference = null;
        Reference previousReference = null;

        int nextInsertionIndex = -1;

        if (subheadings.Any())
        {
            if (!overwriteSubheadings)
            {
                DialogResult result = MessageBox.Show("The filtered list of knowledge items in the category \"" + category.Name + "\" already contains sub-headings.\r\n\r\nIf you continue, these sub-headings will be removed first.\r\n\r\nContinue?", "Citavi", MessageBoxButtons.YesNo);
                if (result == DialogResult.No) return;
            }

            foreach (var subheading in subheadings)
            {
                project.Thoughts.Remove(subheading);
            }

            projectShell.SaveAsync(mainForm);
        }

        foreach (var knowledgeItem in knowledgeItems)
        {
            if (knowledgeItem.KnowledgeItemType == KnowledgeItemType.Subheading)
            {
                knowledgeItem.Categories.Remove(category);
                continue;
            }

            if (knowledgeItem.Reference != null) currentReference = knowledgeItem.Reference;

            string headingText = "No short title available";
            if (currentReference != null)
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
			
                headingText = citationkey; //currentReference.CitationKey;
            }
            else if (knowledgeItem.QuotationType == QuotationType.None)
            {
                headingText = "Thoughts";
            }

            nextInsertionIndex = category.KnowledgeItems.IndexOf(knowledgeItem);
            category.KnowledgeItems.AddNextItemAtIndex = nextInsertionIndex;
            currentReference = knowledgeItem.Reference;

            if (nextInsertionIndex == 0)
            {
                var subheading = new KnowledgeItem(project, KnowledgeItemType.Subheading) { CoreStatement = headingText };
                subheading.Categories.Add(category);

                project.Thoughts.Add(subheading);
                projectShell.SaveAsync(mainForm);
                previousReference = currentReference;
                continue;
            }

            if (nextInsertionIndex > 0 && (currentReference != null && currentReference != previousReference))
            {
                var subheading = new KnowledgeItem(project, KnowledgeItemType.Subheading) { CoreStatement = headingText };
                subheading.Categories.Add(category);

                project.Thoughts.Add(subheading);
                projectShell.SaveAsync(mainForm);
            }

            previousReference = currentReference;
        }
    }
}