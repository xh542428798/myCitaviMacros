using System;
using System.Linq;
using System.ComponentModel;
using System.Collections.Generic;
using System.Windows.Forms;
using SwissAcademic.Citavi;
using SwissAcademic.Citavi.Metadata;
using SwissAcademic.Citavi.Shell;
using SwissAcademic.Collections;
using System.Diagnostics;

public static class CitaviMacro
{
    public static void Main() 
    {
        // 获取当前项目和主窗口的引用
        Project project = null;
        MainForm mainForm = null;
        
        // 健壮性检查：确保项目外壳存在
        if (Program.ActiveProjectShell != null)
        {
            project = Program.ActiveProjectShell.Project;
            mainForm = Program.ActiveProjectShell.PrimaryMainForm;
        }

        // 1. 健壮性检查：确保项目和主窗口存在
        if (project == null || mainForm == null)
        {
            MessageBox.Show("没有打开的Citavi项目。", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            return;
        }

        // 2. 获取所有选中的条目
        List<Reference> selectedReferences = mainForm.GetSelectedReferences();
        List<KnowledgeItem> selectedKnowledgeItems = mainForm.GetSelectedKnowledgeItems();
        List<Location> selectedLocations = mainForm.GetSelectedElectronicLocations();

        // 3. 初始化一个空字符串来存放结果
        string output = "";
        bool foundAnyItem = false;

        // 4. 处理选中的参考文献
        if (selectedReferences != null && selectedReferences.Count > 0)
        {
            foundAnyItem = true;
            output += "--- 选中的参考文献 ---\n";
            foreach (Reference reference in selectedReferences)
            {
                output += string.Format("RefID: {0}\n", reference.Id.ToStringSafe());
            }
            output += "\n"; // 添加一个空行分隔
        }

        // 5. 处理选中的知识条目
        if (selectedKnowledgeItems != null && selectedKnowledgeItems.Count > 0)
        {
            foundAnyItem = true;
            output += "--- 选中的知识条目 ---\n";
            foreach (KnowledgeItem knowledgeItem in selectedKnowledgeItems)
            {
                output += string.Format("KnowID: {0}\n", knowledgeItem.Id.ToStringSafe());
            }
            output += "\n"; // 添加一个空行分隔
        }

        // 6. 处理选中的附件位置
        if (selectedLocations != null && selectedLocations.Count > 0)
        {
            foundAnyItem = true;
            output += "--- 选中的附件位置 ---\n";
            foreach (Location location in selectedLocations)
            {
                output += string.Format("LocaID: {0}\n", location.Id.ToStringSafe());
            }
        }

        // 7. 根据是否找到条目来显示结果
        if (foundAnyItem)
        {
            // 移除末尾可能多余的换行符
            string finalOutput = output.TrimEnd('\n'); 
            DebugMacro.WriteLine("--- 宏执行结果 ---");
            DebugMacro.WriteLine(finalOutput);
            DebugMacro.WriteLine("--------------------");

            Clipboard.SetText(finalOutput);
            MessageBox.Show(string.Format("已找到并复制选中条目的ID！\n\n{0}", finalOutput), "成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        else
        {
            MessageBox.Show("未选中任何参考文献、知识条目或附件位置。\n\n请在Citavi中选中至少一个条目后重试。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }
}