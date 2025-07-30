using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Reflection;
using System.IO;
using Microsoft.Office.Core;
using Excel = Microsoft.Office.Interop.Excel;
using Extensibility;
using System.Drawing;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using System.Collections.Generic;

namespace AIHelper
{
    [ComVisible(true)]
    [Guid("A2F47820-C9D3-4C8F-B3D5-78A982F89E31")]
    [ProgId("AIHelper.Connect")]
    [ClassInterface(ClassInterfaceType.None)]
    public class Connect : IDTExtensibility2
    {
        // Excel 应用程序实例
        private Excel.Application excelApp;
        // 加载项实例
        private object addInInstance;
        // 任务面板
        private object taskPane;
        // 聊天窗体
        private ChatPanel chatPanel;
        // 命令栏菜单
        private CommandBar aiMenu;
        
        // 日志文件路径
        private static readonly string LogFilePath = Path.Combine(
            Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), 
            "AIHelper.log");
        
        // 配置信息
        private string apiKey = string.Empty;
        private string apiEndpoint = "https://api.openai.com/v1/chat/completions";
        private string apiModel = "gpt-3.5-turbo";
        
        #region 构造函数和日志方法

        public Connect()
        {
            WriteLog("Connect 构造函数被调用");
            LoadSettings();
        }
        
        private void WriteLog(string message)
        {
            try
            {
                string timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff");
                File.AppendAllText(LogFilePath, timestamp + ": " + message + "\r\n");
            }
            catch (Exception ex)
            {
                // 如果写入日志失败，尝试创建日志目录
                try
                {
                    string logDir = Path.GetDirectoryName(LogFilePath);
                    if (!Directory.Exists(logDir))
                    {
                        Directory.CreateDirectory(logDir);
                        File.AppendAllText(LogFilePath, DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff") + 
                            ": 日志目录已创建，初始日志信息: " + message + "\r\n");
                    }
                }
                catch
                {
                    // 忽略所有错误，无法写入日志
                }
            }
        }
        
        private void LoadSettings()
        {
            try
            {
                string settingsPath = Path.Combine(
                    Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location),
                    "AIHelper.config");
                
                if (File.Exists(settingsPath))
                {
                    string[] lines = File.ReadAllLines(settingsPath);
                    foreach (string line in lines)
                    {
                        if (line.StartsWith("ApiKey="))
                        {
                            apiKey = line.Substring("ApiKey=".Length);
                        }
                        else if (line.StartsWith("ApiEndpoint="))
                        {
                            apiEndpoint = line.Substring("ApiEndpoint=".Length);
                        }
                        else if (line.StartsWith("ApiModel="))
                        {
                            apiModel = line.Substring("ApiModel=".Length);
                        }
                    }
                }
                WriteLog("配置加载完成");
            }
            catch (Exception ex)
            {
                WriteLog("加载配置错误: " + ex.Message);
            }
        }
        
        private void SaveSettings()
        {
            try
            {
                string settingsPath = Path.Combine(
                    Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location),
                    "AIHelper.config");
                
                using (StreamWriter writer = new StreamWriter(settingsPath))
                {
                    writer.WriteLine("ApiKey=" + apiKey);
                    writer.WriteLine("ApiEndpoint=" + apiEndpoint);
                    writer.WriteLine("ApiModel=" + apiModel);
                }
                
                WriteLog("保存配置成功: " + settingsPath);
            }
            catch (Exception ex)
            {
                WriteLog("保存配置错误: " + ex.Message);
            }
        }

        #endregion

        #region IDTExtensibility2 实现

        public void OnConnection(object Application, Extensibility.ext_ConnectMode ConnectMode, object AddInInst, ref Array custom)
        {
            try
            {
                // 设置静态实例引用
                instance = this;
                
                WriteLog("OnConnection 方法被调用");
                
                excelApp = Application as Microsoft.Office.Interop.Excel.Application;
                addInInstance = AddInInst;
                
                if (excelApp != null)
                {
                    string version = excelApp.Version;
                    WriteLog("成功连接到 Excel 应用程序, 版本: " + version);
                    
                    // 加载设置
                    LoadSettings();
                    
                    // 创建菜单
                    CreateMenu();
                }
                else
                {
                    WriteLog("无法获取Excel应用程序对象");
                }
            }
            catch (Exception ex)
            {
                WriteLog("连接失败: " + ex.Message);
            }
        }

        public void OnDisconnection(Extensibility.ext_DisconnectMode RemoveMode, ref Array custom)
        {
            WriteLog("OnDisconnection 方法被调用，断开模式: " + RemoveMode);
            
            try
            {
                // 清理菜单
                if (aiMenu != null)
                {
                    aiMenu.Delete();
                    aiMenu = null;
                }
                
                // 清理任务面板
                if (taskPane != null && excelApp != null)
                {
                    // 尝试关闭任务面板
                    try
                    {
                        object customTaskPanes = excelApp.GetType().InvokeMember(
                            "CustomTaskPanes", 
                            BindingFlags.GetProperty, 
                            null, 
                            excelApp, 
                            null);
                        
                        if (customTaskPanes != null)
                        {
                            customTaskPanes.GetType().InvokeMember(
                                "Delete",
                                BindingFlags.InvokeMethod,
                                null,
                                customTaskPanes,
                                new object[] { taskPane });
                        }
                    }
                    catch (Exception ex)
                    {
                        WriteLog("清理任务面板失败: " + ex.Message);
                    }
                    
                    taskPane = null;
                }
                
                if (chatPanel != null)
                {
                    if (!chatPanel.IsDisposed)
                    {
                        chatPanel.Dispose();
                    }
                    chatPanel = null;
                }
                
                // 清理资源
                excelApp = null;
                addInInstance = null;
                
                WriteLog("资源已清理");
            }
            catch (Exception ex)
            {
                WriteLog("OnDisconnection 发生错误: " + ex.Message);
            }
        }

        public void OnAddInsUpdate(ref Array custom)
        {
            WriteLog("OnAddInsUpdate 方法被调用");
        }

        public void OnStartupComplete(ref Array custom)
        {
            WriteLog("OnStartupComplete 方法被调用");
        }

        public void OnBeginShutdown(ref Array custom)
        {
            WriteLog("OnBeginShutdown 方法被调用");
        }

        #endregion
        
        #region 菜单和功能实现
        
        private void CreateMenu()
        {
            try
            {
                WriteLog("Creating menu");
                
                // 删除现有菜单（如果存在）
                CommandBars commandBars = (CommandBars)excelApp.CommandBars;
                try
                {
                    CommandBar existingMenu = commandBars["AI Helper"];
                    if (existingMenu != null)
                    {
                        existingMenu.Delete();
                    }
                }
                catch (Exception menuEx)
                {
                    WriteLog("Error deleting existing menu (can be ignored): " + menuEx.Message);
                    // 忽略错误，表示菜单不存在
                }
                
                // 创建新菜单
                aiMenu = commandBars.Add("AI Helper", 1, missing, true);
                
                // 添加菜单项 - 使用OnAction属性而非Click事件
                CommandBarButton showAiButton = (CommandBarButton)aiMenu.Controls.Add(
                    1, missing, missing, missing, true);
                showAiButton.Caption = "Show AI Assistant";
                showAiButton.Style = MsoButtonStyle.msoButtonIconAndCaption;
                showAiButton.FaceId = 59; // Excel内置图标
                showAiButton.OnAction = "!<AIHelper.Connect.OnShowAI>";
                
                CommandBarButton settingsButton = (CommandBarButton)aiMenu.Controls.Add(
                    1, missing, missing, missing, true);
                settingsButton.Caption = "API Settings";
                settingsButton.Style = MsoButtonStyle.msoButtonIconAndCaption;
                settingsButton.FaceId = 23; // 设置图标
                settingsButton.OnAction = "!<AIHelper.Connect.OnSettings>";
                
                CommandBarButton aboutButton = (CommandBarButton)aiMenu.Controls.Add(
                    1, missing, missing, missing, true);
                aboutButton.Caption = "About";
                aboutButton.Style = MsoButtonStyle.msoButtonIconAndCaption;
                aboutButton.FaceId = 487; // 信息图标
                aboutButton.OnAction = "!<AIHelper.Connect.OnAbout>";
                
                // 显示菜单
                aiMenu.Visible = true;
                
                WriteLog("Menu creation completed, OnAction configured");
                // 不显示消息框，避免启动时干扰
            }
            catch (Exception ex)
            {
                WriteLog("Create menu failed: " + ex.Message + "\r\n" + ex.StackTrace);
                // 不显示消息框，避免启动时干扰
            }
        }
        
        // 使用公共静态方法处理菜单点击
        public static void OnShowAI()
        {
            try
            {
                WriteToLog("OnShowAI method called");
                if (instance != null)
                {
                    instance.ShowAiPanel();
                }
                else
                {
                    WriteToLog("Error: instance is null, cannot call ShowAiPanel");
                }
            }
            catch (Exception ex)
            {
                WriteToLog("OnShowAI error: " + ex.Message + "\r\n" + ex.StackTrace);
                MessageBox.Show("Failed to show AI Assistant panel: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        
        public static void OnSettings()
        {
            try
            {
                WriteToLog("OnSettings 方法被调用");
                if (instance != null)
                {
                    instance.ShowSettings();
                }
                else
                {
                    WriteToLog("错误: instance为空，无法调用ShowSettings");
                }
            }
            catch (Exception ex)
            {
                WriteToLog("OnSettings 错误: " + ex.Message + "\r\n" + ex.StackTrace);
                MessageBox.Show("显示设置对话框失败: " + ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        
        public static void OnAbout()
        {
            try
            {
                WriteToLog("OnAbout 方法被调用");
                if (instance != null)
                {
                    instance.ShowAbout();
                }
                else
                {
                    WriteToLog("错误: instance为空，无法调用ShowAbout");
                }
            }
            catch (Exception ex)
            {
                WriteToLog("OnAbout 错误: " + ex.Message + "\r\n" + ex.StackTrace);
                MessageBox.Show("显示关于对话框失败: " + ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        
        // 显示AI助手面板
        private void ShowAiPanel()
        {
            try
            {
                WriteLog("ShowAiPanel called");
                MessageBox.Show("AI Assistant panel function triggered", "AI Helper", MessageBoxButtons.OK, MessageBoxIcon.Information);
                
                // 如果已有面板，则切换显示/隐藏
                if (taskPane != null)
                {
                    bool isVisible = (bool)taskPane.GetType().InvokeMember(
                        "Visible", 
                        BindingFlags.GetProperty, 
                        null, 
                        taskPane, 
                        null);
                    
                    taskPane.GetType().InvokeMember(
                        "Visible", 
                        BindingFlags.SetProperty, 
                        null, 
                        taskPane, 
                        new object[] { !isVisible });
                    
                    WriteLog("Task pane visibility changed to: " + (!isVisible));
                    return;
                }
                
                // TO-DO: 实现实际的任务面板
                WriteLog("Task pane functionality not fully implemented yet");
                MessageBox.Show("AI Assistant panel feature is under development, coming soon!", "AI Helper", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                WriteLog("Show AI Assistant panel error: " + ex.Message + "\r\n" + ex.StackTrace);
                MessageBox.Show("Failed to show AI Assistant panel: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        
        // 静态实例引用和日志方法
        private static Connect instance;
        
        private static void WriteToLog(string message)
        {
            try
            {
                string logPath = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "callback_log.txt");
                using (StreamWriter writer = new StreamWriter(logPath, true))
                {
                    writer.WriteLine(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff") + ": " + message);
                }
            }
            catch
            {
                // 忽略日志写入错误
            }
        }
        
        // 显示设置对话框
        private void ShowSettings()
        {
            try
            {
                WriteLog("ShowSettings called");
                MessageBox.Show("API Settings dialog function triggered", "AI Helper", MessageBoxButtons.OK, MessageBoxIcon.Information);
                
                // 创建设置对话框
                Form settingsForm = new Form
                {
                    Text = "API Settings",
                    Size = new Size(400, 200),
                    StartPosition = FormStartPosition.CenterScreen,
                    FormBorderStyle = FormBorderStyle.FixedDialog,
                    MaximizeBox = false,
                    MinimizeBox = false
                };
                
                // 创建控件
                Label apiKeyLabel = new Label
                {
                    Text = "API Key:",
                    Location = new Point(20, 20),
                    Size = new Size(80, 20)
                };
                
                TextBox apiKeyTextBox = new TextBox
                {
                    Text = apiKey,
                    Location = new Point(110, 20),
                    Size = new Size(250, 20),
                    PasswordChar = '*'
                };
                
                CheckBox showApiKeyCheckBox = new CheckBox
                {
                    Text = "Show API Key",
                    Location = new Point(110, 50),
                    AutoSize = true
                };
                
                Label apiEndpointLabel = new Label
                {
                    Text = "API Endpoint:",
                    Location = new Point(20, 80),
                    Size = new Size(80, 20)
                };
                
                TextBox apiEndpointTextBox = new TextBox
                {
                    Text = apiEndpoint,
                    Location = new Point(110, 80),
                    Size = new Size(250, 20)
                };
                
                Button saveButton = new Button
                {
                    Text = "Save",
                    Location = new Point(110, 120),
                    Size = new Size(80, 30),
                    DialogResult = DialogResult.OK
                };
                
                Button cancelButton = new Button
                {
                    Text = "Cancel",
                    Location = new Point(200, 120),
                    Size = new Size(80, 30),
                    DialogResult = DialogResult.Cancel
                };
                
                // 添加事件处理
                showApiKeyCheckBox.CheckedChanged += (sender, e) =>
                {
                    apiKeyTextBox.PasswordChar = showApiKeyCheckBox.Checked ? '\0' : '*';
                };
                
                // 添加控件到表单
                settingsForm.Controls.Add(apiKeyLabel);
                settingsForm.Controls.Add(apiKeyTextBox);
                settingsForm.Controls.Add(showApiKeyCheckBox);
                settingsForm.Controls.Add(apiEndpointLabel);
                settingsForm.Controls.Add(apiEndpointTextBox);
                settingsForm.Controls.Add(saveButton);
                settingsForm.Controls.Add(cancelButton);
                
                // 显示对话框
                if (settingsForm.ShowDialog() == DialogResult.OK)
                {
                    // 保存设置
                    apiKey = apiKeyTextBox.Text;
                    apiEndpoint = apiEndpointTextBox.Text;
                    SaveSettings();
                    MessageBox.Show("Settings saved", "AI Helper", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                
                WriteLog("Settings dialog closed");
            }
            catch (Exception ex)
            {
                WriteLog("Show settings dialog error: " + ex.Message);
                MessageBox.Show("Failed to show settings dialog: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        
        // 显示关于对话框
        private void ShowAbout()
        {
            try
            {
                WriteLog("ShowAbout called");
                
                MessageBox.Show(
                    "AI Helper Excel Add-in\n" +
                    "Version: 1.0.0\n" +
                    "Build Date: " + Assembly.GetExecutingAssembly().GetName().Version + "\n\n" +
                    "This add-in provides AI assistance for Microsoft Excel.\n" +
                    "You can interact with AI through the AI Assistant panel to perform various spreadsheet operations.",
                    "About AI Helper",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
                
                WriteLog("About dialog displayed");
            }
            catch (Exception ex)
            {
                WriteLog("Show about dialog error: " + ex.Message);
                MessageBox.Show("Failed to show about dialog: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        
        #endregion
        
        #region AI功能实现
        
        internal async Task<string> GetAIResponseAsync(string userMessage)
        {
            if (string.IsNullOrEmpty(apiKey))
            {
                return "请先在设置中配置API密钥。";
            }
            
            using (HttpClient client = new HttpClient())
            {
                client.DefaultRequestHeaders.Add("Authorization", "Bearer " + apiKey);
                
                // 构造请求
                string jsonRequest = "{\"model\":\"" + apiModel + 
                    "\",\"messages\":[{\"role\":\"system\",\"content\":\"你是Excel助手，可以帮助用户完成Excel相关任务。请简洁回答，并提供具体的Excel操作指导。\"},{\"role\":\"user\",\"content\":\"" + 
                    userMessage.Replace("\"", "\\\"") + "\"}],\"temperature\":0.7}";
                
                var content = new StringContent(jsonRequest, Encoding.UTF8, "application/json");
                
                try
                {
                    WriteLog("发送API请求: " + apiEndpoint);
                    var response = await client.PostAsync(apiEndpoint, content);
                    
                    if (response.IsSuccessStatusCode)
                    {
                        var jsonResponse = await response.Content.ReadAsStringAsync();
                        WriteLog("API请求成功");
                        
                        // 简单解析JSON响应
                        int contentStart = jsonResponse.IndexOf("\"content\":\"") + "\"content\":\"".Length;
                        if (contentStart > 0)
                        {
                            int contentEnd = jsonResponse.IndexOf("\"", contentStart);
                            if (contentEnd > contentStart)
                            {
                                string aiMessage = jsonResponse.Substring(contentStart, contentEnd - contentStart);
                                // 处理转义字符
                                aiMessage = aiMessage.Replace("\\n", "\n").Replace("\\\"", "\"").Replace("\\\\", "\\");
                                return aiMessage;
                            }
                        }
                        
                        WriteLog("解析响应失败: " + jsonResponse);
                        return "无法解析AI响应。";
                    }
                    else
                    {
                        var errorContent = await response.Content.ReadAsStringAsync();
                        WriteLog("API请求失败: " + response.StatusCode + " - " + errorContent);
                        return "AI服务返回错误，状态码: " + response.StatusCode;
                    }
                }
                catch (Exception ex)
                {
                    WriteLog("API请求异常: " + ex.Message);
                    return "请求AI服务时出错: " + ex.Message;
                }
            }
        }
        
        internal void ExecuteExcelOperation(string operation)
        {
            try
            {
                WriteLog("执行Excel操作: " + operation);
                
                if (operation.Contains("填充") || operation.Contains("添加数据"))
                {
                    FillDataToSelection(operation);
                }
                else if (operation.Contains("公式") || operation.Contains("计算"))
                {
                    ApplyFormula(operation);
                }
                else if (operation.Contains("格式") || operation.Contains("样式"))
                {
                    ApplyFormatting(operation);
                }
                else
                {
                    WriteLog("无法识别的操作: " + operation);
                    if (chatPanel != null)
                    {
                        chatPanel.AddAIMessage("我无法识别这个操作。请尝试'填充数据'、'添加公式'或'设置格式'等命令。");
                    }
                }
            }
            catch (Exception ex)
            {
                WriteLog("执行Excel操作错误: " + ex.Message);
                if (chatPanel != null)
                {
                    chatPanel.AddAIMessage("执行Excel操作时出错: " + ex.Message);
                }
            }
        }
        
        private void FillDataToSelection(string command)
        {
            if (excelApp == null) 
            {
                WriteLog("Excel应用程序对象为空");
                return;
            }
            
            try
            {
                Excel.Range selection = excelApp.Selection as Excel.Range;
                if (selection == null) 
                {
                    WriteLog("没有选择单元格区域");
                    return;
                }
                
                WriteLog("填充数据到选定区域: " + selection.Address);
                
                // 简单实现：根据命令生成一些示例数据
                if (command.Contains("随机数"))
                {
                    Random random = new Random();
                    for (int row = 1; row <= selection.Rows.Count; row++)
                    {
                        for (int col = 1; col <= selection.Columns.Count; col++)
                        {
                            selection.Cells[row, col].Value = random.Next(1, 100);
                        }
                    }
                    
                    if (chatPanel != null)
                    {
                        chatPanel.AddAIMessage("已在选定区域填充随机数。");
                    }
                }
                else if (command.Contains("日期"))
                {
                    DateTime startDate = DateTime.Today;
                    for (int row = 1; row <= selection.Rows.Count; row++)
                    {
                        for (int col = 1; col <= selection.Columns.Count; col++)
                        {
                            selection.Cells[row, col].Value = startDate.AddDays((row - 1) * selection.Columns.Count + (col - 1));
                            selection.Cells[row, col].NumberFormat = "yyyy-mm-dd";
                        }
                    }
                    
                    if (chatPanel != null)
                    {
                        chatPanel.AddAIMessage("已在选定区域填充日期序列。");
                    }
                }
                else
                {
                    // 默认填充一些文本
                    for (int row = 1; row <= selection.Rows.Count; row++)
                    {
                        for (int col = 1; col <= selection.Columns.Count; col++)
                        {
                            selection.Cells[row, col].Value = "数据 " + row + "-" + col;
                        }
                    }
                    
                    if (chatPanel != null)
                    {
                        chatPanel.AddAIMessage("已在选定区域填充文本数据。");
                    }
                }
            }
            catch (Exception ex)
            {
                WriteLog("填充数据错误: " + ex.Message);
                throw;
            }
        }
        
        private void ApplyFormula(string command)
        {
            if (excelApp == null) 
            {
                WriteLog("Excel应用程序对象为空");
                return;
            }
            
            try
            {
                Excel.Range selection = excelApp.Selection as Excel.Range;
                if (selection == null) 
                {
                    WriteLog("没有选择单元格区域");
                    return;
                }
                
                WriteLog("应用公式到选定区域: " + selection.Address);
                
                string formula = "=SUM(A1:B10)"; // 默认公式
                
                // 解析命令中的公式
                if (command.Contains("求和") || command.Contains("sum"))
                {
                    formula = "=SUM(" + selection.Address + ")";
                }
                else if (command.Contains("平均") || command.Contains("average"))
                {
                    formula = "=AVERAGE(" + selection.Address + ")";
                }
                else if (command.Contains("计数") || command.Contains("count"))
                {
                    formula = "=COUNT(" + selection.Address + ")";
                }
                
                // 应用公式到选定区域旁边的单元格
                Excel.Range targetCell = selection.Offset[0, selection.Columns.Count];
                targetCell.Formula = formula;
                
                if (chatPanel != null)
                {
                    chatPanel.AddAIMessage("已应用公式 '" + formula + "' 到选定区域旁边的单元格。");
                }
            }
            catch (Exception ex)
            {
                WriteLog("应用公式错误: " + ex.Message);
                throw;
            }
        }
        
        private void ApplyFormatting(string command)
        {
            if (excelApp == null) 
            {
                WriteLog("Excel应用程序对象为空");
                return;
            }
            
            try
            {
                Excel.Range selection = excelApp.Selection as Excel.Range;
                if (selection == null) 
                {
                    WriteLog("没有选择单元格区域");
                    return;
                }
                
                WriteLog("应用格式到选定区域: " + selection.Address);
                
                if (command.Contains("粗体") || command.Contains("bold"))
                {
                    selection.Font.Bold = true;
                }
                if (command.Contains("斜体") || command.Contains("italic"))
                {
                    selection.Font.Italic = true;
                }
                if (command.Contains("下划线") || command.Contains("underline"))
                {
                    selection.Font.Underline = true;
                }
                
                // 设置背景色
                if (command.Contains("黄色背景") || command.Contains("yellow"))
                {
                    selection.Interior.Color = ColorTranslator.ToOle(Color.LightYellow);
                }
                else if (command.Contains("蓝色背景") || command.Contains("blue"))
                {
                    selection.Interior.Color = ColorTranslator.ToOle(Color.LightBlue);
                }
                else if (command.Contains("绿色背景") || command.Contains("green"))
                {
                    selection.Interior.Color = ColorTranslator.ToOle(Color.LightGreen);
                }
                
                if (chatPanel != null)
                {
                    chatPanel.AddAIMessage("已应用格式到选定区域。");
                }
            }
            catch (Exception ex)
            {
                WriteLog("应用格式错误: " + ex.Message);
                throw;
            }
        }
        
        #endregion
        
        #region 辅助变量和方法
        
        // missing参数的缓存
        private static object missing = Type.Missing;
        
        #endregion
    }
    
    #region 聊天面板
    
    public class ChatPanel : UserControl
    {
        private Connect connectInstance;
        private TextBox chatHistory;
        private TextBox userInput;
        private Button sendButton;
        private List<string> chatMessages = new List<string>();
        
        public ChatPanel(Connect connect)
        {
            this.connectInstance = connect;
            InitializeComponents();
        }
        
        private void InitializeComponents()
        {
            // 设置控件样式
            this.Size = new Size(300, 600);
            this.Dock = DockStyle.Fill;
            this.BackColor = Color.White;
            
            // 顶部标题
            Panel titlePanel = new Panel
            {
                Height = 40,
                Dock = DockStyle.Top,
                BackColor = Color.SteelBlue
            };
            
            Label titleLabel = new Label
            {
                Text = "AIHelper 助手",
                Font = new Font("微软雅黑", 12, FontStyle.Bold),
                ForeColor = Color.White,
                AutoSize = false,
                TextAlign = ContentAlignment.MiddleCenter,
                Dock = DockStyle.Fill
            };
            
            titlePanel.Controls.Add(titleLabel);
            
            // 聊天历史
            chatHistory = new TextBox
            {
                Multiline = true,
                ReadOnly = true,
                BackColor = Color.White,
                Font = new Font("微软雅黑", 9),
                ScrollBars = ScrollBars.Vertical,
                Dock = DockStyle.Fill
            };
            
            // 底部输入区域
            Panel inputPanel = new Panel
            {
                Height = 100,
                Dock = DockStyle.Bottom
            };
            
            userInput = new TextBox
            {
                Multiline = true,
                Height = 70,
                Dock = DockStyle.Top,
                Font = new Font("微软雅黑", 9)
            };
            
            sendButton = new Button
            {
                Text = "发送",
                Height = 30,
                Dock = DockStyle.Bottom,
                BackColor = Color.RoyalBlue,
                ForeColor = Color.White,
                Font = new Font("微软雅黑", 9, FontStyle.Bold)
            };
            
            // 示例按钮区域
            FlowLayoutPanel examplePanel = new FlowLayoutPanel
            {
                Height = 70,
                Dock = DockStyle.Bottom,
                BackColor = Color.AliceBlue,
                FlowDirection = FlowDirection.LeftToRight,
                WrapContents = true,
                Padding = new Padding(5),
                AutoScroll = true
            };
            
            // 添加示例按钮
            string[] examples = new string[] {
                "在选区填充随机数", 
                "计算选区平均值",
                "设置选区为黄色背景",
                "生成销售数据"
            };
            
            foreach (string example in examples)
            {
                Button exampleButton = new Button
                {
                    Text = example,
                    AutoSize = true,
                    FlatStyle = FlatStyle.Flat,
                    BackColor = Color.LightBlue,
                    ForeColor = Color.DarkBlue,
                    Font = new Font("微软雅黑", 8),
                    Margin = new Padding(3),
                    Cursor = Cursors.Hand
                };
                
                exampleButton.Click += (sender, e) => {
                    userInput.Text = example;
                    SendMessage();
                };
                
                examplePanel.Controls.Add(exampleButton);
            }
            
            // 添加到输入面板
            inputPanel.Controls.Add(userInput);
            inputPanel.Controls.Add(sendButton);
            
            // 添加控件到表单
            this.Controls.Add(chatHistory);
            this.Controls.Add(inputPanel);
            this.Controls.Add(examplePanel);
            this.Controls.Add(titlePanel);
            
            // 添加事件处理
            sendButton.Click += (sender, e) => SendMessage();
            userInput.KeyDown += (sender, e) => {
                if (e.KeyCode == Keys.Enter && e.Control)
                {
                    e.SuppressKeyPress = true;
                    SendMessage();
                }
            };
            
            // 初始欢迎消息
            AddAIMessage("您好，我是Excel AI助手。我可以帮助您完成Excel中的各种操作，如填充数据、应用公式、设置格式等。请告诉我您需要什么帮助？");
        }
        
        private void SendMessage()
        {
            if (string.IsNullOrWhiteSpace(userInput.Text))
                return;
            
            string userMessage = userInput.Text.Trim();
            userInput.Text = string.Empty;
            
            // 添加用户消息到聊天历史
            AddUserMessage(userMessage);
            
            // 显示"正在思考"提示
            AddAIMessage("正在思考...");
            
            // 异步处理AI响应
            Task.Run(async () =>
            {
                try
                {
                    // 获取AI响应
                    string aiResponse = await connectInstance.GetAIResponseAsync(userMessage);
                    
                    // 从聊天中移除"正在思考"消息
                    RemoveLastMessage();
                    
                    // 添加AI响应到聊天历史
                    AddAIMessage(aiResponse);
                    
                    // 执行Excel操作
                    connectInstance.ExecuteExcelOperation(userMessage);
                }
                catch (Exception ex)
                {
                    // 从聊天中移除"正在思考"消息
                    RemoveLastMessage();
                    
                    // 显示错误消息
                    AddAIMessage("抱歉，处理您的请求时出错: " + ex.Message);
                }
            });
        }
        
        public void AddUserMessage(string message)
        {
            if (this.InvokeRequired)
            {
                this.BeginInvoke(new Action(() => AddUserMessage(message)));
                return;
            }
            
            chatMessages.Add("用户: " + message);
            UpdateChatHistory();
        }
        
        public void AddAIMessage(string message)
        {
            if (this.InvokeRequired)
            {
                this.BeginInvoke(new Action(() => AddAIMessage(message)));
                return;
            }
            
            chatMessages.Add("AI: " + message);
            UpdateChatHistory();
        }
        
        private void RemoveLastMessage()
        {
            if (this.InvokeRequired)
            {
                this.BeginInvoke(new Action(RemoveLastMessage));
                return;
            }
            
            if (chatMessages.Count > 0)
            {
                chatMessages.RemoveAt(chatMessages.Count - 1);
                UpdateChatHistory();
            }
        }
        
        private void UpdateChatHistory()
        {
            chatHistory.Text = string.Join(Environment.NewLine + Environment.NewLine, chatMessages);
            chatHistory.SelectionStart = chatHistory.Text.Length;
            chatHistory.ScrollToCaret();
        }
    }
    
    #endregion
}
