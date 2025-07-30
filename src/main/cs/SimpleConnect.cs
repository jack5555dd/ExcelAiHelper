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
    public class Connect : IDTExtensibility2, IRibbonExtensibility
    {
        private Excel.Application excelApp;
        private object addInInstance;
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
                File.AppendAllText(LogFilePath, DateTime.Now.ToString() + ": " + message + "\r\n");
            }
            catch
            {
                // 忽略日志写入错误
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
            }
            catch (Exception ex)
            {
                WriteLog("加载配置错误: " + ex.Message);
            }
        }

        #endregion

        #region IDTExtensibility2 实现

        public void OnConnection(object Application, Extensibility.ext_ConnectMode ConnectMode, object AddInInst, ref Array custom)
        {
            try
            {
                WriteLog("OnConnection 方法被调用，连接模式: " + ConnectMode);
                excelApp = Application as Excel.Application;
                addInInstance = AddInInst;
                
                if (excelApp != null)
                {
                    WriteLog("成功连接到 Excel 应用程序, 版本: " + excelApp.Version);
                }
                else
                {
                    WriteLog("警告：Excel 应用程序对象为空");
                }
            }
            catch (Exception ex)
            {
                WriteLog("OnConnection 发生错误: " + ex.Message + "\r\n" + ex.StackTrace);
                MessageBox.Show("加载项连接错误：" + ex.Message, "错误", 
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void OnDisconnection(Extensibility.ext_DisconnectMode RemoveMode, ref Array custom)
        {
            WriteLog("OnDisconnection 方法被调用，断开模式: " + RemoveMode);
            
            try
            {
                // 清理资源
                excelApp = null;
                addInInstance = null;
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

        #region IRibbonExtensibility 实现

        public string GetCustomUI(string ribbonID)
        {
            WriteLog("GetCustomUI 方法被调用，ribbonID: " + ribbonID);
            
            // 返回自定义 Ribbon XML
            return @"<customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui'>
                <ribbon>
                    <tabs>
                        <tab id='customTab' label='AIHelper'>
                            <group id='aiChatGroup' label='AI Assistant'>
                                <button id='btnShowAIPanel' 
                                        label='Show AI Assistant' 
                                        size='large'
                                        onAction='OnShowAIPanel'
                                        imageMso='ReviewShowMarkupMenu'/>
                            </group>
                            <group id='settingsGroup' label='Settings'>
                                <button id='btnSettings' 
                                        label='API Settings' 
                                        size='normal'
                                        onAction='OnSettings'
                                        imageMso='ServerProperties'/>
                                <button id='btnAbout' 
                                        label='About' 
                                        size='normal'
                                        onAction='OnAbout'
                                        imageMso='Info'/>
                            </group>
                        </tab>
                    </tabs>
                </ribbon>
            </customUI>";
        }

        #endregion

        #region Ribbon 回调方法

        public void OnShowAIPanel(IRibbonControl control)
        {
            WriteLog("OnShowAIPanel 方法被调用，控件ID: " + control.Id);
            MessageBox.Show("AI助手功能将在完整版中提供。\n当前显示的是临时消息，确保功能可用。", 
                "AIHelper", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        public void OnSettings(IRibbonControl control)
        {
            WriteLog("OnSettings 方法被调用，控件ID: " + control.Id);
            
            // 显示简单的设置对话框
            Form settingsForm = new Form();
            settingsForm.Text = "API Settings";
            settingsForm.Size = new Size(400, 200);
            settingsForm.StartPosition = FormStartPosition.CenterScreen;
            
            Label apiKeyLabel = new Label();
            apiKeyLabel.Text = "API Key:";
            apiKeyLabel.Location = new Point(20, 20);
            apiKeyLabel.Size = new Size(80, 20);
            
            TextBox apiKeyTextBox = new TextBox();
            apiKeyTextBox.Location = new Point(100, 20);
            apiKeyTextBox.Size = new Size(250, 20);
            apiKeyTextBox.Text = apiKey;
            
            Button saveButton = new Button();
            saveButton.Text = "Save";
            saveButton.Location = new Point(150, 120);
            saveButton.DialogResult = DialogResult.OK;
            
            settingsForm.Controls.Add(apiKeyLabel);
            settingsForm.Controls.Add(apiKeyTextBox);
            settingsForm.Controls.Add(saveButton);
            
            if (settingsForm.ShowDialog() == DialogResult.OK)
            {
                apiKey = apiKeyTextBox.Text;
                SaveSettings();
                MessageBox.Show("设置已保存", "完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        public void OnAbout(IRibbonControl control)
        {
            WriteLog("OnAbout 方法被调用，控件ID: " + control.Id);
            MessageBox.Show("AIHelper Excel Add-in\n" +
                "Version: 1.0.0\n" +
                "Build Date: " + Assembly.GetExecutingAssembly().GetName().Version + "\n\n" +
                "This add-in provides AI-powered functionality for Microsoft Excel.", 
                "About", MessageBoxButtons.OK, MessageBoxIcon.Information);
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
            }
            catch (Exception ex)
            {
                WriteLog("保存配置错误: " + ex.Message);
            }
        }

        #endregion
    }
} 