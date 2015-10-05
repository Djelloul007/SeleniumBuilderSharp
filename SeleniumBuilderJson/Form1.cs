using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using OpenQA.Selenium;
using OpenQA.Selenium.Firefox;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.IE;
using OpenQA.Selenium.Edge;
using OpenQA.Selenium.PhantomJS;
using OpenQA.Selenium.Support.UI;
using OpenQA.Selenium.Interactions;
using System.Threading;
using Microsoft.Office.Interop.Excel;

namespace SeleniumBuilderJson
{
    public partial class Form1 : Form
    {
        public static IWebDriver driver;
        private StringBuilder verificationErrors;
        private string UserName, Password;
        private string baseURL = "https://is-int1.brandmaker.com";
        public static string filePath = System.Windows.Forms.Application.StartupPath + "\\ShopSetupData.xlsx";
        public string sitProxyUrle = "http://localhost:8080/bodgeit/";
        public static string ZapProxyStatus = "1";
        public int SleepTimemilliseconds = 500;


        public Form1()
        {
            InitializeComponent();

        }

        private void Form1_Load(object sender, EventArgs e)
        {
            //read Root folder
            //LoadFoldersInTreeView(treeView1);
            // Get the current directory.
            string path = Directory.GetCurrentDirectory();
            ListDirectory(treeView1, path);

            //Tree2 Tesplan
            treeView2.BeginUpdate();
            treeView2.Nodes.Add("TestPlan");
            treeView2.Nodes[0].Nodes.Add("TestCase001");
            treeView2.EndUpdate();


            //* Implement Drag&Drop for the Tree
            this.treeView2.AllowDrop = true;
            this.treeView1.ItemDrag += new System.Windows.Forms.ItemDragEventHandler(this.treeView_ItemDrag);
            this.treeView2.DragEnter += new System.Windows.Forms.DragEventHandler(this.treeView_DragEnter);
            this.treeView2.DragDrop += new System.Windows.Forms.DragEventHandler(this.treeView_DragDrop);            
        }

        private void treeView_ItemDrag(object sender,System.Windows.Forms.ItemDragEventArgs e)
        {
            DoDragDrop(e.Item, DragDropEffects.Move);                    
        }

        


        private void treeView_DragEnter(object sender,
            System.Windows.Forms.DragEventArgs e)
        {
            e.Effect = DragDropEffects.Move;
        }


        private void treeView_DragDrop(object sender, System.Windows.Forms.DragEventArgs e)
        {
            TreeNode NewNode;
            if (e.Data.GetDataPresent("System.Windows.Forms.TreeNode", false))
            {
                System.Drawing.Point pt = ((TreeView)sender).PointToClient(new System.Drawing.Point(e.X, e.Y));
                TreeNode DestinationNode = ((TreeView)sender).GetNodeAt(pt);
                NewNode = (TreeNode)e.Data.GetData("System.Windows.Forms.TreeNode");
                {
                    //DestinationNode.Nodes.Add((TreeNode)NewNode.Clone());
                    DestinationNode.Nodes.Add(NewNode.FullPath.ToString());
                    //DestinationNode.Nodes.Add(NewNode.Parent.Nodes.ToString()+"//" + NewNode.ToString());
                    DestinationNode.Expand();
                }
            }
        }


        private static void ListDirectory(TreeView treeView, string path)
        {
            treeView.Nodes.Clear();

            var stack = new Stack<TreeNode>();
            var rootDirectory = new DirectoryInfo(path);
            var node = new TreeNode(rootDirectory.Name) { Tag = rootDirectory };
            stack.Push(node);

            while (stack.Count > 0)
            {
                var currentNode = stack.Pop();
                var directoryInfo = (DirectoryInfo)currentNode.Tag;
                foreach (var directory in directoryInfo.GetDirectories())
                {
                    var childDirectoryNode = new TreeNode(directory.Name) { Tag = directory };
                    currentNode.Nodes.Add(childDirectoryNode);
                    stack.Push(childDirectoryNode);
                }
                foreach (var file in directoryInfo.GetFiles())
                {
                    if (file.Name.ToLower().EndsWith(".json")) {
                        currentNode.Nodes.Add(new TreeNode(file.Name));
                    }
                }
            }

            treeView.Nodes.Add(node);
        }

        public IWebElement GetWebElement(String locatortype, String locatorvalue)
        {
            switch (locatortype)
            {
                case "id":
                    return driver.FindElement(By.Id(locatorvalue));
                   
                case "xpath":
                    return driver.FindElement(By.XPath(locatorvalue));

                case "name":
                    return driver.FindElement(By.Name(locatorvalue));

                case "css":
                    return driver.FindElement(By.CssSelector(locatorvalue));

                case "LinkText":
                    return driver.FindElement(By.LinkText(locatorvalue));

                case "classname":
                    return driver.FindElement(By.ClassName(locatorvalue));

                case "tagname":
                    return driver.FindElement(By.TagName(locatorvalue));

                default:
                    return null;
                    

            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            XmlHandler xmlHandler = new XmlHandler();
            xmlHandler.TreeViewToXml(treeView2, System.Windows.Forms.Application.StartupPath + "\\TestPlan.xml");
        }

        private void button3_Click(object sender, EventArgs e)
        {
            XmlHandler xmlHandler = new XmlHandler();
            //xmlHandler.TreeViewToXml(treeView2, System.Windows.Forms.Application.StartupPath + "\\TestPlan.xml");
            xmlHandler.XmlToTreeView(System.Windows.Forms.Application.StartupPath + "\\TestPlan.xml", treeView2);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            driver = new ChromeDriver();
            richTextBox1.Text = richTextBox1.Text + "Begin Test " + DateTime.Now.ToString() + Environment.NewLine;
            richTextBox1.Text = richTextBox1.Text + Environment.NewLine;
            #region Tree
            foreach (TreeNode tn in treeView2.Nodes[0].Nodes[0].Nodes)
            {
                //MessageBox.Show("Child Node " + tn.ToString());
                int startIndex = tn.ToString().IndexOf("Root");
                //MessageBox.Show(tn.ToString().Substring(startIndex, tn.ToString().Length- startIndex));
                string filename= (tn.ToString().Substring(startIndex, tn.ToString().Length - startIndex));

            





            //richTextBox1.AppendText("T");
          
            String JsonString = File.ReadAllText(System.Windows.Forms.Application.StartupPath + "\\" + filename);
            Selenium selenium = JsonConvert.DeserializeObject<Selenium>(JsonString);

            
            IWebElement webelement;


            #region foreach
            foreach (object element in selenium.steps)
            {
                Step step = JsonConvert.DeserializeObject<Step>(element.ToString());
                //if (step.Type != "get" & step.Type != "verifyTextPresent")
                #region switch
                {

                    switch (step.Type.ToString())
                    {
                        case "get":
                            driver.Navigate().GoToUrl(step.url.ToString());                          
                            driver.Manage().Window.Maximize();
                            Thread.Sleep(SleepTimemilliseconds);
                            Thread.Sleep(SleepTimemilliseconds);
                            Thread.Sleep(SleepTimemilliseconds);
                            break;

                        case "setElementText":
                            webelement = GetWebElement(step.locator.Type.ToString(), step.locator.value.ToString());
                            webelement.Clear();
                            webelement.SendKeys(step.text.ToString());
                            Thread.Sleep(SleepTimemilliseconds);
                            break;
                        case "clickElement":
                            webelement = GetWebElement(step.locator.Type.ToString(), step.locator.value.ToString());
                            webelement.Click();
                            Thread.Sleep(SleepTimemilliseconds);
                            break;

                        case "verifyTextPresent":
                            Thread.Sleep(SleepTimemilliseconds);
                            webelement = GetWebElement(step.locator.Type.ToString(), step.locator.value.ToString());
                            Thread.Sleep(SleepTimemilliseconds);
                            if (webelement.Text== step.text.ToString())
                            {                                
                                richTextBox1.Text = richTextBox1.Text + "     " + step.text.ToString() + " Button " + " was found" + Environment.NewLine;
                                richTextBox1.SelectionColor = Color.Green;
                                int my1stPosition = richTextBox1.Find(step.text.ToString() );
                                richTextBox1.SelectionStart = my1stPosition;
                                richTextBox1.SelectionLength = (step.text.ToString() + " Button " + " was found").Length;
                                //richTextBox1.SelectionFont = fnt;
                                richTextBox1.SelectionColor = Color.Green;

                            }
                            Thread.Sleep(SleepTimemilliseconds);
                            break;

                        case "storeTitle":
                            Thread.Sleep(SleepTimemilliseconds);
                            richTextBox1.Text = richTextBox1.Text + "     " + step.text.ToString() + " storeTitle " + " Not implemented" + Environment.NewLine;
                            //Get Titel
                            break;

                        case "assertTextPresent":
                            Thread.Sleep(SleepTimemilliseconds);
                            webelement = GetWebElement(step.locator.Type.ToString(), step.locator.value.ToString());
                            richTextBox1.Text = richTextBox1.Text + "     " + step.text.ToString() + " assertTextPresent " + " Not implemented" + Environment.NewLine;
                            break;

                        case "switchToFrame":
                            //driver.SwitchTo().Window("windowName");
                            richTextBox1.Text = richTextBox1.Text + "     " + step.text.ToString() + " switchToFrame " + " Not implemented" + Environment.NewLine;
                            //driver.switchTo().frame("frameName");

                            break;

                        case "dragAndDrop":
                            richTextBox1.Text = richTextBox1.Text + "     " + step.text.ToString() + " dragAndDrop " + " Not implemented" + Environment.NewLine;
                            /*
                            WebElement element = driver.findElement(By.name("source"));
                            WebElement target = driver.findElement(By.name("target"));
                            (new Actions(driver)).dragAndDrop(element, target).perform();
                            */
                            break;


                    }
                   
                }
                #endregion switch
            }
            #endregion

           

            }
            #endregion Tree
            richTextBox1.Text = richTextBox1.Text + Environment.NewLine;
            richTextBox1.Text = richTextBox1.Text + "End  Test  " + DateTime.Now.ToString() + Environment.NewLine;
        }//end Button Click
    }


    public class Selenium
    {
        public string Type { get; set; }
        public string seleniumVersion { get; set; }
        public int formatVersion { get; set; }
        public object[] steps { get; set; }

    }

    public class Step
    {
        public string Type { get; set; }
        public string text { get; set; }
        public string url { get; set; }
        public WebElementlocator locator { get; set; }

    }

    public class WebElementlocator
    {
        public string Type { get; set; }
        public string value { get; set; }
    }


}
