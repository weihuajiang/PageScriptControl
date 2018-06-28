using Jint;
using Microsoft.Win32;
using mshtml;
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace FormSubmit2
{
    public delegate string GetDelegent(string a, string b);
    public delegate string GetStringDelegent();
    public delegate int InfoDelegent(string m, string c);
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        Engine engine;
        static ManualResetEvent re = new ManualResetEvent(false);
        static ManualResetEvent getEvent = new ManualResetEvent(false);
        public MainWindow()
        {
            InitializeComponent();
            webbrowser.Navigated += Webbrowser_Navigated;
            webbrowser.LoadCompleted += Webbrowser_LoadCompleted;
            //AllocConsole();
            engine = new Engine(cfg => {
                cfg.AllowClr();
                cfg.AllowDebuggerStatement();
                cfg.DebugMode();
            });
            engine.SetValue("log", new Action<object>(Console.WriteLine));
            engine.SetValue("doc", webbrowser.Document);
            engine.SetValue("go", new Action<string>(Navigate));
            engine.SetValue("setValueById", new Action<string, string>(SetValueById));
            engine.SetValue("setValueByName", new Action<string, string>(SetValueByName));
            engine.SetValue("clickById", new Action<string>(clickById));
            engine.SetValue("clickByName", new Action<string>(clickByName));
            engine.SetValue("value", new Action<string, string>(value));
            engine.SetValue("click", new Action<string>(click));
            engine.SetValue("list", new Action<string>(list));
            engine.SetValue("ready", new Action(WaitForReady));
            engine.SetValue("get", new GetDelegent(get));
            engine.SetValue("paste", new GetStringDelegent(getClipboardText));
            engine.SetValue("view", new Action<string>(view));
            engine.SetValue("info", new InfoDelegent(info));
        }

        [DllImport("kernel32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        static extern bool AllocConsole();

        private void Webbrowser_LoadCompleted(object sender, NavigationEventArgs e)
        {
            Console.WriteLine("LoadCompleted");
            re.Set();
        }

        private void Webbrowser_Navigated(object sender, NavigationEventArgs e)
        {
            Console.WriteLine("Navigated "+e.Uri);
        }

        private void Run(object sender, RoutedEventArgs ex)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Title = "Open Script File";
            if (openFileDialog.ShowDialog() == true)
            {
                StreamReader reader = new StreamReader(openFileDialog.FileName);
                string script = reader.ReadToEnd();
                reader.Close();
                new Thread(new ThreadStart(() =>
                {
                    try
                    {
                        Console.WriteLine(script);
                        engine.Execute(script);
                        
                    }
                    catch (Exception e)
                    {
                        MessageBox.Show(e.Message);
                        //Console.WriteLine(e.StackTrace);
                    }
                })).Start();
            }

        }
        public void SetValueById(string id, string value)
        {
            Application.Current.Dispatcher.Invoke((() =>
            {
                HTMLDocument doc = webbrowser.Document as HTMLDocument;
                IHTMLElement e = doc.getElementById(id);
                if (e != null)
                    e.setAttribute("value", value);
            }));
        }
        public void SetValueByName(string id, string value)
        {
            Application.Current.Dispatcher.Invoke((() =>
            {
                HTMLDocument doc = webbrowser.Document as HTMLDocument;
                IHTMLElementCollection e = doc.getElementsByName(id);
                if (e != null)
                {
                    foreach (IHTMLElement el in e)
                        el.setAttribute("value", value);
                }
            }));
        }
        public void Navigate(string url)
        {
            Application.Current.Dispatcher.Invoke((() =>
            {
                webbrowser.Navigate(url);
            }));
        }
        public void view(string selector)
        {
            Application.Current.Dispatcher.Invoke((() =>
            {
                IHTMLElement[] es = select(selector);
                if (es != null && es.Length > 0)
                {
                    foreach (IHTMLElement e in es)
                        e.scrollIntoView();
                }
            }));
        }
        public void value(string selector, string v)
        {
            Application.Current.Dispatcher.Invoke((() =>
            {
                IHTMLElement[] es = select(selector);
                if (es != null && es.Length > 0)
                {
                    foreach (IHTMLElement e in es)
                        e.setAttribute("value", v);
                }
            }));
        }
        public void click(string selector)
        {
            Application.Current.Dispatcher.Invoke((() =>
            {
                IHTMLElement[] es = select(selector);
                if(es!=null && es.Length>0)
                {
                    foreach (IHTMLElement e in es)
                        e.click();
                }
            }));
        }
        public string get(string selector, string attribute)
        {
            getEvent.Reset();
            string result = "";
            Application.Current.Dispatcher.Invoke((() =>
            {
                IHTMLElement[] es = select(selector);
                if (es != null && es.Length > 0)
                {
                    result = es[0].getAttribute(attribute);
                }
                getEvent.Set();
            }));
            getEvent.WaitOne(50000);
            return result;
        }
        public void list(string selector)
        {
            Application.Current.Dispatcher.Invoke((() =>
            {
                IHTMLElement[] es = select(selector);
                if (es != null && es.Length > 0)
                {
                    foreach (IHTMLElement e in es)
                        Console.WriteLine(e.outerHTML);
                }
            }));
        }
        public IHTMLElement[] select(string selector)
        {
            string[] selectors = selector.Split(new string[] { " " }, StringSplitOptions.RemoveEmptyEntries);
            IHTMLElement[] results = null;
            for (int i = 0; i < selectors.Length; i++)
            {
                if (i == 0)
                {
                    HTMLDocument doc = webbrowser.Document as HTMLDocument;
                    results = select(selectors[0], doc.all);
                }
                else
                {
                    if (results == null)
                        return null;
                    results = select(selectors[i], results);
                }
            }
            return results;
        }
        private IHTMLElement[] select(string option, IEnumerable elements)
        {
            if (option.Equals(">"))
            {
                List<IHTMLElement> list = new List<IHTMLElement>();
                foreach(IHTMLElement e in elements)
                {
                    var child = e.children;
                    if (child is IHTMLElement)
                        list.Add(child);
                    else if(child is IHTMLElementCollection)
                    {
                        foreach (IHTMLElement c in child)
                            list.Add(c);
                    }
                }
                return list.ToArray();
            }
            else if (option.StartsWith("."))
            {
                string c = option.Substring(1);
                List<IHTMLElement> list = new List<IHTMLElement>();
                foreach(IHTMLElement e in elements)
                {
                    if (string.IsNullOrEmpty(e.className))
                        continue;
                    string[] cs = e.className.Split(new string[] { " " }, StringSplitOptions.RemoveEmptyEntries);
                    if(cs!=null && cs.Length > 0)
                    {
                        foreach(string cname in cs)
                        {
                            if (c.Equals(cname))
                            {
                                if(!list.Contains(e))
                                    list.Add(e);
                            }
                        }
                    }
                }
                return list.ToArray();
            }
            else if (option.StartsWith("#"))
            {
                string id = option.Substring(1);
                List<IHTMLElement> list = new List<IHTMLElement>();
                foreach(IHTMLElement e in elements)
                {
                    if (id.Equals(e.id) && !list.Contains(e))
                        list.Add(e);
                }
                return list.ToArray();
            }
            else
            {
                option = option.Trim();
                string tagName = option;
                Dictionary<string, string> attrs = new Dictionary<string, string>();
                int a = option.IndexOf("[");
                int b = option.IndexOf("]");
                if (a >0)
                {
                    tagName = tagName.Substring(0, a);
                }
                if (a == 0)
                    tagName = "";
                //Console.WriteLine("Tag name is " + tagName);
                if(a>=0 && b > a)
                {
                    string attString = option.Substring(a + 1, b - a-1);
                    string[] attStrs = attString.Split(new string[] { "," }, StringSplitOptions.RemoveEmptyEntries);
                    foreach(string s in attStrs)
                    {
                        int c = s.IndexOf('=');
                        if (c > 0)
                        {
                            string aname = s.Substring(0, c).Trim();
                            string avalue = s.Substring(c + 1).Trim();
                            if (avalue.StartsWith("\"") || avalue.StartsWith("'"))
                                avalue = avalue.Substring(1);
                            if (avalue.EndsWith("\"") || avalue.EndsWith("'"))
                                avalue = avalue.Substring(0, avalue.Length - 1);
                            attrs.Add(aname, avalue);
                            //Console.WriteLine("attribute " + aname + "=" + avalue);
                        }
                    }
                }
                List<IHTMLElement> list = new List<IHTMLElement>();
                foreach(IHTMLElement e in elements)
                {
                    bool tagNameOk = false;
                    if (!string.IsNullOrEmpty(tagName))
                    {
                        if (tagName.Equals(e.tagName, StringComparison.CurrentCultureIgnoreCase))
                            tagNameOk = true;
                    }
                    else {
                        tagNameOk = true;
                    }
                    if (!tagNameOk)
                        continue;
                    bool attOk = true;
                    foreach(string key in attrs.Keys)
                    {
                        string attrV = e.getAttribute(key)+"";
                        if (!attrs[key].Equals(attrV))
                            attOk = false;
                    }
                    if (attOk)
                    {
                        //Console.WriteLine("find element " + e.tagName + "   " + e.id);
                        list.Add(e);
                    }
                }
                return list.ToArray();
            }
        }
        public void WaitForReady()
        {
            re.Reset();
            re.WaitOne();
        }
        public void clickById(string id)
        {
            Application.Current.Dispatcher.Invoke((() =>
            {
                HTMLDocument doc = webbrowser.Document as HTMLDocument;
                IHTMLElement e = doc.getElementById(id);
                if (e != null)
                {
                    e.click();
                    Console.WriteLine("click on " + e);
                }
            }));
        }
        public void clickByName(string name)
        {
            Application.Current.Dispatcher.Invoke((() =>
            {
                HTMLDocument doc = webbrowser.Document as HTMLDocument;
                IHTMLElementCollection e = doc.getElementsByName(name);
                if (e != null)
                {
                    foreach (IHTMLElement el in e)
                        el.click();
                }
            }));
        }

        private void OnList(object sender, RoutedEventArgs ex)
        {
            HTMLDocument doc = webbrowser.Document as HTMLDocument;
            foreach(IHTMLElement e in doc.all)
            {
                string name = e.getAttribute("name")+"";
                if(!string.IsNullOrEmpty(e.id) || !string.IsNullOrEmpty(name))
                    Console.WriteLine(e.tagName+"\t"+e.id + "\t" + name+"\t"+e.getAttribute("type")+"\t"+e.className);
            }
        }

        private void CopyData(object sender, RoutedEventArgs e)
        {
            string data=(Clipboard.GetText(TextDataFormat.UnicodeText));
            string[] lines = data.Split(new string[] { "\r\n" }, StringSplitOptions.RemoveEmptyEntries);
            foreach(string l in lines)
            {

            }
        }
        public string getClipboardText()
        {
            getEvent.Reset();
            string result = "";
            Application.Current.Dispatcher.Invoke((() =>
            {
                result= Clipboard.GetText(TextDataFormat.UnicodeText);
                getEvent.Set();
            }));
            getEvent.WaitOne(50000);
            return result;
        }
        public int info(string s, string t)
        {
            int ret = 0;
            getEvent.Reset();
            Application.Current.Dispatcher.Invoke((() =>
            {
                if (MessageBox.Show(s, t, MessageBoxButton.YesNo) == MessageBoxResult.Yes)
                    ret= 1;
                getEvent.Set();
            }));
            getEvent.WaitOne();
            return ret;
        }

        private void RunDefault(object sender, RoutedEventArgs ez)
        {
            StreamReader reader = new StreamReader("roche.txt");
            string script = reader.ReadToEnd();
            reader.Close(); new Thread(new ThreadStart(() =>
            {
                try
                {
                    Console.WriteLine(script);
                    engine.Execute(script);

                }
                catch (Exception e)
                {
                    MessageBox.Show(e.Message);
                    //Console.WriteLine(e.StackTrace);
                }
            })).Start();
        }
    }
}
