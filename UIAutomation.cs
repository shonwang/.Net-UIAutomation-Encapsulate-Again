using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Text;
using System.Windows.Automation;
using System.IO;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Forms;
using Emgu.CV;
using Emgu.CV.CvEnum;
using Emgu.CV.Features2D;
using Emgu.CV.Structure;
using Emgu.CV.UI;
using Emgu.CV.Util;

namespace UIAutomationLibrary
{
    class UIAutomation
    {
        /// <summary>
        /// Click UI object:
        /// args[0] = "Condition"
        /// args[1] = "Click"
        /// args[2] = "left"
        /// Click Image:
        /// args[0] = "E:\modelImage.png"
        /// args[1] = "ClickImageInScreen"
        /// args[2] = "left"
        /// </summary>
        static void Main(string[] args)
        {
            args[0] = args[0].Replace("|", " ");
            Console.WriteLine("Command is " + args[1]);
            UIAutomationClass myUIAutomationClass = new UIAutomationClass();
            AutomationElement myAutomationElement = null;
            if (args[0].IndexOf(@"\") == -1 && (args[0].ToLower().IndexOf(@".png") > -1 || args[0].ToLower().IndexOf(@".jpg") > -1))
            {
                myAutomationElement = myUIAutomationClass.FindFirst(args[0]);
                if (myAutomationElement != null || args[1] == "ControlExists")
                {
                    switch (args[1])
                    {
                        case "RangeValue":
                            myUIAutomationClass.RangeValue(myAutomationElement, args[2]);
                            break;
                        case "Value":
                            myUIAutomationClass.Value(myAutomationElement, args[2]);
                            break;
                        case "Invoke":
                            myUIAutomationClass.Invoke(myAutomationElement);
                            break;
                        case "Scroll":
                            // NoAmount /  NoScroll / LargeDecrement / LargeIncrement / SmallDecrement / SmallIncrement
                            myUIAutomationClass.Scroll(myAutomationElement, args[2], args[3]);
                            break;
                        case "SetScrollPercent":
                            // NoScroll
                            myUIAutomationClass.SetScrollPercent(myAutomationElement, args[2], args[3]);
                            break;
                        case "ExpandCollapse":
                            //option: Expand / Collapse
                            myUIAutomationClass.ExpandCollapse(myAutomationElement, args[2]);
                            break;
                        case "WindowClose":
                            myUIAutomationClass.WindowClose(myAutomationElement);
                            break;
                        case "SetWindowVisualState":
                            //option: Normal / Maximized / Minimized
                            myUIAutomationClass.SetWindowVisualState(myAutomationElement, args[2]);
                            break;
                        case "SelectionItem":
                            // option: AddToSelection / RemoveFromSelection / Select
                            myUIAutomationClass.SelectionItem(myAutomationElement, args[2]);
                            break;
                        case "Dock":
                            //Bottom / Fill / Left / None / Right / Top
                            myUIAutomationClass.Dock(myAutomationElement, args[2]);
                            break;
                        case "Toggle":
                            myUIAutomationClass.Toggle(myAutomationElement);
                            break;
                        case "TransformMove":
                            myUIAutomationClass.TransformMove(myAutomationElement, int.Parse(args[2]), int.Parse(args[3]));
                            break;
                        case "TransformResize":
                            myUIAutomationClass.TransformResize(myAutomationElement, int.Parse(args[2]), int.Parse(args[3]));
                            break;
                        case "TransformRotate":
                            myUIAutomationClass.TransformRotate(myAutomationElement, int.Parse(args[2]));
                            break;
                        case "ScrollItem":
                            myUIAutomationClass.ScrollItem(myAutomationElement);
                            break;
                        case "ControlExists":
                            myUIAutomationClass.ControlExists(myAutomationElement);
                            break;
                        case "SetFocus":
                            myUIAutomationClass.SetFocus(myAutomationElement);
                            break;
                        case "WalkControlElements":
                            myUIAutomationClass.WalkControlElements(myAutomationElement, "   ");
                            break;
                        case "GridGetItemName":
                            myUIAutomationClass.GridGetItem(myAutomationElement, int.Parse(args[2]), int.Parse(args[3]));
                            break;
                        case "MultipleView":
                            myUIAutomationClass.MultipleView(myAutomationElement, int.Parse(args[2]), args[3]);
                            break;
                        case "Click":
                            myUIAutomationClass.Click(myAutomationElement, args[2]);
                            break;
                        case "Shortcut":
                            myUIAutomationClass.Shortcut(myAutomationElement, args[2]);
                            break;
                        case "SendKeys":
                            myUIAutomationClass.SendKeys(myAutomationElement, args[2]);
                            break;
                    }
                }
                else
                {
                    Console.WriteLine("Error: can not find the control...");
                    Console.ReadKey();
                }
            }
            else
            {
                switch (args[1])
                {
                    case "ClickImageScreen":
                        myUIAutomationClass.ClickImageScreen(args[0], args[2]);
                        break;
                    case "ClickImageByForceMatch":
                        myUIAutomationClass.ClickImageByForceMatch(args[0], args[2], args[3]);
                        break;
                    case "ClickImageRegion":
                        myUIAutomationClass.ClickImageRegion(args[0], args[2], args[3]);
                        break;
                    case "ImageExistRegion":
                        myUIAutomationClass.ImageExistRegion(args[0], args[2]);
                        break;
                    case "ImageExistScreen":
                        myUIAutomationClass.ImageExistScreen(args[0]);
                        break;
                }
            }
        }
    }

    /// <summary>
    /// UIAutomationClass
    /// </summary>
    public class UIAutomationClass
    {
        /// <summary>
        /// Find First AutomationElement object according to condition
        /// </summary>
        /// <param name="Condition"></param>
        /// <returns>successed AutomationElement object, unsuccesse null</returns>
        public AutomationElement FindFirst(string Condition)
        {
            //Condition for example: ClassName:CabinetWClass&ControlType:Window&Name:Scripts/ClassName:CabinetWClass&ControlType:Window&Name:Scripts
            AutomationElement result = null;
            AutomationElement[] control = null;
            string[] controlArray = null;

            AutomationElement desktop = AutomationElement.RootElement;
            controlArray = Condition.Split('/');
            control = new AutomationElement[controlArray.Length];
            for (int i = 0; i <= controlArray.Length - 1; i++)
            {
                Console.WriteLine("Level: " + i.ToString());
                string[] conditionArray = controlArray[i].Split('&');
                PropertyCondition[] propertyCondition = new PropertyCondition[conditionArray.Length];
                AutomationProperty[] automationProperty = new AutomationProperty[conditionArray.Length];
                for (int k = 0; k <= conditionArray.Length - 1; k++)
                {
                    Console.WriteLine("Condition: " + k.ToString() + " " + conditionArray[k]);
                    string[] valueArray = conditionArray[k].Split(':');
                    ControlType aControlType = null;
                    switch (valueArray[0])
                    {
                        case "ClassName":
                            automationProperty[k] = AutomationElement.ClassNameProperty;
                            break;
                        case "ControlType":
                            automationProperty[k] = AutomationElement.ControlTypeProperty;
                            break;
                        case "AutomationId":
                            automationProperty[k] = AutomationElement.AutomationIdProperty;
                            //Int32.Parse(valueArray[1]);
                            break;
                        case "Name":
                            automationProperty[k] = AutomationElement.NameProperty;
                            break;
                    }
                    switch (valueArray[1])
                    {
                        case "Window":
                            aControlType = ControlType.Window;
                            break;
                        case "Button":
                            aControlType = ControlType.Button;
                            break;
                        case "Calendar":
                            aControlType = ControlType.Calendar;
                            break;
                        case "CheckBox":
                            aControlType = ControlType.CheckBox;
                            break;
                        case "ComboBox":
                            aControlType = ControlType.ComboBox;
                            break;
                        case "Custom":
                            aControlType = ControlType.Custom;
                            break;
                        case "DataGrid":
                            aControlType = ControlType.DataGrid;
                            break;
                        case "DataItem":
                            aControlType = ControlType.DataItem;
                            break;
                        case "Document":
                            aControlType = ControlType.Document;
                            break;
                        case "Edit":
                            aControlType = ControlType.Edit;
                            break;
                        case "Group":
                            aControlType = ControlType.Group;
                            break;
                        case "Header":
                            aControlType = ControlType.Header;
                            break;
                        case "HeaderItem":
                            aControlType = ControlType.HeaderItem;
                            break;
                        case "Hyperlink":
                            aControlType = ControlType.Hyperlink;
                            break;
                        case "Image":
                            aControlType = ControlType.Image;
                            break;
                        case "List":
                            aControlType = ControlType.List;
                            break;
                        case "ListItem":
                            aControlType = ControlType.ListItem;
                            break;
                        case "Menu":
                            aControlType = ControlType.Menu;
                            break;
                        case "MenuBar":
                            aControlType = ControlType.MenuBar;
                            break;
                        case "MenuItem":
                            aControlType = ControlType.MenuItem;
                            break;
                        case "Pane":
                            aControlType = ControlType.Pane;
                            break;
                        case "ProgressBar":
                            aControlType = ControlType.ProgressBar;
                            break;
                        case "RadioButton":
                            aControlType = ControlType.RadioButton;
                            break;
                        case "ScrollBar":
                            aControlType = ControlType.ScrollBar;
                            break;
                        case "Separator":
                            aControlType = ControlType.Separator;
                            break;
                        case "Slider":
                            aControlType = ControlType.Slider;
                            break;
                        case "Spinner":
                            aControlType = ControlType.Spinner;
                            break;
                        case "SplitButton":
                            aControlType = ControlType.SplitButton;
                            break;
                        case "StatusBar":
                            aControlType = ControlType.StatusBar;
                            break;
                        case "Tab":
                            aControlType = ControlType.Tab;
                            break;
                        case "TabItem":
                            aControlType = ControlType.TabItem;
                            break;
                        case "Table":
                            aControlType = ControlType.Table;
                            break;
                        case "Text":
                            aControlType = ControlType.Text;
                            break;
                        case "Thumb":
                            aControlType = ControlType.Thumb;
                            break;
                        case "TitleBar":
                            aControlType = ControlType.TitleBar;
                            break;
                        case "ToolBar":
                            aControlType = ControlType.ToolBar;
                            break;
                        case "ToolTip":
                            aControlType = ControlType.ToolTip;
                            break;
                        case "Tree":
                            aControlType = ControlType.Tree;
                            break;
                        case "TreeItem":
                            aControlType = ControlType.TreeItem;
                            break;
                    }
                    if (valueArray[0] == "ControlType")
                        propertyCondition[k] = new PropertyCondition(automationProperty[k], aControlType);
                    else
                        propertyCondition[k] = new PropertyCondition(automationProperty[k], valueArray[1]);
                }
                if (conditionArray.Length == 1)
                {
                    if (i == 0)
                    {
                        Console.WriteLine("Parent: " + desktop.Current.Name + " " + desktop.Current.LocalizedControlType);
                        control[i] = desktop.FindFirst(TreeScope.Children, propertyCondition[0]);
                        if (control[i] != null)
                            Console.WriteLine("Self: " + control[i].Current.Name + " " + control[i].Current.LocalizedControlType);
                        else
                            Console.WriteLine("Self: can not find...");
                    }
                    else if (control[i - 1] != null)
                    {
                        Console.WriteLine("Parent: " + control[i - 1].Current.Name + " " + control[i - 1].Current.LocalizedControlType);
                        control[i] = control[i - 1].FindFirst(TreeScope.Children, propertyCondition[0]);
                        if (control[i] != null)
                            Console.WriteLine("Self: " + control[i].Current.Name + " " + control[i].Current.LocalizedControlType);
                        else
                            Console.WriteLine("Self: can not find...");
                    }
                }
                else
                {
                    AndCondition andCondition = new AndCondition(propertyCondition);
                    if (i == 0)
                    {
                        Console.WriteLine("Parent: " + desktop.Current.Name + desktop.Current.LocalizedControlType);
                        control[i] = desktop.FindFirst(TreeScope.Children, andCondition);
                        if (control[i] != null)
                            Console.WriteLine("Self: " + control[i].Current.Name + " " + control[i].Current.LocalizedControlType);
                        else
                            Console.WriteLine("Self: can not find...");
                    }
                    else if (control[i - 1] != null)
                    {
                        Console.WriteLine("Parent: " + control[i - 1].Current.Name + " " + control[i - 1].Current.LocalizedControlType);
                        control[i] = control[i - 1].FindFirst(TreeScope.Children, andCondition);
                        if (control[i] == null)
                            Console.WriteLine("Self: can not find...");
                    }
                }
            }
            result = control[control.Length - 1];
            return result;
        }

        /// <summary>
        /// Set Range Value
        /// </summary>
        /// <param name="automationElement">Automation Element object</param>
        /// <param name="value"></param>
        /// <returns>successed 1, unsuccesse 0</returns>
        public int RangeValue(AutomationElement automationElement, string  value)
        {
            try
            {
                RangeValuePattern rangeValuePattern = (RangeValuePattern)automationElement.GetCurrentPattern(RangeValuePattern.Pattern);
                rangeValuePattern.SetValue(double.Parse(value));
                //this.ReturnResult("1");
                return 1;
            }
            catch (InvalidOperationException e)
            {
                Console.WriteLine(e.Message);
                return 0;
                //this.ReturnResult(e.Message);
            }
        }

        /// <summary>
        /// Set Value
        /// </summary>
        /// <param name="automationElement">Automation Element object</param>
        /// <param name="value"></param>
        /// <returns>successed 1, unsuccesse 0</returns>
        public int Value(AutomationElement automationElement, string value)
        {
            try
            {
                ValuePattern valuePattern = (ValuePattern)automationElement.GetCurrentPattern(ValuePattern.Pattern);
                valuePattern.SetValue(value);
                return 1;
                //this.ReturnResult("1");
            }
            catch (InvalidOperationException e)
            {
                Console.WriteLine(e.Message);
                return 0;
                //this.ReturnResult(e.Message);
            }
        }

        /// <summary>
        /// Invoke
        /// </summary>
        /// <param name="automationElement">Automation Element object</param>
        /// <returns>successed 1, unsuccesse 0</returns>
        public int Invoke(AutomationElement automationElement)
        {
            try
            {
                InvokePattern invokePattern = (InvokePattern)automationElement.GetCurrentPattern(InvokePattern.Pattern);
                invokePattern.Invoke();
                return 1;
                //this.ReturnResult("1");
            }
            catch (InvalidOperationException e)
            {
                Console.WriteLine(e.Message);
                return 0;
                //this.ReturnResult(e.Message);
            }
        }

        /// <summary>
        /// Scroll
        /// </summary>
        /// <param name="automationElement">Automation Element object</param>
        /// <param name="myHorizontalAmount">NoAmount /  NoScroll / LargeDecrement / LargeIncrement / SmallDecrement / SmallIncrement</param>
        /// <param name="myVerticalAmount">NoAmount /  NoScroll / LargeDecrement / LargeIncrement / SmallDecrement / SmallIncrementl</param>
        /// <returns>successed 1, unsuccesse 0</returns>
        public int Scroll(AutomationElement automationElement, string myHorizontalAmount, string myVerticalAmount)
        {
            try
            {
                // 一
                ScrollAmount horizontalAmount = ScrollAmount.NoAmount;
                // |
                ScrollAmount verticalAmount = ScrollAmount.NoAmount;
                switch (myHorizontalAmount)
                {
                    case "NoAmount":
                        horizontalAmount = ScrollAmount.NoAmount;
                        break;
                    case "NoScroll":
                        horizontalAmount = (ScrollAmount)ScrollPattern.NoScroll;
                        break;
                    case "LargeDecrement":
                        horizontalAmount = ScrollAmount.LargeDecrement;
                        break;
                    case "LargeIncrement":
                        horizontalAmount = ScrollAmount.LargeIncrement;
                        break;
                    case "SmallDecrement":
                        horizontalAmount = ScrollAmount.SmallDecrement;
                        break;
                    case "SmallIncrement":
                        horizontalAmount = ScrollAmount.SmallIncrement;
                        break;
                }
                switch (myVerticalAmount)
                {
                    case "NoAmount":
                        verticalAmount = ScrollAmount.NoAmount;
                        break;
                    case "NoScroll":
                        verticalAmount = (ScrollAmount)ScrollPattern.NoScroll;
                        break;
                    case "LargeDecrement":
                        verticalAmount = ScrollAmount.LargeDecrement;
                        break;
                    case "LargeIncrement":
                        verticalAmount = ScrollAmount.LargeIncrement;
                        break;
                    case "SmallDecrement":
                        verticalAmount = ScrollAmount.SmallDecrement;
                        break;
                    case "SmallIncrement":
                        verticalAmount = ScrollAmount.SmallIncrement;
                        break;
                }
                ScrollPattern scrollPattern = (ScrollPattern)automationElement.GetCurrentPattern(ScrollPattern.Pattern);
                scrollPattern.Scroll(horizontalAmount, verticalAmount);
                return 1;
                //this.ReturnResult("1");
            }
            catch (InvalidOperationException e)
            {
                Console.WriteLine(e.Message);
                return 0;
                //this.ReturnResult(e.Message);
            }
        }

        /// <summary>
        /// Set Scroll Percent
        /// </summary>
        /// <param name="automationElement">Automation Element object</param>
        /// <param name="myHorizontalPercent">Percent or NoScroll</param>
        /// <param name="myVerticalPercent">Percent or NoScroll</param>
        /// <returns>successed 1, unsuccesse 0</returns>
        public int SetScrollPercent(AutomationElement automationElement, string myHorizontalPercent, string myVerticalPercent)
        {
            try
            {
                // 一
                double horizontalPercent = 0;
                // |
                double verticalPercent = 0;
                switch (myHorizontalPercent)
                {
                    case "NoScroll":
                        horizontalPercent = ScrollPattern.NoScroll;
                        break;
                }
                switch (myVerticalPercent)
                {
                    case "NoScroll":
                        verticalPercent = ScrollPattern.NoScroll;
                        break;
                }
                ScrollPattern scrollPattern = (ScrollPattern)automationElement.GetCurrentPattern(ScrollPattern.Pattern);
                scrollPattern.SetScrollPercent(horizontalPercent, verticalPercent);
                return 1;
                //this.ReturnResult("1");
            }
            catch (InvalidOperationException e)
            {
                Console.WriteLine(e.Message);
                return 0;
                //this.ReturnResult(e.Message);
            }
        }

        /// <summary>
        /// Expand Collapse
        /// </summary>
        /// <param name="automationElement">Automation Element object</param>
        /// <param name="option">option: Expand / Collapse</param>
        /// <returns>successed 1, unsuccesse 0</returns>
        public int ExpandCollapse(AutomationElement automationElement,string option)
        {
            try
            {
                ExpandCollapsePattern expandCollapsePattern = (ExpandCollapsePattern)automationElement.GetCurrentPattern(ExpandCollapsePattern.Pattern);
                if (option == "Expand")
                    expandCollapsePattern.Expand();
                else if (option == "Collapse")
                    expandCollapsePattern.Collapse();
                return 1;
                //this.ReturnResult("1");
            }
            catch (InvalidOperationException e)
            {
                Console.WriteLine(e.Message);
                return 0;
                //this.ReturnResult(e.Message);
            }
        }

        /// <summary>
        /// Grid Get Item
        /// </summary>
        /// <param name="automationElement">Automation Element object</param>
        /// <param name="row"></param>
        /// <param name="column"></param>
        /// <returns>successed GridItem AutomationElement object, unsuccesse null</returns>
        public AutomationElement GridGetItem(AutomationElement automationElement, int row, int column)
        {
            try
            {
                GridPattern gridPattern = (GridPattern)automationElement.GetCurrentPattern(GridPattern.Pattern);
                AutomationElement gridItem = gridPattern.GetItem(row, column);
                return gridItem;
                //string ItemName = gridItem.Current.Name;
                //this.ReturnResult(ItemName);
            }
            catch (ArgumentOutOfRangeException e)
            {
                Console.WriteLine(e.Message);
                return null;
                //this.ReturnResult(e.Message);
            }
        }

        /// <summary>
        /// Grid Get Item Name
        /// </summary>
        /// <param name="automationElement">Automation Element object</param>
        /// <param name="row"></param>
        /// <param name="column"></param>
        /// <returns>successed Item Name, unsuccesse exception message</returns>
        public string GridGetItemName(AutomationElement automationElement, int row, int column)
        {
            try
            {
                GridPattern gridPattern = (GridPattern)automationElement.GetCurrentPattern(GridPattern.Pattern);
                AutomationElement gridItem = gridPattern.GetItem(row, column);
                string ItemName = gridItem.Current.Name;
                byte[] defaultItemName = Encoding.Default.GetBytes(ItemName);
                byte[] temp = Encoding.Convert(Encoding.ASCII, Encoding.UTF8, defaultItemName);
                ItemName = Encoding.Default.GetString(temp);
                return ItemName;
            }
            catch (ArgumentOutOfRangeException e)
            {
                Console.WriteLine(e.Message);
                return (e.Message);
            }
        }

        /// <summary>
        /// MultipleView
        /// </summary>
        /// <param name="automationElement">Automation Element object</param>
        /// <param name="viewId"> View ID</param>
        /// <param name="option"> option: GetViewName / SetCurrentView</param>
        /// <returns>successed ViewName when option is GetViewName, 1 when option is SetCurrentView, unsuccesse exception message or null</returns>
        public string MultipleView(AutomationElement automationElement, int viewId, string option)
        {
            string result = null;
            try
            {
                MultipleViewPattern multipleViewPattern = (MultipleViewPattern)automationElement.GetCurrentPattern(MultipleViewPattern.Pattern);
                if (option == "GetViewName")
                {
                    result = multipleViewPattern.GetViewName(viewId);
                    byte[] defaultItemName = Encoding.Default.GetBytes(result);
                    byte[] temp = Encoding.Convert(Encoding.ASCII, Encoding.UTF8, defaultItemName);
                    result = Encoding.Default.GetString(temp);
                }
                else if (option == "SetCurrentView")
                {
                    multipleViewPattern.SetCurrentView(viewId);
                    result = "1";
                }
                //this.ReturnResult(result);
                return result;
            }
            catch (InvalidOperationException e)
            {
                Console.WriteLine(e.Message);
                //this.ReturnResult(e.Message);
                return e.Message;
            }
        }

        /// <summary>
        /// Window Close
        /// </summary>
        /// <param name="automationElement">Automation Element object</param>
        /// <returns>successed 1, unsuccesse 0</returns>
        public int WindowClose(AutomationElement automationElement)
        {
            try
            {
                WindowPattern windowPattern = (WindowPattern)automationElement.GetCurrentPattern(WindowPattern.Pattern);
                windowPattern.Close();
                return 1;
                //this.ReturnResult("1");
            }
            catch (InvalidOperationException e)
            {
                Console.WriteLine(e.Message);
                return 0;
                //this.ReturnResult(e.Message);
            }
        }

        /// <summary>
        /// Set Window Visual State
        /// </summary>
        /// <param name="automationElement">Automation Element object</param>
        /// <param name="option"> option: Normal / Maximized / Minimized</param>
        /// <returns>successed 1, unsuccesse 0</returns>
        public int SetWindowVisualState(AutomationElement automationElement, string option)
        {
            WindowVisualState windowVisualState = WindowVisualState.Normal;
            try
            {
                switch (option)
                {
                    case "Normal":
                        windowVisualState = WindowVisualState.Normal;
                        break;
                    case "Maximized":
                        windowVisualState = WindowVisualState.Maximized;
                        break;
                    case "Minimized":
                        windowVisualState = WindowVisualState.Minimized;
                        break;
                }
                WindowPattern windowPattern = (WindowPattern)automationElement.GetCurrentPattern(WindowPattern.Pattern);
                windowPattern.SetWindowVisualState(windowVisualState);
                //this.ReturnResult("1");
                return 1;
            }
            catch (InvalidOperationException e)
            {
                Console.WriteLine(e.Message);
                return 0;
                //this.ReturnResult(e.Message);
            }
        }

        /// <summary>
        /// Selection Item
        /// </summary>
        /// <param name="automationElement">Automation Element object</param>
        /// <param name="option"> option: AddToSelection / RemoveFromSelection / Select</param>
        /// <returns>successed 1, unsuccesse 0</returns>
        public int SelectionItem(AutomationElement automationElement, string option)
        {
            try
            {
                SelectionItemPattern selectionItemPatternPattern = (SelectionItemPattern)automationElement.GetCurrentPattern(SelectionItemPattern.Pattern);
                switch (option)
                {
                    case "AddToSelection":
                        selectionItemPatternPattern.AddToSelection();
                        break;
                    case "RemoveFromSelection":
                        selectionItemPatternPattern.RemoveFromSelection();
                        break;
                    case "Select":
                        selectionItemPatternPattern.Select();
                        break;
                }
                return 1;
                //this.ReturnResult("1");
            }
            catch (InvalidOperationException e)
            {
                Console.WriteLine(e.Message);
                return 0;
                //this.ReturnResult(e.Message);
            }
        }

        /// <summary>
        /// Set Dock Position
        /// </summary>
        /// <param name="automationElement">Automation Element object</param>
        /// <param name="myDockPosition">option: Bottom / Fill / Left / None / Right / Top</param>
        /// <returns>successed 1, unsuccesse 0</returns>
        public int Dock(AutomationElement automationElement, string myDockPosition)
        {
            DockPosition dockPosition = DockPosition.Bottom;
            try
            {
                switch(myDockPosition)
                {
                    case "Bottom":
                        dockPosition = DockPosition.Bottom;
                        break;
                    case "Fill":
                        dockPosition = DockPosition.Fill;
                        break;
                    case "Left":
                        dockPosition = DockPosition.Left;
                        break;
                    case "None":
                        dockPosition = DockPosition.None;
                        break;
                    case "Right":
                        dockPosition = DockPosition.Right;
                        break;
                    case "Top":
                        dockPosition = DockPosition.Top;
                        break;
                }
                DockPattern dockPattern = (DockPattern)automationElement.GetCurrentPattern(DockPattern.Pattern);
                dockPattern.SetDockPosition(dockPosition);
                return 1;
                //this.ReturnResult("1");
            }
            catch (InvalidOperationException e)
            {
                Console.WriteLine(e.Message);
                return 0;
                //this.ReturnResult(e.Message);
            }
        }

        /// <summary>
        /// Table Get Item
        /// </summary>
        /// <param name="automationElement">Automation Element object</param>
        /// <param name="row"></param>
        /// <param name="column"></param>
        /// <returns>successed 1, unsuccesse 0</returns>
        public AutomationElement TableGetItem(AutomationElement automationElement, int row, int column)
        {
            try
            {
                TablePattern tablePattern = (TablePattern)automationElement.GetCurrentPattern(TablePattern.Pattern);
                return tablePattern.GetItem(row, column);
            }
            catch (InvalidOperationException e)
            {
                Console.WriteLine(e.Message);
                return null;
            }
        }

        /// <summary>
        /// Toggle
        /// </summary>
        /// <param name="automationElement">Automation Element object</param>
        /// <returns>successed 1, unsuccesse 0</returns>
        public int Toggle(AutomationElement automationElement)
        {
            try
            {
                TogglePattern togglePattern = (TogglePattern)automationElement.GetCurrentPattern(TogglePattern.Pattern);
                togglePattern.Toggle();
                return 1;
                //this.ReturnResult("1");
            }
            catch (InvalidOperationException e)
            {
                Console.WriteLine(e.Message);
                return 0;
                //this.ReturnResult(e.Message);
            }
        }

        /// <summary>
        /// Transform Move
        /// </summary>
        /// <param name="automationElement">Automation Element object</param>
        /// <param name="x">x</param>
        /// <param name="y">y</param>
        /// <returns>successed 1, unsuccesse 0</returns>
        public int TransformMove(AutomationElement automationElement, int x, int y)
        {
            try
            {
                TransformPattern togglePattern = (TransformPattern)automationElement.GetCurrentPattern(TransformPattern.Pattern);
                togglePattern.Move((double)x, (double)y);
                return 1;
                //this.ReturnResult("1");
            }
            catch (InvalidOperationException e)
            {
                Console.WriteLine(e.Message);
                return 0;
                //this.ReturnResult(e.Message);
            }
        }

        /// <summary>
        /// Transform Resize
        /// </summary>
        /// <param name="automationElement">Automation Element object</param>
        /// <param name="width">width</param>
        /// <param name="height">height</param>
        /// <returns>successed 1, unsuccesse 0</returns>
        public int TransformResize(AutomationElement automationElement, int width, int height)
        {
            try
            {
                TransformPattern togglePattern = (TransformPattern)automationElement.GetCurrentPattern(TransformPattern.Pattern);
                togglePattern.Resize((double)width, (double)height);
                return 1;
                //this.ReturnResult("1");
            }
            catch (InvalidOperationException e)
            {
                Console.WriteLine(e.Message);
                return 0;
                //this.ReturnResult(e.Message);
            }
        }

        /// <summary>
        /// Transform Rotate
        /// </summary>
        /// <param name="automationElement">Automation Element object</param>
        /// <param name="degrees">degrees</param>
        /// <returns>successed 1, unsuccesse 0</returns>
        public int TransformRotate(AutomationElement automationElement, int degrees)
        {
            try
            {
                TransformPattern togglePattern = (TransformPattern)automationElement.GetCurrentPattern(TransformPattern.Pattern);
                togglePattern.Rotate((double)degrees);
                return 1;
                //this.ReturnResult("1");
            }
            catch (InvalidOperationException e)
            {
                Console.WriteLine(e.Message);
                return 0;
                //this.ReturnResult(e.Message);
            }
        }

        /// <summary>
        /// Scroll Into View
        /// </summary>
        /// <param name="automationElement">Automation Element object</param>
        /// <returns>successed 1, unsuccesse 0</returns>
        public int ScrollItem(AutomationElement automationElement)
        {
            try
            {
                ScrollItemPattern scrollItemPattern = (ScrollItemPattern)automationElement.GetCurrentPattern(ScrollItemPattern.Pattern);
                scrollItemPattern.ScrollIntoView();
                return 1;
                //this.ReturnResult("1");
            }
            catch (InvalidOperationException e)
            {
                Console.WriteLine(e.Message);
                return 0;
                //this.ReturnResult(e.Message);
            }
        }

        /// <summary>
        /// Check whether there is AutomationElement in screen
        /// </summary>
        /// <param name="automationElement">Automation Element object</param>
        /// <returns>successed 1, unsuccesse 0</returns>
        public int ControlExists(AutomationElement automationElement)
        {
            int result = 0;
            try
            {
                if (automationElement == null)
                    result = 0;
                else
                    result = 1;
                //this.ReturnResult(result);
                return result;
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                return result;
            }
        }

        /// <summary>
        /// Set Focus
        /// </summary>
        /// <param name="automationElement">Automation Element object</param>
        /// <returns>successed 1, unsuccesse 0</returns>
        public int SetFocus(AutomationElement automationElement)
        {
            try
            {
                automationElement.SetFocus();
                return 1;
                //this.ReturnResult("1");
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                return 0;
                //this.ReturnResult(e.Message);
            }
        }

        /// <summary>
        /// Walk Control Elements
        /// </summary>
        /// <param name="rootElement">Automation Element object</param>
        /// <param name="n">number of split char</param>
        /// <returns>none</returns>
        public void WalkControlElements(AutomationElement rootElement, string n)
        {
            AutomationElement elementNode = TreeWalker.ControlViewWalker.GetFirstChild(rootElement);
            n = n + "   ";
            while (elementNode != null)
            {
                string text = n + "|" + "ControlType:" + elementNode.Current.ControlType.ProgrammaticName + "&" +
                                        "ClassName:" + elementNode.Current.ClassName.ToString() + "&" +
                                        "AutomationId:" + elementNode.Current.AutomationId + "&" +
                                        "Name:" + elementNode.Current.Name;
                this.WriteToText(text);
                WalkControlElements(elementNode, n);
                elementNode = TreeWalker.ControlViewWalker.GetNextSibling(elementNode);
            }
        }

        /// <summary>
        /// Wirte Data to txt
        /// </summary>
        /// <param name="text"></param>
        /// <returns>none</returns>
        private void WriteToText(string text)
        {
            StreamWriter newText = null;
            if (!File.Exists(System.Environment.CurrentDirectory + @"\UIAutomationElementTree.txt"))
                newText = File.CreateText(System.Environment.CurrentDirectory + @"\UIAutomationElementTree.txt");
            else
                newText = File.AppendText(System.Environment.CurrentDirectory + @"\UIAutomationElementTree.txt");
            newText.WriteLine(text);
            newText.Close();
        }

        /// <summary>
        /// Wirte Data to txt
        /// </summary>
        /// <param name="result"></param>
        /// <returns>none</returns>
        private void ReturnResult(string result)
        {
            StreamWriter newText = null;
            if (!File.Exists(System.Environment.CurrentDirectory + @"\Scripts\temp\UIAutomation_return.txt"))
            {
                newText = File.CreateText(System.Environment.CurrentDirectory + @"\Scripts\temp\UIAutomation_return.txt");
            }
            else
            {
                File.Delete(System.Environment.CurrentDirectory + @"\Scripts\temp\UIAutomation_return.txt");
                newText = File.CreateText(System.Environment.CurrentDirectory + @"\Scripts\temp\UIAutomation_return.txt");
            }
            newText.Write(result);
            newText.Close();
        }

        /// <summary>
        /// Create Shortcut
        /// </summary>
        /// <param name="automationElement">Automation Element object</param>
        /// <param name="savePath"></param>
        /// <returns>successed 1, unsuccesse 0</returns>
        public int Shortcut(AutomationElement automationElement, string savePath)
        {
            try
            {
                this.SetFocus(automationElement);
                System.Threading.Thread.Sleep(300);
                double LocationX = automationElement.Current.BoundingRectangle.Location.X;
                double LocationY = automationElement.Current.BoundingRectangle.Location.Y;
                double SizeHeight = automationElement.Current.BoundingRectangle.Size.Height;
                double SizeWidth = automationElement.Current.BoundingRectangle.Size.Width;
                Bitmap myImage = new Bitmap((int)SizeWidth, (int)SizeHeight);
                Graphics g = Graphics.FromImage(myImage);
                g.CopyFromScreen(new System.Drawing.Point((int)LocationX, (int)LocationY), new System.Drawing.Point(0, 0), new System.Drawing.Size((int)SizeWidth, (int)SizeHeight));
                myImage.Save(savePath);
                return 1;
                //this.ReturnResult("1");
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                return 0;
                //this.ReturnResult(e.Message);
            }
        }

        [DllImport("user32.dll")]
        extern static bool SetCursorPos(int x, int y);
        [DllImport("user32.dll")]
        extern static void mouse_event(int mouseEventFlag, int incrementX, int incrementY, uint data, UIntPtr extraInfo);
        const int MOUSEEVENTF_MOVE = 0x0001;
        const int MOUSEEVENTF_LEFTDOWN = 0x0002;
        const int MOUSEEVENTF_LEFTUP = 0x0004;
        const int MOUSEEVENTF_RIGHTDOWN = 0x0008;
        const int MOUSEEVENTF_RIGHTUP = 0x0010;
        const int MOUSEEVENTF_MIDDLEDOWN = 0x0020;
        const int MOUSEEVENTF_MIDDLEUP = 0x0040;
        const int MOUSEEVENTF_ABSOLUTE = 0x8000;
        /// <summary>
        /// Click Automation Element object
        /// </summary>
        /// <param name="automationElement">Automation Element object</param>
        /// <param name="leftOrRight">mouse button option :left / right</param>
        /// <returns>successed 1, unsuccesse 0</returns>
        public int Click(AutomationElement automationElement, string leftOrRight)
        {
            int result = 0;
            try
            {
                double LocationX = automationElement.Current.BoundingRectangle.Location.X;
                double LocationY = automationElement.Current.BoundingRectangle.Location.Y;
                double SizeHeight = automationElement.Current.BoundingRectangle.Size.Height;
                double SizeWidth = automationElement.Current.BoundingRectangle.Size.Width;

                int IncrementX = (int)(LocationX + SizeWidth / 2);
                int IncrementY = (int)(LocationY + SizeHeight / 2);
                if (leftOrRight == "left")
                    result = this.MouseClick(IncrementX, IncrementY, "left");
                else if (leftOrRight == "right")
                    result = this.MouseClick(IncrementX, IncrementY, "right");
                return result;
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                return 0;
            }
        }

        /// <summary>
        /// Mouse Click.
        /// </summary>
        /// <param name="IncrementX">X</param>
        /// <param name="IncrementY">Y</param>
        /// <param name="leftOrRight">mouse button option :left / right</param>
        /// <returns>successed 1, unsuccesse 0</returns>
        private int MouseClick(int IncrementX, int IncrementY, string leftOrRight)
        {
            try
            {
                //Make the cursor position to the element.
                SetCursorPos(IncrementX, IncrementY);
                //Make the left mouse down and up.
                if (leftOrRight == "left")
                {
                    mouse_event(MOUSEEVENTF_LEFTDOWN, IncrementX, IncrementY, 0, UIntPtr.Zero);
                    mouse_event(MOUSEEVENTF_LEFTUP, IncrementX, IncrementY, 0, UIntPtr.Zero);
                }
                else if (leftOrRight == "right")
                {
                    mouse_event(MOUSEEVENTF_RIGHTDOWN, IncrementX, IncrementY, 0, UIntPtr.Zero);
                    mouse_event(MOUSEEVENTF_RIGHTUP, IncrementX, IncrementY, 0, UIntPtr.Zero);
                }
                return 1;
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                return 0;
            }
        }

        /// <summary>
        /// Send Keys.
        /// </summary>
        /// <param name="automationElement">Automation Element object</param>
        /// <param name="keys">Keys http://msdn.microsoft.com/zh-cn/library/system.windows.forms.sendkeys.send(v=vs.80).aspx </param>
        /// <returns>successed 1, unsuccesse 0</returns>
        public int SendKeys(AutomationElement automationElement, string keys)
        {
            try
            {
                this.SetFocus(automationElement);
                System.Windows.Forms.SendKeys.SendWait(keys);
                return 1;
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                return 0;
            }
        }

        /// <summary>
        /// Draw the model image and observed image, the matched features and homography projection.
        /// </summary>
        /// <param name="modelImageFileName">The model image</param>
        /// <param name="observedImageBitmap">The observed image</param>
        /// <param name="matchTime">The output total time for computing the homography matrix.</param>
        /// <returns>The model image and observed image, the matched features and homography projection.</returns>
        private System.Drawing.Point[] DrawBruteForceMatch(String modelImageFileName, Bitmap observedImageBitmap, out long matchTime)
            {
                try
                {
                    Image<Gray, Byte> modelImage = new Image<Gray, byte>(modelImageFileName);
                    Image<Gray, Byte> observedImage = new Image<Gray, byte>(observedImageBitmap);
                    HomographyMatrix homography = null;
                    Stopwatch watch;
                    SURFDetector surfCPU = new SURFDetector(500, false);
                    VectorOfKeyPoint modelKeyPoints;
                    VectorOfKeyPoint observedKeyPoints;
                    Matrix<int> indices;

                    Matrix<byte> mask;
                    int k = 2;
                    double uniquenessThreshold = 0.8;

                    //extract features from the object image
                    modelKeyPoints = surfCPU.DetectKeyPointsRaw(modelImage, null);
                    Matrix<float> modelDescriptors = surfCPU.ComputeDescriptorsRaw(modelImage, null, modelKeyPoints);

                    watch = Stopwatch.StartNew();

                    // extract features from the observed image
                    observedKeyPoints = surfCPU.DetectKeyPointsRaw(observedImage, null);
                    Matrix<float> observedDescriptors = surfCPU.ComputeDescriptorsRaw(observedImage, null, observedKeyPoints);
                    BruteForceMatcher<float> matcher = new BruteForceMatcher<float>(DistanceType.L2);
                    matcher.Add(modelDescriptors);

                    indices = new Matrix<int>(observedDescriptors.Rows, k);
                    Matrix<float> dist = new Matrix<float>(observedDescriptors.Rows, k);
                    matcher.KnnMatch(observedDescriptors, indices, dist, k, null);
                    mask = new Matrix<byte>(dist.Rows, 1);
                    mask.SetValue(255);
                    Features2DToolbox.VoteForUniqueness(dist, uniquenessThreshold, mask);

                    int nonZeroCount = CvInvoke.cvCountNonZero(mask);
                    if (nonZeroCount >= 4)
                    {
                        nonZeroCount = Features2DToolbox.VoteForSizeAndOrientation(modelKeyPoints, observedKeyPoints, indices, mask, 1.5, 20);
                        if (nonZeroCount >= 4)
                            homography = Features2DToolbox.GetHomographyMatrixFromMatchedFeatures(modelKeyPoints, observedKeyPoints, indices, mask, 2);
                    }
                    watch.Stop();

                    //Draw the matched keypoints
                    Image<Bgr, Byte> result = Features2DToolbox.DrawMatches(modelImage, modelKeyPoints, observedImage, observedKeyPoints,
                                                            indices, new Bgr(255, 255, 255), new Bgr(255, 255, 255), mask, Features2DToolbox.KeypointDrawType.DEFAULT);

                    System.Drawing.Point[] newpts = null;
                    #region draw the projected region on the image
                    if (homography != null)
                    {
                        //draw a rectangle along the projected model
                        Rectangle rect = modelImage.ROI;
                        PointF[] pts = new PointF[] { 
                                                               new PointF(rect.Left, rect.Bottom),
                                                               new PointF(rect.Right, rect.Bottom),
                                                               new PointF(rect.Right, rect.Top),
                                                               new PointF(rect.Left, rect.Top)};
                        homography.ProjectPoints(pts);
                        //result.DrawPolyline(Array.ConvertAll<PointF, System.Drawing.Point>(pts, System.Drawing.Point.Round), true, new Bgr(Color.Red), 2);
                        //result.Save(@"E:\1.jpg");
                        newpts = Array.ConvertAll<PointF, System.Drawing.Point>(pts, System.Drawing.Point.Round);

                    }
                    #endregion
                    matchTime = watch.ElapsedMilliseconds;
                    return newpts;
                }
                catch (Exception e)
                {
                    Console.WriteLine(e.Message);
                    matchTime = 0;
                    return new System.Drawing.Point[] { new System.Drawing.Point(-1, -1), new System.Drawing.Point(-1, -1), new System.Drawing.Point(-1, -1), new System.Drawing.Point(-1, -1) };
                }
        }

        /// <summary>
        /// Check the model image and observed image, the matched features.
        /// </summary>
        /// <param name="modelImageFileName">The model image</param>
        /// <param name="observedBitmap">The observed image</param>
        /// <param name="rate">similarity</param>
        /// <returns>The center point object of matched image</returns>
        private System.Drawing.Point ImageMatch(String modelImageFileName, Bitmap observedBitmap, double rate = 0.90)
        {       
            double[] min_val;       
            double[] max_val;       
            System.Drawing.Point[] min_loc;       
            System.Drawing.Point[] max_loc;
            try
            {
                Bitmap modelImageBitmap = new Bitmap(modelImageFileName);
                Image<Gray, Byte> modelImage = new Image<Gray, byte>(modelImageFileName);
                Image<Gray, Byte> observedImage = new Image<Gray, byte>(observedBitmap);
                Image<Gray, float> result = observedImage.MatchTemplate(modelImage, TM_TYPE.CV_TM_CCORR_NORMED);
                Image<Gray, float> resultPow = result.Pow(2);
                resultPow.MinMax(out min_val, out max_val, out min_loc, out max_loc);
                Console.WriteLine("max_val[0] rate : {0} : {1}", max_val[0], rate);
                Console.WriteLine("max_loc[0] locatipon : {0} {1}", max_loc[0].X, max_loc[0].Y);
                System.Drawing.Point targetCenter = new System.Drawing.Point(max_loc[0].X + modelImageBitmap.Width / 2, max_loc[0].Y + modelImageBitmap.Height / 2);
                if (max_val[0] > rate)
                    return targetCenter;
                else
                    return new System.Drawing.Point(-1, -1);
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                return new System.Drawing.Point(-1, -1);
            }
        }

        /// <summary>
        /// Click model image in a region by Force Matcher
        /// </summary>
        /// <param name="modelImageFileName">The model image's full name</param>
        /// <param name="XYWidthHeight">Region's X, Y, Width, Height split by '/' such as "0/0/418/556"</param>
        /// <param name="leftOrRight">mouse button option :left / right</param>
        /// <returns>successed 1, unsuccesse 0</returns>
        public int ClickImageByForceMatch(String modelImageFileName, string XYWidthHeight, string leftOrRight)
        {
            int result = 0;
            try
            {
                int x = int.Parse(XYWidthHeight.Split('/')[0]);
                int y = int.Parse(XYWidthHeight.Split('/')[1]);
                int width = int.Parse(XYWidthHeight.Split('/')[2]);
                int height = int.Parse(XYWidthHeight.Split('/')[3]);
                Bitmap observedImage = new Bitmap(width, height);
                Graphics g = Graphics.FromImage(observedImage);
                g.CopyFromScreen(new System.Drawing.Point(x, y), new System.Drawing.Point(0, 0), new System.Drawing.Size(width, height));
                long matchTime;
                System.Drawing.Point[] observedPoints = this.DrawBruteForceMatch(modelImageFileName, observedImage, out matchTime);
                Console.WriteLine("Match Time is {0} Milliseconds", matchTime);
                int IncrementX = (int)(observedPoints[3].X + observedPoints[1].X / 2);
                int IncrementY = (int)(observedPoints[3].Y + observedPoints[1].Y / 2);
                //Make the left mouse down and up.
                if (leftOrRight == "left")
                    result = this.MouseClick(IncrementX, IncrementY, "left");
                else if (leftOrRight == "right")
                    result = this.MouseClick(IncrementX, IncrementY, "right");
                return result;
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                return 0;
            }
        }

        /// <summary>
        /// Click model image in a region
        /// </summary>
        /// <param name="modelImageFileName">The model image's full name</param>
        /// <param name="XYWidthHeight">Region's X, Y, Width, Height split by '/' such as "0/0/418/556"</param>
        /// <param name="leftOrRight">mouse button option :left / right</param>
        /// <returns>successed 1, unsuccesse 0</returns>
        public int ClickImageRegion(String modelImageFileName, string XYWidthHeight, string leftOrRight)
        {
            int result = 0;
            try
            {
                int x = int.Parse(XYWidthHeight.Split('/')[0]);
                int y = int.Parse(XYWidthHeight.Split('/')[1]);
                int width = int.Parse(XYWidthHeight.Split('/')[2]);
                int height = int.Parse(XYWidthHeight.Split('/')[3]);
                Bitmap observedImage = new Bitmap(width, height);
                Graphics g = Graphics.FromImage(observedImage);
                g.CopyFromScreen(new System.Drawing.Point(x, y), new System.Drawing.Point(0, 0), new System.Drawing.Size(width, height));
                System.Drawing.Point observedPoints = this.ImageMatch(modelImageFileName, observedImage, 0.90);
                //Make the left mouse down and up.
                if (leftOrRight == "left")
                    result = this.MouseClick(observedPoints.X, observedPoints.Y, "left");
                else if (leftOrRight == "right")
                    result = this.MouseClick(observedPoints.X, observedPoints.Y, "right");
                return result;
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                return 0;
            }
        }

        /// <summary>
        /// Click model image in screen.
        /// </summary>
        /// <param name="modelImageFileName">The model image's full name</param>
        /// <param name="leftOrRight">mouse button option :left / right</param>
        /// <returns>successed 1, unsuccesse 0</returns>
        public int ClickImageScreen(String modelImageFileName, string leftOrRight)
        {
            int result = 0;
            try
            {
                Bitmap observedBitmap = new Bitmap(Screen.PrimaryScreen.Bounds.Width, Screen.PrimaryScreen.Bounds.Height);
                Graphics gs = Graphics.FromImage(observedBitmap);
                gs.CopyFromScreen(new System.Drawing.Point(0, 0), new System.Drawing.Point(0, 0),
                                               new System.Drawing.Size(Screen.PrimaryScreen.Bounds.Width, Screen.PrimaryScreen.Bounds.Height));
                System.Drawing.Point observedPoints = this.ImageMatch(modelImageFileName, observedBitmap, 0.90);
                //Make the left mouse down and up.
                if (leftOrRight == "left")
                    result = this.MouseClick(observedPoints.X, observedPoints.Y, "left");
                else if (leftOrRight == "right")
                    result = this.MouseClick(observedPoints.X, observedPoints.Y, "right");
                return result;
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                return 0;
            }
        }

        /// <summary>
        /// Check whether there is model image in a region
        /// </summary>
        /// <param name="modelImageFileName">The model image's full name</param>
        /// <param name="XYWidthHeight">Region's X, Y, Width, Height split by '/' such as "0/0/418/556"</param>
        /// <returns>exist 1, not exist 0</returns>
        public int ImageExistRegion(String modelImageFileName, string XYWidthHeight)
        {
            int result = 0;
            try
            {
                int x = int.Parse(XYWidthHeight.Split('/')[0]);
                int y = int.Parse(XYWidthHeight.Split('/')[1]);
                int width = int.Parse(XYWidthHeight.Split('/')[2]);
                int height = int.Parse(XYWidthHeight.Split('/')[3]);
                Bitmap observedImage = new Bitmap(width, height);
                Graphics g = Graphics.FromImage(observedImage);
                g.CopyFromScreen(new System.Drawing.Point(x, y), new System.Drawing.Point(0, 0), new System.Drawing.Size(width, height));
                System.Drawing.Point observedPoints = this.ImageMatch(modelImageFileName, observedImage, 0.90);
                if (observedPoints.X > -1 && observedPoints.Y > -1)
                    result = 1;
                else
                    result = 0;
                return result;
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                return 0;
            }
        }

        /// <summary>
        /// Check whether there is model image in screen
        /// </summary>
        /// <param name="modelImageFileName">The model image's full name</param>
        /// <returns>exist 1, not exist 0</returns>
        public int ImageExistScreen(String modelImageFileName)
        {
            int result = 0;
            try
            {
                Bitmap observedBitmap = new Bitmap(Screen.PrimaryScreen.Bounds.Width, Screen.PrimaryScreen.Bounds.Height);
                Graphics gs = Graphics.FromImage(observedBitmap);
                gs.CopyFromScreen(new System.Drawing.Point(0, 0), new System.Drawing.Point(0, 0),
                                               new System.Drawing.Size(Screen.PrimaryScreen.Bounds.Width, Screen.PrimaryScreen.Bounds.Height));
                System.Drawing.Point observedPoints = this.ImageMatch(modelImageFileName, observedBitmap, 0.90);
                if (observedPoints.X > -1 && observedPoints.Y > -1)
                    result = 1;
                else
                    result = 0;
                return result;
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                return 0;
            }
        }

        /// <summary>
        /// hover model image in a region
        /// </summary>
        /// <param name="modelImageFileName">The model image's full name</param>
        /// <param name="XYWidthHeight">Region's X, Y, Width, Height split by '/' such as "0/0/418/556"</param>
        /// <returns>successed 1, unsuccesse 0</returns>
        public int HoverImageRegion(String modelImageFileName, string XYWidthHeight)
        {
            try
            {
                int x = int.Parse(XYWidthHeight.Split('/')[0]);
                int y = int.Parse(XYWidthHeight.Split('/')[1]);
                int width = int.Parse(XYWidthHeight.Split('/')[2]);
                int height = int.Parse(XYWidthHeight.Split('/')[3]);
                Bitmap observedImage = new Bitmap(width, height);
                Graphics g = Graphics.FromImage(observedImage);
                g.CopyFromScreen(new System.Drawing.Point(x, y), new System.Drawing.Point(0, 0), new System.Drawing.Size(width, height));
                System.Drawing.Point observedPoints = this.ImageMatch(modelImageFileName, observedImage, 0.90);

                //Make the cursor position to the element.
                SetCursorPos(observedPoints.X, observedPoints.Y);
                return 1;
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                return 0;
            }
        }

        /// <summary>
        /// hover model image in screen.
        /// </summary>
        /// <param name="modelImageFileName">The model image's full name</param>
        /// <returns>successed 1, unsuccesse 0</returns>
        public int HoverImageScreen(String modelImageFileName)
        {
            try
            {
                Bitmap observedBitmap = new Bitmap(Screen.PrimaryScreen.Bounds.Width, Screen.PrimaryScreen.Bounds.Height);
                Graphics gs = Graphics.FromImage(observedBitmap);
                gs.CopyFromScreen(new System.Drawing.Point(0, 0), new System.Drawing.Point(0, 0),
                                               new System.Drawing.Size(Screen.PrimaryScreen.Bounds.Width, Screen.PrimaryScreen.Bounds.Height));
                System.Drawing.Point observedPoints = this.ImageMatch(modelImageFileName, observedBitmap, 0.90);
                //Make the cursor position to the element.
                SetCursorPos(observedPoints.X, observedPoints.Y);
                return 1;
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                return 0;
            }
        }

        /// <summary>
        /// Click Automation Element object
        /// </summary>
        /// <param name="automationElement">Automation Element object</param>
        /// <param name="leftOrRight">mouse button option :left / right</param>
        /// <returns>successed 1, unsuccesse 0</returns>
        public int ClickablePoint(AutomationElement automationElement, string leftOrRight)
        {
            try
            {
                System.Windows.Point ClickablePoint = automationElement.GetClickablePoint();
                this.MouseClick((int)ClickablePoint.X, (int)ClickablePoint.Y, leftOrRight);
                return 1;
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                return 0;
            }
        }
    }
}
