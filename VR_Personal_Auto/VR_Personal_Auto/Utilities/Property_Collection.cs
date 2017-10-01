using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace VR_Personal_Auto
    
{
    public enum Property_type
    {
        Name,
        Id,
        XPath,
        LinkText,
        CssName,

    }
    class Property_Collection
    {
        
        public static IWebDriver driver { get; set; }
    }
    
}
